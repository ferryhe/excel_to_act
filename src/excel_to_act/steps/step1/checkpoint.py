"""Read-only Stage 1 import checkpoint over an existing finalized source run."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_to_act.steps.conversion_workflow import append_stage_artifact, create_workflow, hash_file, read_json


_REQUIRED_ARTIFACTS = ("source.json", "workbook_manifest.json", "inventory.json", "handoff.json", "quality.json")


def create_checkpoint(run_dir: Path, workflow_dir: Path) -> dict[str, Any]:
    try:
        run_dir = run_dir.expanduser().resolve()
        source = read_json(run_dir / "source.json")
        workbook_manifest = read_json(run_dir / "workbook_manifest.json")
        quality = read_json(run_dir / "quality.json")
        promotion = read_json(run_dir / "promotion.json")
        if not all(isinstance(item, dict) for item in (source, workbook_manifest, quality, promotion)):
            raise ValueError("final Step 1 source records must be JSON objects")
        if source.get("schema_version") != "step1.v1" or workbook_manifest.get("artifact_type") != "workbook_manifest":
            raise ValueError("run directory does not contain a finalized Step 1 source")
        if quality.get("ready_for_next_step") is not True:
            raise ValueError("Step 1 quality record is not ready for Step 2")
        workbook = Path(source["source_path"]).expanduser().resolve()
        source_hash = source.get("sha256")
        if not workbook.is_file() or hash_file(workbook) != source_hash or workbook_manifest.get("sha256") != source_hash:
            raise ValueError("source workbook hash does not match the finalized Step 1 records")
        source_copy = run_dir / "source_evidence" / Path(source.get("relative_path", workbook.name)).name
        if not source_copy.is_file() or hash_file(source_copy) != source_hash:
            raise ValueError("finalized Step 1 source copy is missing or changed")
        declared = promotion.get("files")
        if not isinstance(declared, dict):
            raise ValueError("promotion.json does not contain finalized file hashes")
        declared_paths: dict[str, Path] = {}
        for relative, expected in declared.items():
            candidate = Path(relative)
            if candidate.is_absolute():
                raise ValueError("promotion ledger paths must be relative to the finalized run")
            path = (run_dir / candidate).resolve()
            if not path.is_relative_to(run_dir):
                raise ValueError(f"promotion ledger path escapes the finalized run: {relative}")
            if not path.is_file() or not isinstance(expected, str) or hash_file(path) != expected:
                raise ValueError(f"final Step 1 artifact is missing or changed: {relative}")
            declared_paths[Path(relative).as_posix()] = path
        actual_paths = {
            path.relative_to(run_dir).as_posix()
            for path in run_dir.rglob("*")
            if path.is_file() and path.name != "promotion.json"
        }
        if actual_paths != set(declared_paths):
            raise ValueError("final Step 1 artifact set differs from promotion.json")
        missing_required = set(_REQUIRED_ARTIFACTS) - set(declared_paths)
        if missing_required:
            raise ValueError("promotion.json is missing required Step 1 artifacts: " + ", ".join(sorted(missing_required)))
        artifact_refs = {name: {"path": str(path), "sha256": declared[relative]}
                         for relative, path in declared_paths.items()
                         for name in [relative]}
        promotion_path = run_dir / "promotion.json"
        promotion_hash = hash_file(promotion_path)
        artifact_refs["promotion.json"] = {"path": str(promotion_path), "sha256": promotion_hash}
        if artifact_refs.get("source_evidence/" + source_copy.name, {}).get("sha256") != source_hash:
            raise ValueError("promotion ledger does not bind the finalized Step 1 source copy")
        source_record = {"source_id": None, "run_id": source.get("run_id"), "batch_id": source.get("batch_id"),
                         "workbook_path": str(workbook), "workbook_sha256": source_hash,
                         "step1_source_copy_path": str(source_copy.resolve()), "step1_run_path": str(run_dir)}
        root, _manifest = create_workflow(workflow_dir, source_record)
        inputs = {f"step1:{name}": path for name, path in declared_paths.items()}
        inputs["promotion.json"] = promotion_path
        inputs["source_workbook"] = workbook
        payload = {"schema_version": "step1.import_checkpoint.v1", "tool": "step1.report", "status": "ready_for_review",
                   "source": source_record, "artifacts": artifact_refs,
                   "quality": {"status": quality.get("status"),
                               "ready_for_next_step": quality.get("ready_for_next_step"),
                               "metrics_state": quality.get("metrics_state"),
                               "metrics": quality.get("metrics"), "blockers": quality.get("blockers", [])},
                   "promotion": {"batch_id": promotion.get("batch_id"),
                                 "declared_files": len(declared), "promotion_sha256": promotion_hash,
                                 "checks": promotion.get("checks")}}
        markdown = _checkpoint_markdown(payload)
        revision = append_stage_artifact(root, 1, "import_checkpoint.json", payload, markdown,
                                         input_files=inputs)
        return {"schema_version": payload["schema_version"], "tool": payload["tool"], "status": "pass",
                "workflow": str(root), "revision": revision["revision"], "artifact": revision["artifact"]}
    except Exception as exc:
        return {"schema_version": "step1.import_checkpoint.v1", "tool": "step1.report", "status": "blocked",
                "reason": str(exc)}


def _checkpoint_markdown(
    value: dict[str, Any], *, machine_link: str = "import_checkpoint.json",
    workflow_link: str = "../../workflow.json",
) -> str:
    source = value["source"]
    quality = value["quality"]
    metrics = quality.get("metrics") or {}
    blockers = quality.get("blockers", [])
    controls = metrics.get("control_modules") or {}

    def shown(item: Any) -> str:
        return "Not recorded" if item is None else str(item)

    detail = "\n".join([
        "# Step 1 Import Checkpoint", "", f"Status: **{value['status']}**", "",
        f"Source run: `{source['run_id']}`", f"Workbook: `{source['workbook_path']}`",
        f"Workbook SHA-256: `{source['workbook_sha256']}`", "",
        f"Step 1 quality status: `{shown(quality.get('status'))}`; ready for Step 2: `{shown(quality.get('ready_for_next_step'))}`",
        f"Exact logical objects: {shown(metrics.get('logical_objects_accounted'))} / "
        f"{shown(metrics.get('logical_objects_total'))}",
        f"Package preservation: {shown(metrics.get('package_parts_preserved'))} / "
        f"{shown(metrics.get('package_parts_total'))}",
        f"Parsed package parts: {shown(metrics.get('parsed_package_parts'))} / "
        f"{shown(metrics.get('parsed_package_parts_total'))} ({shown(metrics.get('parsed_package_parts_ratio'))})",
        f"Opaque package parts: {shown(metrics.get('opaque_parts'))} ({shown(metrics.get('opaque_rate'))})",
        f"Checkboxes: {shown(controls.get('checkbox_controls'))} total; "
        f"{shown(controls.get('checkbox_resolved'))} resolved; {shown(controls.get('checkbox_invalid'))} invalid",
        f"ActiveX controls/handlers: {shown(controls.get('activex_controls'))} / "
        f"{shown(controls.get('activex_handlers'))}; VBA modules/procedures: "
        f"{shown(controls.get('vba_modules'))} / {shown(controls.get('vba_procedures'))}",
        f"Blockers: {len(quality.get('blockers', []))}", "",
        "This checkpoint references the finalized Step 1 files and source workbook without changing them.",
        "",
    ])
    handoff_json = value.get("artifacts", {}).get("handoff.json", {}).get("path")
    handoff_md = value.get("artifacts", {}).get("handoff.md", {}).get("path")
    front = [
        "# Step 1 Import Checkpoint", "",
        "## What this checkpoint asks you to accept", "",
        f"A checked factual extraction and preservation handoff for run `{source.get('run_id', 'Not recorded')}`. "
        "This accepts source inventory and preservation evidence; it does not interpret business meaning or claim that macros ran.", "",
        "## Purpose and scope", "",
        f"Source workbook: `{Path(source['workbook_path']).name}`. Step 1 quality is "
        f"**{shown(quality.get('status'))}**; ready for Step 2: **{shown(quality.get('ready_for_next_step'))}**.", "",
        "## Verified results", "",
        f"- Logical objects accounted: {shown(metrics.get('logical_objects_accounted'))} / {shown(metrics.get('logical_objects_total'))}.",
        f"- Workbook package parts preserved: {shown(metrics.get('package_parts_preserved'))} / {shown(metrics.get('package_parts_total'))}.",
        f"- Opaque package parts: {shown(metrics.get('opaque_parts'))}.",
        f"- Checkbox bindings: {shown(controls.get('checkbox_resolved'))} resolved / "
        f"{shown(controls.get('checkbox_controls'))}; {shown(controls.get('checkbox_invalid'))} invalid references retained.",
        f"- Step 1 blockers: {len(blockers) if isinstance(blockers, list) else 'Not recorded'}.", "",
        "## Limits and unresolved items", "",
        "Opaque package parts are preserved source bytes, not parsed Excel semantics. Invalid checkbox references remain source facts. ActiveX and VBA are static handoffs only; macros were not run. Formula caches are saved values and are not recalculated.",
    ]
    if isinstance(blockers, list) and blockers:
        front.extend(["", "Recorded blockers:", ""] + [f"- {item}" for item in blockers])
    front.extend([
        "", "## Handover and review", "",
        f"- Machine checkpoint: [JSON handoff](<{machine_link}>).",
        f"- Current workflow ledger: [workflow.json](<{workflow_link}>).",
    ])
    if handoff_json:
        front.append(f"- Step 1 machine handoff: [handoff.json](<{Path(handoff_json).as_posix()}>).")
    if handoff_md:
        front.append(f"- Step 1 person handoff: [handoff.md](<{Path(handoff_md).as_posix()}>).")
    front.extend([
        "- Next: the Agent reviews this exact pair, then the actual human reviews it after the Agent decision. Record human approval only after an explicit response.",
        "- To reject: record a rejection routed to Step 1, preserve this revision, revise the affected evidence, and obtain fresh decisions.",
        "- A successful quality check is technical evidence; formal acceptance is recorded separately in the workflow ledger.", "",
        "## Detailed extraction evidence", "", detail.replace("# Step 1 Import Checkpoint", "### Original extraction details", 1),
    ])
    return "\n".join(front)
