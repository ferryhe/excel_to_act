"""Read-only Stage 2 index checkpoint, gated by the Stage 1 review receipt."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_to_act.steps.conversion_workflow import (
    append_stage_artifact,
    bind_source_id,
    display_report_value,
    hash_file,
    load_workflow,
    read_json,
    require_stage_approved,
)
from excel_to_act.steps.step2.workflow import _validate_saved_index


def create_checkpoint(index_path: Path, step1_root: Path, workflow_dir: Path,
                      source_id: str | None = None) -> dict[str, Any]:
    try:
        workflow_root, manifest = load_workflow(workflow_dir)
        _workflow, stage1_revision = require_stage_approved(workflow_root, 1)
        index_path = index_path.expanduser().resolve()
        step1_root = step1_root.expanduser().resolve()
        checked = _validate_saved_index(index_path, step1_root)
        if checked.get("status") != "pass":
            raise ValueError("Step 2 index validation did not pass")
        index = read_json(index_path)
        entries = index.get("entries", []) if isinstance(index, dict) else []
        source = manifest["source"]
        candidates = [item for item in entries if item.get("run_id") == source.get("run_id")
                      and item.get("source_sha256") == source.get("workbook_sha256")
                      and item.get("ready_for_next_step") is True
                      and item.get("status") not in {"blocked", "error", "fail"}]
        if source_id:
            candidates = [item for item in candidates if item.get("source_id") == source_id]
        if len(candidates) != 1:
            raise ValueError("index must select exactly one ready entry for the reviewed Step 1 source")
        entry = candidates[0]
        expected_run = (step1_root / entry.get("run_path", "")).resolve()
        if expected_run != Path(source["step1_run_path"]).resolve():
            raise ValueError("selected Step 2 entry does not point to the reviewed Step 1 source run")
        bind_source_id(workflow_root, entry["source_id"])
        _workflow, manifest = load_workflow(workflow_root)
        payload = {"schema_version": "step2.index_checkpoint.v1", "tool": "step2.report", "status": "pass",
                   "source": {"source_id": entry["source_id"], "run_id": entry["run_id"],
                              "source_sha256": entry["source_sha256"], "source_path": entry["source_path"]},
                   "index": {"path": str(index_path), "sha256": hash_file(index_path),
                             "batch_id": index.get("batch_id"), "status": checked.get("status"),
                             "metrics": checked.get("metrics")},
                   "step1_validation": {"status": checked.get("status"),
                                        "artifacts_checked": checked.get("metrics", {}).get("artifacts_checked")},
                   "selected_entry": {"status": entry.get("status"),
                                      "ready_for_next_step": entry.get("ready_for_next_step"),
                                      "source_id": entry.get("source_id")}}
        markdown = _checkpoint_markdown(
            payload,
            upstream_link=str(workflow_root / stage1_revision["artifact"]["json"]),
        )
        revision = append_stage_artifact(workflow_root, 2, "index_checkpoint.json", payload, markdown,
                                         input_files={"step2_index": index_path},
                                         input_stages={1: stage1_revision["artifact"]["json_sha256"]})
        return {"schema_version": payload["schema_version"], "tool": payload["tool"], "status": "pass",
                "workflow": str(workflow_root), "revision": revision["revision"], "artifact": revision["artifact"]}
    except Exception as exc:
        return {"schema_version": "step2.index_checkpoint.v1", "tool": "step2.report", "status": "blocked",
                "reason": str(exc)}


def _checkpoint_markdown(
    value: dict[str, Any], *, machine_link: str = "index_checkpoint.json",
    workflow_link: str = "../../workflow.json", upstream_link: str | None = None,
    index_link: str | None = None,
) -> str:
    source = value["source"]
    index = value["index"]
    metrics = index.get("metrics")
    if not isinstance(metrics, dict):
        metrics = {}
    selected_entry = value.get("selected_entry")
    if not isinstance(selected_entry, dict):
        selected_entry = {}
    lines = [
        "# Step 2 Index Checkpoint", "",
        "## What this checkpoint asks you to accept", "",
        f"A validated index entry for `{display_report_value(source.get('source_path'))}`, bound to the Step 1 source run. "
        "It confirms navigation and referenced artifacts; it does not assign business meaning to the workbook.", "",
        "## Purpose and scope", "",
        f"Index validation: **{display_report_value(index.get('status'))}**; selected entry status: "
        f"**{display_report_value(selected_entry.get('status'))}**; ready for next step: "
        f"**{display_report_value(selected_entry.get('ready_for_next_step'))}**.", "",
        "## Verified results", "",
        f"- Source run: `{display_report_value(source.get('run_id'))}`; source ID: `{display_report_value(source.get('source_id'))}`.",
        f"- Index entries checked: {display_report_value(metrics.get('entries_checked'))}; artifacts checked: "
        f"{display_report_value(metrics.get('artifacts_checked'))}.",
        f"- Input handoffs checked: {display_report_value(metrics.get('input_handoff_checked'))}.", "",
        "## Limits and unresolved items", "",
        "The selected source entry is labeled by its recorded status. Index validation does not interpret formulas, "
        "resolve opaque workbook structures, or establish model behavior.", "",
        "## Handover and review", "",
        f"- Machine checkpoint: [index_checkpoint.json](<{machine_link}>).",
        f"- Validated source index: [index.json](<{Path(index_link or index['path']).as_posix()}>).",
        f"- Current acceptance ledger: [workflow.json](<{workflow_link}>).",
    ]
    if upstream_link:
        lines.append(f"- Bound Step 1 checkpoint: [Step 1 JSON](<{Path(upstream_link).as_posix()}>).")
    lines.extend([
        "- Next: review this exact report pair as Agent, then request the actual human decision. A clarification or technical PASS is not an approval.",
        "- To reject: record the return route to Step 1 or Step 2, keep this revision, refresh affected reports, and obtain fresh decisions.",
        "- The report's `pass` status describes index validation. Agent PASS and human acceptance are separate decisions recorded in the current ledger.", "",
        "## Detailed index evidence", "",
        f"Report status: **{value.get('status', 'Not recorded')}**", "",
        f"Selected source: `{source['source_id']}`", f"Step 1 run: `{source['run_id']}`",
        f"Workbook SHA-256: `{source['source_sha256']}`", f"Index: `{index['path']}`",
        f"Index SHA-256: `{index['sha256']}`", "",
        f"Read-only index validation: `{index['status']}`; artifacts checked: "
        f"{metrics.get('artifacts_checked', 'unknown')}.", "",
        "This checkpoint references the validated index without modifying it.", "",
    ])
    return "\n".join(lines)
