"""Final conversion report across the six confirmed workflow checkpoints."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

from excel_to_act.steps.conversion_workflow import (
    append_stage_artifact,
    current_revision,
    hash_file,
    load_workflow,
    read_json,
    require_stage_approved,
    stage_status,
)


def create_conversion_report(workflow_dir: Path) -> dict[str, Any]:
    try:
        root, manifest = load_workflow(workflow_dir)
        workflow, stage5_revision = require_stage_approved(root, 5)
        stage_refs: dict[str, Any] = {}
        input_files: dict[str, Path] = {}
        input_stages: dict[int, str] = {}
        for stage in range(1, 6):
            state = stage_status(root, manifest, stage)
            if state["status"] not in {"approved", "simulated_approved"}:
                raise ValueError(f"Stage {stage} is {state['status']}; Step 6 requires every earlier checkpoint approved")
            revision = current_revision(manifest, stage)
            if revision is None:
                raise ValueError(f"Stage {stage} artifact is missing")
            json_path, md_path = root / revision["artifact"]["json"], root / revision["artifact"]["md"]
            payload = read_json(json_path)
            stage_refs[str(stage)] = {"revision": revision["revision"],
                                      "status": state["status"],
                                      "reviewer_pair": state.get("reviewer_pair", []),
                                      "agent_receipt": state.get("agent"),
                                      "human_receipt": state.get("human"),
                                      "typesafe_receipt": state.get("typesafe"),
                                      "json_path": str(json_path),
                                      "json_sha256": revision["artifact"]["json_sha256"],
                                      "markdown_path": str(md_path),
                                      "markdown_sha256": revision["artifact"]["md_sha256"],
                                      "artifact_status": payload.get("status") if isinstance(payload, dict) else None}
            input_files[f"stage{stage}_json"] = json_path
            input_files[f"stage{stage}_markdown"] = md_path
            input_stages[stage] = revision["artifact"]["json_sha256"]

        step1 = read_json(Path(stage_refs["1"]["json_path"]))
        stage3 = read_json(Path(stage_refs["3"]["json_path"]))
        stage4 = read_json(Path(stage_refs["4"]["json_path"]))
        stage5 = read_json(Path(stage_refs["5"]["json_path"]))
        design = stage3.get("design", {})
        input_catalog, catalog_paths = _load_input_catalog(root, stage3)
        for catalog_kind, catalog_path in catalog_paths.items():
            input_files[f"step3_input_boundary_{catalog_kind}"] = catalog_path
        coverage = design.get("coverage", {})
        known_ids = coverage.get("known_path_ids", [])
        selected_options = [item for item in design.get("options", [])
                            if isinstance(item, dict) and _option_selected(item)]
        selected_numerical_ids = {path_id for item in selected_options for path_id in item.get("numerical_paths", [])
                                  if isinstance(path_id, str)}
        selected_condition_ids = {path_id for item in selected_options for path_id in item.get("condition_only_paths", [])
                                  if isinstance(path_id, str)}
        valid_ids = set(known_ids) if isinstance(known_ids, list) else set()
        stage4_coverage = stage4.get("path_coverage", {})
        stage5_coverage = stage5.get("path_coverage", {})
        stage4_evidence = {item.get("path_id"): item for item in stage4_coverage.get("implementation_evidence", [])
                           if isinstance(item, dict) and item.get("status") == "generated"
                           and item.get("path_id") in valid_ids & selected_numerical_ids
                           and isinstance(item.get("address"), str)
                           and isinstance(item.get("source_formula"), str)}
        stage5_evidence = []
        for item in stage5_coverage.get("runtime_evidence", []):
            if not isinstance(item, dict) or item.get("status") != "matched":
                continue
            path_id = item.get("path_id")
            implementation = stage4_evidence.get(path_id)
            if (implementation and item.get("address") == implementation.get("address")
                    and item.get("source_formula") == implementation.get("source_formula")
                    and path_id in stage5_coverage.get("runtime_verified_path_ids", [])):
                stage5_evidence.append(item)
        implemented_ids = sorted(stage4_evidence)
        runtime_verified_ids = sorted({item["path_id"] for item in stage5_evidence})
        condition_evidence = [item for item in stage5_coverage.get("condition_evidence", [])
                              if isinstance(item, dict) and item.get("status") == "inactive_condition_checked_no_solver"
                              and item.get("path_id") in valid_ids & selected_condition_ids]
        option_summaries = []
        for option in design.get("options", []):
            if not isinstance(option, dict):
                continue
            numerical = [path_id for path_id in option.get("numerical_paths", []) if isinstance(path_id, str)]
            condition_only = [path_id for path_id in option.get("condition_only_paths", []) if isinstance(path_id, str)]
            selected = _option_selected(option)
            option_summaries.append({"option_id": option.get("option_id"), "selected": selected,
                                     "selection_field": ("selected" if option.get("selected") is True
                                                         else "selected_for_draft" if selected else None),
                                     "selected_for_draft": selected,
                                     "scope": option.get("scope"),
                                     "planned_known_path_coverage": option.get("planned_known_path_coverage"),
                                     "cumulative_path_coverage": option.get("cumulative_path_coverage"),
                                     "all_configuration_coverage": option.get("all_configuration_coverage"),
                                     "exclusions": option.get("exclusions", []),
                                     "actual_numeric_coverage": ({
                                         "planned_ids": numerical,
                                         "implemented_ids": sorted(set(numerical) & set(implemented_ids)),
                                         "runtime_verified_ids": sorted(set(numerical) & set(runtime_verified_ids)),
                                         "condition_only_ids": condition_only,
                                         "condition_assessed_inactive_ids": sorted(
                                             set(condition_only) & {item["path_id"] for item in condition_evidence}),
                                     } if selected else None),
                                     "execution_status": "current_saved_scenario_executed" if selected else "planning_only"})
        actual_coverage = {"implemented_path_ids": implemented_ids,
                           "runtime_verified_path_ids": runtime_verified_ids,
                           "condition_only_path_ids": sorted(selected_condition_ids & valid_ids),
                           "condition_assessed_inactive_ids": sorted({item["path_id"] for item in condition_evidence}),
                           "implementation_evidence": list(stage4_evidence.values()),
                           "runtime_evidence": stage5_evidence,
                           "condition_evidence": condition_evidence}
        delegated_stages = set()
        delegation = manifest.get("review_delegation")
        if isinstance(delegation, dict) and isinstance(delegation.get("stages"), list):
            delegated_stages = {stage for stage in delegation["stages"] if isinstance(stage, int)}
        accepted_reviewers = {
            str(stage): {"status": stage_refs[str(stage)]["status"],
                         "reviewer_pair": stage_refs[str(stage)]["reviewer_pair"],
                         "agent": stage_refs[str(stage)]["agent_receipt"],
                         "human": stage_refs[str(stage)]["human_receipt"],
                         "typesafe": stage_refs[str(stage)]["typesafe_receipt"]}
            for stage in range(1, 6)
        }
        stage6_reviewers = ("agent", "typesafe") if 6 in delegated_stages else ("agent", "human")
        review_policy = {
            "default_reviewer_pair": ["agent", "human"],
            "delegated_typesafe_active": 6 in delegated_stages,
            "delegation": ({"authorization_sha256": delegation.get("authorization_sha256"),
                            "source_sha256": delegation.get("source_sha256"),
                            "stages": sorted(delegated_stages),
                            "requested_model": delegation.get("requested_model"),
                            "minimum_approval_probability": delegation.get("minimum_approval_probability")}
                           if isinstance(delegation, dict) and 6 in delegated_stages else None),
            "accepted_stage_reviews": accepted_reviewers,
            "stage6_required_reviewer_pair": list(stage6_reviewers),
        }
        payload = {"schema_version": "step6.conversion_report.v1", "tool": "step6.report",
                   "status": "ready_for_review", "source": workflow.get("source"),
                   "stages": stage_refs,
                   "quality": step1.get("quality", {}),
                   "review_policy": review_policy,
                   "design_summary": {"targets": design.get("targets", []),
                                      "scope": design.get("scope", {}),
                                      "field_groups": design.get("field_groups", []),
                                      "execution_design": design.get("execution_design", {}),
                                      "shared_modules": design.get("shared_modules", []),
                                      "scenarios": design.get("scenarios", []),
                                      "input_boundary_reference": stage3.get("input_boundary_reference"),
                                      "input_catalog": input_catalog,
                                      "scenario_ids": [item.get("scenario_id") for item in design.get("scenarios", [])],
                                      "known_path_count": len(known_ids),
                                      "documented_path_count": len(coverage.get("documented_path_ids", [])),
                                      "resolved_path_count": len(coverage.get("resolved_path_ids", [])),
                                      "planned_path_count": len(coverage.get("planned_path_ids", [])),
                                      "stage3_implemented_path_count": len(coverage.get("implemented_path_ids", [])),
                                      "stage3_runtime_verified_path_count": len(coverage.get("runtime_verified_path_ids", [])),
                                      "stage3_all_configuration_coverage": coverage.get("all_configuration_coverage"),
                                      "actual_coverage": actual_coverage,
                                      "options": option_summaries,
                                      "validation_plan": design.get("validation_plan", []),
                                      "open_questions": design.get("open_questions", [])},
                   "generation_summary": {"bundle": stage4.get("bundle"),
                                          "formula_count": stage4.get("formula_count"),
                                          "active_formula_member_count": stage4.get("active_formula_member_count",
                                                                                     stage4.get("formula_count")),
                                          "compiled_function_count": stage4.get("compiled_function_count",
                                                                                stage4.get("unique_equation_family_count")),
                                          "source_family_count": stage4.get("source_family_count"),
                                          "semantic_equation_group_count": stage4.get(
                                              "semantic_equation_group_count",
                                              stage4.get("unique_equation_family_count")),
                                          "array_instance_count": stage4.get("array_instance_count"),
                                          "array_member_count": stage4.get("array_member_count",
                                                                            stage4.get("array_follower_count")),
                                          "business_input_count": stage4.get("business_input_count"),
                                          "raw_business_input_count": stage4.get("raw_business_input_count"),
                                          "external_business_input_count": stage4.get("external_business_input_count"),
                                          "model_variable_count": stage4.get("model_variable_count"),
                                          "derived_variable_count": stage4.get("derived_variable_count"),
                                          "raw_input_count": stage4.get("raw_input_count"),
                                          "raw_source_coordinate_count": stage4.get("raw_source_coordinate_count"),
                                          "source_metadata_coordinate_count": stage4.get("source_metadata_coordinate_count"),
                                          "external_input_coordinate_count": stage4.get("external_input_coordinate_count"),
                                          "external_capture": stage4.get("external_capture"),
                                          "target_names": stage4.get("target_names", []),
                                          "run_command": "python model.py --out model_result.json",
                                          "runtime_dependencies": stage4.get("runtime_dependencies")},
                   "validation_summary": stage5.get("compared", {}),
                   "validation_context": {key: stage5[key] for key in (
                       "oracle", "standalone_validation", "standalone_project_imports", "formula_cache_inputs"
                   ) if key in stage5},
                   "target_results": stage5.get("target_results", {}),
                   "tolerances": stage5.get("tolerances", {}),
                   "mismatch_count": stage5.get("mismatch_count"),
                   "readiness": {"approved_saved_scenario_only": True,
                                 "all_configurations": False,
                                 "full_workbook_conversion": False,
                                 "gpu_execution": False}}
        revision_number = len(manifest.get("stages", {}).get("6", {}).get("revisions", [])) + 1
        report_directory = root / "stage6" / f"revision-{revision_number:04d}"
        markdown = _report_markdown(payload, workflow_root=root, report_directory=report_directory)
        revision = append_stage_artifact(root, 6, "conversion_report.json", payload, markdown,
                                         input_files=input_files,
                                         input_stages={stage: revision_hash for stage, revision_hash in input_stages.items()})
        return {"schema_version": payload["schema_version"], "tool": payload["tool"], "status": "pass",
                "workflow": str(root), "revision": revision["revision"], "artifact": revision["artifact"]}
    except Exception as exc:
        return {"schema_version": "step6.conversion_report.v1", "tool": "step6.report",
                "status": "blocked", "reason": str(exc)}


def _option_selected(option: dict[str, Any]) -> bool:
    """Accept the current selection field and legacy Step 3 drafts."""
    return option.get("selected") is True or option.get("selected_for_draft") is True


def _load_input_catalog(root: Path, stage3: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Path]]:
    reference = stage3.get("input_boundary_reference")
    if not isinstance(reference, dict):
        return {"available": False, "reason": "No hash-bound input catalog is referenced by the accepted design."}, {}
    artifact = reference.get("artifact")
    if not isinstance(artifact, dict):
        return {"available": False, "reason": "The accepted design has no complete catalog artifact reference."}, {}
    accepted_source = stage3.get("source")
    expected_source = accepted_source.get("workbook_sha256") if isinstance(accepted_source, dict) else None
    if not isinstance(expected_source, str) or not expected_source.strip():
        return {"available": False,
                "reason": "The accepted design has no nonempty source workbook digest to bind the catalog."}, {}
    json_path = artifact.get("json")
    json_digest = artifact.get("json_sha256")
    if not isinstance(json_path, str) or not isinstance(json_digest, str):
        return {"available": False, "reason": "The catalog JSON path or recorded digest is missing."}, {}

    root_path = root.resolve()
    catalog_path = Path(json_path)
    if not catalog_path.is_absolute():
        catalog_path = root_path / catalog_path
    catalog_path = catalog_path.resolve()
    try:
        catalog_path.relative_to(root_path)
    except ValueError:
        return {"available": False, "reason": "The catalog reference is outside the workflow directory."}, {}
    if not catalog_path.is_file() or hash_file(catalog_path) != json_digest:
        return {"available": False, "reason": "The referenced catalog JSON is missing or its recorded digest does not match."}, {}

    input_paths = {"json": catalog_path}
    markdown_path = artifact.get("md")
    markdown_digest = artifact.get("md_sha256")
    if markdown_path is not None or markdown_digest is not None:
        if not isinstance(markdown_path, str) or not isinstance(markdown_digest, str):
            return {"available": False, "reason": "The catalog Markdown path or recorded digest is incomplete."}, {}
        md_path = Path(markdown_path)
        if not md_path.is_absolute():
            md_path = root_path / md_path
        md_path = md_path.resolve()
        try:
            md_path.relative_to(root_path)
        except ValueError:
            return {"available": False, "reason": "The catalog Markdown reference is outside the workflow directory."}, {}
        if not md_path.is_file() or hash_file(md_path) != markdown_digest:
            return {"available": False, "reason": "The referenced catalog Markdown is missing or its recorded digest does not match."}, {}
        input_paths["markdown"] = md_path

    try:
        catalog = read_json(catalog_path)
    except (OSError, ValueError):
        return {"available": False, "reason": "The referenced catalog JSON could not be read."}, {}
    if not isinstance(catalog, dict) or catalog.get("schema_version") != "step3.input_boundary.v1":
        return {"available": False, "reason": "The referenced file is not a supported Step 3 input-boundary catalog."}, {}
    if catalog.get("phase") != "input_boundary" or catalog.get("boundary_sha256") != reference.get("boundary_sha256"):
        return {"available": False, "reason": "The catalog does not match the boundary recorded by the accepted design."}, {}
    catalog_source_info = catalog.get("source")
    catalog_source = catalog_source_info.get("workbook_sha256") if isinstance(catalog_source_info, dict) else None
    if not isinstance(catalog_source, str) or not catalog_source.strip():
        return {"available": False, "reason": "The catalog has no nonempty source workbook digest."}, {}
    if catalog_source != expected_source:
        return {"available": False, "reason": "The catalog source does not match the accepted design source."}, {}
    catalog_summary = {
        "available": True,
        "revision": reference.get("revision"),
        "boundary_sha256": reference.get("boundary_sha256"),
        "artifact_path": str(catalog_path),
        "artifact_sha256": json_digest,
        "content": catalog,
    }
    if "markdown" in input_paths:
        catalog_summary["markdown_artifact_path"] = str(input_paths["markdown"])
        catalog_summary["markdown_artifact_sha256"] = markdown_digest
    return catalog_summary, input_paths


def _plain(value: Any, fallback: str = "Not recorded") -> str:
    if value is None:
        return fallback
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, (dict, list, tuple)):
        return fallback
    return str(value).replace("\r", " ").replace("\n", " ").strip() or fallback


def _cell(value: Any, fallback: str = "Not recorded") -> str:
    if isinstance(value, (list, tuple)):
        return _human_value(value).replace("|", "\\|")
    return _plain(value, fallback).replace("|", "\\|")


def _human_value(value: Any) -> str:
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (list, tuple)):
        return ", ".join(_human_value(item) for item in value)
    if isinstance(value, dict):
        return "structured value"
    return _plain(value)


def _table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    lines = ["| " + " | ".join(headers) + " |",
             "| " + " | ".join("---" for _ in headers) + " |"]
    lines.extend("| " + " | ".join(_cell(value) for value in row) + " |" for row in rows)
    return lines


def _items(values: Any) -> list[str]:
    if not isinstance(values, list):
        return []
    return [str(value).replace("\r", " ").replace("\n", " ").strip()
            for value in values if isinstance(value, (str, int, float))]


def _format_result(value: Any) -> str:
    if isinstance(value, float):
        return repr(value)
    return _human_value(value)


def _result_metadata(target: dict[str, Any] | None) -> str:
    if target is None:
        return "shape Not recorded; kind Not recorded; units Not recorded; axes Not recorded"
    def atom(value: Any) -> str:
        if value is None:
            return "Not recorded"
        if isinstance(value, bool):
            return "True" if value else "False"
        if isinstance(value, list):
            return "[" + ", ".join(atom(item) for item in value) + "]"
        if isinstance(value, dict):
            return "structured value"
        return str(value).replace("|", "\\|").replace("\r", " ").replace("\n", " ")

    shape = atom(target.get("shape"))
    kind = atom(target.get("result_kind"))
    units = atom(target.get("units"))
    raw_axes = target.get("axes", target.get("axis"))
    if raw_axes is None:
        axes = "Not recorded"
    elif isinstance(raw_axes, list):
        axis_items = []
        for axis in raw_axes:
            if not isinstance(axis, dict):
                axis_items.append(atom(axis))
                continue
            fields = [f"{key}: {atom(axis[key])}" for key in ("name", "axis_id", "role", "units", "source_range") if key in axis]
            if "keys" in axis:
                fields.append(f"keys: {atom(axis['keys'])}")
            axis_items.append("; ".join(fields) or "Not recorded")
        axes = "[]" if not axis_items else " / ".join(axis_items)
    else:
        axes = atom(raw_axes)
    return f"shape {shape}; kind {kind}; units {units}; axes {axes}"


def _structured_result(value: Any, target: dict[str, Any] | None, limit: int = 12) -> str:
    leaves: list[tuple[tuple[Any, ...], Any]] = []

    def visit(item: Any, path: tuple[Any, ...]) -> None:
        if len(leaves) > limit:
            return
        if isinstance(item, list):
            if not item:
                leaves.append((path, "[]"))
            else:
                for index, child in enumerate(item):
                    visit(child, (*path, index))
                    if len(leaves) > limit:
                        return
        elif isinstance(item, dict):
            if not item:
                leaves.append((path, "{}"))
            else:
                for key, child in item.items():
                    visit(child, (*path, key))
                    if len(leaves) > limit:
                        return
        else:
            leaves.append((path, item))

    def atom(item: Any) -> str:
        if item is None:
            return "Not recorded"
        if isinstance(item, bool):
            return "True" if item else "False"
        if isinstance(item, (dict, list)):
            return "structured value"
        return str(item).replace("|", "\\|").replace("\r", " ").replace("\n", " ")

    visit(value, ())
    rows = [_result_metadata(target)]
    for path, leaf in leaves[:limit]:
        location = f"position {list(path)}" if all(isinstance(part, int) for part in path) else f"field path {list(path)}"
        rows.append(f"{location} = {atom(leaf)}")
    if len(leaves) > limit:
        rows.append("Additional values omitted; full values remain in the accepted Step 5 JSON linked in Appendix A.")
    return "<br>".join(rows)


def _result_row(
    classification: str, name: str, record: dict[str, Any], target: dict[str, Any] | None = None,
) -> list[Any]:
    model_value = record.get("model")
    excel_value = record.get("excel")
    difference = (model_value - excel_value
                  if isinstance(model_value, (int, float)) and not isinstance(model_value, bool)
                  and isinstance(excel_value, (int, float)) and not isinstance(excel_value, bool)
                  else None)
    check = ("Matched" if record.get("matched") is True else "Not matched"
             if record.get("matched") is False else "Not reported")
    model_text = (_structured_result(model_value, target) if isinstance(model_value, (list, dict))
                  else _format_result(model_value))
    excel_text = (_structured_result(excel_value, target) if isinstance(excel_value, (list, dict))
                  else _format_result(excel_value))
    return [classification, name, model_text, excel_text, _format_result(difference), check]

def _link(label: str, path: Any, workflow_root: Path | None, report_directory: Path | None) -> str:
    if not isinstance(path, str) or not path:
        return label
    target = Path(path)
    if workflow_root is not None and not target.is_absolute():
        target = workflow_root / target
    if report_directory is not None:
        try:
            href = os.path.relpath(target, report_directory)
        except ValueError:
            href = target.as_posix()
    else:
        href = target.as_posix()
    href = href.replace("\\", "/")
    from urllib.parse import quote

    return f"[{label}]({quote(href, safe='/._-')})"


def _catalog_group_rows(catalog: dict[str, Any]) -> list[list[Any]]:
    variables = catalog.get("variables", [])
    if not isinstance(variables, list):
        return []
    grouped: dict[str, list[dict[str, Any]]] = {}
    for variable in variables:
        if isinstance(variable, dict):
            grouped.setdefault(str(variable.get("group_id", "unassigned")), []).append(variable)
    groups = catalog.get("groups", [])
    group_records = [item for item in groups if isinstance(item, dict)] if isinstance(groups, list) else []
    known_ids = set()
    rows = []
    for group in sorted(group_records, key=lambda item: str(item.get("name", item.get("group_id", ""))).casefold()):
        group_id = str(group.get("group_id", "unassigned"))
        known_ids.add(group_id)
        entries = grouped.get(group_id, [])
        if not entries:
            continue
        shapes: dict[str, int] = {}
        for variable in entries:
            kind = _plain(variable.get("kind"), "other")
            shape = variable.get("shape", [])
            dimensions = " × ".join(str(value) for value in shape) if isinstance(shape, list) and shape else ""
            axes = variable.get("axes", [])
            axis_names = [str(axis.get("name")) for axis in axes
                          if isinstance(axis, dict) and axis.get("name")] if isinstance(axes, list) else []
            description = kind
            if dimensions:
                description += f" ({dimensions}"
                if axis_names:
                    description += f"; {', '.join(axis_names)} axis"
                description += ")"
            shapes[description] = shapes.get(description, 0) + 1
        shape_text = "; ".join(f"{name}: {count}" for name, count in sorted(shapes.items()))
        rows.append([group.get("name", group_id), group.get("classification"), len(entries), shape_text])
    for group_id, entries in sorted(grouped.items()):
        if group_id not in known_ids:
            rows.append([group_id, "Not classified", len(entries), ""])
    return rows


def _catalog_kind_counts(catalog: dict[str, Any]) -> dict[str, int]:
    kinds = {"scalar": 0, "vector": 0, "table": 0}
    variables = catalog.get("variables", [])
    if isinstance(variables, list):
        for variable in variables:
            if isinstance(variable, dict):
                kind = str(variable.get("kind", "")).casefold()
                if kind in kinds:
                    kinds[kind] += 1
        if any(kinds.values()):
            return kinds
    counts = catalog.get("counts", {})
    if isinstance(counts, dict):
        for kind, key in (("scalar", "scalar_field_count"),
                          ("vector", "vector_object_count"), ("table", "table_count")):
            value = counts.get(key)
            if isinstance(value, int) and not isinstance(value, bool):
                kinds[kind] = value
    return kinds


def _catalog_source_tab_rows(catalog: dict[str, Any]) -> list[list[Any]]:
    variables = catalog.get("variables", [])
    if not isinstance(variables, list):
        return []
    by_sheet: dict[str, dict[str, set[str]]] = {}
    for index, variable in enumerate(variables):
        if not isinstance(variable, dict):
            continue
        role = str(variable.get("role", "")).casefold()
        if role not in {"source_raw", "formula_derived_external"}:
            continue
        variable_id = str(variable.get("variable_id", variable.get("logical_name", index)))
        extents = variable.get("source_extents", [])
        if not isinstance(extents, list):
            continue
        for extent in extents:
            if not isinstance(extent, dict) or not isinstance(extent.get("sheet"), str):
                continue
            sheet = extent["sheet"]
            by_sheet.setdefault(sheet, {"source_raw": set(), "formula_derived_external": set()})
            by_sheet[sheet][role].add(variable_id)
    rows = []
    for sheet, roles in sorted(by_sheet.items(), key=lambda item: item[0].casefold()):
        raw_count = len(roles["source_raw"])
        external_count = len(roles["formula_derived_external"])
        rows.append([sheet, raw_count, external_count, raw_count + external_count])
    return rows


def _field_group_rows(design: dict[str, Any]) -> list[list[Any]]:
    groups = design.get("field_groups", [])
    if not isinstance(groups, list):
        return []
    rows = []
    for group in groups:
        if not isinstance(group, dict):
            continue
        extents = group.get("source_extents", [])
        source = "; ".join(_items(extents)) if isinstance(extents, list) else _plain(group.get("source"))
        rows.append([group.get("id"), group.get("shape"), source])
    return rows


def _recurrence_notes(execution: dict[str, Any]) -> list[str]:
    notes = execution.get("preserve_formulas", [])
    if not isinstance(notes, list):
        return []
    relevant = [note for note in _items(notes)
                if re.search(r"\b(prior|recurrence|seed|same-year|year-zero|lagged|interval|snapshot|annual)\b",
                             note, re.IGNORECASE)]
    return relevant[:4]


def _result_relationship_notes(execution: dict[str, Any]) -> list[str]:
    notes = execution.get("preserve_formulas", [])
    if not isinstance(notes, list):
        return []
    return [note for note in _items(notes)
            if "=" in note or re.search(r"\b(formula|equation)\b", note, re.IGNORECASE)][:3]


def _validation_coverage_rows(validation: dict[str, Any]) -> list[list[Any]]:
    fields = (
        ("Named result checks", "named_targets", "required_targets"),
        ("Primary source inputs", "primary_inputs", "required_primary_inputs"),
        ("Active formula members", "active_formula_members", "required_active_formula_members"),
        ("Formula identity matches", "active_formula_identity_matches", "required_active_formula_identities"),
        ("External input coordinates", "external_input_coordinates", "required_external_input_coordinates"),
        ("External formula-cut members", "external_formula_cut_members", "required_external_formula_cut_members"),
        ("External key-axis entries", "external_age_axis_keys", "required_external_age_axis_keys"),
    )
    rows = []
    for label, observed_key, required_key in fields:
        observed = validation.get(observed_key)
        if observed is None and observed_key == "active_formula_members":
            observed = validation.get("active_premium_formula_cells")
        if observed is None:
            continue
        required = validation.get(required_key)
        rows.append([label, observed, required if required is not None else "Not recorded"])
    return rows


def _mermaid_flow(modules: list[Any]) -> list[str]:
    valid = [module for module in modules if isinstance(module, dict)]
    valid.sort(key=lambda module: (module.get("order", 0), str(module.get("module_id", ""))))
    if not valid:
        return []
    fence = chr(96) * 3
    lines = [fence + "mermaid", "flowchart LR"]
    for index, module in enumerate(valid):
        label = _plain(module.get("module_id"), f"Module {index + 1}")
        label = re.sub(r'["\[\]{}|]', " ", label).strip()
        lines.append(f'  m{index}["{label}"]')
    if len(valid) > 1:
        lines.append("  " + " --> ".join(f"m{index}" for index in range(len(valid))))
    lines.append(fence)
    return lines


def _question_row(item: Any) -> list[Any]:
    if isinstance(item, dict):
        return [item.get("question_id"), item.get("question"), item.get("disposition")]
    return ["", item, ""]


def _report_markdown(
    report: dict[str, Any],
    *,
    workflow_root: Path | None = None,
    report_directory: Path | None = None,
) -> str:
    code = chr(96)
    source = report.get("source", {})
    design = report.get("design_summary", {})
    actual = design.get("actual_coverage", {})
    generation = report.get("generation_summary", {})
    validation = report.get("validation_summary", {})
    review_policy = report.get("review_policy", {})
    stage6_pair = review_policy.get("stage6_required_reviewer_pair", ["agent", "human"])
    stage6_review_text = ("Agent and delegated TypeSafe" if stage6_pair == ["agent", "typesafe"]
                          else "Agent and human")
    catalog_reference = design.get("input_catalog", {})
    catalog = catalog_reference.get("content", {}) if isinstance(catalog_reference, dict) else {}
    execution = design.get("execution_design", {})
    modules = design.get("shared_modules", [])
    scenarios = design.get("scenarios", [])
    targets = [item for item in design.get("targets", []) if isinstance(item, dict)]
    lines = ["# Model Conversion Report", "",
             "Report status: **Ready for review**", "",
             "## Executive summary", ""]

    target_rows = []
    target_results = report.get("target_results", {})
    if isinstance(target_results, dict):
        ordered_targets = sorted(targets, key=lambda item: (
            item.get("result_order") if isinstance(item.get("result_order"), int) else 10**9,
            str(item.get("selector", item.get("target_id", ""))).casefold(),
        ))
        emitted = set()
        for target in ordered_targets:
            name = target.get("selector", target.get("target_id"))
            target_id = target.get("target_id")
            result_name = name if name in target_results else target_id if target_id in target_results else None
            if result_name is None or result_name in emitted:
                continue
            record = target_results[result_name]
            if not isinstance(record, dict):
                continue
            emitted.add(result_name)
            target_rows.append(_result_row("Requested result", str(name), record, target))
        for target in ordered_targets:
            for diagnostic in _items(target.get("diagnostic_intermediates")):
                record = target_results.get(diagnostic)
                if diagnostic in emitted or not isinstance(record, dict):
                    continue
                emitted.add(diagnostic)
                target_rows.append(_result_row("Diagnostic intermediate", diagnostic, record))
        for name, record in sorted(target_results.items()):
            if name in emitted or not isinstance(record, dict):
                continue
            target_rows.append(_result_row("Additional Step 5 check", str(name), record))
    if target_rows:
        lines.extend(_table(["Classification", "Measure", "Converted model", "Workbook", "Difference", "Check"], target_rows))
    else:
        lines.append("No named result comparisons are recorded in the accepted Step 5 report.")
    matched = validation.get("named_targets")
    required = validation.get("required_targets")
    mismatch_count = report.get("mismatch_count")
    if matched is not None and required is not None:
        lines.extend(["", f"Step 5 compared {_plain(matched)} of {_plain(required)} named results; "
                      f"recorded mismatches: {_plain(mismatch_count)}."])
    scenario_label = ", ".join(_plain(item.get("scenario_id")) for item in scenarios
                               if isinstance(item, dict) and item.get("scenario_id")) or \
        ", ".join(_items(design.get("scenario_ids"))) or "the accepted saved scenario"
    verified_routes = _items(actual.get("runtime_verified_path_ids"))
    numerical_routes = _items(actual.get("implemented_path_ids"))
    if (isinstance(matched, int) and isinstance(required, int) and matched == required
            and mismatch_count == 0):
        conclusion = (f"**Scoped conclusion:** the generated model is reconciled for {scenario_label} "
                      "within the tolerances recorded below.")
    else:
        conclusion = (f"**Scoped conclusion:** the available Step 5 evidence for {scenario_label} "
                      "is summarized below; the recorded data do not establish a complete reconciliation.")
    generated_count = len(numerical_routes) if isinstance(actual.get("implemented_path_ids"), list) else "not recorded"
    verified_count = len(verified_routes) if isinstance(actual.get("runtime_verified_path_ids"), list) else "not recorded"
    lines.extend(["", conclusion + f" Step 4 generated {generated_count} numerical routes; "
                  f"Step 5 verified {verified_count} of them. This evidence supports the declared saved case only.", "",
                  "## Purpose/intended use/scope", "",
                  "This report hands over the generated implementation and its comparison with the accepted workbook run. "
                  "Use it to review the stated case, inspect the delivered bundle, and reproduce the comparison. "
                  "It does not establish actuarial certification, production deployment, or results for untested scenarios."])
    scope = design.get("scope", {})
    if isinstance(scope, dict):
        boundary = scope.get("boundary")
        if boundary:
            lines.extend(["", f"**Declared boundary:** {_cell(boundary)}"])
        included = _items(scope.get("included"))
        excluded = _items(scope.get("excluded"))
        if included:
            lines.extend(["", "**Included in the accepted design:**", ""] + [f"- {item}" for item in included] + [""])
        if excluded:
            lines.extend(["", "**Excluded from this result:**", ""] + [f"- {item}" for item in excluded] + [""])

    lines.extend(["", "## Source model and inputs/assumptions", ""])
    workbook_path = source.get("workbook_path")
    source_name = Path(workbook_path).name if isinstance(workbook_path, str) else "the accepted source workbook"
    lines.append(f"Source: {_link(source_name, workbook_path, workflow_root, report_directory)}. "
                 f"Run {code}{_cell(source.get('run_id'))}{code}; source ID {_cell(source.get('source_id'))}.")
    if targets:
        target_rows = []
        for target in sorted(targets, key=lambda item: item.get("result_order", 0)):
            shape = target.get("shape")
            shape_text = "scalar" if shape == [] else _human_value(shape)
            target_rows.append([target.get("selector", target.get("target_id")),
                                target.get("result_order"), shape_text,
                                target.get("units"), target.get("description")])
        lines.extend(["", "Requested outputs:", ""])
        lines.extend(_table(["Output", "Order", "Shape", "Units / basis", "Meaning"], target_rows))
    if isinstance(catalog, dict) and catalog:
        counts = catalog.get("counts", {})
        if isinstance(counts, dict):
            logical_count = counts.get("logical_input_object_count", sum(_catalog_kind_counts(catalog).values()))
            raw_count = counts.get("source_raw_variable_count", "not recorded")
            external_count = counts.get("formula_derived_external_variable_count", "not recorded")
            kind_counts = _catalog_kind_counts(catalog)
            kind_text = ", ".join(
                f"{kind_counts[kind]} {kind}{'' if kind_counts[kind] == 1 else 's'}"
                for kind in ("scalar", "vector", "table")
            )
            lines.extend(["", f"The verified input catalog records {_plain(logical_count)} logical input objects: "
                          f"{kind_text}. It distinguishes {_plain(raw_count)} source-raw variables "
                          f"from {_plain(external_count)} separately supplied formula-derived external variables. "
                          "A scalar stores one value, a vector is indexed by its declared axis, and a table preserves its "
                          "record and category axes."])
        tab_rows = _catalog_source_tab_rows(catalog)
        if tab_rows:
            lines.extend(["", "Distinct catalog variables by source tab:", ""])
            lines.extend(_table(["Source tab", "Raw variables", "Formula-derived external", "Total variables"], tab_rows))
        group_rows = _catalog_group_rows(catalog)
        if group_rows:
            lines.extend(["", "Input groups with recorded variables and data shapes:", ""])
            lines.extend(_table(["Catalog group", "Classification", "Objects", "Kinds, dimensions, and axes"], group_rows))
    else:
        lines.extend(["", "Input catalog facts are unavailable: the accepted design has no verified hash-bound catalog. "
                      "Counts, shapes, axes, and source-record detail are therefore not inferred."])
    scenario_input_count = sum(
        len(scenario.get("primary_inputs", {}))
        for scenario in scenarios
        if isinstance(scenario, dict) and isinstance(scenario.get("primary_inputs", {}), dict)
    )
    if scenario_input_count:
        lines.extend(["", f"The accepted design records {scenario_input_count} saved primary input literals. "
                      "Their source locations and values are listed in Appendix D."])
    if isinstance(catalog_reference, dict) and catalog_reference.get("available"):
        catalog_markdown = catalog_reference.get("markdown_artifact_path")
        catalog_link_text, catalog_link_path = (
            ("catalog Markdown", catalog_markdown) if isinstance(catalog_markdown, str) and catalog_markdown
            else ("catalog JSON", catalog_reference.get("artifact_path"))
        )
        lines.append(f"Input-boundary catalog: {_link(catalog_link_text, catalog_link_path, workflow_root, report_directory)}.")
    behavior_policy = scope.get("source_behavior_policy") if isinstance(scope, dict) else None
    if behavior_policy:
        lines.extend(["", f"**Source-behavior policy:** {_cell(behavior_policy)}"])

    lines.extend(["", "## Calculation design and time/module order", ""])
    ordering = execution.get("ordering") if isinstance(execution, dict) else None
    if ordering:
        lines.append(f"Accepted calculation order: {_cell(ordering)}")
    emission = execution.get("emission") if isinstance(execution, dict) else None
    if emission:
        lines.extend(["", _cell(emission)])
    input_policy = execution.get("input_policy") if isinstance(execution, dict) else None
    if input_policy:
        lines.extend(["", f"Input handling: {_cell(input_policy)}"])
    if isinstance(modules, list) and modules:
        lines.extend(["", "The flow follows the accepted Step 3 module order. Each responsibility below is recorded in that design; generation and reconciliation evidence are reported separately."])
        lines.extend(["", *_mermaid_flow(modules), "", "Module responsibilities:", ""])
        module_rows = [[item.get("order"), item.get("module_id"), item.get("purpose")]
                       for item in sorted((module for module in modules if isinstance(module, dict)),
                                          key=lambda module: (module.get("order", 0),
                                                              str(module.get("module_id", ""))))]
        lines.extend(_table(["Order", "Module", "Responsibility"], module_rows))
    else:
        lines.append("The accepted design does not provide a shared-module list; module order is unavailable.")
    field_group_rows = _field_group_rows(design)
    if field_group_rows:
        lines.extend(["", "Accepted field-group layouts and source extents (design organization, not runtime proof):", ""])
        lines.extend(_table(["Group", "Shape and time alignment", "Source extents"], field_group_rows))
    recurrence_notes = _recurrence_notes(execution if isinstance(execution, dict) else {})
    if recurrence_notes:
        lines.extend(["", "Time alignment and recurrence notes carried from the accepted design:", ""])
        lines.extend(f"- {_cell(note)}" for note in recurrence_notes)
        lines.append("")
    result_notes = _result_relationship_notes(execution if isinstance(execution, dict) else {})
    if result_notes:
        lines.extend(["", "**Result equation and reductions from the accepted design:**", ""])
        lines.extend(f"- {_cell(note)}" for note in result_notes)
        lines.append("")

    lines.extend(["", "## Conversion approach", ""])
    backend = execution.get("backend") if isinstance(execution, dict) else None
    if backend:
        backend_text = _cell(backend).rstrip(".")
        lines.append(f"Backend: {backend_text}.")
    function_count = generation.get("compiled_function_count")
    function_label = "compiled functions"
    if function_count is None:
        function_count = generation.get("unique_equation_family_count")
        function_label = "unique equation-family functions"
    if function_count is None:
        function_count = generation.get("formula_count")
        function_label = "generated formula functions"
    lines.append(f"The generated bundle contains {_plain(function_count)} {function_label}, "
                 f"{_plain(generation.get('source_family_count'))} source families, and "
                 f"{_plain(generation.get('semantic_equation_group_count', generation.get('unique_equation_family_count')))} "
                 "semantic equation groups.")
    if generation.get("model_variable_count") is not None or generation.get("derived_variable_count") is not None:
        lines.append(f"Generated model variables: {_plain(generation.get('model_variable_count'))}; "
                     f"derived variables: {_plain(generation.get('derived_variable_count'))}.")
    if generation.get("array_instance_count") is not None or generation.get("array_member_count") is not None:
        lines.append(f"Array instances: {_plain(generation.get('array_instance_count'))}; "
                     f"array members: {_plain(generation.get('array_member_count'))}.")
    dependencies = generation.get("runtime_dependencies")
    if isinstance(dependencies, list) and dependencies:
        lines.append("Runtime dependencies: " + "; ".join(_items(dependencies)) + ".")
    external = generation.get("external_capture")
    if isinstance(external, dict):
        ranges = external.get("ranges", [])
        axis = external.get("age_axis", {})
        axis_range = axis.get("source_range") if isinstance(axis, dict) else None
        lines.extend(["", f"Formula-derived external inputs were captured as {len(ranges) if isinstance(ranges, list) else 0} "
                      f"source ranges covering {_plain(external.get('formula_cut_member_count'))} formula-cut members. "
                      f"Declared key axis: {_plain(axis_range)}. Step 5 reports the recapture and comparison results."])
    elif "external_capture" in generation:
        lines.extend(["", "The accepted generation report records no formula-derived external capture."])
    else:
        lines.extend(["", "Formula-derived external capture status is not recorded in this report."])

    lines.extend(["", "## Validation and reconciliation results", ""])
    validation_context = report.get("validation_context", {})
    if isinstance(validation_context, dict):
        standalone = validation_context.get("standalone_validation")
        oracle = validation_context.get("oracle")
        method_lines = []
        if isinstance(standalone, dict):
            standalone_path = standalone.get("path")
            standalone_link = _link("standalone validation artifact", standalone_path,
                                    workflow_root, report_directory) if standalone_path else "standalone validation artifact"
            method_lines.append(f"Standalone Python run: {standalone_link} is recorded.")
        if "standalone_project_imports" in validation_context:
            method_lines.append(f"Project imports in the standalone run: {_plain(validation_context.get('standalone_project_imports'))}.")
        if "formula_cache_inputs" in validation_context:
            method_lines.append(f"Formula-cache inputs: {_plain(validation_context.get('formula_cache_inputs'))}.")
        if isinstance(oracle, dict):
            oracle_engine = oracle.get("engine")
            oracle_path = oracle.get("path")
            oracle_link = _link("native workbook oracle record", oracle_path,
                                workflow_root, report_directory) if oracle_path else "native workbook oracle"
            method_lines.append(f"Native workbook comparison: {_plain(oracle_engine)} was the recorded oracle; {oracle_link} is linked.")
            if "macros_executed" in oracle:
                method_lines.append(f"Macros executed for the oracle run: {_plain(oracle.get('macros_executed'))}.")
            overrides = oracle.get("input_overrides")
            if isinstance(overrides, list):
                method_lines.append("Oracle input overrides: none recorded." if not overrides
                                    else f"Oracle input overrides: {len(overrides)} recorded.")
            settings = oracle.get("engine_settings")
            if isinstance(settings, dict) and settings:
                run_conditions = []
                if settings.get("version") is not None:
                    run_conditions.append(f"Excel version {_plain(settings.get('version'))}")
                if isinstance(settings.get("iteration"), bool):
                    run_conditions.append(f"iteration {'enabled' if settings['iteration'] else 'disabled'}")
                elif settings.get("iteration") is not None:
                    run_conditions.append(f"iteration setting {_plain(settings.get('iteration'))}")
                if settings.get("max_iterations") is not None:
                    run_conditions.append(f"maximum {_plain(settings.get('max_iterations'))} iterations")
                if settings.get("max_change") is not None:
                    run_conditions.append(f"maximum change {_plain(settings.get('max_change'))}")
                if run_conditions:
                    method_lines.append("Recorded Excel run conditions: " + "; ".join(run_conditions) + ".")
        if method_lines:
            lines.extend(["**Validation method:** " + " ".join(method_lines)])
        else:
            lines.append("Validation method metadata is not recorded in the accepted Step 5 report.")
    coverage_rows = _validation_coverage_rows(validation)
    if coverage_rows:
        lines.extend(["", "Comparison coverage:", ""])
        lines.extend(_table(["Measure", "Observed", "Required"], coverage_rows))
        lines.append("")
    max_difference = validation.get("active_formula_max_abs_difference")
    if max_difference is not None:
        lines.append(f"Maximum absolute difference across active formula values: {_format_result(max_difference)}.")
    tolerances = report.get("tolerances", {})
    if isinstance(tolerances, dict):
        lines.append(f"Comparison tolerances: absolute {_plain(tolerances.get('absolute'))}; "
                     f"relative {_plain(tolerances.get('relative'))}. Mismatches: {_plain(mismatch_count)}.")
    if isinstance(actual, dict):
        generated_value = actual.get("implemented_path_ids")
        verified_value = actual.get("runtime_verified_path_ids")
        generated_ids = _items(generated_value)
        verified_ids = _items(verified_value)
        generated_summary = (f"{len(generated_ids)}" + (f" ({', '.join(generated_ids)})" if generated_ids else "")
                            if isinstance(generated_value, list) else "not recorded")
        verified_summary = (f"{len(verified_ids)}" + (f" ({', '.join(verified_ids)})" if verified_ids else "")
                            if isinstance(verified_value, list) else "not recorded")
        lines.extend(["", f"Numerical routes generated by Step 4: {generated_summary}.",
                      f"Numerical routes supported by Step 5 runtime evidence: {verified_summary}."])
        condition_ids = _items(actual.get("condition_only_path_ids"))
        inactive_ids = _items(actual.get("condition_assessed_inactive_ids"))
        if condition_ids:
            lines.append(f"Condition-only route checks assessed inactive: {len(inactive_ids)} of {len(condition_ids)}"
                         + (f" ({', '.join(inactive_ids)})." if inactive_ids else ".")
                         + " No feedback solver is included.")

    lines.extend(["", "## Differences/limitations/open decisions", "",
                  "The comparison above establishes behavior only for the accepted saved scenario and selected outputs. "
                  "A condition checked as inactive is not a solver implementation. The accepted business interpretation, "
                  "source scale, and basis remain limited to what the accepted target metadata and source artifacts state."])
    quality = report.get("quality", {})
    metrics = quality.get("metrics", {}) if isinstance(quality, dict) else {}
    parsed_parts = metrics.get("parsed_package_parts") if isinstance(metrics, dict) else None
    total_parts = metrics.get("parsed_package_parts_total", metrics.get("package_parts_total")) if isinstance(metrics, dict) else None
    opaque_parts = metrics.get("opaque_parts") if isinstance(metrics, dict) else None
    if isinstance(parsed_parts, int) and isinstance(total_parts, int) and parsed_parts < total_parts:
        extraction_note = f"Source extraction was partial: {parsed_parts} of {total_parts} package parts were parsed"
        if isinstance(opaque_parts, int):
            extraction_note += f" and {opaque_parts} are recorded as opaque"
        lines.append(extraction_note + ". The generated model and current reconciliation cover the declared conversion scope, not every workbook part.")
    unresolved_interpretations = []
    questions = design.get("open_questions", [])
    if isinstance(questions, list):
        for item in questions:
            if not isinstance(item, dict):
                continue
            question = item.get("question")
            disposition = item.get("disposition")
            searchable = f"{question or ''} {disposition or ''}".casefold()
            if any(term in searchable for term in ("interpretation", "actuarial", "business review", "business meaning")):
                unresolved_interpretations.append(item)
    for item in unresolved_interpretations:
        lines.append(f"**Unresolved business interpretation:** {_cell(item.get('question'))} "
                     f"Step 3 disposition: {_cell(item.get('disposition'))}")
    external = generation.get("external_capture")
    external_values = validation.get("external_input_coordinates")
    external_required = validation.get("required_external_input_coordinates")
    cut_values = validation.get("external_formula_cut_members")
    cut_required = validation.get("required_external_formula_cut_members")
    axis_values = validation.get("external_age_axis_keys")
    axis_required = validation.get("required_external_age_axis_keys")
    if (isinstance(external, dict) and isinstance(external_values, int) and isinstance(external_required, int)
            and external_values == external_required and isinstance(cut_values, int) and isinstance(cut_required, int)
            and cut_values == cut_required and isinstance(axis_values, int) and isinstance(axis_required, int)
            and axis_values == axis_required):
        lines.append("The design-time missing external-input capture prerequisite was completed for this saved run, "
                     "with the captured coordinates, formula-cut members, and key-axis entries compared in Step 5.")
    options = [item for item in design.get("options", []) if isinstance(item, dict)]
    selected_names = [str(item.get("name", item.get("scope", item.get("option_id"))))
                      for item in options if _option_selected(item)]
    if selected_names:
        lines.append("Selected design option: " + "; ".join(selected_names) + ". Other options remain planning history.")
    lines.append("The original Step 3 questions remain in the machine report and Appendix C. Later capture and native comparison evidence is summarized above; reconciliation does not by itself resolve business interpretation.")
    lines.append("This report does not claim all-configuration closure, full-workbook conversion, a feedback solver, GPU performance, or production deployment.")

    lines.extend(["", "## Handover/run instructions", ""])
    bundle = generation.get("bundle", {})
    bundle_path = bundle.get("path") if isinstance(bundle, dict) else None
    bundle_files = bundle.get("files", {}) if isinstance(bundle, dict) else {}
    if isinstance(bundle_path, str):
        lines.append(f"Generated bundle: {_link('open bundle', bundle_path, workflow_root, report_directory)}.")
        if isinstance(bundle_files, dict):
            for filename, label in (("model.py", "Runnable model"),
                                    ("model_manifest.json", "Model manifest"),
                                    ("source_map.json", "Source map")):
                if filename in bundle_files:
                    model_path = str(Path(bundle_path) / filename)
                    lines.append(f"{label}: {_link(filename, model_path, workflow_root, report_directory)}.")
    stage_links = []
    for number, label in ((3, "Accepted design"), (4, "Generation report"), (5, "Validation report")):
        stage = report.get("stages", {}).get(str(number), {})
        if isinstance(stage, dict):
            stage_path = stage.get("markdown_path") or stage.get("json_path")
            if stage_path:
                stage_links.append(_link(label, stage_path, workflow_root, report_directory))
    if stage_links:
        lines.append("Supporting reports: " + "; ".join(stage_links) + ".")
    run_command = generation.get("run_command")
    if run_command:
        lines.extend(["", "From the generated bundle directory, run:", "", chr(96) * 3 + "sh",
                      str(run_command), chr(96) * 3])
    lines.extend(["", "## Acceptance", "",
                  f"This report is generated ready for review. Step 6 requires {stage6_review_text} approval of this exact JSON/Markdown pair. "
                  "Later acceptance receipts are recorded in the workflow ledger; do not edit this report after acceptance. "
                  "The earlier stages and their review records are listed in Appendix B."])
    if workflow_root is not None:
        ledger_path = workflow_root / "workflow.json"
        lines.append(f"Current workflow acceptance ledger: {_link('workflow.json', str(ledger_path), workflow_root, report_directory)}.")

    lines.extend(["", "## Appendices", "", "### A. Detailed evidence and artifact hashes", "",
                  "These identifiers and hashes support audit and repeatability; the main report above is written for review of the result."])
    stage_rows = []
    for stage, record in sorted(report.get("stages", {}).items(), key=lambda item: int(item[0])):
        stage_rows.append([
            f"Step {stage}", record.get("revision"), record.get("status"),
            _link("JSON", record.get("json_path"), workflow_root, report_directory),
            record.get("json_sha256"),
            _link("Markdown", record.get("markdown_path"), workflow_root, report_directory),
            record.get("markdown_sha256"),
        ])
    if stage_rows:
        lines.extend(["", "Accepted upstream artifacts:", ""])
        lines.extend(_table(["Stage", "Revision", "Status", "JSON", "JSON SHA-256",
                             "Markdown", "Markdown SHA-256"], stage_rows))
    validation_artifact_rows = []
    validation_context = report.get("validation_context", {})
    if isinstance(validation_context, dict):
        for key, label in (("standalone_validation", "Standalone Python validation"),
                           ("oracle", "Native workbook oracle")):
            record = validation_context.get(key)
            if isinstance(record, dict) and record.get("path"):
                validation_artifact_rows.append([
                    label,
                    _link(Path(str(record.get("path"))).name, record.get("path"), workflow_root, report_directory),
                    record.get("sha256"),
                ])
    if validation_artifact_rows:
        lines.extend(["", "Step 5 method artifacts:", ""])
        lines.extend(_table(["Artifact", "File", "SHA-256"], validation_artifact_rows))
    oracle = validation_context.get("oracle") if isinstance(validation_context, dict) else None
    engine_settings = oracle.get("engine_settings") if isinstance(oracle, dict) else None
    if isinstance(engine_settings, dict) and engine_settings:
        setting_rows = [[name.replace("_", " "), value] for name, value in sorted(engine_settings.items())
                        if isinstance(value, (str, int, float, bool))]
        if setting_rows:
            lines.extend(["", "Recorded Excel engine settings:", ""])
            lines.extend(_table(["Setting", "Value"], setting_rows))
    if isinstance(catalog_reference, dict) and catalog_reference.get("available"):
        lines.extend(["", f"Input catalog: {_link('hash-bound catalog JSON', catalog_reference.get('artifact_path'), workflow_root, report_directory)}; "
                      f"SHA-256 {catalog_reference.get('artifact_sha256')}."])
    if isinstance(bundle, dict) and isinstance(bundle_files, dict):
        bundle_rows = [[name, digest] for name, digest in sorted(bundle_files.items())]
        if bundle_rows:
            lines.extend(["", "Generated bundle files:", ""])
            lines.extend(_table(["File", "SHA-256"], bundle_rows))
    implementation_rows = []
    for item in actual.get("implementation_evidence", []) if isinstance(actual, dict) else []:
        if isinstance(item, dict):
            implementation_rows.append([item.get("path_id"), item.get("address"), item.get("source_formula")])
    if implementation_rows:
        lines.extend(["", "Step 4 implementation evidence:", ""])
        lines.extend(_table(["Route", "Source location", "Source formula"], implementation_rows))
    runtime_rows = []
    for item in actual.get("runtime_evidence", []) if isinstance(actual, dict) else []:
        if isinstance(item, dict):
            comparison = item.get("comparison", {})
            kind = comparison.get("kind") if isinstance(comparison, dict) else None
            runtime_rows.append([item.get("path_id"), item.get("address"), item.get("source_formula"), kind])
    if runtime_rows:
        lines.extend(["", "Step 5 native comparison evidence:", ""])
        lines.extend(_table(["Route", "Source location", "Source formula", "Comparison"], runtime_rows))
    condition_rows = []
    for item in actual.get("condition_evidence", []) if isinstance(actual, dict) else []:
        if isinstance(item, dict):
            condition_rows.append([item.get("path_id"), item.get("status"), item.get("evidence")])
    if condition_rows:
        lines.extend(["", "Step 5 condition-only evidence:", ""])
        lines.extend(_table(["Route", "Assessment", "Evidence"], condition_rows))
    if isinstance(source, dict) and source.get("workbook_sha256"):
        lines.extend(["", f"Source workbook SHA-256: {source.get('workbook_sha256')}."])
    quality = report.get("quality", {})
    metrics = quality.get("metrics", {}) if isinstance(quality, dict) else {}
    if isinstance(metrics, dict) and metrics:
        quality_rows = [[name.replace("_", " "), value] for name, value in sorted(metrics.items())
                        if isinstance(value, (str, int, float, bool))]
        if quality_rows:
            lines.extend(["", "Step 1 source-quality metrics:", ""])
            lines.extend(_table(["Measure", "Recorded value"], quality_rows))

    lines.extend(["", "### B. Review history", ""])
    review_rows = []
    for stage, record in sorted(report.get("stages", {}).items(), key=lambda item: int(item[0])):
        for reviewer in ("agent", "human", "typesafe"):
            receipt = record.get(f"{reviewer}_receipt")
            if not isinstance(receipt, dict):
                continue
            review_rows.append([f"Step {stage}", reviewer.title(), receipt.get("decision"),
                                receipt.get("recorded_at"), receipt.get("message")])
    if review_rows:
        lines.extend(_table(["Stage", "Reviewer", "Decision", "Recorded", "Review note"], review_rows))
    else:
        lines.append("No upstream review receipts are recorded.")
    delegation = review_policy.get("delegation")
    if isinstance(delegation, dict):
        lines.extend(["", f"Delegation record: stages {_human_value(delegation.get('stages'))}; "
                      f"requested model {_plain(delegation.get('requested_model'))}; "
                      f"minimum approval probability {_plain(delegation.get('minimum_approval_probability'))}; "
                      f"authorization SHA-256 {delegation.get('authorization_sha256')}."])

    lines.extend(["", "### C. Prior option and question history", "",
                  "Options and questions below preserve Step 3 planning history; they do not replace the current execution and reconciliation findings above."])
    option_rows = []
    for option in options:
        actual_option = option.get("actual_numeric_coverage")
        actual_text = "Planning only"
        if isinstance(actual_option, dict):
            actual_text = (f"{len(_items(actual_option.get('runtime_verified_ids')))} of "
                           f"{len(_items(actual_option.get('planned_ids')))} numerical routes verified")
        option_rows.append([option.get("option_id"), option.get("name", option.get("scope")),
                            "Selected" if _option_selected(option) else "Not selected",
                            option.get("status"), actual_text,
                            "; ".join(_items(option.get("exclusions"))) or "None recorded"])
    if option_rows:
        lines.extend(["", "Option history:", ""])
        lines.extend(_table(["Option", "Description", "Selection", "Step 3 status",
                             "Execution evidence", "Exclusions"], option_rows))
    questions = design.get("open_questions", [])
    question_rows = [_question_row(item) for item in questions] if isinstance(questions, list) else []
    if question_rows:
        lines.extend(["", "Questions preserved verbatim with their original design-time disposition:", ""])
        lines.extend(_table(["Question ID", "Question", "Original disposition"], question_rows))
    preserved = execution.get("preserve_formulas", []) if isinstance(execution, dict) else []
    if isinstance(preserved, list) and preserved:
        lines.extend(["", "Formula-behavior notes carried forward from Step 3:", ""])
        lines.extend(f"- {item}" for item in _items(preserved))

    lines.extend(["", "### D. Input-boundary source records", ""])
    scenario_rows = []
    for scenario in scenarios:
        if not isinstance(scenario, dict):
            continue
        scenario_id = scenario.get("scenario_id")
        primary_inputs = scenario.get("primary_inputs", {})
        if not isinstance(primary_inputs, dict):
            continue
        for name, details in sorted(primary_inputs.items()):
            if isinstance(details, dict):
                scenario_rows.append([scenario_id, name, details.get("source_cell"),
                                      _human_value(details.get("source_literal"))])
            else:
                scenario_rows.append([scenario_id, name, "", _human_value(details)])
    if scenario_rows:
        lines.extend(["Saved scenario literals recorded by the accepted design:", ""])
        lines.extend(_table(["Scenario", "Input", "Source location", "Saved literal"], scenario_rows))
        lines.append("")
    source_records = catalog.get("source_records", []) if isinstance(catalog, dict) else []
    source_record_rows = []
    if isinstance(source_records, list):
        for record in source_records:
            if not isinstance(record, dict):
                continue
            extents = record.get("source_extents", [])
            extent_text = ", ".join(
                f"{_plain(extent.get('sheet'))}!{_plain(extent.get('range'))}"
                for extent in extents if isinstance(extent, dict)
            ) if isinstance(extents, list) else ""
            provenance = record.get("provenance", {})
            description = provenance.get("description") if isinstance(provenance, dict) else None
            source_record_rows.append([record.get("record_id"), record.get("classification"),
                                       record.get("group_id"), extent_text, description])
    if source_record_rows:
        lines.extend(["The descriptions below are copied from the hash-bound input-boundary catalog and retain its source-record context.", ""])
        lines.extend(_table(["Record", "Class", "Group", "Source extent", "Catalog description"], source_record_rows))
    else:
        lines.append("No verified input-boundary source records are available.")

    lines.extend(["", "### E. Report organization references", "",
                  "The section order follows common model-conversion reporting practice: state purpose and scope, explain the design and handover, then show testing and reconciliation. These references inform presentation only; this report makes no claim of professional-standard compliance."])
    lines.extend([
        "",
        "- [Society of Actuaries, All You Need to Know About a Successful Model Conversion](https://www.soa.org/digital-publishing-platform/emerging-topics/all-you-need-to-know-about/) — lifecycle view across requirements, design, testing, and handover.",
        "- [Actuarial Standards Board, ASOP No. 56: Modeling](https://www.actuarialstandardsboard.org/asops/modeling-3/) — discussion areas include purpose, model structure, input reliance, testing, limitations, and documentation.",
        "- [Milliman, Confidence in numbers: Validating actuarial results](https://www.milliman.com/en/insight/confidence-in-numbers-validating-actuarial-results) — practical emphasis on defined outputs, comparison criteria, and documented results.",
    ])
    return "\n".join(lines)


def tool_catalog() -> dict[str, Any]:
    from excel_to_act.steps.step6.tools import tool_catalog as catalog

    return catalog()
