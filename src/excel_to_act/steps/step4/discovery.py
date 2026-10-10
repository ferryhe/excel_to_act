"""Discover the formula-text dependency closure for an approved Stage 3 target set."""

from __future__ import annotations

import json
from pathlib import Path
import os
import shutil
from typing import Any

from openpyxl import load_workbook

from excel_to_act.steps.conversion_workflow import hash_file, load_workflow, read_json, require_stage_approved
from excel_to_act.steps.step3.calculation import CellRange, FormulaEvaluator, _json_value, _parse_reference
from excel_to_act.steps.step4.external_inputs import (
    _addresses,
    load_current_external_bundle,
)


def _write(path: Path, payload: bytes) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_bytes(payload)
    os.replace(temporary, path)


def _cell_key(address: str) -> str:
    reference = _parse_reference(address, "")
    if reference is None or not reference.single or not reference.sheet:
        raise ValueError(f"trace member is not a qualified source cell: {address}")
    from openpyxl.utils.cell import get_column_letter
    return f"{reference.sheet.casefold()}!{get_column_letter(reference.min_col)}{reference.min_row}"


def _record_external_override(overrides: dict[str, dict[str, Any]], address: str,
                              record: dict[str, Any]) -> None:
    """Store an approved formula cut using the same key form as trace membership."""
    key = _cell_key(address)
    if key in overrides:
        raise ValueError(f"captured external formula member is duplicated: {address}")
    overrides[key] = record


def _declared_source_addresses(context: dict[str, Any]) -> set[str]:
    """Expand only source literals declared by the accepted catalog and semantic plan."""
    result: set[str] = set()

    def add_extent(record: Any, *, default_sheet: str | None = None) -> None:
        if not isinstance(record, dict):
            return
        if isinstance(record.get("ref"), str):
            extent = record["ref"]
        elif isinstance(record.get("range"), str):
            extent = f"{record.get('sheet', default_sheet)}!{record['range']}"
        else:
            return
        parsed = _parse_reference(extent, "")
        if parsed is None or not parsed.sheet:
            raise ValueError(f"accepted source extent is not a finite qualified range: {extent}")
        result.update(_cell_key(address) for address in _addresses(extent))

    catalog = context["catalog"]
    for record in catalog.get("variables", []):
        if isinstance(record, dict) and record.get("role") in {"raw_business_input", "formula_derived_external"}:
            for extent in record.get("source_extents", []):
                add_extent(extent)
    for field in ("axis_metadata", "source_records"):
        for record in catalog.get(field, []):
            if isinstance(record, dict):
                for extent in record.get("source_extents", []):
                    add_extent(extent)
    for record in context["semantic_plan"].get("variables", {}).values():
        if isinstance(record, dict) and record.get("role") in {"raw", "external"}:
            for extent in record.get("source_extents", []):
                add_extent(extent)
        initial = record.get("initial_condition") if isinstance(record, dict) else None
        if isinstance(initial, dict) and isinstance(initial.get("seed_ref"), str):
            add_extent({"ref": initial["seed_ref"]})
    for record in catalog.get("source_records", []):
        if isinstance(record, dict) and record.get("classification") in {"adapter_metadata", "source_constant"}:
            for extent in record.get("source_extents", []):
                add_extent(extent)
    return result


def _approved_stored_empty_addresses(source_path: Path, source_hash: str,
                                     model_scope: dict[str, Any],
                                     candidate_trace: dict[str, Any]) -> set[str]:
    """Accept only traced single-cell blanks explicitly fenced by the approved scope."""
    scoped_cells: set[str] = set()
    for region in model_scope.get("allowed_regions", []):
        if not isinstance(region, dict) or not isinstance(region.get("source_ref"), str):
            continue
        reference = _parse_reference(region["source_ref"], "")
        if reference is not None and reference.single and reference.sheet:
            scoped_cells.add(_cell_key(region["source_ref"]))

    traced_candidates: set[str] = set()
    for unknown in candidate_trace.get("unknowns", []):
        if (not isinstance(unknown, dict)
                or unknown.get("category") != "unrepresented_source_coordinate"
                or not isinstance(unknown.get("evidence"), str)):
            continue
        try:
            key = _cell_key(unknown["evidence"])
        except ValueError:
            continue
        if key in scoped_cells:
            traced_candidates.add(key)
    if not traced_candidates:
        return set()

    before = hash_file(source_path)
    if before != source_hash:
        raise ValueError("source workbook changed before stored-empty adapter verification")
    workbook = load_workbook(source_path, data_only=False, read_only=False, keep_vba=True)
    verified: set[str] = set()
    try:
        sheet_titles = {sheet.title.casefold(): sheet.title for sheet in workbook.worksheets}
        for address in sorted(traced_candidates):
            sheet_key, local = address.split("!", 1)
            title = sheet_titles.get(sheet_key)
            if title is None:
                raise ValueError(f"approved stored-empty source sheet is missing: {sheet_key}")
            reference = _parse_reference(address, "")
            if reference is None:
                raise ValueError(f"approved stored-empty source coordinate is malformed: {address}")
            cell = workbook[title]._cells.get((reference.min_row, reference.min_col))
            if cell is None or cell.value is not None or cell.data_type == "f":
                raise ValueError(f"approved adapter blank is not a stored, formula-free blank: {address}")
            verified.add(address)
    finally:
        workbook.close()
    if hash_file(source_path) != source_hash:
        raise ValueError("source workbook changed during stored-empty adapter verification")
    return verified


def discover_active_trace(workflow_dir: Path) -> dict[str, Any]:
    """Evaluate approved targets from formula text and raw values; do not invoke Excel."""
    evaluator: FormulaEvaluator | None = None
    out: Path | None = None
    created_out = False
    stage4_root: Path | None = None
    try:
        root, manifest = load_workflow(workflow_dir)
        _workflow, stage3_revision = require_stage_approved(root, 3)
        stage3_path = root / stage3_revision["artifact"]["json"]
        stage3 = read_json(stage3_path)
        design = stage3.get("design", {})
        source = manifest["source"]
        source_path = Path(source["workbook_path"]).expanduser().resolve()
        source_hash = source["workbook_sha256"]
        if not source_path.is_file() or hash_file(source_path) != source_hash:
            raise ValueError("bound source workbook is missing or changed")
        stage4_root = (root / "stage4").resolve()
        if stage4_root.is_relative_to(source_path.parent) or source_path.is_relative_to(stage4_root):
            raise ValueError("Step 4 discovery output must stay outside the source workbook directory")
        targets = design.get("targets", [])
        if not targets:
            raise ValueError("approved Stage 3 design has no declared targets")

        external_context = None
        external_capture = None
        external_capture_path = None
        external_capture_file_hash = None
        candidate_trace_path = None
        candidate_trace_sha = None
        candidate_profile_path = None
        candidate_profile_sha = None
        candidate_formulae: dict[str, str] = {}
        external_overrides: dict[str, dict[str, Any]] = {}
        allowed_source_values: set[str] = set()
        if stage3.get("design", {}).get("semantic_mapping_required") is True:
            external_context, external_capture, external_capture_path, external_capture_file_hash = (
                load_current_external_bundle(root))
            if external_context["stage3_revision"]["artifact"]["json_sha256"] != stage3_revision["artifact"]["json_sha256"]:
                raise ValueError("external capture context differs from the current approved Stage 3 design")
            candidate_trace = external_context["candidate_trace"]
            candidate_profile_path, candidate_profile_sha = external_context["evidence_paths"]["candidate_profile"]
            candidate_trace_path, candidate_trace_sha = external_context["evidence_paths"]["candidate_trace"]
            if (candidate_trace.get("source_sha256") != source_hash
                    or hash_file(candidate_trace_path) != candidate_trace_sha
                    or hash_file(candidate_profile_path) != candidate_profile_sha):
                raise ValueError("approved static candidate trace/profile is stale or bound to another source")
            for record in candidate_trace.get("cells", []):
                if not isinstance(record, dict) or record.get("role") != "calculated_formula":
                    continue
                address, formula = record.get("address"), record.get("formula")
                if not isinstance(address, str) or not isinstance(formula, str):
                    raise ValueError("static candidate trace contains a malformed formula member")
                key = _cell_key(address)
                if key in candidate_formulae:
                    raise ValueError(f"static candidate trace duplicates a formula member: {address}")
                candidate_formulae[key] = formula
            if not candidate_formulae:
                raise ValueError("approved static candidate trace contains no formula members")

            formula_by_address = {_cell_key(item["address"]): item
                                  for item in external_capture["formula_cut_members"]}
            for boundary in external_context["boundaries"]:
                variable_id = boundary["boundary_variable_id"]
                values = external_capture["external_values"].get(variable_id)
                addresses = _addresses(boundary["source_range"])
                if not isinstance(values, list) or len(values) != len(addresses):
                    raise ValueError(f"captured external input has an invalid vector shape: {variable_id}")
                for index, address in enumerate(addresses):
                    key = _cell_key(address)
                    formula_record = formula_by_address.get(key)
                    if formula_record is None:
                        raise ValueError(f"captured external formula member is missing: {address}")
                    _record_external_override(external_overrides, address, {
                        "boundary_variable_id": variable_id,
                        "value": values[index],
                        "source_formula": formula_record["formula"],
                        "source_formula_sha256": formula_record["formula_sha256"],
                        "capture_artifact_sha256": external_capture_file_hash,
                    })
            allowed_source_values = _declared_source_addresses(external_context)
            model_scope = read_json(external_context["evidence_paths"]["model_scope"][0])
            allowed_source_values.update(_approved_stored_empty_addresses(
                source_path, source_hash, model_scope, candidate_trace))

        evaluator = FormulaEvaluator(source_path, external_overrides=external_overrides)
        target_values: dict[str, Any] = {}
        requested_cells = 0
        for item in targets:
            selector = item.get("selector")
            if not isinstance(selector, str) or not selector:
                raise ValueError("each approved target must have a non-empty selector")
            context = item.get("sheet_context", item.get("sheet", "Main"))
            reference = _parse_reference(selector, context)
            resolved: Any = reference if reference is not None else evaluator._name(selector, context)
            if isinstance(resolved, CellRange):
                size = ((resolved.max_row - resolved.min_row + 1)
                        * (resolved.max_col - resolved.min_col + 1))
                requested_cells += size
                if requested_cells > 50_000:
                    raise ValueError("approved target ranges exceed the 50,000-cell Step 4 discovery limit")
                rows = [[evaluator.evaluate_cell(resolved.sheet, row, col, dependency=False)
                         for col in range(resolved.min_col, resolved.max_col + 1)]
                        for row in range(resolved.min_row, resolved.max_row + 1)]
                value: Any = rows[0][0] if resolved.single else rows
            else:
                value = resolved
            target_id = item.get("target_id")
            if not isinstance(target_id, str) or not target_id:
                raise ValueError("each approved target must have a target_id")
            target_values[target_id] = {"selector": selector, "result_order": item.get("result_order"),
                                        "shape": item.get("shape"), "value": _json_value(value)}

        if hash_file(source_path) != source_hash:
            raise ValueError("source workbook changed during Step 4 discovery")
        formula_cells = [item for item in evaluator.cells.values()
                         if item.get("role") in {"calculated_formula", "calculated_array_formula"}]
        external_cells = [item for item in evaluator.cells.values()
                          if item.get("role") == "external_boundary_input"]
        source_cells = [item for item in evaluator.cells.values() if item.get("role") == "source_value"]
        if external_context is not None:
            candidate_keys = set(candidate_formulae)
            for item in formula_cells:
                key = _cell_key(item["address"])
                if key not in candidate_keys:
                    raise ValueError(f"active formula is outside the approved static candidate set: {item['address']}")
                if item.get("formula") != candidate_formulae[key]:
                    raise ValueError(f"active formula text differs from its approved candidate: {item['address']}")
            unbound_raw = [_cell_key(item["address"]) for item in source_cells
                           if _cell_key(item["address"]) not in allowed_source_values]
            if unbound_raw:
                raise ValueError("active source reads are outside the accepted raw or adapter metadata boundary: "
                                 + ", ".join(unbound_raw[:8]))
        if set(_cell_key(item["address"]) for item in external_cells) - set(external_overrides):
            raise ValueError("active external boundary read is not in the approved five-vector capture")

        trace = {"schema_version": "step4.discovery_trace.v1", "source_path": str(source_path),
                 "source_sha256": source_hash,
                 "design_sha256": stage3_revision["artifact"]["json_sha256"],
                 "external_capture_path": str(external_capture_path) if external_capture_path else None,
                 "external_capture_sha256": external_capture_file_hash,
                 "candidate_trace_path": str(candidate_trace_path) if candidate_trace_path else None,
                 "candidate_trace_sha256": candidate_trace_sha,
                 "candidate_profile_path": str(candidate_profile_path) if candidate_profile_path else None,
                 "candidate_profile_sha256": candidate_profile_sha,
                 "cells": sorted(evaluator.cells.values(), key=lambda item: item["address"]),
                 "discovery_order": list(evaluator.cells),
                 "completion_order": [evaluator._address(evaluator.sheet_names[key[0]], key[1], key[2])
                                      for key in evaluator.cache],
                 "edges": sorted(evaluator.edges.values(), key=lambda item: (item["consumer"], item["prerequisite"])),
                 "names": sorted(evaluator.names.values(), key=lambda item: (item["scope"], item["name"])),
                 "lookups": evaluator.lookups, "branches": evaluator.branches,
                 "dynamic_ranges": evaluator.dynamic_ranges,
                 "source_error_values": [item for item in evaluator.cells.values()
                                         if isinstance(item.get("value"), dict) and "excel_error" in item["value"]],
                 "targets": target_values, "formula_cache_inputs": False,
                 "native_excel_called": False, "external_values_from_native_capture": external_context is not None,
                 "unresolved_frontier_count": 0,
                 "candidate_formula_member_count": len(candidate_formulae) if candidate_formulae else None,
                 "active_formula_member_count": len(formula_cells),
                 "external_boundary_cell_count": len(external_cells),
                 "raw_source_cell_count": len(source_cells)}
        if not trace["cells"]:
            raise ValueError("approved targets did not reach any source cells")
        discovery_root = stage4_root / "discovery"
        revisions = [path for path in discovery_root.glob("revision-*") if path.is_dir()] if discovery_root.exists() else []
        revision_number = max((int(path.name.removeprefix("revision-")) for path in revisions
                               if path.name.removeprefix("revision-").isdigit()), default=0) + 1
        out = discovery_root / f"revision-{revision_number:04d}"
        if not out.resolve().is_relative_to(stage4_root):
            raise ValueError("Step 4 discovery output resolves outside this workflow's Stage 4 folder")
        if out.exists():
            raise ValueError("Step 4 discovery output already exists; preserve it and retry with a clean workflow")
        out.mkdir(parents=True)
        created_out = True
        trace_path = out / "active_trace.json"
        trace_bytes = (json.dumps(trace, indent=2, sort_keys=True, ensure_ascii=False, default=str) + "\n").encode("utf-8")
        _write(trace_path, trace_bytes)
        trace_hash = hash_file(trace_path)
        payload = {"schema_version": "step4.discovery.v1", "tool": "step4.discover", "status": "pass",
                   "workflow": str(root), "source_sha256": source_hash,
                   "stage3_revision": stage3_revision["revision"],
                   "stage3_artifact_sha256": stage3_revision["artifact"]["json_sha256"],
                   "active_trace_path": str(trace_path), "active_trace_sha256": trace_hash,
                   "target_values": target_values,
                 "counts": {"cells": len(trace["cells"]),
                              "formula_cells": len(formula_cells),
                              "candidate_formula_member_count": trace["candidate_formula_member_count"],
                              "external_boundary_cells": len(external_cells),
                              "raw_source_cells": len(source_cells),
                              "dependency_edges": len(trace["edges"]),
                              "names": len(trace["names"]), "lookups": len(trace["lookups"]),
                              "branches": len(trace["branches"]),
                              "dynamic_ranges": len(trace["dynamic_ranges"]),
                              "source_errors": len(trace["source_error_values"])},
                   "formula_cache_inputs": False, "native_excel_called": False}
        summary_path = out / "discovery.json"
        _write(summary_path, (json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n").encode("utf-8"))
        rows = ["# Step 4 Active-Trace Discovery", "", "Status: **Pass**", "",
                f"Source SHA-256: `{source_hash}`",
                f"Approved Stage 3 revision/hash: {stage3_revision['revision']} / `{payload['stage3_artifact_sha256']}`", "",
                f"Reached cells: {payload['counts']['cells']}; active generated formulas: {payload['counts']['formula_cells']}; external boundary inputs: {payload['counts']['external_boundary_cells']}; raw source cells: {payload['counts']['raw_source_cells']}; edges: {payload['counts']['dependency_edges']}.",
                f"Names: {payload['counts']['names']}; lookups: {payload['counts']['lookups']}; branches: {payload['counts']['branches']}; dynamic ranges: {payload['counts']['dynamic_ranges']}.",
                "The evaluator used formula text and raw workbook values; saved formula caches were not read.",
                "Discovery used formula text and raw literals only. External vectors came from the separately bound native FullRebuild capture; discovery itself did not call Excel.",
                (f"The active formula set is a strict subset of {payload['counts']['candidate_formula_member_count']} reviewed candidates; no formula caches were read."
                 if payload["counts"]["candidate_formula_member_count"] is not None
                 else "This legacy-compatible trace has no separately bound static candidate profile; no formula caches were read."),
                "This trace does not prove Excel equivalence.", "",
                "## Requested target results", ""]
        rows.extend(f"- `{target_id}` · `{record['selector']}` · `{json.dumps(record['value'], ensure_ascii=False, sort_keys=True)}`"
                    for target_id, record in target_values.items())
        rows.extend(["", "## Source errors", ""])
        rows.extend(f"- `{record['address']}`: `{record['value'].get('excel_error')}`"
                    for record in trace["source_error_values"])
        if not trace["source_error_values"]:
            rows.append("- None reached.")
        _write(out / "discovery.md", ("\n".join(rows) + "\n").encode("utf-8"))
        return {"schema_version": payload["schema_version"], "tool": payload["tool"], "status": "pass",
                "workflow": str(root), "discovery": str(summary_path),
                "active_trace": str(trace_path), "active_trace_sha256": trace_hash,
                "counts": payload["counts"], "target_values": target_values,
                "formula_cache_inputs": False, "native_excel_called": False}
    except Exception as exc:
        if (created_out and out is not None and out.exists()
                and stage4_root is not None and out.resolve().is_relative_to(stage4_root)):
            shutil.rmtree(out, ignore_errors=True)
        return {"schema_version": "step4.discovery.v1", "tool": "step4.discover",
                "status": "blocked", "reason": str(exc)}
    finally:
        if evaluator is not None:
            evaluator.close()
