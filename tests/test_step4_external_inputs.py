from __future__ import annotations

import hashlib
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps.conversion_workflow import append_stage_artifact, create_workflow, record_decision
from excel_to_act.steps.step4 import external_inputs


_EXTENTS = {
    "external.ci_rate_by_age": "Qtable!W5:W110",
    "external.multiple_ci_rate_by_age": "Qtable!Y5:Y110",
    "external.minor_ci_rate_by_age": "Qtable!Z5:Z110",
    "external.special_ci_rate_by_age": "Qtable!AD5:AD110",
    "external.moderate_ci_rate_by_age": "Qtable!AH5:AH110",
}


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _json(path: Path, value: dict) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True), encoding="utf-8")
    return _sha(path)


def _external_workflow(tmp_path: Path, extents: dict[str, str] | None = None,
                       age_axis: str = "Qtable!C5:C110") -> tuple[Path, Path]:
    extents = _EXTENTS if extents is None else extents
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    source_path = source_dir / "model.xlsx"
    workbook = Workbook()
    main = workbook.active
    main.title = "Main"
    sheets = {main.title: main}
    for address in [age_axis, *extents.values()]:
        sheet_name = address.split("!", 1)[0]
        if sheet_name not in sheets:
            sheets[sheet_name] = workbook.create_sheet(sheet_name)
    axis_sheet, axis_local = age_axis.split("!", 1)
    axis_min_col, axis_min_row, axis_max_col, axis_max_row = range_boundaries(axis_local)
    if axis_min_col != axis_max_col:
        raise ValueError("fixture external key axis must be vertical")
    for offset, row in enumerate(range(axis_min_row, axis_max_row + 1)):
        sheets[axis_sheet].cell(row, axis_min_col, offset)
    for extent in extents.values():
        sheet, local = extent.split("!", 1)
        min_col, min_row, max_col, max_row = range_boundaries(local)
        if min_col != max_col:
            raise ValueError("fixture external extent must be vertical")
        for row in range(min_row, max_row + 1):
            sheets[sheet].cell(row, min_col, "=UNSUPPORTED()")
    premium = workbook.create_sheet("Premium")
    premium["J1"] = "=1"
    workbook.defined_names.add(__import__("openpyxl").workbook.defined_name.DefinedName(
        "GP", attr_text="'Premium'!$J$1"))
    workbook.save(source_path)
    workbook.close()
    source_hash = _sha(source_path)

    workflow = tmp_path / "workflow"
    create_workflow(workflow, {"workbook_path": str(source_path), "workbook_sha256": source_hash,
                               "source_id": "fixture-source", "run_id": "fixture-run"})
    for stage in (1, 2):
        append_stage_artifact(workflow, stage, "checkpoint.json", {"status": "ready_for_review"},
                              f"Synthetic Stage {stage}.")
        for reviewer in ("agent", "human"):
            record_decision(workflow, stage, reviewer, "approve", "Synthetic fixture approval.")

    scenario = {"scenario_id": "saved-pricing-no-overrides", "scenario_sha256": "d" * 64,
                "input_overrides": []}
    catalog_variables = []
    semantic_variables = {}
    external_descriptors = []
    cuts = []
    axis_id = "fixture_age"
    for variable_id, extent in extents.items():
        sheet, local = extent.split("!", 1)
        min_col, min_row, max_col, max_row = range_boundaries(local)
        shape = [max_row - min_row + 1]
        if min_col != max_col:
            raise ValueError("fixture external extent must be vertical")
        catalog_variables.append({"variable_id": variable_id, "role": "formula_derived_external",
                                  "logical_name": variable_id, "shape": shape,
                                  "axes": [{"axis_id": axis_id, "role": "key"}],
                                  "source_extents": [{"sheet": sheet, "range": local, "source_role": "formula_output"}]})
        plan_extent = extent
        semantic_variables[variable_id] = {
            "variable_id": variable_id, "role": "external", "shape": shape,
            "axes": [{"axis_id": axis_id, "name": "age", "source": age_axis}],
            "source_extents": [{"ref": plan_extent, "role": "confirmed_external_cut"}],
        }
        external_descriptors.append({"variable_id": variable_id, "boundary_variable_id": variable_id,
                                     "source_extents": [plan_extent], "shape": shape})
        members = []
        for row in range(min_row, max_row + 1):
            address = f"{sheet}!{get_column_letter(min_col)}{row}"
            formula = "=UNSUPPORTED()"
            members.append({"address": address, "formula": formula,
                            "formula_sha256": hashlib.sha256(formula.encode()).hexdigest(),
                            "source_workbook_sha256": source_hash})
        cuts.append({"variable_id": variable_id, "source_range": plan_extent,
                     "coordinate_count": shape[0], "traversal": "stopped_before_prerequisites",
                     "formula_members": members})

    boundary_payload = {
        "schema_version": "step3.input_boundary.v1", "phase": "input_boundary", "status": "ready_for_review",
        "boundary_sha256": "e" * 64,
        "source": {"workbook_sha256": source_hash},
        "target_selection": {"scenario": scenario},
        "variables": catalog_variables,
        "axis_metadata": [{"axis_id": axis_id, "shape": [axis_max_row - axis_min_row + 1],
                            "source_extents": [{"sheet": axis_sheet, "range": axis_local}]}],
    }
    boundary_revision = append_stage_artifact(workflow, 3, "input_boundary.json", boundary_payload,
                                              "Synthetic boundary.")
    boundary_reference = {
        "schema_version": "step3.input_boundary.reference.v1", "revision": boundary_revision["revision"],
        "boundary_sha256": boundary_payload["boundary_sha256"], "artifact": boundary_revision["artifact"],
        "review_receipts": {reviewer: {"receipt_id": f"{reviewer}-receipt", "stage": 3,
                                       "revision": boundary_revision["revision"], "decision": "approve",
                                       "artifact_sha256": boundary_revision["artifact"]["json_sha256"],
                                       "artifact_md_sha256": boundary_revision["artifact"]["md_sha256"]}
                            for reviewer in ("agent", "human")},
    }
    manifest_path = workflow / "workflow.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["input_boundary"] = {"revision": boundary_revision["revision"],
                                  "boundary_sha256": boundary_payload["boundary_sha256"],
                                  "artifact": boundary_revision["artifact"], "source_sha256": source_hash}
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")

    evidence_dir = workflow / "evidence"
    model_scope_path = evidence_dir / "model_scope.json"
    candidate_trace_path = evidence_dir / "candidate_trace.json"
    candidate_profile_path = evidence_dir / "candidate_profile.json"
    _json(model_scope_path, {"source_sha256": source_hash})
    _json(candidate_profile_path, {"source_sha256": source_hash, "families": []})
    candidate_trace = {"source_sha256": source_hash, "external_formula_cuts": cuts,
                       "cells": [{"address": "Premium!J1", "role": "calculated_formula",
                                  "formula": "=1"}]}
    _json(candidate_trace_path, candidate_trace)
    plan_path = evidence_dir / "semantic_plan.json"
    plan = {"status": "pass", "source": {"source_sha256": source_hash},
            "variables": semantic_variables, "external_boundary_variables": external_descriptors,
            "evidence": {"model_scope": {"path": str(model_scope_path), "sha256": _sha(model_scope_path)},
                         "source_candidate_trace": {"path": str(candidate_trace_path), "sha256": _sha(candidate_trace_path)},
                         "source_family_profile": {"path": str(candidate_profile_path), "sha256": _sha(candidate_profile_path)}}}
    plan_hash = _json(plan_path, plan)
    check_path = evidence_dir / "semantic_check.json"
    check_hash = _json(check_path, {"status": "pass", "semantic_plan_sha256": plan_hash})
    stage3_payload = {
        "schema_version": "step3.analysis_design.v1", "status": "ready_for_review",
        "source": {"workbook_sha256": source_hash}, "input_boundary_reference": boundary_reference,
        "analysis": {"artifacts": {
            "semantic_plan": {"path": str(plan_path), "sha256": plan_hash},
            "semantic_check": {"path": str(check_path), "sha256": check_hash},
        }},
        "design": {"semantic_mapping_required": True, "scenarios": [scenario], "targets": [
            {"target_id": "T_GP", "selector": "GP", "shape": [], "result_order": 0}],
        },
    }
    append_stage_artifact(workflow, 3, "analysis_design.json", stage3_payload, "Synthetic approved design.")
    record_decision(workflow, 3, "agent", "approve", "Synthetic fixture approval.")
    record_decision(workflow, 3, "human", "approve", "Synthetic fixture approval.")
    return workflow, source_path


def _mock_oracle(source: Path, out: Path, *, targets: list[str], ranges: list[str]) -> dict:
    assert targets == []
    out.mkdir(parents=True)
    oracle_ranges = {}
    for selector in ranges:
        _sheet, local = selector.split("!", 1)
        min_col, min_row, max_col, max_row = range_boundaries(local)
        row_count, column_count = max_row - min_row + 1, max_col - min_col + 1
        formulas = [["=UNSUPPORTED()" for _ in range(column_count)] for _ in range(row_count)]
        oracle_ranges[selector] = {
            "qualified_address": selector,
            "values": [[float(row * column_count + col) for col in range(column_count)]
                       for row in range(row_count)],
            "excel_errors": [[None for _ in range(column_count)] for _ in range(row_count)],
            "formulas": formulas,
        }
    source_hash = _sha(source)
    payload = {"schema_version": "gp.excel_oracle.v1", "tool": "step5.oracle", "status": "pass",
               "engine": "Microsoft Excel", "engine_settings": {"full_rebuild": True},
               "source_sha256": source_hash, "source_copy_sha256": source_hash,
               "macros_executed": False, "input_overrides": [], "requested_ranges": ranges,
               "ranges": oracle_ranges}
    (out / "excel_oracle.json").write_text(json.dumps(payload), encoding="utf-8")
    (out / "excel_oracle.md").write_text("Synthetic native capture fixture.", encoding="utf-8")
    return {"status": "pass"}


def test_fixed_external_capture_binds_exact_catalog_and_formula_cut_set(tmp_path: Path, monkeypatch) -> None:
    workflow, source = _external_workflow(tmp_path)
    monkeypatch.setattr(external_inputs, "capture_excel_oracle", _mock_oracle)
    result = external_inputs.capture_external_inputs(workflow)
    assert result["status"] == "pass", result
    artifact = json.loads(Path(result["artifact"]).read_text(encoding="utf-8"))
    assert artifact["external_coordinate_count"] == 530
    assert artifact["formula_cut_member_count"] == 530
    assert set(artifact["external_values"]) == set(_EXTENTS)
    assert artifact["native_capture"]["source_copy_sha256"] == _sha(source)
    assert artifact["age_axis"]["values"] == list(range(106))
    assert artifact["native_capture"]["input_overrides"] == []
    assert artifact["native_capture"]["formula_cache_inputs"] is False
    current, _path, digest = external_inputs.load_current_external_inputs(workflow)
    assert digest == _sha(Path(result["artifact"]))
    assert current["external_values_sha256"] == artifact["external_values_sha256"]


def test_external_capture_uses_catalog_ids_and_declared_vector_length(tmp_path: Path, monkeypatch) -> None:
    extents = {"external.life_rate": "Rates!B2:B4", "external.accident_rate": "Rates!D2:D4"}
    workflow, _source = _external_workflow(tmp_path, extents, "Rates!A2:A4")
    monkeypatch.setattr(external_inputs, "capture_excel_oracle", _mock_oracle)
    result = external_inputs.capture_external_inputs(workflow)
    assert result["status"] == "pass", result
    artifact = json.loads(Path(result["artifact"]).read_text(encoding="utf-8"))
    assert artifact["external_coordinate_count"] == 6
    assert artifact["formula_cut_member_count"] == 6
    assert set(artifact["external_values"]) == set(extents)
    assert artifact["age_axis"]["axis_id"] == "fixture_age"
    assert artifact["age_axis"]["source_range"] == "Rates!A2:A4"
    assert artifact["age_axis"]["values"] == [0, 1, 2]


def test_external_capture_rejects_changed_catalog_pair_before_native_call(tmp_path: Path, monkeypatch) -> None:
    workflow, _source = _external_workflow(tmp_path)
    monkeypatch.setattr(external_inputs, "capture_excel_oracle",
                        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("must not call Excel")))
    catalog = workflow / "stage3" / "revision-0001" / "input_boundary.json"
    catalog.write_text(catalog.read_text(encoding="utf-8") + "\n", encoding="utf-8")
    result = external_inputs.capture_external_inputs(workflow)
    assert result["status"] == "blocked"
    assert "catalog artifact pair" in result["reason"]


def test_step4_capture_external_cli_has_no_user_range_override(tmp_path: Path, monkeypatch) -> None:
    workflow, _source = _external_workflow(tmp_path)
    monkeypatch.setattr(external_inputs, "capture_excel_oracle", _mock_oracle)
    result = CliRunner().invoke(app, ["step4", "capture-external", "--workflow", str(workflow)])
    assert result.exit_code == 0, result.output
    payload = json.loads(result.output)
    assert payload["status"] == "pass"
    assert payload["external_coordinate_count"] == 530
