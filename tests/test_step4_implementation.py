from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from click import unstyle
from openpyxl import Workbook
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps.step3.calculation import CalculationBlocked
from excel_to_act.steps.step4.implementation import (
    _active_projection,
    _apply_family_execution_refinements,
    _implementation_plan_binding_mismatches,
    _preflight_context,
)


def test_active_projection_recounts_exact_fresh_members(tmp_path: Path) -> None:
    candidate_path = tmp_path / "candidate_trace.json"
    source_path = tmp_path / "source.xlsx"
    workbook = Workbook()
    workbook.active.title = "Main"
    workbook.active["A1"] = "=1"
    workbook.save(source_path)
    workbook.close()
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    candidate = {"source_sha256": source_hash,
                 "cells": [{"address": "Main!A1", "formula": "=1", "role": "calculated_formula"}]}
    candidate_path.write_text(json.dumps(candidate), encoding="utf-8")
    plan = {
        "evidence": {"source_candidate_trace": {"path": str(candidate_path)}},
        "source": {"source_sha256": source_hash},
        "source_mappings": [{"source_family_id": "family.one", "equation_family_id": "equation.one",
                             "variable_id": "result", "module_id": "projection",
                             "source_members": ["Main!A1"], "member_count": 1}],
        "array_mappings": [],
        "equation_families": {"equation.one": {"family_id": "equation.one"}},
        "variables": {"result": {"role": "derived", "shape": [], "axes": [], "source_extents": []}},
        "coverage": {},
        "modules": [{"module_id": "projection"}],
    }
    profile = {"source_sha256": source_hash,
               "families": [{"family_id": "family.one", "source_members": ["Main!A1"], "member_count": 1}]}
    trace = {"source_sha256": source_hash,
             "cells": [{"address": "Main!A1", "formula": "=1", "role": "calculated_formula"}],
             "names": [], "edges": []}
    implementation = {"reference_resolution": {"named_references": [], "cell_ranges": []},
                      "execution_plan": {"passes": [{"pass_id": "projection", "module_id": "projection",
                                                        "order": 0, "source_family_ids": ["family.one"],
                                                        "variable_ids": ["result"]}]}}

    projected, projected_profile, counts = _active_projection(
        plan, profile, trace, implementation, {"axis_metadata": [], "source_records": []},
        {"allowed_regions": []}, source_path, source_hash,
    )

    assert counts["ordinary_formula_count"] == 1
    assert projected["coverage"]["ordinary_formula_members_mapped"] == 1
    assert projected["coverage"]["array_formula_followers_mapped"] == 0
    assert projected["source_mappings"][0]["source_members"] == ["Main!A1"]
    assert projected_profile["families"][0]["member_count"] == 1


def _schedule_fixture(consumer: str = "Main!A1", prerequisite: str = "Main!B1") -> tuple[dict, dict, dict]:
    segments = [
        {"source_family_ids": ["family.consumer"], "execution": "ascending_recurrence",
         "index_axis": "year", "recurrence_group": "annual", "snapshot_before_update": True,
         "segment_id": "segment.consumer"},
        {"source_family_ids": ["family.prerequisite"], "execution": "ascending_recurrence",
         "index_axis": "year", "recurrence_group": "annual", "snapshot_before_update": True,
         "segment_id": "segment.prerequisite"},
    ]
    plan = {
        "variables": {"projection": {"role": "derived", "shape": [2, 2],
                                       "axes": [{"name": "year"}, {"name": "category"}],
                                       "equation_segments": segments,
                                       "source_extents": [
                                           {"ref": consumer, "role": "formula_output",
                                            "coordinate_mapping": {"source_family_id": "family.consumer",
                                                "axes": [{"axis": "year", "dimension": 0,
                                                          "source_coordinate": "row", "origin": 1,
                                                          "index_origin": 0},
                                                         {"axis": "category", "dimension": 1,
                                                          "source_coordinate": "column", "origin": 1,
                                                          "index_origin": 0}]}},
                                           {"ref": prerequisite, "role": "formula_output",
                                            "coordinate_mapping": {"source_family_id": "family.prerequisite",
                                                "axes": [{"axis": "year", "dimension": 0,
                                                          "source_coordinate": "row", "origin": 1,
                                                          "index_origin": 0},
                                                         {"axis": "category", "dimension": 1,
                                                          "source_coordinate": "column", "origin": 1,
                                                          "index_origin": 0}]}}
                                       ]}},
        "source_mappings": [
            {"source_family_id": "family.consumer", "variable_id": "projection",
             "source_members": [consumer], "module_id": "category_projection"},
            {"source_family_id": "family.prerequisite", "variable_id": "projection",
             "source_members": [prerequisite], "module_id": "category_projection"},
        ],
    }
    trace = {
        "active_trace_sha256": "b" * 64,
        "cells": [
            {"address": consumer, "formula": "=B1", "role": "calculated_formula"},
            {"address": prerequisite, "formula": "=1", "role": "calculated_formula"},
        ],
        "edges": [{"consumer": consumer, "prerequisite": prerequisite,
                   "relationship": "cell_reference"}],
    }
    implementation = {"execution_plan": {"family_execution": [
        {"source_family_id": "family.consumer", "execution": "vectorized"}]}}
    return plan, trace, implementation


def test_vectorized_schedule_refinement_accepts_same_index_formula_dependency() -> None:
    plan, trace, implementation = _schedule_fixture()

    proof = _apply_family_execution_refinements(plan, implementation, trace)

    assert plan["variables"]["projection"]["equation_segments"][0]["execution"] == "vectorized"
    assert proof == [{"source_family_id": "family.consumer", "variable_id": "projection",
                      "segment_id": "segment.consumer", "axis": "year",
                      "old_execution": "ascending_recurrence", "new_execution": "vectorized",
                      "active_member_count": 1, "direct_formula_edge_count": 1,
                      "same_year_formula_edge_count": 1, "cross_time_formula_edge_count": 0,
                      "trace_sha256": "b" * 64}]


def test_vectorized_schedule_refinement_rejects_prior_period_formula_dependency() -> None:
    plan, trace, implementation = _schedule_fixture("Main!A2", "Main!B1")
    with pytest.raises(CalculationBlocked, match="prior-period"):
        _apply_family_execution_refinements(plan, implementation, trace)


def test_step4_plan_command_and_catalog_are_public() -> None:
    runner = CliRunner()
    help_result = runner.invoke(app, ["step4", "plan", "--help"])
    assert help_result.exit_code == 0, help_result.stdout
    assert "--implementation" in unstyle(help_result.stdout)

    catalog_result = runner.invoke(app, ["step4", "tools"])
    assert catalog_result.exit_code == 0, catalog_result.stdout
    catalog = json.loads(catalog_result.stdout)
    tool = next(item for item in catalog["tools"] if item["name"] == "step4.plan")
    assert tool["command"] == "step4 plan --workflow DIR --implementation INPUT.json"
    assert any("does not approve or advance Stage 4" in item for item in catalog["limits"])


def test_implementation_plan_json_roundtrip_binds_the_external_capture_path(tmp_path: Path) -> None:
    record = {"schema_version": "step4.implementation_plan.v1", "status": "pass",
              "external_capture": {"path": r"C:\workflow\stage4\external_inputs\revision-0001\external_inputs.json",
                                   "artifact_sha256": "capture-file-sha",
                                   "capture_sha256": "capture-sha", "values_sha256": "values-sha"}}
    path = tmp_path / "implementation_plan.json"
    path.write_text(json.dumps(record), encoding="utf-8")
    loaded = json.loads(path.read_text(encoding="utf-8"))

    assert _implementation_plan_binding_mismatches(loaded, record) == []
    changed = {**record, "external_capture": {**record["external_capture"],
                                               "path": r"C:\workflow\stage4\external_inputs\revision-0002\external_inputs.json"}}
    assert _implementation_plan_binding_mismatches(loaded, changed) == ["external_capture"]


def test_preflight_context_preserves_formula_mode_source_path_for_projection(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    import excel_to_act.steps.step4.implementation as implementation_module
    import excel_to_act.steps.step4.generator as generator_module

    root = tmp_path / "workflow"
    root.mkdir()
    source_path = tmp_path / "source.xlsx"
    source_path.write_bytes(b"read-only source identity")
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    stage3_path = root / "stage3.json"
    stage3_path.write_text("{}", encoding="utf-8")
    trace_path = tmp_path / "active_trace.json"
    trace_path.write_text(json.dumps({"cells": []}), encoding="utf-8")
    scope_path = tmp_path / "scope.json"
    scope_path.write_text(json.dumps({"allowed_regions": []}), encoding="utf-8")
    candidate_path = tmp_path / "candidate.json"
    candidate_path.write_text("{}", encoding="utf-8")
    profile_path = tmp_path / "profile.json"
    profile_path.write_text("{}", encoding="utf-8")
    plan_path = tmp_path / "semantic_plan.json"
    check_path = tmp_path / "semantic_check.json"
    capture_path = tmp_path / "external_inputs.json"
    stage3_revision = {
        "revision": "revision-0001",
        "artifact": {"json": "stage3.json", "json_sha256": "stage3-json",
                     "md_sha256": "stage3-markdown"},
    }
    context = {
        "root": root,
        "source_path": source_path,
        "source_sha256": source_hash,
        "stage3_revision": stage3_revision,
        "semantic_plan": {},
        "semantic_plan_path": plan_path,
        "semantic_check_path": check_path,
        "semantic_plan_artifact": {"sha256": "semantic-plan"},
        "semantic_check_artifact": {"sha256": "semantic-check"},
        "source_profile": {},
        "candidate_profile": {},
        "evidence_paths": {
            "candidate_trace": (candidate_path, "candidate-trace"),
            "candidate_profile": (profile_path, "candidate-profile"),
            "model_scope": (scope_path, "model-scope"),
        },
        "catalog": {},
        "catalog_path": tmp_path / "catalog.json",
    }
    capture_file_sha = "capture-file"
    monkeypatch.setattr(implementation_module, "load_workflow", lambda _path: (root, {}))
    monkeypatch.setattr(implementation_module, "require_stage_approved",
                        lambda _root, _stage: (None, stage3_revision))
    monkeypatch.setattr(implementation_module, "load_current_external_bundle",
                        lambda _path: (context, {}, capture_path, capture_file_sha))
    monkeypatch.setattr(generator_module, "_latest_discovery_trace",
                        lambda _root, _revision, _source_hash: trace_path)
    implementation_input = tmp_path / "implementation.json"
    implementation_input.write_text(json.dumps({
        "schema_version": "step4.active_implementation.v1",
        "binding": {
            "active_trace_sha256": implementation_module.hash_file(trace_path),
            "candidate_trace_sha256": "candidate-trace",
            "external_capture_sha256": capture_file_sha,
            "semantic_check_sha256": "semantic-check",
            "semantic_plan_sha256": "semantic-plan",
            "source_profile_sha256": "candidate-profile",
            "source_sha256": source_hash,
            "stage3_artifact_sha256": "stage3-json",
            "stage3_revision": "revision-0001",
        },
    }), encoding="utf-8")

    loaded = _preflight_context(root, implementation_input)

    assert loaded["source_path"] == source_path
    assert loaded["source_sha256"] == source_hash
