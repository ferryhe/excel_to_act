from __future__ import annotations

import json
import hashlib
from copy import deepcopy
from pathlib import Path
import zipfile

import pytest
from openpyxl import Workbook

from excel_to_act.steps.step3 import input_boundary


def _synthetic_sample_boundary(tmp_path):
    workbook_path = tmp_path / "synthetic.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Model"
    sheet["A1"] = 7
    sheet["B1"] = 0.1
    sheet["C1"] = None
    sheet["F1"] = 1
    sheet["F2"] = 2
    sheet["E1"] = 0
    sheet["E2"] = 1
    workbook.save(workbook_path)
    workbook.close()
    source_sha = input_boundary.conversion.hash_file(workbook_path)
    cases = []
    for age in (1, 2):
        path = tmp_path / f"synthetic-{age}.json"
        inputs = {"Age": {"axes": [], "values": age}, "Rate": {"axes": [], "values": 0.1},
                  "Options": {"axes": [{"keys": [0, 1], "name": "option", "role": "key"}],
                              "values": [{"indices": [0], "value": 1}]}}
        digest = hashlib.sha256((json.dumps(inputs, sort_keys=True, ensure_ascii=False,
                                           separators=(",", ":")) + "\n").encode("utf-8")).hexdigest()
        path.write_text(json.dumps({"inputs": inputs,
                                    "external_data": {"curve": {"source_sha256": source_sha,
                                                                "input_sha256": digest,
                                                                "version": "synthetic-producer-v1",
                                                                "axes": [{"axis_id": "age_axis", "name": "age",
                                                                          "role": "key", "keys": [0, 1]}],
                                                                "values": [0.2, 0.3]}}}), encoding="utf-8")
        cases.append({"path": str(path), "sha256": input_boundary.conversion.hash_file(path)})
    sample = {
        "schema_version": "step3.sample_boundary.v1", "source_workbook_sha256": source_sha,
        "request_cases": cases, "case_axes": {"Age": [1, 2]},
        "request_fields": [
            {"name": "Age", "variable_id": "age", "source_cells": ["Model!A1"]},
            {"name": "Rate", "variable_id": "rate", "source_cells": ["Model!B1"],
             "fixed_source_literal": 0.1},
            {"name": "Options", "variable_id": "options", "source_cells": ["Model!F1"]},
        ],
        "external_vectors": [{"name": "curve", "variable_id": "curve", "source_extent": "Model!D1:D2",
                              "age_extent": "Model!E1:E2", "age_keys": [0, 1],
                              "producer_version": "synthetic-producer-v1"}],
        "fixed_source_fields": [{"name": "Rate", "address": "Model!B1", "source_literal": 0.1}],
        "fixed_source_literals": [{"address": "Model!C1", "source_literal": None, "present": False}],
    }
    fixed_contract = tmp_path / "synthetic-fixed-closure.json"
    fixed_contract.write_text(json.dumps({"schema_version": "step4.private_fixed_literals.v1",
                                          "source_sha256": source_sha,
                                          "implementation_plan_sha256": "a" * 64,
                                          "interface_sha256": "b" * 64,
                                          "records": [{"address": "Model!C1", "indices": [0],
                                                       "value": None, "variable_id": "internal.fixed"}]}),
                              encoding="utf-8")
    sample["fixed_literal_contract"] = {"path": str(fixed_contract),
                                         "sha256": input_boundary.conversion.hash_file(fixed_contract)}
    variables = [
        {"variable_id": "age", "logical_name": "Age", "kind": "scalar", "role": "source_raw",
         "logical_role": "business_input",
         "source_extents": [{"sheet": "Model", "range": "A1"}]},
        {"variable_id": "rate", "logical_name": "Rate", "kind": "scalar", "role": "source_raw",
         "logical_role": "fixed_source",
         "source_extents": [{"sheet": "Model", "range": "B1"}]},
        {"variable_id": "curve", "logical_name": "Curve", "kind": "vector", "role": "formula_derived_external",
         "logical_role": "series",
         "axes": [{"axis_id": "age_axis", "name": "age", "role": "key"}],
         "source_extents": [{"sheet": "Model", "range": "D1:D2"}]},
        {"variable_id": "options", "logical_name": "Options", "kind": "table", "role": "source_raw",
         "logical_role": "configuration_table",
         "shape": [2], "axes": [{"name": "option", "role": "key"}],
         "source_extents": [{"sheet": "Model", "range": "F1:F2"}],
         "coordinate_map": [{"logical_index": [0], "source_cell": "Model!F1"},
                            {"logical_index": [1], "source_cell": "Model!F2"}]},
    ]
    axis_metadata = [{"axis_id": "age_axis", "shape": [2],
                      "source_extents": [{"sheet": "Model", "range": "E1:E2"}]}]
    return sample, variables, axis_metadata, workbook_path


def test_synthetic_sample_boundary_checks_request_and_source_contract(tmp_path) -> None:
    sample, variables, axes, workbook_path = _synthetic_sample_boundary(tmp_path)
    checked = input_boundary._validate_sample_boundary(sample, variables, axes, workbook_path,
                                                       sample["source_workbook_sha256"])
    assert checked["request_field_count"] == 3
    assert checked["request_raw_cell_count"] == 3
    assert checked["fixed_source_literal_count"] == 1
    assert checked["external_vector_count"] == 1
    assert checked["request_case_count"] == 2
    _trace, catalog_fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    cut = input_boundary._source_candidate_report(targets, catalog, catalog_fields)
    groups, recorded, axes, records, exclusions, _policy = input_boundary._validate_catalog(
        catalog, targets, catalog_fields, cut)
    report = input_boundary._boundary_markdown({
        "status": "ready_for_review", "boundary_sha256": "d" * 64,
        "source": targets["source"], "scenario": targets["scenario"],
        "target_selection": targets, "counts": input_boundary._input_counts(recorded, records, axes),
        "groups": groups, "variables": recorded, "axis_metadata": axes,
        "source_records": records, "upstream_exclusions": exclusions,
        "sample_boundary": checked, "pruned_source_analysis": cut, "open_questions": []})
    assert "Supported request and fixed source boundary" in report
    assert "Model!C1" in report and "Rate" in report
    assert "Separate fixed-literal source ledger" in report
    assert "synthetic-producer-v1" in report
    assert "Source catalog objects" in report
    assert "Logical input objects" not in report


def test_synthetic_sample_boundary_rejects_rebound_requests_and_literals(tmp_path) -> None:
    sample, variables, axes, workbook_path = _synthetic_sample_boundary(tmp_path)
    sample["request_cases"][0]["sha256"] = "0" * 64
    with pytest.raises(ValueError, match="request.*SHA"):
        input_boundary._validate_sample_boundary(sample, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])
    sample["request_cases"][0]["sha256"] = input_boundary.conversion.hash_file(
        Path(sample["request_cases"][0]["path"]))
    sample["fixed_source_literals"][0]["present"] = True
    with pytest.raises(ValueError, match="fixed source literal"):
        input_boundary._validate_sample_boundary(sample, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])
    sample["fixed_source_literals"][0]["present"] = False
    sample["request_fields"][2]["source_cells"] = ["Model!F2"]
    with pytest.raises(ValueError, match="sparse cells"):
        input_boundary._validate_sample_boundary(sample, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])
    sample["request_fields"][2]["source_cells"] = ["Model!F1"]
    variables[1]["logical_role"] = "business_input"
    with pytest.raises(ValueError, match="fixed-source classification"):
        input_boundary._validate_sample_boundary(sample, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])


@pytest.mark.parametrize("change,reason", [
    ("input_digest", "wrong source or age keys"),
    ("producer_version", "wrong source or age keys"),
    ("vector_axes", "wrong source or age keys"),
    ("scalar_axes", "non-scalar"),
    ("sparse_shape", "sparse-table axes or shape"),
])
def test_synthetic_sample_rejects_runtime_incompatible_request(tmp_path, change, reason) -> None:
    sample, variables, axes, workbook_path = _synthetic_sample_boundary(tmp_path)
    case = sample["request_cases"][0]
    request = json.loads(Path(case["path"]).read_text(encoding="utf-8"))
    if change == "input_digest":
        request["external_data"]["curve"]["input_sha256"] = "0" * 64
    elif change == "producer_version":
        request["external_data"]["curve"]["version"] = "wrong-producer"
    elif change == "vector_axes":
        request["external_data"]["curve"]["axes"][0]["axis_id"] = "wrong-axis"
    elif change == "scalar_axes":
        request["inputs"]["Age"]["axes"] = [{"keys": [1, 2]}]
    else:
        request["inputs"]["Options"]["axes"][0]["keys"] = [0]
    changed = tmp_path / f"synthetic-{change}.json"
    changed.write_text(json.dumps(request), encoding="utf-8")
    case.update(path=str(changed), sha256=input_boundary.conversion.hash_file(changed))
    with pytest.raises(ValueError, match=reason):
        input_boundary._validate_sample_boundary(sample, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])


def test_synthetic_sample_rejects_wrong_source_axis_and_fixed_set(tmp_path) -> None:
    sample, variables, axes, workbook_path = _synthetic_sample_boundary(tmp_path)
    wrong_extent = deepcopy(sample)
    wrong_extent["external_vectors"][0]["age_extent"] = "Model!G1:G2"
    with pytest.raises(ValueError, match="age extent"):
        input_boundary._validate_sample_boundary(wrong_extent, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])
    wrong_keys = deepcopy(sample)
    wrong_keys["external_vectors"][0]["age_keys"] = [8, 9]
    for case in wrong_keys["request_cases"]:
        request = json.loads(Path(case["path"]).read_text(encoding="utf-8"))
        request["external_data"]["curve"]["axes"][0]["keys"] = [8, 9]
        changed = tmp_path / f"synthetic-wrong-keys-{Path(case['path']).stem}.json"
        changed.write_text(json.dumps(request), encoding="utf-8")
        case.update(path=str(changed), sha256=input_boundary.conversion.hash_file(changed))
    with pytest.raises(ValueError, match="source workbook axis"):
        input_boundary._validate_sample_boundary(wrong_keys, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])
    wrong_literals = deepcopy(sample)
    wrong_literals["fixed_source_literals"][0] = {
        "address": "Model!Z99", "source_literal": None, "present": False}
    with pytest.raises(ValueError, match="bound source closure"):
        input_boundary._validate_sample_boundary(wrong_literals, variables, axes, workbook_path,
                                                 sample["source_workbook_sha256"])


def test_synthetic_fixed_source_is_a_supported_catalog_role() -> None:
    _trace, _fields, _targets, _binding, catalog = _boundary_fixture()
    fixed = catalog["variables"][0]
    fixed["logical_role"] = "fixed_source"
    input_boundary._validate_variable(fixed, {"policy"}, {"gp"})


def _boundary_fixture() -> tuple[dict, dict, dict, dict, dict]:
    source = {"source_id": "source-1", "run_id": "run-1", "workbook_sha256": "a" * 64,
              "analysis_binding_sha256": "b" * 64}
    scenario_basis = {"scenario_id": "base-case", "selectors": {"level": "Level100", "benefit": "None"}}
    scenario = {**scenario_basis, "scenario_sha256": input_boundary.conversion.hash_bytes(
        input_boundary.conversion.json_bytes(scenario_basis))}
    target = {"target_id": "gp", "selector": "GP", "result_order": 0, "shape": [],
              "start_cells": ["Main!C1"]}
    target_selection = {"source": source, "scenario": scenario, "targets": [target]}
    binding = {"source": {"source_id": "source-1", "run_id": "run-1", "source_sha256": "a" * 64}}
    trace = {
        "schema_version": "gp.active_trace.v1",
        "source_sha256": "a" * 64,
        "native_excel_called": False,
        "formula_cache_inputs": False,
        "targets": {"gp": {"selector": "GP", "result_order": 0, "shape": [], "value": "must-not-be-copied"}},
        "cells": [
            {"address": "Main!C1", "role": "calculated_formula", "formula": "=B1+SUM(A1:A2)", "value": "cached-target"},
            {"address": "Main!B1", "role": "calculated_formula", "formula": "=A1*2", "value": "cached-boundary"},
            {"address": "Main!A1", "role": "source_raw", "formula": None, "value": "private-raw-value"},
            {"address": "Main!A2", "role": "source_raw", "formula": None, "value": "private-axis-value"},
        ],
        "edges": [
            {"consumer": "Main!C1", "prerequisite": "Main!B1", "relationship": "cell_reference"},
            {"consumer": "Main!C1", "prerequisite": "Main!A1:A2", "relationship": "range_read"},
            {"consumer": "Main!B1", "prerequisite": "Main!A1", "relationship": "cell_reference"},
        ],
        "lookups": [],
    }
    fields = {"fields": [
        {"field_id": "source-a1", "sheet": "Main", "role": "source", "members": [{"address": "A1"}]},
        {"field_id": "source-a2", "sheet": "Main", "role": "source", "members": [{"address": "A2"}]},
        {"field_id": "formula-b1", "sheet": "Main", "role": "calculated", "members": [{"address": "B1"}]},
        {"field_id": "formula-c1", "sheet": "Main", "role": "calculated", "members": [{"address": "C1"}]},
    ], "descriptors": {"defined_names": [{"name": "BoundaryRate", "sheet": "Main", "address": "B1"}]}}
    catalog = {
        "schema_version": "step3.input_boundary.input.v1",
        "source": source,
        "scenario": scenario,
        "target_selection": target_selection,
        "groups": [{"group_id": "policy", "name": "Policy inputs", "classification": "business_configuration"}],
        "variables": [
            {"variable_id": "input.policy_flag", "group_id": "policy", "logical_name": "Policy flag",
             "role": "source_raw", "kind": "scalar", "logical_role": "business_input", "shape": [], "axes": [],
             "source_extents": [{"sheet": "Main", "range": "A1", "source_role": "raw_source"}],
             "consumer_targets": ["gp"], "provenance": {"description": "Raw source coordinate; label remains under review."},
             "value_source_policy": {"origin": "workbook_formula_mode_source",
                                     "formula_cache_allowed": False}},
            {"variable_id": "external.ci_rate", "group_id": "policy", "logical_name": "External CI rate",
             "role": "formula_derived_external", "kind": "scalar", "logical_role": "business_input", "shape": [], "axes": [],
             "source_extents": [{"sheet": "Main", "range": "B1", "source_role": "formula_output"}],
             "consumer_targets": ["gp"], "provenance": {"description": "Formula-derived candidate boundary; no value supplied."},
             "value_source_policy": {"allowed_future_sources": ["approved_excel_capture", "user_rate_file"],
                                     "formula_cache_allowed": False, "available_now": False}},
        ],
        "axis_metadata": [],
        "source_records": [{"record_id": "pending.policy.axis", "group_id": "policy",
                             "classification": "pending_classification", "open_question_id": "Q1",
                             "source_extents": [{"sheet": "Main", "range": "A2", "source_role": "raw_source"}]}],
        "upstream_exclusions": [{"boundary_variable_id": "external.ci_rate", "named_source": "BoundaryRate",
                                 "source_range": "Main!B1", "reason": "Candidate external boundary for review."}],
        "value_source_policy": {"formula_cache_allowed": False},
        "topology_evidence": {"use": "source_topology_only", "path": "unused.json", "sha256": "c" * 64},
        "open_questions": [{"question_id": "Q1", "question": "Confirm whether Main!A2 is a policy axis or adapter metadata."}],
    }
    return trace, fields, target_selection, binding, catalog


def test_source_topology_cut_and_partition_keep_values_out_of_catalog() -> None:
    trace, fields, targets, binding, catalog = _boundary_fixture()
    cut = input_boundary._cut_report(trace, targets, catalog, fields, binding)
    groups, variables, axes, records, exclusions, policy = input_boundary._validate_catalog(
        catalog, targets, fields, cut)
    counts = input_boundary._input_counts(variables, records, axes)

    assert cut["status"] == "historical_topology_candidate"
    assert cut["retained_formula_positions"] == 1
    assert cut["retained_raw_source_coordinates"] == 2
    assert cut["candidate_boundary_reached_count"] == 1
    assert counts["logical_input_object_count"] == 2
    assert counts["source_raw_variable_count"] == 1
    assert counts["formula_derived_external_variable_count"] == 1
    assert counts["source_coordinate_count"] == 3
    assert counts["pending_classification_coordinate_count"] == 1
    assert groups and axes == [] and records[0]["classification"] == "pending_classification"
    assert exclusions and policy["formula_cache_allowed"] is False
    serialized = json.dumps({"cut": {key: value for key, value in cut.items() if not key.startswith("_")},
                             "variables": variables, "records": records})
    for forbidden in ("private-raw-value", "private-axis-value", "cached-target", "cached-boundary", "must-not-be-copied"):
        assert forbidden not in serialized


def test_upstream_exclusion_can_keep_a_named_range_column_for_runtime() -> None:
    trace, fields, targets, binding, catalog = _boundary_fixture()
    fields["descriptors"]["defined_names"][0]["address"] = "B1:D2"
    exclusion = catalog["upstream_exclusions"][0]
    exclusion["source_range"] = "Main!B1:D2"
    exclusion["excluded_range"] = "Main!C1:C2"
    cut = input_boundary._cut_report(trace, targets, catalog, fields, binding)

    input_boundary._validate_catalog(catalog, targets, fields, cut)
    from excel_to_act.steps.step3.profiling import _catalog_upstream_exclusions

    coordinates, _records = _catalog_upstream_exclusions(catalog)
    assert coordinates == {"main!C1", "main!C2"}

    for reversed_ref in ("Main!C2:C1", "Main!D1:B1"):
        exclusion["excluded_range"] = reversed_ref
        with pytest.raises(ValueError, match="reversed"):
            input_boundary._validate_catalog(catalog, targets, fields, cut)

    exclusion["excluded_range"] = "Main!E1:E2"
    with pytest.raises(ValueError, match="within the named source range"):
        input_boundary._validate_catalog(catalog, targets, fields, cut)

    exclusion["excluded_range"] = None
    with pytest.raises(ValueError, match="within the named source range"):
        input_boundary._validate_catalog(catalog, targets, fields, cut)


@pytest.mark.parametrize("mutation, error", [
    ("formula_as_raw", "includes a formula cell"),
    ("promote_target", "cannot be promoted to an input"),
    ("unknown_question", "undeclared open question"),
    ("missing_coordinate", "account for every retained raw coordinate in the declared topology candidate"),
])
def test_catalog_rejects_wrong_roles_missing_raw_or_unsupported_target(
    mutation: str, error: str,
) -> None:
    trace, fields, targets, binding, catalog = _boundary_fixture()
    cut = input_boundary._cut_report(trace, targets, catalog, fields, binding)
    if mutation == "formula_as_raw":
        catalog["variables"][0]["source_extents"] = [
            {"sheet": "Main", "range": "B1", "source_role": "raw_source"}]
    elif mutation == "promote_target":
        catalog["variables"][0]["source_extents"] = [
            {"sheet": "Main", "range": "C1", "source_role": "raw_source"}]
    elif mutation == "unknown_question":
        catalog["source_records"][0]["open_question_id"] = "Q404"
    else:
        catalog["source_records"] = []
    with pytest.raises(ValueError, match=error):
        input_boundary._validate_catalog(catalog, targets, fields, cut)


def test_catalog_value_cache_fields_are_rejected() -> None:
    _trace, _fields, _targets, _binding, catalog = _boundary_fixture()
    catalog["variables"][0]["provenance"]["value"] = 123
    with pytest.raises(ValueError, match="must describe value provenance"):
        input_boundary._assert_no_embedded_values(catalog, "catalog")


def test_catalog_without_historical_trace_is_reviewable_but_closure_stays_unknown() -> None:
    _trace, fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    cut = input_boundary._source_candidate_report(targets, catalog, fields)
    input_boundary._validate_catalog(catalog, targets, fields, cut)

    assert cut["status"] == "source_candidate_only_target_closure_unavailable"
    assert cut["candidate_boundary_coordinate_count"] == 1
    assert cut["retained_raw_source_coordinates"] == 2
    assert cut["unknown_count"] is None
    assert cut["unrepresented_source_coordinate_count"] is None
    assert cut["input_coverage"]["status"] == "target_closure_unproven_without_topology"
    assert cut["input_coverage"]["unassigned_raw_coordinate_count"] is None


def test_catalog_accepts_only_source_verified_stored_blank(tmp_path) -> None:
    _trace, fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    catalog["source_records"].append({
        "record_id": "fixed.blank", "group_id": "policy", "classification": "source_constant",
        "source_extents": [{"sheet": "Main", "range": "K81", "source_role": "raw_source"}],
    })
    catalog["source_scope_notes"] = {"stored_coordinate_reconciliation": {
        "source_copy_sha256": "", "historical_empty_coordinates": [{
            "address": "Main!K81", "data_type": "n", "source_value_status": "stored_blank_no_literal",
        }],
    }}
    inventory = {"sheets": [{"name": "Main", "cells": [{
        "address": "A1", "source_location": {"ooxml_part": "xl/worksheets/sheet1.xml"},
    }]}]}
    source = tmp_path / "source.xlsx"

    def write_cell(xml: str) -> None:
        with zipfile.ZipFile(source, "w") as package:
            package.writestr("xl/worksheets/sheet1.xml", "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
                             "<sheetData><row r='81'>" + xml + "</row></sheetData></worksheet>")

    write_cell("<c r='K81' s='1'/>")
    source_sha = input_boundary.conversion.hash_file(source)
    catalog["source_scope_notes"]["stored_coordinate_reconciliation"]["source_copy_sha256"] = source_sha
    augmented = input_boundary._with_verified_stored_blanks(catalog, fields, inventory, source, source_sha)
    cut = input_boundary._source_candidate_report(targets, catalog, augmented)
    input_boundary._validate_catalog(catalog, targets, augmented, cut)
    assert fields != augmented
    assert cut["retained_raw_source_coordinates"] == 3

    with pytest.raises(ValueError, match="checksum"):
        input_boundary._with_verified_stored_blanks(catalog, fields, inventory, source, "0" * 64)
    write_cell("<c r='K81' s='1'><v>0</v></c>")
    catalog["source_scope_notes"]["stored_coordinate_reconciliation"]["source_copy_sha256"] = input_boundary.conversion.hash_file(source)
    with pytest.raises(ValueError, match="stored blank"):
        input_boundary._with_verified_stored_blanks(catalog, fields, inventory, source,
                                                     catalog["source_scope_notes"]["stored_coordinate_reconciliation"]["source_copy_sha256"])


def test_input_boundary_rejects_malformed_stored_coordinate_before_reading_it(tmp_path, monkeypatch) -> None:
    _trace, fields, targets, binding, catalog = _boundary_fixture()
    catalog["source_scope_notes"] = {"stored_coordinate_reconciliation": {
        "historical_empty_coordinates": [None],
    }}
    targets_path = tmp_path / "targets.json"
    catalog_path = tmp_path / "catalog.json"
    targets_path.write_text(json.dumps(targets), encoding="utf-8")
    catalog_path.write_text(json.dumps(catalog), encoding="utf-8")
    analysis = {"binding_sha256": "b" * 64}
    monkeypatch.setattr(input_boundary.step3, "_load_context", lambda path: (
        path, analysis, binding, tmp_path, tmp_path / "inventory.json"))
    monkeypatch.setattr(input_boundary.conversion, "load_workflow", lambda path: (
        path, {"schema_version": "conversion.workflow.v2",
                   "origin": "native",
               "source": {"step1_source_copy_path": str(tmp_path / "source.xlsx")}}))
    monkeypatch.setattr(input_boundary, "_require_bound_workflow", lambda *args: None)
    monkeypatch.setattr(input_boundary.conversion, "require_stage_approved", lambda *args: (None, {"artifact": {}}))
    monkeypatch.setattr(input_boundary, "_validate_targets", lambda *args: targets)
    monkeypatch.setattr(input_boundary, "_validate_catalog_identity", lambda *args: None)
    monkeypatch.setattr(input_boundary.step3, "_read_stage", lambda *args: (fields, "c" * 64))
    monkeypatch.setattr(input_boundary.step3, "_load_inventory", lambda *args: {"sheets": []})
    appended = []
    monkeypatch.setattr(input_boundary.conversion, "append_stage_artifact", lambda *args, **kwargs: appended.append(args))

    result = input_boundary.create_input_boundary(tmp_path, tmp_path, targets_path, catalog_path)

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["message"] == "historical empty coordinate must be an object"
    assert appended == []


def test_historical_trace_reports_range_coordinates_absent_from_fields_separately() -> None:
    trace, fields, targets, binding, catalog = _boundary_fixture()
    trace["edges"][1]["prerequisite"] = "Main!A1:A3"
    cut = input_boundary._cut_report(trace, targets, catalog, fields, binding)

    assert cut["unknown_count"] == 0
    assert cut["unrepresented_source_coordinate_count"] == 1
    assert cut["unrepresented_source_coordinate_examples"] == ["main!a3"]
    assert "blank-read behavior" in cut["unrepresented_semantics"]


def test_input_counts_distinguish_objects_from_vector_fields_and_curves() -> None:
    variables = [
        {"kind": "scalar", "role": "source_raw", "shape": [], "axes": []},
        {"kind": "vector", "logical_role": "other", "role": "source_raw", "shape": [3], "axes": [{}]},
        {"kind": "vector", "logical_role": "series", "role": "source_raw", "shape": [5], "axes": [{}]},
        {"kind": "table", "role": "source_raw", "shape": [2, 2], "axes": [{}, {}], "curve_count": 2,
         "coordinate_map": [{"source_cell": "Main!A1"}, {"source_cell": "Main!B2"}]},
    ]
    counts = input_boundary._input_counts(variables, [], [])

    assert counts["logical_input_object_count"] == 4
    assert counts["scalar_field_count"] == 1
    assert counts["vector_object_count"] == 2
    assert counts["non_curve_vector_count"] == 1
    assert counts["vector_curve_count"] == 1
    assert counts["field_or_curve_count"] == 4
    assert counts["declared_logical_element_count"] == 11
    assert counts["sparse_table_mapped_coordinate_count"] == 2
    assert counts["sparse_table_bounding_slot_count"] == 4


def test_boundary_markdown_calls_empty_source_coordinates_stored_but_empty() -> None:
    trace, fields, targets, binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    cut = input_boundary._source_candidate_report(targets, catalog, fields)
    groups, variables, axes, records, exclusions, _policy = input_boundary._validate_catalog(
        catalog, targets, fields, cut)
    variables[0]["source_column_labels"] = [{"source_cell": "Premium!H8", "label": "MinorCI\npx+t-1"}]
    counts = input_boundary._input_counts(variables, records, axes)
    payload = {
        "status": "draft_not_approved", "boundary_sha256": "d" * 64,
        "source": targets["source"], "scenario": targets["scenario"],
        "target_selection": targets, "counts": counts, "groups": groups,
        "variables": variables, "axis_metadata": axes, "pruned_source_analysis": cut,
        "source_scope_notes": {
            "note": "No current trace was supplied.",
            "stored_coordinate_reconciliation": {
                "historical_empty_coordinates": ["Benefit_Table!AG3"],
                "interpretation": "The stored cell has no formula or literal content; later blank-read semantics remain open.",
                "stored_coordinate_audit": {
                    "path": "stored-coordinate-audit.json", "sha256": "e" * 64,
                    "old_trace_raw_coordinate_count": 2631,
                    "current_retained_raw_coordinate_count": 2623,
                    "empty_source_coordinate_count": 8,
                    "coordinates": [{"address": "Benefit_Table!AG3"}],
                },
            },
        },
        "upstream_exclusions": exclusions, "source_records": records,
        "open_questions": catalog["open_questions"],
    }
    report = input_boundary._boundary_markdown(payload)

    assert "cell element present; no formula or literal content" in report
    assert "stored-but-empty source cells" in report
    assert "physical absence" not in report
    assert "always requires an independent Agent decision and an explicit human confirmation" in report
    assert "cannot approve or reject this checkpoint" in report
    assert "unknown fields such as `sample_rates` or `rates` block submission" in report
    group = next(item for item in groups if item["group_id"] == variables[0]["group_id"])
    assert f"**{group['name']}**" in report
    reader_front = report.split("## Detailed source and boundary evidence", 1)[0]
    variable_row = next(line for line in reader_front.splitlines() if "`input.policy_flag`" in line)
    assert f"role `{variables[0]['role']}`" in variable_row
    assert "Main!A1" in variable_row
    technical_row = next(line for line in report.splitlines() if line.startswith("| `input.policy_flag`"))
    assert "MinorCI<br>px+t-1" in technical_row
    for extent in variables[0]["source_extents"]:
        assert f"{extent['sheet']}!{extent['range']}" in variable_row
    assert "Recorded disposition/status: Not recorded." in report
    assert "MinorCI\npx+t-1" not in technical_row
    assert variables[0]["source_column_labels"][0]["label"] == "MinorCI\npx+t-1"
    assert input_boundary._markdown_table_cell("label|with pipe") == "label\\|with pipe"



@pytest.mark.parametrize("with_note", [True, False])
@pytest.mark.parametrize("with_interpretation", [True, False])
def test_boundary_markdown_without_optional_audit_keeps_source_blank_claim_limited(
    with_note, with_interpretation
) -> None:
    _trace, fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    cut = input_boundary._source_candidate_report(targets, catalog, fields)
    groups, variables, axes, records, exclusions, _policy = input_boundary._validate_catalog(
        catalog, targets, fields, cut)
    payload = {
        "status": "ready_for_review", "boundary_sha256": "d" * 64,
        "source": targets["source"], "scenario": targets["scenario"],
        "target_selection": targets, "counts": input_boundary._input_counts(variables, records, axes),
        "groups": groups, "variables": variables, "axis_metadata": axes,
        "pruned_source_analysis": cut, "upstream_exclusions": exclusions,
        "source_records": records, "open_questions": catalog["open_questions"],
        "source_scope_notes": {"note": "Source verification only.",
                               "stored_coordinate_reconciliation": {
                                   "historical_empty_coordinates": [{
                                       "address": "Main!K81", "data_type": "n",
                                       "source_value_status": "stored_blank_no_literal",
                                   }],
                                   "interpretation": "Later blank-read semantics remain unresolved.",
                               }},
    }
    if not with_note:
        payload["source_scope_notes"].pop("note")
    if not with_interpretation:
        payload["source_scope_notes"]["stored_coordinate_reconciliation"].pop("interpretation")
    report = input_boundary._boundary_markdown(payload)

    assert "Main!K81" in report
    assert "stored blank" in report.lower()
    assert "Later blank-read semantics remain unresolved." in report
    if not with_interpretation:
        assert "No interpretation was supplied" in report
    assert "No separate source-coordinate audit" in report
    assert "The source-coordinate audit" not in report
    assert "`None`" not in report


@pytest.mark.parametrize("change", ["valid", "missing_identity", "byte_mismatch",
                                    "source_mismatch", "content_mismatch"])
def test_supplied_coordinate_audit_requires_bound_file_and_source(tmp_path, change) -> None:
    _trace, fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    audit_data = {
        "source_sha256": targets["source"]["workbook_sha256"],
        "old_trace_raw_coordinate_count": 1,
        "current_retained_raw_coordinate_count": 1,
        "empty_source_coordinate_count": 0,
        "coordinates": [],
    }
    audit_path = tmp_path / "synthetic-coordinate-audit.json"
    audit_path.write_text(json.dumps(audit_data), encoding="utf-8")
    audit = {**audit_data, "path": str(audit_path),
             "sha256": input_boundary.conversion.hash_file(audit_path)}
    catalog["source_scope_notes"] = {"stored_coordinate_reconciliation": {
        "stored_coordinate_audit": audit,
    }}
    if change == "missing_identity":
        for key in ("path", "sha256", "source_sha256"):
            audit.pop(key)
    elif change == "byte_mismatch":
        audit_path.write_text(json.dumps({**audit_data, "extra": True}), encoding="utf-8")
    elif change == "source_mismatch":
        audit_data["source_sha256"] = "b" * 64
        audit_path.write_text(json.dumps(audit_data), encoding="utf-8")
        audit.update(source_sha256=audit_data["source_sha256"],
                     sha256=input_boundary.conversion.hash_file(audit_path))
    else:
        if change == "content_mismatch":
            audit["old_trace_raw_coordinate_count"] = 2
    cut = input_boundary._source_candidate_report(targets, catalog, fields)
    if change == "valid":
        groups, variables, axes, records, exclusions, _policy = input_boundary._validate_catalog(
            catalog, targets, fields, cut)
        payload = {
            "status": "ready_for_review", "boundary_sha256": "d" * 64,
            "source": targets["source"], "scenario": targets["scenario"],
            "target_selection": targets,
            "counts": input_boundary._input_counts(variables, records, axes),
            "groups": groups, "variables": variables, "axis_metadata": axes,
            "pruned_source_analysis": cut, "upstream_exclusions": exclusions,
            "source_records": records, "open_questions": catalog["open_questions"],
            "source_scope_notes": catalog["source_scope_notes"],
        }
        report = input_boundary._boundary_markdown(payload)
        assert "The source-coordinate audit" in report
        assert str(audit_path) in report
        assert audit["sha256"] in report
        assert "1 historical raw candidates" in report
        assert "No separate source-coordinate audit" not in report
    else:
        with pytest.raises(ValueError, match="stored_coordinate_audit"):
            input_boundary._validate_catalog(catalog, targets, fields, cut)


def test_boundary_reader_null_metadata_is_unknown_and_preserves_zero_false_empty_shape() -> None:
    import copy

    trace, fields, targets, binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    cut = input_boundary._source_candidate_report(targets, catalog, fields)
    groups, variables, axes, records, exclusions, _policy = input_boundary._validate_catalog(
        catalog, targets, fields, cut)
    counts = input_boundary._input_counts(variables, records, axes)
    payload = {
        "status": "draft_not_approved", "boundary_sha256": "d" * 64,
        "source": {**targets["source"], "run_id": None, "source_id": "None"},
        "scenario": targets["scenario"], "target_selection": targets,
        "counts": {**counts, "logical_input_object_count": 0, "scalar_field_count": None},
        "groups": groups, "variables": variables, "axis_metadata": axes,
        "pruned_source_analysis": cut, "upstream_exclusions": exclusions,
        "source_records": records,
        "open_questions": [{"question_id": "Q_UNKNOWN", "question": "Which unit applies?", "status": None}],
    }
    payload["variables"][0].update({"kind": None, "shape": [], "role": False})
    payload["variables"][0]["source_extents"][0]["range"] = None
    payload["target_selection"]["targets"][0].update({"selector": None, "shape": []})
    original = copy.deepcopy(payload)
    report = input_boundary._boundary_markdown(payload, machine_link="boundary.json")
    front = report.split("## Detailed source and boundary evidence", 1)[0]

    assert "Bound source run: `Not recorded`; source ID: `None`" in front
    assert "| Source catalog objects | 0 |" in front and "| Scalar objects | Not recorded |" in front
    assert "Not recorded []" in front and "role `False`" in front
    assert "Main!Not recorded" in front
    assert "(order 0; kind Not recorded; shape []; units Not recorded; declared axes Not recorded); source start" in front
    assert "Recorded disposition/status: Not recorded." in front
    assert payload == original


def test_boundary_target_and_input_handoffs_show_only_recorded_axes() -> None:
    import copy

    _trace, fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    catalog["variables"][0].update({
        "kind": "vector", "logical_role": "series", "shape": [1],
        "axes": [{"name": "Policy duration", "axis_id": "policy-duration", "role": "observation"}],
        "source_extents": [{"sheet": "Main", "range": "A1", "source_role": "raw_source"}],
    })
    cut = input_boundary._source_candidate_report(targets, catalog, fields)
    groups, variables, axes, records, exclusions, _policy = input_boundary._validate_catalog(
        catalog, targets, fields, cut)
    payload = {
        "status": "draft_not_approved", "boundary_sha256": "d" * 64,
        "source": targets["source"], "scenario": targets["scenario"],
        "target_selection": targets, "counts": input_boundary._input_counts(variables, records, axes),
        "groups": groups, "variables": variables, "axis_metadata": axes,
        "pruned_source_analysis": cut, "upstream_exclusions": exclusions,
        "source_records": records, "open_questions": catalog["open_questions"],
    }
    original = copy.deepcopy(payload)
    report = input_boundary._boundary_markdown(payload, machine_link="boundary.json")
    front = report.split("## Detailed source and boundary evidence", 1)[0]

    target_line = next(line for line in front.splitlines() if line.startswith("- `gp`:"))
    input_line = next(line for line in front.splitlines() if "`input.policy_flag`" in line)
    detail_line = next(line for line in report.splitlines() if line.startswith("| `input.policy_flag`"))
    assert "declared axes Not recorded" in target_line
    assert "input axes name: Policy duration; axis_id: policy-duration; role: observation" in input_line
    assert "name: Policy duration; axis_id: policy-duration; role: observation" in detail_line
    assert "Machine checkpoint: [input_boundary.json](<boundary.json>)" in front
    assert payload == original
    assert input_boundary.conversion.display_report_axes({"axes": None}) == "Not recorded"
    assert input_boundary.conversion.display_report_axes({}) == "Not recorded"
    assert input_boundary.conversion.display_report_axes({"axes": []}) == "[]"
    assert input_boundary.conversion.display_report_axes({
        "axes": [{"keys": [0, False, "None", None]}],
    }) == 'keys: [0, False, "None", Not recorded]'


def test_input_target_metadata_binding_and_handoffs_preserve_order_and_unknowns() -> None:
    import copy

    _trace, fields, original_selection, binding, catalog = _boundary_fixture()
    scenario_basis = {key: value for key, value in original_selection["scenario"].items()
                      if key != "scenario_sha256"}
    scenario_basis["overrides"] = {}
    scenario = {**scenario_basis, "scenario_sha256": input_boundary.conversion.hash_bytes(
        input_boundary.conversion.json_bytes(scenario_basis))}
    legacy_selection = {**original_selection, "scenario": scenario}
    minimal_input = {"schema_version": "step3.input_targets.v1", **legacy_selection}
    minimal = input_boundary._validate_targets(
        minimal_input, binding, legacy_selection["source"]["analysis_binding_sha256"])
    assert minimal == legacy_selection
    assert input_boundary.conversion.hash_bytes(input_boundary.conversion.json_bytes(minimal)) == \
        input_boundary.conversion.hash_bytes(input_boundary.conversion.json_bytes(legacy_selection))

    submitted = {
        "schema_version": "step3.input_targets.v1",
        "source": legacy_selection["source"],
        "scenario": scenario,
        "targets": [
            {"target_id": "vector", "selector": "Vector output", "result_order": 1,
             "shape": [3], "start_cells": ["Main!C1"], "result_kind": "vector",
             "units": "yearly amount", "axes": [{"name": "policy year", "axis_id": "policy_year",
                                                      "role": "time", "keys": [1, 2, 3]}]},
            {"target_id": "matrix", "selector": "Matrix output", "result_order": 0,
             "shape": [2, 2], "start_cells": ["Main!C1"], "result_kind": None,
             "units": None, "axes": None},
        ],
    }
    original = copy.deepcopy(submitted)
    selection = input_boundary._validate_targets(
        submitted, binding, original_selection["source"]["analysis_binding_sha256"])
    assert [item["target_id"] for item in selection["targets"]] == ["vector", "matrix"]
    assert selection["targets"][0]["axes"][0]["keys"] == [1, 2, 3]
    assert all(key in selection["targets"][1] and selection["targets"][1][key] is None
               for key in ("result_kind", "units", "axes"))
    assert submitted == original

    bound_catalog = {**catalog, "scenario": scenario, "target_selection": selection}
    input_boundary._validate_catalog_identity(bound_catalog, binding, selection)
    mismatched_catalog = copy.deepcopy(bound_catalog)
    mismatched_catalog["target_selection"]["targets"][0]["units"] = "different unit"
    with pytest.raises(ValueError, match="target order and selectors must exactly match"):
        input_boundary._validate_catalog_identity(mismatched_catalog, binding, selection)

    catalog.pop("topology_evidence")
    cut = input_boundary._source_candidate_report(original_selection, catalog, fields)
    groups, variables, axes, records, exclusions, _policy = input_boundary._validate_catalog(
        catalog, original_selection, fields, cut)
    payload = {
        "status": "draft_not_approved", "boundary_sha256": "d" * 64,
        "source": selection["source"], "scenario": selection["scenario"],
        "target_selection": selection, "counts": input_boundary._input_counts(variables, records, axes),
        "groups": groups, "variables": variables, "axis_metadata": axes,
        "pruned_source_analysis": cut, "upstream_exclusions": exclusions,
        "source_records": records, "open_questions": [],
    }
    report = input_boundary._boundary_markdown(payload, machine_link="boundary.json")
    front = report.split("## Detailed source and boundary evidence", 1)[0]
    requested = front.split("Requested targets:\n\n", 1)[1].split("\nRecorded logical input kinds:", 1)[0]
    assert requested.index("`matrix`") < requested.index("`vector`")
    assert "kind Not recorded; shape [2, 2]; units Not recorded; declared axes Not recorded" in requested
    assert "kind vector; shape [3]; units yearly amount; declared axes name: policy year; axis_id: policy_year; role: time; keys: [1, 2, 3]" in requested
    detail_targets = report.split("## Detailed source and boundary evidence", 1)[1]
    assert detail_targets.index("`matrix` · order 0") < detail_targets.index("`vector` · order 1")

def test_source_candidate_catalog_rejects_formula_coordinates_claimed_as_raw() -> None:
    _trace, fields, targets, _binding, catalog = _boundary_fixture()
    catalog.pop("topology_evidence")
    catalog["variables"][0]["source_extents"] = [
        {"sheet": "Main", "range": "B1", "source_role": "raw_source"}]
    with pytest.raises(ValueError, match="does not match a retained raw source field"):
        cut = input_boundary._source_candidate_report(targets, catalog, fields)
        input_boundary._validate_catalog(catalog, targets, fields, cut)


def test_sparse_configuration_table_preserves_exact_nonrectangular_cells() -> None:
    trace, fields, targets, binding, catalog = _boundary_fixture()
    catalog["variables"][0].update({
        "kind": "table", "shape": [2, 2],
        "axes": [{"name": "configuration_row", "role": "record"},
                 {"name": "configuration_family", "role": "category"}],
        "source_extents": [
            {"sheet": "Main", "range": "A1", "source_role": "raw_source"},
            {"sheet": "Main", "range": "A2", "source_role": "raw_source"}],
        "coordinate_map": [
            {"source_cell": "Main!A1", "logical_index": [0, 0]},
            {"source_cell": "Main!A2", "logical_index": [1, 1]}],
    })
    catalog["source_records"] = []
    catalog["open_questions"] = []
    cut = input_boundary._cut_report(trace, targets, catalog, fields, binding)
    input_boundary._validate_catalog(catalog, targets, fields, cut)

    catalog["variables"][0]["coordinate_map"].pop()
    with pytest.raises(ValueError, match="must cover every exact source coordinate"):
        input_boundary._validate_catalog(catalog, targets, fields, cut)
