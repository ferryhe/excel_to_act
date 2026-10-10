from __future__ import annotations

import json

import pytest

from excel_to_act.steps.step3 import input_boundary


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
    assert "| Logical input objects | 0 |" in front and "| Scalar objects | Not recorded |" in front
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
