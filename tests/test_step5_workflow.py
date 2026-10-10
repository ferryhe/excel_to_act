from __future__ import annotations

import json
import hashlib
from pathlib import Path

from openpyxl import Workbook

from excel_to_act.steps.conversion_workflow import (
    _write_manifest,
    append_stage_artifact,
    create_workflow,
    hash_file,
    load_workflow,
    record_decision,
)
from excel_to_act.steps.step5.workflow import (
    _active_formula_addresses,
    _active_formula_details,
    _compare_active_formula_members,
    _compare_external_boundary,
    _oracle_range_cells,
    _preflight_oracle,
    _reconciliation_markdown,
    _verify_python,
    reconcile,
    validate_generated,
)


def _canonical_hash(value: object) -> str:
    return hashlib.sha256((json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":")) + "\n")
                           .encode("utf-8")).hexdigest()


def test_stage5_compares_every_active_formula_across_ranges_and_preserves_errors() -> None:
    oracle = {"ranges": {
        "Premium!C1:D2": {
            "qualified_address": "Premium!C1:D2",
            "values": [[2.5, None], [None, None]],
            "formulas": [["=1+1.5", ""], ["", "=#N/A"]],
            "excel_errors": [[None, None], [None, "#N/A"]],
        },
        "Main!F34:F34": {
            "qualified_address": "Main!F34",
            "values": [[7]], "formulas": [["=3+4"]], "excel_errors": [[None]],
        },
    }}
    native, ranges = _oracle_range_cells(oracle)
    assert ranges == ["Main!F34", "Premium!C1:D2"]
    addresses = ["Premium!C1", "Premium!D2", "Main!F34"]
    model = {"premium!C1": 2.5000000000001, "premium!D2": {"excel_error": "#N/A"}, "main!F34": 7}
    expected_formulas = {"premium!C1": "=1+1.5", "premium!D2": "=#N/A",
                         "main!F34": "=SUM(3,\n 4)"}
    native["main!F34"]["formula"] = "=SUM(3,4)"

    compared, identity_matches, max_abs, matched_addresses, mismatches = _compare_active_formula_members(
        addresses, native, model, expected_formulas, 1e-12, 1e-12)
    assert compared == 3
    assert identity_matches == 3
    assert max_abs is not None and 0 < max_abs < 1e-12
    assert matched_addresses == {"premium!C1", "premium!D2", "main!F34"}
    assert mismatches == []

    changed_formula = dict(native)
    changed_formula["main!F34"] = {**native["main!F34"], "formula": "=SUM(2,5)"}
    compared, identity_matches, _max_abs, _matched, mismatches = _compare_active_formula_members(
        addresses, changed_formula, model, expected_formulas, 1e-12, 1e-12)
    assert compared == 3
    assert identity_matches == 2
    assert any(item["kind"] == "active_formula_identity" and item["address"] == "Main!F34"
               for item in mismatches)
    assert not any(item["kind"] == "formula_value" and item["address"] == "Main!F34"
                   for item in mismatches)

    missing_formula = dict(native)
    missing_formula["main!F34"] = {**native["main!F34"], "formula": ""}
    compared, identity_matches, _max_abs, _matched, mismatches = _compare_active_formula_members(
        addresses, missing_formula, model, expected_formulas, 1e-12, 1e-12)
    assert compared == 3
    assert identity_matches == 2
    assert any(item["kind"] == "active_formula_identity" and item["address"] == "Main!F34"
               for item in mismatches)

    compared, _identity_matches, _max_abs, _matched_addresses, mismatches = _compare_active_formula_members(
        addresses, native, {"premium!C1": 2.5}, expected_formulas, 1e-12, 1e-12)
    assert compared == 1
    assert [item["kind"] for item in mismatches] == [
        "active_formula_missing_from_model", "active_formula_missing_from_model"]


def test_step5_oracle_requires_completed_native_calculation_state() -> None:
    oracle = {"schema_version": "gp.excel_oracle.v1", "tool": "step5.oracle", "status": "pass",
              "engine": "Microsoft Excel", "source_sha256": "source", "source_copy_sha256": "source",
              "stage4_artifact_sha256": "stage4", "macros_executed": False, "input_overrides": [],
              "primary_inputs": {address: {"value": 1} for address in
                                 ("C3", "C4", "C6", "C7", "C8", "J48", "O26")}}
    _preflight_oracle({**oracle, "calculation_state": 0}, "source", "stage4")
    for state in (None, 1, 2):
        try:
            _preflight_oracle({**oracle, "calculation_state": state}, "source", "stage4")
        except ValueError as exc:
            assert "calculation must be complete" in str(exc)
        else:
            raise AssertionError(f"native calculation state {state!r} must block reconciliation")


def test_stage5_active_member_set_comes_from_the_bound_trace(tmp_path: Path) -> None:
    trace_path = tmp_path / "active_trace.json"
    trace = {"source_sha256": "source-hash", "cells": [
        {"address": "Premium!J1", "role": "calculated_formula", "formula": "=1+1"},
        {"address": "CI_MaleChoosen!F9", "role": "calculated_array_formula", "formula": "=A1+1"},
        {"address": "Main!C3", "role": "source_value"},
    ]}
    trace_path.write_text(json.dumps(trace), encoding="utf-8")
    report = {"source_sha256": "source-hash",
              "active_trace": {"path": str(trace_path), "sha256": hash_file(trace_path)}}
    manifest = {"active_formula_addresses": ["Premium!J1", "CI_MaleChoosen!F9"]}

    assert _active_formula_addresses(report, manifest) == ["Premium!J1", "CI_MaleChoosen!F9"]
    stale = {**manifest, "active_formula_addresses": ["Premium!J1"]}
    try:
        _active_formula_addresses(report, stale)
    except ValueError as exc:
        assert "differ from the bound trace" in str(exc)
    else:
        raise AssertionError("a manifest that omits an active array follower must be rejected")


def test_legacy_formula_addresses_are_bound_to_formula_mode_source(tmp_path: Path) -> None:
    source_path = tmp_path / "legacy_source.xlsx"
    workbook = Workbook()
    workbook.active.title = "Main"
    workbook.active["A1"] = "=1+1"
    workbook.save(source_path)
    workbook.close()
    source_hash = hash_file(source_path)
    addresses, formulas = _active_formula_details(
        {}, {"active_formula_addresses": ["Main!A1"]}, source_path, source_hash)
    assert addresses == ["Main!A1"]
    assert formulas == {"main!A1": "=1+1"}

    try:
        _active_formula_details({}, {"active_formula_addresses": ["Main!A1"]}, source_path, "stale")
    except ValueError as exc:
        assert "differs from the bound workbook" in str(exc)
    else:
        raise AssertionError("a legacy formula address set must not detach from its source workbook")


def test_overlapping_native_ranges_must_agree_on_values_and_errors() -> None:
    base = {"qualified_address": "Premium!C1", "values": [[None]],
            "formulas": [["=#N/A"]], "excel_errors": [["#N/A"]]}
    oracle = {"ranges": {"Premium!C1": base, "Premium!C1:C1": dict(base)}}
    cells, _ranges = _oracle_range_cells(oracle)
    assert cells["premium!C1"]["value"] == {"excel_error": "#N/A"}

    conflicting = {"ranges": {"Premium!C1": base,
                               "Premium!C1:C1": {**base, "excel_errors": [[None]]}}}
    try:
        _oracle_range_cells(conflicting)
    except ValueError as exc:
        assert "overlapping native oracle ranges disagree" in str(exc)
    else:
        raise AssertionError("overlapping native ranges with different error semantics must be rejected")


def test_stage5_compares_producer_shaped_external_cuts_and_literal_native_axis(tmp_path: Path) -> None:
    bundle = tmp_path / "bundle"
    bundle.mkdir()
    values = {"external.rate": [0.1, 0.2]}
    age_values = [0, 1]
    formulas = ["=A1*0+0.1", "=A2*0+0.2"]
    source_path = tmp_path / "source.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Rates"
    worksheet["A1"], worksheet["A2"] = age_values
    worksheet["B1"], worksheet["B2"] = formulas
    workbook.save(source_path)
    workbook.close()
    source_hash = hash_file(source_path)
    # This matches the Step 4 producer: source identity is bound at payload level;
    # each cut member carries its own exact formula text and formula hash.
    members = [{"address": f"Rates!B{row}", "boundary_variable_id": "external.rate",
                "formula": formula,
                "formula_sha256": hashlib.sha256(formula.encode("utf-8")).hexdigest()}
               for row, formula in zip((1, 2), formulas)]
    values_hash, age_hash = _canonical_hash(values), _canonical_hash(age_values)
    external_payload = {
        "schema_version": "step4.external_values.v1", "source_sha256": source_hash,
        "capture_sha256": "capture-hash", "values_sha256": values_hash, "values": values,
        "boundaries": [{"boundary_variable_id": "external.rate", "source_range": "Rates!B1:B2",
                        "shape": [2]}],
        "formula_cut_members": members, "formula_cut_member_count": 2,
        "formula_cut_formula_set_sha256": "cut-set-hash",
        # The native producer does not duplicate the generated metadata-group binding.
        "age_axis": {"source_range": "Rates!A1:A2", "values": age_values, "value_sha256": age_hash},
    }
    (bundle / "external_values.json").write_text(json.dumps(external_payload), encoding="utf-8")
    (bundle / "input_layout.json").write_text(json.dumps({
        "external_bindings": [{"variable_id": "external.rate"}],
        "external_capture": {"age_axis": {"source_range": "Rates!A1:A2", "metadata_group": "axis.age"}},
    }), encoding="utf-8")
    (bundle / "adapter_metadata.json").write_text(json.dumps({"groups": {
        "axis.age": {"addresses": ["rates!A1", "rates!A2"], "values": age_values},
    }}), encoding="utf-8")
    manifest = {
        "source_sha256": source_hash, "external_bindings": [{"variable_id": "external.rate"}],
        "external_input_addresses": ["rates!B1", "rates!B2"], "external_input_count": 2,
        "external_capture_sha256": "capture-hash", "external_values_sha256": values_hash,
        "external_formula_cut_formula_set_sha256": "cut-set-hash",
        "external_formula_cut_member_count": 2, "external_age_axis_sha256": age_hash,
        "external_age_axis": {"source_range": "Rates!A1:A2", "values_sha256": age_hash,
                              "metadata_group": "axis.age"},
    }
    native = {
        # Excel COM returns literal constants through Range.Formula, e.g. "0"/"1".
        "rates!A1": {"address": "Rates!A1", "value": 0, "formula": "0"},
        "rates!A2": {"address": "Rates!A2", "value": 1, "formula": "1"},
        "rates!B1": {"address": "Rates!B1", "value": 0.1, "formula": formulas[0]},
        "rates!B2": {"address": "Rates!B2", "value": 0.2, "formula": formulas[1]},
    }
    model = {"external_inputs": values, "external_age_axis": age_values}
    metrics, mismatches = _compare_external_boundary(
        manifest, bundle, model, native, 1e-12, 1e-12, source_path)
    assert metrics == {"external_input_coordinates": 2, "external_age_axis_keys": 2,
                       "external_formula_cut_members": 2}
    assert mismatches == []

    wrong_source_manifest = {**manifest, "source_sha256": "b" * 64}
    _metrics, mismatches = _compare_external_boundary(
        wrong_source_manifest, bundle, model, native, 1e-12, 1e-12, source_path)
    assert any(item["kind"] == "external_boundary_binding" for item in mismatches)

    layout_path = bundle / "input_layout.json"
    layout_record = json.loads(layout_path.read_text(encoding="utf-8"))
    wrong_layout = json.loads(json.dumps(layout_record))
    wrong_layout["external_capture"]["age_axis"]["metadata_group"] = "axis.wrong"
    layout_path.write_text(json.dumps(wrong_layout), encoding="utf-8")
    _metrics, mismatches = _compare_external_boundary(
        manifest, bundle, model, native, 1e-12, 1e-12, source_path)
    assert any(item["kind"] == "external_age_axis_binding" for item in mismatches)
    layout_path.write_text(json.dumps(layout_record), encoding="utf-8")

    metadata_path = bundle / "adapter_metadata.json"
    metadata_record = json.loads(metadata_path.read_text(encoding="utf-8"))
    metadata_record["groups"]["axis.age"]["addresses"].reverse()
    metadata_path.write_text(json.dumps(metadata_record), encoding="utf-8")
    _metrics, mismatches = _compare_external_boundary(
        manifest, bundle, model, native, 1e-12, 1e-12, source_path)
    assert any(item["kind"] == "external_age_axis_bundle_value" for item in mismatches)
    metadata_record["groups"]["axis.age"]["addresses"].reverse()
    metadata_path.write_text(json.dumps(metadata_record), encoding="utf-8")

    changed_native = dict(native)
    changed_native["rates!B2"] = {**native["rates!B2"], "formula": "=A2/200"}
    metrics, mismatches = _compare_external_boundary(
        manifest, bundle, model, changed_native, 1e-12, 1e-12, source_path)
    assert any(item["kind"] == "external_formula_cut_changed" for item in mismatches)

    # A value in the source range that is actually a formula remains invalid, even
    # when Excel's calculated value matches the captured age key.
    formula_source_path = tmp_path / "formula_source.xlsx"
    formula_workbook = Workbook()
    formula_sheet = formula_workbook.active
    formula_sheet.title = "Rates"
    formula_sheet["A1"], formula_sheet["A2"] = 0, "=1"
    formula_sheet["B1"], formula_sheet["B2"] = formulas
    formula_workbook.save(formula_source_path)
    formula_workbook.close()
    changed_source_hash = hash_file(formula_source_path)
    external_payload["source_sha256"] = changed_source_hash
    (bundle / "external_values.json").write_text(
        json.dumps(external_payload), encoding="utf-8")
    manifest["source_sha256"] = changed_source_hash
    changed_native["rates!A2"] = {**native["rates!A2"], "formula": "=1"}
    _metrics, mismatches = _compare_external_boundary(
        manifest, bundle, model, changed_native, 1e-12, 1e-12, formula_source_path)
    assert any(item["kind"] == "external_age_axis_source_formula" for item in mismatches)


def test_step5_report_states_all_member_coverage_and_delegated_reviewer_pair() -> None:
    payload = {"status": "ready_for_review", "source_sha256": "source", "stage4_artifact_sha256": "stage4",
               "oracle": {"engine": "Microsoft Excel", "source_copy_sha256": "source",
                          "macros_executed": False, "input_overrides": []},
               "compared": {"primary_inputs": 7, "named_targets": 4, "required_targets": 4,
                            "active_formula_members": 23098, "required_active_formula_members": 23098,
                            "active_formula_identity_matches": 23098,
                            "required_active_formula_identities": 23098,
                            "active_formula_max_abs_difference": 2.22e-16,
                            "active_premium_formula_cells": 7767},
               "tolerances": {"absolute": 1e-12, "relative": 1e-12},
               "mismatch_count": 0, "mismatch_examples": [], "target_results": {},
               "path_coverage": {"implemented_path_ids": [], "numerical_path_ids": [],
                                 "runtime_verified_path_ids": [], "condition_only_path_ids": [],
                                 "runtime_evidence": [], "condition_evidence": []},
               "review_policy": {"required_reviewer_pair": ["agent", "typesafe"],
                                 "delegation": {"stages": [3, 4, 5, 6],
                                                "requested_model": "jev-latest",
                                                "authorization_sha256": "auth-sha"}}}

    markdown = _reconciliation_markdown(payload)
    assert "Active source formula members compared: 23098 / 23098" in markdown
    assert "Active source formula identities matched: 23098 / 23098" in markdown
    assert "Maximum absolute difference across numeric active formula values: 2.22e-16" in markdown
    assert "Agent and delegated TypeSafe confirmation are pending." in markdown
    assert "Agent and human confirmation are pending." not in markdown
    assert "TypeSafe delegation: stages [3, 4, 5, 6]" in markdown


def test_step5_reader_front_orders_requested_result_before_diagnostics_and_shows_tolerances() -> None:
    from excel_to_act.steps.step5.workflow import _reconciliation_markdown

    payload = {
        "status": "ready_for_review",
        "source_sha256": "source-hash",
        "stage4_artifact_sha256": "stage4-hash",
        "oracle": {"engine": "Fixture", "source_copy_sha256": "source-hash",
                   "macros_executed": False, "input_overrides": []},
        "compared": {"primary_inputs": 7, "named_targets": 2, "required_targets": 2,
                      "active_formula_members": 4, "required_active_formula_members": 4,
                      "active_formula_identity_matches": 4, "required_active_formula_identities": 4,
                      "active_formula_max_abs_difference": 0.0, "active_premium_formula_cells": 4,
                      "external_input_coordinates": 0, "required_external_input_coordinates": 0,
                      "external_formula_cut_members": 0, "required_external_formula_cut_members": 0},
        "mismatch_count": 0,
        "mismatch_examples": [],
        "tolerances": {"absolute": 1e-12, "relative": 2e-12},
        "target_results": {
            "PVFB": {"model": 14.7, "excel": 14.7, "matched": True},
            "GP": {"model": 2.7, "excel": 2.8, "matched": False},
        },
        "path_coverage": {"numerical_path_ids": [], "runtime_verified_path_ids": [],
                          "condition_only_path_ids": [], "condition_evidence": [],
                          "runtime_evidence": []},
        "review_policy": {"required_reviewer_pair": ["agent", "human"]},
    }
    design = {"targets": [{"selector": "GP", "result_order": 0,
                            "diagnostic_intermediates": ["PVFB"]}]}

    front = _reconciliation_markdown(payload, design=design).split("## Detailed comparison evidence", 1)[0]

    assert "| GP | Requested target | Not recorded | Not recorded | Not recorded | Not recorded | 2.7 | 2.8 | False |" in front
    assert "| PVFB | Diagnostic | Not recorded | Not recorded | Not recorded | Not recorded | 14.7 | 14.7 | True |" in front
    assert front.index("| GP | Requested target") < front.index("| PVFB | Diagnostic")
    assert "absolute=1e-12, relative=2e-12" in front



def test_step5_reader_shows_declared_shapes_ordered_array_values_and_unknowns_without_mutation() -> None:
    import copy

    payload = {
        "status": "ready_for_review", "source_sha256": "source-fixture",
        "stage4_artifact_sha256": "stage4-fixture", "mismatch_count": None,
        "oracle": {"engine": "fixture", "source_copy_sha256": "copy-fixture",
                   "macros_executed": False, "input_overrides": []},
        "compared": {"primary_inputs": 0, "named_targets": 3, "required_targets": 3,
                     "active_formula_members": 0, "required_active_formula_members": 0,
                     "active_formula_identity_matches": False, "required_active_formula_identities": None,
                     "active_formula_max_abs_difference": None, "active_premium_formula_cells": 0,
                     "external_input_coordinates": 0, "required_external_input_coordinates": 0,
                     "external_formula_cut_members": 0, "required_external_formula_cut_members": 0,
                     "external_age_axis_keys": 0, "required_external_age_axis_keys": 0},
        "tolerances": {"absolute": None, "relative": 0}, "mismatch_examples": [],
        "target_results": {
            "Diagnostic": {"model": [5, 6], "excel": [5, 6], "matched": True},
            "Matrix": {"model": [[1, 2], [3, 4]], "excel": [[1, 2], [3, 0]], "matched": False},
            "Scalar": {"model": False, "excel": False, "matched": None},
            "Vector": {"model": [0, False, None], "excel": [0, False, 8], "matched": True},
        },
        "path_coverage": {"numerical_path_ids": [], "runtime_verified_path_ids": [],
                          "condition_only_path_ids": [], "condition_evidence": []},
        "review_policy": {"required_reviewer_pair": ["agent", "human"]},
    }
    original = copy.deepcopy(payload)
    design = {
        "targets": [
            {"selector": "Scalar", "result_order": 2, "shape": [], "result_kind": None, "units": None},
            {"selector": "Vector", "result_order": 0, "shape": [3], "result_kind": "vector", "units": "USD",
             "axes": [{"name": "duration", "role": "projection", "keys": ["None", 0, False]}]},
            {"selector": "Matrix", "result_order": 1, "shape": [2, 2], "result_kind": "table", "units": "count",
             "axes": []},
            {"selector": "Other", "result_order": 3, "shape": None, "diagnostic_intermediates": ["Diagnostic"]},
        ],
        "open_questions": [
            {"question_id": "Q_CAPTURE", "question": "A capture was required before execution.",
             "answer": "Capture is recorded in the current run.", "status": "captured", "disposition": "execution prerequisite"},
            {"question_id": "Q_MEANING", "question": "What does the selected axis mean?",
             "answer": None, "status": None, "disposition": None},
        ],
    }
    front = _reconciliation_markdown(payload, design=design, machine_link="validation.json").split(
        "## Detailed comparison evidence", 1)[0]

    assert front.index("| Vector | Requested target") < front.index("| Matrix | Requested target") < front.index("| Scalar | Requested target")
    assert front.index("| Scalar | Requested target") < front.index("| Diagnostic | Diagnostic")
    assert "| Vector | Requested target | vector | [3] | USD |" in front
    assert "keys: position [0] = None" in front and "position [2] = False" in front
    assert "position [0] = 0" in front and "position [2] = Not recorded" in front
    assert "position [2] = 8" in front and "| True |" in front
    assert "| Scalar | Requested target | Not recorded | [] | Not recorded | Not recorded | False | False | Not recorded |" in front
    assert "position [0, 0] = 1" in front and "position [1, 1] = 4" in front
    assert "position [1, 0] = 3" in front and "position [1, 1] = 0" in front
    assert "Mismatch count: Not recorded" in front
    assert "Primary inputs matched: 0 / 7" in front and "absolute=Not recorded, relative=0" in front
    assert "Recorded answer: Capture is recorded in the current run." in front
    assert "execution prerequisite" in front and "**Q_MEANING**" in front
    assert "status: Not recorded; disposition: Not recorded" in front
    assert "Position labels do not imply a quarter, year, or other business-axis origin" in front
    assert "[validation_report.json](<validation.json>)" in front
    assert "[[1, 2], [3, 4]]" not in front and "[0, False, None]" not in front
    assert payload == original

def test_step5_accepts_modular_runtime_but_keeps_legacy_runtime_contract(tmp_path: Path) -> None:
    bundle = tmp_path / "bundle"
    bundle.mkdir()
    (bundle / "model.py").write_text("def main(): pass\n", encoding="utf-8")
    (bundle / "source_values.json").write_text("{}\n", encoding="utf-8")
    (bundle / "modular_runtime.py").write_text("VALUE = 1\n", encoding="utf-8")
    (bundle / "input_layout.json").write_text('{"external_bindings": []}\n', encoding="utf-8")
    (bundle / "input_adapter.py").write_text(
        "import hashlib\n"
        "def content_hash(payload):\n"
        "    return hashlib.sha256(payload).hexdigest()\n",
        encoding="utf-8",
    )
    modular_files = {path.name: hash_file(path) for path in bundle.iterdir()}
    modular_manifest = {
        "schema_version": "step4.modular_model.v1", "execution_kind": "modular",
        "files": modular_files, "raw_input_addresses": [], "formula_addresses": [],
        "formula_cache_inputs": False,
    }
    assert _verify_python(bundle, modular_manifest) == []

    legacy_manifest = {**modular_manifest, "schema_version": "step4.generated_model.v1",
                       "execution_kind": "legacy"}
    assert "generated bundle is missing runtime.py" in _verify_python(bundle, legacy_manifest)


def test_step5_validates_separate_external_ledger_and_disallows_overlap(tmp_path: Path) -> None:
    bundle = tmp_path / "external-bundle"
    bundle.mkdir()
    values = {"rates.external": [0.1, 0.2]}
    values_sha = _canonical_hash(values)
    layout = {
        "external_bindings": [{"variable_id": "rates.external", "source_range": "Rates!B1:B2",
                               "shape": [2], "addresses": ["rates!B1", "rates!B2"],
                               "indices": [[0], [1]]}],
        "external_capture": {"artifact_sha256": "artifact-sha", "capture_sha256": "capture-sha",
                             "values_sha256": values_sha},
    }
    external = {"schema_version": "step4.external_values.v1", "source_sha256": "source-sha",
                "capture_artifact_sha256": "artifact-sha", "capture_sha256": "capture-sha",
                "values_sha256": values_sha, "values": values}
    contents = {"model.py": "def main(): pass\n", "modular_runtime.py": "VALUE = 1\n",
                "source_values.json": "{}\n", "input_layout.json": json.dumps(layout),
                "external_values.json": json.dumps(external)}
    for name, content in contents.items():
        (bundle / name).write_text(content, encoding="utf-8")
    manifest = {
        "schema_version": "step4.modular_model.v1", "execution_kind": "modular",
        "source_sha256": "source-sha", "files": {name: hash_file(bundle / name) for name in contents},
        "raw_input_addresses": [], "formula_addresses": [],
        "external_input_addresses": ["Rates!B1", "Rates!B2"], "external_input_count": 2,
        "external_capture_artifact_sha256": "artifact-sha", "external_capture_sha256": "capture-sha",
        "external_values_sha256": values_sha, "formula_cache_inputs": False,
    }
    assert _verify_python(bundle, manifest) == []

    overlap = {**manifest, "formula_addresses": ["rates!B2"]}
    assert any("overlaps a raw or calculated formula address" in failure
               for failure in _verify_python(bundle, overlap))
    stale = {**manifest, "external_capture_artifact_sha256": "stale-artifact"}
    assert any("external_values.json is stale" in failure for failure in _verify_python(bundle, stale))


def test_step5_validates_a_synthetic_modular_generation_end_to_end(tmp_path: Path) -> None:
    root = tmp_path / "workflow"
    source = root / "source" / "fixture.xlsm"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"synthetic source identity; not loaded by standalone validation")
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    create_workflow(root, {"workbook_sha256": source_hash, "workbook_path": str(source),
                           "run_id": "synthetic-run", "source_id": "synthetic-source"})
    _, workflow = load_workflow(root)
    workflow["mode"] = "synthetic_fixture"
    _write_manifest(root, workflow)

    previous = None
    for stage in (1, 2, 3):
        revision = append_stage_artifact(
            root, stage, f"stage{stage}.json", {"status": "ready_for_review"},
            f"Synthetic stage {stage}.",
            input_stages={stage - 1: previous["artifact"]["json_sha256"]} if previous else None)
        for reviewer in ("agent", "human"):
            record_decision(root, stage, reviewer, "approve", f"Synthetic stage {stage} review.")
        previous = revision

    bundle = root / "stage4" / "bundles" / "revision-0001" / "bundle"
    bundle.mkdir(parents=True)
    files = {
        "model.py": (
            "import argparse, json\n"
            "from pathlib import Path\n"
            "def main():\n"
            "    parser = argparse.ArgumentParser()\n"
            "    parser.add_argument('--out', type=Path, default=Path('model_result.json'))\n"
            "    args = parser.parse_args()\n"
            "    args.out.write_text(json.dumps({'status':'pass','targets':{'GP':1},"
            "'cells':{'Main!A1':1},'cell_count':1,'formula_count':1}), encoding='utf-8')\n"
            "if __name__ == '__main__': main()\n"
        ),
        "modular_runtime.py": "VALUE = 1\n",
        "input_layout.json": '{"external_bindings": [], "external_capture": null}\n',
        "source_values.json": '{"main!B1": 2}\n',
    }
    for name, content in files.items():
        (bundle / name).write_text(content, encoding="utf-8")
    bundle_files = {name: hash_file(bundle / name) for name in files}
    manifest = {
        "schema_version": "step4.modular_model.v1", "execution_kind": "modular",
        "source_sha256": source_hash, "files": bundle_files,
        "formula_count": 1, "array_member_count": 0, "array_instance_count": 0,
        "formula_addresses": ["Main!A1"], "raw_input_addresses": ["main!B1"],
        "active_formula_addresses": ["Main!A1"], "formula_cache_inputs": False,
    }
    (bundle / "model_manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    report_files = {**bundle_files, "model_manifest.json": hash_file(bundle / "model_manifest.json")}
    append_stage_artifact(
        root, 4, "generation_report.json",
        {"schema_version": "step4.generation_report.v1", "status": "ready_for_review",
         "source_sha256": source_hash, "bundle": {"path": str(bundle), "files": report_files}},
        "Synthetic modular generation.", input_stages={3: previous["artifact"]["json_sha256"]})
    for reviewer in ("agent", "human"):
        record_decision(root, 4, reviewer, "approve", "Synthetic Stage 4 review.")

    result = validate_generated(root)
    assert result["status"] == "pass", result
    evidence = json.loads(Path(result["validation"]).read_text(encoding="utf-8"))
    assert evidence["status"] == "pass"
    assert evidence["result_summary"]["targets"] == {"GP": 1}

    current_root, current_workflow = load_workflow(root)
    stage4_hash = current_workflow["stages"]["4"]["revisions"][-1]["artifact"]["json_sha256"]
    stage5_root = current_root / "stage5"
    incomplete_oracle_path = stage5_root / "oracle-incomplete.json"
    incomplete_oracle_path.write_text(json.dumps({
        "schema_version": "gp.excel_oracle.v1", "tool": "step5.oracle", "status": "pass",
        "engine": "Microsoft Excel", "source_sha256": source_hash, "source_copy_sha256": source_hash,
        "stage4_artifact_sha256": stage4_hash, "macros_executed": False, "input_overrides": [],
        "primary_inputs": {address: {"value": 1} for address in
                           ("C3", "C4", "C6", "C7", "C8", "J48", "O26")},
        # calculation_state intentionally omitted
    }), encoding="utf-8")
    revisions_before = len(current_workflow["stages"]["5"]["revisions"])
    blocked = reconcile(root, Path(result["validation"]), incomplete_oracle_path)
    assert blocked["status"] == "blocked"
    assert "calculation must be complete" in blocked["reason"]
    _current_root, after_workflow = load_workflow(root)
    assert len(after_workflow["stages"]["5"]["revisions"]) == revisions_before
