from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import sys

import pytest
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.formula import ArrayFormula

from excel_to_act.steps.step3.calculation import CalculationBlocked
from excel_to_act.steps.step4.generator import _smoke_bundle
from excel_to_act.steps.step4.modular_bundle import (
    _augment_literal_sources,
    _named_bindings,
    _raw_bindings,
    add_adapter_metadata,
    build_modular_bundle_files,
)
from excel_to_act.steps.step4.modular_compiler import build_source_cell_index
from excel_to_act.steps.step4.modular_runtime import binary
from excel_to_act.steps.step4 import modular_runtime as runtime


def _workbook(path: Path) -> str:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    for row, key in enumerate((1, 2, 3), 1):
        sheet.cell(row, 1).value = key
        sheet.cell(row, 2).value = key + 10
        sheet.cell(row, 3).value = f"=A{row}*10"
    sheet["D1"] = "=VLOOKUP(2,RawTable,3,FALSE)"
    sheet["F1"] = "No"
    sheet["E1"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
    sheet["G1"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
    workbook.save(path)
    workbook.close()
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _input_records():
    return [
        {"variable_id": "raw.key", "role": "raw", "kind": "projection_series", "shape": [3],
         "axes": [{"name": "year"}], "dependencies": [], "initial_condition": None,
         "source_extents": [{"ref": "Main!A1:A3", "role": "raw_input",
             "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        {"variable_id": "raw.value", "role": "raw", "kind": "projection_series", "shape": [3],
         "axes": [{"name": "year"}], "dependencies": [], "initial_condition": None,
         "source_extents": [{"ref": "Main!B1:B3", "role": "raw_input",
             "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        {"variable_id": "raw.config", "role": "raw", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [], "initial_condition": None,
         "source_extents": [{"ref": "Main!F1", "role": "raw_input",
             "coordinate_mapping": {"kind": "scalar", "fixed_axes": []}}]},
        {"variable_id": "table.value", "role": "derived", "kind": "projection_series", "shape": [3],
         "axes": [{"name": "year"}], "dependencies": [{"variable_id": "raw.key"}],
         "initial_condition": None,
         "source_extents": [{"ref": "Main!C1:C3", "role": "formula_output",
             "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}],
         "equation_segments": [{"segment_id": "value_pass", "source_family_ids": ["family-value"],
             "equation_family_ids": ["eq-value"], "index_axis": "year", "index_start": 0,
             "index_stop_exclusive": 3, "execution": "vectorized", "recurrence_group": None,
             "snapshot_before_update": False}]},
        {"variable_id": "model.answer", "role": "derived", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [{"variable_id": "raw.key"}, {"variable_id": "table.value"}],
         "initial_condition": None,
         "source_extents": [{"ref": "Main!D1", "role": "formula_output",
             "coordinate_mapping": {"kind": "scalar", "fixed_axes": []}}],
         "equation_segments": [{"segment_id": "answer_pass", "source_family_ids": ["family-answer"],
             "equation_family_ids": ["eq-answer"], "index_axis": None, "index_start": 0,
             "index_stop_exclusive": 1, "execution": "scalar", "recurrence_group": None,
             "snapshot_before_update": False}]},
    ]


def test_modular_bundle_runs_named_table_and_grouped_family_driver_in_isolation(tmp_path: Path) -> None:
    source_path = tmp_path / "source.xlsx"
    source_hash = _workbook(source_path)
    variables = _input_records()
    formulas = [
        ("Main!C1", "=A1*10"), ("Main!C2", "=A2*10"), ("Main!C3", "=A3*10"),
        ("Main!D1", "=VLOOKUP(2,RawTable,3,FALSE)"),
    ]
    trace = {"source_sha256": source_hash, "names": [
        {"scope": "workbook", "name": "RawTable"}, {"scope": "workbook", "name": "Answer"}],
        "cells": [
            *[{"address": f"Main!A{row}", "sheet": "Main", "role": "source_value"} for row in range(1, 4)],
            *[{"address": f"Main!B{row}", "sheet": "Main", "role": "source_value"} for row in range(1, 4)],
            *[{"address": address, "sheet": "Main", "role": "calculated_formula", "formula": formula}
              for address, formula in formulas],
        ]}
    plan = {
        "schema_version": "step3.semantic_plan.v1", "source": {"source_sha256": source_hash},
        "variables": variables,
        "equation_families": [
            {"family_id": "eq-value", "function_name": "table_value"},
            {"family_id": "eq-answer", "function_name": "answer_value"},
        ],
        "source_mappings": [
            {"source_family_id": "family-value", "equation_family_id": "eq-value",
             "variable_id": "table.value", "source_members": ["Main!C1", "Main!C2", "Main!C3"],
             "member_count": 3},
            {"source_family_id": "family-answer", "equation_family_id": "eq-answer",
             "variable_id": "model.answer", "source_members": ["Main!D1"], "member_count": 1},
        ],
        "array_mappings": [],
        "execution_plan": {"passes": [
            {"pass_id": "table-values", "source_family_ids": ["family-value"]},
            {"pass_id": "answer", "source_family_ids": ["family-answer"]},
        ]},
        "coverage": {
            "ordinary_formula_members_mapped": 4,
            "array_formula_followers_mapped": 0,
            "active_formula_members_mapped": 4,
            "source_family_count": 2,
            "semantic_family_count": 2,
            "array_instance_count": 0,
        },
        "reference_resolution": {"named_references": [
            {"name": "RawTable", "kind": "range", "source_ref": "Main!A1:C3",
             "table_binding": {"kind": "column_overlay_table", "key_variable_id": "raw.key",
                               "value_variable_id": "raw.value", "value_column_index": 2,
                               "formula_columns": {"C": "table.value"}}},
            {"name": "Answer", "kind": "scalar", "source_ref": "Main!D1",
             "variable_id": "model.answer", "resolution": "derived_formula_output"},
        ]},
    }
    profile = {"source_sha256": source_hash, "families": [
        {"family_id": "family-value", "source_members": ["Main!C1", "Main!C2", "Main!C3"],
         "member_count": 3},
        {"family_id": "family-answer", "source_members": ["Main!D1"], "member_count": 1},
    ]}
    raw_values = {f"main!A{row}": row for row in range(1, 4)}
    raw_values.update({f"main!B{row}": row + 10 for row in range(1, 4)})
    # The compiler must source every declared raw binding from the workbook, even
    # when the active formula trace omitted a literal cell from its frontier.
    trace["cells"] = [cell for cell in trace["cells"] if cell["address"] != "Main!A2"]
    raw_values.pop("main!A2")
    output = tmp_path / "revision"
    built = build_modular_bundle_files(
        semantic_plan=plan, source_profile=profile, trace=trace, raw_source_values=raw_values,
        source_path=source_path, semantic_plan_path=tmp_path / "semantic_plan.json",
        source_profile_path=tmp_path / "source_profile.json", semantic_map_path=tmp_path / "semantic_map.json",
        revision_dir=output, target_names=["Answer"], path_coverage={"implemented_path_ids": []},
        source_sha256=source_hash, design_sha256="d" * 64,
        semantic_plan_sha256="p" * 64, source_profile_sha256="f" * 64,
        trace_sha256="t" * 64, trace_path=tmp_path / "trace.json", trace_source="synthetic_fixture",
        scenario_ids=["synthetic"],
    )

    bundle = Path(built["bundle"])
    result = _smoke_bundle(bundle, output / "smoke", ["Answer"])
    payload = json.loads(Path(result["result_path"]).read_text(encoding="utf-8"))
    assert payload["targets"]["Answer"] == 20
    assert payload["calculated_variables"]["table.value"] == {
        "shape": [3],
        "demanded_values": [
            {"indices": [0], "value": 10},
            {"indices": [1], "value": 20},
            {"indices": [2], "value": 30},
        ],
    }
    assert payload["calculated_variables"]["model.answer"]["demanded_values"] == [
        {"indices": [], "value": 20}]
    assert payload["calculated_value_count"] == 4
    assert payload["cells"] == {"main!C1": 10, "main!C2": 20, "main!C3": 30, "main!D1": 20}
    assert built["manifest"]["compiled_variant_count"] == 2
    assert built["manifest"]["array_instance_count"] == built["coverage"]["array_instance_count"] == 0
    assert built["coverage"]["active_formula_member_count"] == 4
    assert "Main!" not in (bundle / "pricing.py").read_text(encoding="utf-8")
    assert "excel_to_act" not in "\n".join(path.read_text(encoding="utf-8")
                                            for path in bundle.glob("*.py"))
    model_source = (bundle / "model.py").read_text(encoding="utf-8")
    assert "main!A1" not in model_source and "main!F1" not in model_source
    raw_bundle = json.loads((bundle / "source_values.json").read_text(encoding="utf-8"))
    bundle_manifest = json.loads((bundle / "model_manifest.json").read_text(encoding="utf-8"))
    assert bundle_manifest["array_instance_count"] == built["coverage"]["array_instance_count"]
    assert set(raw_bundle) == {"main!A1", "main!A2", "main!A3", "main!B1", "main!B2", "main!B3", "main!F1"}
    assert raw_bundle["main!A2"] == 2
    layout = json.loads((bundle / "input_layout.json").read_text(encoding="utf-8"))
    assert layout["scenario_guard"]["source_values"]["main!A1"] == 1
    assert built["manifest"]["scenario_guard_coordinate_count"] == 7
    layout_path = bundle / "input_layout.json"
    layout["scenario_guard"]["source_values"]["main!A1"] = -1
    layout_path.write_text(json.dumps(layout), encoding="utf-8")
    changed_layout = subprocess.run(
        [sys.executable, "model.py", "--out", "changed_layout.json"], cwd=bundle,
        capture_output=True, text=True, check=False,
    )
    assert changed_layout.returncode != 0
    assert "bundle scenario guard differs from its generated contract" in changed_layout.stderr
    layout["scenario_guard"]["source_values"]["main!A1"] = 1
    layout_path.write_text(json.dumps(layout), encoding="utf-8")
    raw_bundle["main!A1"] = 99
    (bundle / "source_values.json").write_text(json.dumps(raw_bundle), encoding="utf-8")
    changed_scenario = subprocess.run(
        [sys.executable, "model.py", "--out", "changed.json"], cwd=bundle,
        capture_output=True, text=True, check=False,
    )
    assert changed_scenario.returncode != 0
    assert "saved-scenario source selector changed" in changed_scenario.stderr
    raw_bundle["main!A1"] = 1
    raw_bundle["main!F1"] = "Yes"
    (bundle / "source_values.json").write_text(json.dumps(raw_bundle), encoding="utf-8")
    changed_configuration = subprocess.run(
        [sys.executable, "model.py", "--out", "changed_config.json"], cwd=bundle,
        capture_output=True, text=True, check=False,
    )
    assert changed_configuration.returncode != 0
    assert "saved-scenario source selector changed: main!F1" in changed_configuration.stderr


def test_approved_stored_empty_adapter_metadata_reaches_bundle_and_unapproved_blank_blocks(tmp_path: Path) -> None:
    source_path = tmp_path / "source.xlsx"
    source_hash = _workbook(source_path)
    variables = _input_records()
    formulas = [
        ("Main!C1", "=A1*10"), ("Main!C2", "=A2*10"), ("Main!C3", "=A3*10"),
        ("Main!D1", "=VLOOKUP(2,RawTable,3,FALSE)"),
    ]
    base_cells = [
        *[{"address": f"Main!A{row}", "sheet": "Main", "role": "source_value", "value": row}
          for row in range(1, 4)],
        *[{"address": f"Main!B{row}", "sheet": "Main", "role": "source_value", "value": row + 10}
          for row in range(1, 4)],
        *[{"address": address, "sheet": "Main", "role": "calculated_formula", "formula": formula}
          for address, formula in formulas],
    ]
    trace = {"source_sha256": source_hash, "names": [
        {"scope": "workbook", "name": "RawTable"}, {"scope": "workbook", "name": "Answer"}],
        "cells": [*base_cells, {"address": "Main!E1", "sheet": "Main", "role": "source_value", "value": None}]}
    plan = {
        "schema_version": "step3.semantic_plan.v1", "source": {"source_sha256": source_hash},
        "variables": variables,
        "equation_families": [
            {"family_id": "eq-value", "function_name": "table_value"},
            {"family_id": "eq-answer", "function_name": "answer_value"},
        ],
        "source_mappings": [
            {"source_family_id": "family-value", "equation_family_id": "eq-value",
             "variable_id": "table.value", "source_members": ["Main!C1", "Main!C2", "Main!C3"],
             "member_count": 3},
            {"source_family_id": "family-answer", "equation_family_id": "eq-answer",
             "variable_id": "model.answer", "source_members": ["Main!D1"], "member_count": 1},
        ],
        "array_mappings": [],
        "execution_plan": {"passes": [
            {"pass_id": "table-values", "source_family_ids": ["family-value"]},
            {"pass_id": "answer", "source_family_ids": ["family-answer"]},
        ]},
        "coverage": {"ordinary_formula_members_mapped": 4, "array_formula_followers_mapped": 0,
                     "active_formula_members_mapped": 4, "source_family_count": 2,
                     "semantic_family_count": 2, "array_instance_count": 0},
        "reference_resolution": {"named_references": [
            {"name": "RawTable", "kind": "range", "source_ref": "Main!A1:C3",
             "table_binding": {"kind": "column_overlay_table", "key_variable_id": "raw.key",
                               "value_variable_id": "raw.value", "value_column_index": 2,
                               "formula_columns": {"C": "table.value"}}},
            {"name": "Answer", "kind": "scalar", "source_ref": "Main!D1",
             "variable_id": "model.answer", "resolution": "derived_formula_output"},
        ]},
    }
    profile = {"source_sha256": source_hash, "families": [
        {"family_id": "family-value", "source_members": ["Main!C1", "Main!C2", "Main!C3"],
         "member_count": 3},
        {"family_id": "family-answer", "source_members": ["Main!D1"], "member_count": 1},
    ]}
    scope = {"allowed_regions": [{"source_ref": "Main!E1",
                                  "rationale": "source-audited stored-empty adapter boundary"}]}
    candidate_trace = {"unknowns": [
        {"category": "unrepresented_source_coordinate", "evidence": "Main!E1"}]}
    blank_plan = json.loads(json.dumps(plan))
    metadata = add_adapter_metadata(blank_plan, {"axis_metadata": [], "source_records": []},
                                    scope, trace, source_path, source_hash, candidate_trace)
    assert metadata["source_coordinate_count"] == 1
    assert blank_plan["adapter_metadata_groups"]["source.verified_empty"]["addresses"] == ["main!E1"]

    output = tmp_path / "approved-empty-revision"
    built = build_modular_bundle_files(
        semantic_plan=blank_plan, source_profile=profile, trace=trace,
        raw_source_values={f"main!A{row}": row for row in range(1, 4)}
                         | {f"main!B{row}": row + 10 for row in range(1, 4)},
        source_path=source_path, semantic_plan_path=tmp_path / "semantic_plan.json",
        source_profile_path=tmp_path / "source_profile.json", semantic_map_path=tmp_path / "semantic_map.json",
        revision_dir=output, target_names=["Answer"], path_coverage={"implemented_path_ids": []},
        source_sha256=source_hash, design_sha256="d" * 64,
        semantic_plan_sha256="p" * 64, source_profile_sha256="f" * 64,
        trace_sha256="t" * 64, trace_path=tmp_path / "trace.json", trace_source="synthetic_fixture",
        scenario_ids=["synthetic"],
    )
    metadata_file = json.loads((Path(built["bundle"]) / "adapter_metadata.json").read_text(encoding="utf-8"))
    assert metadata_file["groups"]["source.verified_empty"]["values"] == [None]

    unapproved_trace = {**trace, "cells": [*trace["cells"], {
        "address": "Main!G1", "sheet": "Main", "role": "source_value", "value": None}]}
    unapproved_plan = json.loads(json.dumps(plan))
    add_adapter_metadata(unapproved_plan, {"axis_metadata": [], "source_records": []},
                         scope, unapproved_trace, source_path, source_hash, candidate_trace)
    with pytest.raises(CalculationBlocked, match="source values are outside the approved logical inputs and metadata: main!G1"):
        build_modular_bundle_files(
            semantic_plan=unapproved_plan, source_profile=profile, trace=unapproved_trace,
            raw_source_values={f"main!A{row}": row for row in range(1, 4)}
                             | {f"main!B{row}": row + 10 for row in range(1, 4)},
            source_path=source_path, semantic_plan_path=tmp_path / "semantic_plan.json",
            source_profile_path=tmp_path / "source_profile.json", semantic_map_path=tmp_path / "semantic_map.json",
            revision_dir=tmp_path / "unapproved-empty-revision", target_names=["Answer"],
            path_coverage={"implemented_path_ids": []}, source_sha256=source_hash,
            design_sha256="d" * 64, semantic_plan_sha256="p" * 64,
            source_profile_sha256="f" * 64, trace_sha256="u" * 64,
            trace_path=tmp_path / "unapproved-trace.json", trace_source="synthetic_fixture",
            scenario_ids=["synthetic"],
        )


def test_axis_metadata_initial_condition_is_checked_metadata_not_raw_input(tmp_path: Path) -> None:
    source_path = tmp_path / "axis-seed.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    sheet["A1"] = 7
    sheet["A2"] = "=A1*2"
    sheet["B1"] = "=A1+E1"
    sheet["E1"] = 11
    workbook.save(source_path)
    workbook.close()
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()

    variables = [
        {"variable_id": "raw.operand", "role": "raw", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [], "initial_condition": None,
         "source_extents": [{"ref": "Main!E1", "role": "raw_input",
                              "coordinate_mapping": {"kind": "scalar", "fixed_axes": []}}]},
        {"variable_id": "raw.seed_alias", "role": "derived", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [],
         "initial_condition": {"seed_ref": "Main!E1", "source_role": "raw_input_seed"},
         "source_extents": []},
        {"variable_id": "amr.age_key", "role": "derived", "kind": "projection_series", "shape": [2],
         "axes": [{"name": "age_index"}], "dependencies": [{"variable_id": "amr.age_key"}],
         "initial_condition": {"seed_ref": "Main!A1", "logical_index": 0,
                               "source_role": "accepted_axis_metadata"},
         "source_extents": [{"ref": "Main!A2", "role": "formula_output",
                              "coordinate_mapping": {"kind": "row_ordinal", "axes": [
                                  {"axis": "age_index", "dimension": 0,
                                   "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}],
         "equation_segments": [{"segment_id": "age_update", "source_family_ids": ["family-age"],
                                "equation_family_ids": ["eq-age"], "index_axis": "age_index",
                                "index_start": 1, "index_stop_exclusive": 2,
                                "execution": "vectorized", "recurrence_group": None,
                                "snapshot_before_update": False}]},
        {"variable_id": "model.answer", "role": "derived", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [{"variable_id": "amr.age_key"},
                                       {"variable_id": "raw.operand"}],
         "initial_condition": None,
         "source_extents": [{"ref": "Main!B1", "role": "formula_output",
                              "coordinate_mapping": {"kind": "scalar", "fixed_axes": []}}],
         "equation_segments": [{"segment_id": "answer", "source_family_ids": ["family-answer"],
                                "equation_family_ids": ["eq-answer"], "index_axis": None,
                                "index_start": 0, "index_stop_exclusive": 1,
                                "execution": "scalar", "recurrence_group": None,
                                "snapshot_before_update": False}]},
    ]
    plan = {
        "schema_version": "step3.semantic_plan.v1", "source": {"source_sha256": source_hash},
        "variables": variables,
        "equation_families": [{"family_id": "eq-age", "function_name": "age_update"},
                              {"family_id": "eq-answer", "function_name": "answer"}],
        "source_mappings": [
            {"source_family_id": "family-age", "equation_family_id": "eq-age",
             "variable_id": "amr.age_key", "source_members": ["Main!A2"], "member_count": 1},
            {"source_family_id": "family-answer", "equation_family_id": "eq-answer",
             "variable_id": "model.answer", "source_members": ["Main!B1"], "member_count": 1},
        ],
        "array_mappings": [],
        "execution_plan": {"passes": [
            {"pass_id": "age", "source_family_ids": ["family-age"]},
            {"pass_id": "answer", "source_family_ids": ["family-answer"]},
        ]},
        "coverage": {"ordinary_formula_members_mapped": 2, "array_formula_followers_mapped": 0,
                     "active_formula_members_mapped": 2, "source_family_count": 2,
                     "semantic_family_count": 2, "array_instance_count": 0},
        "reference_resolution": {"named_references": [
            {"name": "Answer", "kind": "scalar", "source_ref": "Main!B1",
             "variable_id": "model.answer", "resolution": "derived_formula_output"}]},
    }
    formulas = {"Main!A2": "=A1*2", "Main!B1": "=A1+E1"}
    trace = {"source_sha256": source_hash,
             "names": [{"scope": "workbook", "name": "Answer", "destination": "Main!B1"}],
             "cells": [
                 {"address": "Main!A1", "sheet": "Main", "role": "source_value", "value": 7},
                 {"address": "Main!E1", "sheet": "Main", "role": "source_value", "value": 11},
                 *[{"address": address, "sheet": "Main", "role": "calculated_formula", "formula": formula}
                   for address, formula in formulas.items()],
             ]}
    profile = {"source_sha256": source_hash, "families": [
        {"family_id": "family-age", "source_members": ["Main!A2"], "member_count": 1},
        {"family_id": "family-answer", "source_members": ["Main!B1"], "member_count": 1},
    ]}
    projection = add_adapter_metadata(plan, {"axis_metadata": [{"axis_id": "age_origin",
                                                                  "source_extents": [{"sheet": "Main", "range": "A1"}]}],
                                             "source_records": []},
                                      {"allowed_regions": []}, trace, source_path, source_hash, {"unknowns": []})
    assert projection["raw_source_coordinate_count"] == 1
    assert projection["source_metadata_coordinate_count"] == 1
    assert projection["metadata_seed_binding_count"] == 1
    assert plan["metadata_seed_bindings"] == [{
        "variable_id": "amr.age_key", "indices": [0], "metadata_group": "axis.age_origin",
        "metadata_indices": [0], "source_address": "main!A1"}]
    source_index = build_source_cell_index(plan)
    raw_bindings = _raw_bindings(plan, source_index, {item["variable_id"]: item for item in variables})
    assert {item["address"].casefold() for item in raw_bindings} == {"main!e1"}
    assert {item["variable_id"] for item in raw_bindings} == {"raw.operand", "raw.seed_alias"}

    revision = tmp_path / "bundle-revision"
    built = build_modular_bundle_files(
        semantic_plan=plan, source_profile=profile, trace=trace,
        raw_source_values={"main!E1": 11}, source_path=source_path,
        semantic_plan_path=tmp_path / "semantic_plan.json", source_profile_path=tmp_path / "profile.json",
        semantic_map_path=tmp_path / "map.json", revision_dir=revision, target_names=["Answer"],
        path_coverage={"implemented_path_ids": []}, source_sha256=source_hash,
        design_sha256="d" * 64, semantic_plan_sha256="p" * 64,
        source_profile_sha256="f" * 64, trace_sha256="t" * 64,
        trace_path=tmp_path / "trace.json", trace_source="synthetic_fixture", scenario_ids=["saved"],
    )
    bundle = Path(built["bundle"])
    result = _smoke_bundle(bundle, revision / "smoke", ["Answer"])
    output = json.loads(Path(result["result_path"]).read_text(encoding="utf-8"))
    assert output["targets"]["Answer"] == 18
    assert built["manifest"]["raw_source_coordinate_count"] == 1
    assert built["manifest"]["source_metadata_coordinate_count"] == 1
    assert built["manifest"]["metadata_seed_binding_count"] == 1
    raw_values = json.loads((bundle / "source_values.json").read_text(encoding="utf-8"))
    assert raw_values == {"main!E1": 11}
    metadata = json.loads((bundle / "adapter_metadata.json").read_text(encoding="utf-8"))
    original_metadata_bytes = (bundle / "adapter_metadata.json").read_bytes()
    assert metadata["groups"]["axis.age_origin"]["values"] == [7]

    def run_model(label: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run([sys.executable, "model.py", "--out", f"{label}.json"], cwd=bundle,
                              capture_output=True, text=True, check=False)

    metadata["groups"].clear()
    (bundle / "adapter_metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
    missing_group = run_model("missing_group")
    assert missing_group.returncode != 0
    assert "adapter metadata file is stale or incomplete" in missing_group.stderr
    metadata["groups"]["axis.age_origin"] = {"addresses": ["main!A1"], "values": [7],
                                               "sha256": "bad"}
    (bundle / "adapter_metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
    changed_group = run_model("changed_group")
    assert changed_group.returncode != 0
    assert "adapter metadata group changed: axis.age_origin" in changed_group.stderr
    (bundle / "adapter_metadata.json").write_bytes(original_metadata_bytes)
    layout_path = bundle / "input_layout.json"
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    layout["metadata_seed_bindings"] = []
    layout_path.write_text(json.dumps(layout), encoding="utf-8")
    changed_binding = run_model("changed_seed_binding")
    assert changed_binding.returncode != 0
    assert "metadata seed bindings are missing or differ" in changed_binding.stderr


def test_named_scalar_can_bind_external_coordinate_or_declared_source_literal() -> None:
    plan = {"variables": [
        {"variable_id": "external.rate", "role": "external", "shape": [2],
         "axes": [{"name": "age"}], "source_extents": [{
             "ref": "Rates!W5:W6", "role": "formula_output",
             "coordinate_mapping": {"kind": "row_ordinal", "axes": [{
                 "axis": "age", "dimension": 0, "source_coordinate": "row",
                 "origin": 5, "index_origin": 0}]}}]},
        {"variable_id": "raw.rate_header", "role": "raw", "shape": [], "axes": [],
         "source_extents": [{"ref": "Rates!W3", "role": "adapter_metadata",
             "coordinate_mapping": {"kind": "scalar", "axes": []}}]},
    ], "source_mappings": [], "array_mappings": [], "reference_resolution": {
        "named_references": [
            {"name": "SelectedRate", "kind": "scalar", "source_ref": "Rates!W5",
             "binding": {"kind": "variable_coordinate", "variable_id": "external.rate", "indices": [0]}},
            {"name": "SelectedRateDirect", "kind": "variable_coordinate", "source_ref": "Rates!W5",
             "variable_id": "external.rate", "indices": [0]},
            {"name": "RateLabel", "kind": "scalar", "source_ref": "Rates!W3",
             "binding": {"kind": "source_literal", "address": "Rates!W3"}},
        ]}}

    index = build_source_cell_index(plan)
    names = _named_bindings(plan, index, {"rates!W3"})
    values = {"external.rate": [0.25, 0.5], "__source_values__": {"rates!W3": "CI_2017_M"},
              "__names__": names}

    assert runtime.lookup_name(values, "SelectedRate") == 0.25
    assert runtime.lookup_name(values, "SelectedRateDirect") == 0.25
    assert runtime.lookup_name(values, "RateLabel") == "CI_2017_M"
    values["__source_values__"].clear()
    with pytest.raises(runtime.CalculationBlocked, match="missing"):
        runtime.lookup_name(values, "RateLabel")


def test_raw_binding_that_is_now_a_formula_is_rejected(tmp_path: Path) -> None:
    source_path = tmp_path / "source.xlsx"
    source_hash = _workbook(source_path)
    plan = {"source": {"source_sha256": source_hash}, "variables": [
        {"variable_id": "raw.formula", "role": "raw", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [], "initial_condition": None,
         "source_extents": [{"ref": "Main!C1", "role": "raw_input",
                              "coordinate_mapping": {"kind": "scalar", "fixed_axes": []}}]},
    ]}
    with pytest.raises(CalculationBlocked, match="declared raw source boundary is a formula cell"):
        _augment_literal_sources(source_path, source_hash, plan, {"cells": []}, {})


def test_raw_binding_that_overlaps_an_array_follower_is_rejected(tmp_path: Path) -> None:
    source_path = tmp_path / "array.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    sheet["A1"] = 1
    sheet["A2"] = 2
    sheet["C1"] = ArrayFormula(ref="C1:C2", text="=A1:A2")
    workbook.save(source_path)
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    plan = {"source": {"source_sha256": source_hash}, "variables": [
        {"variable_id": "raw.array_follower", "role": "raw", "kind": "scalar", "shape": [],
         "axes": [], "dependencies": [], "initial_condition": None,
         "source_extents": [{"ref": "Main!C2", "role": "raw_input",
                              "coordinate_mapping": {"kind": "scalar", "fixed_axes": []}}]},
    ]}
    trace = {"cells": [{"address": "Main!C2", "role": "calculated_array_formula",
                        "array_ref": "C1:C2"}]}
    with pytest.raises(CalculationBlocked, match="formula or array member"):
        _augment_literal_sources(source_path, source_hash, plan, trace, {})


def test_blank_cells_compare_as_zero_in_numeric_conditions() -> None:
    assert binary("<=", lambda: None, lambda: 1) is True
    assert binary(">", lambda: 1, lambda: None) is True
    assert binary("=", lambda: None, lambda: 0) is True
    assert binary("<", lambda: None, lambda: None) is False
