from __future__ import annotations

import pytest
from openpyxl.utils import get_column_letter

from excel_to_act.steps.step3.calculation import CalculationBlocked, _parse_reference
from excel_to_act.steps.step4 import modular_runtime as runtime
from excel_to_act.steps.step4.modular_compiler import (
    SourceCell,
    _build_execution_schedule,
    build_source_cell_index,
    compile_formula_function,
    compile_mapped_families,
    lower_mapped_reference,
    render_grouped_driver,
    validate_family_coverage,
)


def _reference(address: str) -> str:
    if address == "Main!A1":
        return "runtime.logical_reference(values['policy.issue_age'], (), ())"
    if address == "Main!B1:B2":
        return "runtime.logical_reference(values['policy.year_keys'], (0,), (2,))"
    if address == "Main!C1":
        return "runtime.logical_reference(values['policy.premium'], (), ())"
    raise AssertionError(f"unexpected source reference {address}")


def _name(name: str) -> str:
    raise CalculationBlocked(f"unmapped name: {name}")


def _compile(formula: str, *, resolver=_reference, names=_name, indices=()) -> str:
    return compile_formula_function(
        "source_equation", formula, "Main", "Main!D1",
        lambda ref, sheet, owner: resolver(ref.address),
        lambda name, sheet, owner: names(name),
        index_parameters=indices,
    )


def _execute(source: str, values: dict[str, object]) -> object:
    namespace = {"runtime": runtime}
    exec(compile(source, "formula_families.py", "exec"), namespace)
    return namespace["source_equation"](values)


def test_direct_family_function_uses_logical_values_and_python_literals() -> None:
    source = _compile("=IF(A1>0,SUM(B1:B2)*2,IFERROR(C1/0,0))")

    assert "Main!" not in source
    assert "A1" not in source
    assert "env.get" not in source
    assert _execute(source, {"policy.issue_age": 1, "policy.year_keys": [2, 3],
                             "policy.premium": 9}) == 10
    assert _execute(source, {"policy.issue_age": 0, "policy.year_keys": [2, 3],
                             "policy.premium": 9}) == 0


def test_false_and_blank_operands_emit_valid_python_and_if_is_lazy() -> None:
    source = compile_formula_function(
        "source_equation", "=IF(FALSE,1/0,\"\")", "Main", "Main!D1",
        lambda ref, sheet, owner: "None",
        lambda name, sheet, owner: "None",
    )

    assert "False" in source and "None" not in source
    assert _execute(source, {}) == ""

    ifna = _compile("=_xlfn.IFNA(1,2)")
    assert _execute(ifna, {}) == 1


def test_dynamic_reference_is_lazy_but_active_unmapped_name_blocks() -> None:
    source = compile_formula_function(
        "source_equation", '=IF(FALSE,INDIRECT("Missing"),7)', "Main", "Main!D1",
        lambda ref, sheet, owner: "None",
        lambda name, sheet, owner: "None",
    )
    assert _execute(source, {}) == 7

    active = compile_formula_function(
        "source_equation", '=IFERROR(INDIRECT("Missing"),3)', "Main", "Main!D1",
        lambda ref, sheet, owner: "None",
        lambda name, sheet, owner: "None",
    )
    with pytest.raises(runtime.CalculationBlocked, match="not mapped"):
        _execute(active, {})


def test_formula_result_scalarizes_cell_references_and_preserves_range_values() -> None:
    scalar = _compile("=IF(TRUE,A1,0)")
    assert _execute(scalar, {"policy.issue_age": 10}) == 10

    vector = _compile("=IF(TRUE,B1:B2,0)")
    assert _execute(vector, {"policy.year_keys": [10, 20]}) == [10, 20]


def test_transpose_coerces_only_blank_referenced_cells_to_zero() -> None:
    source_values = [None, "nonblank", 4.5, {"excel_error": "#N/A"}]
    source_range = runtime.logical_reference(source_values, (0,), (4,), source_shape=(1, 4))

    result = runtime.transpose(lambda: source_range)

    assert result[:3] == [[0], ["nonblank"], [4.5]]
    assert isinstance(result[3][0], runtime.ExcelError)
    assert result[3][0].code == "#N/A"
    assert source_values[0] is None

    table_view = runtime.table_view([[None, "nonblank"], [4.5, {"excel_error": "#N/A"}]])
    table_result = runtime.transpose(lambda: table_view)
    assert table_result == [[0, 4.5], ["nonblank", runtime.ExcelError("#N/A")]]
    assert table_view.rows[0][0] is None
    assert runtime.transpose(lambda: [None, "nonblank"]) == [[None, "nonblank"]]


def test_column_overlay_vlookup_forces_only_key_and_selected_return_column() -> None:
    forced: list[str] = []
    table = runtime.ColumnOverlayTable(
        row_count=3, column_count=4,
        columns={
            1: [1, 2, 3],
            2: lambda: [None, 0.4, 0.5],
            4: lambda: (forced.append("inactive") or [10, 20, 30]),
        },
    )

    assert runtime.vlookup(lambda: 2, lambda: table, lambda: 2, lambda: False) == 0.4
    assert runtime.vlookup(lambda: 1, lambda: table, lambda: 2, lambda: False) == 0
    assert forced == []
    with pytest.raises(runtime.CalculationBlocked, match="column 3 is unavailable"):
        runtime.vlookup(lambda: 2, lambda: table, lambda: 3, lambda: False)
    assert forced == []
    assert runtime.vlookup(lambda: 9, lambda: table, lambda: 3, lambda: False).code == "#N/A"
    assert forced == []
    with pytest.raises(runtime.CalculationBlocked, match="column 3 is unavailable"):
        runtime.excel_iferror(
            lambda: runtime.vlookup(lambda: 2, lambda: table, lambda: 3, lambda: False),
            lambda: 99,
        )
    assert isinstance(runtime.vlookup(lambda: 0, lambda: table, lambda: 2), runtime.ExcelError)

    mixed_header = runtime.ColumnOverlayTable(
        row_count=3, column_count=2, columns={1: ["ppp", 1, 2], 2: [10, 20, 30]})
    assert runtime.vlookup(lambda: 0, lambda: mixed_header, lambda: 2).code == "#N/A"


def test_column_overlay_vlookup_reads_only_the_selected_return_row() -> None:
    values = {
        "__source_values__": {
            "amr_table!a7": 1,
            "amr_table!a8": 2,
        },
        "test.array_wrapper": [runtime.LogicalReference([10, 20], (1,), ())],
        "__names__": {
            "amr_table": {
                "kind": "table",
                "row_count": 2,
                "column_count": 2,
                "columns": {
                    "1": [
                        {"address": "amr_table!a7"},
                        {"address": "amr_table!a8"},
                    ],
                    "2": [
                        {"missing_address": "amr_table!b6"},
                        {"variable_id": "test.array_wrapper", "indices": [0]},
                    ],
                },
            },
        },
    }
    table = runtime.lookup_name(values, "AMR_Table")

    # B6 is absent from the approved inputs, but key 2 selects B7 and must not
    # force that unrelated header cell.
    assert runtime.vlookup(lambda: 2, lambda: table, lambda: 2, lambda: False) == 20

    # If the lookup actually selects the absent cell, preserve the explicit
    # source-missing blocker instead of substituting a header or blank.
    with pytest.raises(runtime.CalculationBlocked, match="amr_table!b6"):
        runtime.vlookup(lambda: 1, lambda: table, lambda: 2, lambda: False)


def test_match_and_vlookup_propagate_excel_errors_from_mode_arguments() -> None:
    missing = runtime.ExcelError("#N/A")
    table = runtime.TableView(((1, 10), (2, 20)))

    assert runtime.match(lambda: 1, lambda: [1, 2], lambda: missing).code == "#N/A"
    assert runtime.vlookup(lambda: 1, lambda: table, lambda: missing).code == "#N/A"
    assert runtime.vlookup(lambda: 1, lambda: table, lambda: 2, lambda: missing).code == "#N/A"


def test_unsupported_external_reference_is_deferred_to_the_selected_branch() -> None:
    external = "'[External.xlsx]Sheet1'!A1"
    inactive = _compile(f"=IF(FALSE,{external},7)")
    assert _execute(inactive, {}) == 7

    active = _compile(f"=IF(TRUE,{external},7)")
    with pytest.raises(runtime.CalculationBlocked, match="not mapped"):
        _execute(active, {})

    wrapped = _compile(f"=IFERROR({external},3)")
    with pytest.raises(runtime.CalculationBlocked, match="not mapped"):
        _execute(wrapped, {})


def test_unbound_formula_name_is_deferred_to_the_selected_branch() -> None:
    inactive = _compile("=IF(FALSE,AnnuityPaymentFix,7)")
    assert _execute(inactive, {}) == 7

    active = _compile("=IF(TRUE,AnnuityPaymentFix,7)")
    with pytest.raises(runtime.CalculationBlocked, match="not mapped"):
        _execute(active, {})

    wrapped = _compile("=IFERROR(AnnuityPaymentFix,3)")
    with pytest.raises(runtime.CalculationBlocked, match="not mapped"):
        _execute(wrapped, {})


def test_offset_range_uses_logical_vector_and_active_unsupported_function_blocks() -> None:
    source = compile_formula_function(
        "source_equation", "=SUM(A1:OFFSET(A1,2,0))", "Main", "Main!D1",
        lambda ref, sheet, owner: "runtime.logical_reference(values['policy.year_keys'], (0,), ())",
        lambda name, sheet, owner: "None",
    )
    namespace = {"runtime": runtime}
    exec(compile(source, "formula_families.py", "exec"), namespace)
    assert namespace["source_equation"]({"policy.year_keys": [10, 20, 30, 40]}) == 60

    with pytest.raises(CalculationBlocked, match="not supported"):
        compile_formula_function(
            "source_equation", "=INDIRECT(A1)&NOW()", "Main", "Main!D1",
            lambda ref, sheet, owner: "None",
            lambda name, sheet, owner: "None",
        )


def test_formula_compiler_rejects_unmapped_and_bad_names() -> None:
    with pytest.raises(CalculationBlocked, match="not supported"):
        _compile("=NOW()")
    with pytest.raises(ValueError, match="function name"):
        compile_formula_function(
            "bad-name", "=1", "Main", "Main!D1",
            lambda ref, sheet, owner: "None", lambda name, sheet, owner: "None")


def _seed_overlap_plan(formula_owner: str, *, owner_depends_on_seed: bool) -> dict[str, object]:
    seed = {
        "variable_id": "input.seed", "role": "raw", "kind": "scalar", "shape": [],
        "axes": [], "dependencies": [], "source_extents": [{
            "ref": "Main!A1", "role": "raw_seed",
            "coordinate_mapping": {"kind": "scalar", "axes": []},
        }],
    }
    series = {
        "variable_id": "model.series", "role": "derived", "kind": "projection_series",
        "shape": [2], "axes": [{"name": "step"}],
        "dependencies": ([{"variable_id": "input.seed"}, {"variable_id": "model.series", "lag": 1}]
                         if owner_depends_on_seed else []),
        "initial_condition": {"seed_ref": "Main!A1", "logical_index": 0},
        "source_extents": [{
            "ref": "Main!A2", "role": "formula_output",
            "coordinate_mapping": {"kind": "projection_series", "axes": [
                {"axis": "step", "dimension": 0, "source_coordinate": "row",
                 "origin": 1, "index_origin": 0}], "fixed_axes": []},
        }],
    }
    if formula_owner == "model.series":
        output = series
        formula_address = "Main!A2"
    else:
        output = {
            "variable_id": formula_owner, "role": "derived", "kind": "scalar", "shape": [],
            "axes": [], "dependencies": [], "source_extents": [{
                "ref": "Main!B1", "role": "formula_output",
                "coordinate_mapping": {"kind": "scalar", "axes": []},
            }],
        }
        formula_address = "Main!B1"
    return {
        "variables": [seed, series] if output is series else [seed, series, output],
        "equation_families": [{"family_id": "eq.seed_test", "function_name": "calculate_seed_test"}],
        "source_mappings": [{"source_family_id": "seed-test", "equation_family_id": "eq.seed_test",
                             "variable_id": formula_owner, "source_members": [formula_address]}],
        "array_mappings": [],
    }


@pytest.mark.parametrize(
    ("formula_owner", "owner_depends_on_seed"),
    [("model.series", True), ("loading.amount", False)],
)
def test_declared_initial_condition_owns_overlapping_raw_seed_reference(
    formula_owner: str, owner_depends_on_seed: bool,
) -> None:
    plan = _seed_overlap_plan(formula_owner, owner_depends_on_seed=owner_depends_on_seed)
    formula_address = "Main!A2" if formula_owner == "model.series" else "Main!B1"
    seed_cells = build_source_cell_index(plan)[("main", 1, 1)]
    assert {(cell.variable_id, cell.role) for cell in seed_cells} == {
        ("input.seed", "raw_seed"), ("model.series", "initial_condition")}
    trace = {
        "cells": [{"address": formula_address, "formula": "=A1+1", "role": "calculated_formula"}],
        "edges": [{"consumer": formula_address, "prerequisite": "Main!A1",
                   "relationship": "cell_reference"}],
        "names": [],
    }

    compiled = compile_mapped_families(
        plan, trace, lambda name, sheet, owner: (_ for _ in ()).throw(CalculationBlocked(f"unmapped {name}")))

    assert compiled["source_family_bindings"]["seed-test"] == ["model.series"]
    assert "runtime.unsupported_reference" not in compiled["functions"]["calculate_seed_test"]


def test_sparse_formula_extent_does_not_shadow_raw_source_gap() -> None:
    row_axis = [{"axis": "row", "dimension": 0, "source_coordinate": "row",
                 "origin": 1, "index_origin": 0}]
    plan = {
        "variables": [
            {"variable_id": "raw.bridge", "role": "raw", "shape": [3], "axes": [{"name": "row"}],
             "dependencies": [], "source_extents": [{"ref": "Main!A1:A3", "role": "raw_source_bridge",
                 "coordinate_mapping": {"kind": "row_ordinal", "axes": row_axis}}]},
            {"variable_id": "derived.sparse", "role": "derived", "shape": [3],
             "axes": [{"name": "row"}], "dependencies": [], "source_extents": [{
                 "ref": "Main!A1:A3", "role": "formula_output",
                 "coordinate_mapping": {"kind": "row_ordinal", "axes": row_axis}}]},
            {"variable_id": "model.result", "role": "derived", "shape": [], "axes": [],
             "dependencies": [{"variable_id": "derived.sparse"}, {"variable_id": "raw.bridge"}],
             "source_extents": [{"ref": "Main!B1", "role": "formula_output",
                 "coordinate_mapping": {"kind": "scalar", "axes": []}}]},
        ],
        "equation_families": [
            {"family_id": "eq.sparse", "function_name": "calculate_sparse"},
            {"family_id": "eq.result", "function_name": "calculate_result"},
        ],
        "source_mappings": [
            {"source_family_id": "sparse-family", "equation_family_id": "eq.sparse",
             "variable_id": "derived.sparse", "source_members": ["Main!A1", "Main!A3"]},
            {"source_family_id": "result-family", "equation_family_id": "eq.result",
             "variable_id": "model.result", "source_members": ["Main!B1"]},
        ],
        "array_mappings": [],
    }
    trace = {
        "cells": [
            {"address": "Main!A1", "formula": "=10", "role": "calculated_formula"},
            {"address": "Main!A3", "formula": "=30", "role": "calculated_formula"},
            {"address": "Main!B1", "formula": "=A2+1", "role": "calculated_formula"},
        ],
        "edges": [{"consumer": "Main!B1", "prerequisite": "Main!A2",
                   "relationship": "cell_reference"}],
        "names": [],
    }

    source_index = build_source_cell_index(plan)
    assert [(cell.variable_id, cell.role) for cell in source_index[("main", 2, 1)]] == [
        ("raw.bridge", "raw_source_bridge")]
    compiled = compile_mapped_families(plan, trace,
                                       lambda name, sheet, owner: (_ for _ in ()).throw(CalculationBlocked(name)))

    assert compiled["source_family_bindings"]["result-family"] == ["raw.bridge"]


def _unsupported_preflight_plan(formula: str) -> tuple[dict[str, object], dict[str, object]]:
    plan = {
        "variables": [{"variable_id": "model.output", "role": "derived", "shape": [],
                       "axes": [], "dependencies": [], "source_extents": [{
                           "ref": "Main!B1", "role": "formula_output",
                           "coordinate_mapping": {"kind": "scalar", "axes": []},
                       }]}],
        "equation_families": [{"family_id": "eq.preflight", "function_name": "calculate_preflight"}],
        "source_mappings": [{"source_family_id": "preflight", "equation_family_id": "eq.preflight",
                             "variable_id": "model.output", "source_members": ["Main!B1"]}],
        "array_mappings": [],
    }
    trace = {"cells": [{"address": "Main!B1", "formula": formula,
                        "role": "calculated_formula"}], "edges": [], "names": []}
    return plan, trace


def test_compiler_preflight_blocks_unsupported_active_reference_and_name() -> None:
    plan, trace = _unsupported_preflight_plan("=A1+1")
    trace["edges"] = [{"consumer": "Main!B1", "prerequisite": "Main!A1",
                       "relationship": "cell_reference"}]
    with pytest.raises(CalculationBlocked, match="active source trace has unsupported emitted reference/name"):
        compile_mapped_families(plan, trace,
                                lambda name, sheet, owner: (_ for _ in ()).throw(CalculationBlocked(name)))

    plan, trace = _unsupported_preflight_plan("=Mystery+1")
    trace["names"] = [{"name": "Mystery", "destination": "Main!A1"}]
    with pytest.raises(CalculationBlocked, match="active source trace has unsupported emitted reference/name"):
        compile_mapped_families(plan, trace,
                                lambda name, sheet, owner: (_ for _ in ()).throw(CalculationBlocked(name)))


def test_compiler_preflight_keeps_unselected_unsupported_reference_lazy() -> None:
    plan, trace = _unsupported_preflight_plan("=IF(FALSE,'[External.xlsx]Sheet1'!A1,7)")
    compiled = compile_mapped_families(
        plan, trace, lambda name, sheet, owner: (_ for _ in ()).throw(CalculationBlocked(name)))

    assert compiled["deferred_unsupported_references"] == [{
        "source_family_id": "preflight", "formula_cell": "main!B1",
        "reference": "'[External.xlsx]Sheet1'!A1",
    }]


def _coverage_inputs(*, array: bool = False):
    source_sha = "a" * 64
    profile = {"source_sha256": source_sha, "families": [
        {"family_id": "source-1", "member_count": 1 if array else 2,
         "source_members": ["Main!A1"] if array else ["Main!A1", "Main!A2"]},
    ]}
    ordinary = [{"address": "Main!A1", "formula": "=1", "role": "calculated_formula"}]
    followers = []
    array_mappings = []
    if array:
        formula = "=TRANSPOSE(Main!B1:B2)"
        ordinary[0]["formula"] = formula
        followers = [{"address": "Main!B1", "formula": formula,
                      "role": "calculated_array_formula", "array_ref": "A1:B2"}]
        array_mappings = [{"array_family_id": "array-1", "instances": [{
            "anchor": "Main!A1", "anchor_source_family_id": "source-1", "array_ref": "A1:B2",
            "source_formula": formula, "source_members": ["Main!B1"], "member_count": 1,
            "shape": [2, 2],
        }]}]
    else:
        ordinary.append({"address": "Main!A2", "formula": "=1", "role": "calculated_formula"})
    mapped_members = ["Main!A1"] if array else ["Main!A1", "Main!A2"]
    mapped_count = len(mapped_members)
    follower_count = len(followers)
    plan = {
        "source": {"source_sha256": source_sha},
        "source_mappings": [{"source_family_id": "source-1", "source_members": mapped_members,
                             "member_count": mapped_count, "variable_id": "values",
                             "equation_family_id": "eq-1"}],
        "variables": [{"variable_id": "values"}],
        "equation_families": [{"family_id": "eq-1"}],
        "array_mappings": array_mappings,
        "coverage": {"ordinary_formula_members_mapped": mapped_count,
                     "array_formula_followers_mapped": follower_count,
                     "active_formula_members_mapped": mapped_count + follower_count,
                     "source_family_count": 1, "semantic_family_count": 1,
                     "array_instance_count": 1 if array else 0},
    }
    trace = {"source_sha256": source_sha, "cells": ordinary + followers}
    return plan, profile, trace


def test_source_map_coverage_requires_exact_profile_and_trace_members() -> None:
    plan, profile, trace = _coverage_inputs()
    assert validate_family_coverage(plan, profile, trace) == {
        "ordinary_formula_count": 2, "array_follower_count": 0,
        "active_formula_member_count": 2, "source_family_count": 1,
        "semantic_family_count": 1, "array_instance_count": 0,
    }

    plan["source_mappings"][0]["source_members"].pop()
    plan["source_mappings"][0]["member_count"] = 1
    with pytest.raises(CalculationBlocked, match="differs from its profile"):
        validate_family_coverage(plan, profile, trace)


def test_array_follower_demand_may_be_a_checked_subset_of_array_shape() -> None:
    plan, profile, trace = _coverage_inputs(array=True)
    assert validate_family_coverage(plan, profile, trace)["active_formula_member_count"] == 2

    trace["cells"].append({"address": "Main!B2", "formula": "=TRANSPOSE(Main!B1:B2)",
                           "role": "calculated_array_formula", "array_ref": "A1:B2"})
    with pytest.raises(CalculationBlocked, match="array-follower coverage differ"):
        validate_family_coverage(plan, profile, trace)


def test_source_coordinates_lower_to_named_vector_indices() -> None:
    plan = {"variables": [
        {"variable_id": "raw.rate", "shape": [3], "dependencies": [], "source_extents": [
            {"ref": "Main!A1:A3", "role": "literal_input", "coordinate_mapping": {
                "kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                    "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        {"variable_id": "model.rate", "shape": [3], "dependencies": [
            {"variable_id": "raw.rate"}, {"variable_id": "model.rate", "lag": 1}],
         "source_extents": [{"ref": "Main!B1:B3", "role": "formula_output", "coordinate_mapping": {
             "kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
    ]}
    source_index = build_source_cell_index(plan)
    compiled = compile_formula_function(
        "source_equation", "=A2+10", "Main", "Main!B2",
        lambda reference, sheet, owner: lower_mapped_reference(
            reference, owner_variable_id="model.rate", owner_indices=(1,), owner_axes=("year",),
            semantic_plan=plan, source_index=source_index),
        lambda name, sheet, owner: (_ for _ in ()).throw(CalculationBlocked("unmapped name")),
        index_parameters=("i0",),
    )

    namespace = {"runtime": runtime}
    exec(compile(compiled, "formula_families.py", "exec"), namespace)
    assert namespace["source_equation"]({"raw.rate": [1, 2, 3]}, 1) == 12
    assert "Main!" not in compiled and "A2" not in compiled


def test_approved_adapter_metadata_reference_lowers_to_grouped_logical_slot() -> None:
    plan = {
        "adapter_metadata_groups": {"premium.labels": {"addresses": ["Premium!AL7"]}},
        "adapter_metadata_cells": [{"address": "Premium!AL7", "group_id": "premium.labels",
                                    "indices": [0], "axes": [None]}],
        "variables": [{"variable_id": "premium.rate", "shape": [1], "axes": [{"name": "year"}],
                       "dependencies": []}],
    }
    reference = _parse_reference("Premium!AL7", "Premium")
    assert reference is not None
    source_index = build_source_cell_index(plan)
    expression = lower_mapped_reference(
        reference, owner_variable_id="premium.rate", owner_indices=(0,), owner_axes=("year",),
        semantic_plan=plan, source_index=source_index,
    )

    assert expression == "runtime.adapter_metadata(values, 'premium.labels', (0,))"
    assert "Premium!AL7" not in expression
    assert runtime.adapter_metadata({"__adapter_metadata__": {"premium.labels": ["MultipleCIInc"]}},
                                    "premium.labels", (0,)) == "MultipleCIInc"


def test_translated_formula_rows_share_metadata_axis_reference_body() -> None:
    plan = {"variables": [{"variable_id": "premium.rate", "shape": [2],
                           "axes": [{"name": "projection_year"}], "dependencies": []}],
            "adapter_metadata_groups": {
                "axis.premium_year": {"addresses": ["Premium!B10", "Premium!B11"]}},
            "adapter_metadata_cells": [
                {"address": "Premium!B10", "group_id": "axis.premium_year",
                 "indices": [0], "axes": ["projection_year"]},
                {"address": "Premium!B11", "group_id": "axis.premium_year",
                 "indices": [1], "axes": ["projection_year"]},
            ]}
    source_index = build_source_cell_index(plan)
    expressions = []
    for address, index in (("Premium!B10", 0), ("Premium!B11", 1)):
        reference = _parse_reference(address, "Premium")
        assert reference is not None
        expressions.append(lower_mapped_reference(
            reference, owner_variable_id="premium.rate", owner_indices=(index,),
            owner_axes=("projection_year",), semantic_plan=plan, source_index=source_index))
    assert expressions[0] == expressions[1] == (
        "runtime.adapter_metadata(values, 'axis.premium_year', (i0,))")
    assert "Premium!B10" not in expressions[0] and "Premium!B11" not in expressions[0]


def _compile_array_transpose(formula: str, inputs: list[dict[str, object]],
                             output_shape: list[int],
                             function_name: str = "transpose_array") -> tuple[dict[str, object],
                                                                               dict[str, object],
                                                                               str, dict[str, int]]:
    axes = [{"name": "array_row"}, {"name": "component"}]
    output_extent = {
        "ref": "Main!A1", "role": "formula_output",
        "coordinate_mapping": {"kind": "array_anchor", "fixed_axes": [
            {"dimension": 0, "index": 0, "axis": "array_row"},
            {"dimension": 1, "index": 0, "axis": "component"},
        ]},
    }
    variables = [*inputs, {
        "variable_id": "output.array", "shape": output_shape, "axes": axes,
        "dependencies": [{"variable_id": item["variable_id"]} for item in inputs],
        "equation_segments": [{"segment_id": "transpose_array", "source_family_ids": ["array-anchor"],
            "equation_family_ids": ["transpose"], "index_axis": "array_row", "index_start": 0,
            "index_stop_exclusive": output_shape[0], "execution": "vectorized", "recurrence_group": None,
            "snapshot_before_update": False}],
        "source_extents": [output_extent],
    }]
    array_ref = f"A1:{get_column_letter(output_shape[1])}{output_shape[0]}"
    follower_addresses = [f"Main!{get_column_letter(column)}{row}"
                          for row in range(1, output_shape[0] + 1)
                          for column in range(1, output_shape[1] + 1)
                          if (row, column) != (1, 1)]
    plan = {
        "source": {"source_sha256": "a" * 64},
        "variables": variables,
        "equation_families": {function_name: {"family_id": function_name,
                                                "function_name": function_name}},
        "source_mappings": [{"source_family_id": "array-anchor", "equation_family_id": function_name,
                             "variable_id": "output.array", "source_members": ["Main!A1"],
                             "member_count": 1}],
        "array_mappings": [{"array_family_id": "transpose-array", "instances": [{
            "anchor": "Main!A1", "anchor_source_family_id": "array-anchor", "array_ref": array_ref,
            "variable_id": "output.array", "source_formula": formula, "source_members": follower_addresses,
            "member_count": len(follower_addresses), "shape": output_shape,
        }]}],
        "execution_plan": {"passes": [{"pass_id": "transpose", "source_family_ids": ["array-anchor"]}]},
        "coverage": {"ordinary_formula_members_mapped": 1,
                     "array_formula_followers_mapped": len(follower_addresses),
                     "active_formula_members_mapped": 1 + len(follower_addresses),
                     "source_family_count": 1, "semantic_family_count": 1, "array_instance_count": 1},
    }
    trace = {"source_sha256": "a" * 64, "cells": [
        {"address": "Main!A1", "formula": formula, "role": "calculated_formula"},
        *[{"address": address, "formula": formula, "role": "calculated_array_formula",
           "array_ref": array_ref} for address in follower_addresses],
    ]}
    profile = {"source_sha256": "a" * 64, "families": [{"family_id": "array-anchor",
               "member_count": 1, "source_members": ["Main!A1"]}]}
    coverage = validate_family_coverage(plan, profile, trace)
    compiled = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    namespace = {"runtime": runtime}
    exec(compile(compiled["functions"][function_name], "formula_families.py", "exec"), namespace)
    driver = render_grouped_driver(plan, compiled)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    return namespace, compiled, driver, coverage


def test_transpose_mixed_raw_and_computed_columns_uses_table_view_and_preserves_array_offsets() -> None:
    inputs = [
        {"variable_id": "input.raw_f", "shape": [24], "axes": [{"name": "component"}],
         "dependencies": [], "source_extents": [{"ref": "Main!F1:F24", "role": "literal_input",
             "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "component", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        # This fixture supplies the already-computed logical vector directly;
        # production formula_output extents are restricted to mapped members.
        {"variable_id": "derived.g", "shape": [24], "axes": [{"name": "component"}],
         "dependencies": [], "source_extents": [{"ref": "Main!G1:G24", "role": "logical_input",
             "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "component", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
    ]
    namespace, compiled, driver, coverage = _compile_array_transpose(
        "=TRANSPOSE(F1:G24)", inputs, [2, 24], function_name="transpose_ci_headers")
    raw_f = [100 + index for index in range(24)]
    raw_f[1] = None
    values = {"input.raw_f": raw_f,
              "derived.g": [1000 + index for index in range(24)], "output.array": None}
    result = namespace["run_pricing"](values, namespace)["output.array"]

    assert len(result) == 2 and all(len(row) == 24 for row in result)
    assert result[0][0] == 100  # Anchor (0, 0).
    assert result[0][1] == 0  # A blank TableView reference is zero in TRANSPOSE.
    assert values["input.raw_f"][1] is None  # Coercion does not rewrite source inputs.
    assert result[0][23] == 123  # First-row follower at offset (0, 23).
    assert result[1][0] == 1000  # Second-row follower at offset (1, 0).
    assert result[1][23] == 1023  # Final follower at offset (1, 23).
    assert "runtime.table_view" in compiled["functions"]["transpose_ci_headers"]
    assert "Main!" not in driver
    assert coverage["array_follower_count"] == 47
    assert coverage["active_formula_member_count"] == 48


def test_transpose_computed_vertical_vector_preserves_one_by_24_shape() -> None:
    inputs = [{"variable_id": "derived.j", "shape": [24], "axes": [{"name": "component"}],
        "dependencies": [], "source_extents": [{"ref": "Main!J1:J24", "role": "logical_input",
            "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "component", "dimension": 0,
                "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]}]
    namespace, compiled, driver, coverage = _compile_array_transpose("=TRANSPOSE(J1:J24)", inputs, [1, 24])
    source = [10 * (index + 1) for index in range(24)]
    values = {"derived.j": source, "output.array": None}
    result = namespace["run_pricing"](values, namespace)["output.array"]

    assert len(result) == 1 and len(result[0]) == 24
    assert result[0][0] == 10
    assert result[0][23] == 240
    assert "source_shape=(24, 1)" in compiled["functions"]["transpose_array"]
    assert coverage["array_follower_count"] == 23
    assert coverage["active_formula_member_count"] == 24


def test_transpose_preserves_excel_errors_and_blocks_unresolved_lazy_tables() -> None:
    error = runtime.ExcelError("#N/A")
    assert runtime.transpose(lambda: error) is error

    table = runtime.ColumnOverlayTable(row_count=2, column_count=2, columns={1: [1, 2]})
    with pytest.raises(runtime.CalculationBlocked, match="unresolved lazy table overlay"):
        runtime.transpose(lambda: table)


def test_unmapped_axis_label_extent_is_not_treated_as_a_value_coordinate() -> None:
    plan = {"variables": [{"variable_id": "state.grid", "shape": [4, 2], "dependencies": [],
        "source_extents": [{"ref": "Main!A1:B1", "role": "state_label_axis", "coordinate_mapping": {
            "kind": "column_ordinal", "axes": [{"axis": "state", "dimension": 1,
                "source_coordinate": "column", "origin": "A", "index_origin": 0}]}}]}]}
    assert build_source_cell_index(plan) == {}


def test_copied_source_family_compiles_once_to_a_direct_indexed_function() -> None:
    plan = {"variables": [
        {"variable_id": "raw.rate", "shape": [3], "dependencies": [], "source_extents": [
            {"ref": "Main!A1:A3", "role": "literal_input", "coordinate_mapping": {
                "kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                    "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        {"variable_id": "model.rate", "shape": [3], "dependencies": [{"variable_id": "raw.rate"}],
         "axes": [{"name": "year"}], "source_extents": [{"ref": "Main!B1:B3", "role": "formula_output",
            "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year", "dimension": 0,
                "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
    ], "equation_families": [{"family_id": "eq.rate", "function_name": "rate_step"}],
        "source_mappings": [{"source_family_id": "source.rate", "equation_family_id": "eq.rate",
            "variable_id": "model.rate", "source_members": ["Main!B2", "Main!B3"]}]}
    trace = {"cells": [
        {"address": "Main!B2", "formula": "=A2+10", "role": "calculated_formula"},
        {"address": "Main!B3", "formula": "=A3+10", "role": "calculated_formula"},
    ]}

    result = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")

    assert result["function_count"] == 1
    assert result["compiled_member_count"] == 2
    assert result["source_family_to_function"] == {"source.rate": "rate_step"}
    assert "Main!" not in result["functions"]["rate_step"]
    namespace = {"runtime": runtime}
    exec(compile(result["functions"]["rate_step"], "formula_families.py", "exec"), namespace)
    assert namespace["rate_step"]({"raw.rate": [1, 2, 3]}, ["raw.rate"], 2) == 13


@pytest.mark.parametrize(
    ("formulas", "expected_expression", "expected_value"),
    [
        (["=A1+10", "=A2+10", "=A3+10"], "i0", 15),
        (["=A$1+10", "=A$1+10", "=A$1+10"], "(0,)", 13),
    ],
)
def test_family_alignment_uses_translated_source_indices_but_keeps_fixed_rows(
    formulas: list[str], expected_expression: str, expected_value: int,
) -> None:
    plan = {"variables": [
        {"variable_id": "raw.other", "shape": [3], "dependencies": [], "source_extents": [
            {"ref": "Main!A1:A3", "role": "literal_input", "coordinate_mapping": {
                "kind": "row_ordinal", "axes": [{"axis": "source_row", "dimension": 0,
                    "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        {"variable_id": "model.age", "shape": [3], "dependencies": [{"variable_id": "raw.other"}],
         "axes": [{"name": "attained_age"}], "source_extents": [
            {"ref": "Main!B1:B3", "role": "formula_output", "coordinate_mapping": {
                "kind": "row_ordinal", "axes": [{"axis": "attained_age", "dimension": 0,
                    "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
    ], "equation_families": [{"family_id": "eq.age", "function_name": "age_step"}],
        "source_mappings": [{"source_family_id": "source.age", "equation_family_id": "eq.age",
            "variable_id": "model.age", "source_members": ["Main!B1", "Main!B2", "Main!B3"]}]}
    trace = {"cells": [{"address": f"Main!B{row}", "formula": formula, "role": "calculated_formula"}
                       for row, formula in enumerate(formulas, 1)]}

    result = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    source = result["functions"]["age_step"]
    namespace = {"runtime": runtime}
    exec(compile(source, "formula_families.py", "exec"), namespace)
    assert namespace["age_step"]({"raw.other": [3, 4, 5]}, ["raw.other"], 2) == expected_value
    assert expected_expression in source


def test_grouped_driver_runs_a_named_projection_loop_without_source_addresses() -> None:
    plan = {
        "variables": {
            "input.rate": {"variable_id": "input.rate", "shape": [3], "role": "raw",
                "axes": [{"name": "year"}], "dependencies": [], "equation_segments": [],
                "source_extents": [{"ref": "Main!A1:A3", "role": "literal_input",
                    "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year",
                        "dimension": 0, "source_coordinate": "row", "origin": 1,
                        "index_origin": 0}]}}]},
            "model.premium": {"variable_id": "model.premium", "shape": [3], "role": "derived",
                "axes": [{"name": "year"}], "dependencies": [{"variable_id": "input.rate", "lag": 0}],
                "equation_segments": [{"segment_id": "years", "source_family_ids": ["source.premium"],
                    "equation_family_ids": ["eq.premium"], "execution": "vectorized",
                    "index_axis": "year", "index_start": 0, "index_stop_exclusive": 3,
                    "recurrence_group": None, "snapshot_before_update": False}],
                "source_extents": [{"ref": "Main!B1:B3", "role": "formula_output",
                    "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year",
                        "dimension": 0, "source_coordinate": "row", "origin": 1,
                        "index_origin": 0}]}}]},
        },
        "equation_families": {"eq.premium": {"family_id": "eq.premium", "function_name": "premium_step"}},
        "source_mappings": [{"source_family_id": "source.premium", "equation_family_id": "eq.premium",
            "variable_id": "model.premium", "source_members": ["Main!B1", "Main!B2", "Main!B3"]}],
        "execution_plan": {"passes": [{"pass_id": "premium_projection",
            "source_family_ids": ["source.premium"]}]},
    }
    trace = {"cells": [
        {"address": "Main!B1", "formula": "=A1+10", "role": "calculated_formula"},
        {"address": "Main!B2", "formula": "=A2+10", "role": "calculated_formula"},
        {"address": "Main!B3", "formula": "=A3+10", "role": "calculated_formula"},
    ]}
    compilation = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    driver = render_grouped_driver(plan, compilation)
    assert "for _year in range(0, 3):" in driver
    assert "Main!" not in driver and "A1" not in driver and "B1" not in driver

    namespace = {"runtime": runtime}
    for source in compilation["functions"].values():
        exec(compile(source, "formula_families.py", "exec"), namespace)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    values = {"input.rate": [1, 2, 3], "model.premium": [None, None, None]}
    family_functions = {name: namespace[name] for name in compilation["functions"]}
    namespace["run_pricing"](values, family_functions)
    assert values["model.premium"] == [11, 12, 13]


def test_range_read_schedules_every_same_year_formula_member_before_aggregate() -> None:
    component_members = [f"Main!{column}{row}" for row in (1, 2) for column in ("B", "C", "D")]
    total_members = ["Main!E1", "Main!E2"]
    plan = {
        "variables": {
            "input.base": {"variable_id": "input.base", "shape": [2], "role": "raw",
                "axes": [{"name": "year"}], "dependencies": [], "equation_segments": [],
                "source_extents": [{"ref": "Main!A1:A2", "role": "raw_input",
                    "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year",
                        "dimension": 0, "source_coordinate": "row", "origin": 1,
                        "index_origin": 0}]}}]},
            "model.components": {"variable_id": "model.components", "shape": [2, 3], "role": "derived",
                "axes": [{"name": "year"}, {"name": "component"}],
                "dependencies": [{"variable_id": "input.base"}],
                "equation_segments": [{"segment_id": "component-values",
                    "source_family_ids": ["source.components"], "equation_family_ids": ["eq.components"],
                    "execution": "vectorized", "index_axis": "year", "index_start": 0,
                    "index_stop_exclusive": 2, "recurrence_group": None,
                    "snapshot_before_update": False}],
                "source_extents": [{"ref": "Main!B1:D2", "role": "formula_output",
                    "coordinate_mapping": {"kind": "row_column_ordinal", "axes": [
                        {"axis": "year", "dimension": 0, "source_coordinate": "row", "origin": 1,
                         "index_origin": 0},
                        {"axis": "component", "dimension": 1, "source_coordinate": "column", "origin": "B",
                         "index_origin": 0}]}}]},
            "model.total": {"variable_id": "model.total", "shape": [2], "role": "derived",
                "axes": [{"name": "year"}], "dependencies": [{"variable_id": "model.components"}],
                "equation_segments": [{"segment_id": "annual-total", "source_family_ids": ["source.total"],
                    "equation_family_ids": ["eq.total"], "execution": "vectorized", "index_axis": "year",
                    "index_start": 0, "index_stop_exclusive": 2, "recurrence_group": None,
                    "snapshot_before_update": False}],
                "source_extents": [{"ref": "Main!E1:E2", "role": "formula_output",
                    "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "year",
                        "dimension": 0, "source_coordinate": "row", "origin": 1,
                        "index_origin": 0}]}}]},
        },
        "equation_families": {
            "eq.components": {"family_id": "eq.components", "function_name": "component_step"},
            "eq.total": {"family_id": "eq.total", "function_name": "total_step"},
        },
        "source_mappings": [
            {"source_family_id": "source.total", "equation_family_id": "eq.total",
             "variable_id": "model.total", "source_members": total_members,
             "index_mapping": {"axes": [{"axis": "year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
            {"source_family_id": "source.components", "equation_family_id": "eq.components",
             "variable_id": "model.components", "source_members": component_members,
             "index_mapping": {"axes": [{"axis": "year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
        ],
        "execution_plan": {"passes": [{"pass_id": "annual", "order": 0,
            # Put the aggregate first so only traced range prerequisites can order it safely.
            "source_family_ids": ["source.total", "source.components"]}]},
    }
    trace = {"cells": [
        *[{"address": address, "formula": f"=A{address.split('!')[1][1:]}+1",
           "role": "calculated_formula"} for address in component_members],
        {"address": "Main!E1", "formula": "=SUM(B1:D1)", "role": "calculated_formula"},
        {"address": "Main!E2", "formula": "=SUM(B2:D2)", "role": "calculated_formula"},
    ], "edges": [
        {"consumer": "Main!E1", "prerequisite": "Main!B1:D1", "relationship": "range_read"},
        {"consumer": "Main!E2", "prerequisite": "Main!B2:D2", "relationship": "range_read"},
    ]}

    compilation = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    driver = render_grouped_driver(plan, compilation)
    assert set(compilation["execution_schedule"]["same_index_dependencies"]["source.total"]) == {
        "source.components"}
    assert driver.index("family_functions['component_step']") < driver.index("family_functions['total_step']")

    namespace = {"runtime": runtime}
    for source in compilation["functions"].values():
        exec(compile(source, "formula_families.py", "exec"), namespace)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    values = {"input.base": [1, 2], "model.components": [[None] * 3 for _ in range(2)],
              "model.total": [None, None]}
    family_functions = {name: namespace[name] for name in compilation["functions"]}
    namespace["run_pricing"](values, family_functions)
    assert values["model.components"] == [[2, 2, 2], [3, 3, 3]]
    assert values["model.total"] == [6, 9]


def test_lookup_table_schedules_only_key_and_selected_return_columns() -> None:
    def vector_variable(variable_id: str, family_id: str, address: str) -> dict[str, object]:
        return {
            "variable_id": variable_id, "shape": [2], "role": "derived",
            "axes": [{"name": "row"}], "dependencies": [],
            "equation_segments": [{"segment_id": family_id, "source_family_ids": [family_id],
                "equation_family_ids": [f"eq.{family_id}"], "execution": "vectorized",
                "index_axis": "row", "index_start": 0, "index_stop_exclusive": 2,
                "recurrence_group": None, "snapshot_before_update": False}],
            "source_extents": [{"ref": address, "role": "formula_output",
                "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "row",
                    "dimension": 0, "source_coordinate": "row", "origin": 1,
                    "index_origin": 0}]}}],
        }

    plan = {
        "variables": {
            "model.key": vector_variable("model.key", "source.key", "Main!A1:A2"),
            "model.return": vector_variable("model.return", "source.return", "Main!B1:B2"),
            "model.unselected": vector_variable(
                "model.unselected", "source.unselected", "Main!C1:C2"),
            "model.lookup": {
                "variable_id": "model.lookup", "shape": [], "role": "derived", "axes": [],
                "dependencies": [], "equation_segments": [{"segment_id": "lookup",
                    "source_family_ids": ["source.lookup"], "equation_family_ids": ["eq.lookup"],
                    "execution": "scalar", "index_start": 0, "index_stop_exclusive": 1,
                    "recurrence_group": None, "snapshot_before_update": False}],
                "source_extents": [{"ref": "Main!D1", "role": "formula_output",
                    "coordinate_mapping": {"kind": "scalar", "axes": []}}],
            },
        },
        "source_mappings": [
            {"source_family_id": "source.key", "variable_id": "model.key",
             "source_members": ["Main!A1", "Main!A2"],
             "index_mapping": {"axes": [{"axis": "row", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
            {"source_family_id": "source.return", "variable_id": "model.return",
             "source_members": ["Main!B1", "Main!B2"],
             "index_mapping": {"axes": [{"axis": "row", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
            {"source_family_id": "source.unselected", "variable_id": "model.unselected",
             "source_members": ["Main!C1", "Main!C2"],
             "index_mapping": {"axes": [{"axis": "row", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
            {"source_family_id": "source.lookup", "variable_id": "model.lookup",
             "source_members": ["Main!D1"], "index_mapping": {"axes": []}},
        ],
        "execution_plan": {"passes": [
            {"pass_id": "table_inputs", "order": 0,
             "source_family_ids": ["source.key", "source.return"]},
            {"pass_id": "lookup", "order": 1, "source_family_ids": ["source.lookup"]},
            {"pass_id": "unselected_table_column", "order": 2,
             "source_family_ids": ["source.unselected"]},
        ]},
    }
    trace = {
        "edges": [{"consumer": "Main!D1", "prerequisite": "Main!A1:C2",
                   "relationship": "lookup_table"}],
        "lookups": [{"formula_cell": "Main!D1", "function": "VLOOKUP",
                     "table_range": "Main!A1:C2", "lookup_column": 1,
                     "return_column": 2, "matched_row": 1}],
    }

    schedule = _build_execution_schedule(plan, trace, build_source_cell_index(plan))

    assert schedule["cross_pass_edge_count"] == 4
    assert schedule["family_pass"]["source.lookup"] == "lookup"
    with pytest.raises(CalculationBlocked, match="lacks a matching VLOOKUP footprint"):
        _build_execution_schedule(
            plan, {"edges": trace["edges"], "lookups": []}, build_source_cell_index(plan))


def test_grouped_recurrence_buffers_each_state_step_before_commit() -> None:
    def variable(variable_id: str, output_range: str, seed_ref: str, family_id: str,
                 dependency: str) -> dict[str, object]:
        return {
            "variable_id": variable_id, "shape": [3], "role": "derived",
            "axes": [{"name": "step"}], "dependencies": [{"variable_id": dependency, "lag": 1}],
            "initial_condition": {"seed_ref": seed_ref, "logical_index": 0},
            "equation_segments": [{"segment_id": family_id, "source_family_ids": [family_id],
                "equation_family_ids": [f"eq.{family_id}"], "execution": "ascending_recurrence",
                "index_axis": "step", "index_start": 1, "index_stop_exclusive": 3,
                "recurrence_group": "coupled", "snapshot_before_update": True}],
            "source_extents": [{"ref": output_range, "role": "formula_output",
                "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "step",
                    "dimension": 0, "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}],
        }

    plan = {
        "variables": {
            "model.left": variable("model.left", "Main!C2:C3", "Main!C1", "source.left", "model.right"),
            "model.right": variable("model.right", "Main!D2:D3", "Main!D1", "source.right", "model.left"),
        },
        "equation_families": {
            "eq.source.left": {"family_id": "eq.source.left", "function_name": "left_step"},
            "eq.source.right": {"family_id": "eq.source.right", "function_name": "right_step"},
        },
        "source_mappings": [
            {"source_family_id": "source.left", "equation_family_id": "eq.source.left",
             "variable_id": "model.left", "source_members": ["Main!C2", "Main!C3"]},
            {"source_family_id": "source.right", "equation_family_id": "eq.source.right",
             "variable_id": "model.right", "source_members": ["Main!D2", "Main!D3"]},
        ],
        "execution_plan": {"passes": [{"pass_id": "state_projection",
            "source_family_ids": ["source.left", "source.right"]}]},
    }
    trace = {"cells": [
        {"address": "Main!C2", "formula": "=D1+1", "role": "calculated_formula"},
        {"address": "Main!C3", "formula": "=D2+1", "role": "calculated_formula"},
        {"address": "Main!D2", "formula": "=C1+2", "role": "calculated_formula"},
        {"address": "Main!D3", "formula": "=C2+3", "role": "calculated_formula"},
    ]}
    compilation = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    driver = render_grouped_driver(plan, compilation)
    assert driver.count("for _step in range(1, 3):") == 1
    assert "_pending.append" in driver
    assert driver.index("_pending.append") < driver.index("for _variable_id, _indices, _value in _pending:")

    namespace = {"runtime": runtime}
    for source in compilation["functions"].values():
        exec(compile(source, "formula_families.py", "exec"), namespace)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    values = {"model.left": [1, None, None], "model.right": [2, None, None]}
    family_functions = {name: namespace[name] for name in compilation["functions"]}
    namespace["run_pricing"](values, family_functions)
    assert values["model.left"] == [1, 3, 4]
    assert values["model.right"] == [2, 3, 6]


def test_grouped_driver_interleaves_prior_state_and_current_year_dependencies() -> None:
    def annual(variable_id: str, column: str, family_id: str, *, recurrence: bool) -> dict[str, object]:
        start = 1 if recurrence else 0
        return {
            "variable_id": variable_id, "role": "derived", "shape": [4 if recurrence else 3],
            "axes": [{"name": "policy_year"}], "dependencies": [],
            "initial_condition": ({"seed_ref": f"Main!{column}1", "logical_index": 0}
                                  if recurrence else None),
            "equation_segments": [{"segment_id": family_id, "source_family_ids": [family_id],
                "equation_family_ids": [f"eq.{family_id}"], "execution":
                    "ascending_recurrence" if recurrence else "vectorized",
                "index_axis": "policy_year", "index_start": start,
                "index_stop_exclusive": 4 if recurrence else 3,
                "recurrence_group": "premium_states" if recurrence else None,
                "snapshot_before_update": recurrence}],
            "source_extents": [{"ref": f"Main!{column}{2 if recurrence else 1}:"
                                          f"{column}{4 if recurrence else 3}",
                "role": "formula_output", "coordinate_mapping": {"kind": "row_ordinal",
                    "axes": [{"axis": "policy_year", "dimension": 0,
                        "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}],
        }

    h_id, k_id = "premium.state_h", "premium.state_k"
    h_family, k_family = "source.h", "source.k"
    h_members = ["Main!A2", "Main!A3", "Main!A4"]
    k_members = ["Main!B1", "Main!B2", "Main!B3"]
    plan = {
        "variables": {
            h_id: annual(h_id, "A", h_family, recurrence=True),
            k_id: annual(k_id, "B", k_family, recurrence=False),
        },
        "equation_families": {
            "eq.source.h": {"family_id": "eq.source.h", "function_name": "advance_h"},
            "eq.source.k": {"family_id": "eq.source.k", "function_name": "calculate_k"},
        },
        "source_mappings": [
            {"source_family_id": h_family, "equation_family_id": "eq.source.h",
             "variable_id": h_id, "source_members": h_members,
             "index_mapping": {"axes": [{"axis": "policy_year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
            {"source_family_id": k_family, "equation_family_id": "eq.source.k",
             "variable_id": k_id, "source_members": k_members,
             "index_mapping": {"axes": [{"axis": "policy_year", "dimension": 0,
                 "source_coordinate": "row", "origin": 1, "index_origin": 0}]}},
        ],
        "execution_plan": {"passes": [{"pass_id": "premium_projection", "order": 0,
            "source_family_ids": [h_family, k_family]}]},
    }
    trace = {"cells": [
        *[{"address": address, "formula": f"=B{row - 1}+1", "role": "calculated_formula"}
          for row, address in zip(range(2, 5), h_members)],
        *[{"address": address, "formula": f"=A{row}+10", "role": "calculated_formula"}
          for row, address in zip(range(1, 4), k_members)],
    ], "edges": [
        *[{"consumer": f"Main!A{row}", "prerequisite": f"Main!B{row - 1}",
           "relationship": "cell_reference"} for row in range(2, 5)],
        *[{"consumer": f"Main!B{row}", "prerequisite": f"Main!A{row}",
           "relationship": "cell_reference"} for row in range(2, 4)],
    ]}

    compilation = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    driver = render_grouped_driver(plan, compilation)
    assert "for _policy_year in range(0, 4):" in driver
    assert driver.index("_pending.append") < driver.index("for _variable_id, _indices, _value in _pending:")

    namespace = {"runtime": runtime}
    for source in compilation["functions"].values():
        exec(compile(source, "formula_families.py", "exec"), namespace)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    values = {h_id: [1, None, None, None], k_id: [None, None, None]}
    functions = {name: namespace[name] for name in compilation["functions"]}
    namespace["run_pricing"](values, functions)

    assert values[h_id] == [1, 12, 23, 34]
    assert values[k_id] == [11, 22, 33]
    assert compilation["emitted_schedule"]["premium_projection"][0]["family_order"] == [h_family, k_family]


def test_106_year_projection_uses_one_named_axis_loop_and_family_function() -> None:
    years = 106
    members = [f"Main!B{row}" for row in range(1, years + 1)]
    trace = {"cells": [{"address": address, "formula": f"=A{index + 1}+10",
                        "role": "calculated_formula"}
                       for index, address in enumerate(members)]}
    plan = {
        "variables": {
            "input.premium_rate": {"variable_id": "input.premium_rate", "role": "raw",
                "kind": "projection_series", "shape": [years], "axes": [{"name": "policy_year"}],
                "dependencies": [], "equation_segments": [],
                "source_extents": [{"ref": f"Main!A1:A{years}", "role": "raw_input",
                    "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "policy_year",
                        "dimension": 0, "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
            "premium.gross_projection": {"variable_id": "premium.gross_projection", "role": "derived",
                "kind": "projection_series", "shape": [years], "axes": [{"name": "policy_year"}],
                "dependencies": [{"variable_id": "input.premium_rate", "lag": 0}],
                "equation_segments": [{"segment_id": "annual_projection", "source_family_ids": ["source.premium"],
                    "equation_family_ids": ["eq.premium"], "execution": "vectorized",
                    "index_axis": "policy_year", "index_start": 0, "index_stop_exclusive": years,
                    "recurrence_group": None, "snapshot_before_update": False}],
                "source_extents": [{"ref": f"Main!B1:B{years}", "role": "formula_output",
                    "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "policy_year",
                        "dimension": 0, "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}]},
        },
        "equation_families": {"eq.premium": {"family_id": "eq.premium", "function_name": "premium_projection_step"}},
        "source_mappings": [{"source_family_id": "source.premium", "equation_family_id": "eq.premium",
            "variable_id": "premium.gross_projection", "source_members": members}],
        "execution_plan": {"passes": [{"pass_id": "annual_projection", "source_family_ids": ["source.premium"]}]},
    }

    compilation = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    driver = render_grouped_driver(plan, compilation)

    assert compilation["compiled_member_count"] == years
    assert compilation["function_count"] == 1
    assert "for _policy_year in range(0, 106):" in driver
    assert driver.count("family_functions[") == 1
    assert "Main!" not in driver and "B1" not in driver

    namespace = {"runtime": runtime}
    for source in compilation["functions"].values():
        exec(compile(source, "formula_families.py", "exec"), namespace)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    values = {"input.premium_rate": list(range(years)), "premium.gross_projection": [None] * years}
    family_functions = {name: namespace[name] for name in compilation["functions"]}
    namespace["run_pricing"](values, family_functions)
    assert values["premium.gross_projection"] == [rate + 10 for rate in range(years)]


def test_104_step_markov_projection_uses_snapshot_recurrence_loop() -> None:
    steps = 104
    stop = steps + 1

    def state(variable_id: str, col: str, family_id: str, equation_id: str,
              dependency_id: str, seed: int, increment: int) -> dict[str, object]:
        return {
            "variable_id": variable_id, "role": "derived", "kind": "matrix", "shape": [stop],
            "axes": [{"name": "markov_step"}],
            "dependencies": [{"variable_id": dependency_id, "lag": 1}],
            "initial_condition": {"seed_ref": f"Main!{col}1", "logical_index": 0,
                                  "source_note": f"literal test seed {seed}"},
            "equation_segments": [{"segment_id": family_id, "source_family_ids": [family_id],
                "equation_family_ids": [equation_id], "execution": "ascending_recurrence",
                "index_axis": "markov_step", "index_start": 1, "index_stop_exclusive": stop,
                "recurrence_group": "two_state", "snapshot_before_update": True}],
            "source_extents": [{"ref": f"Main!{col}2:{col}{stop}", "role": "formula_output",
                "coordinate_mapping": {"kind": "row_ordinal", "axes": [{"axis": "markov_step",
                    "dimension": 0, "source_coordinate": "row", "origin": 1, "index_origin": 0}]}}],
        }

    left_id, right_id = "markov.state_left", "markov.state_right"
    left_family, right_family = "source.left", "source.right"
    left_members = [f"Main!C{row}" for row in range(2, stop + 1)]
    right_members = [f"Main!D{row}" for row in range(2, stop + 1)]
    trace = {"cells": [
        *[{"address": address, "formula": f"=D{row - 1}+1", "role": "calculated_formula"}
          for row, address in zip(range(2, stop + 1), left_members)],
        *[{"address": address, "formula": f"=C{row - 1}+2", "role": "calculated_formula"}
          for row, address in zip(range(2, stop + 1), right_members)],
    ]}
    plan = {
        "variables": {
            left_id: state(left_id, "C", left_family, "eq.left", right_id, 1, 1),
            right_id: state(right_id, "D", right_family, "eq.right", left_id, 2, 2),
        },
        "equation_families": {
            "eq.left": {"family_id": "eq.left", "function_name": "left_update"},
            "eq.right": {"family_id": "eq.right", "function_name": "right_update"},
        },
        "source_mappings": [
            {"source_family_id": left_family, "equation_family_id": "eq.left",
             "variable_id": left_id, "source_members": left_members},
            {"source_family_id": right_family, "equation_family_id": "eq.right",
             "variable_id": right_id, "source_members": right_members},
        ],
        "execution_plan": {"passes": [{"pass_id": "markov_recurrence",
            "source_family_ids": [left_family, right_family]}]},
    }

    compilation = compile_mapped_families(plan, trace, lambda name, sheet, owner: "None")
    driver = render_grouped_driver(plan, compilation)

    assert compilation["compiled_member_count"] == 2 * steps
    assert compilation["function_count"] == 2
    assert driver.count("for _markov_step in range(1, 105):") == 1
    assert "_pending.append" in driver
    assert driver.index("_pending.append") < driver.index("for _variable_id, _indices, _value in _pending:")
    assert "Main!" not in driver

    namespace = {"runtime": runtime}
    for source in compilation["functions"].values():
        exec(compile(source, "formula_families.py", "exec"), namespace)
    exec(compile(driver, "pricing.py", "exec"), namespace)
    values = {left_id: [1] + [None] * steps, right_id: [2] + [None] * steps}
    family_functions = {name: namespace[name] for name in compilation["functions"]}
    namespace["run_pricing"](values, family_functions)

    expected_left, expected_right = [1], [2]
    for _ in range(steps):
        next_left, next_right = expected_right[-1] + 1, expected_left[-1] + 2
        expected_left.append(next_left)
        expected_right.append(next_right)
    assert values[left_id] == expected_left
    assert values[right_id] == expected_right


def test_external_formula_extent_is_indexed_as_an_input_not_an_emitted_family() -> None:
    plan = {"variables": {
        "external.ci_rate": {
            "variable_id": "external.ci_rate", "role": "external", "shape": [3],
            "axes": [{"name": "attained_age"}], "source_extents": [{
                "ref": "Rates!W5:W7", "role": "formula_output",
                "coordinate_mapping": {"kind": "row_ordinal", "axes": [{
                    "axis": "attained_age", "dimension": 0, "source_coordinate": "row",
                    "origin": 5, "index_origin": 0}]},
            }],
        }
    }, "source_mappings": [], "array_mappings": []}

    index = build_source_cell_index(plan)

    assert index[("rates", 5, 23)] == [
        SourceCell("external.ci_rate", (0,), ("attained_age",), "external_boundary_input")]
    assert index[("rates", 7, 23)][0].indices == (2,)
