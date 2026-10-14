from __future__ import annotations

import json
from pathlib import Path

import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps import conversion_workflow as workflow
from excel_to_act.steps.step3.input_boundary import confirmed_boundary_reference
from excel_to_act.steps.step3.tools.interface import (
    _build_contract, export_tool_interface, validate_tool_interface_declarations,
)


def _approved_workflow(tmp_path: Path, *, declarations: list[dict] | None = None,
                       support_domain: dict | None = None, target: dict | None = None,
                       wrong_source: bool = False) -> Path:
    root, _ = workflow.create_workflow(tmp_path / "workflow", {
        "source_id": "pricing", "run_id": "run-1", "workbook_sha256": "a" * 64})
    _, manifest = workflow.load_workflow(root)
    manifest["input_boundary_required"] = True
    workflow._write_manifest(root, manifest)
    stage1 = workflow.append_stage_artifact(root, 1, "stage1.json", {"status": "pass"}, "Stage 1")
    for reviewer in ("agent", "human"):
        workflow.record_decision(root, 1, reviewer, "approve", "Fixture approval")
    stage2 = workflow.append_stage_artifact(root, 2, "stage2.json", {"status": "pass"}, "Stage 2",
                                            input_stages={1: stage1["artifact"]})
    for reviewer in ("agent", "human"):
        workflow.record_decision(root, 2, reviewer, "approve", "Fixture approval")
    variables = []
    for name, kind, shape, role, logical_role in (
        ("IssueAge", "scalar", [], "source_raw", "business_input"),
        ("RateCurve", "vector", [2], "formula_derived_external", "series"),
        ("SwitchTable", "table", [2, 2], "source_raw", "configuration_table"),
    ):
        variables.append({
            "variable_id": name.lower(), "logical_name": name, "kind": kind, "shape": shape,
            "axes": [{"name": f"axis-{index}", "role": "category"}
                     for index in range(len(shape))], "role": role, "logical_role": logical_role,
            "group_id": "pricing", "consumer_targets": ["GP"],
            "source_extents": [{"sheet": "Main", "range": "A1", "source_role": "raw_source"}],
            "value_source_policy": {"formula_cache_allowed": False,
                                    "available_now": role != "formula_derived_external"},
        })
    source = {"source_id": "pricing", "run_id": "run-1", "workbook_sha256": "a" * 64,
              "analysis_binding_sha256": "b" * 64}
    boundary = {"schema_version": "step3.input_boundary.v1", "phase": "input_boundary",
                "status": "ready_for_review", "source": source, "boundary_sha256": "c" * 64,
                "variables": variables, "target_selection": {"targets": [{"target_id": "GP"}]}}
    workflow.append_stage_artifact(root, 3, "input_boundary.json", boundary, "Boundary",
                                   input_stages={2: stage2["artifact"]})
    for reviewer in ("agent", "human"):
        workflow.record_decision(root, 3, reviewer, "approve", "Fixture boundary approval")
    _, manifest = workflow.load_workflow(root)
    boundary_reference, boundary_json, boundary_md = confirmed_boundary_reference(root, manifest)
    design_source = {**source, "source_id": "other"} if wrong_source else source
    design = {"schema_version": "step3.analysis_design.v1", "status": "ready_for_review",
              "source": design_source, "binding_sha256": "b" * 64,
              "input_boundary_reference": boundary_reference,
              "design": {"targets": [target or {"target_id": "GP", "selector": "GP", "result_order": 0,
                                      "shape": [], "units": "Currency and basis under review"}],
                         "scenarios": [{"scenario_id": "saved", "primary_inputs": {
                             "IssueAge": "saved-value-must-not-become-default"}}],
                         **({"tool_interface": {"objects": declarations,
                                                **({"support_domain": support_domain}
                                                   if support_domain is not None else {})}}
                            if declarations is not None else {})}}
    workflow.append_stage_artifact(root, 3, "analysis_design.json", design, "Design",
                                   input_files={"boundary_json": boundary_json,
                                                "boundary_markdown": boundary_md},
                                   input_stages={2: stage2["artifact"]})
    for reviewer in ("agent", "human"):
        workflow.record_decision(root, 3, reviewer, "approve", "Fixture design approval")
    return root


def _declaration(name: str) -> dict:
    return {"variable_id": name.lower(), "mutability": "mutable", "value_type": "number",
            "required": True, "default_policy": {"status": "none"},
            "allowed_range": {"status": "declared", "minimum": 0, "maximum": 100},
            "units": "rate", "unit_basis": "per period"}


def test_export_is_deterministic_and_preserves_logical_objects(tmp_path: Path) -> None:
    root = _approved_workflow(tmp_path, declarations=[_declaration(name) for name in
                                                      ("IssueAge", "RateCurve", "SwitchTable")])
    stage3_before = workflow.load_workflow(root)[1]["stages"]["3"]
    first = export_tool_interface(root)
    assert first["status"] == "pass" and first["message"] == "interface generated"
    definition_path = root / first["artifact"]["json"]
    first_bytes = definition_path.read_bytes()
    definition = json.loads(first_bytes)
    assert definition["schema_id"] == "step3.tool_interface.v1"
    assert set(definition["parameters"]["properties"]) == {"IssueAge", "RateCurve", "SwitchTable"}
    assert definition["parameters"]["required"] == ["IssueAge", "RateCurve", "SwitchTable"]
    assert definition["parameters"]["properties"]["RateCurve"]["minItems"] == 2
    assert definition["parameters"]["properties"]["SwitchTable"]["items"]["minItems"] == 2
    assert definition["parameters"]["properties"]["IssueAge"]["minimum"] == 0
    assert definition["parameters"]["additionalProperties"] is False
    assert definition["source_provenance"]["inputs"]["ratecurve"]["role"] == "formula_derived_external"
    assert definition["bindings"]["design"]["review_receipts"]["agent"]["decision"] == "approve"
    assert definition["bindings"]["input_boundary"]["review_receipts"]["human"]["decision"] == "approve"
    assert definition["contract_readiness"]["status"] == "blocked"
    assert any("external value is unavailable" in item for item in
               definition["contract_readiness"]["blockers"])
    assert "output GP: units are unknown" in definition["contract_readiness"]["unknown_facts"]
    assert definition["runtime_readiness"]["status"] == "blocked"
    assert "saved-value-must-not-become-default" not in first_bytes.decode()
    assert "Main!A1" not in json.dumps(definition["parameters"])
    assert export_tool_interface(root) == first
    assert definition_path.read_bytes() == first_bytes
    assert len(workflow.load_workflow(root)[1]["tool_interfaces"]) == 1
    assert workflow.load_workflow(root)[1]["stages"]["3"] == stage3_before
    assert workflow.workflow_status(root)["tool_interface"]["status"] == "interface_generated"


def test_missing_business_declarations_remain_unknown(tmp_path: Path) -> None:
    root = _approved_workflow(tmp_path)
    result = export_tool_interface(root)
    definition = workflow.read_json(root / result["artifact"]["json"])
    assert result["status"] == "pass"
    assert definition["parameters"]["properties"] == {}
    assert {item["mutability"] for item in definition["objects"]} == {"unknown"}
    assert len([item for item in definition["contract_readiness"]["blockers"]
                if item.endswith("mutability is unknown")]) == 3


def test_descriptive_unknowns_do_not_block_typed_sample_contract(tmp_path: Path) -> None:
    fixed = {"variable_id": "switchtable", "mutability": "immutable", "value_type": "unknown",
             "required": False, "default_policy": {"status": "none"},
             "allowed_range": {"status": "unknown"}}
    external = {**_declaration("RateCurve"), "allowed_range": {"status": "unbounded"}}
    domain = {"status": "fixed_trace", "active_trace_sha256": "a" * 64,
              "external_policy": {"ratecurve": {"mode": "same_request",
                                                "source_sha256": "a" * 64, "version": "v1"}},
              "tolerances": {"absolute": 1e-12, "relative": 1e-12}}
    target = {"target_id": "Result", "selector": "Result", "result_order": 0,
              "shape": [], "axes": [], "value_type": "number", "units": None,
              "unit_basis": None, "allowed_range": {"status": "unknown"}}
    root = _approved_workflow(tmp_path, declarations=[_declaration("IssueAge"), external, fixed],
                              support_domain=domain, target=target)
    result = export_tool_interface(root)
    assert result["status"] == "pass"
    definition = workflow.read_json(root / result["artifact"]["json"])
    readiness = definition["contract_readiness"]
    assert readiness["status"] == "defined" and readiness["blockers"] == []
    assert "SwitchTable: value type is unknown" in readiness["unknown_facts"]
    assert "SwitchTable: allowed range is unknown" in readiness["unknown_facts"]
    assert "output Result: units are unknown" in readiness["unknown_facts"]
    assert "output Result: unit basis is unknown" in readiness["unknown_facts"]
    assert "output Result: allowed range is unknown" in readiness["unknown_facts"]
    assert definition["outputs"]["properties"]["Result"]["type"] == "number"
    assert definition["outputs"]["properties"]["Result"]["x-allowed-range"] == {"status": "unknown"}
    assert "output Result: allowed range is unknown" in (root / result["artifact"]["md"]).read_text()


@pytest.mark.parametrize(("change", "expected"), [
    ({"value_type": "unknown", "allowed_range": {"status": "unbounded"}}, "value type is unknown"),
    ({"allowed_range": {"status": "unknown"}}, "allowed range is unknown"),
    ({"default_policy": {"status": "unknown"}}, "default policy is unknown"),
    ({"required": "omit"}, "required policy is unknown"),
])
def test_mutable_scalar_unknown_execution_policy_blocks(change: dict, expected: str) -> None:
    declaration = {**_declaration("Count"), **change}
    if declaration.get("required") == "omit":
        declaration.pop("required")
    boundary = {"variables": [{"variable_id": "count", "logical_name": "Count",
                               "kind": "scalar", "shape": [], "axes": []}]}
    _, _, _, blockers = _build_contract(
        {"design": {"tool_interface": {"objects": [declaration]}}}, boundary)
    assert f"Count: {expected}" in blockers


def test_unknown_table_container_needs_source_mapped_typed_cells() -> None:
    declaration = {"variable_id": "grid", "mutability": "mutable", "value_type": "unknown",
                   "required": True, "default_policy": {"status": "none"},
                   "allowed_range": {"status": "unknown"}, "mutable_cells": [
                       {"indices": [1, 1], "value_type": "string", "allow_blank": False,
                        "allowed_range": {"status": "enum", "values": ["A"]}}]}
    boundary = {"variables": [{"variable_id": "grid", "logical_name": "Grid", "kind": "table",
                               "shape": [2, 2], "axes": [{"name": "row"}, {"name": "column"}],
                               "coordinate_map": [{"logical_index": [1, 1],
                                                   "source_cell": "Sheet!B2"}]}]}
    design = {"design": {"tool_interface": {"objects": [declaration]}}}
    _, parameters, _, blockers = _build_contract(design, boundary)
    assert blockers == []
    assert parameters["properties"]["Grid"]["x-mutable-cells"] == declaration["mutable_cells"]
    boundary["variables"][0]["coordinate_map"][0]["logical_index"] = [0, 0]
    with pytest.raises(ValueError, match="source coordinate"):
        _build_contract(design, boundary)
    boundary["variables"][0]["coordinate_map"][0]["logical_index"] = [1, 1]
    declaration["mutable_cells"][0].pop("allow_blank")
    with pytest.raises(ValueError, match="blank policy"):
        _build_contract(design, boundary)


def test_external_unknown_leaf_type_or_missing_same_request_policy_blocks() -> None:
    declaration = {**_declaration("Cut"), "value_type": "unknown",
                   "allowed_range": {"status": "unbounded"}}
    boundary = {"variables": [{"variable_id": "cut", "logical_name": "Cut", "kind": "vector",
                               "shape": [2], "axes": [{"name": "age", "keys": [0, 1]}],
                               "value_source_policy": {"available_now": False}}]}
    design = {"design": {"tool_interface": {"objects": [declaration]}}}
    _, _, _, blockers = _build_contract(design, boundary)
    assert "Cut: value type is unknown" in blockers
    assert "Cut: approved current external value is unavailable" in blockers
    _, _, _, blockers = _build_contract(design, boundary, {"external_policy": {
        "cut": {"mode": "same_request", "source_sha256": "a" * 64, "version": "v1"}}})
    assert "Cut: value type is unknown" in blockers
    assert not any("external value is unavailable" in item for item in blockers)


@pytest.mark.parametrize("mutation", ["no_human", "typesafe_only", "cross_revision",
                                      "json_tamper", "md_tamper", "wrong_source",
                                      "design_json_tamper", "design_md_tamper",
                                      "stale_upstream", "rejected_upstream"])
def test_export_rejects_noncurrent_approval_or_source(tmp_path: Path, mutation: str) -> None:
    root = _approved_workflow(tmp_path, wrong_source=mutation == "wrong_source")
    _, manifest = workflow.load_workflow(root)
    boundary_revision = manifest["stages"]["3"]["revisions"][0]
    if mutation == "no_human":
        boundary_revision["decisions"] = [receipt for receipt in boundary_revision["decisions"]
                                           if receipt["reviewer"] != "human"]
        workflow._write_manifest(root, manifest)
    elif mutation == "typesafe_only":
        boundary_revision["decisions"][1]["reviewer"] = "typesafe"
        workflow._write_manifest(root, manifest)
    elif mutation == "cross_revision":
        boundary_revision["decisions"][1]["revision"] += 1
        workflow._write_manifest(root, manifest)
    elif mutation in {"json_tamper", "md_tamper"}:
        path = root / boundary_revision["artifact"]["json" if mutation == "json_tamper" else "md"]
        path.write_bytes(path.read_bytes() + b"changed")
    elif mutation in {"design_json_tamper", "design_md_tamper"}:
        kind = "json" if mutation == "design_json_tamper" else "md"
        path = root / manifest["stages"]["3"]["revisions"][1]["artifact"][kind]
        path.write_bytes(path.read_bytes() + b"changed")
    elif mutation == "stale_upstream":
        path = root / manifest["stages"]["2"]["revisions"][0]["artifact"]["md"]
        path.write_bytes(path.read_bytes() + b"changed")
    elif mutation == "rejected_upstream":
        workflow.record_decision(root, 2, "agent", "reject", "Revise upstream", return_to=2)
    result = export_tool_interface(root)
    assert result["status"] == "blocked"
    assert not (root / "tool_interfaces").exists()


def test_unknown_fields_and_literal_defaults_are_rejected(tmp_path: Path) -> None:
    unknown = {**_declaration("IssueAge"), "surprise": "no"}
    root = _approved_workflow(tmp_path, declarations=[unknown])
    result = export_tool_interface(root)
    assert result["status"] == "blocked"
    assert "unsupported fields" in result["diagnostics"][0]["message"]
    with pytest.raises(ValueError, match="literal defaults need later approved value capture"):
        validate_tool_interface_declarations({"objects": [{**_declaration("IssueAge"),
            "default_policy": {"status": "literal", "value": 65}}]})


def test_variable_domain_declarations_need_exact_trace_producer_and_tolerances() -> None:
    declaration = {**_declaration("IssueAge"), "required": False,
                   "default_policy": {"status": "saved"}, "allow_blank": False,
                   "allowed_range": {"status": "enum", "values": [60, 65]}}
    domain = {"status": "fixed_trace", "active_trace_sha256": "a" * 64,
              "external_policy": {"rate": {"mode": "same_request",
                                            "source_sha256": "b" * 64, "version": "v1"}},
              "tolerances": {"absolute": 1e-12, "relative": 1e-12}}
    validate_tool_interface_declarations({"objects": [declaration], "support_domain": domain})
    with pytest.raises(ValueError, match="absolute and relative tolerances"):
        validate_tool_interface_declarations({"objects": [declaration],
                                              "support_domain": {**domain, "tolerances": {"absolute": 0}}})
    with pytest.raises(ValueError, match="distinct finite scalar values"):
        validate_tool_interface_declarations({"objects": [{**declaration,
            "allowed_range": {"status": "enum", "values": [65, 65]}}]})


def test_exact_case_domain_requires_complete_hashed_evidence() -> None:
    reference = {"path": "case.json", "sha256": "b" * 64}
    domain = {"status": "exact_cases", "active_trace_sha256": "a" * 64,
              "cases": [{name: reference for name in
                         ("trace", "raw_request", "public_request", "producer", "plan", "profile")}],
              "external_policy": {}, "tolerances": {"absolute": 1e-12, "relative": 1e-12}}
    validate_tool_interface_declarations({"objects": [_declaration("Count")],
                                          "support_domain": domain})
    with pytest.raises(ValueError, match="distinct plan and profile"):
        validate_tool_interface_declarations({"objects": [_declaration("Count")],
            "support_domain": {**domain, "cases": [{"trace": reference}]}})
    with pytest.raises(ValueError, match="malformed"):
        validate_tool_interface_declarations({"objects": [_declaration("Count")],
            "support_domain": {**domain, "cases": [{**domain["cases"][0],
                "producer": {"path": "case.json", "sha256": "bad"}}]}})


def test_exact_cases_can_defer_trace_binding_until_current_discovery() -> None:
    reference = {"path": "case.json", "sha256": "b" * 64}
    domain = {"status": "exact_cases", "trace_binding": "post_discovery",
              "cases": [{name: reference for name in
                         ("trace", "raw_request", "public_request", "producer", "plan", "profile")}],
              "external_policy": {}, "tolerances": {"absolute": 0, "relative": 0}}
    validate_tool_interface_declarations({"objects": [_declaration("Count")],
                                          "support_domain": domain})
    with pytest.raises(ValueError, match="post-discovery|active-trace"):
        validate_tool_interface_declarations({"objects": [_declaration("Count")],
            "support_domain": {**domain, "active_trace_sha256": "a" * 64}})
    with pytest.raises(ValueError, match="post-discovery|active-trace"):
        validate_tool_interface_declarations({"objects": [_declaration("Count")],
            "support_domain": {**domain, "status": "fixed_trace"}})


def test_dependent_enum_and_per_cell_policies_are_explicit() -> None:
    count = {**_declaration("Count"), "value_type": "integer",
             "allowed_range": {"status": "declared", "minimum": 2, "maximum": 7}}
    period = {**_declaration("Period"), "value_type": "integer",
              "allowed_range": {"status": "dependent_enum", "values": [1, 2],
                                "variable_id": "count"}}
    table = {**_declaration("Grid"), "mutable_cells": [
        {"indices": [0, 1], "value_type": "string", "allow_blank": False,
         "allowed_range": {"status": "enum", "values": ["A", "B"]}}]}
    validate_tool_interface_declarations({"objects": [count, period, table]})
    with pytest.raises(ValueError, match="nonnegative indices"):
        validate_tool_interface_declarations({"objects": [{**table, "mutable_cells": [
            {**table["mutable_cells"][0], "indices": [True, 1]}]}]})
    with pytest.raises(ValueError, match="dependent enum needs a variable_id"):
        validate_tool_interface_declarations({"objects": [{**period,
            "allowed_range": {"status": "dependent_enum", "values": [1]}}]})
    self_referencing = {**period, "variable_id": "period",
                        "allowed_range": {"status": "dependent_enum", "values": [1],
                                          "variable_id": "period"}}
    boundary = {"variables": [{"variable_id": "period", "logical_name": "Period",
                               "kind": "scalar", "shape": [], "axes": []}]}
    with pytest.raises(ValueError, match="dependent enum needs a mutable scalar"):
        _build_contract({"design": {"tool_interface": {"objects": [self_referencing]}}}, boundary)


def test_design_axis_keys_extend_boundary_identity_and_export_to_parameters() -> None:
    boundary = {"variables": [{"variable_id": "grid", "logical_name": "Grid", "kind": "table",
                              "shape": [2, 2], "axes": [{"name": "row", "role": "record"},
                                                       {"name": "column", "role": "category",
                                                        "axis_id": "source-column"}]}]}
    declaration = {**_declaration("Grid"), "axes": [
        {"name": "row", "role": "record", "keys": [0, False]},
        {"name": "column", "role": "category", "axis_id": "source-column",
         "keys": ["A", "B"]}]}
    objects, parameters, *_ = _build_contract(
        {"design": {"tool_interface": {"objects": [declaration]}}}, boundary)
    assert objects[0]["axes"] == declaration["axes"]
    assert parameters["properties"]["Grid"]["x-axes"] == declaration["axes"]
    numeric_types = {**declaration, "axes": [
        {**declaration["axes"][0], "keys": [0, 0.0]}, declaration["axes"][1]]}
    typed_objects, *_ = _build_contract(
        {"design": {"tool_interface": {"objects": [numeric_types]}}}, boundary)
    assert typed_objects[0]["axes"][0]["keys"] == [0, 0.0]
    omitted, legacy_parameters, *_ = _build_contract(
        {"design": {"tool_interface": {"objects": [{key: value for key, value in declaration.items()
                                                    if key != "axes"}]}}}, boundary)
    assert omitted[0]["axes"] == boundary["variables"][0]["axes"]
    assert legacy_parameters["properties"]["Grid"]["x-axes"] == boundary["variables"][0]["axes"]


@pytest.mark.parametrize(("first_keys", "second_keys"), [
    ([1, 2], [2, 1]),
    ([1, 2], None),
    ([0, False], [False, 0]),
    ([0, 1], [0.0, 1]),
])
def test_shared_axis_id_rejects_conflicting_keys(first_keys: list, second_keys: list | None) -> None:
    boundary = {"variables": [
        {"variable_id": name.lower(), "logical_name": name, "kind": "vector",
         "shape": [2], "axes": [{"name": "position", "role": "key", "axis_id": "position"}]}
        for name in ("Rate", "Factor")]}
    first = {**_declaration("Rate"), "axes": [
        {**boundary["variables"][0]["axes"][0], "keys": first_keys}]}
    second_axis = {**boundary["variables"][1]["axes"][0]}
    if second_keys is not None:
        second_axis["keys"] = second_keys
    second = {**_declaration("Factor"), "axes": [second_axis]}
    with pytest.raises(ValueError, match="shared axis"):
        _build_contract({"design": {"tool_interface": {"objects": [first, second]}}}, boundary)


def test_shared_axis_id_exports_compatible_typed_keys() -> None:
    boundary = {"variables": [
        {"variable_id": name.lower(), "logical_name": name, "kind": "vector",
         "shape": [2], "axes": [{"name": "position", "role": "key", "axis_id": "position"}]}
        for name in ("Rate", "Factor")]}
    declarations = [{**_declaration(name), "axes": [
        {**boundary["variables"][0]["axes"][0], "keys": [0, False]}]}
        for name in ("Rate", "Factor")]
    objects, parameters, *_ = _build_contract(
        {"design": {"tool_interface": {"objects": declarations}}}, boundary)
    assert [item["axes"][0]["keys"] for item in objects] == [[0, False], [0, False]]
    assert parameters["properties"]["Rate"]["x-axes"][0]["keys"] == [0, False]
    assert parameters["properties"]["Factor"]["x-axes"][0]["keys"] == [0, False]
    unkeyed, *_ = _build_contract({"design": {"tool_interface": {"objects": [
        _declaration("Rate"), _declaration("Factor")]}}}, boundary)
    assert all("keys" not in item["axes"][0] for item in unkeyed)
    other_boundary = {"variables": [boundary["variables"][0],
                                    {**boundary["variables"][1], "axes": [
                                        {**boundary["variables"][1]["axes"][0],
                                         "axis_id": "other-position"}]}]}
    other_declarations = [declarations[0], {**declarations[1], "axes": [
        {**other_boundary["variables"][1]["axes"][0], "keys": [False, 0]}]}]
    separate, *_ = _build_contract(
        {"design": {"tool_interface": {"objects": other_declarations}}}, other_boundary)
    by_id = {item["variable_id"]: item["axes"][0] for item in separate}
    assert by_id["rate"]["keys"] == [0, False]
    assert [type(key) for key in by_id["factor"]["keys"]] == [bool, int]


def test_shared_axis_id_rejects_inconsistent_lengths() -> None:
    boundary = {"variables": [
        {"variable_id": name.lower(), "logical_name": name, "kind": "vector",
         "shape": [size], "axes": [{"name": "period", "role": "key", "axis_id": "period"}]}
        for name, size in (("Rate", 2), ("Factor", 3))]}
    design = {"design": {"tool_interface": {"objects": [
        _declaration("Rate"), _declaration("Factor")]}}}
    with pytest.raises(ValueError, match="shared axis period.*length"):
        _build_contract(design, boundary)
    boundary["axis_metadata"] = [{"axis_id": "period", "shape": [2]}]
    with pytest.raises(ValueError, match="axis metadata period.*length"):
        _build_contract(design, boundary)
    boundary["variables"][1]["axes"][0]["axis_id"] = "other-period"
    boundary["axis_metadata"].append({"axis_id": "other-period", "shape": [3]})
    objects, parameters, *_ = _build_contract(design, boundary)
    assert [item["shape"] for item in objects] == [[3], [2]]
    assert parameters["properties"]["Factor"]["x-axes"][0]["axis_id"] == "other-period"


@pytest.mark.parametrize(("axes", "message"), [
    ([{"name": "row", "role": "record", "keys": [1]},
      {"name": "column", "role": "category", "axis_id": "source-column", "keys": ["A", "B"]}],
     "key count"),
    ([{"name": "row", "role": "record", "keys": [1, 1]},
      {"name": "column", "role": "category", "axis_id": "source-column", "keys": ["A", "B"]}],
     "duplicate"),
    ([{"name": "row", "role": "record", "keys": [-0.0, 0.0]},
      {"name": "column", "role": "category", "axis_id": "source-column", "keys": ["A", "B"]}],
     "duplicate"),
    ([{"name": "row", "role": "record", "keys": [1, [2]]},
      {"name": "column", "role": "category", "axis_id": "source-column", "keys": ["A", "B"]}],
     "finite JSON scalars"),
    ([{"name": "row", "role": "record", "keys": [1, float("nan")]},
      {"name": "column", "role": "category", "axis_id": "source-column", "keys": ["A", "B"]}],
     "finite JSON scalars"),
    ([{"name": "row", "role": "record", "keys": [1, 2]}], "axis rank"),
    ([{"name": "other", "role": "record", "keys": [1, 2]},
      {"name": "column", "role": "category", "axis_id": "source-column", "keys": ["A", "B"]}],
     "boundary identity"),
    ([{"name": "row", "role": "record", "keys": [1, 2]},
      {"name": "column", "role": "category", "keys": ["A", "B"]}],
     "boundary identity"),
])
def test_design_axis_keys_reject_invalid_or_changed_boundary_axes(axes: list[dict], message: str) -> None:
    boundary = {"variables": [{"variable_id": "grid", "logical_name": "Grid", "kind": "table",
                              "shape": [2, 2], "axes": [{"name": "row", "role": "record"},
                                                       {"name": "column", "role": "category",
                                                        "axis_id": "source-column"}]}]}
    declaration = {**_declaration("Grid"), "axes": axes}
    with pytest.raises(ValueError, match=message):
        _build_contract({"design": {"tool_interface": {"objects": [declaration]}}}, boundary)


def test_cli_reports_generation_without_execution_claim(tmp_path: Path) -> None:
    root = _approved_workflow(tmp_path)
    result = CliRunner().invoke(app, ["step3", "interface", "--workflow", str(root)])
    assert result.exit_code == 2
    assert "read-only" in result.output
    output = export_tool_interface(root)
    assert output["message"] == "interface generated"
    assert output["runtime_readiness"]["status"] == "blocked"
    assert "Excel equivalence" not in json.dumps(output)


def test_status_marks_changed_interface_artifact_stale(tmp_path: Path) -> None:
    root = _approved_workflow(tmp_path)
    result = export_tool_interface(root)
    path = root / result["artifact"]["md"]
    path.write_bytes(path.read_bytes() + b"changed")
    assert workflow.workflow_status(root)["tool_interface"]["status"] == "stale"
    assert export_tool_interface(root)["status"] == "blocked"


@pytest.mark.parametrize("failed_write", ["interface.md", "workflow.json"])
def test_export_retries_matching_orphans_after_write_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failed_write: str,
) -> None:
    root = _approved_workflow(tmp_path)
    original_write = workflow._atomic_write

    def interrupted_write(path: Path, content: bytes) -> None:
        if path.name == failed_write:
            raise OSError("injected write failure")
        original_write(path, content)

    with monkeypatch.context() as patch:
        patch.setattr(workflow, "_atomic_write", interrupted_write)
        first = export_tool_interface(root)
    assert first["status"] == "blocked"
    assert "injected write failure" in first["diagnostics"][0]["message"]
    assert "tool_interfaces" not in workflow.load_workflow(root)[1]
    orphan_dir = next((root / "tool_interfaces").iterdir())
    json_path, md_path = orphan_dir / "interface.json", orphan_dir / "interface.md"
    json_before = json_path.read_bytes()
    md_before = md_path.read_bytes() if md_path.exists() else None

    retry = export_tool_interface(root)
    assert retry["status"] == "pass"
    assert retry["definition_sha256"] == orphan_dir.name
    assert json_path.read_bytes() == json_before
    if md_before is not None:
        assert md_path.read_bytes() == md_before
    assert workflow.hash_file(md_path) == retry["artifact"]["md_sha256"]
    assert len(workflow.load_workflow(root)[1]["tool_interfaces"]) == 1


def test_export_rejects_conflicting_orphan_bytes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    root = _approved_workflow(tmp_path)
    original_write = workflow._atomic_write

    def interrupted_write(path: Path, content: bytes) -> None:
        if path.name == "interface.md":
            raise OSError("injected write failure")
        original_write(path, content)

    with monkeypatch.context() as patch:
        patch.setattr(workflow, "_atomic_write", interrupted_write)
        assert export_tool_interface(root)["status"] == "blocked"
    orphan_dir = next((root / "tool_interfaces").iterdir())
    json_path = orphan_dir / "interface.json"
    json_path.write_bytes(b"conflicting bytes")

    retry = export_tool_interface(root)
    assert retry["status"] == "blocked"
    assert "different contents" in retry["diagnostics"][0]["message"]
    assert json_path.read_bytes() == b"conflicting bytes"
    assert "tool_interfaces" not in workflow.load_workflow(root)[1]
