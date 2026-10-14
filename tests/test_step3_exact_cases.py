"""Finite case production uses bound evidence; these fixtures are synthetic only."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path

from openpyxl import Workbook
import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps.step3 import exact_cases
from excel_to_act.steps.step3 import design as step3_design
from excel_to_act.steps.step3.design import _validate_exact_cases_for_design


def _write(path: Path, value: dict) -> dict[str, str]:
    path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")
    return {"path": str(path), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def _semantic_pair_fixture(tmp_path: Path) -> tuple[dict, str, str]:
    core = {name: _write(tmp_path / f"{name}.json", {"kind": name})
            for name in ("trace", "raw_request", "public_request", "producer")}
    manifest = _write(tmp_path / "cases.json", {"cases": [{"case_id": "case-1", **core}]})
    source_sha, boundary_sha = "a" * 64, "b" * 64
    identity = {"case_id": "case-1", "case_manifest": manifest,
                "source_sha256": source_sha, "boundary_sha256": boundary_sha, **core}
    inventory = _write(tmp_path / "inventory.json", {"synthetic": True})
    profile = _write(tmp_path / "profile.json", {
        "schema_version": "step3.exact_case_family_profile.v1", "status": "pass",
        "selection_basis": "exact_case", "case_identity": identity,
        "exact_case_trace": core["trace"], "inventory": inventory})
    source = {"source_sha256": source_sha}
    semantic_map = _write(tmp_path / "map.json", {
        "selection_basis": "exact_case", "case_id": "case-1", "case_manifest": manifest,
        "exact_case_trace": core["trace"], "source_family_profile": profile,
        "source": source})
    plan = _write(tmp_path / "plan.json", {
        "schema_version": "step3.semantic_plan.v1", "status": "pass",
        "selection_basis": "exact_case", "source": source,
        "case_identity": identity, "semantic_map_input": semantic_map,
        "reference_resolution": {}, "execution_plan": {"passes": []},
        "evidence": {"source_family_profile": profile, "exact_case_trace": core["trace"]}})
    return {**core, "profile": profile, "plan": plan}, source_sha, boundary_sha


@pytest.mark.parametrize("label", ("trace", "raw_request", "public_request", "producer"))
def test_exact_case_semantic_pair_binds_distinct_trace_request_and_nested_map(
        tmp_path: Path, label: str) -> None:
    case, source_sha, boundary_sha = _semantic_pair_fixture(tmp_path)
    plan, profile, nested = exact_cases.verify_case_semantic_pair(case, source_sha, boundary_sha)
    assert plan["case_identity"] == profile["case_identity"]
    assert set(nested) == {"case_manifest", "semantic_map_input", "inventory"}
    with pytest.raises(ValueError, match="another source, boundary, trace or request"):
        exact_cases.verify_case_semantic_pair(case, source_sha, "c" * 64)
    other_raw = _write(tmp_path / f"other-{label}.json", {"kind": label, "case": 2})
    with pytest.raises(ValueError, match="another source, boundary, trace or request"):
        exact_cases.verify_case_semantic_pair({**case, label: other_raw}, source_sha, boundary_sha)
    with pytest.raises(ValueError, match="another source, boundary, trace or request"):
        exact_cases.verify_case_semantic_pair(case, "c" * 64, boundary_sha)
    no_execution = json.loads(Path(case["plan"]["path"]).read_text())
    no_execution["execution_plan"] = None
    no_execution_ref = _write(tmp_path / "no-execution.json", no_execution)
    with pytest.raises(ValueError, match="another source, boundary, trace or request"):
        exact_cases.verify_case_semantic_pair({**case, "plan": no_execution_ref}, source_sha, boundary_sha)
    map_path = nested["semantic_map_input"]
    changed = json.loads(map_path.read_text())
    changed["case_id"] = "case-2"
    _write(map_path, changed)
    with pytest.raises(ValueError, match="missing or changed"):
        exact_cases.verify_case_semantic_pair(case, source_sha, boundary_sha)


def test_exact_case_design_checks_analysis_artifacts_after_loading_case_files(
        tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"synthetic workbook")
    source_sha = hashlib.sha256(source.read_bytes()).hexdigest()
    raw = _write(tmp_path / "raw.json", {"input_overrides": []})
    public = _write(tmp_path / "public.json", {"inputs": {"Case": {"axes": [], "values": 1}}})
    boundary = _write(tmp_path / "boundary.json", {"sample_boundary": {
        "request_cases": [public], "request_case_count": 1, "case_axes": {"Case": [1]}}})
    native = _write(tmp_path / "native.json", {"synthetic": True})
    producer = _write(tmp_path / "producer.json", {
        "evidence": {"native_oracle": native, "raw_request": raw, "public_request": public},
        "boundary_json": {"sha256": boundary["sha256"]},
        "source_sha256": source_sha, "input_sha256": exact_cases._digest([])})
    trace = _write(tmp_path / "trace.json", {
        "case_id": "case-1", "source_sha256": source_sha,
        "input_sha256": exact_cases._digest([]), "request_sha256": raw["sha256"],
        "producer_sha256": producer["sha256"]})
    plan_ref = _write(tmp_path / "plan.json", {"synthetic": True})
    profile_ref = _write(tmp_path / "profile.json", {"synthetic": True})
    fields = _write(tmp_path / "fields.json", {"synthetic": True})
    dependencies = _write(tmp_path / "dependencies.json", {"synthetic": True})
    artifacts = {"fields": fields, "dependencies": dependencies}
    analysis = {"binding_sha256": "c" * 64,
                "binding": {"artifacts": {"inventory_sha256": "d" * 64}}}
    plan = {"case_identity": {"case_id": "case-1"}, "binding_sha256": "c" * 64,
            "coverage": {"unmapped_active_formula_members": 0,
                         "duplicate_active_formula_members": 0,
                         "active_formula_members": 1, "active_formula_members_mapped": 1},
            "source": {"fields_sha256": fields["sha256"],
                       "dependencies_sha256": dependencies["sha256"]}}
    profile = {"binding_sha256": "c" * 64, "inventory": {"sha256": "d" * 64}}
    monkeypatch.setattr(step3_design, "verify_native_case_producer", lambda *_: None)
    monkeypatch.setattr(step3_design, "verify_case_semantic_pair", lambda *_: (plan, profile, {}))
    case = {"public_request": public, "raw_request": raw, "producer": producer,
            "trace": trace, "plan": plan_ref, "profile": profile_ref}
    _validate_exact_cases_for_design([case], {"revision": 1}, Path(boundary["path"]),
                                     source, analysis=analysis, artifacts=artifacts)
    wrong_fields = {**artifacts, "fields": {**fields, "sha256": "e" * 64}}
    with pytest.raises(ValueError, match="analysis fields, dependencies or inventory differ"):
        _validate_exact_cases_for_design([case], {"revision": 1}, Path(boundary["path"]),
                                         source, analysis=analysis, artifacts=wrong_fields)


def test_exact_case_profile_plan_and_check_are_supported_cli_routes(tmp_path: Path) -> None:
    runner = CliRunner()
    profile_help = runner.invoke(app, ["step3", "profile", "--help"])
    plan_help = runner.invoke(app, ["step3", "plan", "--help"])
    check_help = runner.invoke(app, ["step3", "check", "--help"])
    assert profile_help.exit_code == plan_help.exit_code == check_help.exit_code == 0
    assert "--case-manifest" in profile_help.stdout and "--case-id" in profile_help.stdout
    assert "--case-out" in plan_help.stdout and "--case-plan" in check_help.stdout
    missing_case_id = runner.invoke(app, ["step3", "profile", "--analysis", str(tmp_path),
                                          "--case-manifest", str(_write(tmp_path / "manifest.json", {})["path"]),
                                          "--out", str(tmp_path / "out")])
    assert missing_case_id.exit_code == 1
    assert "--case-manifest needs --case-id" in missing_case_id.stdout


def _fixture(tmp_path: Path, monkeypatch):
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    main = workbook.active
    main.title = "Main"
    main["A1"] = 1
    main["B1"] = "=Rates!B1"
    rates = workbook.create_sheet("Rates")
    rates["A1"] = 0
    rates["A2"] = 1
    rates["B1"] = "=Main!A1+2"
    rates["B2"] = "=Main!A1+3"
    workbook.save(source)
    workbook.close()
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    fixed_ref = _write(tmp_path / "fixed.json", {"synthetic": True})
    boundary = {
        "boundary_sha256": "a" * 64,
        "sample_boundary": {"source_workbook_sha256": source_hash,
                            "case_axes": {"Case": [1]}, "request_case_count": 1,
                            "request_raw_cell_count": 1,
                            "fixed_literal_contract": fixed_ref,
                            "fixed_source_literals": [], "fixed_source_literal_count": 0,
                            "fixed_source_fields": [],
                            "request_fields": [{"name": "Case", "source_cells": ["Main!A1"],
                                                "variable_id": "main.a1"}],
                            "external_vectors": [{"variable_id": "external.rate", "source_extent": "Rates!B1:B2",
                                                  "age_extent": "Rates!A1:A2",
                                                  "age_keys": [0, 1], "producer_version": "synthetic-v1"}]},
        "variables": [{"variable_id": "main.a1", "axes": [], "shape": [], "kind": "scalar"},
                      {"variable_id": "external.rate", "shape": [2],
                       "axes": [{"axis_id": "age", "name": "age", "role": "key"}]}],
        "target_selection": {"targets": [{"start_cells": ["Main!B1"]}]},
    }
    public = {"inputs": {"Case": {"axes": [], "values": 1}},
              "external_data": {"external.rate": {
                  "axes": [{"axis_id": "age", "name": "age", "role": "key", "keys": [0, 1]}],
                  "values": [3.0, 4.0], "source_sha256": source_hash,
                  "version": "synthetic-v1", "input_sha256": exact_cases._public_digest(
                      {"Case": {"axes": [], "values": 1}})}}}
    records = [{"sheet": "Main", "address": "A1", "qualified_address": "Main!A1", "value": 1}]
    raw = {"source_sha256": source_hash, "scenario_id": "synthetic-case-1",
           "input_overrides": records, "input_sha256": exact_cases._digest(records)}
    native = {"schema_version": "gp.excel_oracle.v1", "tool": "step5.oracle", "status": "pass",
              "engine": "Microsoft Excel", "engine_settings": {"version": "16.0", "calculation": -4105},
              "calculation_state": 0, "full_rebuild_seconds": 1.0, "macros_executed": False,
              "source_sha256": source_hash, "source_copy_sha256": source_hash,
              "scenario_id": "synthetic-case-1", "input_sha256": exact_cases._digest(records),
              "input_overrides": records, "requested_ranges": ["Rates!B1:B2"],
              "applied_inputs": {"Main!A1": {"value": 1, "formula": "1", "excel_error": None}},
              "ranges": {"Rates!B1:B2": {"address": "Rates!B1:B2", "values": [[3], [4]],
                                          "formulas": [["=Main!A1+2"], ["=Main!A1+3"]],
                                          "excel_errors": [[None], [None]]}}}
    public_ref = _write(tmp_path / "public.json", public)
    boundary["sample_boundary"]["request_cases"] = [public_ref]
    boundary_ref = _write(tmp_path / "boundary.json", boundary)
    raw_ref = _write(tmp_path / "raw.json", raw)
    native_ref = _write(tmp_path / "native.json", native)
    case = {"case_id": "synthetic-1", "raw_request": raw_ref, "public_request": public_ref,
            "native_oracle": native_ref}
    manifest_path = tmp_path / "manifest.json"
    _write(manifest_path, {"schema_version": "step3.exact_cases.manifest.v1", "cases": [case]})
    monkeypatch.setattr(exact_cases, "load_workflow", lambda _: (
        tmp_path, {"source": {"step1_source_copy_path": str(source), "workbook_sha256": source_hash}}))
    monkeypatch.setattr(exact_cases, "confirmed_boundary_reference", lambda *_: (
        {"revision": 1}, Path(boundary_ref["path"]), None))
    return source, case, manifest_path, boundary_ref


def test_synthetic_exact_case_produces_sealed_native_and_trace(tmp_path: Path, monkeypatch) -> None:
    source, case, manifest, boundary_ref = _fixture(tmp_path, monkeypatch)
    result = exact_cases.produce_exact_cases(tmp_path, manifest, tmp_path / "cases")
    assert result["status"] == "pass", result
    output = json.loads(Path(result["artifact"]).read_text())
    item = output["cases"][0]
    trace = json.loads(Path(item["trace"]["path"]).read_text())
    producer = json.loads(Path(item["producer"]["path"]).read_text())
    assert trace["producer_sha256"] == item["producer"]["sha256"]
    assert trace["cells"] and producer["numeric_equivalence"]["representation_only_count"] == 2
    exact_cases.verify_native_case_producer(producer, case, source,
                                             json.loads(Path(case["raw_request"]["path"]).read_text()),
                                             json.loads(Path(case["public_request"]["path"]).read_text()))
    refs = [{label: item[label] for label in ("trace", "raw_request", "public_request", "producer")}]
    with pytest.raises(ValueError, match="requires a distinct plan and profile"):
        _validate_exact_cases_for_design(refs, {"revision": 1}, Path(boundary_ref["path"]), source)
    with pytest.raises(ValueError, match="omits or repeats"):
        _validate_exact_cases_for_design([], {"revision": 1}, Path(boundary_ref["path"]), source)


def test_synthetic_exact_case_rejects_bad_native_cut_and_request(tmp_path: Path, monkeypatch) -> None:
    _, case, manifest, _ = _fixture(tmp_path, monkeypatch)
    native_path = Path(case["native_oracle"]["path"])
    native = json.loads(native_path.read_text())
    native["ranges"]["Rates!B1:B2"]["values"][0][0] = 9
    case["native_oracle"] = _write(native_path, native)
    _write(manifest, {"schema_version": "step3.exact_cases.manifest.v1", "cases": [case]})
    result = exact_cases.produce_exact_cases(tmp_path, manifest, tmp_path / "bad-cut")
    assert result["status"] == "blocked" and "native cut differs" in result["reason"]
    assert not (tmp_path / "bad-cut" / "exact_cases.json").exists()

    public_path = Path(case["public_request"]["path"])
    public = json.loads(public_path.read_text())
    public["inputs"]["Case"]["values"] = 2
    case["public_request"] = _write(public_path, public)
    result = exact_cases.produce_exact_cases(tmp_path, manifest, tmp_path / "bad-request")
    assert result["status"] == "blocked"


def test_synthetic_exact_case_requires_complete_unique_domain(tmp_path: Path, monkeypatch) -> None:
    _, case, manifest, boundary_ref = _fixture(tmp_path, monkeypatch)
    boundary_path = Path(boundary_ref["path"])
    boundary = json.loads(boundary_path.read_text())
    boundary["sample_boundary"]["case_axes"]["Case"] = [1, 2]
    boundary["sample_boundary"]["request_case_count"] = 2
    boundary["sample_boundary"]["request_cases"] = [case["public_request"],
                                                     {"path": "missing", "sha256": "b" * 64}]
    _write(boundary_path, boundary)
    missing = exact_cases.produce_exact_cases(tmp_path, manifest, tmp_path / "missing")
    assert missing["status"] == "blocked" and "complete case set" in missing["reason"]
    duplicate = copy.deepcopy(case)
    duplicate["case_id"] = "synthetic-2"
    _write(manifest, {"schema_version": "step3.exact_cases.manifest.v1", "cases": [case, duplicate]})
    repeated = exact_cases.produce_exact_cases(tmp_path, manifest, tmp_path / "duplicate")
    assert repeated["status"] == "blocked"


@pytest.mark.parametrize("mutation,expected", [
    ("version", "source, version, input, or axis"),
    ("axis", "source, version, input, or axis"),
    ("source", "source, version, input, or axis"),
    ("formula", "native cut differs"),
    ("readback", "native input readback"),
])
def test_synthetic_exact_case_rechecks_public_and_native_lineage(
    tmp_path: Path, monkeypatch, mutation: str, expected: str,
) -> None:
    _, case, manifest, boundary_ref = _fixture(tmp_path, monkeypatch)
    if mutation in {"version", "axis", "source"}:
        path = Path(case["public_request"]["path"])
        public = json.loads(path.read_text())
        data = public["external_data"]["external.rate"]
        if mutation == "version":
            data["version"] = "another-producer"
        elif mutation == "axis":
            data["axes"][0]["keys"] = [1, 0]
        else:
            data["source_sha256"] = "b" * 64
        case["public_request"] = _write(path, public)
        boundary = json.loads(Path(boundary_ref["path"]).read_text())
        boundary["sample_boundary"]["request_cases"] = [case["public_request"]]
        _write(Path(boundary_ref["path"]), boundary)
    else:
        path = Path(case["native_oracle"]["path"])
        native = json.loads(path.read_text())
        if mutation == "formula":
            native["ranges"]["Rates!B1:B2"]["formulas"][0][0] = "=0"
        else:
            native["applied_inputs"]["Main!A1"]["value"] = 4
        case["native_oracle"] = _write(path, native)
    _write(manifest, {"schema_version": "step3.exact_cases.manifest.v1", "cases": [case]})
    result = exact_cases.produce_exact_cases(tmp_path, manifest, tmp_path / f"bad-{mutation}")
    assert result["status"] == "blocked" and expected in result["reason"]
