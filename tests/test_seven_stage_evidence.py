"""Synthetic evidence-contract fixtures; these never attach or register a tool."""

import hashlib
import json
from pathlib import Path

import pytest

from excel_to_act.steps import seven_stage
from excel_to_act.steps.conversion_workflow import hash_file, json_bytes


def _write(path: Path, value: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(json_bytes(value))
    return path


def _input_sha(inputs: dict) -> str:
    payload = json.dumps(inputs, sort_keys=True, ensure_ascii=False, separators=(",", ":")) + "\n"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _evidence(tmp_path: Path) -> tuple[dict, Path, dict, Path, dict]:
    source, stage4 = "a" * 64, "b" * 64
    interface = "d" * 64
    contract = {"source_sha256": source, "interface_sha256": interface,
                "support_domain": {"synthetic-domain": True},
                "mutable": {f"synthetic-axis-{number}": {
                    "bindings": [{"address": f"synthetic!A{number + 1}"}]}
                    for number in range(4)}}
    layout = _write(tmp_path / "package" / "input_layout.json", {
        "source_sha256": source, "variable_contract": contract})
    bundle_manifest = _write(tmp_path / "package" / "model_manifest.json", {
        "schema_version": "step4.modular_model.v1", "source_sha256": source,
        "formula_cache_inputs": False, "target_names": ["synthetic-target"],
        "variable_input_contract_sha256": _input_sha(contract),
        "files": {"input_layout.json": hash_file(layout)},
    })
    bundle = hash_file(bundle_manifest)
    oracle = _write(tmp_path / "oracle.json", {
        "schema_version": "gp.excel_oracle.v1", "status": "pass", "engine": "Microsoft Excel",
        "source_sha256": source, "source_copy_sha256": source,
        "stage4_artifact_sha256": stage4,
    })
    standalone_result = _write(tmp_path / "standalone-result.json", {"synthetic": True})
    standalone = _write(tmp_path / "standalone.json", {
        "schema_version": "step5.execution_validation.v1", "status": "pass",
        "standalone_run_pass": True, "compile_pass": True,
        "source_sha256": source, "stage4_artifact_sha256": stage4,
        "generation_report_sha256": stage4, "result_path": str(standalone_result),
        "result_sha256": hash_file(standalone_result),
    })
    stage5 = {
        "schema_version": "step5.validation_report.v1", "status": "ready_for_review",
        "source_sha256": source, "stage4_artifact_sha256": stage4, "mismatch_count": 0,
        "formula_cache_inputs": False, "standalone_project_imports": False,
        "oracle": {"path": str(oracle), "sha256": hash_file(oracle),
                   "engine": "Microsoft Excel", "source_copy_sha256": source},
        "standalone_validation": {"path": str(standalone), "sha256": hash_file(standalone),
                                  "result_sha256": hash_file(standalone_result)},
        "test_fixture": True,
    }
    cases = []
    axis_cases = []
    for index in range(32):
        scenario = f"synthetic-{index}"
        folder = tmp_path / scenario
        axes = (index // 8, (index // 4) % 2, (index // 2) % 2, index % 2)
        inputs = {f"synthetic-axis-{number}": {"values": value, "axes": []}
                  for number, value in enumerate(axes)}
        input_sha = _input_sha(inputs)
        native_input_sha = hashlib.sha256(f"synthetic-native-{index}".encode()).hexdigest()
        native_overrides = [{"sheet": "Synthetic", "address": f"A{number + 1}",
                             "qualified_address": f"Synthetic!A{number + 1}", "value": value}
                            for number, value in enumerate(axes)]
        request = _write(folder / "request.json", {
            "synthetic": True, "inputs": inputs, "external_data": {}})
        native_request = _write(folder / "native-request.json", {
            "source_sha256": source, "scenario_id": scenario,
            "input_sha256": native_input_sha, "input_overrides": native_overrides,
            "targets": ["synthetic-target", "synthetic-diagnostic"]})
        native_result = _write(folder / "native-result.json", {
            "source_sha256": source, "source_copy_sha256": source,
            "engine": "Microsoft Excel", "scenario_id": scenario,
            "input_sha256": native_input_sha, "input_overrides": native_overrides,
            "named_targets": {"synthetic-target": {"value": index},
                              "synthetic-diagnostic": {"value": index + 1}}})
        producer = _write(folder / "producer.json", {
            "synthetic": True, "case": scenario, "source_sha256": source,
            "native_request_sha256": hash_file(native_request),
            "input_sha256": native_input_sha, "raw_input_count": len(native_overrides),
            "formula_cache_inputs": False})
        public = _write(folder / "public-case" / "model_result.json", {
            "synthetic": True, "index": index, "input_sha256": input_sha,
            "schema_version": "step4.modular_model_result.v1", "status": "pass",
            "source_sha256": source, "dependency_sha256": _input_sha({}),
            "interface_sha256": interface,
            "raw_inputs": {f"synthetic!A{number + 1}": value
                           for number, value in enumerate(axes)},
            "support_domain_sha256": _input_sha(contract["support_domain"]),
            "targets": {"synthetic-target": index}})
        case = {"case": scenario, "status": "pass", "mismatch_count": 0,
                "age": axes[0], "sex": axes[1], "benefit_term": axes[2],
                "payment_period": axes[3],
                "synthetic-target": index, "synthetic-diagnostic": index + 1,
                "public_input_sha256": input_sha,
                "request_sha256": hash_file(request),
                "native_request_sha256": hash_file(native_request),
                "native_result_sha256": hash_file(native_result),
                "producer_sha256": hash_file(producer),
                "public_result_sha256": hash_file(public),
                "execution_workdir": str(public.parent)}
        cases.append(case)
        axis_cases.append({"case": scenario, "source_sha256": source,
                           "native_input_sha256": native_input_sha,
                           "public_input_sha256": input_sha,
                           "formula_cache_inputs": False, "public_native_raw_inputs_equal": True,
                           "raw_input_mismatch_count": 0,
                           "external_values_equal_approved_capture_source_mapping": True,
                           **{key: case[key] for key in (
                               "request_sha256", "native_request_sha256",
                               "native_result_sha256", "producer_sha256")},
                           "request_path": str(request), "native_request_path": str(native_request),
                           "native_result_path": str(native_result), "producer_path": str(producer)})
        if index == 0:
            first_native_request, first_native_result, first_public = native_request, native_result, public
    axis = {"schema_version": "issue45.current_request_axis_rebind.v1", "case_count": 32,
            "source_sha256": source, "stage4_report_sha256": stage4,
            "bundle_manifest_sha256": bundle,
            "cases": axis_cases}
    axis_path = _write(tmp_path / "current-request-axis-rebind-manifest.json", axis)
    comparison_path = tmp_path / "public-execution" / "comparison.json"
    comparison = {"schema_version": "issue45.public_matrix_comparison.v2", "status": "pass",
                  "mismatch_count": 0, "source_sha256": source,
                  "stage4_report_sha256": stage4,
                  "stage4_bundle_manifest_sha256": bundle,
                  "formula_cache_inputs": False, "public_pass_count": 32,
                  "tolerance": {"atol": 0, "rtol": 0},
                  "sample": {"count": 32, "ages": list(range(4)), "sex_tokens": [0, 1],
                             "benefit_terms": [0, 1], "payment_periods": [0, 1]},
                  "request_axis_manifest_sha256": hash_file(axis_path), "cases": cases}
    _write(comparison_path, comparison)
    return stage5, oracle, comparison, comparison_path, {
        "source": source, "stage4": stage4, "bundle": bundle,
        "bundle_manifest": bundle_manifest,
        "native_request": first_native_request, "native_result": first_native_result,
        "axis": axis_path, "public": first_public}


def test_supported_evidence_shape_and_referenced_files(tmp_path: Path) -> None:
    stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    assert len(comparison["cases"]) == 32
    seven_stage._require_stage5(stage5, paths["source"], paths["stage4"])
    seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                    paths["stage4"], paths["bundle"], hash_file(comparison_path), paths["bundle_manifest"])


def test_synthetic_duplicated_effective_domain_rejected(tmp_path: Path) -> None:
    """Distinct files cannot turn one effective request into a 32-case sample."""
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    first = axis["cases"][0]
    first_request = json.loads(Path(first["request_path"]).read_text())
    first_native = json.loads(Path(first["native_request_path"]).read_text())
    first_input_sha = comparison["cases"][0]["public_input_sha256"]
    for case, bound in zip(comparison["cases"][1:], axis["cases"][1:]):
        request_path = Path(bound["request_path"])
        request = json.loads(request_path.read_text())
        request["inputs"] = first_request["inputs"]
        request["case_label"] = case["case"]
        _write(request_path, request)
        case["request_sha256"] = bound["request_sha256"] = hash_file(request_path)
        native_path = Path(bound["native_request_path"])
        native = json.loads(native_path.read_text())
        native["input_overrides"] = first_native["input_overrides"]
        native["input_sha256"] = first_native["input_sha256"]
        _write(native_path, native)
        case["native_request_sha256"] = bound["native_request_sha256"] = hash_file(native_path)
        bound["native_input_sha256"] = first_native["input_sha256"]
        result_path = Path(bound["native_result_path"])
        result = json.loads(result_path.read_text())
        result["input_overrides"] = first_native["input_overrides"]
        result["input_sha256"] = first_native["input_sha256"]
        _write(result_path, result)
        case["native_result_sha256"] = bound["native_result_sha256"] = hash_file(result_path)
        producer_path = Path(bound["producer_path"])
        producer = json.loads(producer_path.read_text())
        producer["native_request_sha256"] = hash_file(native_path)
        producer["input_sha256"] = first_native["input_sha256"]
        _write(producer_path, producer)
        case["producer_sha256"] = bound["producer_sha256"] = hash_file(producer_path)
        public_path = Path(case["execution_workdir"]) / "model_result.json"
        public = json.loads(public_path.read_text())
        public["input_sha256"] = first_input_sha
        public["raw_inputs"] = {f"{item['sheet'].casefold()}!{item['address']}": item["value"]
                                for item in first_native["input_overrides"]}
        _write(public_path, public)
        case["public_result_sha256"] = hash_file(public_path)
        case["public_input_sha256"] = bound["public_input_sha256"] = first_input_sha
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                        paths["bundle_manifest"])


def test_synthetic_axis_marker_rejected_even_when_comparison_unmarked(tmp_path: Path) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    axis["lineage"] = {"test_fixture": True}
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                        paths["bundle_manifest"])


@pytest.mark.parametrize("fault", ["duplicate_tuple", "outside_tuple", "sample_marker"])
def test_synthetic_declared_domain_must_be_complete_and_unmarked(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    if fault == "duplicate_tuple":
        for field in ("age", "sex", "benefit_term", "payment_period"):
            comparison["cases"][1][field] = comparison["cases"][0][field]
    elif fault == "outside_tuple":
        comparison["cases"][0]["age"] = 99
    else:
        comparison["sample"]["lineage"] = {"test_fixture": True}
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                        paths["bundle_manifest"])


@pytest.mark.parametrize("field,value", [
    ("schema_version", "wrong.result.v1"), ("status", "failed"),
    ("source_sha256", "0" * 64), ("interface_sha256", "0" * 64),
    ("dependency_sha256", "0" * 64), ("support_domain_sha256", "0" * 64),
    ("source_sha256", None), ("interface_sha256", []),
])
def test_rehashed_synthetic_public_result_identity_cannot_pass(
        tmp_path: Path, field: str, value: object) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    public = json.loads(paths["public"].read_text())
    public[field] = value
    _write(paths["public"], public)
    comparison["cases"][0]["public_result_sha256"] = hash_file(paths["public"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"],
                                        hash_file(comparison_path), paths["bundle_manifest"])


def test_isolated_synthetic_entry_hash_cannot_validate_wrong_source(tmp_path: Path) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    public = json.loads(paths["public"].read_text())
    public["source_sha256"] = "0" * 64
    _write(paths["public"], public)
    package = tmp_path / "synthetic-entry"
    package.mkdir()
    (package / "model.py").write_text(
        "import argparse\nfrom pathlib import Path\n"
        "p=argparse.ArgumentParser();p.add_argument('--inputs');p.add_argument('--out');a=p.parse_args()\n"
        "Path(a.out).write_bytes((Path(__file__).parent/'payload.json').read_bytes())\n",
        encoding="utf-8")
    (package / "payload.json").write_bytes(paths["public"].read_bytes())
    request = tmp_path / "synthetic-0" / "request.json"
    result = tmp_path / "isolated-result.json"
    invocation = seven_stage._run_entry(package, request, result)
    assert invocation["exit_code"] == 0
    assert hash_file(result) == hash_file(paths["public"])
    comparison["cases"][0]["public_result_sha256"] = hash_file(result)
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"],
                                        hash_file(comparison_path), paths["bundle_manifest"])


@pytest.mark.parametrize("fault", [
    "one_case", "coherent_one_case", "missing_case", "duplicate_case",
    "duplicate_request", "wrong_axis_membership",
])
def test_original_sample_requires_32_distinct_bound_cases(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    if fault in {"one_case", "coherent_one_case"}:
        comparison["cases"] = comparison["cases"][:1]
        comparison["public_pass_count"] = 1
        comparison["sample"]["count"] = 1
        if fault == "coherent_one_case":
            axis["cases"] = axis["cases"][:1]
            axis["case_count"] = 1
    elif fault == "missing_case":
        comparison["cases"].pop()
        comparison["public_pass_count"] = 31
        comparison["sample"]["count"] = 31
        axis["cases"].pop()
        axis["case_count"] = 31
    elif fault == "duplicate_case":
        comparison["cases"][1]["case"] = comparison["cases"][0]["case"]
    elif fault == "duplicate_request":
        comparison["cases"][1]["request_sha256"] = comparison["cases"][0]["request_sha256"]
        axis["cases"][1]["request_sha256"] = axis["cases"][0]["request_sha256"]
        axis["cases"][1]["request_path"] = axis["cases"][0]["request_path"]
    else:
        axis["cases"][1]["case"] = "synthetic-wrong"
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path), paths["bundle_manifest"])


@pytest.mark.parametrize("fault", ["replaced_id", "substituted_request"])
def test_pinned_comparison_rejects_replaced_sample_identity(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    pinned_sha = hash_file(comparison_path)
    axis = json.loads(paths["axis"].read_text())
    comparison["cases"][0]["case"] = axis["cases"][0]["case"] = "synthetic-outside"
    if fault == "substituted_request":
        replacement = _write(tmp_path / "outside-request.json", {"synthetic": True, "outside": True})
        comparison["cases"][0]["request_sha256"] = hash_file(replacement)
        axis["cases"][0]["request_sha256"] = hash_file(replacement)
        axis["cases"][0]["request_path"] = str(replacement)
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="pinned identity"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], pinned_sha, paths["bundle_manifest"])


@pytest.mark.parametrize("fault", ["test_fixture", "empty_targets", "missing_targets",
                                   "malformed_targets", "incomplete_targets", "wrong_native_value",
                                   "wrong_public_value"])
def test_native_targets_must_support_declared_comparison(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    if fault == "wrong_public_value":
        public = json.loads(paths["public"].read_text())
        public["targets"]["synthetic-target"] = -1
        _write(paths["public"], public)
        comparison["cases"][0]["public_result_sha256"] = hash_file(paths["public"])
    else:
        native = json.loads(paths["native_result"].read_text())
        if fault == "test_fixture":
            native["test_fixture"] = True
        elif fault == "empty_targets":
            native["named_targets"] = {}
        elif fault == "missing_targets":
            native.pop("named_targets")
        elif fault == "malformed_targets":
            native["named_targets"] = []
        elif fault == "incomplete_targets":
            native["named_targets"] = {"other": {"value": 0}}
        else:
            native["named_targets"]["synthetic-target"]["value"] = -1
        _write(paths["native_result"], native)
        comparison["cases"][0]["native_result_sha256"] = hash_file(paths["native_result"])
        axis["cases"][0]["native_result_sha256"] = hash_file(paths["native_result"])
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path), paths["bundle_manifest"])


@pytest.mark.parametrize("fault", [
    "reused_native_request", "reused_native_scenario", "wrong_producer_case", "wrong_raw_input",
])
def test_native_evidence_joins_exact_public_case(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    first, second = axis["cases"][:2]
    case = comparison["cases"][1]
    native_request_path = Path(second["native_request_path"])
    native_result_path = Path(second["native_result_path"])
    producer_path = Path(second["producer_path"])
    native_request = json.loads(native_request_path.read_text())
    native_result = json.loads(native_result_path.read_text())
    producer = json.loads(producer_path.read_text())
    if fault == "reused_native_request":
        reused = json.loads(Path(first["native_request_path"]).read_text())
        second["native_request_path"] = first["native_request_path"]
        second["native_request_sha256"] = case["native_request_sha256"] = first["native_request_sha256"]
        second["native_input_sha256"] = first["native_input_sha256"]
        native_result["scenario_id"] = reused["scenario_id"]
        native_result["input_sha256"] = reused["input_sha256"]
    elif fault == "reused_native_scenario":
        native_request["scenario_id"] = native_result["scenario_id"] = json.loads(
            Path(first["native_request_path"]).read_text())["scenario_id"]
        _write(native_request_path, native_request)
        second["native_request_sha256"] = case["native_request_sha256"] = hash_file(native_request_path)
        producer["native_request_sha256"] = hash_file(native_request_path)
        _write(producer_path, producer)
        second["producer_sha256"] = case["producer_sha256"] = hash_file(producer_path)
    elif fault == "wrong_producer_case":
        producer["case"] = first["case"]
        _write(producer_path, producer)
        second["producer_sha256"] = case["producer_sha256"] = hash_file(producer_path)
    else:
        native_request["input_overrides"][0]["value"] = -1
        native_result["input_overrides"][0]["value"] = -1
        _write(native_request_path, native_request)
        second["native_request_sha256"] = case["native_request_sha256"] = hash_file(native_request_path)
        producer["native_request_sha256"] = hash_file(native_request_path)
        _write(producer_path, producer)
        second["producer_sha256"] = case["producer_sha256"] = hash_file(producer_path)
    if fault in {"reused_native_request", "reused_native_scenario", "wrong_raw_input"}:
        _write(native_result_path, native_result)
        second["native_result_sha256"] = case["native_result_sha256"] = hash_file(native_result_path)
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                        paths["bundle_manifest"])


def test_missing_declared_output_rejected_even_with_native_diagnostic(tmp_path: Path) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    public = json.loads(paths["public"].read_text())
    public["targets"] = {"synthetic-diagnostic": 1}
    _write(paths["public"], public)
    comparison["cases"][0]["public_result_sha256"] = hash_file(paths["public"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                        paths["bundle_manifest"])


def test_published_output_and_extra_native_diagnostic_are_valid(tmp_path: Path) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    native_request = json.loads(paths["native_request"].read_text())
    assert len(native_request["targets"]) > len(json.loads(paths["public"].read_text())["targets"])
    assert axis["cases"][0]["native_request_sha256"] == comparison["cases"][0]["native_request_sha256"]
    seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                    paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                    paths["bundle_manifest"])


def test_output_contract_manifest_must_match_bound_package(tmp_path: Path) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    manifest = json.loads(paths["bundle_manifest"].read_text())
    manifest["target_names"] = ["synthetic-diagnostic"]
    _write(paths["bundle_manifest"], manifest)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path),
                                        paths["bundle_manifest"])


@pytest.mark.parametrize("fault", [
    "native_result_wrong", "native_request_wrong", "native_axis_wrong",
    "native_result_missing", "native_result_malformed",
    "public_result_wrong", "public_case_wrong", "public_axis_wrong",
    "public_result_missing", "public_result_malformed", "public_request_changed",
])
def test_exact_request_input_lineage_rejects_rebound_evidence(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    bound = axis["cases"][0]
    case = comparison["cases"][0]
    if fault.startswith("native_result"):
        result = json.loads(paths["native_result"].read_text())
        if fault.endswith("missing"):
            result.pop("input_sha256")
        else:
            result["input_sha256"] = "bad" if fault.endswith("malformed") else "0" * 64
        _write(paths["native_result"], result)
        bound["native_result_sha256"] = case["native_result_sha256"] = hash_file(paths["native_result"])
    elif fault == "native_request_wrong":
        request = json.loads(paths["native_request"].read_text())
        request["input_sha256"] = "0" * 64
        _write(paths["native_request"], request)
        bound["native_request_sha256"] = case["native_request_sha256"] = hash_file(paths["native_request"])
    elif fault == "native_axis_wrong":
        bound["native_input_sha256"] = "0" * 64
    elif fault.startswith("public_result"):
        result = json.loads(paths["public"].read_text())
        if fault.endswith("missing"):
            result.pop("input_sha256")
        else:
            result["input_sha256"] = "bad" if fault.endswith("malformed") else "0" * 64
        _write(paths["public"], result)
        case["public_result_sha256"] = hash_file(paths["public"])
    elif fault == "public_case_wrong":
        case["public_input_sha256"] = "0" * 64
    elif fault == "public_axis_wrong":
        bound["public_input_sha256"] = "0" * 64
    else:
        request = json.loads(Path(bound["request_path"]).read_text())
        request["inputs"]["synthetic-axis-0"]["values"] = -1
        _write(Path(bound["request_path"]), request)
        bound["request_sha256"] = case["request_sha256"] = hash_file(Path(bound["request_path"]))
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path), paths["bundle_manifest"])


@pytest.mark.parametrize("where", ["native_request", "native_target", "public_result"])
def test_nested_synthetic_marker_rejected_from_comparison(tmp_path: Path, where: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    axis = json.loads(paths["axis"].read_text())
    if where == "native_request":
        path, key = paths["native_request"], "native_request_sha256"
        record = json.loads(path.read_text())
        record["metadata"] = [{"test_fixture": True}]
    elif where == "native_target":
        path, key = paths["native_result"], "native_result_sha256"
        record = json.loads(path.read_text())
        record["named_targets"]["synthetic-target"]["test_fixture"] = True
    else:
        path, key = paths["public"], "public_result_sha256"
        record = json.loads(path.read_text())
        record["metadata"] = {"test_fixture": True}
    _write(path, record)
    comparison["cases"][0][key] = hash_file(path)
    if where != "public_result":
        axis["cases"][0][key] = hash_file(path)
    _write(paths["axis"], axis)
    comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path), paths["bundle_manifest"])


@pytest.mark.parametrize("where", ["oracle", "oracle_nested", "standalone", "result"])
def test_stage5_referenced_synthetic_marker_rejected(tmp_path: Path, where: str) -> None:
    stage5, oracle, _comparison, _comparison_path, paths = _evidence(tmp_path)
    standalone = Path(stage5["standalone_validation"]["path"])
    if where.startswith("oracle"):
        record = json.loads(oracle.read_text())
        if where == "oracle_nested":
            record["named_targets"] = {"synthetic-target": {"test_fixture": True}}
        else:
            record["test_fixture"] = True
        _write(oracle, record)
        stage5["oracle"]["sha256"] = hash_file(oracle)
    elif where == "standalone":
        record = json.loads(standalone.read_text())
        record["test_fixture"] = True
        _write(standalone, record)
        stage5["standalone_validation"]["sha256"] = hash_file(standalone)
    else:
        record = json.loads(Path(json.loads(standalone.read_text())["result_path"]).read_text())
        record["test_fixture"] = True
        result = Path(json.loads(standalone.read_text())["result_path"])
        _write(result, record)
        validation = json.loads(standalone.read_text())
        validation["result_sha256"] = stage5["standalone_validation"]["result_sha256"] = hash_file(result)
        _write(standalone, validation)
        stage5["standalone_validation"]["sha256"] = hash_file(standalone)
    with pytest.raises(ValueError, match="Step 5"):
        seven_stage._require_stage5(stage5, paths["source"], paths["stage4"])


@pytest.mark.parametrize("changed", ["native_result", "public"])
def test_pinned_evidence_file_changes_invalidate_comparison(tmp_path: Path, changed: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    pinned_sha = hash_file(comparison_path)
    paths[changed].write_bytes(b"changed")
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], pinned_sha, paths["bundle_manifest"])


@pytest.mark.parametrize("fault", [
    "declaration_only", "wrong_schema", "missing_axis", "stale_axis",
    "missing_native_request", "corrupt_native_request", "missing_native_result",
    "corrupt_native_result", "stale_scenario", "stale_raw_binding",
    "missing_public_result", "malformed_axis",
])
def test_comparison_declarations_cannot_replace_native_files(tmp_path: Path, fault: str) -> None:
    _stage5, _oracle, comparison, comparison_path, paths = _evidence(tmp_path)
    if fault == "declaration_only":
        comparison.pop("request_axis_manifest_sha256")
        comparison["cases"][0].pop("native_request_sha256")
        comparison["cases"][0].pop("native_result_sha256")
    elif fault == "wrong_schema":
        comparison["schema_version"] = "made-up"
    elif fault == "missing_axis":
        paths["axis"].unlink()
    elif fault == "stale_axis":
        paths["axis"].write_bytes(b"changed")
    elif fault == "missing_native_request":
        paths["native_request"].unlink()
    elif fault == "corrupt_native_request":
        paths["native_request"].write_bytes(b"changed")
    elif fault == "missing_native_result":
        paths["native_result"].unlink()
    elif fault == "corrupt_native_result":
        paths["native_result"].write_bytes(b"changed")
    elif fault == "stale_scenario":
        comparison["cases"][0]["request_sha256"] = "0" * 64
    elif fault == "stale_raw_binding":
        axis = json.loads(paths["axis"].read_text())
        axis["cases"][0]["public_native_raw_inputs_equal"] = False
        _write(paths["axis"], axis)
        comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    elif fault == "malformed_axis":
        paths["axis"].write_text("[]", encoding="utf-8")
        comparison["request_axis_manifest_sha256"] = hash_file(paths["axis"])
    else:
        paths["public"].unlink()
    _write(comparison_path, comparison)
    with pytest.raises(ValueError, match="native comparison"):
        seven_stage._require_comparison(comparison, comparison_path, paths["source"],
                                        paths["stage4"], paths["bundle"], hash_file(comparison_path), paths["bundle_manifest"])


@pytest.mark.parametrize("fault", ["declaration_only", "missing_oracle", "corrupt_oracle",
                                   "stale_oracle", "malformed_oracle",
                                   "missing_standalone", "corrupt_standalone"])
def test_stage5_declarations_cannot_replace_oracle_files(tmp_path: Path, fault: str) -> None:
    stage5, oracle, _comparison, _path, paths = _evidence(tmp_path)
    standalone = Path(stage5["standalone_validation"]["path"])
    if fault == "declaration_only":
        stage5["oracle"] = {"engine": "Microsoft Excel"}
        stage5.pop("standalone_validation")
    elif fault == "missing_oracle":
        oracle.unlink()
    elif fault == "corrupt_oracle":
        oracle.write_bytes(b"changed")
    elif fault == "stale_oracle":
        value = json.loads(oracle.read_text())
        value["stage4_artifact_sha256"] = "0" * 64
        _write(oracle, value)
        stage5["oracle"]["sha256"] = hash_file(oracle)
    elif fault == "malformed_oracle":
        oracle.write_text("[]", encoding="utf-8")
        stage5["oracle"]["sha256"] = hash_file(oracle)
    elif fault == "missing_standalone":
        standalone.unlink()
    else:
        standalone.write_bytes(b"changed")
    with pytest.raises(ValueError, match="Step 5"):
        seven_stage._require_stage5(stage5, paths["source"], paths["stage4"])
