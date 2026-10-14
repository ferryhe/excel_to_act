"""Attached seven-stage workflow and exact internal package delivery."""

from __future__ import annotations

import json
from itertools import product
import os
import math
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from typing import Any

from excel_to_act.steps.conversion_workflow import (
    MANIFEST_NAME, V2_STAGES, _now, _write_manifest, append_stage_artifact,
    current_revision, hash_bytes, hash_file, input_boundary_status, json_bytes,
    load_workflow, read_json, require_stage_approved, stage_status,
    workflow_origin,
)
from excel_to_act.steps.step5.workflow import _bundle_file_paths, _stage4, _verify_case_receipts


def _has_test_fixture(value: Any) -> bool:
    if isinstance(value, dict):
        return "test_fixture" in value or any(_has_test_fixture(item) for item in value.values())
    if isinstance(value, list):
        return any(_has_test_fixture(item) for item in value)
    return False


def _canonical_hash(value: Any) -> str:
    return hash_bytes((json.dumps(value, sort_keys=True, ensure_ascii=False,
                                  separators=(",", ":"), allow_nan=False) + "\n").encode("utf-8"))


def _pinned_comparison(manifest: dict[str, Any]) -> tuple[Path, str]:
    record = manifest.get("comparison")
    if not isinstance(record, dict):
        raise ValueError("seven-stage workflow has no pinned native comparison")
    path, digest = record.get("path"), record.get("sha256")
    if (not isinstance(path, str) or not Path(path).is_absolute()
            or not isinstance(digest, str) or not re.fullmatch(r"[0-9a-f]{64}", digest)
            or not Path(path).is_file() or hash_file(Path(path)) != digest):
        raise ValueError("pinned native comparison changed or is missing")
    return Path(path), digest


def require_parent(root: Path, manifest: dict[str, Any] | None = None) -> tuple[Path, dict[str, Any]]:
    if manifest is None:
        root, manifest = load_workflow(root)
    if workflow_origin(manifest) != "attached":
        raise ValueError("attached seven-stage workflow required")
    parent = manifest.get("parent")
    if not isinstance(parent, dict) or not parent.get("root"):
        raise ValueError("seven-stage workflow has no pinned parent")
    parent_root, old = load_workflow(Path(parent["root"]))
    if (old["schema_version"] != "conversion.workflow.v1"
            or hash_file(parent_root / MANIFEST_NAME) != parent.get("manifest_sha256")):
        raise ValueError("pinned parent workflow changed")
    if old.get("mode") != "production" or manifest.get("mode") != "production":
        raise ValueError("synthetic workflow cannot release a real package")
    if old.get("source", {}).get("fixture") or manifest.get("source", {}).get("fixture"):
        raise ValueError("synthetic fixture cannot release a real package")
    if manifest.get("source") != old.get("source"):
        raise ValueError("pinned source identity changed")
    source = old.get("source", {})
    workbook = Path(source.get("workbook_path", ""))
    if not workbook.is_file() or hash_file(workbook) != source.get("workbook_sha256"):
        raise ValueError("source workbook changed or is missing")
    for stage in range(1, 6):
        state = stage_status(parent_root, old, stage)
        if state["status"] != "approved":
            raise ValueError(f"pinned Stage {stage} is {state['status']}")
        require_stage_approved(parent_root, stage)
    if old.get("input_boundary_required"):
        boundary = input_boundary_status(parent_root, old)
        if (boundary["status"] != "input_boundary_confirmed"
                or boundary.get("human", {}).get("decision") != "approve"):
            raise ValueError("Stage 3 input boundary lacks actual human confirmation")
    _pinned_comparison(manifest)
    return parent_root, old


def require_upstream(root: Path, manifest: dict[str, Any], *,
                     comparison_required: bool = True) -> tuple[Path, dict[str, Any]]:
    if workflow_origin(manifest) == "attached":
        return require_parent(root, manifest)
    source = manifest.get("source", {})
    workbook = Path(source.get("workbook_path", ""))
    if (manifest.get("mode") != "production" or source.get("fixture")
            or not workbook.is_file() or hash_file(workbook) != source.get("workbook_sha256")):
        raise ValueError("native source workbook changed or is synthetic")
    for stage in range(1, 6):
        require_stage_approved(root, stage)
    if manifest.get("input_boundary_required"):
        boundary = input_boundary_status(root, manifest)
        if (boundary["status"] != "input_boundary_confirmed"
                or boundary.get("human", {}).get("decision") != "approve"):
            raise ValueError("Stage 3 input boundary lacks actual human confirmation")
    if comparison_required:
        _, stage5_revision = require_stage_approved(root, 5)
        stage5 = read_json(root / stage5_revision["artifact"]["json"])
        if stage5.get("execution_mode") != "exact_cases":
            _pinned_comparison(manifest)
    return root, manifest


def bind_native_comparison(root: Path, comparison: Path, comparison_sha256: str) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if workflow_origin(manifest) != "native":
        raise ValueError("comparison binding requires a native seven-stage workflow")
    require_upstream(root, manifest, comparison_required=False)
    _, stage4_revision, stage4, bundle = _stage4(root)
    _, stage5_revision = require_stage_approved(root, 5)
    stage4_path = root / stage4_revision["artifact"]["json"]
    stage5 = read_json(root / stage5_revision["artifact"]["json"])
    _require_stage5(stage5, manifest["source"]["workbook_sha256"], hash_file(stage4_path))
    comparison = comparison.expanduser().resolve()
    if not comparison.is_file() or hash_file(comparison) != comparison_sha256:
        raise ValueError("separately verified native comparison digest does not match")
    _require_comparison(read_json(comparison), comparison, manifest["source"]["workbook_sha256"],
                        hash_file(stage4_path), hash_file(bundle / "model_manifest.json"),
                        comparison_sha256, bundle / "model_manifest.json")
    record = {"path": str(comparison), "sha256": comparison_sha256}
    if manifest.get("comparison") not in (None, record):
        raise ValueError("native comparison is already bound to another identity")
    if manifest.get("comparison") is None:
        manifest["comparison"] = record
        _write_manifest(root, manifest)
    return {"status": "pass", "workflow": str(root), "comparison": record}


def parent_stage_status(root: Path, manifest: dict[str, Any], stage: int) -> dict[str, Any]:
    try:
        parent_root, old = require_parent(root, manifest)
        state = stage_status(parent_root, old, stage)
        return {**state, "inherited_from": str(parent_root)}
    except (OSError, ValueError, KeyError, TypeError) as exc:
        return {"stage": stage, "status": "stale", "revision": None,
                "inherited_from": manifest.get("parent", {}).get("root"), "reasons": [str(exc)]}


def create_attached_workflow(root: Path, parent: Path, comparison: Path,
                             comparison_sha256: str) -> dict[str, Any]:
    root = root.expanduser().resolve()
    parent_root, old = load_workflow(parent)
    if old["schema_version"] != "conversion.workflow.v1":
        raise ValueError("parent must be a six-stage workflow")
    if root == parent_root or root.exists() and any(root.iterdir()):
        raise ValueError("seven-stage root must be new and separate from its parent")
    comparison = comparison.expanduser().resolve()
    manifest = {"schema_version": "conversion.workflow.v2", "origin": "attached", "created_at": _now(),
                "mode": "production", "source": old.get("source"),
                "parent": {"root": str(parent_root),
                           "manifest_sha256": hash_file(parent_root / MANIFEST_NAME)},
                "comparison": {"path": str(comparison), "sha256": comparison_sha256},
                "stages": {str(stage): {"current_revision": 0, "revisions": []}
                           for stage in V2_STAGES}, "delivery": None, "delivery_history": []}
    require_parent(root, manifest)
    _write_manifest(root, manifest)
    return {"schema_version": manifest["schema_version"], "workflow": str(root),
            "parent": manifest["parent"], "comparison": manifest["comparison"], "status": "pass"}


def v2_status(root: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    stages = [stage_status(root, manifest, stage) for stage in V2_STAGES]
    stages[5]["meaning"] = "package candidate and internal delivery"
    stages[6]["meaning"] = "final conversion report"
    effective = None
    try:
        effective = require_delivery(root)
    except (OSError, ValueError, KeyError, TypeError):
        pass
    approved = {"approved"}
    if manifest.get("mode") == "synthetic_fixture":
        approved.add("simulated_approved")
    if any(state["status"] not in approved for state in stages[:5]):
        next_stage = (next((state.get("return_to", state["stage"]) for state in stages[:5]
                       if state["status"] not in approved), None)
                      if workflow_origin(manifest) == "native" else None)
    elif stages[5]["status"] == "rejected":
        next_stage = 6
    elif stages[5]["status"] != "approved" or effective is None:
        next_stage = 6
    elif stages[6]["status"] == "rejected":
        next_stage = stages[6]["return_to"]
    elif stages[6]["status"] != "approved":
        next_stage = 7
    else:
        next_stage = None
    result = {"schema_version": "conversion.workflow_status.v2", "workflow": str(root),
              "source": manifest["source"], "mode": manifest["mode"],
              "origin": workflow_origin(manifest), "parent": manifest.get("parent"), "stages": stages,
              "active_delivery": effective, "next_permitted_stage": next_stage}
    if workflow_origin(manifest) == "native":
        from excel_to_act.steps.conversion_workflow import _tool_interface_status

        result["input_boundary"] = input_boundary_status(root, manifest)
        result["tool_interface"] = _tool_interface_status(root, manifest, stages[2])
    return result


def _package_files(bundle: Path, ledger: dict[str, str]) -> dict[str, str]:
    files = _bundle_file_paths(bundle, ledger)
    expected = {Path(name.replace("\\", "/")).as_posix() for name, _path in files}
    directories = {parent.as_posix() for name in expected
                   for parent in Path(name).parents if parent != Path(".")}
    for path in bundle.rglob("*"):
        relative = path.relative_to(bundle).as_posix()
        if (path.is_dir() and relative not in directories
                or not path.is_dir() and (not path.is_file() or relative not in expected)):
            raise ValueError(f"package has undeclared file or directory: {relative}")
    return {name: hash_file(path) for name, path in files}


def _resume_package_copy(source: Path, destination: Path, ledger: dict[str, str]) -> None:
    for name, source_path in _bundle_file_paths(source, ledger):
        target = destination / name
        if target.exists() and not target.is_file():
            raise ValueError("interrupted package copy has a non-file target")
        if not target.exists() or hash_file(target) != ledger[name]:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, target)


def _run_entry(package: Path, request: Path, result: Path,
               ledger: dict[str, str] | None = None) -> dict[str, Any]:
    result.parent.mkdir(parents=True, exist_ok=True)
    environment = os.environ.copy()
    environment.pop("PYTHONPATH", None)
    with tempfile.TemporaryDirectory(prefix="excel-to-act-step6-") as cwd:
        execution_package = Path(cwd) / "package" if ledger is not None else package
        if ledger is not None:
            _resume_package_copy(package, execution_package, ledger)
            if _package_files(execution_package, ledger) != ledger:
                raise ValueError("isolated package copy differs from the approved bundle")
        scratch = Path(cwd) / "model_result.json"
        command = [sys.executable, "-E", "-s", "-B", str(package / "model.py"),
                   "--inputs", str(request), "--out", str(result)]
        executed_command = [sys.executable, "-E", "-s", "-B", str(execution_package / "model.py"),
                            "--inputs", str(request), "--out", str(scratch)]
        completed = subprocess.run(executed_command, cwd=cwd, env=environment,
                                   capture_output=True, text=True, timeout=300, check=False)
        record = {"command": command, "cwd": cwd, "isolated_python": True,
                  "executed_command": executed_command,
                  "package_copy_verified": ledger is not None,
                  "pythonpath_unset": True, "exit_code": completed.returncode,
                  "stdout": completed.stdout, "stderr": completed.stderr,
                  "request_sha256": hash_file(request)}
        if completed.returncode != 0 or not scratch.is_file():
            raise ValueError(f"isolated package entry failed: {record}")
        shutil.copy2(scratch, result)
    record["result_sha256"] = hash_file(result)
    return record


def _comparison_case(comparison: dict[str, Any], request: Path) -> dict[str, Any]:
    request_hash = hash_file(request)
    cases = comparison.get("cases")
    if not isinstance(cases, list) or not all(isinstance(item, dict) for item in cases):
        raise ValueError("native comparison has no valid scenario list")
    if comparison.get("schema_version") == "step5.validation_report.v1":
        matches = [item for item in cases if item.get("request_sha256") == request_hash]
        if len(matches) != 1:
            raise ValueError("native comparison does not bind this exact request scenario")
        case = matches[0]
        result = case.get("model_result")
        if (not isinstance(case.get("case_id"), str) or not isinstance(result, dict)
                or not isinstance(result.get("sha256"), str)):
            raise ValueError("native comparison case result is incomplete")
        return {"case": case["case_id"], "public_result_sha256": result["sha256"]}
    matches = [item for item in cases
               if item.get("request_sha256") == request_hash and item.get("status") == "pass"]
    if len(matches) != 1:
        raise ValueError("native comparison does not bind this exact request scenario")
    return matches[0]


def _candidate_comparison(manifest: dict[str, Any], comparison_path: Path,
                          stage5_path: Path, stage5: dict[str, Any], source_sha: str,
                          stage4_sha: str, bundle_manifest: Path) -> str:
    if workflow_origin(manifest) == "native" and stage5.get("execution_mode") == "exact_cases":
        if (comparison_path.resolve() != stage5_path.resolve()
                or hash_file(comparison_path) != current_revision(manifest, 5)["artifact"]["json_sha256"]):
            raise ValueError("candidate must use the current approved exact-case Step 5 report")
        return hash_file(comparison_path)
    pinned_path, pinned_sha = _pinned_comparison(manifest)
    if comparison_path.resolve() != pinned_path.resolve():
        raise ValueError("candidate uses a different native comparison")
    _require_comparison(read_json(comparison_path), comparison_path, source_sha, stage4_sha,
                        hash_file(bundle_manifest), pinned_sha, bundle_manifest)
    return pinned_sha


def _require_stage5(stage5: dict[str, Any], source_sha: str, stage4_sha: str) -> None:
    if (not isinstance(stage5, dict)
            or stage5.get("schema_version") != "step5.validation_report.v1"
            or stage5.get("status") != "ready_for_review"
            or stage5.get("mismatch_count") != 0
            or stage5.get("source_sha256") != source_sha
            or stage5.get("stage4_artifact_sha256") != stage4_sha
            or stage5.get("formula_cache_inputs") is not False
            or stage5.get("standalone_project_imports") is not False):
        raise ValueError("Step 5 validation is not current for this source and Step 4 report")
    if stage5.get("execution_mode") == "exact_cases":
        bundle_manifest_path = Path(stage5.get("stage4_bundle_manifest_path", ""))
        if not bundle_manifest_path.is_file() or len(bundle_manifest_path.parents) < 5:
            raise ValueError("formal exact-case Step 5 needs its current workflow and package manifest")
        root = bundle_manifest_path.parents[4]
        model_manifest = read_json(bundle_manifest_path)
        layout_path = bundle_manifest_path.parent / "input_layout.json"
        if (not layout_path.is_file()
                or hash_file(layout_path) != model_manifest.get("files", {}).get("input_layout.json")
                or hash_file(bundle_manifest_path) != stage5.get("stage4_bundle_manifest_sha256")):
            raise ValueError("formal exact-case Step 5 package identity changed")
        contract = read_json(layout_path).get("variable_contract")
        if (not isinstance(contract, dict)
                or stage5.get("variable_input_contract_sha256") != model_manifest.get("variable_input_contract_sha256")
                or stage5.get("support_certificate_sha256") != model_manifest.get("support_certificate_sha256")
                or stage5.get("stage3_design_sha256") != contract.get("design", {}).get("artifact", {}).get("json_sha256")
                or stage5.get("tolerances") != contract.get("support_domain", {}).get("tolerances")):
            raise ValueError("formal exact-case Step 5 design, support or tolerance binding changed")
        cases = stage5.get("cases")
        if not isinstance(cases, list) or len(cases) != 32 or stage5.get("case_count") != 32:
            raise ValueError("formal exact-case Step 5 does not contain all 32 cases")
        records, _inputs = _verify_case_receipts(
            root, [Path(item.get("report", {}).get("path", "")) for item in cases],
            contract, source_sha, stage4_sha, hash_file(bundle_manifest_path),
            model_manifest.get("target_names", []),
            {**model_manifest["files"], "model_manifest.json": hash_file(bundle_manifest_path)})
        if records != cases:
            raise ValueError("formal exact-case Step 5 receipt ledger differs from current evidence")
        return
    oracle_record = stage5.get("oracle")
    standalone_record = stage5.get("standalone_validation")
    if not isinstance(oracle_record, dict) or not isinstance(standalone_record, dict):
        raise ValueError("Step 5 oracle or standalone validation evidence is missing")
    for label, record in (("oracle", oracle_record), ("standalone validation", standalone_record)):
        path = Path(record.get("path", ""))
        if not path.is_file() or hash_file(path) != record.get("sha256"):
            raise ValueError(f"Step 5 {label} evidence is missing or changed")
    oracle = read_json(Path(oracle_record["path"]))
    standalone = read_json(Path(standalone_record["path"]))
    if (not isinstance(oracle, dict) or not isinstance(standalone, dict)
            or _has_test_fixture(oracle) or _has_test_fixture(standalone)
            or oracle_record.get("engine") != "Microsoft Excel"
            or oracle_record.get("source_copy_sha256") != source_sha
            or oracle.get("schema_version") != "gp.excel_oracle.v1"
            or oracle.get("status") != "pass"
            or oracle.get("engine") != "Microsoft Excel"
            or oracle.get("source_sha256") != source_sha
            or oracle.get("source_copy_sha256") != source_sha
            or oracle.get("stage4_artifact_sha256") != stage4_sha):
        raise ValueError("Step 5 native Excel oracle does not bind this source and Step 4 report")
    result = Path(standalone.get("result_path", ""))
    if (standalone.get("schema_version") != "step5.execution_validation.v1"
            or standalone.get("status") != "pass"
            or standalone.get("standalone_run_pass") is not True
            or standalone.get("compile_pass") is not True
            or standalone.get("source_sha256") != source_sha
            or standalone.get("stage4_artifact_sha256") != stage4_sha
            or standalone.get("generation_report_sha256") != stage4_sha
            or standalone.get("result_sha256") != standalone_record.get("result_sha256")
            or not result.is_file() or hash_file(result) != standalone["result_sha256"]):
        raise ValueError("Step 5 standalone validation evidence is stale or missing")
    linked_result = read_json(result)
    if not isinstance(linked_result, dict) or _has_test_fixture(linked_result):
        raise ValueError("Step 5 standalone result contains synthetic evidence")


def _require_comparison(comparison: dict[str, Any], comparison_path: Path, source_sha: str,
                        stage4_sha: str, bundle_manifest_sha: str, expected_sha: str,
                        bundle_manifest_path: Path) -> None:
    if not isinstance(comparison, dict) or _has_test_fixture(comparison):
        raise ValueError("native comparison is not an object")
    if hash_file(comparison_path) != expected_sha or read_json(comparison_path) != comparison:
        raise ValueError("native comparison differs from the pinned identity")
    if (not bundle_manifest_path.is_file()
            or hash_file(bundle_manifest_path) != bundle_manifest_sha):
        raise ValueError("native comparison package manifest is missing or changed")
    bundle_manifest = read_json(bundle_manifest_path)
    published = bundle_manifest.get("target_names") if isinstance(bundle_manifest, dict) else None
    if (not isinstance(bundle_manifest, dict)
            or bundle_manifest.get("schema_version") != "step4.modular_model.v1"
            or bundle_manifest.get("source_sha256") != source_sha
            or bundle_manifest.get("formula_cache_inputs") is not False
            or not isinstance(published, list) or not published
            or not all(isinstance(name, str) and name for name in published)
            or len(set(published)) != len(published)):
        raise ValueError("native comparison has no current published output contract")
    layout_path = bundle_manifest_path.parent / "input_layout.json"
    files = bundle_manifest.get("files")
    if (not isinstance(files, dict) or not layout_path.is_file()
            or hash_file(layout_path) != files.get("input_layout.json")):
        raise ValueError("native comparison package input contract is missing or changed")
    layout = read_json(layout_path)
    contract = layout.get("variable_contract") if isinstance(layout, dict) else None
    if (not isinstance(contract, dict) or layout.get("source_sha256") != source_sha
            or contract.get("source_sha256") != source_sha
            or bundle_manifest.get("variable_input_contract_sha256") != _canonical_hash(contract)
            or not isinstance(contract.get("interface_sha256"), str)
            or not re.fullmatch(r"[0-9a-f]{64}", contract["interface_sha256"])
            or not isinstance(contract.get("mutable"), dict) or not contract["mutable"]
            or not isinstance(contract.get("support_domain"), dict)):
        raise ValueError("native comparison package input contract is stale or malformed")
    support_domain_sha = _canonical_hash(contract["support_domain"])
    cases = comparison.get("cases")
    if not isinstance(cases, list) or not all(isinstance(item, dict) for item in cases):
        raise ValueError("native comparison has no valid scenario list")
    case_ids = [item.get("case") for item in cases]
    request_hashes = [item.get("request_sha256") for item in cases]
    native_request_hashes = [item.get("native_request_sha256") for item in cases]
    if (comparison.get("schema_version") != "issue45.public_matrix_comparison.v2"
            or comparison.get("status") != "pass" or comparison.get("mismatch_count") != 0
            or comparison.get("source_sha256") != source_sha
            or comparison.get("stage4_report_sha256") != stage4_sha
            or comparison.get("stage4_bundle_manifest_sha256") != bundle_manifest_sha
            or comparison.get("formula_cache_inputs") is not False
            or comparison.get("public_pass_count") != len(cases)
            or not isinstance(comparison.get("sample"), dict)
            or comparison["sample"].get("count") != len(cases)
            or len(cases) != 32
            or not all(isinstance(value, str) and value for value in case_ids + request_hashes)
            or len(set(case_ids)) != len(cases)
            or len(set(request_hashes)) != len(cases)
            or not all(isinstance(value, str) and value for value in native_request_hashes)
            or len(set(native_request_hashes)) != len(cases)
            or any(item.get("status") != "pass" or item.get("mismatch_count") != 0
                   or not item.get("native_request_sha256") or not item.get("native_result_sha256")
                   for item in cases)):
        raise ValueError("native comparison is not current for this exact source and package")
    tolerance = comparison.get("tolerance")
    if (not isinstance(tolerance, dict)
            or any(isinstance(tolerance.get(key), bool)
                   or not isinstance(tolerance.get(key), (int, float))
                   or not math.isfinite(tolerance[key]) or tolerance[key] < 0
                   for key in ("atol", "rtol"))):
        raise ValueError("native comparison has no valid numerical tolerance")
    sample = comparison["sample"]
    dimensions = (("age", "ages", 4), ("sex", "sex_tokens", 2),
                  ("benefit_term", "benefit_terms", 2),
                  ("payment_period", "payment_periods", 2))
    values = []
    for _, sample_key, size in dimensions:
        options = sample.get(sample_key)
        if (not isinstance(options, list) or len(options) != size
                or any(type(value) not in (str, int, float)
                       or isinstance(value, float) and not math.isfinite(value)
                       for value in options)
                or len({(type(value), value) for value in options}) != size):
            raise ValueError("native comparison has an incomplete sample domain")
        values.append(options)
    expected_axes = set(product(*values))
    declared_axes = []
    for case in cases:
        coordinates = tuple(case.get(field) for field, _, _ in dimensions)
        if any(type(actual) not in (str, int, float)
               or isinstance(actual, float) and not math.isfinite(actual)
               for actual in coordinates):
            raise ValueError("native comparison has malformed sample axes")
        declared_axes.append(coordinates)
    if set(declared_axes) != expected_axes or len(set(declared_axes)) != len(cases):
        raise ValueError("native comparison does not cover the complete sample domain")
    axis_path = comparison_path.parent.parent / "current-request-axis-rebind-manifest.json"
    if not axis_path.is_file() or hash_file(axis_path) != comparison.get("request_axis_manifest_sha256"):
        raise ValueError("native comparison axis manifest is missing or changed")
    axis = read_json(axis_path)
    if (not isinstance(axis, dict) or _has_test_fixture(axis)
            or axis.get("schema_version") != "issue45.current_request_axis_rebind.v1"
            or axis.get("source_sha256") != source_sha
            or axis.get("stage4_report_sha256") != stage4_sha
            or axis.get("bundle_manifest_sha256") != bundle_manifest_sha
            or axis.get("case_count") != len(cases)
            or not isinstance(axis.get("cases"), list) or len(axis["cases"]) != len(cases)):
        raise ValueError("native comparison axis does not bind the source and package")
    by_case = {item.get("case"): item for item in axis["cases"] if isinstance(item, dict)}
    if len(by_case) != len(cases):
        raise ValueError("native comparison axis has duplicate or missing scenarios")
    native_scenarios: set[str] = set()
    axis_fields = [None] * len(dimensions)
    for case in cases:
        bound = by_case.get(case.get("case"))
        if (not isinstance(bound, dict) or bound.get("source_sha256") != source_sha
                or bound.get("formula_cache_inputs") is not False
                or bound.get("public_native_raw_inputs_equal") is not True
                or bound.get("raw_input_mismatch_count") != 0
                or bound.get("external_values_equal_approved_capture_source_mapping") is not True):
            raise ValueError("native comparison scenario has no current source and input binding")
        for key, path_key in (("request_sha256", "request_path"),
                              ("native_request_sha256", "native_request_path"),
                              ("native_result_sha256", "native_result_path"),
                              ("producer_sha256", "producer_path")):
            path = Path(bound.get(path_key, ""))
            if (case.get(key) != bound.get(key) or not path.is_file()
                    or hash_file(path) != case.get(key)):
                raise ValueError(f"native comparison {key} evidence is missing or changed")
        public_result = Path(case.get("execution_workdir", "")) / "model_result.json"
        public_request = read_json(Path(bound["request_path"]))
        native_request = read_json(Path(bound["native_request_path"]))
        native_result = read_json(Path(bound["native_result_path"]))
        producer = read_json(Path(bound["producer_path"]))
        public = (read_json(public_result) if public_result.is_file()
                  and hash_file(public_result) == case.get("public_result_sha256") else None)
        if (not isinstance(native_request, dict) or not isinstance(native_result, dict)
                or not isinstance(producer, dict)
                or not isinstance(public_request, dict)
                or native_request.get("source_sha256") != source_sha
                or native_result.get("source_sha256") != source_sha
                or native_result.get("source_copy_sha256") != source_sha
                or native_result.get("engine") != "Microsoft Excel"
                or native_result.get("scenario_id") != native_request.get("scenario_id")
                or native_result.get("input_overrides") != native_request.get("input_overrides")
                or _has_test_fixture(native_request) or _has_test_fixture(native_result)
                or _has_test_fixture(public_request)
                or not isinstance(public, dict) or _has_test_fixture(public)
                or public.get("schema_version") != "step4.modular_model_result.v1"
                or public.get("status") != "pass"
                or public.get("source_sha256") != source_sha
                or public.get("interface_sha256") != contract["interface_sha256"]
                or public.get("support_domain_sha256") != support_domain_sha):
            raise ValueError("native comparison result or source lineage is invalid")
        scenario = native_request.get("scenario_id")
        if (not isinstance(scenario, str) or not scenario.endswith(case["case"])
                or scenario in native_scenarios):
            raise ValueError("native comparison reuses a native scenario")
        native_scenarios.add(scenario)
        native_input = bound.get("native_input_sha256")
        public_input = bound.get("public_input_sha256")
        inputs = public_request.get("inputs")
        external_data = public_request.get("external_data")
        if (not isinstance(native_input, str) or not re.fullmatch(r"[0-9a-f]{64}", native_input)
                or native_request.get("input_sha256") != native_input
                or native_result.get("input_sha256") != native_input
                or not isinstance(public_input, str) or not re.fullmatch(r"[0-9a-f]{64}", public_input)
                or case.get("public_input_sha256") != public_input
                or public.get("input_sha256") != public_input
                or not isinstance(inputs, dict) or set(inputs) != set(contract["mutable"])
                or not isinstance(external_data, dict)
                or public.get("dependency_sha256") != _canonical_hash(external_data)):
            raise ValueError("native comparison input lineage is invalid")
        raw_inputs = public.get("raw_inputs")
        overrides = native_request.get("input_overrides")
        if (producer.get("case") != case.get("case")
                or producer.get("native_request_sha256") != case.get("native_request_sha256")
                or producer.get("input_sha256") != native_input
                or producer.get("source_sha256") != source_sha
                or producer.get("formula_cache_inputs") is not False
                or not isinstance(raw_inputs, dict) or not raw_inputs
                or not isinstance(overrides, list) or not overrides
                or producer.get("raw_input_count") != len(overrides)):
            raise ValueError("native comparison producer and public case do not bind the native request")
        native_raw = {}
        for item in overrides:
            if (not isinstance(item, dict) or not isinstance(item.get("sheet"), str)
                    or not isinstance(item.get("address"), str)
                    or item.get("qualified_address") != f"{item['sheet']}!{item['address']}"):
                raise ValueError("native comparison raw input binding is malformed")
            address = f"{item['sheet'].casefold()}!{item['address']}"
            if address in native_raw:
                raise ValueError("native comparison repeats a raw input address")
            native_raw[address] = item.get("value")
        if native_raw != raw_inputs:
            raise ValueError("native comparison raw inputs differ from the public case")
        scalar_inputs = {key: value["values"] for key, value in inputs.items()
                         if isinstance(value, dict) and value.get("axes") == []
                         and "values" in value}
        for index, (field, _, _) in enumerate(dimensions):
            declared = case[field]
            matches = set()
            for name, value in scalar_inputs.items():
                entry = contract["mutable"][name]
                bindings = entry.get("bindings") if isinstance(entry, dict) else None
                if (type(value) is type(declared) and value == declared
                        and isinstance(bindings, list)
                        and any(isinstance(binding, dict)
                                and isinstance(binding.get("address"), str)
                                and type(native_raw.get(binding["address"])) is type(declared)
                                and native_raw[binding["address"]] == declared
                                for binding in bindings)):
                    matches.add(name)
            axis_fields[index] = matches if axis_fields[index] is None else axis_fields[index] & matches
        # The delivered model hashes the effective input object, not raw request bytes.
        input_sha = _canonical_hash(inputs)
        if input_sha != public_input:
            raise ValueError("native comparison public result does not bind the request inputs")
        targets = native_request.get("targets")
        named = native_result.get("named_targets")
        public_targets = public.get("targets")
        if (not isinstance(targets, list) or not targets
                or not all(isinstance(name, str) and name for name in targets)
                or len(set(targets)) != len(targets)
                or not isinstance(named, dict) or set(named) != set(targets)
                or not isinstance(public_targets, dict) or not public_targets
                or not set(published).issubset(public_targets)
                or not set(public_targets).issubset(set(targets))):
            raise ValueError("native comparison has missing or malformed targets")
        for name in targets:
            value = named[name].get("value") if isinstance(named[name], dict) else None
            declared = case.get(name)
            if (isinstance(value, bool) or not isinstance(value, (int, float))
                    or not math.isfinite(value) or isinstance(declared, bool)
                    or not isinstance(declared, (int, float)) or not math.isfinite(declared)
                    or not math.isclose(value, declared, rel_tol=tolerance["rtol"],
                                        abs_tol=tolerance["atol"])):
                raise ValueError("native comparison target differs from its declared result")
            if name in public_targets:
                actual = public_targets[name]
                if (isinstance(actual, bool) or not isinstance(actual, (int, float))
                        or not math.isfinite(actual)
                        or not math.isclose(value, actual, rel_tol=tolerance["rtol"],
                                            abs_tol=tolerance["atol"])):
                    raise ValueError("native comparison target differs from the public result")
    if any(not fields for fields in axis_fields):
        raise ValueError("native comparison requests do not exercise their declared sample axes")


def candidate_staleness(root: Path, manifest: dict[str, Any], revision: dict[str, Any]) -> list[str]:
    try:
        candidate = read_json(root / revision["artifact"]["json"])
        if not isinstance(candidate, dict):
            raise ValueError("candidate is not an object")
        if (candidate.get("schema_version") != "step6.delivery_candidate.v1"
                or candidate.get("tool") != "step6.candidate"
                or candidate.get("status") != "ready_for_review"
                or candidate.get("simulation") is True
                or candidate.get("parent") != manifest.get("parent")
                or candidate.get("origin", "attached") != workflow_origin(manifest)
                or candidate.get("permission") != "internal"
                or candidate.get("entry") != "model.py"):
            raise ValueError("candidate schema, parent, permission, or state is invalid")
        version = candidate.get("version")
        if not isinstance(version, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,63}", version):
            raise ValueError("candidate version is invalid")
        package = Path(candidate["package_path"])
        if package.resolve() != (root / "stage6" / "packages" / version).resolve():
            raise ValueError("candidate package path is outside its versioned staging area")
        ledger = candidate["files"]
        if (not isinstance(ledger, dict) or not ledger
                or candidate.get("package_sha256") != hash_bytes(json_bytes(ledger))
                or _package_files(package, ledger) != ledger):
            raise ValueError("candidate package ledger changed")
        inputs = revision.get("input_files", {})
        request = Path(inputs["smoke_request"]["path"])
        smoke_result = Path(inputs["smoke_result"]["path"])
        comparison = Path(inputs["native_comparison"]["path"])
        stage4 = read_json(Path(inputs["stage4_report"]["path"]))
        stage5 = read_json(Path(inputs["stage5_report"]["path"]))
        bundle_manifest = package / "model_manifest.json"
        if not isinstance(stage4, dict) or not isinstance(stage5, dict):
            raise ValueError("approved generation or validation report is not an object")
        if stage5.get("test_fixture"):
            raise ValueError("synthetic Step 5 evidence cannot authorize publication")
        if (stage4.get("schema_version") != "step4.generation_report.v1"
                or stage4.get("bundle", {}).get("files") != ledger
                or candidate.get("runtime_dependencies") != stage4.get("runtime_dependencies")
                or candidate.get("interface_sha256") != stage4.get("variable_input_contract_sha256")
                or candidate.get("interface_sha256") != read_json(bundle_manifest).get("variable_input_contract_sha256")):
            raise ValueError("candidate no longer binds the approved code, interface, dependencies, or validation")
        _require_stage5(stage5, manifest["source"]["workbook_sha256"],
                        hash_file(Path(inputs["stage4_report"]["path"])))
        compared = read_json(comparison)
        if isinstance(compared, dict) and compared.get("test_fixture"):
            raise ValueError("synthetic comparison cannot authorize publication")
        pinned_sha = _candidate_comparison(
            manifest, comparison, Path(inputs["stage5_report"]["path"]), stage5,
            manifest["source"]["workbook_sha256"],
            hash_file(Path(inputs["stage4_report"]["path"])), bundle_manifest)
        if not isinstance(candidate.get("smoke"), dict):
            raise ValueError("candidate smoke record is missing")
        if (candidate.get("request_sha256") != hash_file(request)
                or candidate.get("comparison_sha256") != pinned_sha
                or candidate.get("smoke", {}).get("exit_code") != 0
                or candidate["smoke"].get("request_sha256") != hash_file(request)
                or candidate["smoke"].get("result_sha256") != hash_file(smoke_result)
                or candidate["smoke"].get("result_sha256") != _comparison_case(
                    compared, request)["public_result_sha256"]
                or candidate.get("comparison_case") != _comparison_case(compared, request)["case"]):
            raise ValueError("candidate smoke or scenario binding changed")
        if (candidate.get("stage4_artifact_sha256") != inputs["stage4_report"]["sha256"]
                or candidate.get("stage5_artifact_sha256") != inputs["stage5_report"]["sha256"]):
            raise ValueError("candidate upstream report identity changed")
    except (OSError, ValueError, KeyError, TypeError) as exc:
        return [str(exc)]
    return []


def create_candidate(root: Path, version: str, request: Path, comparison_path: Path,
                     *, permission: str = "internal") -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if manifest["schema_version"] != "conversion.workflow.v2":
        raise ValueError("Step 6 candidate requires a seven-stage workflow")
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,63}", version):
        raise ValueError("version must be a safe, nonempty directory name")
    if permission != "internal":
        raise ValueError("only internal delivery is supported")
    upstream_root, upstream = require_upstream(root, manifest)
    _, stage5_revision = require_stage_approved(upstream_root, 5)
    _, stage4_revision, stage4, bundle = _stage4(upstream_root)
    stage4_path = upstream_root / stage4_revision["artifact"]["json"]
    stage5_path = upstream_root / stage5_revision["artifact"]["json"]
    stage5 = read_json(stage5_path)
    _require_stage5(stage5, upstream["source"]["workbook_sha256"], hash_file(stage4_path))
    if stage5.get("test_fixture"):
        raise ValueError("synthetic Step 5 evidence cannot authorize publication")
    bundle_manifest = bundle / "model_manifest.json"
    comparison_path = comparison_path.expanduser().resolve()
    request = request.expanduser().resolve()
    comparison = read_json(comparison_path)
    if isinstance(comparison, dict) and comparison.get("test_fixture"):
        raise ValueError("synthetic comparison cannot authorize publication")
    _candidate_comparison(
        manifest, comparison_path, stage5_path, stage5,
        upstream["source"]["workbook_sha256"], hash_file(stage4_path), bundle_manifest)
    case = _comparison_case(comparison, request)
    ledger = stage4["bundle"]["files"]
    _bundle_file_paths(bundle, ledger)
    previous = current_revision(manifest, 6)
    package_hash = hash_bytes(json_bytes(ledger))
    if previous is not None:
        previous_payload = read_json(root / previous["artifact"]["json"])
        if (previous_payload.get("version") == version
                and previous_payload.get("package_sha256") == package_hash
                and previous_payload.get("request_sha256") == hash_file(request)
                and previous_payload.get("comparison_sha256") == hash_file(comparison_path)):
            if any(item.get("decision") == "reject" for item in previous.get("decisions", [])):
                raise ValueError("this candidate was rejected; create a new version/candidate")
            if any(item.get("schema_version") == "step6.delivery_revocation.v1"
                   and item.get("version") == version
                   for item in manifest.get("delivery_history", [])):
                raise ValueError("this acceptance was revoked; create a new version/candidate")
            if stage_status(root, manifest, 6)["status"] != "stale":
                return {"status": "pass", "workflow": str(root), "revision": previous["revision"],
                        "artifact": previous["artifact"], "idempotent": True}
    staging = root / "stage6" / "packages" / version
    if any(read_json(root / old_revision["artifact"]["json"]).get("version") == version
           for old_revision in manifest["stages"]["6"]["revisions"]):
        raise ValueError("this version already has a reviewed candidate; use a new version")
    staging.mkdir(parents=True, exist_ok=True)
    _resume_package_copy(bundle, staging, ledger)
    if _package_files(staging, ledger) != ledger:
        raise ValueError("staged package differs from the approved Step 4 bundle")
    smoke_result = root / "stage6" / "smoke" / version / f"{uuid.uuid4().hex}.json"
    smoke = _run_entry(staging, request, smoke_result, ledger)
    if smoke["result_sha256"] != case.get("public_result_sha256"):
        raise ValueError("candidate result differs from the exact scenario comparison")
    payload = {"schema_version": "step6.delivery_candidate.v1", "tool": "step6.candidate",
               "status": "ready_for_review", "version": version, "permission": permission,
               "origin": workflow_origin(manifest), "parent": manifest.get("parent"),
               "stage4_artifact_sha256": hash_file(stage4_path),
               "stage5_artifact_sha256": hash_file(stage5_path),
               "package_path": str(staging), "package_sha256": package_hash,
               "files": ledger, "entry": "model.py", "runtime_dependencies": stage4.get("runtime_dependencies"),
               "interface_sha256": read_json(bundle_manifest).get("variable_input_contract_sha256"),
               "comparison_sha256": hash_file(comparison_path), "request_sha256": hash_file(request),
               "comparison_case": case["case"], "smoke": smoke}
    markdown = (f"# Step 6 internal delivery candidate\n\nVersion: `{version}`; permission: `{permission}`.\n\n"
                f"Package SHA-256: `{package_hash}`. Entry: `model.py --inputs REQUEST.json --out RESULT.json`.\n\n"
                f"Current {workflow_origin(manifest)} upstream: "
                f"`{manifest['parent']['manifest_sha256'] if workflow_origin(manifest) == 'attached' else stage5_revision['artifact']['json_sha256']}`.\n\n"
                f"Exact scenario: `{case['case']}`; request SHA-256: `{hash_file(request)}`; "
                f"result SHA-256: `{smoke['result_sha256']}`.\n\n"
                "Candidate smoke passed. Agent then actual human review of this exact JSON/Markdown pair is required "
                "before internal delivery. This candidate is not an active registration.\n")
    input_files = {"stage4_report": stage4_path, "stage5_report": stage5_path,
                   "native_comparison": comparison_path, "smoke_request": request,
                   "smoke_result": smoke_result,
                   **{f"package:{name}": path for name, path in _bundle_file_paths(staging, ledger)}}
    revision = append_stage_artifact(root, 6, "delivery_candidate.json", payload, markdown,
                                     input_files=input_files,
                                     input_stages={5: stage5_revision["artifact"]})
    return {"status": "pass", "workflow": str(root), "revision": revision["revision"],
            "artifact": revision["artifact"], "package_sha256": package_hash}


def require_delivery(root: Path) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if manifest["schema_version"] != "conversion.workflow.v2":
        raise ValueError("delivery requires a seven-stage workflow")
    require_upstream(root, manifest)
    state = stage_status(root, manifest, 6)
    if state["status"] != "approved":
        raise ValueError(f"Step 6 is {state['status']}, not genuinely accepted")
    revision = current_revision(manifest, 6)
    candidate = read_json(root / revision["artifact"]["json"])
    record = manifest.get("delivery")
    if not isinstance(record, dict):
        raise ValueError("no effective delivery record")
    accepted = state.get("human") or state.get("typesafe")
    if (accepted is None or accepted.get("decision") != "approve"
            or record.get("schema_version") != "step6.delivery_record.v1"
            or record not in manifest.get("delivery_history", [])
            or record.get("receipt_id") != accepted.get("receipt_id")
            or record.get("candidate_artifact") != revision["artifact"]
            or record.get("package_sha256") != candidate.get("package_sha256")
            or record.get("version") != candidate.get("version")
            or record.get("permission") != candidate.get("permission")):
        raise ValueError("delivery does not bind the current accepted candidate")
    if any(item.get("schema_version") == "step6.delivery_revocation.v1"
           and item.get("receipt_id") == record["receipt_id"]
           and item.get("version") == record["version"]
           for item in manifest.get("delivery_history", [])):
        raise ValueError("delivery receipt was revoked")
    package = Path(record["path"])
    if package.resolve() != (root / "deliveries" / candidate["version"]).resolve():
        raise ValueError("delivery is outside the versioned internal registration")
    if _package_files(package, candidate["files"]) != candidate["files"]:
        raise ValueError("delivered package changed")
    readback = record.get("readback")
    if not isinstance(readback, dict) or not isinstance(readback.get("invocation"), dict):
        raise ValueError("delivery readback record is missing")
    request = Path(revision["input_files"]["smoke_request"]["path"])
    result = Path(readback.get("result_path", ""))
    invocation = readback.get("invocation", {})
    if (readback.get("files") != candidate["files"]
            or readback.get("entry") != candidate["entry"]
            or readback.get("result_sha256") != candidate["smoke"]["result_sha256"]
            or readback.get("request_sha256") != candidate["request_sha256"]
            or not isinstance(invocation.get("command"), list)
            or invocation["command"][1:] != ["-E", "-s", "-B", str(package / "model.py"),
                                             "--inputs", str(request), "--out", str(result)]
            or invocation.get("exit_code") != 0
            or invocation.get("pythonpath_unset") is not True
            or invocation.get("isolated_python") is not True
            or invocation.get("request_sha256") != candidate["request_sha256"]
            or invocation.get("result_sha256") != readback.get("result_sha256")):
        raise ValueError("delivery readback does not match the accepted package and scenario")
    if not result.is_file() or hash_file(result) != readback["result_sha256"]:
        raise ValueError("delivered entry result changed or is missing")
    return record


def deliver(root: Path) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if manifest["schema_version"] != "conversion.workflow.v2":
        raise ValueError("delivery requires a seven-stage workflow")
    require_upstream(root, manifest)
    state = stage_status(root, manifest, 6)
    if state["status"] != "approved":
        raise ValueError("Step 6 requires current Agent and actual human acceptance")
    if manifest.get("delivery"):
        record = require_delivery(root)
        return {"status": "pass", "idempotent": True, "delivery": record}
    revision = current_revision(manifest, 6)
    candidate = read_json(root / revision["artifact"]["json"])
    accepted = state.get("human") or state.get("typesafe")
    if any(item.get("schema_version") == "step6.delivery_revocation.v1"
           and item.get("receipt_id") == accepted["receipt_id"]
           for item in manifest.get("delivery_history", [])):
        raise ValueError("this acceptance was revoked; create and review a new candidate")
    staging = Path(candidate["package_path"])
    ledger = candidate["files"]
    if _package_files(staging, ledger) != ledger:
        raise ValueError("accepted candidate package changed")
    final = root / "deliveries" / candidate["version"]
    if final.exists():
        if _package_files(final, ledger) != ledger:
            raise ValueError("existing delivery version differs; never overwrite it")
    else:
        temporary = root / "deliveries" / ".pending" / candidate["version"]
        temporary.mkdir(parents=True, exist_ok=True)
        _resume_package_copy(staging, temporary, ledger)
        if _package_files(temporary, ledger) != ledger:
            raise ValueError("delivery copy failed its file readback")
        os.replace(temporary, final)
    request = Path(next(item["path"] for key, item in revision["input_files"].items()
                        if key == "smoke_request"))
    result = root / "deliveries" / ".readback" / candidate["version"] / f"{uuid.uuid4().hex}.json"
    invocation = _run_entry(final, request, result, ledger)
    if invocation["result_sha256"] != candidate["smoke"]["result_sha256"]:
        raise ValueError("delivered entry readback differs from the accepted candidate")
    readback = {"files": _package_files(final, ledger), "entry": "model.py",
                "request_sha256": invocation["request_sha256"],
                "result_path": str(result), "result_sha256": invocation["result_sha256"],
                "invocation": invocation}
    record = {"schema_version": "step6.delivery_record.v1", "recorded_at": _now(),
              "receipt_id": accepted["receipt_id"], "candidate_artifact": revision["artifact"],
              "package_sha256": candidate["package_sha256"], "version": candidate["version"],
              "permission": candidate["permission"], "path": str(final), "readback": readback}
    manifest["delivery"] = record
    manifest.setdefault("delivery_history", []).append(record)
    _write_manifest(root, manifest)
    return {"status": "pass", "idempotent": False, "delivery": record}


def revoke_delivery(root: Path, reason: str) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if manifest["schema_version"] != "conversion.workflow.v2" or not reason.strip():
        raise ValueError("a seven-stage delivery and a reason are required")
    if manifest.get("delivery") is None:
        raise ValueError("there is no effective delivery to revoke")
    manifest.setdefault("delivery_history", []).append(
        {"schema_version": "step6.delivery_revocation.v1", "recorded_at": _now(),
         "version": manifest["delivery"]["version"],
         "receipt_id": manifest["delivery"]["receipt_id"], "reason": reason.strip()})
    manifest["delivery"] = None
    _write_manifest(root, manifest)
    return {"status": "pass", "active_delivery": None}
