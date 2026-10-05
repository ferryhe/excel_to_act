import hashlib
import json
from pathlib import Path

import pytest

from excel_to_act.steps.step1.workflow import _write_handoff
from excel_to_act.steps.step2.workflow import build_index, tool_catalog, validate_saved_index


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value), encoding="utf-8")


def source_handoff(**updates: object) -> dict:
    handoff = {
        "schema_version": "step1.v1",
        "run_path": "runs/source-a",
        "artifact_paths_relative_to": "runs/source-a/artifacts",
        "source": {"path": "models/a.xlsx", "sha256": "source-hash"},
        "run_id": "run-a",
        "batch_id": "batch-a",
        "status": "pass",
        "ready_for_next_step": True,
        "metrics_state": "current",
        "metrics": {"traceability_ratio": 1.0},
        "thresholds": {"traceability": 1.0},
        "artifacts": [{"name": "inventory", "path": "inventory.json", "sha256": "artifact-hash"}],
        "final_output": "final/batches/batch-a/source-a",
        "diagnostics": [],
        "blockers": [],
        "next_actions": [],
    }
    handoff.update(updates)
    return handoff


def source_metadata() -> dict:
    return {
        "schema_version": "step1.v1",
        "batch_id": "batch-a",
        "run_id": "run-a",
        "source_path": "C:/models/a.xlsx",
        "relative_path": "models/a.xlsx",
        "sha256": None,
        "suffix": ".xlsx",
        "thresholds": {},
        "allow_opaque": True,
        "attempt_limit": 3,
    }


def test_builds_single_current_handoff_with_relative_artifact_refs_and_quality(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    input_path = root / "batches/batch-a/sources/source-a/handoff.json"
    inventory_path = root / "runs/source-a/artifacts/inventory.json"
    inventory = {
        "schema_version": "phase1.v1", "artifact_type": "workbook_inventory",
        "workbook_sha256": "source-hash", "sheets": [],
        "coverage": {"recognized_inventory_objects": 0, "unsupported_or_opaque_objects": 0, "discovered_workbook_objects": 0},
    }
    write_json(inventory_path, inventory)
    write_json(input_path, source_handoff(artifacts=[{
        "name": "inventory", "path": "inventory.json",
        "sha256": hashlib.sha256(inventory_path.read_bytes()).hexdigest(),
    }]))

    result = build_index(input_path, root, tmp_path / "index")
    payload = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))

    assert result["status"] == "partial"
    assert len(payload["entries"]) == 1
    entry = payload["entries"][0]
    assert entry["source_path"] == "models/a.xlsx"
    assert entry["source_sha256"] == "source-hash"
    assert entry["status"] == "pass"
    assert entry["handoff_path"] == "batches/batch-a/sources/source-a/handoff.json"
    assert entry["artifacts"][0]["path"] == "runs/source-a/artifacts/inventory.json"
    assert entry["metrics"] == {"traceability_ratio": 1.0}
    assert entry["thresholds"] == {"traceability": 1.0}
    assert entry["final_output"] == "final/batches/batch-a/source-a"
    assert "large content" not in (tmp_path / "index/index.json").read_text(encoding="utf-8")


def test_accepts_legitimate_failed_handoff_with_missing_hash_and_no_artifacts(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = root / "failed/handoff.json"
    write_json(handoff, {
        "schema_version": "step1.v1",
        "source": {"path": None, "sha256": None},
        "run_id": "failed-run",
        "status": "fail",
        "ready_for_next_step": False,
        "metrics_state": "unavailable",
        "metrics": {},
        "artifacts": [],
    })

    result = build_index(handoff, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["status"] == "fail"
    assert entry["source_sha256"] is None
    assert entry["artifacts"] == []
    assert entry["diagnostics"] == []


def test_rejects_source_metadata_shape_as_root_handoff(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = root / "run/source.json"
    write_json(handoff, source_metadata())

    result = build_index(handoff, root, tmp_path / "index")

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["code"] == "handoff_invalid"
    assert "source" in result["diagnostics"][0]["message"]


def test_rejects_source_metadata_shape_in_batch_and_preserves_batch_evidence(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    metadata_path = root / "run/source.json"
    batch_path = root / "batch_handoff.json"
    write_json(metadata_path, source_metadata())
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{
            "relative_path": "models/a.xlsx",
            "sha256": None,
            "run_id": "run-a",
            "status": "fail",
            "metrics_state": "unavailable",
            "metrics": {"preserved": True},
            "diagnostics": [{"code": "source_read_failed"}],
            "handoff_path": "run/source.json",
        }],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["source_path"] == "models/a.xlsx"
    assert entry["run_id"] == "run-a"
    assert entry["status"] == "fail"
    assert entry["metrics"] == {"preserved": True}
    assert entry["diagnostics"][0]["code"] == "source_read_failed"
    assert entry["diagnostics"][1]["code"] == "source_handoff_invalid"


def test_batch_keeps_failed_sources_and_missing_handoffs_with_evidence(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batches/batch-b/batch_handoff.json"
    write_json(root / "sources/pass/handoff.json", source_handoff(
        source={"path": "models/a.xlsx", "sha256": "sha-a"}, status="partial", run_id="run-pass", artifacts=[]
    ))
    write_json(root / "sources/fail/handoff.json", source_handoff(
        source={"path": "models/b.xlsx", "sha256": "sha-b"},
        status="partial",
        ready_for_next_step=False,
        run_id="run-fail",
        metrics={"opaque_rate": 0.2},
        artifacts=[],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "batch_id": "batch-b",
        "entries": [
            {"relative_path": "models/a.xlsx", "sha256": "sha-a", "run_id": "run-pass", "status": "partial", "handoff_path": "sources/pass/handoff.json", "metrics": {"fallback": 1}},
            {"relative_path": "models/b.xlsx", "sha256": "sha-b", "run_id": "run-fail", "status": "fail", "handoff_path": "sources/fail/handoff.json", "diagnostics": [{"code": "failed"}]},
            {"relative_path": "models/c.xlsx", "sha256": "sha-c", "run_id": "run-missing", "status": "fail", "metrics": {"retained": 3}, "diagnostics": [{"code": "source_error"}]},
        ],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entries = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"]

    assert result["status"] == "partial"
    assert [entry["status"] for entry in entries] == ["partial", "fail", "fail"]
    assert [entry["source_path"] for entry in entries] == ["models/a.xlsx", "models/b.xlsx", "models/c.xlsx"]
    assert [entry["source_sha256"] for entry in entries] == ["sha-a", "sha-b", "sha-c"]
    assert [entry["run_id"] for entry in entries] == ["run-pass", "run-fail", "run-missing"]
    assert all(entry["source_id"] for entry in entries)
    assert entries[1]["diagnostics"][0]["code"] == "failed"
    assert entries[2]["metrics"] == {"retained": 3}
    assert entries[2]["diagnostics"][0]["code"] == "source_error"
    assert any(item["code"] == "source_handoff_unavailable" for item in entries[2]["diagnostics"])
    assert entries[2]["handoff_path"] is None


@pytest.mark.parametrize(
    ("handoff_update", "mismatch"),
    [
        ({"source": {"path": "models/other.xlsx", "sha256": "sha-b"}}, "source path"),
        ({"source": {"path": "models/b.xlsx", "sha256": "other-sha"}}, "source SHA-256"),
        ({"run_id": "other-run"}, "run ID"),
    ],
)
def test_batch_rejects_handoff_with_mismatched_identity_and_keeps_batch_evidence(
    tmp_path: Path,
    handoff_update: dict,
    mismatch: str,
) -> None:
    root = tmp_path / "step1"
    updates = {"source": {"path": "models/b.xlsx", "sha256": "sha-b"}, "run_id": "run-b"}
    updates.update(handoff_update)
    handoff = source_handoff(**updates, metrics={"wrong": True})
    handoff_path = root / "sources/other/handoff.json"
    batch_path = root / "batch.json"
    write_json(handoff_path, handoff)
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "batch_id": "batch-a",
        "entries": [{
            "relative_path": "models/b.xlsx",
            "sha256": "sha-b",
            "run_id": "run-b",
            "status": "fail",
            "metrics": {"batch_evidence": True},
            "diagnostics": [{"code": "batch_diagnostic"}],
            "handoff_path": "sources/other/handoff.json",
        }],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["source_path"] == "models/b.xlsx"
    assert entry["source_sha256"] == "sha-b"
    assert entry["run_id"] == "run-b"
    assert entry["status"] == "fail"
    assert entry["metrics"] == {"batch_evidence": True}
    assert entry["artifacts"] == []
    assert entry["diagnostics"][0]["code"] == "batch_diagnostic"
    assert entry["diagnostics"][1]["code"] == "source_handoff_mismatch"
    assert mismatch in entry["diagnostics"][1]["message"]


def test_batch_identity_check_allows_missing_source_hash_and_run_id(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batch.json"
    write_json(root / "sources/a/handoff.json", source_handoff(
        source={"path": "models/a.xlsx", "sha256": None},
        run_id=None,
        status="fail",
        ready_for_next_step=False,
        artifacts=[],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{
            "relative_path": "models/a.xlsx",
            "sha256": "batch-sha",
            "run_id": "batch-run",
            "status": "fail",
            "handoff_path": "sources/a/handoff.json",
        }],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["source_sha256"] == "batch-sha"
    assert entry["run_id"] == "batch-run"
    assert entry["diagnostics"] == []


def test_batch_handoff_and_artifact_paths_use_layered_bases_not_final_output(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batch_handoff.json"
    write_json(root / "per-source/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"},
        run_id="1",
        run_path="runs/alternate",
        artifact_paths_relative_to="runs/chosen",
        final_output="final/elsewhere",
        artifacts=[{"name": "facts", "path": "facts.json", "sha256": "facts-hash"}],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{"relative_path": "a.xlsx", "sha256": "a", "run_id": "1", "status": "pass", "handoff_path": "per-source/handoff.json"}],
    })

    build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert entry["handoff_path"] == "per-source/handoff.json"
    assert entry["artifacts"][0]["path"] == "runs/chosen/facts.json"
    assert entry["run_path"] == "runs/alternate"
    assert entry["final_output"] == "final/elsewhere"


def test_artifacts_fall_back_to_run_path_when_relative_base_is_absent(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = source_handoff(artifact_paths_relative_to=None, run_path="runs/fallback")
    input_path = root / "handoff.json"
    write_json(input_path, handoff)

    build_index(input_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert entry["artifacts"][0]["path"] == "runs/fallback/inventory.json"


def test_batch_path_outside_root_is_not_retained_as_a_handoff_reference(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batch.json"
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{"relative_path": "a.xlsx", "sha256": "a", "status": "fail", "handoff_path": "../../outside.json"}],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["handoff_path"] is None
    assert any(item["code"] == "source_handoff_unavailable" for item in entry["diagnostics"])


def test_rejects_outside_root_and_malformed_root_handoffs(tmp_path: Path, monkeypatch) -> None:
    root = tmp_path / "step1"
    root.mkdir()
    monkeypatch.chdir(tmp_path)
    outside = tmp_path / "outside.json"
    write_json(outside, source_handoff())
    write_json(root / "outside.json", source_handoff())

    outside_result = build_index(Path("outside.json"), root, tmp_path / "outside-index")
    malformed_path = root / "bad.json"
    malformed_path.write_text("[]", encoding="utf-8")
    malformed_result = build_index(malformed_path, root, tmp_path / "bad-index")

    assert outside_result["status"] == "blocked"
    assert outside_result["diagnostics"][0]["code"] == "handoff_invalid"
    assert malformed_result["status"] == "blocked"
    assert malformed_result["diagnostics"][0]["code"] == "handoff_invalid"


def test_rejects_legacy_schema_and_keeps_escaping_artifact_out_of_refs(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    legacy_path = root / "legacy.json"
    write_json(legacy_path, {"schema_version": "phase1.v1", "artifacts": []})
    result = build_index(legacy_path, root, tmp_path / "legacy-index")

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["code"] == "handoff_invalid"

    handoff_path = root / "handoff.json"
    write_json(handoff_path, source_handoff(artifacts=[{"name": "bad", "path": "../../../../outside.json", "sha256": "x"}]))
    build_index(handoff_path, root, tmp_path / "artifact-index")
    entry = json.loads((tmp_path / "artifact-index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert entry["artifacts"] == []
    assert any(item["code"] == "artifact_path_invalid" for item in entry["diagnostics"])


def test_unchanged_inputs_produce_identical_index_files(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff())
    first = tmp_path / "first"
    second = tmp_path / "second"

    build_index(handoff, root, first)
    build_index(handoff, root, second)

    assert (first / "index.json").read_bytes() == (second / "index.json").read_bytes()
    assert (first / "INDEX.md").read_bytes() == (second / "INDEX.md").read_bytes()


def test_tools_catalogue_documents_resolver_and_builder_contracts() -> None:
    tools = tool_catalog()["tools"]

    assert [tool["name"] for tool in tools] == ["step2.handoff.resolve", "step2.index.build", "step2.index.validate"]
    for tool in tools:
        assert tool["command"]
        assert tool["inputs"]
        assert tool["outputs"]
        assert "next" in tool


def test_saved_index_validator_checks_paths_hashes_and_known_schemas(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    entry = index["entries"][0]
    graph = root / "runs/source-a/dependency_graph.json"
    write_json(graph, {"artifact_type": "wrong", "source_location": {"workbook_path": "a.xlsx", "object_type": "cell"}})
    outside = tmp_path / "outside.json"
    write_json(outside, {"opaque": True})
    opaque = root / "runs/source-a/unknown.json"
    write_json(opaque, {"opaque": True})
    changed_graph = root / "runs/source-a/changed_graph.json"
    write_json(changed_graph, {"nodes": [{"source_location": {"workbook_path": "a.xlsx", "sheet_name": "Sheet1", "address": "A1", "object_type": "cell"}}]})
    malformed = root / "runs/source-a/malformed.json"
    malformed.parent.mkdir(parents=True, exist_ok=True)
    malformed.write_text("{", encoding="utf-8")
    entry["artifacts"] = [
        {"name": "dependency_graph.json", "path": graph.relative_to(root).as_posix(), "sha256": hashlib.sha256(graph.read_bytes()).hexdigest()},
        {"name": "inventory.json", "path": "runs/source-a/missing.json", "sha256": "missing"},
        {"name": "outside.json", "path": "../outside.json", "sha256": hashlib.sha256(outside.read_bytes()).hexdigest()},
        {"name": "unknown.json", "path": opaque.relative_to(root).as_posix(), "sha256": hashlib.sha256(opaque.read_bytes()).hexdigest()},
        {"name": "inventory.json", "path": malformed.relative_to(root).as_posix(), "sha256": hashlib.sha256(malformed.read_bytes()).hexdigest()},
        {"name": "dependency_graph.json", "path": changed_graph.relative_to(root).as_posix(), "sha256": "old-checksum"},
        {"name": "inventory.json", "path": malformed.relative_to(root).as_posix(), "sha256": "old-checksum"},
    ]
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    assert {item["code"] for item in result["diagnostics"]} == {
        "artifact_path_invalid", "artifact_missing", "artifact_schema_invalid", "artifact_json_invalid", "artifact_checksum_mismatch"
    }
    assert all(item["source_id"] == entry["source_id"] for item in result["diagnostics"])
    assert all(item["artifact"] for item in result["diagnostics"])
    schema_error = next(item for item in result["diagnostics"] if item["code"] == "artifact_schema_invalid")
    assert schema_error["source_location"]["workbook_path"] == "a.xlsx"
    changed_error = next(item for item in result["diagnostics"] if item["code"] == "artifact_checksum_mismatch" and item["artifact"] == "dependency_graph.json")
    assert changed_error["source_location"]["address"] == "A1"
    malformed_error = next(item for item in result["diagnostics"] if item["code"] == "artifact_checksum_mismatch" and item["artifact"] == "inventory.json")
    assert "source_location" not in malformed_error
    assert result["next_tool"] is None


def test_saved_index_validator_checks_current_step1_identity_fields(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    source_record = source_metadata()
    source_record.update({"relative_path": "models/other.xlsx", "run_id": "run-b", "sha256": "another-workbook"})
    artifacts = {
        "workbook_manifest.json": {
            "schema_version": "phase1.v1", "artifact_type": "workbook_manifest",
            "workbook_path": "a.xlsx", "file_name": "a.xlsx", "file_size": 0,
            "sha256": "another-workbook", "sheets": [],
        },
        "source.json": source_record,
        "quality.json": {
            "schema_version": "step1.v1", "source_sha256": "another-workbook", "run_id": "run-a",
        },
        "custom.json": {"sha256": "opaque-hash"},
    }
    refs = []
    for name, payload in artifacts.items():
        path = root / "runs/source-a" / name
        write_json(path, payload)
        refs.append({"name": name, "path": path.relative_to(root).as_posix(), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    index["entries"][0]["artifacts"] = refs
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    mismatches = [item for item in result["diagnostics"] if item["code"] == "artifact_identity_mismatch"]
    assert {(item["artifact"], "source path" in item["message"]) for item in mismatches} == {
        ("workbook_manifest.json", False), ("source.json", True),
        ("source.json", False), ("source.json", False), ("quality.json", False),
    }
    assert not any(item["artifact"] == "custom.json" for item in mismatches)


@pytest.mark.parametrize(
    ("name", "payload"),
    [
        ("workbook_manifest.json", {
            "schema_version": "phase1.v1", "artifact_type": "workbook_manifest",
            "workbook_path": "a.xlsx", "file_name": "a.xlsx", "file_size": 0,
            "sha256": "wrong-source", "sheets": [],
        }),
        ("source.json", {"relative_path": "models/other.xlsx", "sha256": "wrong-source"}),
        ("quality.json", {"source_sha256": "wrong-source", "run_id": "run-a"}),
    ],
)
def test_build_index_blocks_mismatched_current_step1_artifacts(tmp_path: Path, name: str, payload: dict) -> None:
    root = tmp_path / "step1"
    handoff = root / "handoff.json"
    if name == "source.json":
        payload = source_metadata()
        payload.update({"relative_path": "models/other.xlsx", "sha256": "wrong-source"})
    artifact = root / "runs/source-a/artifacts" / name
    write_json(artifact, payload)
    write_json(handoff, source_handoff(artifacts=[{
        "name": name, "path": name,
        "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]))

    result = build_index(handoff, root, tmp_path / "index")

    assert result["status"] == "blocked"
    assert any(item["code"] == "artifact_identity_mismatch" and item["artifact"] == name for item in result["diagnostics"])
    assert json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["validation_status"] == "blocked"


def test_saved_index_validator_accepts_legacy_phase1_handoff(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    artifact = root / "runs/source-a/handoff.json"
    write_json(artifact, {
        "schema_version": "phase1.v1", "artifact_type": "handoff",
        "workbook_sha256": "source-hash", "run_id": "run-a",
    })
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    index["entries"][0]["artifacts"] = [{
        "name": "handoff.json", "path": artifact.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]
    write_json(index_path, index)

    assert validate_saved_index(index_path, root)["status"] == "pass"


def test_saved_index_validator_passes_known_and_opaque_artifacts(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    artifacts = [
        ("dependency_graph.json", {"schema_version": "phase1.v1", "artifact_type": "formula_graph", "nodes": [], "edges": []}),
        ("custom.json", {"custom": "opaque"}),
    ]
    refs = []
    for name, payload in artifacts:
        path = root / "runs/source-a" / name
        write_json(path, payload)
        refs.append({"name": name, "path": path.relative_to(root).as_posix(), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    index["entries"][0]["artifacts"] = refs
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "pass"
    assert result["diagnostics"] == []
    assert result["metrics"]["artifacts_checked"] == 2
    assert result["retryable"] is False


def test_saved_index_validator_accepts_and_repairs_writer_emitted_current_handoff(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    run_dir = root / "batches/batch-a/sources/source-a/runs/run-a"
    source = {
        "schema_version": "step1.v1", "batch_id": "batch-a", "run_id": "run-a",
        "source_path": "C:/models/a.xlsx", "relative_path": "models/a.xlsx",
        "sha256": "source-hash", "suffix": ".xlsx", "thresholds": {},
        "allow_opaque": True, "attempt_limit": 3,
    }
    write_json(run_dir / "source.json", source)
    write_json(run_dir / "quality.json", {
        "source_sha256": "source-hash", "run_id": "run-a", "status": "pass",
        "ready_for_next_step": True, "metrics_state": "current", "metrics": {},
    })
    quality = json.loads((run_dir / "quality.json").read_text(encoding="utf-8"))
    _write_handoff(run_dir, quality)
    handoff_path = run_dir / "handoff.json"
    build_index(handoff_path, root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    artifact = handoff_path
    index["entries"][0]["artifacts"] = [{
        "name": "handoff.json", "path": artifact.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]
    write_json(index_path, index)

    valid = validate_saved_index(index_path, root)
    assert valid["status"] == "pass"

    emitted = json.loads(artifact.read_text(encoding="utf-8"))
    emitted.pop("source")
    write_json(artifact, emitted)
    index["entries"][0]["artifacts"][0]["sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    index["input_handoff_sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    write_json(index_path, index)
    invalid = validate_saved_index(index_path, root)

    assert invalid["status"] == "blocked"
    assert any(item["code"] == "artifact_schema_invalid" and item["artifact"] == "handoff.json" for item in invalid["diagnostics"])
    artifact_error = next(item for item in invalid["diagnostics"] if item["code"] == "artifact_schema_invalid")
    assert artifact_error["next_tool"] == "report.handoff"
    assert invalid["next_tool"] is None

    _write_handoff(run_dir, quality)
    index["entries"][0]["artifacts"][0]["sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    index["input_handoff_sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    write_json(index_path, index)
    assert validate_saved_index(index_path, root)["status"] == "pass"


def test_saved_index_validator_checks_input_handoff_reference(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, index_dir)
    handoff.write_text("{}", encoding="utf-8")

    result = validate_saved_index(index_dir / "index.json", root)

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["code"] == "input_handoff_checksum_mismatch"


def test_saved_index_validator_checks_input_handoff_schema_even_with_updated_digest(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    write_json(handoff, {})
    index["input_handoff_sha256"] = hashlib.sha256(handoff.read_bytes()).hexdigest()
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    assert any(item["code"] == "input_handoff_schema_invalid" and item["artifact"] == "input_handoff" for item in result["diagnostics"])


@pytest.mark.parametrize("invalid_handoff_path", ["../outside.json", "sources/pass/missing.json"])
def test_saved_index_validator_rechecks_batch_entry_handoff_path(tmp_path: Path, invalid_handoff_path: str) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    batch_path = root / "batch_handoff.json"
    per_source_handoff = root / "sources/pass/handoff.json"
    write_json(per_source_handoff, source_handoff(
        source={"path": "models/a.xlsx", "sha256": "sha-a"},
        run_id="run-a",
        artifacts=[],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{
            "relative_path": "models/a.xlsx",
            "sha256": "sha-a",
            "run_id": "run-a",
            "status": "pass",
            "handoff_path": "sources/pass/handoff.json",
        }],
    })
    build_index(batch_path, root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    assert validate_saved_index(index_path, root)["status"] == "pass"

    index["entries"][0]["handoff_path"] = invalid_handoff_path
    write_json(index_path, index)
    if invalid_handoff_path == "sources/pass/missing.json":
        per_source_handoff.unlink()

    result = validate_saved_index(index_path, root)

    expected_code = "source_handoff_path_invalid" if invalid_handoff_path.startswith("..") else "source_handoff_missing"
    diagnostic = next(item for item in result["diagnostics"] if item["code"] == expected_code)
    assert result["status"] == "blocked"
    assert diagnostic["artifact"] == "source_handoff"
    assert diagnostic["source_path"] == "models/a.xlsx"


def test_saved_index_validator_rejects_incomplete_source_record_artifact(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    source_path = root / "runs/source-a/source.json"
    source_record = source_metadata()
    source_record.pop("attempt_limit")
    write_json(source_path, source_record)
    index["entries"][0]["artifacts"] = [{
        "name": "source.json",
        "path": source_path.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(source_path.read_bytes()).hexdigest(),
    }]
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    assert any(item["code"] == "artifact_schema_invalid" and item["artifact"] == "source.json" for item in result["diagnostics"])
