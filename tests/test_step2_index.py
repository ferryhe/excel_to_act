import json
from pathlib import Path

import pytest

from excel_to_act.steps.step2.workflow import build_index, tool_catalog


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
    write_json(input_path, source_handoff())
    (root / "runs/source-a/artifacts/inventory.json").parent.mkdir(parents=True)
    (root / "runs/source-a/artifacts/inventory.json").write_text("large content stays outside the index", encoding="utf-8")

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
        source={"path": "models/a.xlsx", "sha256": "sha-a"}, status="partial", run_id="run-pass"
    ))
    write_json(root / "sources/fail/handoff.json", source_handoff(
        source={"path": "models/b.xlsx", "sha256": "sha-b"},
        status="partial",
        ready_for_next_step=False,
        run_id="run-fail",
        metrics={"opaque_rate": 0.2},
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

    assert [tool["name"] for tool in tools] == ["step2.handoff.resolve", "step2.index.build"]
    for tool in tools:
        assert tool["command"]
        assert tool["inputs"]
        assert tool["outputs"]
        assert "next" in tool
