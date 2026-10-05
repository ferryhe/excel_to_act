import hashlib
import json
from pathlib import Path

import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app


def test_step2_tools_lists_initial_actions() -> None:
    result = CliRunner().invoke(app, ["step2", "tools"])

    assert result.exit_code == 0
    assert [item["name"] for item in json.loads(result.stdout)["tools"]] == [
        "step2.handoff.resolve",
        "step2.index.build",
        "step2.index.validate",
    ]


def test_step2_index_exits_successfully_for_usable_partial_index(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    root.mkdir()
    handoff = root / "handoff.json"
    handoff.write_text(json.dumps({
        "schema_version": "step1.v1",
        "source": {"path": "a.xlsx", "sha256": "sha"},
        "status": "partial",
        "ready_for_next_step": False,
        "metrics_state": "unavailable",
        "metrics": {},
        "artifacts": [],
        "next_actions": [{"name": "source.review"}],
    }), encoding="utf-8")

    result = CliRunner().invoke(app, ["step2", "index", "--handoff", str(handoff), "--step1-root", str(root), "--out", str(tmp_path / "index")])

    assert result.exit_code == 0
    assert json.loads(result.stdout)["status"] == "partial"
    assert (tmp_path / "index/index.json").exists()


def test_step2_validate_passes_and_index_blocks_corrupt_references(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    root.mkdir()
    handoff = root / "handoff.json"
    handoff.write_text(json.dumps({
        "schema_version": "step1.v1", "source": {"path": "a.xlsx", "sha256": "sha"},
        "status": "pass", "artifacts": [],
    }), encoding="utf-8")
    index_dir = tmp_path / "index"
    built = CliRunner().invoke(app, ["step2", "index", "--handoff", str(handoff), "--step1-root", str(root), "--out", str(index_dir)])
    assert built.exit_code == 0
    valid = CliRunner().invoke(app, ["step2", "validate", "--index", str(index_dir / "index.json"), "--step1-root", str(root)])
    assert valid.exit_code == 0

    index = json.loads((index_dir / "index.json").read_text(encoding="utf-8"))
    index["entries"][0]["artifacts"] = [{"name": "inventory.json", "path": "missing.json", "sha256": "x"}]
    (index_dir / "index.json").write_text(json.dumps(index), encoding="utf-8")
    invalid = CliRunner().invoke(app, ["step2", "validate", "--index", str(index_dir / "index.json"), "--step1-root", str(root)])

    assert invalid.exit_code == 1
    result = json.loads(invalid.stdout)
    assert result["status"] == "blocked"
    assert {"status", "diagnostics", "metrics", "retryable", "next_tool"} <= result.keys()


def test_step2_index_blocks_source_metadata_run_id_mismatch(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    source_path = root / "runs/run-a/source.json"
    source_path.parent.mkdir(parents=True)
    source_path.write_text(json.dumps({
        "schema_version": "phase1.v1", "batch_id": "batch-a", "run_id": "run-b",
        "source_path": "C:/models/a.xlsx", "relative_path": "models/a.xlsx",
        "sha256": "source-hash", "suffix": ".xlsx", "thresholds": {},
        "allow_opaque": True, "attempt_limit": 3,
    }), encoding="utf-8")
    handoff = root / "handoff.json"
    handoff.write_text(json.dumps({
        "schema_version": "step1.v1", "source": {"path": "models/a.xlsx", "sha256": "source-hash"},
        "run_id": "run-a", "status": "pass", "run_path": "runs/run-a",
        "artifacts": [{
            "name": "source.json", "path": "source.json",
            "sha256": hashlib.sha256(source_path.read_bytes()).hexdigest(),
        }],
    }), encoding="utf-8")
    index_dir = tmp_path / "index"

    built = CliRunner().invoke(app, ["step2", "index", "--handoff", str(handoff), "--step1-root", str(root), "--out", str(index_dir)])

    assert built.exit_code == 1
    build_result = json.loads(built.stdout)
    mismatch = next(item for item in build_result["diagnostics"] if item["code"] == "artifact_identity_mismatch")
    assert mismatch["artifact"] == "source.json"
    assert "run ID" in mismatch["message"]
    assert json.loads((index_dir / "index.json").read_text(encoding="utf-8"))["validation_status"] == "blocked"
    checked = CliRunner().invoke(app, ["step2", "validate", "--index", str(index_dir / "index.json"), "--step1-root", str(root)])
    assert checked.exit_code == 1
    assert json.loads(checked.stdout)["status"] == "blocked"


@pytest.mark.parametrize(
    ("artifact_ref", "expected_code", "expected_path"),
    [
        ({"name": "inventory.json", "path": "../../../outside.json", "sha256": "x"}, "artifact_path_invalid", "../../../outside.json"),
        ({"name": "inventory.json", "path": "inventory.json"}, "artifact_ref_incomplete", "inventory.json"),
    ],
)
def test_step2_index_blocks_rejected_artifact_refs_through_saved_recheck(
    tmp_path: Path, artifact_ref: dict, expected_code: str, expected_path: str,
) -> None:
    root = tmp_path / "step1"
    root.mkdir()
    handoff = root / "handoff.json"
    handoff.write_text(json.dumps({
        "schema_version": "step1.v1", "source": {"path": "models/a.xlsx", "sha256": "source-hash"},
        "run_id": "run-a", "status": "pass", "run_path": "runs/run-a",
        "artifact_paths_relative_to": "runs/run-a", "artifacts": [artifact_ref],
    }), encoding="utf-8")
    index_dir = tmp_path / "index"

    built = CliRunner().invoke(app, ["step2", "index", "--handoff", str(handoff), "--step1-root", str(root), "--out", str(index_dir)])

    assert built.exit_code == 1
    result = json.loads(built.stdout)
    diagnostic = next(item for item in result["diagnostics"] if item["code"] == expected_code)
    assert diagnostic["artifact"] == "inventory.json"
    assert diagnostic["path"] == expected_path
    assert diagnostic["source_path"] == "models/a.xlsx"
    index = json.loads((index_dir / "index.json").read_text(encoding="utf-8"))
    assert index["validation_status"] == "blocked"
    assert index["entries"][0]["artifacts"] == []

    checked = CliRunner().invoke(app, ["step2", "validate", "--index", str(index_dir / "index.json"), "--step1-root", str(root)])

    assert checked.exit_code == 1
    checked_result = json.loads(checked.stdout)
    assert checked_result["status"] == "blocked"
    assert any(item["code"] == expected_code and item["artifact"] == "inventory.json" for item in checked_result["diagnostics"])


def test_step2_index_exits_nonzero_when_handoff_is_outside_root(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    root.mkdir()
    handoff = tmp_path / "outside.json"
    handoff.write_text(json.dumps({"schema_version": "step1.v1"}), encoding="utf-8")

    result = CliRunner().invoke(app, ["step2", "index", "--handoff", str(handoff), "--step1-root", str(root)])

    assert result.exit_code != 0
    assert json.loads(result.stdout)["diagnostics"][0]["severity"] == "error"


def test_step2_index_rejects_source_json_with_same_version_marker(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    source = root / "run/source.json"
    source.parent.mkdir(parents=True)
    source.write_text(json.dumps({
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
    }), encoding="utf-8")

    result = CliRunner().invoke(app, ["step2", "index", "--handoff", str(source), "--step1-root", str(root), "--out", str(tmp_path / "index")])

    assert result.exit_code != 0
    assert json.loads(result.stdout)["status"] == "blocked"
    assert "source" in json.loads(result.stdout)["diagnostics"][0]["message"]
