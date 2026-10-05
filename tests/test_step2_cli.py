import json
from pathlib import Path

from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app


def test_step2_tools_lists_initial_actions() -> None:
    result = CliRunner().invoke(app, ["step2", "tools"])

    assert result.exit_code == 0
    assert [item["name"] for item in json.loads(result.stdout)["tools"]] == [
        "step2.handoff.resolve",
        "step2.index.build",
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
