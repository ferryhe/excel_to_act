"""Synthetic routing fixtures for fresh seven-stage workflows; never publication evidence."""

import json
from pathlib import Path

import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps import conversion_workflow as workflow
from excel_to_act.steps.seven_stage import bind_native_comparison, create_candidate
from excel_to_act.steps.step6.workflow import create_conversion_report as legacy_report
from excel_to_act.steps.step1.checkpoint import create_checkpoint
from excel_to_act.steps.step2 import checkpoint as step2_checkpoint
from excel_to_act.steps.step3.workflow import prepare_analysis
from excel_to_act.steps.step3.tools.interface import export_tool_interface
from excel_to_act.steps.step5.workflow import validate_generated


def _json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value), encoding="utf-8")


def _run(tmp_path: Path) -> Path:
    run = tmp_path / "synthetic-run"
    workbook = tmp_path / "synthetic.xlsx"
    workbook.write_bytes(b"synthetic source bytes")
    digest = workflow.hash_file(workbook)
    source = {"schema_version": "step1.v1", "source_path": str(workbook),
              "relative_path": workbook.name, "sha256": digest, "run_id": "synthetic-run"}
    _json(run / "source.json", source)
    _json(run / "workbook_manifest.json", {"artifact_type": "workbook_manifest", "sha256": digest})
    _json(run / "quality.json", {"status": "pass", "ready_for_next_step": True})
    _json(run / "inventory.json", {})
    _json(run / "handoff.json", {})
    copy = run / "source_evidence" / workbook.name
    copy.parent.mkdir(parents=True)
    copy.write_bytes(workbook.read_bytes())
    files = {path.relative_to(run).as_posix(): workflow.hash_file(path)
             for path in run.rglob("*") if path.is_file()}
    _json(run / "promotion.json", {"files": files})
    return run


def _native(tmp_path: Path) -> Path:
    source = tmp_path / "synthetic.xlsx"
    source.write_bytes(b"synthetic source bytes")
    root, manifest = workflow.create_workflow(tmp_path / "fresh", {
        "workbook_path": str(source), "workbook_sha256": workflow.hash_file(source),
        "run_id": "synthetic", "step1_run_path": str(tmp_path / "synthetic-run"),
        "fixture": True,
    })
    manifest["mode"] = "synthetic_fixture"
    workflow._write_manifest(root, manifest)
    return root


def _append(root: Path, stage: int, *, boundary: bool = False) -> dict:
    _, manifest = workflow.load_workflow(root)
    previous = workflow.current_revision(manifest, stage - 1) if stage > 1 else None
    payload = ({"schema_version": "step3.input_boundary.v1", "phase": "input_boundary",
                "boundary_sha256": "b" * 64, "source": manifest["source"], "target_selection": {}}
               if boundary else {"status": "ready_for_review", "stage": stage})
    return workflow.append_stage_artifact(root, stage, f"stage{stage}.json", payload, "synthetic",
                                          input_stages={stage - 1: previous["artifact"]} if previous else None)


def _approve(root: Path, stage: int) -> None:
    workflow.record_decision(root, stage, "agent", "approve", "synthetic Agent review")
    workflow.record_decision(root, stage, "human", "approve", "synthetic human fixture")


def test_legacy_report_api_rejects_native_v2_without_changing_delivery_or_history(tmp_path: Path) -> None:
    root = _native(tmp_path)
    for stage in range(1, 6):
        _append(root, stage)
        _approve(root, stage)
    _, manifest = workflow.load_workflow(root)
    pointer = {"version": "synthetic-active", "receipt_id": "synthetic-receipt"}
    manifest["delivery"] = pointer
    manifest["delivery_history"] = [pointer]
    workflow._write_manifest(root, manifest)
    before = {path.relative_to(root): path.read_bytes() for path in root.rglob("*") if path.is_file()}

    result = legacy_report(root)

    assert result["status"] == "blocked" and "legacy Step 6" in result["reason"]
    assert {path.relative_to(root): path.read_bytes() for path in root.rglob("*") if path.is_file()} == before
    assert workflow.load_workflow(root)[1]["delivery"] == pointer


def test_fresh_step1_public_api_and_cli(tmp_path: Path) -> None:
    run = _run(tmp_path)
    root = tmp_path / "api"
    result = create_checkpoint(run, root)
    assert result["status"] == "pass", result
    _, manifest = workflow.load_workflow(root)
    assert workflow.workflow_origin(manifest) == "native"
    assert manifest["schema_version"] == "conversion.workflow.v2"
    assert "parent" not in manifest
    assert workflow.stage_status(root, manifest, 1)["status"] == "awaiting_confirmation"
    cli_root = tmp_path / "cli"
    cli = CliRunner().invoke(app, ["step1", "report", "--run", str(run), "--out", str(cli_root)])
    assert cli.exit_code == 0, cli.output
    assert workflow.workflow_origin(workflow.load_workflow(cli_root)[1]) == "native"
    for path in (root, cli_root):
        _, synthetic = workflow.load_workflow(path)
        synthetic["mode"] = "synthetic_fixture"
        synthetic["source"]["fixture"] = True
        workflow._write_manifest(path, synthetic)


def test_native_step2_existing_checkpoint_producer_and_cli(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    run = _run(tmp_path)
    root = tmp_path / "native"
    assert create_checkpoint(run, root)["status"] == "pass"
    _, manifest = workflow.load_workflow(root)
    manifest["mode"] = "synthetic_fixture"
    manifest["source"]["fixture"] = True
    workflow._write_manifest(root, manifest)
    _approve(root, 1)
    index = tmp_path / "synthetic-index.json"
    _json(index, {"schema_version": "step2.index.v1", "status": "pass", "entries": [
        {"source_id": "synthetic-source", "run_id": "synthetic-run",
         "source_sha256": manifest["source"]["workbook_sha256"],
         "run_path": run.name, "source_path": str(tmp_path / "synthetic.xlsx"),
         "ready_for_next_step": True, "status": "pass"}]})
    monkeypatch.setattr(step2_checkpoint, "_validate_saved_index",
                        lambda *_args: {"status": "pass", "metrics": {"artifacts_checked": 1}})
    cli = CliRunner().invoke(app, ["step2", "report", "--index", str(index),
                                   "--step1-root", str(tmp_path), "--workflow", str(root)])
    assert cli.exit_code == 0, cli.output
    _, manifest = workflow.load_workflow(root)
    assert manifest["source"]["source_id"] == "synthetic-source"
    assert workflow.stage_status(root, manifest, 2)["status"] == "awaiting_confirmation"


def test_native_stage_review_return_and_comparison_gate(tmp_path: Path) -> None:
    root = _native(tmp_path)
    for stage in range(1, 6):
        _append(root, stage)
        with pytest.raises(ValueError, match="Agent must approve"):
            workflow.record_decision(root, stage, "human", "approve", "synthetic premature review")
        _approve(root, stage)
        assert workflow.stage_status(root, workflow.load_workflow(root)[1], stage)["status"] == "simulated_approved"
    assert workflow.workflow_status(root)["next_permitted_stage"] == 6
    with pytest.raises(ValueError, match="synthetic|comparison|Stage"):
        bind_native_comparison(root, tmp_path / "none.json", "0" * 64)
    with pytest.raises(ValueError, match="synthetic|comparison|Stage"):
        create_candidate(root, "v1", tmp_path / "request.json", tmp_path / "none.json")
    workflow.record_decision(root, 5, "agent", "reject", "synthetic correction", return_to=3)
    assert workflow.stage_status(root, workflow.load_workflow(root)[1], 3)["status"] == "stale"
    assert workflow.workflow_status(root)["next_permitted_stage"] == 3


def test_native_input_boundary_requires_human_and_attached_upstream_is_read_only(tmp_path: Path) -> None:
    root = _native(tmp_path)
    for stage in (1, 2):
        _append(root, stage)
        _approve(root, stage)
    _append(root, 3, boundary=True)
    workflow.record_decision(root, 3, "agent", "approve", "synthetic Agent review")
    assert workflow.input_boundary_status(root, workflow.load_workflow(root)[1])["status"] != "input_boundary_confirmed"
    with pytest.raises(ValueError, match="TypeSafe"):
        workflow.record_decision(root, 3, "typesafe", "approve", "synthetic bypass")
    workflow.record_decision(root, 3, "human", "approve", "synthetic human fixture")
    assert workflow.input_boundary_status(root, workflow.load_workflow(root)[1])["status"] == "simulated_input_boundary_confirmed"
    with pytest.raises(ValueError, match="separately reviewed final design"):
        workflow.require_stage_approved(root, 3)


def test_synthetic_native_upstream_revision_revokes_pointer(tmp_path: Path) -> None:
    root = _native(tmp_path)
    _append(root, 1)
    _, manifest = workflow.load_workflow(root)
    manifest["delivery"] = {"version": "synthetic", "receipt_id": "synthetic-only"}
    workflow._write_manifest(root, manifest)
    _append(root, 1)
    _, changed = workflow.load_workflow(root)
    assert changed["delivery"] is None
    assert changed["delivery_history"][-1]["schema_version"] == "step6.delivery_revocation.v1"
    assert changed["delivery_history"][-1]["receipt_id"] == "synthetic-only"


def test_malformed_origin_and_legacy_write_reject_without_mutation(tmp_path: Path) -> None:
    root = _native(tmp_path)
    _, manifest = workflow.load_workflow(root)
    for change in ({"origin": "unknown"}, {"origin": "attached"}, {"origin": None},
                   {"parent": {"root": str(root), "manifest_sha256": "0" * 64}}):
        changed = {**manifest, **change}
        workflow._write_manifest(root, changed)
        with pytest.raises(ValueError, match="origin"):
            workflow.load_workflow(root)
    workflow._write_manifest(root, manifest)
    legacy = tmp_path / "legacy"
    legacy.mkdir()
    old = {"schema_version": "conversion.workflow.v1", "source": manifest["source"],
           "stages": {str(stage): {"current_revision": 0, "revisions": []}
                      for stage in workflow.STAGES}}
    workflow._write_manifest(legacy, old)
    before = (legacy / "workflow.json").read_bytes()
    with pytest.raises(ValueError, match="read-only"):
        workflow.create_workflow(legacy, manifest["source"])
    with pytest.raises(ValueError, match="read-only"):
        workflow.append_stage_artifact(legacy, 1, "stage1.json", {}, "synthetic")
    assert (legacy / "workflow.json").read_bytes() == before

    attached = tmp_path / "attached"
    attached.mkdir()
    old_attached = {"schema_version": "conversion.workflow.v2", "mode": "production",
                    "source": manifest["source"],
                    "parent": {"root": str(legacy.resolve()),
                               "manifest_sha256": workflow.hash_file(legacy / "workflow.json")},
                    "stages": {str(stage): {"current_revision": 0, "revisions": []}
                               for stage in workflow.V2_STAGES}, "delivery": None,
                    "delivery_history": []}
    workflow._write_manifest(attached, old_attached)
    assert workflow.workflow_origin(workflow.load_workflow(attached)[1]) == "attached"
    attached_before = (attached / "workflow.json").read_bytes()
    with pytest.raises(ValueError, match="unsupported workflow stage"):
        workflow.append_stage_artifact(attached, 1, "stage1.json", {}, "synthetic")
    with pytest.raises(ValueError, match="allowed stage"):
        workflow.record_decision(attached, 1, "agent", "approve", "synthetic")
    prepared = prepare_analysis(tmp_path / "absent-index.json", tmp_path,
                                tmp_path / "adjacent-analysis", workflow_path=attached)
    exported = export_tool_interface(attached)
    validated = validate_generated(attached)
    for result in (prepared, exported, validated):
        assert result["status"] == "blocked"
        assert "read-only" in str(result)
    assert not (tmp_path / "adjacent-analysis").exists()
    assert (attached / "workflow.json").read_bytes() == attached_before
