"""Direct legacy write APIs must reject without changing synthetic history."""

from excel_to_act.steps import conversion_workflow as conversion
from excel_to_act.steps.step3.tools.interface import export_tool_interface
from excel_to_act.steps.step3.workflow import prepare_analysis
from excel_to_act.steps.step3.input_boundary import create_input_boundary
from excel_to_act.steps.step4.discovery import discover_active_trace
from excel_to_act.steps.step4.external_inputs import capture_external_inputs
from excel_to_act.steps.step4.generator import generate_model
from excel_to_act.steps.step4.implementation import create_implementation_plan
from excel_to_act.steps.step5.workflow import capture_oracle, reconcile, validate_generated
from test_step3_workflow import _prepared_fixture


def test_direct_legacy_entry_points_leave_workflow_and_outputs_untouched(tmp_path, monkeypatch):
    with monkeypatch.context() as setup:
        setup.setattr(conversion, "_require_writable", lambda _manifest: None)
        root = tmp_path / "legacy"
        conversion._write_manifest(root, {
            "schema_version": "conversion.workflow.v1", "mode": "production",
            "source": {"workbook_sha256": "source-sha", "source_id": "source-0", "run_id": "run-0"},
            "stages": {str(stage): {"current_revision": 0, "revisions": []}
                       for stage in conversion.STAGES},
        })
        previous = None
        for stage in range(1, 5):
            revision = conversion.append_stage_artifact(
                root, stage, f"stage{stage}.json", {"status": "ready_for_review"},
                f"Synthetic Stage {stage}.",
                input_stages={stage - 1: previous["artifact"]} if previous else None)
            for reviewer in ("agent", "human"):
                conversion.record_decision(root, stage, reviewer, "approve", "Synthetic review.")
            previous = revision
    _analysis, step1_root, _result = _prepared_fixture(tmp_path / "index-fixture", monkeypatch)
    index = tmp_path / "index-fixture" / "index.json"
    before = {p.relative_to(tmp_path): p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    outputs = tmp_path / "new-analysis"
    prepared = prepare_analysis(index, step1_root, outputs, workflow_path=root)
    exported = export_tool_interface(root)
    validated = validate_generated(root)
    assert prepared["status"] == exported["status"] == validated["status"] == "blocked"
    assert "read-only" in prepared["diagnostics"][0]["message"]
    assert "read-only" in exported["diagnostics"][0]["message"]
    assert "read-only" in validated["reason"]
    for result in (
        create_input_boundary(outputs, root, index, index),
        discover_active_trace(root),
        capture_external_inputs(root),
        create_implementation_plan(root, index),
        generate_model(root),
        capture_oracle(root),
        reconcile(root, index, index),
    ):
        assert result["status"] == "blocked" and "read-only" in str(result)
    assert not outputs.exists()
    assert not (root / "tool_interfaces").exists()
    assert not (root / "stage5").exists()
    assert {p.relative_to(tmp_path): p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()} == before
