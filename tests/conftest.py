"""Keep pre-migration workflow fixtures available for historical evaluator tests."""

from pathlib import Path

import pytest

from excel_to_act.steps import conversion_workflow
from excel_to_act.steps.step6 import workflow as legacy_report_workflow


_HISTORICAL_FIXTURE_MODULES = {
    "test_conversion_workflow.py", "test_seven_stage.py", "test_step3_tool_interface.py",
    "test_step3_workflow.py", "test_step4_external_inputs.py", "test_step4_generation.py",
    "test_step5_workflow.py", "test_step6_report.py", "test_workflow_catalogs.py",
}


def _legacy_fixture_workflow(root: Path, source: dict):
    """Build old v1 fixtures only in the historical test modules below."""
    root = root.expanduser().resolve()
    manifest_path = root / conversion_workflow.MANIFEST_NAME
    if manifest_path.exists():
        return conversion_workflow.load_workflow(root)
    manifest = {"schema_version": "conversion.workflow.v1", "created_at": conversion_workflow._now(),
                "source": source, "mode": "production",
                "stages": {str(stage): {"current_revision": 0, "revisions": []}
                           for stage in conversion_workflow.STAGES}}
    conversion_workflow._write_manifest(root, manifest)
    return root, manifest


@pytest.fixture(autouse=True)
def historical_workflow_fixture_writes(request: pytest.FixtureRequest, monkeypatch: pytest.MonkeyPatch) -> None:
    if request.node.name == "test_legacy_python_write_entries_leave_history_untouched":
        return
    if Path(str(request.fspath)).name in _HISTORICAL_FIXTURE_MODULES:
        # These existing tests construct v1 history. Public read-only behavior is tested unpatched.
        original = conversion_workflow.create_workflow
        if getattr(request.module, "create_workflow", None) is original:
            monkeypatch.setattr(request.module, "create_workflow", _legacy_fixture_workflow)
        monkeypatch.setattr(conversion_workflow, "create_workflow", _legacy_fixture_workflow)
        monkeypatch.setattr(conversion_workflow, "_require_writable", lambda _manifest: None)
        monkeypatch.setattr(legacy_report_workflow, "_require_writable", lambda _manifest: None)
