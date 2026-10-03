"""Completeness verification and the handoff artifact that ends Step 1."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.orchestrator.phase1 import Phase1Orchestrator
from excel_to_act.schemas import (
    CompletenessReport,
    CompletenessStatus,
    FormulaGraph,
    Handoff,
    ModuleClassification,
    SheetManifest,
    WorkbookInventory,
    WorkbookManifest,
)
from excel_to_act.schemas.artifacts import CoverageSummary
from excel_to_act.store.local_store import LocalArtifactStore
from excel_to_act.verify.completeness import verify_completeness


def make_fixture(path: Path) -> Path:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    ws["B1"] = 0.05
    ws["B2"] = 100
    ws["C2"] = "=B1*B2"
    wb.save(path)
    wb.close()
    return path


def _manifest() -> WorkbookManifest:
    return WorkbookManifest(
        workbook_path="wb.xlsx",
        file_name="wb.xlsx",
        file_size=1,
        sha256="a" * 64,
        sheets=[SheetManifest(name="Inputs", index=0, max_row=2, max_column=3)],
    )


def test_completeness_and_handoff_are_written(tmp_path: Path) -> None:
    fixture = make_fixture(tmp_path / "fixture.xlsx")
    out = tmp_path / "artifacts"
    metadata = Phase1Orchestrator().run(fixture, out)

    assert metadata.completeness_status in {"pass", "warn"}

    completeness = LocalArtifactStore(out).read_json(
        next(a.path for a in metadata.artifacts if a.name == "completeness.json"), CompletenessReport
    )
    assert completeness.status in {"pass", "warn"}
    check_names = {check.name for check in completeness.checks}
    assert {
        "sheets_accounted",
        "cells_accounted",
        "content_parts_accounted",
        "formulas_linked",
        "coverage_arithmetic",
    } <= check_names

    handoff_path = Path(next(a.path for a in metadata.artifacts if a.name == "handoff.json"))
    markdown_path = Path(next(a.path for a in metadata.artifacts if a.name == "handoff.md"))
    assert handoff_path.exists() and markdown_path.exists()
    # The markdown is the human-facing twin: it must carry a one-glance summary
    # and be readable without opening any JSON.
    markdown = markdown_path.read_text(encoding="utf-8")
    assert "Handoff" in markdown and "一眼看懂" in markdown and "下一步" in markdown

    handoff = LocalArtifactStore(out).read_json(handoff_path, Handoff)
    assert handoff.summary.get("sheets") == 1
    assert handoff.summary.get("cells", 0) > 0
    assert handoff.summary.get("formula_cells") == 1
    assert any(artifact.kind == "inventory" for artifact in handoff.artifacts)

    # The convenience alias at the output root keeps the handoff discoverable.
    assert (out / "handoff.md").exists()

    store = LocalArtifactStore(out)
    reloaded = store.read_run(metadata.workbook_sha256, metadata.run_id)
    assert reloaded.completeness_status == metadata.completeness_status


def test_missing_sheet_blocks_completeness() -> None:
    manifest = _manifest()
    inventory = WorkbookInventory(
        workbook_sha256=manifest.sha256,
        sheets=[],  # the single sheet from workbook.xml never made it into the output
        coverage=CoverageSummary(
            recognized_inventory_objects=0,
            unsupported_or_opaque_objects=0,
            discovered_workbook_objects=0,
        ),
    )
    report = verify_completeness(manifest, inventory, FormulaGraph(), ModuleClassification())

    assert report.status == "fail"
    assert report.blocking_reasons
    assert any(gap.object_type == "sheet" for gap in report.gaps)
    assert report.status == CompletenessStatus.fail.value


def test_cli_exits_nonzero_when_completeness_fails(tmp_path: Path) -> None:
    fixture = make_fixture(tmp_path / "fixture.xlsx")
    out = tmp_path / "artifacts"
    failing = CompletenessReport(
        workbook_sha256="a" * 64,
        status=CompletenessStatus.fail,
        blocking_reasons=["simulated gap"],
    )
    with patch("excel_to_act.orchestrator.phase1.verify_completeness", return_value=failing):
        result = CliRunner().invoke(app, ["inspect", str(fixture), "--out", str(out)])
    assert result.exit_code == 1
