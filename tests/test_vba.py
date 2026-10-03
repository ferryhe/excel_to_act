"""VBA extraction and VBA→cell edge building (no oletools needed for the logic)."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile

import builtins
from openpyxl import Workbook

from excel_to_act.ingest.vba import extract_vba_project
from excel_to_act.inventory.vba_links import build_vba_edges, extract_vba_cell_links
from excel_to_act.schemas import VbaModule


def _project(modules: list[VbaModule]):
    return SimpleNamespace(modules=modules)


def _module(code: str, name: str = "Module1") -> VbaModule:
    return VbaModule(name=name, kind="StdModule", code=code, procedures=[])


def test_cell_refs_are_extracted_with_confidence() -> None:
    code = (
        'Range("B7").Value = 1\n'
        'Range("Inputs!$B$2").Value = Rate\n'
        'Names("Mort_qx").RefersToRange.Value = 1\n'
        "Cells(2, 3).Value = 9\n"
    )
    refs = extract_vba_cell_links(_project([_module(code)]))
    by_kind = {(r.kind, r.raw): r for r in refs}

    range_bare = by_kind[("range", 'Range("B7")')]
    assert range_bare.address == "B7"
    assert range_bare.sheet is None
    assert range_bare.confidence == 0.6

    range_qualified = by_kind[("range", 'Range("Inputs!$B$2")')]
    assert range_qualified.address == "B2"
    assert range_qualified.sheet == "Inputs"
    assert range_qualified.confidence == 0.9

    name_ref = by_kind[("name", 'Names("Mort_qx")')]
    assert name_ref.name == "Mort_qx"
    assert name_ref.unresolved is True

    cells_ref = by_kind[("cells", "Cells(2, 3)")]
    assert cells_ref.unresolved is True


def test_edges_share_formula_graph_node_space() -> None:
    code = 'Range("Inputs!B2").Value = 1\nNames("Mort_qx").RefersToRange.Value = 1\n'
    refs = extract_vba_cell_links(_project([_module(code)]))
    nodes, edges = build_vba_edges(refs, workbook_path="wb.xlsx")

    assert any(n.id == "vba:Module1" and n.kind.value == "vba" for n in nodes)
    cell_edge = next(e for e in edges if e.target == "cell:Inputs!B2")
    assert cell_edge.relationship == "vba_ref"
    assert cell_edge.confidence == 0.9
    name_edge = next(e for e in edges if e.target == "name:Mort_qx")
    assert name_edge.relationship == "vba_ref"
    # Unresolved Cells references must not create a phantom edge.
    assert all(e.relationship == "vba_ref" for e in edges)


def test_project_without_vba_reports_unavailable(tmp_path: Path) -> None:
    wb = Workbook()
    wb.save(tmp_path / "book.xlsx")
    wb.close()
    result = extract_vba_project(tmp_path / "book.xlsx")
    assert result.available is False
    assert result.oletools_missing is False


def test_vba_present_without_oletools_is_graceful(tmp_path: Path) -> None:
    wb = Workbook()
    wb.save(tmp_path / "macro.xlsm")
    wb.close()
    # Inject a dummy vbaProject.bin so the package looks macro-enabled.
    with ZipFile(tmp_path / "macro.xlsm", "a") as zf:
        zf.writestr("xl/vbaProject.bin", b"dummy")

    original_import = builtins.__import__

    def fake_import(name, *args, **kwargs):  # noqa: ANN002, ANN003
        if name.startswith("oletools"):
            raise ImportError("oletools not installed")
        return original_import(name, *args, **kwargs)

    with patch("builtins.__import__", side_effect=fake_import):
        result = extract_vba_project(tmp_path / "macro.xlsm")
    assert result.available is True
    assert result.oletools_missing is True
    assert result.modules == []
