"""Cached formula values: dual-load extraction and missing-cache diagnostics."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook

from excel_to_act.ingest.cached_values import read_cached_values
from excel_to_act.ingest.openpyxl_reader import OpenpyxlWorkbookReader
from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor

FORMULA_CELLS = {"C2": "=B1*B2", "D2": "=SUM(B1:B2)"}


def build_fixture(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["B1"] = 0.05
    ws["B2"] = 100
    for address, formula in FORMULA_CELLS.items():
        ws[address] = formula
    wb.save(path)
    wb.close()
    return path


def inject_cached_values(
    path: Path, values: dict[str, str], cell_types: dict[str, str] | None = None
) -> None:
    """Make an openpyxl-written workbook look like Excel recalculated it.

    openpyxl writes formula cells as ``<f>...</f><v></v>``; filling ``<v>`` is
    the only way to obtain a workbook with cached results without Excel.
    """
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        payload = {name: zf.read(name) for name in names}
    part = "xl/worksheets/sheet1.xml"
    root = ET.fromstring(payload[part])
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    for cell in root.findall(".//s:c", ns):
        address = cell.attrib["r"]
        if address not in values:
            continue
        value = cell.find("s:v", ns)
        assert value is not None
        value.text = values[address]
        if cell_types and address in cell_types:
            cell.set("t", cell_types[address])
    payload[part] = ET.tostring(root, encoding="utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, payload[name])


def _inventory(path: Path):
    manifest = OpenpyxlWorkbookReader().read_manifest(path)
    return OpenpyxlInventoryExtractor().extract(path, manifest)


def _cell(inventory, address: str):
    sheet = inventory.sheets[0]
    return next(c for c in sheet.cells if c.address == address)


def test_openpyxl_written_workbook_has_no_cached_values(tmp_path: Path) -> None:
    fixture = build_fixture(tmp_path / "no_cache.xlsx")
    inventory = _inventory(fixture)

    for address in FORMULA_CELLS:
        cell = _cell(inventory, address)
        assert cell.formula is not None
        assert cell.cached_value is None
        assert cell.cached_value_available is False

    missing = [f for f in inventory.unsupported_features if f.feature_type == "missing_cached_values"]
    assert len(missing) == 1
    assert missing[0].severity == "warning"
    assert missing[0].metadata["formula_cells"] == len(FORMULA_CELLS)


def test_cached_values_are_read_when_excel_stored_them(tmp_path: Path) -> None:
    fixture = build_fixture(tmp_path / "with_cache.xlsx")
    inject_cached_values(fixture, {"C2": "5", "D2": "105"})

    values = read_cached_values(fixture)
    assert values[("Sheet1", "C2")] == 5
    assert values[("Sheet1", "D2")] == 105

    inventory = _inventory(fixture)
    c2 = _cell(inventory, "C2")
    assert c2.formula == "=B1*B2"
    assert c2.cached_value == 5
    assert c2.cached_value_available is True
    assert _cell(inventory, "D2").cached_value == 105
    assert not [f for f in inventory.unsupported_features if f.feature_type == "missing_cached_values"]


def test_literal_cells_report_no_cached_value(tmp_path: Path) -> None:
    fixture = build_fixture(tmp_path / "literals.xlsx")
    inject_cached_values(fixture, {"C2": "5"})
    inventory = _inventory(fixture)

    literal = _cell(inventory, "B1")
    assert literal.kind == "value"
    assert literal.value == 0.05
    assert literal.cached_value is None
    assert literal.cached_value_available is False


def test_cached_value_boundaries_preserve_availability_and_types(tmp_path: Path) -> None:
    fixture = tmp_path / "cache_boundaries.xlsx"
    formulas = {"A1": "=0", "B1": "=FALSE()", "C1": '=\"\"', "D1": "=1+1"}
    expected = {"A1": 0, "B1": False, "C1": "", "D1": None}
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for address, formula in formulas.items():
        ws[address] = formula
    ws["E1"] = "literal"
    wb.save(fixture)
    wb.close()
    inject_cached_values(
        fixture,
        {"A1": "0", "B1": "0", "C1": "", "D1": ""},
        cell_types={"B1": "b", "C1": "str"},
    )

    values = read_cached_values(fixture)
    assert values == {
        ("Sheet1", "A1"): 0,
        ("Sheet1", "B1"): False,
        ("Sheet1", "C1"): "",
        ("Sheet1", "E1"): "literal",
    }
    assert {key: type(value) for key, value in values.items()} == {
        ("Sheet1", "A1"): int,
        ("Sheet1", "B1"): bool,
        ("Sheet1", "C1"): str,
        ("Sheet1", "E1"): str,
    }

    inventory = _inventory(fixture)
    sheet = inventory.sheets[0]
    assert len(sheet.cells) == 5
    assert {cell.address for cell in sheet.cells} == {*formulas, "E1"}
    assert len({cell.address for cell in sheet.cells}) == len(sheet.cells)
    for address, expected_value in expected.items():
        cell = _cell(inventory, address)
        assert cell.formula == formulas[address]
        assert cell.cached_value == expected_value
        assert type(cell.cached_value) is type(expected_value)
        assert cell.cached_value_available is (address != "D1")

    literal = _cell(inventory, "E1")
    assert literal.kind == "value"
    assert literal.formula is None
    assert literal.cached_value is None
    assert literal.cached_value_available is False


def test_dual_load_does_not_double_count_coverage(tmp_path: Path) -> None:
    fixture = build_fixture(tmp_path / "count.xlsx")
    inject_cached_values(fixture, {"C2": "5"})
    inventory = _inventory(fixture)

    sheet = inventory.sheets[0]
    assert len(sheet.cells) == 4  # B1, B2, C2, D2
    assert len({c.address for c in sheet.cells}) == len(sheet.cells)
    assert (
        inventory.coverage.recognized_inventory_objects + inventory.coverage.unsupported_or_opaque_objects
        == inventory.coverage.discovered_workbook_objects
    )
