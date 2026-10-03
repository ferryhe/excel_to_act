"""Excel what-if data tables (``<f t="dataTable">``) read straight from sheet XML.

A What-If data table is not a separate OOXML element: Excel stores it as the
formula of the table's corner cell, marked with ``t="dataTable"`` and carrying
the table range plus the row/column input cells as A1 cell references:

    <c r="B7"><f t="dataTable" ref="B7:C9" dt2D="1" dtr="0" r1="B2" r2="D3">=B2*B3</f></c>

openpyxl reads that cell into a ``DataTableFormula`` object but **drops the
formula text**, and any caller doing ``str(cell.value)`` silently turns it into
a repr. Both the geometry and the text are therefore read from XML here.
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils.exceptions import CellCoordinatesException

_DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}


@dataclass(frozen=True)
class DataTableSpec:
    """One what-if data table anchored at ``corner_cell``."""

    corner_cell: str
    ref: str | None = None
    formula_text: str | None = None
    row_input_cell: str | None = None
    col_input_cell: str | None = None
    two_dimensional: bool = False
    raw_r1: str | None = None
    raw_r2: str | None = None
    dtr: str | None = None


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def resolve_package_part(base_dir: str, target: str) -> str:
    """Resolve an OPC relationship target into a package part name.

    Targets may be package-absolute (``/xl/worksheets/sheet1.xml``) or relative
    to the source part (``worksheets/sheet1.xml``, possibly with ``../``).
    """

    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _cell_ref_to_a1(value: str | None) -> str | None:
    """Normalize an OOXML A1 cell reference such as ``B2`` to ``$B$2``."""
    if not value:
        return None
    try:
        column, row = coordinate_from_string(value.strip().replace("$", ""))
    except (CellCoordinatesException, TypeError):
        return None
    return f"${column}${row}"


def sheet_xml_parts(workbook_path: Path) -> dict[str, str]:
    """Return ``{sheet_name: xl/worksheets/sheetN.xml}`` for a workbook package."""

    workbook_path = workbook_path.expanduser().resolve()
    if workbook_path.suffix.lower() not in _SUPPORTED_SUFFIXES:
        return {}
    parts: dict[str, str] = {}
    with zipfile.ZipFile(workbook_path) as zf:
        try:
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        except (KeyError, ET.ParseError):
            return {}
        target_by_id = {rel.attrib.get("Id"): rel.attrib.get("Target", "") for rel in rels}
        for element in workbook.iter():
            if _local(element.tag) != "sheet":
                continue
            target = target_by_id.get(element.attrib.get(f"{_DOC_REL}id") or element.attrib.get("id"), "")
            if not target:
                continue
            parts[element.attrib.get("name", "")] = resolve_package_part("xl", target)
    return parts


def read_data_tables(workbook_path: Path) -> dict[str, list[DataTableSpec]]:
    """Return ``{sheet_name: [DataTableSpec]}`` for every what-if data table found.

    Sheets without data tables are absent from the result. Malformed packages
    yield an empty mapping rather than raising, so inventory extraction never
    depends on this optional enrichment succeeding.
    """

    workbook_path = workbook_path.expanduser().resolve()
    result: dict[str, list[DataTableSpec]] = {}
    parts = sheet_xml_parts(workbook_path)
    if not parts:
        return result
    with zipfile.ZipFile(workbook_path) as zf:
        for sheet_name, part in parts.items():
            if part not in zf.namelist():
                continue
            try:
                root = ET.fromstring(zf.read(part))
            except (KeyError, ET.ParseError):
                continue
            tables: list[DataTableSpec] = []
            for cell in root.iter():
                if _local(cell.tag) != "c":
                    continue
                formula = next(
                    (child for child in cell if _local(child.tag) == "f" and child.attrib.get("t") == "dataTable"),
                    None,
                )
                if formula is None:
                    continue
                text = (formula.text or "").strip()
                two_dimensional = formula.attrib.get("dt2D") in {"1", "true"}
                row_oriented = formula.attrib.get("dtr") in {"1", "true"}
                r1 = _cell_ref_to_a1(formula.attrib.get("r1"))
                r2 = _cell_ref_to_a1(formula.attrib.get("r2"))
                if two_dimensional:
                    row_input_cell, col_input_cell = r1, r2
                elif row_oriented:
                    row_input_cell, col_input_cell = r1, None
                else:
                    row_input_cell, col_input_cell = None, r1
                tables.append(
                    DataTableSpec(
                        corner_cell=cell.attrib.get("r", ""),
                        ref=formula.attrib.get("ref"),
                        formula_text=f"={text}" if text and not text.startswith("=") else (text or None),
                        row_input_cell=row_input_cell,
                        col_input_cell=col_input_cell,
                        two_dimensional=two_dimensional,
                        raw_r1=formula.attrib.get("r1"),
                        raw_r2=formula.attrib.get("r2"),
                        dtr=formula.attrib.get("dtr"),
                    )
                )
            if tables:
                result[sheet_name] = tables
    return result
