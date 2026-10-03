"""Excel what-if data tables (``<f t="dataTable">``) read straight from sheet XML.

A What-If data table is not a separate OOXML element: Excel stores it as the
formula of the table's corner cell, marked with ``t="dataTable"`` and carrying
the table range plus the row/column input cells in R1C1 notation:

    <c r="B7"><f t="dataTable" ref="B7:C9" dt2D="1" dtr="0" r1="R2C2" r2="R3C4">=B2*B3</f></c>

openpyxl reads that cell into a ``DataTableFormula`` object but **drops the
formula text**, and any caller doing ``str(cell.value)`` silently turns it into
a repr. Both the geometry and the text are therefore read from XML here.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.utils.cell import get_column_letter

_DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_R1C1 = re.compile(r"R(\d+)C(\d+)")
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


def _r1c1_to_a1(value: str | None) -> str | None:
    """Convert absolute R1C1 (``R2C2``) to ``$B$2``; relative refs are not resolvable."""

    if not value:
        return None
    match = _R1C1.fullmatch(value.strip())
    if match is None:
        return None
    return f"${get_column_letter(int(match.group(2)))}${match.group(1)}"


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
                tables.append(
                    DataTableSpec(
                        corner_cell=cell.attrib.get("r", ""),
                        ref=formula.attrib.get("ref"),
                        formula_text=f"={text}" if text and not text.startswith("=") else (text or None),
                        row_input_cell=_r1c1_to_a1(formula.attrib.get("r1")),
                        col_input_cell=_r1c1_to_a1(formula.attrib.get("r2")),
                        two_dimensional=formula.attrib.get("dt2D") in {"1", "true"},
                        raw_r1=formula.attrib.get("r1"),
                        raw_r2=formula.attrib.get("r2"),
                        dtr=formula.attrib.get("dtr"),
                    )
                )
            if tables:
                result[sheet_name] = tables
    return result
