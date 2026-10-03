"""Cached formula results stored inside the workbook.

Excel writes the last calculated result next to every formula. Reading those
results is the only zero-dependency source of "what the workbook actually
evaluates to", which makes it the primary oracle for later reconciliation.

Workbooks that were never recalculated by Excel (for example files written by
openpyxl) contain empty ``<v/>`` elements, so openpyxl reports ``None``. A
``None`` entry therefore means "no cached value available", never "the value
is None".
"""

from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

CellValue = str | int | float | bool | None

# Coordinates are keyed by sheet title because sheet-scoped addresses are the
# only stable identity openpyxl exposes for both the formula and value passes.
CachedValueMap = dict[tuple[str, str], CellValue]


def _safe_value(value: Any) -> CellValue:
    """Coerce a raw openpyxl value into the scalar union allowed by contracts."""

    if value is None or isinstance(value, str | int | float | bool):
        return value
    if isinstance(value, datetime | date | time):
        return value.isoformat()
    return str(value)


def read_cached_values(workbook_path: Path) -> CachedValueMap:
    """Return ``{(sheet_title, coordinate): cached value}`` for every stored result.

    Cells without a stored result are absent from the map, so callers must use
    ``key in map`` rather than relying on a ``None`` value.
    """

    workbook_path = workbook_path.expanduser().resolve()
    values: CachedValueMap = {}
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            # In read-only mode the iteration bounds come from the declared
            # <dimension>, which non-Excel writers can omit or understate.
            reset = getattr(ws, "reset_dimensions", None)
            if callable(reset):
                reset()
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    values[(ws.title, cell.coordinate)] = _safe_value(cell.value)
    finally:
        wb.close()
    return values
