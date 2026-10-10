"""Core inventory extraction from openpyxl with source locations."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from excel_to_act.ingest.cached_values import read_cached_values
from excel_to_act.ingest.data_table import read_data_tables
from excel_to_act.ingest.form_controls import read_form_controls
from excel_to_act.schemas import (
    CellInventory,
    CellKind,
    CoverageSummary,
    RangeInventory,
    SheetInventory,
    SourceLocation,
    UnsupportedFeature,
    UnsupportedSeverity,
    WorkbookInventory,
    WorkbookManifest,
)


def _loc(path: Path, object_type: str, sheet: str | None = None, index: int | None = None, address: str | None = None, object_id: str | None = None) -> SourceLocation:
    return SourceLocation(workbook_path=str(path), sheet_name=sheet, sheet_index=index, address=address, object_type=object_type, object_id=object_id)


def _safe_value(value: Any) -> str | int | float | bool | None:
    if value is None or isinstance(value, str | int | float | bool):
        return value
    return str(value)


class OpenpyxlInventoryExtractor:
    name = "openpyxl_inventory"

    def extract(self, workbook_path: Path, manifest: WorkbookManifest) -> WorkbookInventory:
        workbook_path = workbook_path.expanduser().resolve()
        wb = load_workbook(workbook_path, data_only=False, read_only=False, keep_vba=workbook_path.suffix.lower() == ".xlsm")
        cached_values = read_cached_values(workbook_path)
        data_tables = read_data_tables(workbook_path)
        form_controls = read_form_controls(workbook_path)
        sheets: list[SheetInventory] = []
        workbook_ranges: list[RangeInventory] = []
        unsupported = list(manifest.unsupported_features)
        recognized = 0
        formula_cells = 0
        cached_hits = 0
        try:
            for i, ws in enumerate(wb.worksheets):
                sheet = SheetInventory(
                    source_location=_loc(workbook_path, "sheet", ws.title, i),
                    name=ws.title,
                    index=i,
                    max_row=ws.max_row or 0,
                    max_column=ws.max_column or 0,
                    state=ws.sheet_state,
                )
                tables_by_corner = {spec.corner_cell: spec for spec in data_tables.get(ws.title, [])}
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        assert isinstance(cell, Cell)
                        # openpyxl models what-if data tables and array formulas as
                        # objects, not strings; str()-ing them would destroy the cell.
                        table_spec = tables_by_corner.get(cell.coordinate)
                        if table_spec is not None:
                            is_formula = True
                            formula_text = table_spec.formula_text
                        elif isinstance(cell.value, ArrayFormula):
                            is_formula = True
                            formula_text = cell.value.text
                        elif isinstance(cell.value, DataTableFormula):
                            is_formula = True
                            formula_text = None
                        else:
                            is_formula = isinstance(cell.value, str) and cell.value.startswith("=")
                            formula_text = str(cell.value) if is_formula else None
                        cache_key = (ws.title, cell.coordinate)
                        has_cached = is_formula and cache_key in cached_values
                        if is_formula:
                            formula_cells += 1
                            cached_hits += 1 if has_cached else 0
                        sheet.cells.append(
                            CellInventory(
                                source_location=_loc(workbook_path, "cell", ws.title, i, cell.coordinate, f"{ws.title}!{cell.coordinate}"),
                                address=cell.coordinate,
                                row=cell.row,
                                column=cell.column,
                                kind=CellKind.formula if is_formula else CellKind.literal,
                                value=None if is_formula else _safe_value(cell.value),
                                formula=formula_text,
                                data_type=cell.data_type,
                                number_format=cell.number_format,
                                style_id=getattr(cell, "style_id", None),
                                cached_value=cached_values.get(cache_key) if has_cached else None,
                                cached_value_available=has_cached,
                            )
                        )
                        recognized += 1
                for merged in ws.merged_cells.ranges:
                    sheet.ranges.append(RangeInventory(source_location=_loc(workbook_path, "merged_range", ws.title, i, str(merged), str(merged)), address=str(merged), kind="merged_range"))
                    recognized += 1
                for table in ws.tables.values():
                    totals_row_count = table.totalsRowCount
                    if totals_row_count is None:
                        totals_row_count = 1 if table.totalsRowShown is True else 0
                    table_columns = [column.name for column in table.tableColumns if column.name]
                    sheet.ranges.append(RangeInventory(
                        source_location=_loc(workbook_path, "table", ws.title, i, table.ref, table.name),
                        name=table.name,
                        address=table.ref,
                        kind="table",
                        metadata={"display_name": table.displayName, "columns": table_columns,
                                  "header_row_count": table.headerRowCount, "totals_row_count": totals_row_count},
                    ))
                    recognized += 1
                for spec in data_tables.get(ws.title, []):
                    sheet.ranges.append(
                        RangeInventory(
                            source_location=_loc(workbook_path, "data_table", ws.title, i, spec.ref or spec.corner_cell, f"data_table:{spec.corner_cell}"),
                            address=spec.ref or spec.corner_cell,
                            kind="data_table",
                            metadata={
                                "corner_cell": spec.corner_cell,
                                "formula": spec.formula_text,
                                "row_input_cell": spec.row_input_cell,
                                "col_input_cell": spec.col_input_cell,
                                "two_dimensional": spec.two_dimensional,
                                "raw_r1": spec.raw_r1,
                                "raw_r2": spec.raw_r2,
                                "dtr": spec.dtr,
                            },
                        )
                    )
                    recognized += 1
                for control in form_controls.get(ws.title, []):
                    sheet.ranges.append(
                        RangeInventory(
                            source_location=_loc(workbook_path, "form_control", ws.title, i, control.linked_address or control.control_id or "", f"control:{control.control_id}"),
                            # `address` stays a bare A1 like every other range object;
                            # the sheet-qualified reference lives in metadata.
                            address=control.linked_address or control.control_id or "",
                            kind="form_control",
                            metadata={
                                "control_id": control.control_id,
                                "control_type": control.control_type,
                                "linked_cell": control.linked_cell,
                                "linked_sheet": control.linked_sheet,
                                "list_fill_range": control.list_fill_range,
                                "macro": control.macro,
                            },
                        )
                    )
                    recognized += 1
                self._layout(ws, workbook_path, i, sheet)
                recognized += len(sheet.layout_objects)
                sheets.append(sheet)

            for name, defined_name in wb.defined_names.items():
                try:
                    destinations = list(defined_name.destinations)
                except Exception:
                    destinations = []
                if not destinations:
                    loc = _loc(workbook_path, "defined_name", object_id=name)
                    workbook_ranges.append(RangeInventory(source_location=loc, name=name, address=str(getattr(defined_name, "attr_text", name)), kind="defined_name", metadata={"scope": "workbook"}))
                    recognized += 1
                for sheet_name, address in destinations:
                    idx = wb.sheetnames.index(sheet_name) if sheet_name in wb.sheetnames else None
                    workbook_ranges.append(RangeInventory(source_location=_loc(workbook_path, "defined_name", sheet_name, idx, address, name), name=name, address=address, kind="defined_name", metadata={"scope": "workbook"}))
                    recognized += 1
            for ws in wb.worksheets:
                for name, defined_name in ws.defined_names.items():
                    try:
                        destinations = list(defined_name.destinations)
                    except Exception:
                        destinations = []
                    if not destinations:
                        address = str(getattr(defined_name, "attr_text", name))
                        destinations = [(ws.title, address)]
                    for destination_sheet, address in destinations:
                        idx = wb.sheetnames.index(destination_sheet) if destination_sheet in wb.sheetnames else None
                        workbook_ranges.append(
                            RangeInventory(
                                source_location=_loc(workbook_path, "defined_name", destination_sheet, idx, address, name),
                                name=name,
                                address=address,
                                kind="defined_name",
                                metadata={"scope": ws.title, "sheet": ws.title},
                            )
                        )
                        recognized += 1
        finally:
            wb.close()
        if formula_cells and cached_hits == 0:
            unsupported.append(
                UnsupportedFeature(
                    feature_type="missing_cached_values",
                    description=(
                        f"Workbook stores no cached formula results ({formula_cells} formula cells); "
                        "it was never recalculated by Excel, so cached-value reconciliation is unavailable."
                    ),
                    source_location=SourceLocation(workbook_path=str(workbook_path), object_type="workbook"),
                    severity=UnsupportedSeverity.warning,
                    opaque=False,
                    metadata={"formula_cells": formula_cells, "cells_with_cached_value": cached_hits},
                )
            )
        opaque_count = len(unsupported)
        coverage = CoverageSummary(recognized_inventory_objects=recognized, unsupported_or_opaque_objects=opaque_count, discovered_workbook_objects=recognized + opaque_count)
        return WorkbookInventory(workbook_sha256=manifest.sha256, sheets=sheets, workbook_ranges=workbook_ranges, unsupported_features=unsupported, coverage=coverage)

    def _layout(self, ws: Any, workbook_path: Path, index: int, sheet: SheetInventory) -> None:
        for row_idx, dim in ws.row_dimensions.items():
            if dim.hidden or dim.height:
                sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "row_layout", ws.title, index, str(row_idx), f"row:{row_idx}"), address=str(row_idx), kind="row_layout", metadata={"hidden": bool(dim.hidden), "height": dim.height}))
        for col, dim in ws.column_dimensions.items():
            if dim.hidden or dim.width:
                start = dim.min or column_index_from_string(col)
                end = dim.max or start
                first = get_column_letter(start)
                last = get_column_letter(end)
                address = f"{first}:{last}" if start != end else first
                sheet.layout_objects.append(
                    RangeInventory(
                        source_location=_loc(workbook_path, "column_layout", ws.title, index, address, f"col:{address}"),
                        address=address,
                        kind="column_layout",
                        metadata={"hidden": bool(dim.hidden), "width": dim.width, "min": start, "max": end},
                    )
                )
        if ws.freeze_panes:
            sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "freeze_panes", ws.title, index, str(ws.freeze_panes), "freeze_panes"), address=str(ws.freeze_panes), kind="freeze_panes"))
        for dv in getattr(ws.data_validations, "dataValidation", []):
            sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "data_validation", ws.title, index, str(dv.sqref), f"dv:{dv.sqref}"), address=str(dv.sqref), kind="data_validation", metadata={"type": dv.type, "formula1": dv.formula1, "formula2": dv.formula2}))
        for cf_range in getattr(ws.conditional_formatting, "_cf_rules", {}):
            sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "conditional_formatting", ws.title, index, str(cf_range), f"cf:{cf_range}"), address=str(cf_range), kind="conditional_formatting"))
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment:
                    sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "comment", ws.title, index, cell.coordinate, f"comment:{cell.coordinate}"), address=cell.coordinate, kind="comment", metadata={"text": cell.comment.text, "author": cell.comment.author}))
                if cell.hyperlink:
                    sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "hyperlink", ws.title, index, cell.coordinate, f"hyperlink:{cell.coordinate}"), address=cell.coordinate, kind="hyperlink", metadata={"target": cell.hyperlink.target, "location": cell.hyperlink.location}))
        if ws.protection and ws.protection.sheet:
            sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "sheet_protection", ws.title, index, object_id="sheet_protection"), address="sheet", kind="sheet_protection"))
