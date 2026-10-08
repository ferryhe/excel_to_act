"""Core inventory extraction from openpyxl with source locations."""

from __future__ import annotations

from collections import defaultdict, deque
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from excel_to_act.ingest.cached_values import read_cached_values
from excel_to_act.ingest.data_table import PACKAGE_READ_ERRORS, read_data_tables
from excel_to_act.ingest.form_controls import read_form_controls
from excel_to_act.steps.step1.source_scan import scan_step1_source
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


def _identity_key(kind: str, value: str | None) -> str | None:
    return " ".join(sorted(value.split())) if value and kind in {"conditional_formatting", "data_validation"} else value


def _attach_source_identities(inventory: WorkbookInventory, scan: dict[str, Any]) -> int:
    """Mark only records actually extracted from a matching source declaration."""

    sheets = {sheet.name: sheet for sheet in inventory.sheets}
    cells = {(sheet.name, cell.address): cell for sheet in inventory.sheets for cell in sheet.cells}
    ranges: dict[tuple[str, str | None, str | None], deque[RangeInventory]] = defaultdict(deque)
    for sheet in inventory.sheets:
        for item in (*sheet.ranges, *sheet.layout_objects):
            key = (item.name if item.kind == "table" else
                   item.metadata.get("corner_cell") if item.kind == "data_table" else
                   item.metadata.get("control_id") if item.kind == "form_control" else
                   "" if item.kind == "sheet_protection" else item.address)
            ranges[(item.kind, sheet.name, _identity_key(item.kind, key))].append(item)
    for item in inventory.workbook_ranges:
        owner = item.metadata.get("scope", "workbook") if item.kind == "defined_name" else item.source_location.sheet_name
        key = item.name if item.kind == "defined_name" else item.address
        ranges[(item.kind, owner, _identity_key(item.kind, key))].append(item)
    matched: set[str] = set()
    for source in scan["objects"]:
        kind = source["kind"]
        details = source["details"]
        sheet_name = details.get("sheet") or (details.get("name") if kind == "sheet" else None)
        record = None
        if kind == "sheet":
            record = sheets.get(sheet_name)
        elif kind == "cell":
            record = cells.get((sheet_name, details.get("address")))
        else:
            owner = details.get("scope", "workbook") if kind == "defined_name" else sheet_name
            key = (details.get("name") if kind in {"defined_name", "table"} else
                   details.get("corner_cell") if kind == "data_table" else
                   details.get("control_id") if kind == "form_control" else
                   "" if kind == "sheet_protection" else details.get("address"))
            bucket = ranges.get((kind, owner, _identity_key(kind, key)))
            record = bucket.popleft() if bucket else None
        if record is None:
            continue
        record.source_identity = source["identity"]
        record.source_location.source_identity = source["identity"]
        record.source_location.ooxml_part = source["part"]
        matched.add(source["identity"])
    return len(matched)


class OpenpyxlInventoryExtractor:
    name = "openpyxl_inventory"

    def extract(self, workbook_path: Path, manifest: WorkbookManifest) -> WorkbookInventory:
        workbook_path = workbook_path.expanduser().resolve()
        try:
            wb = load_workbook(workbook_path, data_only=False, read_only=False, keep_vba=workbook_path.suffix.lower() == ".xlsm")
        except Exception as exc:
            raise InvalidFileException(
                f"openpyxl could not read workbook: {type(exc).__name__}: {exc}"
            ) from exc
        try:
            cached_values = read_cached_values(workbook_path)
            try:
                data_tables = read_data_tables(workbook_path)
            except PACKAGE_READ_ERRORS as exc:
                raise InvalidFileException(
                    f"Could not read workbook data-table package parts: {type(exc).__name__}: {exc}"
                ) from exc
            try:
                form_controls = read_form_controls(workbook_path)
            except PACKAGE_READ_ERRORS as exc:
                raise InvalidFileException(
                    f"Could not read workbook form-control package parts: {type(exc).__name__}: {exc}"
                ) from exc
            sheets: list[SheetInventory] = []
            workbook_ranges: list[RangeInventory] = []
            unsupported = list(manifest.unsupported_features)
            recognized = 0
            formula_cells = 0
            cached_hits = 0
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
                    table_columns = [column.name for column in table.tableColumns]
                    if not all(isinstance(name, str) and name for name in table_columns):
                        table_columns = []
                    totals_row_count = table.totalsRowCount
                    if totals_row_count is None:
                        if table.totalsRowShown is True:
                            totals_row_count = 1
                        elif table.totalsRowShown in {None, False}:
                            totals_row_count = 0
                    sheet.ranges.append(RangeInventory(
                        source_location=_loc(workbook_path, "table", ws.title, i, table.ref, table.name),
                        name=table.name,
                        address=table.ref,
                        kind="table",
                        metadata={
                            "display_name": table.displayName,
                            "header_row_count": table.headerRowCount,
                            "totals_row_count": totals_row_count,
                            "totals_row_shown": table.totalsRowShown,
                            "columns": table_columns,
                        },
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
            print_areas = {ws.title: str(ws.print_area) for ws in wb.worksheets if ws.print_area}
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
        inventory = WorkbookInventory(
            workbook_sha256=manifest.sha256,
            sheets=sheets,
            workbook_ranges=workbook_ranges,
            unsupported_features=unsupported,
            coverage=CoverageSummary(
                recognized_inventory_objects=recognized,
                unsupported_or_opaque_objects=0,
                discovered_workbook_objects=recognized,
            ),
        )
        try:
            scan = scan_step1_source(workbook_path)
        except Exception as exc:
            raise InvalidFileException(f"Could not scan source XML: {type(exc).__name__}: {exc}") from exc
        for source in scan["objects"]:
            details = source["details"]
            if source["kind"] != "defined_name" or details.get("name") != "_xlnm.Print_Area":
                continue
            scope = details.get("scope")
            area = print_areas.get(scope)
            if area and not any(item.kind == "defined_name" and item.name == "_xlnm.Print_Area" and item.metadata.get("scope") == scope for item in inventory.workbook_ranges):
                inventory.workbook_ranges.append(RangeInventory(
                    source_location=_loc(workbook_path, "defined_name", scope, next(sheet.index for sheet in sheets if sheet.name == scope), area, "_xlnm.Print_Area"),
                    name="_xlnm.Print_Area", address=area, kind="defined_name",
                    metadata={"scope": scope, "sheet": scope},
                ))
        inventory.coverage = CoverageSummary(
            recognized_inventory_objects=_attach_source_identities(inventory, scan),
            unsupported_or_opaque_objects=0,
            discovered_workbook_objects=len(scan["objects"]),
        )
        return inventory

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
            address = str(cf_range.sqref)
            sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "conditional_formatting", ws.title, index, address, f"cf:{address}"), address=address, kind="conditional_formatting"))
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment:
                    sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "comment", ws.title, index, cell.coordinate, f"comment:{cell.coordinate}"), address=cell.coordinate, kind="comment", metadata={"text": cell.comment.text, "author": cell.comment.author}))
                if cell.hyperlink:
                    sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "hyperlink", ws.title, index, cell.coordinate, f"hyperlink:{cell.coordinate}"), address=cell.coordinate, kind="hyperlink", metadata={"target": cell.hyperlink.target, "location": cell.hyperlink.location}))
        if ws.protection and ws.protection.sheet:
            sheet.layout_objects.append(RangeInventory(source_location=_loc(workbook_path, "sheet_protection", ws.title, index, object_id="sheet_protection"), address="sheet", kind="sheet_protection"))
