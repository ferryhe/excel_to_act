"""Dependency graph builder that parses formula references without evaluation."""

from __future__ import annotations

import re

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils.cell import column_index_from_string

from excel_to_act.schemas import (
    FormulaGraph,
    GraphEdge,
    GraphNode,
    GraphNodeKind,
    RangeInventory,
    SourceLocation,
    UnsupportedFeature,
    WorkbookInventory,
)

_CELL = r"\$?[A-Z]{1,3}\$?[1-9]\d*"
_CELL_REF_RE = re.compile(rf"^(?P<first>{_CELL})(?::(?P<last>{_CELL}))?$", re.IGNORECASE)
_CELL_ADDRESS_RE = re.compile(r"^\$?(?P<column>[A-Z]{1,3})\$?(?P<row>[1-9]\d*)$", re.IGNORECASE)
_TABLE_REF_RE = re.compile(r"^(?P<table>[A-Za-z_][A-Za-z0-9_.]*)\[(?P<selector>.+)\]$")
# ponytail: static A1, defined-name, and table-column refs only; add operand handlers as coverage needs grow.
_DYNAMIC_REFERENCE_FUNCTIONS = {"INDIRECT", "OFFSET"}


def _node_id(kind: str, label: str) -> str:
    return f"{kind}:{label}"


def _unquote_sheet(sheet: str) -> str:
    sheet = sheet.strip()
    if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
        return sheet[1:-1].replace("''", "'")
    return sheet


def _formula_parts(value: str) -> tuple[str | None, str]:
    if "!" not in value:
        return None, value
    sheet, reference = value.rsplit("!", 1)
    return _unquote_sheet(sheet), reference


def _valid_a1_reference(value: str) -> bool:
    match = _CELL_REF_RE.fullmatch(value)
    if not match:
        return False
    for address in (match.group("first"), match.group("last")):
        if address is None:
            continue
        cell = _CELL_ADDRESS_RE.fullmatch(address)
        if not cell:
            return False
        try:
            column = column_index_from_string(cell.group("column"))
        except ValueError:
            return False
        row = cell.group("row")
        if column > 16_384 or len(row) > 7 or (len(row) == 7 and row > "1048576"):
            return False
    return True


def _name_node(definition: RangeInventory, scope: str) -> GraphNode:
    name = definition.name or ""
    label = f"{scope}!{name}" if scope != "workbook" else name
    return GraphNode(
        id=_node_id(GraphNodeKind.name.value, label),
        kind=GraphNodeKind.name,
        label=label,
        source_location=definition.source_location,
        metadata={"name": name, "scope": scope, "address": definition.address},
    )


class RegexFormulaGraphBuilder:
    """Keep the Phase 1 builder name while using tokenizer operands internally."""

    name = "regex_formula_graph"

    def build(self, inventory: WorkbookInventory) -> FormulaGraph:
        nodes: dict[str, GraphNode] = {}
        edges: list[GraphEdge] = []
        unsupported: list[UnsupportedFeature] = []

        def add_node(node: GraphNode) -> GraphNode:
            return nodes.setdefault(node.id, node)

        sheets = {sheet.name.casefold(): sheet.name for sheet in inventory.sheets}
        tables = {
            item.name.casefold(): item
            for sheet in inventory.sheets
            for item in sheet.ranges
            if item.kind == "table" and item.name
        }
        names = [item for item in inventory.workbook_ranges if item.kind == "defined_name" and item.name]

        def diagnostic(cell, description: str) -> None:
            unsupported.append(
                UnsupportedFeature(
                    feature_type="formula_reference_unresolved",
                    description=description,
                    source_location=cell.source_location,
                    opaque=False,
                )
            )

        def name_reference(name: str, sheet_name: str) -> GraphNode | None:
            candidates = [item for item in names if (item.name or "").casefold() == name.casefold()]
            local = next(
                (
                    item
                    for item in candidates
                    if str(item.metadata.get("scope", "workbook")).casefold() == sheet_name.casefold()
                ),
                None,
            )
            definition = local or next(
                (item for item in candidates if item.metadata.get("scope", "workbook") == "workbook"),
                None,
            )
            if definition is None:
                return None
            scope = str(definition.metadata.get("scope", "workbook"))
            return _name_node(definition, scope)

        def table_reference(operand: str) -> GraphNode | None:
            match = _TABLE_REF_RE.fullmatch(operand)
            if not match:
                return None
            table = tables.get(match.group("table").casefold())
            if table is None:
                return None
            selector = match.group("selector")
            selector_key = selector.casefold()
            item_selector = {"#all": "#All", "#data": "#Data"}.get(selector_key)
            current_row = selector.startswith("@")
            column = (
                selector[1:]
                if current_row
                else selector
                if item_selector is None and "[" not in selector and "]" not in selector
                else None
            )
            if current_row and not column:
                return None
            headers = table.metadata.get("columns")
            if not headers:
                # Header text is field evidence, not authoritative table-column metadata.
                headers = []
            if column and headers and column.casefold() not in {str(header).casefold() for header in headers}:
                return None
            if column and headers:
                column = next(str(header) for header in headers if str(header).casefold() == column.casefold())
            label_selector = item_selector or (f"@{column}" if current_row else column or selector)
            label = f"{table.name}[{label_selector}]"
            return GraphNode(
                id=_node_id(GraphNodeKind.range.value, label),
                kind=GraphNodeKind.range,
                label=label,
                source_location=table.source_location,
                metadata={
                    "structured_reference": operand,
                    "current_row": current_row,
                    "table": table.name,
                    "table_ref": table.address,
                    "sheet": table.source_location.sheet_name,
                    "column": column,
                    "table_columns": headers or [],
                },
            )

        def reference_node(operand: str, current_sheet: str, workbook_path: str | None) -> GraphNode | None:
            sheet_name, reference = _formula_parts(operand)
            if sheet_name is not None:
                external_match = re.search(r"\[(?P<workbook>[^\]]+)\](?P<sheet>.*)$", sheet_name)
                if external_match:
                    workbook = external_match.group("workbook")
                    external_path = sheet_name[:external_match.start()]
                    external_sheet = external_match.group("sheet")
                    if not workbook or not external_sheet or not _valid_a1_reference(reference):
                        return None
                    normalized = reference.replace("$", "").upper()
                    label = f"{external_path}[{workbook}]{external_sheet}!{normalized}"
                    return GraphNode(
                        id=_node_id(GraphNodeKind.external.value, label),
                        kind=GraphNodeKind.external,
                        label=label,
                        metadata={
                            "workbook": workbook,
                            "path": external_path,
                            "sheet": external_sheet,
                            "address": normalized,
                        },
                    )
            if not _CELL_REF_RE.fullmatch(reference):
                if sheet_name is not None and sheet_name.casefold() not in sheets:
                    return None
                scope_sheet = sheets.get(sheet_name.casefold(), sheet_name) if sheet_name else current_sheet
                return name_reference(reference, scope_sheet)
            if not _valid_a1_reference(reference):
                return None
            actual_sheet = sheets.get(sheet_name.casefold()) if sheet_name else current_sheet
            if actual_sheet is None:
                return None
            normalized = reference.replace("$", "").upper()
            label = f"{actual_sheet}!{normalized}"
            kind = GraphNodeKind.range if ":" in normalized else GraphNodeKind.cell
            return GraphNode(
                id=_node_id(kind.value, label),
                kind=kind,
                label=label,
                source_location=SourceLocation(
                    workbook_path=workbook_path,
                    sheet_name=actual_sheet,
                    address=normalized,
                    object_type=kind.value,
                    object_id=label,
                ),
            )

        for sheet in inventory.sheets:
            for cell in sheet.cells:
                if not cell.formula:
                    continue
                source_label = f"{sheet.name}!{cell.address}"
                source_id = _node_id("cell", source_label)
                source_node = add_node(GraphNode(
                    id=source_id,
                    kind=GraphNodeKind.cell,
                    label=source_label,
                    source_location=cell.source_location,
                    metadata={"references_resolved": True},
                ))
                source_node.source_location = cell.source_location
                source_node.metadata["references_resolved"] = True
                try:
                    tokens = Tokenizer(cell.formula).items
                except (TokenizerError, IndexError) as exc:
                    source_node.metadata["references_resolved"] = False
                    diagnostic(cell, f"Could not tokenize formula {cell.formula!r}: {exc}")
                    continue
                if not tokens:
                    source_node.metadata["references_resolved"] = False
                    diagnostic(cell, f"Formula {cell.formula!r} is empty and has no operands")
                    continue

                balance = 0
                for token in tokens:
                    if token.subtype == "OPEN":
                        balance += 1
                    elif token.subtype == "CLOSE":
                        balance -= 1
                    if balance < 0:
                        break
                if balance:
                    source_node.metadata["references_resolved"] = False
                    diagnostic(cell, f"Unbalanced parentheses in formula {cell.formula!r}")
                    continue
                if tokens[-1].type == "OPERATOR-INFIX":
                    source_node.metadata["references_resolved"] = False
                    diagnostic(
                        cell,
                        f"Formula {cell.formula!r} ends with an infix operator {tokens[-1].value!r}",
                    )
                    continue

                unresolved = False
                for token in tokens:
                    if token.type == "OPERAND" and token.subtype == "ERROR" and token.value.upper() == "#REF!":
                        unresolved = True
                        diagnostic(cell, f"Unresolved #REF! reference in formula {cell.formula!r}")
                        continue
                    if token.type == "FUNC" and token.subtype == "OPEN":
                        function = token.value[:-1].rsplit(".", 1)[-1].upper()
                        if function in _DYNAMIC_REFERENCE_FUNCTIONS:
                            unresolved = True
                            diagnostic(
                                cell,
                                f"Dynamic reference from {function} cannot be statically resolved in {cell.formula!r}",
                            )
                    if token.type != "OPERAND" or token.subtype != "RANGE":
                        continue
                    target = table_reference(token.value) or reference_node(
                        token.value, sheet.name, cell.source_location.workbook_path
                    )
                    if target is None:
                        unresolved = True
                        diagnostic(
                            cell,
                            f"Could not resolve reference operand {token.value!r} in formula {cell.formula!r}",
                        )
                        continue
                    add_node(target)
                    edges.append(
                        GraphEdge(
                            source=source_id,
                            target=target.id,
                            formula=cell.formula,
                            source_location=cell.source_location,
                        )
                    )
                source_node.metadata["references_resolved"] = not unresolved

        return FormulaGraph(nodes=list(nodes.values()), edges=edges, unsupported_features=unsupported)
