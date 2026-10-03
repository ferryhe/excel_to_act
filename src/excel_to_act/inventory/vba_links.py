"""Turn VBA source into cell/name references and graph edges.

VBA drives calculations through ``Range("B7")``, ``Names("Mort_qx")``, and
``Cells(r, c)``. Those references are invisible to the formula tokenizer, so a
workbook that depends on VBA would yield a fragmented dependency graph without
this step.

Literal matching is inherently lossy (``"B" & i``, indirect addressing, variables
holding addresses). Every reference therefore carries a ``confidence`` and an
``unresolved`` flag rather than pretending to be certain.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from excel_to_act.schemas import GraphEdge, GraphNode, GraphNodeKind, SourceLocation

# `(?<!\w)` keeps matches for `ws.Range(...)` and `.Range(...)` while rejecting
# `MyRange("A1")`. `Cells(...)` allows one level of nesting (e.g. Cells(1, UBound(a))).
_RANGE_RE = re.compile(r'(?<!\w)Range\s*\(\s*"([^"]+)"\s*\)')
_NAMES_RE = re.compile(r'(?<!\w)Names\s*\(\s*"([^"]+)"\s*\)')
_CELLS_RE = re.compile(r"(?<!\w)Cells\s*\((?:[^()]|\([^()]*\))*\)")
_A1_QUALIFIED = re.compile(r"^'?(?P<sheet>[^'!]+)'?!\s*(?P<ref>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$")
_A1_BARE = re.compile(r"^(?P<ref>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$")


@dataclass(frozen=True)
class VbaCellRef:
    module: str
    raw: str
    kind: str  # "range" | "cells" | "name"
    sheet: str | None
    address: str | None
    name: str | None
    confidence: float
    unresolved: bool


def _parse_range_literal(text: str) -> tuple[str | None, str | None]:
    text = text.replace("$", "")
    match = _A1_QUALIFIED.match(text) if "!" in text else _A1_BARE.match(text)
    if not match:
        return None, None
    return match.groupdict().get("sheet"), match.group("ref")


def _strip_vba_noise(code: str) -> str:
    """Remove comments and mask strings except literal Range/Names arguments."""

    chars = list(code)
    line_start = 0
    statement_start = True
    i = 0

    def blank(start: int, end: int) -> None:
        for index in range(start, end):
            if chars[index] not in "\r\n":
                chars[index] = " "

    while i < len(code):
        char = code[i]
        if char in "\r\n":
            line_start = i + 1
            statement_start = True
            i += 1
            continue
        if char == "'":
            end = i
            while end < len(code) and code[end] not in "\r\n":
                end += 1
            blank(i, end)
            i = end
            continue
        if statement_start and re.match(r"Rem\b", code[i:], re.IGNORECASE):
            end = i
            while end < len(code) and code[end] not in "\r\n":
                end += 1
            blank(i, end)
            i = end
            continue
        if char == '"':
            end = i + 1
            while end < len(code):
                if code[end] == '"':
                    if end + 1 < len(code) and code[end + 1] == '"':
                        end += 2
                        continue
                    end += 1
                    break
                end += 1
            prefix = code[line_start:i]
            suffix = code[end:]
            is_reference_argument = bool(
                re.search(r"\b(?:Range|Names)\s*\(\s*$", prefix, re.IGNORECASE)
                and re.match(r"\s*\)", suffix)
            )
            if not is_reference_argument:
                blank(i, end)
            statement_start = False
            i = end
            continue
        if char == ":":
            statement_start = True
        elif not char.isspace():
            statement_start = False
        i += 1
    return "".join(chars)


def extract_vba_cell_links(project) -> list[VbaCellRef]:
    """Scan every module's source for cell/name references."""

    refs: list[VbaCellRef] = []
    for module in getattr(project, "modules", []):
        code = _strip_vba_noise(module.code or "")
        for match in _RANGE_RE.finditer(code):
            sheet, address = _parse_range_literal(match.group(1))
            if address and sheet:
                confidence, unresolved = 0.9, False
            elif address:
                confidence, unresolved = 0.6, False
            else:
                confidence, unresolved = 0.3, True
            refs.append(
                VbaCellRef(module.name, match.group(0), "range", sheet, address, None, confidence, unresolved)
            )
        for match in _NAMES_RE.finditer(code):
            refs.append(
                VbaCellRef(module.name, match.group(0), "name", None, None, match.group(1), 0.4, True)
            )
        for match in _CELLS_RE.finditer(code):
            # Cells(r, c) cannot be resolved without its worksheet context.
            refs.append(
                VbaCellRef(module.name, match.group(0), "cells", None, None, None, 0.3, True)
            )
    return refs


def build_vba_edges(
    refs: list[VbaCellRef],
    workbook_path: str | None = None,
    known_sheets: set[str] | None = None,
) -> tuple[list[GraphNode], list[GraphEdge]]:
    """Materialize VBA references into graph nodes/edges sharing the formula graph's id space.

    Edges use ``relationship="vba_ref"`` and land in the same ``FormulaGraph``, so
    downstream reconciliation sees one node space. Target ids follow ``builder.py``
    exactly: ``cell:<sheet>!<A1>`` for a single cell, ``range:<sheet>!<A1:A2>`` for a
    multi-cell reference (the builder keys off ``":"``), and ``name:<name>``.

    An unqualified ``Range("B7")`` has no worksheet context. It is only resolved
    when the workbook has exactly one sheet; otherwise no edge is produced, because
    inventing ``cell:B7`` would create a second, unreachable node namespace.
    Unresolved references are still returned by ``extract_vba_cell_links`` so the
    completeness report can surface them.
    """

    nodes: list[GraphNode] = []
    edges: list[GraphEdge] = []
    seen: set[str] = set()

    def add_node(node: GraphNode) -> None:
        if node.id not in seen:
            seen.add(node.id)
            nodes.append(node)

    for ref in refs:
        source_id = f"vba:{ref.module}"
        add_node(GraphNode(id=source_id, kind=GraphNodeKind.vba, label=ref.module))
        if ref.kind == "name" and ref.name:
            target_id = f"name:{ref.name}"
            add_node(GraphNode(id=target_id, kind=GraphNodeKind.name, label=ref.name))
            edges.append(
                GraphEdge(source=source_id, target=target_id, relationship="vba_ref", formula=ref.raw, confidence=ref.confidence)
            )
        elif ref.address:
            sheet = ref.sheet
            if not sheet and known_sheets and len(known_sheets) == 1:
                sheet = next(iter(known_sheets))
            if not sheet:
                continue
            label = f"{sheet}!{ref.address}"
            kind = GraphNodeKind.range if ":" in ref.address else GraphNodeKind.cell
            target_id = f"{kind.value}:{label}"
            add_node(
                GraphNode(
                    id=target_id,
                    kind=kind,
                    label=label,
                    source_location=SourceLocation(
                        workbook_path=workbook_path,
                        sheet_name=sheet,
                        address=ref.address,
                        object_type="cell",
                        object_id=label,
                    ),
                )
            )
            edges.append(
                GraphEdge(source=source_id, target=target_id, relationship="vba_ref", formula=ref.raw, confidence=ref.confidence)
            )
    return nodes, edges
