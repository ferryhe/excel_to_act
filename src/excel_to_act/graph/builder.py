"""Dependency graph builder that parses formula references without evaluation."""

from __future__ import annotations

import re

import networkx as nx

from excel_to_act.schemas import (
    FormulaGraph,
    GraphEdge,
    GraphNode,
    GraphNodeKind,
    SourceLocation,
    UnsupportedFeature,
    WorkbookInventory,
)

# Supports common A1 references and ranges, optionally sheet-qualified. Intentionally conservative.
REF_RE = re.compile(r"(?P<sheet>(?:'[^']+'|[A-Za-z_][\w .]*)!)?(?P<ref>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)")
EXTERNAL_RE = re.compile(r"\[[^\]]+\]")


def _node_id(kind: str, label: str) -> str:
    return f"{kind}:{label}"


def _clean_sheet(sheet: str | None, default: str | None) -> str | None:
    if not sheet:
        return default
    return sheet[:-1].strip("'")


class RegexFormulaGraphBuilder:
    name = "regex_formula_graph"

    def build(self, inventory: WorkbookInventory) -> FormulaGraph:
        graph = nx.DiGraph()
        nodes: dict[str, GraphNode] = {}
        edges: list[GraphEdge] = []
        unsupported: list[UnsupportedFeature] = []

        def add_node(node: GraphNode) -> None:
            nodes.setdefault(node.id, node)
            graph.add_node(node.id)

        for sheet in inventory.sheets:
            for cell in sheet.cells:
                if not cell.formula:
                    continue
                source_label = f"{sheet.name}!{cell.address}"
                source_id = _node_id("cell", source_label)
                add_node(GraphNode(id=source_id, kind=GraphNodeKind.cell, label=source_label, source_location=cell.source_location))
                formula = cell.formula
                if EXTERNAL_RE.search(formula):
                    loc = cell.source_location
                    target_id = _node_id("external", formula)
                    add_node(GraphNode(id=target_id, kind=GraphNodeKind.external, label=formula, source_location=loc))
                    edges.append(GraphEdge(source=source_id, target=target_id, formula=formula, source_location=loc))
                matches = list(REF_RE.finditer(formula))
                if not matches and any(ch.isalpha() for ch in formula):
                    unsupported.append(UnsupportedFeature(feature_type="formula_reference_parse", description=f"No references parsed from formula {formula!r}", source_location=cell.source_location, opaque=False))
                    target_id = _node_id("unsupported", source_label)
                    add_node(GraphNode(id=target_id, kind=GraphNodeKind.unsupported, label=formula, source_location=cell.source_location))
                    edges.append(GraphEdge(source=source_id, target=target_id, formula=formula, source_location=cell.source_location))
                for match in matches:
                    ref = match.group("ref").replace("$", "")
                    ref_sheet = _clean_sheet(match.group("sheet"), sheet.name)
                    label = f"{ref_sheet}!{ref}" if ref_sheet else ref
                    kind = GraphNodeKind.range if ":" in ref else GraphNodeKind.cell
                    target_id = _node_id(kind.value, label)
                    add_node(GraphNode(id=target_id, kind=kind, label=label, source_location=SourceLocation(workbook_path=cell.source_location.workbook_path, sheet_name=ref_sheet, address=ref, object_type=kind.value, object_id=label)))
                    graph.add_edge(source_id, target_id)
                    edges.append(GraphEdge(source=source_id, target=target_id, formula=formula, source_location=cell.source_location))
        return FormulaGraph(nodes=list(nodes.values()), edges=edges, unsupported_features=unsupported)
