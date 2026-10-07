"""Bounded static traversal of prepared formula edges and structural destinations."""

from __future__ import annotations

from collections import defaultdict, deque
import re
from pathlib import Path
from typing import Any

from excel_to_act.schemas import FormulaGraph, GraphNode, GraphNodeKind, SourceLocation
from excel_to_act.steps.step2.prepare import _destination, _parse_destination
from excel_to_act.steps.step2.query import (
    QueryFailure, _contains, _finish_packet, _is_excluded, _json_bytes, _load_manifest,
    _load_views, _normalize_address, _read_ref, _selector, _sha, _sheet_ref, _source,
    _scope_guard,
)
from excel_to_act.views import _tokens, validate_agent_output


def trace(
    manifest_path: Path, source_id: str, kind: str, target: str, *, sheet: str | None = None,
    direction: str, max_depth: int = 8, max_nodes: int = 100, max_edges: int = 200,
) -> dict[str, Any]:
    """Read prepared evidence only; upstream follows formula-cell -> input edges."""
    if not source_id:
        raise QueryFailure("unsupported", "source_id_required", "--source-id is required.")
    if kind not in {"cell", "range", "name"} or direction not in {"upstream", "downstream", "both"}:
        raise QueryFailure("unsupported", "trace_selector_invalid", "Trace needs kind cell/range/name and direction upstream/downstream/both.")
    limits = {"max_depth": max_depth, "max_nodes": max_nodes, "max_edges": max_edges}
    ceilings = {"max_depth": 100, "max_nodes": 10_000, "max_edges": 20_000}
    for key, value in limits.items():
        if type(value) is not int or value < (0 if key == "max_depth" else 1) or value > ceilings[key]:
            raise QueryFailure("unsupported", "trace_limit_invalid", f"{key} is outside its supported bound.", limits=limits, ceilings=ceilings)
    selector, bounds = _selector(kind, target, sheet, None)
    manifest_bytes, manifest, _, roots = _load_manifest(manifest_path)
    source = _source(manifest, source_id)
    graph_ref = source.dependency_graph
    if graph_ref is None:
        raise QueryFailure("unavailable", "trace_graph_unavailable", "Source has no prepared dependency graph.")
    graph_bytes = _read_ref(graph_ref, roots, manifest)
    try:
        graph = FormulaGraph.model_validate_json(graph_bytes)
    except ValueError as exc:
        raise QueryFailure("integrity_failed", "trace_graph_invalid", f"Prepared dependency graph is invalid: {exc}") from exc
    nodes = {node.id: node for node in graph.nodes}
    if len(nodes) != len(graph.nodes) or any(edge.source not in nodes or edge.target not in nodes for edge in graph.edges):
        raise QueryFailure("integrity_failed", "trace_graph_invalid", "Graph has duplicate identities or missing edge endpoints.")
    node_proofs = {node.id: {"file_ref": graph_ref.model_dump(mode="json"), "pointer": f"/nodes/{i}"}
                   for i, node in enumerate(graph.nodes)}
    loaded: dict[str, Any] = {}
    records: dict[str, Any] = {}
    selected_records: dict[tuple[str, str], Any] = {}

    def sheet_view(name: str):
        ref = _sheet_ref(source, name)
        if ref is None:
            return None
        if ref.path not in loaded:
            entry = _load_views([ref], source, manifest, roots)[0]
            loaded[ref.path] = entry
            for record in entry[1].records:
                if record.record_type == "cell" and record.source_location.address:
                    address = _normalize_address(record.source_location.address, cell=True)[0]
                    records[f"cell:{entry[1].sheet_name}!{address}"] = (entry, record)
        return loaded[ref.path]

    def actual_sheet(name: str | None) -> str | None:
        return next((value for value in source.lookup.sheet_views if name and value.casefold() == name.casefold()), None)

    def add_address(owner: str, address: str, proof: dict[str, Any]) -> str:
        normalized, _ = _normalize_address(address)
        node_kind = "range" if ":" in normalized else "cell"
        label = f"{owner}!{normalized}"
        node_id = f"{node_kind}:{label}"
        if node_id not in nodes:
            nodes[node_id] = GraphNode(id=node_id, kind=GraphNodeKind(node_kind), label=label,
                source_location=SourceLocation(workbook_path=source.source_path, sheet_name=owner,
                    address=normalized, object_type=node_kind, object_id=label))
            node_proofs[node_id] = proof
        return node_id

    name_groups: dict[str, list[Any]] = defaultdict(list)
    name_problems: dict[str, str] = {}
    destinations: dict[str, dict[str, Any]] = {}
    for i, item in enumerate(source.lookup.defined_names):
        node_id = item.node_id or f"name:{item.scope}!{item.name}"
        name_groups[node_id].append(item)
        if node_id not in nodes:
            nodes[node_id] = GraphNode(id=node_id, kind=GraphNodeKind.name, label=item.name,
                source_location=item.source_location,
                metadata={"name": item.name, "scope": item.scope, "address": item.address})
        node_proofs[node_id] = {"manifest_sha256": _sha(manifest_bytes), "pointer": f"/sources/{manifest.sources.index(source)}/lookup/defined_names/{i}"}
    for node_id, declarations in name_groups.items():
        if len(declarations) != 1:
            name_problems[node_id] = "ambiguous_or_multi_area_name"
            continue
        item = declarations[0]
        owner, address = _parse_destination(item.address, item.source_location.sheet_name)
        owner = actual_sheet(owner)
        if not owner or not address:
            if re.fullmatch(r'[-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?|"(?:[^"]|"")*"|TRUE|FALSE', item.address.strip().lstrip("="), re.IGNORECASE):
                name_problems[node_id] = "constant_name"
            else:
                text = item.address.strip()
                name_problems[node_id] = (
                    "external_name" if "[" in text else "dynamic_name" if "(" in text
                    else "unsupported_multi_area_name" if "," in text
                    else "relative_or_unsupported_name"
                )
            continue
        destination = add_address(owner, address, node_proofs[node_id])
        destinations[node_id] = {"source": node_id, "target": destination,
            "relationship": "name_destination", "label": "structural_derivation",
            "provenance": node_proofs[node_id]}

    # Resolve supported table selectors with #30's resolver. Current-row selectors need
    # caller context; retain that explicit boundary rather than invent a local edge.
    for node in list(nodes.values()):
        if node.kind != GraphNodeKind.range or not node.metadata.get("structured_reference"):
            continue
        owner, address = _destination(node)
        owner = actual_sheet(owner)
        if owner and address:
            destination = add_address(owner, address, node_proofs[node.id])
            destinations[node.id] = {"source": node.id, "target": destination,
                "relationship": "range_destination", "label": "structural_derivation",
                "provenance": node_proofs[node.id]}

    if kind == "name":
        candidates = [(node_id, item) for node_id, items in name_groups.items() for item in items
                      if item.name.casefold() == selector["target"]]
        if selector["sheet"]:
            local = [row for row in candidates if row[1].scope.casefold() == selector["sheet"]]
            candidates = local or [row for row in candidates if row[1].scope.casefold() == "workbook"]
        if not candidates:
            raise QueryFailure("not_found", "name_not_found", "No matching defined name exists in the source.")
        root_ids = list(dict.fromkeys(row[0] for row in candidates))
        ambiguous = len(candidates) > 1
        if not ambiguous and root_ids[0] in destinations:
            destination = nodes[destinations[root_ids[0]]["target"]].source_location
            _scope_guard(source, destination.sheet_name, _normalize_address(destination.address)[1], selector)
    else:
        owner = actual_sheet(selector["sheet"])
        if owner is None or bounds is None:
            raise QueryFailure("unavailable", "sheet_unavailable", "Trace needs an available --sheet or Sheet!A1 target.")
        _scope_guard(source, owner, bounds, selector)
        entry = sheet_view(owner)
        root_id = add_address(owner, selector["target"], {"selector": selector,
            "manifest_sha256": _sha(manifest_bytes)})
        if entry is None or kind == "cell" and root_id not in records:
            raise QueryFailure("not_found", "cell_missing", "Requested target has no stored canonical cell record.")
        root_ids, ambiguous = [root_id], False

    outgoing: dict[str, list[dict[str, Any]]] = defaultdict(list)
    incoming: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for i, edge in enumerate(graph.edges):
        loc = nodes[edge.source].source_location
        if _is_excluded(source, loc.sheet_name if loc else None):
            continue
        relation = {"source": edge.source, "target": edge.target, "relationship": edge.relationship,
            "label": "source_formula_edge", "provenance": {"file_ref": graph_ref.model_dump(mode="json"), "pointer": f"/edges/{i}"}}
        outgoing[edge.source].append(relation)
        incoming[edge.target].append(relation)
    for relation in destinations.values():
        outgoing[relation["source"]].append(relation)
        incoming[relation["target"]].append(relation)
    # ponytail: scan prepared range nodes for reverse membership; index by sheet if
    # repeated bounded reads become slow. No blank worksheet cells are generated.
    ranges = [node for node in nodes.values() if node.kind == GraphNodeKind.range and not node.metadata.get("structured_reference")]

    def membership(range_node: GraphNode, cell_id: str, entry, record) -> dict[str, Any]:
        return {"source": range_node.id, "target": cell_id, "relationship": "stored_range_member",
            "label": "structural_derivation", "provenance": {"range": node_proofs[range_node.id],
                "canonical_view_ref": entry[0].model_dump(mode="json"), "view_id": entry[1].view_id, "record_id": record.record_id}}

    def neighbors(node: GraphNode):
        if direction in {"upstream", "both"}:
            yield from outgoing[node.id]
            if node.kind == GraphNodeKind.range and node in ranges and node.source_location:
                owner = node.source_location.sheet_name
                entry = sheet_view(owner) if owner else None
                if entry:
                    box = _normalize_address(node.source_location.address or "")[1]
                    for record in entry[1].records:
                        if record.record_type != "cell" or not record.source_location.address:
                            continue
                        address, cell_box = _normalize_address(record.source_location.address, cell=True)
                        if _contains(box, cell_box):
                            cell_id = add_address(owner, address, {"canonical_view_ref": entry[0].model_dump(mode="json"), "record_id": record.record_id})
                            yield membership(node, cell_id, entry, record)
        if direction in {"downstream", "both"}:
            yield from incoming[node.id]
            if node.kind == GraphNodeKind.cell and node.source_location:
                owner, address = node.source_location.sheet_name, node.source_location.address
                if owner and address:
                    sheet_view(owner)
                    entry_record = records.get(node.id)
                    if entry_record:
                        box = _normalize_address(address, cell=True)[1]
                        for range_node in ranges:
                            loc = range_node.source_location
                            if loc and loc.sheet_name == owner and loc.address and _contains(_normalize_address(loc.address)[1], box):
                                yield membership(range_node, node.id, *entry_record)

    visited: dict[str, int] = {}
    edges: list[dict[str, Any]] = []
    edge_keys: set[bytes] = set()
    frontier: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []
    boundaries: list[dict[str, Any]] = []
    terminals: list[dict[str, Any]] = []
    delivered_nodes: list[dict[str, Any]] = []
    queue: deque[tuple[str, int]] = deque()
    frontier_count = 0

    def defer(node_id: str, reason: str, depth: int) -> None:
        nonlocal frontier_count
        frontier_count += 1
        if len(frontier) < 100:
            frontier.append({"node_id": node_id, "reason": reason, "depth": depth})

    if ambiguous:
        unresolved.append({"reason": "name_ambiguous", "next_action": "Select an explicit name scope; multi-area destinations remain unsupported."})
    else:
        visited[root_ids[0]] = 0
        queue.append((root_ids[0], 0))
    while queue:
        node_id, depth = queue.popleft()
        node = nodes[node_id]
        loc = node.source_location
        excluded = _is_excluded(source, loc.sheet_name if loc else None)
        row = {"id": node_id, "kind": node.kind.value, "label": node.label, "depth": depth,
               "source_location": loc.model_dump(mode="json") if loc else None,
               "provenance": node_proofs[node_id]}
        delivered_nodes.append(row)
        if excluded:
            approved = False
            try:
                _scope_guard(source, loc.sheet_name, _normalize_address(loc.address or "")[1], selector)
                approved = True
            except QueryFailure:
                pass
            row["read_only_dependency"] = approved
            boundaries.append({"node_id": node_id, "sheet": loc.sheet_name, "address": loc.address,
                "reason": "excluded_sheet", "approved_read_only": approved,
                "next_action": "Continue approved structural relations or retained consumers; excluded formula/code expansion requires a new approved scope."
                    if approved else "Keep this boundary; broader formula/code exploration requires a new approved scope."})
            if not approved:
                continue
        if node.kind == GraphNodeKind.cell and loc and loc.sheet_name:
            sheet_view(loc.sheet_name)
            if node_id in records:
                entry, record = records[node_id]
                selected_records[(entry[1].view_id, record.record_id)] = record
            else:
                unresolved.append({"node_id": node_id, "reason": "stored_cell_unavailable", "next_action": "Query prepared source evidence; missing cells are not synthesized."})
            if not excluded and node.metadata.get("references_resolved") is False:
                unresolved.append({"node_id": node_id, "reason": "formula_reference_unresolved", "next_action": "Query this cell and its prepared feature diagnostics."})
        problem = name_problems.get(node_id)
        if problem == "constant_name":
            terminals.append({"node_id": node_id, "reason": problem})
        elif problem or node.kind in {GraphNodeKind.external, GraphNodeKind.unsupported, GraphNodeKind.vba} or node.kind == GraphNodeKind.range and node.metadata.get("structured_reference") and node_id not in destinations:
            unresolved.append({"node_id": node_id, "reason": problem or "unsupported_destination_or_auxiliary_vba",
                "next_action": "Inspect the declaration or auxiliary literal evidence; dynamic references need independent evidence."})
        has_neighbor = False
        for relation in neighbors(node):
            has_neighbor = True
            other = relation["target"] if relation["source"] == node_id else relation["source"]
            key = _json_bytes(relation)
            if key in edge_keys:
                continue
            if depth >= max_depth:
                defer(other, "max_depth", depth + 1)
                continue
            if len(edges) >= max_edges:
                defer(other, "max_edges", depth + 1)
                continue
            if other not in visited and len(visited) >= max_nodes:
                defer(other, "max_nodes", depth + 1)
                continue
            edges.append(relation)
            edge_keys.add(key)
            if other not in visited:
                visited[other] = depth + 1
                queue.append((other, depth + 1))
        if not has_neighbor and not problem:
            terminals.append({"node_id": node_id, "reason": "no_supported_static_neighbors"})
    selected = [entry for entry in loaded.values() if any(pair[0] == entry[1].view_id for pair in selected_records)]
    stream = [(entry[1].view_id, record) for entry in selected for record in entry[1].records
              if (entry[1].view_id, record.record_id) in selected_records]
    status = "needs_selection" if ambiguous else "truncated" if frontier_count else "partial" if unresolved or boundaries else "ok"
    result = _finish_packet(manifest_path=manifest_path.expanduser().resolve(), manifest_bytes=manifest_bytes,
        manifest=manifest, source=source, selector=selector, status=status, selected=selected, stream=stream,
        budget=max(1, sum(_tokens(record) for _, record in stream)), cursor=None,
        summary={"label": "derivation", "kind": "static_trace"}, diagnostics=[], source_file_refs=[], source_bytes=len(graph_bytes))
    result["trace"] = {"direction": direction, "limits": limits, "roots": root_ids,
        "edge_direction": "formula_cell_to_referenced_input", "nodes": delivered_nodes, "edges": edges,
        "frontier": frontier, "frontier_count": frontier_count, "frontier_ids_truncated": frontier_count > len(frontier),
        "truncated": bool(frontier_count), "unresolved": unresolved, "scope_boundaries": boundaries,
        "terminals": terminals, "closure": "not_proven",
        "supported_static_traversal_complete": not bool(frontier_count or unresolved or boundaries),
        "next_actions": (["Continue a specific frontier target with explicit bounds."] if frontier_count else [])
            + [item["next_action"] for item in unresolved + boundaries],
        "vba": "Auxiliary literal references only: no read/write direction, runtime order, full call graph, or resolution of Range(Var1).",
        "runtime_verified": False, "recalculated": False}
    result["metrics"].update({"prepared_graph_loads": 1, "prepared_graph_bytes": len(graph_bytes),
        "selected_view_loads": len(loaded), "indexing_recovery_attempts": 0})
    return result


def validate_trace_packet(packet: dict[str, Any], output: dict[str, Any]) -> dict[str, Any]:
    """Replay the bounded derivation against canonical evidence, then reuse claim validation."""
    spec = packet["trace"]
    selector = packet["selector"]
    expected = trace(Path(packet["manifest"]["path"]), packet["source"]["source_id"],
        selector["kind"], selector["target"], sheet=selector["sheet"], direction=spec["direction"], **spec["limits"])
    if _json_bytes(packet) != _json_bytes(expected):
        raise QueryFailure("integrity_failed", "trace_canonical_mismatch", "Trace packet differs from its bounded canonical replay.")
    from excel_to_act.schemas import WorkbookView
    result = validate_agent_output(output, [WorkbookView.model_validate(view) for view in expected["views"]])
    return {"valid": True, "provenance_only": True, "claim_validation": result, "diagnostics": []}
