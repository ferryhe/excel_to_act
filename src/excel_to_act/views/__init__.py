"""Deterministic, source-addressable projections for downstream readers."""

from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from openpyxl.utils.cell import range_boundaries

from excel_to_act.schemas import (
    FormulaGraph,
    RangeInventory,
    SourceLocation,
    Step2Index,
    ViewChunk,
    ViewRecord,
    WorkbookInventory,
    WorkbookManifest,
    WorkbookView,
)
from excel_to_act.graph.builder import RegexFormulaGraphBuilder
from excel_to_act.steps.step2.workflow import build_index, validate_saved_index


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode("utf-8")


def _stable_id(*parts: str) -> str:
    return hashlib.sha256("\0".join(parts).encode("utf-8")).hexdigest()[:24]


def _tokens(record: ViewRecord) -> int:
    # Stable estimate: one token per four UTF-8 bytes, rounded up. This is not a tokenizer.
    return math.ceil(len(_json_bytes(record.model_dump(mode="json"))) / 4)


def _chunks(view_id: str, records: list[ViewRecord], budget: int) -> list[ViewChunk]:
    chunks: list[ViewChunk] = []
    ids: list[str] = []
    size = 0
    over: list[str] = []

    def flush() -> None:
        nonlocal ids, size, over
        if ids:
            chunks.append(ViewChunk(
                chunk_id=_stable_id(view_id, str(len(chunks))),
                record_ids=ids,
                estimated_tokens=size,
                over_budget_record_ids=over,
            ))
        ids, size, over = [], 0, []

    for record in records:
        cost = _tokens(record)
        if ids and size + cost > budget:
            flush()
        ids.append(record.record_id)
        size += cost
        if cost > budget:
            over.append(record.record_id)
            flush()
    flush()
    return chunks


def _record(source_id: str, kind: str, location: SourceLocation, facts: dict[str, Any]) -> ViewRecord:
    identity = json.dumps(location.model_dump(mode="json"), ensure_ascii=False, sort_keys=True)
    return ViewRecord(
        record_id=_stable_id(source_id, kind, identity, json.dumps(facts, ensure_ascii=False, sort_keys=True)),
        record_type=kind,
        source_location=location,
        facts=facts,
    )


def _make_view(
    *, source_id: str, source_sha: str, run_id: str, schema: str, scope: str,
    sheet_name: str | None, records: list[ViewRecord], budget: int,
    region_address: str | None = None, view_key: str = "",
) -> WorkbookView:
    view_id = _stable_id(source_id, scope, sheet_name or "", view_key)
    return WorkbookView(
        view_id=view_id,
        source_id=source_id,
        source_sha256=source_sha,
        source_run_id=run_id,
        source_schema_version=schema,
        scope=scope,
        sheet_name=sheet_name,
        region_address=region_address,
        records=records,
        chunks=_chunks(view_id, records, budget),
        opaque_report=[record.record_id for record in records if record.facts.get("opaque") is True],
    )


def _range_bounds(address: str) -> tuple[int, int, int, int] | None:
    try:
        bounds = range_boundaries(address)
    except ValueError:
        return None
    return bounds if all(value is not None for value in bounds) else None


def _range_facts(item: RangeInventory) -> dict[str, Any]:
    return {"name": item.name, "address": item.address, "kind": item.kind, "metadata": item.metadata}


def _cell_facts(cell: Any) -> dict[str, Any]:
    return {
        "address": cell.address, "kind": cell.kind.value, "value": cell.value,
        "formula": cell.formula, "cached_value": cell.cached_value,
        "cached_value_available": cell.cached_value_available,
    }


def _region_records(source_id: str, sheet: Any, region: RangeInventory) -> list[ViewRecord]:
    bounds = _range_bounds(region.address)
    if bounds is None:
        return []
    min_col, min_row, max_col, max_row = bounds
    records = [_record(source_id, "range", region.source_location, _range_facts(region))]
    records.extend(
        _record(source_id, "cell", cell.source_location, _cell_facts(cell))
        for cell in sheet.cells
        if cell.row >= min_row and cell.row <= max_row and cell.column >= min_col and cell.column <= max_col
    )
    for item in sheet.ranges:
        if item is region or item.source_location.sheet_name != sheet.name:
            continue
        child = _range_bounds(item.address)
        if child and child[0] >= min_col and child[1] >= min_row and child[2] <= max_col and child[3] <= max_row:
            records.append(_record(source_id, "range", item.source_location, _range_facts(item)))
    return records


def compile_views(handoff_path: str | Path, step1_root: str | Path, *, budget: int = 1200) -> list[WorkbookView]:
    """Validate Step 1 through Step 2, then compile one deterministic sheet view per source sheet."""
    if budget < 1:
        raise ValueError("budget must be positive")
    root, handoff = Path(step1_root).expanduser().resolve(), Path(handoff_path).expanduser()
    with TemporaryDirectory(prefix="excel-to-act-views-") as tmp:
        result = build_index(handoff, root, Path(tmp))
        if result.get("status") == "blocked":
            raise ValueError(f"Step 2 rejected handoff: {result.get('diagnostics', [])}")
        index_file = next((Path(item["path"]) for item in result.get("artifacts", []) if item.get("name") == "index.json"), None)
        if index_file is None:
            raise ValueError("Step 2 did not produce index.json")
        validation = validate_saved_index(index_file, root)
        if validation.get("status") != "pass":
            raise ValueError(f"Step 2 validation failed: {validation.get('diagnostics', [])}")
        index = Step2Index.model_validate_json(index_file.read_bytes())
    views: list[WorkbookView] = []
    for entry in index.entries:
        if entry.status not in {"pass", "partial"} or not entry.source_sha256 or not entry.run_id:
            raise ValueError(f"Step 2 entry lacks a usable source identity: {entry.source_id}")
        artifacts = {Path(ref.name).name: root / ref.path for ref in entry.artifacts}
        required = {"inventory.json", "workbook_manifest.json"}
        if not required <= artifacts.keys():
            raise ValueError(f"Step 2 entry {entry.source_id} is missing {sorted(required - artifacts.keys())}")
        inventory = WorkbookInventory.model_validate_json(artifacts["inventory.json"].read_bytes())
        graph = (
            FormulaGraph.model_validate_json(artifacts["dependency_graph.json"].read_bytes())
            if "dependency_graph.json" in artifacts
            else RegexFormulaGraphBuilder().build(inventory)
        )
        manifest = WorkbookManifest.model_validate_json(artifacts["workbook_manifest.json"].read_bytes())
        identities = (inventory.workbook_sha256, manifest.sha256)
        if any(value != entry.source_sha256 for value in identities):
            raise ValueError(f"Artifact identity mismatch for Step 2 source {entry.source_id}")
        if inventory.schema_version != manifest.schema_version or graph.schema_version != inventory.schema_version:
            raise ValueError(f"Artifact schema mismatch for Step 2 source {entry.source_id}")
        schema = inventory.schema_version
        workbook_records: list[ViewRecord] = []
        for item in inventory.workbook_ranges:
            workbook_records.append(_record(entry.source_id, "workbook_range", item.source_location, {
                "name": item.name, "address": item.address, "kind": item.kind, "metadata": item.metadata,
            }))
        for part in manifest.package_parts:
            workbook_records.append(_record(entry.source_id, "package_part", part.source_location, {
                "name": part.name, "content_type": part.content_type,
                "relationship_type": part.relationship_type, "size": part.size, "opaque": part.opaque,
            }))
        vba_part = next((part for part in manifest.package_parts if part.name.lower().endswith("vbaproject.bin")), None)
        if vba_part:
            for module in inventory.vba_modules:
                workbook_records.append(_record(entry.source_id, "vba_module", vba_part.source_location, {
                    "name": module.name, "kind": module.kind, "code": module.code,
                    "procedures": module.procedures,
                }))
        graph_nodes = {node.id: node for node in graph.nodes}
        for node in graph.nodes:
            if node.kind.value == "vba":
                location = node.source_location or (vba_part.source_location if vba_part else None)
                if location is None:
                    raise ValueError(f"VBA graph node {node.id} has no source package location")
                workbook_records.append(_record(entry.source_id, "vba", location, {
                    "id": node.id, "label": node.label, "metadata": node.metadata,
                }))
        for edge in graph.edges:
            if edge.relationship == "vba_ref":
                source_node = graph_nodes.get(edge.source)
                location = edge.source_location
                if location is None and source_node and source_node.kind.value == "vba" and vba_part:
                    location = vba_part.source_location
                if location is None:
                    raise ValueError(f"VBA graph edge {edge.source!r} -> {edge.target!r} has no source location")
                workbook_records.append(_record(entry.source_id, "dependency", location, {
                    "source": edge.source, "target": edge.target,
                    "relationship": edge.relationship, "formula": edge.formula,
                    "confidence": edge.confidence,
                }))
        for feature in inventory.unsupported_features:
            workbook_records.append(_record(entry.source_id, "unsupported", feature.source_location, {
                "feature_type": feature.feature_type, "description": feature.description,
                "opaque": feature.opaque, "severity": feature.severity.value,
            }))
        views.append(_make_view(source_id=entry.source_id, source_sha=entry.source_sha256,
                                run_id=entry.run_id, schema=schema, scope="workbook",
                                sheet_name=None, records=workbook_records, budget=budget))

        for sheet in inventory.sheets:
            records: list[ViewRecord] = []
            graph_nodes = {node.id: node for node in graph.nodes}
            for cell in sheet.cells:
                records.append(_record(entry.source_id, "cell", cell.source_location, _cell_facts(cell)))
            for item in [*sheet.ranges, *sheet.layout_objects]:
                records.append(_record(entry.source_id, "range", item.source_location, _range_facts(item)))
            for edge in graph.edges:
                loc = edge.source_location
                if loc and loc.sheet_name == sheet.name:
                    target = graph_nodes.get(edge.target)
                    target_sheet = target.source_location.sheet_name if target and target.source_location else None
                    records.append(_record(entry.source_id, "dependency", loc, {
                        "source": edge.source, "target": edge.target,
                        "relationship": edge.relationship, "formula": edge.formula,
                        "target_sheet_name": target_sheet,
                        "cross_sheet": target_sheet is not None and target_sheet != sheet.name,
                    }))
            for feature in graph.unsupported_features:
                if feature.source_location.sheet_name == sheet.name:
                    records.append(_record(entry.source_id, "unresolved_reference", feature.source_location, {
                        "feature_type": feature.feature_type, "description": feature.description,
                        "opaque": feature.opaque, "severity": feature.severity.value,
                    }))
            views.append(_make_view(source_id=entry.source_id, source_sha=entry.source_sha256,
                                    run_id=entry.run_id, schema=schema, scope="sheet",
                                    sheet_name=sheet.name, records=records, budget=budget))
            for region in sheet.ranges:
                if region.source_location.sheet_name != sheet.name or _range_bounds(region.address) is None:
                    continue
                region_record = _record(entry.source_id, "range", region.source_location, _range_facts(region))
                region_records = _region_records(entry.source_id, sheet, region)
                views.append(_make_view(
                    source_id=entry.source_id, source_sha=entry.source_sha256,
                    run_id=entry.run_id, schema=schema, scope="region", sheet_name=sheet.name,
                    region_address=region.address, records=region_records, budget=budget,
                    view_key=region_record.record_id,
                ))
    return views


def validate_agent_output(output: dict[str, Any], views: list[WorkbookView]) -> dict[str, Any]:
    """Validate source-bound facts; derivations and unverified claims must be labeled."""
    if not isinstance(output, dict) or not isinstance(output.get("claims"), list):
        raise ValueError("agent output must contain a claims list")
    view_map = {view.view_id: view for view in views}
    opaque: list[dict[str, Any]] = []
    for position, claim in enumerate(output["claims"]):
        if not isinstance(claim, dict):
            raise ValueError(f"claims[{position}] must be an object")
        view = view_map.get(claim.get("view_id"))
        if view is None:
            raise ValueError(f"claims[{position}] has unknown view_id")
        if claim.get("source_run_id") != view.source_run_id or claim.get("source_schema_version") != view.source_schema_version:
            raise ValueError(f"claims[{position}] source run/schema mismatch")
        record = next((item for item in view.records if item.record_id == claim.get("record_id")), None)
        if record is None:
            raise ValueError(f"claims[{position}] has unknown record_id")
        if _json_bytes(claim.get("source_location")) != _json_bytes(record.source_location.model_dump(mode="json")):
            raise ValueError(f"claims[{position}] source_location mismatch")
        kind = claim.get("kind")
        if kind == "fact":
            if _json_bytes(claim.get("value")) != _json_bytes(record.facts):
                raise ValueError(f"claims[{position}] fact differs from source record")
        elif kind == "derivation":
            if not isinstance(claim.get("derivation"), str) or not claim["derivation"].strip():
                raise ValueError(f"claims[{position}] derivation must be described")
        elif kind == "unverified":
            if claim.get("verified") is not False:
                raise ValueError(f"claims[{position}] unverified claim must set verified=false")
        elif kind == "opaque":
            if record.facts.get("opaque") is not True or claim.get("reported_opaque") is not True or "value" in claim:
                raise ValueError(f"claims[{position}] opaque content must be reported without guessing")
            opaque.append({"view_id": view.view_id, "record_id": record.record_id,
                           "source_location": claim["source_location"]})
        else:
            raise ValueError(f"claims[{position}] kind must be fact, derivation, unverified, or opaque")
    return {"valid": True, "opaque_reports": opaque}


def serialize_views(views: list[WorkbookView]) -> bytes:
    return _json_bytes([view.model_dump(mode="json") for view in views])
