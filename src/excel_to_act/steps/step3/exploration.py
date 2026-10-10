"""Bounded, source-bound Step 3 evidence selection and semantic draft checks."""

from __future__ import annotations

from collections import defaultdict, deque
import hashlib
import json
from pathlib import Path
import re
from typing import Any

from openpyxl.utils.cell import range_boundaries

from excel_to_act.steps.step3 import workflow


_PAGE_DEFAULT = 100
_PAGE_MAX = 500
_TRACE_DEPTH_DEFAULT = 4
_TRACE_DEPTH_MAX = 50
_TRACE_FIELDS_DEFAULT = 100
_TRACE_FIELDS_MAX = 500
_TRACE_EDGE_MAX = 500
_TRACE_FRONTIER_MAX = 500
_SPEC_SECTIONS = (
    "field_roles",
    "logical_axes",
    "calculation_groups",
    "outputs",
    "macro_workflow",
    "open_questions",
    "validation_plan",
)
_ADDRESS = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$", re.IGNORECASE)


def _digest(value: Any) -> str:
    return hashlib.sha256(workflow._json_bytes(value)).hexdigest()


def _source_id(kind: str, facts: dict[str, Any]) -> str:
    return f"source:{kind}:{_digest(facts)}"


def _record_id(kind: str, facts: dict[str, Any]) -> str:
    return f"{kind}:{_digest(facts)}"


def _same_json(left: Any, right: Any) -> bool:
    return workflow._json_bytes(left) == workflow._json_bytes(right)


def _context(analysis_dir: Path) -> dict[str, Any]:
    analysis, manifest, binding, root, inventory_path = workflow._load_context(analysis_dir)
    fields, fields_sha = workflow._read_stage(analysis, manifest, "fields")
    dependencies, dependencies_sha = workflow._read_stage(analysis, manifest, "dependencies")
    if dependencies.get("fields_sha256") != fields_sha:
        raise ValueError("dependencies.json is stale against fields.json; rebuild dependencies")
    inventory = workflow._load_inventory(inventory_path)
    if inventory.get("workbook_sha256") != binding["source"]["source_sha256"]:
        raise ValueError("inventory workbook_sha256 does not match the selected source")
    return {
        "analysis": analysis,
        "manifest": manifest,
        "binding": binding,
        "root": root,
        "inventory": inventory,
        "fields": fields,
        "dependencies": dependencies,
        "fields_sha256": fields_sha,
        "dependencies_sha256": dependencies_sha,
    }


def _stage_binding(context: dict[str, Any]) -> dict[str, Any]:
    return {
        "source": context["binding"]["source"],
        "scope": context["binding"]["scope"],
        "binding_sha256": context["manifest"]["binding_sha256"],
        "stage_hashes": {
            "fields": context["fields_sha256"],
            "dependencies": context["dependencies_sha256"],
        },
    }


def _all_source_records(inventory: dict[str, Any]) -> dict[str, tuple[str, dict[str, Any]]]:
    records: dict[str, tuple[str, dict[str, Any]]] = {}
    for sheet in inventory.get("sheets", []):
        for cell in sheet.get("cells", []):
            if cell.get("kind") in {"value", "formula"}:
                facts = cell
                records[_source_id("cell", facts)] = ("source_cell", facts)
        for item in sheet.get("ranges", []):
            if isinstance(item, dict):
                kind = "table" if item.get("kind") == "table" else "range"
                records[_source_id(kind, item)] = (f"source_{kind}", item)
    for item in inventory.get("workbook_ranges", []):
        if isinstance(item, dict):
            records[_source_id("name", item)] = ("source_name", item)
    for item in inventory.get("vba_modules", []):
        if isinstance(item, dict):
            records[_source_id("vba", item)] = ("source_vba", item)
    return records


def _structural_records(fields: dict[str, Any], dependencies: dict[str, Any]) -> dict[str, tuple[str, dict[str, Any]]]:
    records: dict[str, tuple[str, dict[str, Any]]] = {}
    for item in fields.get("fields", []):
        if isinstance(item, dict) and isinstance(item.get("field_id"), str):
            records[f"field:{item['field_id']}"] = ("field", item)
    for section, prefix in (("edges", "edge"), ("internal_dependencies", "internal"),
                            ("boundaries", "boundary"), ("unresolved", "unresolved")):
        for item in dependencies.get(section, []):
            if isinstance(item, dict):
                records[_record_id(prefix, item)] = (prefix, item)
    return records


def _evidence_entry(kind: str, facts: dict[str, Any]) -> dict[str, Any]:
    return {"kind": kind, "facts": facts}


def _field_summary(facts: dict[str, Any]) -> dict[str, Any]:
    keys = ("field_id", "sheet", "role", "physical_shape", "member_count", "address_extent", "names",
            "formula_family", "formula_examples", "operation_tags", "axis_candidates", "logical_axes")
    return {key: facts.get(key) for key in keys}


def _source_summary(kind: str, facts: dict[str, Any]) -> dict[str, Any]:
    location = facts.get("source_location") or {}
    if kind == "source_cell":
        return {"sheet": location.get("sheet_name"), "address": facts.get("address"), "kind": facts.get("kind"),
                "value": facts.get("value"), "formula": facts.get("formula"), "data_type": facts.get("data_type"),
                "source_identity": facts.get("source_identity")}
    if kind == "source_vba":
        return {"name": facts.get("name"), "kind": facts.get("kind"), "procedures": facts.get("procedures"),
                "source_file": facts.get("source_file"), "sha256": facts.get("sha256"), "code_characters": len(str(facts.get("code", "")))}
    return {"name": facts.get("name"), "kind": facts.get("kind"), "sheet": location.get("sheet_name"),
            "address": facts.get("address"), "metadata": facts.get("metadata")}


def _structural_summary(kind: str, facts: dict[str, Any], fields_by_id: dict[str, dict[str, Any]]) -> dict[str, Any]:
    if kind in {"edge", "internal"}:
        consumer, prerequisite = facts.get("consumer"), facts.get("prerequisite")
        return {"consumer": consumer, "consumer_label": _field_label(fields_by_id[consumer]) if consumer in fields_by_id else None,
                "prerequisite": prerequisite, "prerequisite_label": _field_label(fields_by_id[prerequisite]) if prerequisite in fields_by_id else None,
                "relationship": facts.get("relationship", "prerequisite_before_consumer"),
                "source_formula_cells": facts.get("source_formula_cells", []), "relative_offsets": facts.get("relative_offsets", []),
                "evidence": facts.get("evidence"), "occurrences": facts.get("occurrences"),
                "same_cell": facts.get("same_cell"), "recurrence_candidate": facts.get("recurrence_candidate")}
    consumer = facts.get("consumer")
    return {"consumer": consumer, "consumer_label": _field_label(fields_by_id[consumer]) if consumer in fields_by_id else None,
            "status": facts.get("status"), "boundary_type": facts.get("boundary_type"), "category": facts.get("category"),
            "target_sheet": facts.get("target_sheet"), "evidence": facts.get("evidence"),
            "occurrences": facts.get("occurrences"), "source_formula_cells": facts.get("source_formula_cells", []),
            "target_details": facts.get("target_details"), "evidence_samples": facts.get("evidence_samples")}


def _sheet_map(inventory: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {str(item["name"]).casefold(): item for item in inventory.get("sheets", [])}


def _parse_address(target: str, sheet_option: str | None) -> tuple[str, str, tuple[int, int, int, int]] | None:
    sheet_name, address = workflow._split_qualified(target, sheet_option)
    if not _ADDRESS.fullmatch(address):
        address_like = bool(re.fullmatch(r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*", address, re.IGNORECASE))
        if "!" in target or _ADDRESS.match(target) or address_like:
            raise ValueError("cell selection must be a finite A1 cell or rectangle such as Main!A1:B8")
        return None
    if not sheet_name:
        raise ValueError("an unqualified cell or range requires --sheet")
    try:
        boundaries = range_boundaries(address.replace("$", ""))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"invalid finite A1 address: {target}") from exc
    if any(value is None for value in boundaries):
        raise ValueError("whole-row and whole-column selections are unsupported; use a finite A1 range")
    min_col, min_row, max_col, max_row = boundaries
    if max_col > 16_384 or max_row > 1_048_576:
        raise ValueError("cell selection exceeds Excel worksheet bounds")
    return sheet_name, address, (min_col, min_row, max_col, max_row)


def _field_maps(fields: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[tuple[str, str], str]]:
    by_id = {item["field_id"]: item for item in fields.get("fields", []) if isinstance(item, dict) and item.get("field_id")}
    by_cell: dict[tuple[str, str], str] = {}
    for field_id, item in by_id.items():
        for member in item.get("members", []):
            if isinstance(member, dict) and isinstance(member.get("address"), str):
                by_cell[(str(item.get("sheet", "")).casefold(), member["address"].replace("$", "").upper())] = field_id
    return by_id, by_cell


def _stored_cells(sheet: dict[str, Any]) -> list[dict[str, Any]]:
    return sorted(
        (item for item in sheet.get("cells", []) if item.get("kind") in {"value", "formula"}),
        key=lambda item: (int(item.get("row", 0)), int(item.get("column", 0)), str(item.get("address", ""))),
    )


def _resolve_selector(context: dict[str, Any], target: str, sheet_option: str | None) -> dict[str, Any]:
    inventory = context["inventory"]
    ignored = {name.casefold() for name in context["binding"]["scope"].get("ignored_sheets", [])}
    sheets = _sheet_map(inventory)
    if sheet_option:
        if sheet_option.casefold() not in sheets:
            raise ValueError(f"worksheet not found: {sheet_option}")
        if sheet_option.casefold() in ignored:
            raise ValueError(f"worksheet is excluded by the bound scope: {sheet_option}")

    if target.startswith("vba:"):
        module_name = target[4:]
        if not module_name:
            raise ValueError("vba: target requires an exact module name")
        matches = [item for item in inventory.get("vba_modules", [])
                   if str(item.get("name", "")).casefold() == module_name.casefold()]
        if not matches:
            raise ValueError(f"VBA module not found in canonical inventory: {module_name}")
        if len(matches) > 1:
            return {"kind": "ambiguous", "target": target, "candidates": [_source_id("vba", item) for item in matches],
                    "needs_selection": True, "source_records": [], "field_ids": [], "unknowns": []}
        facts = matches[0]
        return {"kind": "vba_module", "target": target, "source_records": [("source_vba", facts)],
                "field_ids": [], "members": [], "unknowns": ["VBA source is static evidence only; it is not executed and does not establish worksheet ownership."]}

    field_by_id, cell_to_field = _field_maps(context["fields"])
    if target in field_by_id:
        field = field_by_id[target]
        sheet_key = str(field.get("sheet", "")).casefold()
        if sheet_key in ignored:
            raise ValueError(f"field belongs to an excluded worksheet: {field.get('sheet')}")
        sheet_record = sheets.get(sheet_key)
        if sheet_record is None:
            raise ValueError(f"field worksheet is absent from inventory: {field.get('sheet')}")
        member_addresses = {str(member.get("address", "")).replace("$", "").upper() for member in field.get("members", [])}
        cells = [item for item in _stored_cells(sheet_record) if str(item.get("address", "")).replace("$", "").upper() in member_addresses]
        return {"kind": "field", "target": target, "source_records": [("source_cell", item) for item in cells],
                "field_ids": [target], "members": cells, "unknowns": []}

    address_selection = _parse_address(target, sheet_option)
    if address_selection is not None:
        sheet_name, address, (min_col, min_row, max_col, max_row) = address_selection
        if "!" in target and sheet_option and sheet_name.casefold() != sheet_option.casefold():
            raise ValueError("--sheet conflicts with the worksheet qualified in --target")
        key = sheet_name.casefold()
        if key in ignored:
            raise ValueError(f"worksheet is excluded by the bound scope: {sheet_name}")
        sheet = sheets.get(key)
        if sheet is None:
            raise ValueError(f"worksheet not found: {sheet_name}")
        cells = [item for item in _stored_cells(sheet)
                 if min_col <= int(item.get("column", 0)) <= max_col and min_row <= int(item.get("row", 0)) <= max_row]
        field_ids = sorted({cell_to_field[(key, str(item.get("address", "")).replace("$", "").upper())]
                            for item in cells if (key, str(item.get("address", "")).replace("$", "").upper()) in cell_to_field})
        return {"kind": "cell" if min_col == max_col and min_row == max_row else "range", "target": target,
                "sheet": sheet_name, "address": address, "source_records": [("source_cell", item) for item in cells],
                "field_ids": field_ids, "members": cells,
                "unknowns": ["No stored source cell exists at this selection."] if not cells else []}

    # Exact workbook name/table selectors. Sheet context selects a local definition when present.
    candidates: list[tuple[str, dict[str, Any], str | None, str | None]] = []
    for item in inventory.get("workbook_ranges", []):
        if not isinstance(item, dict) or str(item.get("name", "")).casefold() != target.casefold():
            continue
        location = item.get("source_location") or {}
        local_sheet = location.get("sheet_name")
        scope = str((item.get("metadata") or {}).get("scope", "workbook")).casefold()
        if sheet_option and scope not in {"workbook", "global"} and str(local_sheet or "").casefold() != sheet_option.casefold():
            continue
        candidates.append(("source_name", item, local_sheet, item.get("address")))
    for sheet in inventory.get("sheets", []):
        for item in sheet.get("ranges", []):
            if not isinstance(item, dict) or item.get("kind") != "table" or str(item.get("name", "")).casefold() != target.casefold():
                continue
            local_sheet = str(sheet.get("name", ""))
            if sheet_option and local_sheet.casefold() != sheet_option.casefold():
                continue
            candidates.append(("source_table", item, local_sheet, item.get("address")))
    if not candidates:
        raise ValueError(f"exact name or selector not found: {target}")
    if sheet_option:
        local_candidates = [item for item in candidates if item[2] and str(item[2]).casefold() == sheet_option.casefold()
                            and str((item[1].get("metadata") or {}).get("scope", "workbook")).casefold() not in {"workbook", "global"}]
        if local_candidates:
            candidates = local_candidates
    if len(candidates) > 1:
        return {"kind": "ambiguous", "target": target,
                "candidates": [{"evidence_id": _source_id(kind.removeprefix("source_"), facts), "sheet": local_sheet, "address": address}
                               for kind, facts, local_sheet, address in candidates],
                "needs_selection": True, "source_records": [], "field_ids": [], "members": [], "unknowns": []}
    source_kind, descriptor, descriptor_sheet, descriptor_address = candidates[0]
    address_sheet, local_address = workflow._split_qualified(str(descriptor_address or ""), descriptor_sheet)
    descriptor_sheet = address_sheet or descriptor_sheet
    if descriptor_sheet and str(descriptor_sheet).casefold() in ignored:
        raise ValueError(f"name resolves to an excluded worksheet: {descriptor_sheet}")
    source_records: list[tuple[str, dict[str, Any]]] = [(source_kind, descriptor)]
    member_cells: list[dict[str, Any]] = []
    field_ids: list[str] = []
    if descriptor_address:
        try:
            parsed = _parse_address(f"{descriptor_sheet}!{local_address}", descriptor_sheet) if descriptor_sheet else None
        except ValueError:
            parsed = None
        if parsed:
            resolved_sheet, _local, (min_col, min_row, max_col, max_row) = parsed
            key = resolved_sheet.casefold()
            sheet_record = sheets.get(key)
            if sheet_record:
                member_cells = [item for item in _stored_cells(sheet_record)
                                if min_col <= int(item.get("column", 0)) <= max_col and min_row <= int(item.get("row", 0)) <= max_row]
                field_ids = sorted({cell_to_field[(key, str(item.get("address", "")).replace("$", "").upper())]
                                    for item in member_cells if (key, str(item.get("address", "")).replace("$", "").upper()) in cell_to_field})
    source_records.extend(("source_cell", item) for item in member_cells)
    unknowns = [] if member_cells else ["The name descriptor has no resolvable stored cells; no cells were inferred."]
    return {"kind": "name", "target": target, "sheet": descriptor_sheet, "address": descriptor_address,
            "source_records": source_records, "field_ids": field_ids, "members": member_cells, "unknowns": unknowns}


def _selector_output(target: str, selection: dict[str, Any], sheet_context: str | None) -> dict[str, Any]:
    selector = {"target": target, "kind": selection["kind"], "sheet": selection.get("sheet"), "address": selection.get("address")}
    resolved_sheet = selection.get("sheet")
    if sheet_context is not None and (resolved_sheet is None or str(resolved_sheet).casefold() != sheet_context.casefold()):
        # A local name may be scoped on one sheet but refer to cells on another.
        selector["sheet_context"] = sheet_context
    return selector


def _packet_evidence(context: dict[str, Any], selected_source: list[tuple[str, dict[str, Any]]], field_ids: list[str],
                     structural_ids: list[str] | None = None) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for kind, facts in selected_source:
        prefix = {"source_cell": "cell", "source_name": "name", "source_table": "table", "source_range": "range", "source_vba": "vba"}[kind]
        result[_source_id(prefix, facts)] = _evidence_entry(kind, facts)
    field_lookup = {item["field_id"]: item for item in context["fields"].get("fields", [])}
    for field_id in field_ids:
        if field_id in field_lookup:
            result[f"field:{field_id}"] = _evidence_entry("field", field_lookup[field_id])
    if structural_ids:
        canonical = _structural_records(context["fields"], context["dependencies"])
        for evidence_id in structural_ids:
            if evidence_id in canonical:
                kind, facts = canonical[evidence_id]
                result[evidence_id] = _evidence_entry(kind, facts)
    return dict(sorted(result.items()))


def _render_packet_markdown(packet: dict[str, Any]) -> str:
    lines = [f"# Step 3 {packet['kind'].title()} Evidence", "",
             f"**Source:** `{packet['source'].get('source_id')}` · run `{packet['source'].get('run_id')}` · binding `{packet['binding_sha256']}`",
             f"**Selector:** `{packet['selector'].get('target')}` · status `{packet['status']}`", ""]
    if packet["kind"] == "query":
        page = packet.get("page", {})
        lines.extend([f"**Records:** {page.get('total_matches', 0)} total · showing {page.get('shown', 0)} from offset {page.get('offset', 0)} · truncated `{page.get('truncated', False)}`", ""])
        for record in packet.get("records", []):
            facts = record.get("facts", {})
            if record.get("kind") == "source_cell":
                summary = record.get("summary", {})
                lines.append(f"- `{summary.get('sheet')}!{summary.get('address')}` ({summary.get('kind')}): `{summary.get('formula') or summary.get('value')}` · `{record['evidence_id']}`")
            else:
                summary = record.get("summary", {})
                lines.append(f"- {record.get('kind')}: `{summary.get('name') or summary.get('field_id') or summary.get('address')}` · `{record['evidence_id']}`")
        for record in packet.get("descriptors", []):
            summary = record.get("summary", {})
            lines.append(f"- {record.get('kind')} descriptor `{summary.get('name')}` · `{summary.get('address')}` · `{record['evidence_id']}`")
        for item in packet.get("fields", []):
            facts = item["summary"]
            lines.append(f"- Field `{facts.get('field_id')}` · {facts.get('role')} · {facts.get('physical_shape')} · {facts.get('member_count')} stored members · {facts.get('names')}")
    else:
        lines.extend([f"**Direction:** `{packet['direction']}` · visited {len(packet.get('fields', []))} fields · {len(packet.get('edges', []))} edge records", ""])
        for item in packet.get("fields", []):
            facts = item["summary"]
            label = f"{facts.get('sheet')}!{facts.get('address_extent')}"
            lines.append(f"- `{facts.get('field_id')}` {label} ({facts.get('role')}, {facts.get('physical_shape')}) {facts.get('names')}")
        lines.extend(["", "## Dependency evidence", ""])
        for item in packet.get("edges", []):
            summary = item.get("summary", {})
            relationship = "internal" if item.get("kind") == "internal" else "prerequisite before consumer"
            cells = ", ".join(summary.get("source_formula_cells", [])) or "formula cell not sampled"
            lines.append(f"- {summary.get('prerequisite_label')} → {summary.get('consumer_label')} ({relationship}; formula cells: {cells}) · `{item['evidence_id']}`")
        for item in packet.get("boundaries", []):
            summary = item.get("summary", {})
            boundary = summary.get("boundary_type") or summary.get("category") or item.get("kind")
            lines.append(f"- {summary.get('consumer_label')} · {boundary} · {summary.get('evidence')} · `{item['evidence_id']}`")
        for item in packet.get("frontier", []):
            if item.get("reason") == "frontier_items_omitted":
                lines.append(f"- Frontier limit: omitted {item.get('omitted_count')} entries for {item.get('omitted_reason')}; examples {item.get('example_field_ids')}")
            else:
                lines.append(f"- Frontier `{item.get('field_id', '—')}`: {item.get('reason')}")
    if packet.get("unknowns"):
        lines.extend(["", "## Unknowns", "", *[f"- {item}" for item in packet["unknowns"]]])
    lines.extend(["", "This packet is source-bound static evidence. It does not prove business meaning, formula results, or runtime behavior.", ""])
    return "\n".join(lines)


def _output_dir(context: dict[str, Any], out_dir: Path) -> Path:
    output = out_dir.expanduser().resolve()
    binding = context["binding"]
    protected = [context["analysis"], context["root"], Path(binding["index_path"]).resolve()]
    protected.extend([context["root"] / binding["artifacts"][key + "_path"] for key in ("inventory", "manifest")])
    if binding["scope"].get("path"):
        protected.append(Path(binding["scope"]["path"]).resolve())
    for path in protected:
        path = path.resolve()
        if output == path or output in path.parents or path in output.parents:
            raise ValueError(f"packet output directory overlaps a bound input: {path}")
    return output


def _write_packet(context: dict[str, Any], out_dir: Path, packet: dict[str, Any]) -> dict[str, Any]:
    output = _output_dir(context, out_dir)
    stem = packet["kind"]
    json_path, markdown_path = output / f"{stem}.json", output / f"{stem}.md"
    packet_bytes = workflow._json_bytes(packet)
    markdown_bytes = _render_packet_markdown(packet).encode("utf-8")
    workflow._write_bytes(json_path, packet_bytes)
    workflow._write_bytes(markdown_path, markdown_bytes)
    return {"json": str(json_path), "markdown": str(markdown_path), "json_sha256": workflow._hash_bytes(packet_bytes),
            "markdown_sha256": workflow._hash_bytes(markdown_bytes)}


def query(analysis_dir: Path, target: str, *, sheet: str | None = None, offset: int = 0,
          limit: int = _PAGE_DEFAULT, out_dir: Path | None = None,
          _checked_context: dict[str, Any] | None = None) -> dict[str, Any]:
    """Return a deterministic bounded page of canonical source and field evidence."""
    try:
        if offset < 0:
            raise ValueError("offset must be zero or greater")
        if not 1 <= limit <= _PAGE_MAX:
            raise ValueError(f"limit must be between 1 and {_PAGE_MAX}")
        if not target:
            raise ValueError("target must not be empty")
        # Spec validation replays packets against its already checked context to avoid
        # reparsing the potentially large canonical inventory for each packet.
        context = _checked_context if _checked_context is not None else _context(analysis_dir)
        selection = _resolve_selector(context, target, sheet)
        if selection["kind"] == "ambiguous":
            return {"tool": "step3.query", "status": "blocked", "needs_selection": True,
                    "candidates": selection["candidates"], "diagnostics": [{"code": "ambiguous_selector", "message": "selector matches multiple names; supply --sheet or a more exact target"}]}
        source_records = selection["source_records"]
        if selection["kind"] == "vba_module":
            source_records = source_records[offset:offset + limit]
            all_records = selection["source_records"]
            total = len(all_records)
            members: list[dict[str, Any]] = []
        else:
            all_records = selection["members"]
            total = len(all_records)
            members = all_records[offset:offset + limit]
            member_ids = {_source_id("cell", item) for item in members}
            source_records = [(kind, facts) for kind, facts in source_records
                              if kind != "source_cell" or _source_id("cell", facts) in member_ids]
            # A selected exact name/table descriptor remains evidence on every page.
            if selection["kind"] == "name":
                source_records = [item for item in selection["source_records"] if item[0] != "source_cell"] + [item for item in source_records if item[0] == "source_cell"]
        _by_id, cell_to_field = _field_maps(context["fields"])
        field_ids = selection["field_ids"] if selection["kind"] == "vba_module" else sorted({
            field_id for item in members
            for field_id in [cell_to_field.get((str(selection.get("sheet") or (item.get("source_location") or {}).get("sheet_name", "")).casefold(), str(item.get("address", "")).replace("$", "").upper()))]
            if field_id
        })
        # A field selector preserves its owning field even when a page begins past member zero.
        if selection["kind"] == "field":
            field_ids = selection["field_ids"]
        source_prefixes = {"source_cell": "cell", "source_name": "name", "source_table": "table", "source_range": "range", "source_vba": "vba"}
        evidence = _packet_evidence(context, source_records, field_ids)
        record_items = ([{"evidence_id": _source_id("vba", facts), "kind": kind, "summary": _source_summary(kind, facts)} for kind, facts in source_records]
                        if selection["kind"] == "vba_module" else
                        [{"evidence_id": _source_id("cell", item), "kind": "source_cell", "summary": _source_summary("source_cell", item)} for item in members])
        descriptor_items = [{"evidence_id": _source_id(source_prefixes[kind], facts), "kind": kind, "summary": _source_summary(kind, facts)}
                            for kind, facts in source_records if kind != "source_cell"]
        fields_by_id = {item["field_id"]: item for item in context["fields"].get("fields", [])}
        field_items = [{"evidence_id": f"field:{field_id}", "kind": "field", "summary": _field_summary(fields_by_id[field_id])}
                       for field_id in field_ids if field_id in fields_by_id]
        shown = len(record_items) if selection["kind"] == "vba_module" else len(members)
        page = {"offset": offset, "limit": limit, "total_matches": total, "shown": shown,
                "next_offset": offset + shown if offset + shown < total else None,
                "truncated": offset + shown < total}
        packet = {"schema_version": "step3.query.v1", "kind": "query", "status": "pass",
                  **_stage_binding(context), "selector": _selector_output(target, selection, sheet),
                  "page": page, "records": record_items, "descriptors": descriptor_items, "fields": field_items, "evidence": evidence,
                  "unknowns": selection["unknowns"]}
        result = {"tool": "step3.query", "status": "pass", "packet": packet}
        if out_dir is not None:
            result["artifacts"] = _write_packet(context, out_dir, packet)
        return result
    except Exception as exc:
        return workflow._blocked("step3.query", exc)


def _field_label(field: dict[str, Any]) -> str:
    return f"{field.get('sheet')}!{field.get('address_extent')} ({', '.join(field.get('names', [])) or field.get('role')})"


def trace(analysis_dir: Path, target: str, *, sheet: str | None = None, direction: str = "upstream",
          max_depth: int = _TRACE_DEPTH_DEFAULT, max_fields: int = _TRACE_FIELDS_DEFAULT,
          out_dir: Path | None = None, _checked_context: dict[str, Any] | None = None) -> dict[str, Any]:
    """Trace checked prerequisite edges while reporting every traversal limit."""
    try:
        if direction not in {"upstream", "downstream"}:
            raise ValueError("direction must be upstream or downstream")
        if not 0 <= max_depth <= _TRACE_DEPTH_MAX:
            raise ValueError(f"max-depth must be between 0 and {_TRACE_DEPTH_MAX}")
        if not 1 <= max_fields <= _TRACE_FIELDS_MAX:
            raise ValueError(f"max-fields must be between 1 and {_TRACE_FIELDS_MAX}")
        # Spec validation replays packets against its already checked context to avoid
        # reparsing the potentially large canonical inventory for each packet.
        context = _checked_context if _checked_context is not None else _context(analysis_dir)
        selection = _resolve_selector(context, target, sheet)
        if selection["kind"] == "ambiguous":
            return {"tool": "step3.trace", "status": "blocked", "needs_selection": True,
                    "candidates": selection["candidates"], "diagnostics": [{"code": "ambiguous_selector", "message": "selector matches multiple names; supply --sheet or a more exact target"}]}
        if selection["kind"] == "vba_module":
            raise ValueError("VBA modules are source-only and cannot be dependency-traced")
        fields_by_id = {item["field_id"]: item for item in context["fields"].get("fields", [])}
        start_all = sorted(selection["field_ids"])
        if not start_all:
            raise ValueError("selector contains no retained primary field to trace")
        frontier: list[dict[str, Any]] = []
        frontier_overflow: dict[str, dict[str, Any]] = {}

        def add_frontier(item: dict[str, Any]) -> None:
            if len(frontier) < _TRACE_FRONTIER_MAX:
                frontier.append(item)
                return
            reason = str(item.get("reason", "unknown"))
            summary = frontier_overflow.setdefault(reason, {"count": 0, "examples": []})
            summary["count"] += 1
            if len(summary["examples"]) < 5 and item.get("field_id") and item["field_id"] not in summary["examples"]:
                summary["examples"].append(item["field_id"])

        starts = start_all[:max_fields]
        for field_id in start_all[max_fields:]:
            add_frontier({"field_id": field_id, "label": _field_label(fields_by_id[field_id]), "reason": "max_fields_start_limit"})
        if len(start_all) > max_fields:
            stopping_reasons = {"max_fields"}
        else:
            stopping_reasons: set[str] = set()
        adjacency: dict[str, list[tuple[str, str, dict[str, Any]]]] = defaultdict(list)
        for item in context["dependencies"].get("edges", []):
            consumer, prerequisite = item.get("consumer"), item.get("prerequisite")
            if consumer not in fields_by_id or prerequisite not in fields_by_id:
                continue
            if direction == "upstream":
                adjacency[consumer].append((prerequisite, "edge", item))
            else:
                adjacency[prerequisite].append((consumer, "edge", item))
        for item in context["dependencies"].get("internal_dependencies", []):
            consumer, prerequisite = item.get("consumer"), item.get("prerequisite")
            if consumer in fields_by_id and prerequisite in fields_by_id:
                # Internal records are evidence of recurrence/self-reference, not a new field hop.
                adjacency[consumer].append((consumer, "internal", item))
        for field_id in adjacency:
            adjacency[field_id].sort(key=lambda entry: (entry[0], entry[1], _digest(entry[2])))
        boundary_by_field: dict[str, list[tuple[str, dict[str, Any]]]] = defaultdict(list)
        for item in context["dependencies"].get("boundaries", []):
            boundary_by_field[item.get("consumer")].append(("boundary", item))
        for item in context["dependencies"].get("unresolved", []):
            boundary_by_field[item.get("consumer")].append(("unresolved", item))
        for values in boundary_by_field.values():
            values.sort(key=lambda entry: (entry[0], _digest(entry[1])))

        visited: dict[str, int] = {field_id: 0 for field_id in starts}
        queue = deque(starts)
        visited_edges: list[tuple[str, dict[str, Any]]] = []
        visited_boundaries: list[tuple[str, dict[str, Any]]] = []
        cycles: list[dict[str, Any]] = []
        edge_count = 0
        while queue:
            current = queue.popleft()
            depth = visited[current]
            for record_kind, record in boundary_by_field.get(current, []):
                if len(visited_boundaries) >= _TRACE_EDGE_MAX:
                    stopping_reasons.add("record_limit")
                    add_frontier({"field_id": current, "label": _field_label(fields_by_id[current]), "reason": "boundary_record_limit"})
                    break
                visited_boundaries.append((record_kind, record))
            neighbors = adjacency.get(current, [])
            for position, (neighbor, record_kind, record) in enumerate(neighbors):
                if edge_count >= _TRACE_EDGE_MAX:
                    stopping_reasons.add("edge_limit")
                    add_frontier({"field_id": current, "label": _field_label(fields_by_id[current]), "reason": "edge_record_limit",
                                  "remaining_at_field": len(neighbors) - position})
                    for queued in sorted(set(queue)):
                        add_frontier({"field_id": queued, "label": _field_label(fields_by_id[queued]), "reason": "not_expanded_after_edge_limit"})
                    queue.clear()
                    break
                edge_count += 1
                visited_edges.append((record_kind, record))
                if record_kind == "internal":
                    if record.get("same_cell") or record.get("recurrence_candidate"):
                        cycles.append({"field_id": current, "evidence_id": _record_id("internal", record),
                                       "kind": "same_cell" if record.get("same_cell") else "recurrence_candidate"})
                    continue
                if neighbor in visited:
                    if neighbor == current or visited[neighbor] <= depth:
                        cycles.append({"from_field_id": current, "to_field_id": neighbor, "evidence_id": _record_id("edge", record), "kind": "cycle_or_back_edge"})
                    continue
                if depth >= max_depth:
                    stopping_reasons.add("max_depth")
                    add_frontier({"field_id": neighbor, "label": _field_label(fields_by_id[neighbor]), "reason": "max_depth",
                                  "from_field_id": current, "evidence_id": _record_id("edge", record)})
                    continue
                if len(visited) >= max_fields:
                    stopping_reasons.add("max_fields")
                    add_frontier({"field_id": neighbor, "label": _field_label(fields_by_id[neighbor]), "reason": "max_fields",
                                  "from_field_id": current, "evidence_id": _record_id("edge", record)})
                    continue
                visited[neighbor] = depth + 1
                queue.append(neighbor)
        visited_fields = sorted(visited, key=lambda field_id: (visited[field_id], fields_by_id[field_id].get("sheet", "").casefold(),
                                                               fields_by_id[field_id].get("address_extent", ""), field_id))
        field_items = [{"evidence_id": f"field:{field_id}", "kind": "field", "summary": _field_summary(fields_by_id[field_id])} for field_id in visited_fields]
        structural_ids: list[str] = []
        for kind, item in visited_edges:
            structural_ids.append(_record_id("internal" if kind == "internal" else "edge", item))
        for kind, item in visited_boundaries:
            structural_ids.append(_record_id(kind, item))
        if frontier_overflow:
            stopping_reasons.add("frontier_limit")
            frontier.extend({"reason": "frontier_items_omitted", "omitted_reason": reason,
                             "omitted_count": values["count"], "example_field_ids": values["examples"]}
                            for reason, values in sorted(frontier_overflow.items()))
        # Address/range selectors may span thousands of cells; field evidence is complete and
        # bounded, so retain only an explicit single-cell fact or the selected name descriptor.
        trace_sources = [(kind, facts) for kind, facts in selection["source_records"]
                         if kind != "source_cell" or selection["kind"] == "cell"]
        evidence = _packet_evidence(context, trace_sources, visited_fields, structural_ids)
        source_kind_prefix = {"source_cell": "cell", "source_name": "name", "source_table": "table", "source_range": "range", "source_vba": "vba"}
        selector_evidence_ids = sorted(_source_id(source_kind_prefix[kind], facts) for kind, facts in trace_sources)
        edge_items = [{"evidence_id": _record_id("internal" if kind == "internal" else "edge", item), "kind": kind,
                       "summary": _structural_summary(kind, item, fields_by_id)}
                      for kind, item in visited_edges]
        boundary_items = [{"evidence_id": _record_id(kind, item), "kind": kind, "summary": _structural_summary(kind, item, fields_by_id)}
                          for kind, item in visited_boundaries]
        frontier = sorted(frontier, key=lambda item: (item.get("reason", ""), item.get("field_id", ""), item.get("from_field_id", ""), item.get("evidence_id", "")))
        unknowns = sorted({str(item.get("category", item.get("boundary_type", item.get("reason", "unknown"))))
                           for kind, item in visited_boundaries})
        selector_output = _selector_output(target, selection, sheet)
        selector_output["evidence_ids"] = selector_evidence_ids
        packet = {"schema_version": "step3.trace.v1", "kind": "trace", "status": "pass", **_stage_binding(context),
                  "selector": selector_output,
                  "direction": direction, "limits": {"max_depth": max_depth, "max_fields": max_fields, "max_edge_records": _TRACE_EDGE_MAX},
                  "fields": field_items, "edges": edge_items, "boundaries": boundary_items, "frontier": frontier,
                  "truncated": bool(frontier), "stopping_reasons": sorted(stopping_reasons), "cycles": cycles,
                  "evidence": evidence, "unknowns": unknowns}
        result = {"tool": "step3.trace", "status": "pass", "packet": packet}
        if out_dir is not None:
            result["artifacts"] = _write_packet(context, out_dir, packet)
        return result
    except Exception as exc:
        return workflow._blocked("step3.trace", exc)


def _canonical_evidence(context: dict[str, Any]) -> dict[str, tuple[str, dict[str, Any]]]:
    records = _all_source_records(context["inventory"])
    records.update(_structural_records(context["fields"], context["dependencies"]))
    return records


def _check_packet_records(packet: dict[str, Any], packet_evidence: dict[str, dict[str, Any]],
                          all_fields: dict[str, dict[str, Any]]) -> None:
    sections = {"query": ("records", "descriptors", "fields"), "trace": ("fields", "edges", "boundaries")}
    kind = packet.get("kind")
    if kind not in sections:
        raise ValueError("evidence packet kind is unsupported")
    fields_by_id = all_fields
    listed_ids: set[str] = set()
    for section in sections[kind]:
        items = packet.get(section)
        if not isinstance(items, list):
            raise ValueError(f"evidence packet {section} must be an array")
        for item in items:
            if not isinstance(item, dict) or not isinstance(item.get("evidence_id"), str) or not isinstance(item.get("summary"), dict):
                raise ValueError(f"evidence packet {section} record is malformed")
            evidence_id = item["evidence_id"]
            canonical = packet_evidence.get(evidence_id)
            if canonical is None:
                raise ValueError(f"evidence packet {section} record is not listed in its evidence map: {evidence_id}")
            record_kind, facts = canonical["kind"], canonical["facts"]
            if section in {"records", "descriptors"}:
                expected_summary = _source_summary(record_kind, facts)
            elif record_kind == "field":
                expected_summary = _field_summary(facts)
            else:
                expected_summary = _structural_summary(record_kind, facts, fields_by_id)
            if item.get("kind") != record_kind or not _same_json(item.get("summary"), expected_summary):
                raise ValueError(f"evidence packet {section} record differs from its checked evidence: {evidence_id}")
            listed_ids.add(evidence_id)
    if kind == "trace":
        selector = packet.get("selector")
        selector_ids = selector.get("evidence_ids", []) if isinstance(selector, dict) else []
        if not isinstance(selector_ids, list) or any(not isinstance(evidence_id, str) for evidence_id in selector_ids):
            raise ValueError("trace selector evidence_ids must be an array of evidence ID strings")
        for evidence_id in selector_ids:
            if evidence_id not in packet_evidence or not packet_evidence[evidence_id].get("kind", "").startswith("source_"):
                raise ValueError(f"trace selector contains an unknown source evidence ID: {evidence_id}")
            listed_ids.add(evidence_id)
    extra_ids = sorted(set(packet_evidence) - listed_ids)
    if extra_ids:
        raise ValueError(f"{kind} evidence packet contains evidence IDs absent from its selected records: {extra_ids}")
    referenced_ids = set()
    for section in ("frontier", "cycles"):
        items = packet.get(section, [])
        if not isinstance(items, list):
            raise ValueError(f"evidence packet {section} must be an array")
        referenced_ids.update(item["evidence_id"] for item in items if isinstance(item, dict) and isinstance(item.get("evidence_id"), str))
    missing = sorted(referenced_ids - set(packet_evidence))
    if missing:
        raise ValueError(f"evidence packet traversal references evidence missing from its evidence map: {missing}")
    if kind == "trace" and bool(packet.get("truncated")) != bool(packet.get("frontier")):
        raise ValueError("trace packet truncated flag does not match its frontier")


def _check_packet_replay(context: dict[str, Any], packet: dict[str, Any]) -> None:
    """Rebuild a bounded packet from its selector using the already checked source context."""
    selector = packet.get("selector")
    if not isinstance(selector, dict) or not isinstance(selector.get("target"), str):
        raise ValueError("evidence packet selector requires a target string")
    target = selector["target"]
    sheet = selector.get("sheet")
    if sheet is not None and not isinstance(sheet, str):
        raise ValueError("evidence packet selector sheet must be a string or null")
    sheet_context = selector.get("sheet_context", sheet)
    if sheet_context is not None and not isinstance(sheet_context, str):
        raise ValueError("evidence packet selector sheet_context must be a string or null")

    if packet.get("schema_version") == "step3.query.v1":
        page = packet.get("page")
        if not isinstance(page, dict):
            raise ValueError("query packet page metadata is missing")
        offset, limit = page.get("offset"), page.get("limit")
        if type(offset) is not int or type(limit) is not int:
            raise ValueError("query packet offset and limit must be integers")
        replay = query(context["analysis"], target, sheet=sheet_context, offset=offset, limit=limit,
                       _checked_context=context)
    elif packet.get("schema_version") == "step3.trace.v1":
        limits = packet.get("limits")
        if not isinstance(limits, dict):
            raise ValueError("trace packet limits are missing")
        max_depth, max_fields = limits.get("max_depth"), limits.get("max_fields")
        if type(max_depth) is not int or type(max_fields) is not int:
            raise ValueError("trace packet depth and field limits must be integers")
        direction = packet.get("direction")
        if not isinstance(direction, str):
            raise ValueError("trace packet direction must be a string")
        replay = trace(context["analysis"], target, sheet=sheet_context, direction=direction,
                       max_depth=max_depth, max_fields=max_fields, _checked_context=context)
    else:
        raise ValueError("evidence packet schema is unsupported")

    if replay.get("status") != "pass" or not isinstance(replay.get("packet"), dict):
        raise ValueError("evidence packet selector cannot be replayed against the checked source")
    if not _same_json(packet, replay["packet"]):
        raise ValueError("evidence packet selector, page, or traversal differs from its checked source selection")


def _validate_spec_input(context: dict[str, Any], spec_path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], set[str]]:
    spec_path = spec_path.expanduser().resolve()
    spec = workflow._read_json(spec_path)
    if not isinstance(spec, dict) or spec.get("schema_version") != "step3.model_spec.input.v1":
        raise ValueError("spec must use schema_version step3.model_spec.input.v1")
    binding = context["binding"]
    source = binding["source"]
    expected_source = {key: source.get(key) for key in ("source_id", "run_id", "source_sha256")}
    if spec.get("source") != expected_source:
        raise ValueError("spec source identity is stale or does not match the selected Step 2 source")
    if spec.get("binding_sha256") != context["manifest"]["binding_sha256"]:
        raise ValueError("spec binding_sha256 is stale")
    expected_stages = {"fields": context["fields_sha256"], "dependencies": context["dependencies_sha256"]}
    if spec.get("stage_hashes") != expected_stages:
        raise ValueError("spec fields/dependencies stage hashes are stale")
    packet_declarations = spec.get("evidence_packets")
    if not isinstance(packet_declarations, list) or not packet_declarations:
        raise ValueError("spec evidence_packets must declare at least one packet path and SHA-256")
    canonical = _canonical_evidence(context)
    packet_evidence: dict[str, dict[str, Any]] = {}
    loaded_packets: list[dict[str, Any]] = []
    seen_paths: set[str] = set()
    current_binding = _stage_binding(context)
    for declaration in packet_declarations:
        if not isinstance(declaration, dict) or not isinstance(declaration.get("path"), str) or not isinstance(declaration.get("sha256"), str):
            raise ValueError("each evidence packet declaration requires relative path and sha256")
        relative = Path(declaration["path"])
        if relative.is_absolute():
            raise ValueError("evidence packet paths must be relative to the spec file")
        packet_path = (spec_path.parent / relative).resolve()
        if not packet_path.is_file():
            raise ValueError(f"evidence packet does not exist: {declaration['path']}")
        if str(packet_path).casefold() in seen_paths:
            raise ValueError("spec declares the same evidence packet more than once")
        seen_paths.add(str(packet_path).casefold())
        actual_sha = workflow._hash_file(packet_path)
        if actual_sha != declaration["sha256"]:
            raise ValueError(f"evidence packet SHA-256 mismatch: {declaration['path']}")
        packet = workflow._read_json(packet_path)
        if not isinstance(packet, dict) or packet.get("schema_version") not in {"step3.query.v1", "step3.trace.v1"}:
            raise ValueError(f"unsupported evidence packet schema: {declaration['path']}")
        if packet.get("binding_sha256") != current_binding["binding_sha256"] or not _same_json(packet.get("source"), current_binding["source"]):
            raise ValueError(f"evidence packet source binding is stale: {declaration['path']}")
        if packet.get("stage_hashes") != current_binding["stage_hashes"]:
            raise ValueError(f"evidence packet stage hashes are stale: {declaration['path']}")
        packet_map = packet.get("evidence")
        if not isinstance(packet_map, dict):
            raise ValueError(f"evidence packet has no evidence map: {declaration['path']}")
        this_packet_evidence: dict[str, dict[str, Any]] = {}
        for evidence_id, entry in packet_map.items():
            if evidence_id not in canonical:
                raise ValueError(f"evidence packet contains an unknown or undelivered evidence ID: {evidence_id}")
            if not isinstance(entry, dict) or set(entry) != {"kind", "facts"}:
                raise ValueError(f"evidence packet record is malformed: {evidence_id}")
            expected_kind, expected_facts = canonical[evidence_id]
            if entry.get("kind") != expected_kind or not _same_json(entry.get("facts"), expected_facts):
                raise ValueError(f"evidence packet facts do not match current canonical artifacts: {evidence_id}")
            packet_evidence[evidence_id] = entry
            this_packet_evidence[evidence_id] = entry
        all_fields = {item["field_id"]: item for item in context["fields"].get("fields", []) if isinstance(item, dict) and item.get("field_id")}
        _check_packet_records(packet, this_packet_evidence, all_fields)
        _check_packet_replay(context, packet)
        loaded_packets.append({"path": declaration["path"], "sha256": actual_sha})
    expected_keys = {"schema_version", "source", "binding_sha256", "stage_hashes", "evidence_packets", *_SPEC_SECTIONS}
    unknown_keys = sorted(set(spec) - expected_keys)
    missing_keys = sorted(expected_keys - set(spec))
    if unknown_keys or missing_keys:
        raise ValueError(f"spec top-level sections mismatch; missing={missing_keys}, unknown={unknown_keys}")
    for section in _SPEC_SECTIONS:
        items = spec.get(section)
        if not isinstance(items, list):
            raise ValueError(f"spec section {section} must be an array")
        for position, item in enumerate(items):
            if not isinstance(item, dict):
                raise ValueError(f"{section}[{position}] must be an object")
            status = item.get("status")
            if status == "fact":
                references = item.get("evidence")
                if not isinstance(references, list) or not references:
                    raise ValueError(f"{section}[{position}] fact requires evidence records")
                for reference in references:
                    if not isinstance(reference, dict) or not isinstance(reference.get("id"), str) or not isinstance(reference.get("facts"), dict):
                        raise ValueError(f"{section}[{position}] fact evidence requires id and complete facts")
                    evidence_id = reference["id"]
                    if evidence_id not in packet_evidence:
                        raise ValueError(f"{section}[{position}] cites evidence not delivered in declared packets: {evidence_id}")
                    if not _same_json(reference["facts"], canonical[evidence_id][1]):
                        raise ValueError(f"{section}[{position}] fact payload differs from current canonical evidence: {evidence_id}")
                    if not _same_json(reference["facts"], packet_evidence[evidence_id]["facts"]):
                        raise ValueError(f"{section}[{position}] fact payload differs from its evidence packet: {evidence_id}")
                if not isinstance(item.get("claim"), str) or not item["claim"].strip():
                    raise ValueError(f"{section}[{position}] fact requires a concise claim")
            elif status == "inferred":
                refs = item.get("evidence_refs")
                if not isinstance(item.get("statement"), str) or not item["statement"].strip() or not isinstance(refs, list) or not refs:
                    raise ValueError(f"{section}[{position}] inferred item requires statement and evidence_refs")
                if any(not isinstance(ref, str) for ref in refs):
                    raise ValueError(f"{section}[{position}] evidence_refs must contain only evidence ID strings")
                missing = sorted({ref for ref in refs if not isinstance(ref, str) or ref not in packet_evidence})
                if missing:
                    raise ValueError(f"{section}[{position}] inferred item cites evidence not delivered in packets: {missing}")
            elif status == "pending":
                missing = item.get("missing_evidence")
                if not isinstance(item.get("question"), str) or not item["question"].strip() or not isinstance(missing, list) or not missing:
                    raise ValueError(f"{section}[{position}] pending item requires question and missing_evidence")
                if any(not isinstance(entry, str) or not entry.strip() for entry in missing):
                    raise ValueError(f"{section}[{position}] missing_evidence must contain non-empty strings")
            else:
                raise ValueError(f"{section}[{position}] status must be fact, inferred, or pending")
            details = item.get("details")
            if details is not None:
                if not isinstance(details, dict):
                    raise ValueError(f"{section}[{position}] details must be an object")
                field_ids = details.get("field_ids")
                if field_ids is not None:
                    if not isinstance(field_ids, list) or any(not isinstance(field_id, str) for field_id in field_ids):
                        raise ValueError(f"{section}[{position}] details.field_ids must be an array of field IDs")
                    delivered_fields = {evidence_id.removeprefix("field:") for evidence_id, entry in packet_evidence.items()
                                        if entry.get("kind") == "field"}
                    unknown_fields = sorted(set(field_ids) - delivered_fields)
                    if unknown_fields:
                        raise ValueError(f"{section}[{position}] details.field_ids must cite fields delivered in evidence packets: {unknown_fields}")
    return spec, loaded_packets, set(packet_evidence)


def _spec_markdown(result: dict[str, Any]) -> str:
    lines = ["# Step 3 Semantic Draft", "", f"**Validation:** `{result['status']}` · **Source:** `{result['source']['source_id']}` · run `{result['source']['run_id']}`",
             f"**Binding:** `{result['binding_sha256']}` · fields `{result['stage_hashes']['fields']}` · dependencies `{result['stage_hashes']['dependencies']}`", "",
             "This is a cited semantic draft. Validation checks source and structural references; it does not prove business meaning, numerical behavior, formula execution, or macro runtime behavior.", ""]
    for section in _SPEC_SECTIONS:
        lines.extend([f"## {section.replace('_', ' ').title()}", ""])
        items = result.get(section, [])
        if not items:
            lines.append("- No items supplied.")
        for item in items:
            status = item["status"]
            body = item.get("claim", item.get("statement", item.get("question", "")))
            refs = [record.get("id", "") for record in item.get("evidence", [])] if status == "fact" else item.get("evidence_refs", [])
            suffix = f" · evidence: {', '.join(f'`{ref}`' for ref in refs)}" if refs else ""
            if status == "pending":
                suffix += f" · missing: {', '.join(item.get('missing_evidence', []))}"
            lines.append(f"- **{status}** {body}{suffix}")
            if isinstance(item.get("details"), dict):
                lines.extend(["", "  ```json", *[f"  {line}" for line in json.dumps(item["details"], ensure_ascii=False, indent=2, sort_keys=True).splitlines()], "  ```"])
    lines.extend(["", "## Readiness", "", "- Runtime verified: no", "- Generation ready: no", "- GPU executable: no", ""])
    return "\n".join(lines)


def _attach_spec_handoff(context: dict[str, Any], result: dict[str, Any], json_sha: str, markdown_sha: str) -> None:
    handoff_path = context["analysis"] / "handoff.json"
    handoff = workflow._read_json(handoff_path)
    if handoff.get("binding_sha256") != context["manifest"]["binding_sha256"]:
        raise ValueError("handoff binding changed before model spec could be attached")
    handoff["model_spec"] = {"json": "model_spec.json", "markdown": "model_spec.md", "json_sha256": json_sha,
                             "markdown_sha256": markdown_sha, "fields_sha256": context["fields_sha256"],
                             "dependencies_sha256": context["dependencies_sha256"], "status": result["status"]}
    workflow._write_json(handoff_path, handoff)
    markdown_path = context["analysis"] / "handoff.md"
    previous = markdown_path.read_text(encoding="utf-8") if markdown_path.is_file() else "# Step 3 Draft Handoff\n"
    lines = previous.rstrip().splitlines()
    if "## Semantic draft" in previous:
        lines = previous.split("## Semantic draft", 1)[0].rstrip().splitlines()
    lines.extend(["", "## Semantic draft", "", "Validated semantic specification: `model_spec.json` and `model_spec.md`.", ""])
    workflow._write_bytes(markdown_path, "\n".join(lines).encode("utf-8"))


def validate_model_spec(analysis_dir: Path, spec_path: Path) -> dict[str, Any]:
    """Check a cited semantic draft and write its paired outputs only after full validation."""
    try:
        context = _context(analysis_dir)
        spec, packets, evidence_ids = _validate_spec_input(context, spec_path)
        result = {
            "schema_version": "step3.model_spec.v1",
            "status": "pass",
            "source": {key: context["binding"]["source"].get(key) for key in ("source_id", "run_id", "source_sha256")},
            "binding_sha256": context["manifest"]["binding_sha256"],
            "stage_hashes": {"fields": context["fields_sha256"], "dependencies": context["dependencies_sha256"]},
            "evidence_packets": packets,
            "validated_evidence_ids": sorted(evidence_ids),
            "diagnostics": [],
            "readiness": {"draft_only": True, "runtime_verified": False, "generation_ready": False, "gpu_executable": False},
            **{section: spec[section] for section in _SPEC_SECTIONS},
        }
        json_bytes = workflow._json_bytes(result)
        markdown_bytes = _spec_markdown(result).encode("utf-8")
        json_path = context["analysis"] / "model_spec.json"
        markdown_path = context["analysis"] / "model_spec.md"
        # Validation is complete before either prior valid spec artifact can be replaced.
        workflow._write_bytes(json_path, json_bytes)
        workflow._write_bytes(markdown_path, markdown_bytes)
        _attach_spec_handoff(context, result, workflow._hash_bytes(json_bytes), workflow._hash_bytes(markdown_bytes))
        return {"tool": "step3.validate_spec", "status": "pass", "output": "model_spec.json", "model_spec": result,
                "artifacts": {"json": str(json_path), "markdown": str(markdown_path)}}
    except Exception as exc:
        return workflow._blocked("step3.validate_spec", exc)
