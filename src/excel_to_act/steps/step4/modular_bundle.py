"""Build a standalone bundle from an accepted logical semantic plan."""

from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries

from excel_to_act.steps.conversion_workflow import hash_file
from excel_to_act.steps.step3.calculation import CalculationBlocked, _parse_reference
from excel_to_act.steps.step4.modular_compiler import (
    SourceCell,
    build_source_cell_index,
    compile_mapped_families,
    render_grouped_driver,
    validate_family_coverage,
)


def _address_key(address: str) -> str:
    if "!" not in address:
        raise CalculationBlocked(f"source map contains an unqualified address: {address!r}")
    sheet, local = address.rsplit("!", 1)
    bounds = range_boundaries(local.replace("$", ""))
    if any(value is None for value in bounds) or bounds[0] != bounds[2] or bounds[1] != bounds[3]:
        raise CalculationBlocked(f"source member is not a single cell: {address}")
    return f"{sheet.casefold()}!{get_column_letter(bounds[0])}{bounds[1]}"


def _address(sheet: str, row: int, column: int) -> str:
    return f"{sheet}!{get_column_letter(column)}{row}"


def _addresses(address: str) -> list[str]:
    reference = _parse_reference(address, "Main")
    if reference is None or not reference.sheet:
        raise CalculationBlocked(f"source range is not a qualified finite range: {address}")
    return [_address(reference.sheet, row, column)
            for row in range(reference.min_row, reference.max_row + 1)
            for column in range(reference.min_col, reference.max_col + 1)]


def _source_index(plan: dict[str, Any]) -> tuple[dict[tuple[str, int, int], list[SourceCell]], dict[str, dict[str, Any]]]:
    index = build_source_cell_index(plan)
    variables = plan.get("variables", [])
    if isinstance(variables, dict):
        variable_records = variables
    elif isinstance(variables, list):
        variable_records = {item.get("variable_id"): item for item in variables}
    else:
        raise CalculationBlocked("semantic plan has no variable records")
    return index, variable_records


def _metadata_seed_addresses(
    source_index: dict[tuple[str, int, int], list[SourceCell]],
    variables: dict[str, dict[str, Any]],
) -> set[str]:
    """Return seed cells whose accepted source role places them in axis metadata."""
    addresses: set[str] = set()
    for (sheet, row, column), cells in source_index.items():
        accepted = []
        for cell in cells:
            variable = variables.get(cell.variable_id, {})
            initial = variable.get("initial_condition")
            if (cell.role == "initial_condition" and isinstance(initial, dict)
                    and initial.get("source_role") == "accepted_axis_metadata"):
                accepted.append(cell)
        if accepted:
            if len(accepted) != len(cells):
                raise CalculationBlocked("accepted metadata seed overlaps another logical source coordinate")
            addresses.add(_address_key(_address(sheet, row, column)))
    return addresses


def _metadata_seed_bindings(
    plan: dict[str, Any],
    source_index: dict[tuple[str, int, int], list[SourceCell]] | None = None,
    variables: dict[str, dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Bind accepted metadata seeds to checked group slots, separately from raw inputs."""
    if source_index is None or variables is None:
        source_index, variables = _source_index(plan)
    metadata_cells = plan.get("adapter_metadata_cells", [])
    metadata_groups = plan.get("adapter_metadata_groups", {})
    if not isinstance(metadata_cells, list) or not isinstance(metadata_groups, dict):
        raise CalculationBlocked("accepted metadata seed groups are malformed")
    by_address: dict[str, list[dict[str, Any]]] = {}
    for item in metadata_cells:
        if not isinstance(item, dict) or not isinstance(item.get("address"), str):
            raise CalculationBlocked("adapter metadata coordinate is malformed")
        by_address.setdefault(_address_key(item["address"]), []).append(item)

    result = []
    for (sheet, row, column), cells in source_index.items():
        for cell in cells:
            if cell.role != "initial_condition":
                continue
            variable = variables.get(cell.variable_id, {})
            initial = variable.get("initial_condition")
            if not isinstance(initial, dict) or initial.get("source_role") != "accepted_axis_metadata":
                continue
            if variable.get("role") != "derived":
                raise CalculationBlocked("accepted axis metadata may seed only a derived logical variable")
            shape = variable.get("shape")
            if (not isinstance(shape, list)
                    or any(isinstance(size, bool) or not isinstance(size, int) or size < 0
                           for size in shape)
                    or len(cell.indices) != len(shape)
                    or any(isinstance(index, bool) or not isinstance(index, int)
                           or index < 0 or index >= shape[dimension]
                           for dimension, index in enumerate(cell.indices))):
                raise CalculationBlocked(f"accepted metadata seed index is outside its variable: {cell.variable_id}")
            address = _address_key(_address(sheet, row, column))
            matches = by_address.get(address, [])
            if len(matches) != 1:
                raise CalculationBlocked(f"accepted metadata seed has no unique metadata slot: {address}")
            metadata_cell = matches[0]
            group_id = metadata_cell.get("group_id")
            indices = metadata_cell.get("indices")
            group = metadata_groups.get(group_id) if isinstance(group_id, str) else None
            if (not isinstance(group, dict) or group.get("classification") != "axis_metadata"
                    or not isinstance(indices, list) or len(indices) != 1
                    or isinstance(indices[0], bool) or not isinstance(indices[0], int)):
                raise CalculationBlocked(f"accepted metadata seed is not bound to an axis metadata group: {address}")
            group_addresses = group.get("addresses")
            if (not isinstance(group_addresses, list) or indices[0] < 0
                    or indices[0] >= len(group_addresses)
                    or _address_key(group_addresses[indices[0]]) != address):
                raise CalculationBlocked(f"accepted metadata seed slot does not match its source address: {address}")
            result.append({"variable_id": cell.variable_id, "indices": list(cell.indices),
                           "metadata_group": group_id, "metadata_indices": indices,
                           "source_address": address})
    result.sort(key=lambda item: (item["variable_id"], item["indices"], item["source_address"]))
    return result


def _source_axis_name(plan: dict[str, Any], addresses: list[str]) -> str | None:
    coordinates = []
    for address in addresses:
        parsed = _parse_reference(address, "Main")
        if parsed is None or not parsed.single:
            return None
        coordinates.append((parsed.sheet.casefold(), parsed.min_row, parsed.min_col))
    matches: set[str] = set()
    variables = plan.get("variables", {})
    records = variables.values() if isinstance(variables, dict) else variables
    for variable in records:
        for axis in variable.get("axes", []):
            source = _parse_reference(axis.get("source", ""), "Main")
            if source is None:
                continue
            if all(source.sheet.casefold() == sheet and source.min_row <= row <= source.max_row
                   and source.min_col <= col <= source.max_col for sheet, row, col in coordinates):
                name = axis.get("name")
                if isinstance(name, str):
                    matches.add(name)
    return next(iter(matches)) if len(matches) == 1 else None


def add_adapter_metadata(plan: dict[str, Any], catalog: dict[str, Any],
                         model_scope: dict[str, Any], trace: dict[str, Any],
                         source_path: Path, source_sha256: str,
                         candidate_trace: dict[str, Any]) -> dict[str, Any]:
    """Bind catalog metadata coordinates to stable group slots for the private adapter."""
    if plan.get("adapter_metadata_cells") or plan.get("adapter_metadata_groups"):
        raise CalculationBlocked("semantic plan already has adapter metadata bindings")
    source_index, variables = _source_index(plan)
    metadata_seed_addresses = _metadata_seed_addresses(source_index, variables)
    groups: dict[str, dict[str, Any]] = {}
    cells: list[dict[str, Any]] = []

    def add_group(group_id: str, addresses: list[str], axis: str | None,
                  classification: str) -> None:
        canonical = sorted({_address_key(address) for address in addresses},
                           key=lambda address: (address.split("!", 1)[0],
                                               range_boundaries(address.split("!", 1)[1])[1],
                                               range_boundaries(address.split("!", 1)[1])[0]))
        if not canonical:
            return
        # Model variables own overlapping coordinates. Their approved logical binding
        # remains authoritative; metadata groups cover only adapter-only source facts.
        available = []
        for address in canonical:
            sheet, local = address.split("!", 1)
            col, row, _max_col, _max_row = range_boundaries(local)
            if (not source_index.get((sheet, row, col))
                    or _address_key(address) in metadata_seed_addresses):
                available.append(address)
        if not available:
            return
        groups[group_id] = {"addresses": available, "classification": classification}
        for index, address in enumerate(available):
            cells.append({"address": address, "group_id": group_id,
                          "indices": [index], "axes": [axis]})

    for record in catalog.get("axis_metadata", []):
        axis_id = record.get("axis_id")
        if not isinstance(axis_id, str):
            raise CalculationBlocked("confirmed catalog has malformed axis metadata")
        addresses: list[str] = []
        for extent in record.get("source_extents", []):
            parsed = _parse_reference(f"{extent.get('sheet', '')}!{extent.get('range', '')}", "Main")
            if parsed is None:
                raise CalculationBlocked(f"catalog axis extent is invalid: {axis_id}")
            addresses.extend(_address(parsed.sheet, row, col)
                             for row in range(parsed.min_row, parsed.max_row + 1)
                             for col in range(parsed.min_col, parsed.max_col + 1))
        add_group(f"axis.{axis_id}", addresses, _source_axis_name(plan, addresses), "axis_metadata")

    for record in catalog.get("source_records", []):
        record_id = record.get("record_id")
        classification = record.get("classification")
        if (not isinstance(record_id, str)
                or classification not in {"source_constant", "adapter_metadata"}):
            raise CalculationBlocked("confirmed catalog has an unsupported source metadata record")
        addresses = []
        for extent in record.get("source_extents", []):
            parsed = _parse_reference(f"{extent.get('sheet', '')}!{extent.get('range', '')}", "Main")
            if parsed is None:
                raise CalculationBlocked(f"catalog source record extent is invalid: {record_id}")
            addresses.extend(_address(parsed.sheet, row, col)
                             for row in range(parsed.min_row, parsed.max_row + 1)
                             for col in range(parsed.min_col, parsed.max_col + 1))
        add_group(f"source.{record_id}", addresses, None, classification)

    # Scope prose is not proof. Reuse discovery's source-hash, candidate-trace,
    # exact-single-cell and physical formula-free-blank checks before projection.
    from excel_to_act.steps.step4.discovery import _approved_stored_empty_addresses

    try:
        approved_blank_addresses = _approved_stored_empty_addresses(
            source_path, source_sha256, model_scope, candidate_trace
        )
    except (OSError, ValueError) as exc:
        raise CalculationBlocked(str(exc)) from exc
    active_trace_blanks = {
        _address_key(item["address"])
        for item in trace.get("cells", [])
        if isinstance(item, dict) and item.get("role") == "source_value"
        and item.get("value", object()) is None and isinstance(item.get("address"), str)
    }
    if not approved_blank_addresses <= active_trace_blanks:
        raise CalculationBlocked("approved stored-empty coordinate is not a current active-trace blank")
    add_group("source.verified_empty", sorted(approved_blank_addresses), None, "source_audited_blank")
    projected_blank_addresses = {
        _address_key(address)
        for group_id, group in groups.items()
        if group_id == "source.verified_empty"
        for address in group["addresses"]
    }
    if projected_blank_addresses != approved_blank_addresses:
        raise CalculationBlocked("a verified stored-empty coordinate overlaps an existing logical mapping")

    flat = [cell["address"] for cell in cells]
    if len(flat) != len(set(flat)):
        raise CalculationBlocked("catalog metadata groups overlap at an adapter-owned source coordinate")
    plan["adapter_metadata_groups"] = groups
    plan["adapter_metadata_cells"] = cells
    plan["metadata_seed_bindings"] = _metadata_seed_bindings(plan, source_index, variables)
    return {"groups": groups, "cells": cells,
            "source_coordinate_count": len(flat),
            "raw_source_coordinate_count": len({_address_key(item["address"])
                                                  for item in _raw_bindings(plan, source_index, variables)}),
            "source_metadata_coordinate_count": len(flat),
            "metadata_seed_binding_count": len(plan["metadata_seed_bindings"])}


def _raw_literal(value: Any, data_type: str) -> Any:
    if data_type == "e" and isinstance(value, str):
        return {"excel_error": value.upper()}
    if hasattr(value, "isoformat"):
        return {"excel_datetime": value.isoformat(), "date_only": type(value).__name__ == "date"}
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    raise CalculationBlocked(f"raw source value cannot be serialized: {type(value).__name__}")


def _canonical_hash(value: Any) -> str:
    payload = (json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":")) + "\n")
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _external_bindings(plan: dict[str, Any], capture: dict[str, Any] | None,
                       capture_artifact_sha256: str | None,
                       index: dict[tuple[str, int, int], list[SourceCell]]) -> list[dict[str, Any]]:
    variables = plan.get("variables", {})
    if isinstance(variables, list):
        variables = {item.get("variable_id"): item for item in variables}
    if not isinstance(variables, dict):
        raise CalculationBlocked("semantic plan has no logical-variable map")
    external = {key: item for key, item in variables.items()
                if isinstance(item, dict) and item.get("role") == "external"}
    if not external:
        if capture is not None:
            raise CalculationBlocked("external capture is present but the active plan declares no external variables")
        return []
    if (not isinstance(capture, dict) or not isinstance(capture_artifact_sha256, str)
            or capture.get("status") != "pass"
            or capture.get("schema_version") != "step4.external_inputs.v1"
            or capture.get("source_sha256") != plan.get("source", {}).get("source_sha256")):
        raise CalculationBlocked("formula-derived external variables require a current Step 4 capture artifact")
    values = capture.get("external_values")
    if not isinstance(values, dict) or _canonical_hash(values) != capture.get("external_values_sha256"):
        raise CalculationBlocked("external capture values are malformed or have a stale value hash")
    if set(values) != set(external):
        raise CalculationBlocked("external capture variable IDs differ from the approved plan")
    boundaries = {item.get("boundary_variable_id"): item for item in capture.get("boundaries", [])
                  if isinstance(item, dict)}
    if set(boundaries) != set(external):
        raise CalculationBlocked("external capture boundary records differ from the approved plan")
    result: list[dict[str, Any]] = []
    for variable_id, variable in sorted(external.items()):
        shape = variable.get("shape")
        extents = variable.get("source_extents", [])
        boundary = boundaries[variable_id]
        if not isinstance(shape, list) or len(shape) != 1 or not isinstance(extents, list) or len(extents) != 1:
            raise CalculationBlocked(f"external variable must be one declared vector extent: {variable_id}")
        extent = extents[0].get("ref") if isinstance(extents[0], dict) else None
        parsed = _parse_reference(extent or "", "Main")
        boundary_ref = _parse_reference(boundary.get("source_range", ""), "Main")
        if (parsed is None or boundary_ref is None or not parsed.sheet or not boundary_ref.sheet
                or (parsed.sheet.casefold(), parsed.min_row, parsed.max_row, parsed.min_col, parsed.max_col)
                != (boundary_ref.sheet.casefold(), boundary_ref.min_row, boundary_ref.max_row,
                    boundary_ref.min_col, boundary_ref.max_col)
                or shape != boundary.get("shape")):
            raise CalculationBlocked(f"external source extent or shape differs from capture: {variable_id}")
        vector = values[variable_id]
        addresses = [
            _address(parsed.sheet, row, col)
            for row in range(parsed.min_row, parsed.max_row + 1)
            for col in range(parsed.min_col, parsed.max_col + 1)
        ]
        if len(addresses) != shape[0] or not isinstance(vector, list) or len(vector) != shape[0]:
            raise CalculationBlocked(f"external vector length differs from its declared shape: {variable_id}")
        indices: list[list[int]] = []
        for index_in_vector, address in enumerate(addresses):
            key = _address_key(address)
            sheet, local = key.split("!", 1)
            col, row, _max_col, _max_row = range_boundaries(local)
            matches = [item for item in index.get((sheet, row, col), [])
                       if item.variable_id == variable_id and item.role == "external_boundary_input"]
            if len(matches) != 1 or matches[0].indices != (index_in_vector,):
                raise CalculationBlocked(f"external coordinate is not uniquely indexed by its approved axis: {address}")
            indices.append([index_in_vector])
            value = vector[index_in_vector]
            if isinstance(value, bool) or not (
                    value is None or isinstance(value, (str, int, float, dict))):
                raise CalculationBlocked(f"external vector contains an unsupported value at {address}")
            if isinstance(value, float) and not math.isfinite(value):
                raise CalculationBlocked(f"external vector contains a non-finite value at {address}")
            if isinstance(value, dict) and (set(value) != {"excel_error"} or not isinstance(value["excel_error"], str)):
                raise CalculationBlocked(f"external vector contains an invalid error value at {address}")
        result.append({"variable_id": variable_id, "source_range": boundary["source_range"],
                       "shape": shape, "addresses": [_address_key(address) for address in addresses],
                       "indices": indices})
    return result


def _augment_literal_sources(source_path: Path, source_hash: str, plan: dict[str, Any],
                             trace: dict[str, Any], values: dict[str, Any],
                             scenario_guard: dict[str, Any] | None = None) -> dict[str, Any]:
    """Reread declared raw inputs and table boundaries in formula mode."""
    if not source_path.is_file() or hash_file(source_path) != source_hash:
        raise CalculationBlocked("bound workbook is missing or changed before modular input capture")
    requested: set[str] = set()
    required_literals: set[str] = set()
    required_blanks: set[str] = set()
    source_index, variables = _source_index(plan)
    for record in variables.values():
        initial = record.get("initial_condition")
        if isinstance(initial, dict) and isinstance(initial.get("seed_ref"), str):
            if initial.get("source_role") == "accepted_axis_metadata":
                continue
            parsed = _parse_reference(initial["seed_ref"], "Main")
            if parsed is not None and parsed.single:
                address = _address(parsed.sheet, parsed.min_row, parsed.min_col)
                requested.add(address)
                required_literals.add(_address_key(address))
    for binding in _raw_bindings(plan, source_index, variables):
        address = binding["address"]
        requested.add(address)
        required_literals.add(_address_key(address))
    for record in plan.get("reference_resolution", {}).get("named_references", []):
        binding = record.get("table_binding", {})
        if binding.get("kind") == "header_vector":
            for field in ("source_ref",):
                reference = _parse_reference(binding.get(field, record.get("source_ref", "")), "Main")
                if reference is not None:
                    requested.update(_address(reference.sheet, row, column)
                                     for row in range(reference.min_row, reference.max_row + 1)
                                     for column in range(reference.min_col, reference.max_col + 1))
        for field in ("headers_ref", "row_key_ref", "key_ref"):
            reference = _parse_reference(binding.get(field, ""), "Main")
            if reference is not None:
                requested.update(_address(reference.sheet, row, column)
                                 for row in range(reference.min_row, reference.max_row + 1)
                                 for column in range(reference.min_col, reference.max_col + 1))
        for item in binding.get("source_blank_inputs", []):
            reference = _parse_reference(item.get("source_ref", ""), "Main")
            if reference is not None and reference.single:
                address = _address(reference.sheet, reference.min_row, reference.min_col)
                requested.add(address)
                required_blanks.add(_address_key(address))
    for address in (scenario_guard or {}):
        key = _address_key(address)
        requested.add(key)
        required_literals.add(key)
    for record in trace.get("cells", []):
        if isinstance(record, dict) and record.get("role") == "source_value" and isinstance(record.get("address"), str):
            requested.add(record["address"])
    metadata_groups = plan.get("adapter_metadata_groups", {})
    if not isinstance(metadata_groups, dict):
        raise CalculationBlocked("adapter metadata group map is malformed")
    metadata_addresses: set[str] = set()
    for group_id, record in metadata_groups.items():
        addresses = record.get("addresses") if isinstance(record, dict) else None
        if not isinstance(group_id, str) or not isinstance(addresses, list):
            raise CalculationBlocked("adapter metadata group has no coordinate list")
        for address in addresses:
            key = _address_key(address)
            metadata_addresses.add(key)
            requested.add(key)
            if record.get("classification") == "source_audited_blank":
                required_blanks.add(key)
            else:
                required_literals.add(key)
    formula_addresses = {
        _address_key(record["address"])
        for record in trace.get("cells", [])
        if isinstance(record, dict)
        and record.get("role") in {"calculated_formula", "calculated_array_formula"}
        and isinstance(record.get("address"), str)
    }
    overlap = sorted(required_literals & formula_addresses)
    if overlap:
        raise CalculationBlocked("declared raw source boundary overlaps a formula or array member: "
                                 + ", ".join(overlap[:5]))

    # Open formula mode once. The workbook is never saved; normal worksheet access
    # avoids repeatedly rescanning XML for thousands of sparse source coordinates.
    workbook = load_workbook(source_path, data_only=False, read_only=False, keep_vba=True)
    try:
        sheet_lookup = {sheet.title.casefold(): sheet.title for sheet in workbook.worksheets}
        for address in requested:
            key = _address_key(address)
            sheet_name, local = address.rsplit("!", 1)
            title = sheet_lookup.get(sheet_name.casefold())
            if title is None:
                raise CalculationBlocked(f"source support sheet is missing: {sheet_name}")
            cell = workbook[title][local.replace("$", "")]
            if cell.data_type == "f":
                # Formula cells are computed by the generated family driver, never captured as literals.
                if key in required_literals or key in required_blanks:
                    raise CalculationBlocked(f"declared raw source boundary is a formula cell: {address}")
                continue
            literal = _raw_literal(cell.value, cell.data_type)
            if key in required_blanks and literal is not None:
                raise CalculationBlocked(f"declared source blank is not blank: {address}")
            if key in metadata_addresses and key not in required_blanks and literal is None:
                raise CalculationBlocked(f"catalog metadata source coordinate is unexpectedly blank: {address}")
            if key in values and values[key] != literal:
                raise CalculationBlocked(f"trace source value differs from current workbook: {address}")
            values[key] = literal
    finally:
        workbook.close()
    if hash_file(source_path) != source_hash:
        raise CalculationBlocked("bound workbook changed during modular input capture")
    return values


def _unique_binding(index: dict[tuple[str, int, int], list[SourceCell]], address: str,
                    variable_id: str | None = None) -> SourceCell:
    key = _address_key(address)
    sheet, local = key.split("!", 1)
    bounds = range_boundaries(local)
    candidates = index.get((sheet, bounds[1], bounds[0]), [])
    if variable_id is not None:
        candidates = [item for item in candidates if item.variable_id == variable_id]
    unique = {(item.variable_id, item.indices, item.axes): item for item in candidates}
    if len(unique) != 1:
        raise CalculationBlocked(f"formula output coordinate is not uniquely mapped: {address}")
    return next(iter(unique.values()))


def _formula_cell_bindings(plan: dict[str, Any], trace: dict[str, Any],
                           index: dict[tuple[str, int, int], list[SourceCell]]) -> list[dict[str, Any]]:
    records: dict[str, dict[str, Any]] = {}
    for mapping in plan.get("source_mappings", []):
        variable_id = mapping.get("variable_id")
        for address in mapping.get("source_members", []):
            cell = _unique_binding(index, address, variable_id)
            key = _address_key(address)
            records[key] = {"address": key, "variable_id": cell.variable_id,
                            "indices": list(cell.indices), "role": "ordinary_formula"}

    trace_roles = {_address_key(item["address"]): item.get("role") for item in trace.get("cells", [])
                   if isinstance(item, dict) and isinstance(item.get("address"), str)}
    for array_family in plan.get("array_mappings", []):
        for instance in array_family.get("instances", []):
            variable_id = instance.get("variable_id")
            for address in instance.get("source_members", []):
                cell = _unique_binding(index, address, variable_id)
                key = _address_key(address)
                record = {"address": key, "variable_id": cell.variable_id,
                          "indices": list(cell.indices), "role": "array_follower",
                          "array_anchor": _address_key(instance["anchor"])}
                previous = records.get(key)
                if previous is not None and previous != record:
                    raise CalculationBlocked(f"formula cell has conflicting logical mappings: {address}")
                records[key] = record

    expected = {key for key, role in trace_roles.items()
                if role in {"calculated_formula", "calculated_array_formula"}}
    if set(records) != expected:
        raise CalculationBlocked("logical source map differs from active formula cells ("
                                  f"{len(expected - set(records))} unmapped, "
                                  f"{len(set(records) - expected)} extra)")
    return [records[key] for key in sorted(records)]


def _raw_bindings(plan: dict[str, Any], index: dict[tuple[str, int, int], list[SourceCell]],
                  variables: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    bindings: set[tuple[str, str, tuple[int, ...]]] = set()
    for (sheet, row, column), cells in index.items():
        address = _address(sheet, row, column)
        for cell in cells:
            variable = variables.get(cell.variable_id, {})
            initial = variable.get("initial_condition")
            metadata_seed = (cell.role == "initial_condition" and isinstance(initial, dict)
                             and initial.get("source_role") == "accepted_axis_metadata")
            if variable.get("role") == "raw" or (cell.role == "initial_condition" and not metadata_seed):
                bindings.add((address, cell.variable_id, cell.indices))
    return [{"address": address, "variable_id": variable_id, "indices": list(indices)}
            for address, variable_id, indices in sorted(bindings)]


def _coordinate_item(index: dict[tuple[str, int, int], list[SourceCell]],
                     raw_addresses: set[str], address: str,
                     expected_variable: str | None = None) -> dict[str, Any]:
    key = _address_key(address)
    sheet, local = key.split("!", 1)
    col, row, _max_col, _max_row = range_boundaries(local)
    candidates = index.get((sheet, row, col), [])
    if expected_variable:
        matches = [item for item in candidates if item.variable_id == expected_variable]
        if len({(item.variable_id, item.indices) for item in matches}) == 1:
            item = matches[0]
            return {"variable_id": item.variable_id, "indices": list(item.indices)}
    unique = {(item.variable_id, item.indices): item for item in candidates}
    if len(unique) == 1:
        item = next(iter(unique.values()))
        if item.role == "adapter_metadata" and item.metadata_group is not None:
            return {"adapter_metadata_group": item.metadata_group, "indices": list(item.indices)}
        return {"variable_id": item.variable_id, "indices": list(item.indices)}
    if key in raw_addresses:
        return {"address": key}
    return {"missing_address": key}


def _named_bindings(plan: dict[str, Any], index: dict[tuple[str, int, int], list[SourceCell]],
                    raw_addresses: set[str]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    references = plan.get("reference_resolution", {}).get("named_references", [])
    for item in references:
        name = item.get("name")
        if not isinstance(name, str) or not name:
            continue
        key = name.casefold()
        if key in result:
            raise CalculationBlocked(f"case-insensitive duplicate named binding: {name}")
        if item.get("kind") == "variable_coordinate":
            variable_id, indices = item.get("variable_id"), item.get("indices")
            if (not isinstance(variable_id, str) or not isinstance(indices, list)
                    or any(not isinstance(index, int) or index < 0 for index in indices)):
                raise CalculationBlocked(f"named coordinate {name} has an invalid logical binding")
            result[key] = {"kind": "coordinate", "variable_id": variable_id, "indices": indices}
            continue
        if item.get("kind") == "source_literal":
            address = item.get("source_ref")
            if not isinstance(address, str) or _address_key(address) not in raw_addresses:
                raise CalculationBlocked(f"named scalar {name} source literal is not in declared source values")
            result[key] = {"kind": "source_literal", "address": _address_key(address)}
            continue
        if item.get("kind") == "scalar":
            binding = item.get("binding")
            if binding is None:
                binding = {"kind": "variable", "variable_id": item.get("variable_id")}
            if not isinstance(binding, dict):
                raise CalculationBlocked(f"named scalar {name} has a malformed binding")
            binding_kind = binding.get("kind")
            if binding_kind in {"variable", "variable_coordinate"}:
                variable_id = binding.get("variable_id")
                indices = binding.get("indices", [])
                if not isinstance(variable_id, str) or not isinstance(indices, list):
                    raise CalculationBlocked(f"named scalar {name} has an invalid logical binding")
                if binding_kind == "variable_coordinate":
                    result[key] = {"kind": "coordinate", "variable_id": variable_id, "indices": indices}
                elif indices:
                    raise CalculationBlocked(f"named scalar {name} uses indices without variable_coordinate")
                else:
                    result[key] = {"kind": "variable", "variable_id": variable_id}
            elif binding_kind == "source_literal":
                address = binding.get("address", binding.get("source_ref"))
                if not isinstance(address, str) or _address_key(address) not in raw_addresses:
                    raise CalculationBlocked(f"named scalar {name} source literal is not in declared source values")
                result[key] = {"kind": "source_literal", "address": _address_key(address)}
            else:
                raise CalculationBlocked(f"named scalar {name} has unsupported binding kind: {binding_kind!r}")
            continue
        reference = item.get("source_ref")
        cell_range = _parse_reference(reference or "", "Main")
        if cell_range is None:
            result[key] = {"kind": "missing", "name": name}
            continue
        table_binding = item.get("table_binding", {})
        table_kind = table_binding.get("kind")
        if table_kind == "header_vector":
            addresses = [_address(cell_range.sheet, row, column)
                         for row in range(cell_range.min_row, cell_range.max_row + 1)
                         for column in range(cell_range.min_col, cell_range.max_col + 1)]
            result[key] = {"kind": "vector", "items": [
                _coordinate_item(index, raw_addresses, address) for address in addresses]}
            continue
        row_count = cell_range.max_row - cell_range.min_row + 1
        column_count = cell_range.max_col - cell_range.min_col + 1
        columns: dict[str, list[dict[str, Any]]] = {}
        formula_columns = table_binding.get("formula_columns", {})
        column_variables = table_binding.get("column_variables", {})
        value_column_index = table_binding.get("value_column_index")
        value_variable_id = table_binding.get("value_variable_id")
        key_column_index = table_binding.get("key_column_index", 1)
        key_variable_id = table_binding.get("key_variable_id")
        raw_variable_id = table_binding.get("raw_variable_id")
        for column_offset in range(column_count):
            absolute_column = cell_range.min_col + column_offset
            column_letter = get_column_letter(absolute_column)
            expected_variable = formula_columns.get(column_letter) or column_variables.get(column_letter)
            relative_index = column_offset + 1
            if relative_index == value_column_index and isinstance(value_variable_id, str):
                expected_variable = value_variable_id
            if relative_index == key_column_index and isinstance(key_variable_id, str):
                expected_variable = key_variable_id
            if expected_variable is None and isinstance(raw_variable_id, str):
                expected_variable = raw_variable_id
            cells = [_address(cell_range.sheet, row, absolute_column)
                     for row in range(cell_range.min_row, cell_range.max_row + 1)]
            columns[str(relative_index)] = [
                _coordinate_item(index, raw_addresses, address, expected_variable) for address in cells]
        result[key] = {"kind": "table", "row_count": row_count,
                       "column_count": column_count, "columns": columns,
                       "table_kind": table_kind}
    return result


def _target_bindings(target_names: list[str], plan: dict[str, Any]) -> dict[str, str]:
    refs = {item.get("name", "").casefold(): item
            for item in plan.get("reference_resolution", {}).get("named_references", [])
            if isinstance(item.get("name"), str)}
    result: dict[str, str] = {}
    for name in target_names:
        record = refs.get(name.casefold())
        if not isinstance(record, dict) or not isinstance(record.get("variable_id"), str):
            raise CalculationBlocked(f"approved target has no scalar logical binding: {name}")
        result[name] = record["variable_id"]
    return result


def _adapter_source() -> str:
    return '''"""Standalone source-address to logical-variable adapter."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import modular_runtime as runtime

def _canonical_hash(value):
    payload = json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":")) + "\\n"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()

def _allocate(shape, missing):
    if not shape:
        return missing
    if len(shape) == 1:
        return [runtime.MissingSourceValue(missing) for _ in range(shape[0])]
    if len(shape) == 2:
        return [[runtime.MissingSourceValue(missing) for _ in range(shape[1])]
                for _ in range(shape[0])]
    raise runtime.CalculationBlocked("standalone adapter supports at most two logical axes")

def _assign(values, variable_id, indices, value):
    if not indices:
        values[variable_id] = value
        return
    target = values[variable_id]
    for index in indices[:-1]:
        target = target[index]
    target[indices[-1]] = value

def build_values(source_values=None, scenario_guard_sha256=None, scenario_ids=None,
                 metadata_seed_bindings_sha256=None):
    base = Path(__file__).resolve().parent
    if source_values is None:
        source_values = json.loads((base / "source_values.json").read_text(encoding="utf-8"))
    layout = json.loads((base / "input_layout.json").read_text(encoding="utf-8"))
    guard = layout.get("scenario_guard", {})
    if scenario_ids is not None and scenario_ids != guard.get("scenario_ids"):
        raise runtime.CalculationBlocked("bundle scenario identity differs from its generated guard")
    guard_values = guard.get("source_values")
    if (not isinstance(guard_values, dict) or not isinstance(scenario_guard_sha256, str)
            or _canonical_hash(guard_values) != scenario_guard_sha256):
        raise runtime.CalculationBlocked("bundle scenario guard differs from its generated contract")
    for address, expected in guard_values.items():
        if address not in source_values or source_values[address] != expected:
            raise runtime.CalculationBlocked(f"saved-scenario source selector changed: {address}")
    values = {"__source_values__": source_values}
    metadata_groups = layout.get("adapter_metadata_groups", {})
    metadata_values = {}
    if metadata_groups:
        metadata_payload = json.loads((base / "adapter_metadata.json").read_text(encoding="utf-8"))
        groups = metadata_payload.get("groups", {})
        if (metadata_payload.get("schema_version") != "step4.adapter_metadata.v1"
                or metadata_payload.get("source_sha256") != layout["source_sha256"]
                or not isinstance(groups, dict) or set(groups) != set(metadata_groups)):
            raise runtime.CalculationBlocked("adapter metadata file is stale or incomplete")
        metadata_values = {}
        for group_id, expected in metadata_groups.items():
            record = groups.get(group_id)
            items = record.get("values") if isinstance(record, dict) else None
            if (not isinstance(expected, dict) or not isinstance(record, dict)
                    or not isinstance(items, list) or len(items) != expected.get("length")
                    or record.get("addresses") != expected.get("addresses")
                    or _canonical_hash(items) != expected.get("sha256")
                    or record.get("sha256") != expected.get("sha256")):
                raise runtime.CalculationBlocked(f"adapter metadata group changed: {group_id}")
            metadata_values[group_id] = items
        values["__adapter_metadata__"] = metadata_values
    for variable_id, shape in layout["variable_shapes"].items():
        values[variable_id] = _allocate(shape, f"uninitialized logical value: {variable_id}")
    for binding in layout["raw_bindings"]:
        address = binding["address"]
        value = source_values.get(address, runtime.MissingSourceValue(address))
        _assign(values, binding["variable_id"], binding["indices"], value)
    metadata_seed_bindings = layout.get("metadata_seed_bindings")
    if (not isinstance(metadata_seed_bindings, list)
            or not isinstance(metadata_seed_bindings_sha256, str)
            or _canonical_hash(metadata_seed_bindings) != metadata_seed_bindings_sha256):
        raise runtime.CalculationBlocked("metadata seed bindings are missing or differ from the generated contract")
    metadata_layout = layout.get("adapter_metadata_groups", {})
    for binding in metadata_seed_bindings:
        if not isinstance(binding, dict):
            raise runtime.CalculationBlocked("metadata seed binding is malformed")
        variable_id = binding.get("variable_id")
        indices = binding.get("indices")
        metadata_group = binding.get("metadata_group")
        metadata_indices = binding.get("metadata_indices")
        source_address = binding.get("source_address")
        expected_group = metadata_layout.get(metadata_group) if isinstance(metadata_group, str) else None
        if (not isinstance(variable_id, str) or variable_id not in values
                or not isinstance(indices, list) or not isinstance(metadata_indices, list)
                or not isinstance(source_address, str) or not isinstance(expected_group, dict)
                or metadata_group not in metadata_values):
            raise runtime.CalculationBlocked("metadata seed binding has no checked value group")
        shape = layout["variable_shapes"].get(variable_id)
        if (not isinstance(shape, list) or len(indices) != len(shape)
                or any(not isinstance(index, int) or isinstance(index, bool)
                       or index < 0 or index >= shape[dimension]
                       for dimension, index in enumerate(indices))):
            raise runtime.CalculationBlocked(f"metadata seed index is outside its variable: {variable_id}")
        group_addresses = expected_group.get("addresses")
        if (not isinstance(group_addresses, list) or len(metadata_indices) != 1
                or not isinstance(metadata_indices[0], int) or isinstance(metadata_indices[0], bool)
                or metadata_indices[0] < 0 or metadata_indices[0] >= len(group_addresses)
                or group_addresses[metadata_indices[0]] != source_address):
            raise runtime.CalculationBlocked(f"metadata seed does not match its checked source slot: {source_address}")
        seed_value = metadata_values[metadata_group]
        for index in metadata_indices:
            if not isinstance(seed_value, list) or index >= len(seed_value):
                raise runtime.CalculationBlocked(f"metadata seed group is missing its source slot: {metadata_group}")
            seed_value = seed_value[index]
        _assign(values, variable_id, indices, seed_value)
    external_bindings = layout.get("external_bindings", [])
    if external_bindings:
        external_path = base / "external_values.json"
        external_payload = json.loads(external_path.read_text(encoding="utf-8"))
        expected_capture = layout.get("external_capture")
        if (not isinstance(expected_capture, dict)
                or external_payload.get("schema_version") != "step4.external_values.v1"
                or external_payload.get("source_sha256") != layout["source_sha256"]
                or external_payload.get("capture_artifact_sha256") != expected_capture.get("artifact_sha256")
                or external_payload.get("capture_sha256") != expected_capture.get("capture_sha256")
                or external_payload.get("values_sha256") != expected_capture.get("values_sha256")):
            raise runtime.CalculationBlocked("external input file is stale or bound to another capture")
        external_values = external_payload.get("values")
        if not isinstance(external_values, dict) or _canonical_hash(external_values) != expected_capture.get("values_sha256"):
            raise runtime.CalculationBlocked("external input values are incomplete or changed")
        expected_ids = {binding["variable_id"] for binding in external_bindings}
        if set(external_values) != expected_ids:
            raise runtime.CalculationBlocked("external input IDs differ from the generated logical bindings")
        for binding in external_bindings:
            variable_id = binding["variable_id"]
            vector = external_values[variable_id]
            if (not isinstance(vector, list) or len(vector) != len(binding["indices"])
                    or len(vector) != binding["shape"][0]):
                raise runtime.CalculationBlocked(f"external vector shape changed: {variable_id}")
            for indices, item in zip(binding["indices"], vector):
                _assign(values, variable_id, indices, item)
    elif layout.get("external_capture") is not None:
        raise runtime.CalculationBlocked("bundle external capture has no external logical bindings")
    values["__names__"] = layout["named_bindings"]
    return values
'''


def _model_source(scenario_ids: list[str], scenario_guard_sha256: str,
                  metadata_seed_bindings_sha256: str) -> str:
    template = '''"""Run the standalone named-variable model for its approved saved scenario."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import modular_runtime as runtime
from input_adapter import build_values
from pricing import run_pricing
from formula_families import FAMILY_FUNCTIONS

SCENARIO_IDS = __SCENARIO_IDS__
SCENARIO_GUARD_SHA256 = __SCENARIO_GUARD_SHA256__
METADATA_SEED_BINDINGS_SHA256 = __METADATA_SEED_BINDINGS_SHA256__

def _json_value(value):
    if isinstance(value, runtime.ExcelError):
        return {"excel_error": value.code}
    if isinstance(value, runtime.MissingSourceValue):
        raise runtime.CalculationBlocked(f"active output is missing: {value.address}")
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    if isinstance(value, dict) and "excel_error" in value:
        return value
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    raise runtime.CalculationBlocked(f"model output cannot serialize {type(value).__name__}")

def run_model():
    base = Path(__file__).resolve().parent
    layout = json.loads((base / "input_layout.json").read_text(encoding="utf-8"))
    values = build_values(scenario_guard_sha256=SCENARIO_GUARD_SHA256, scenario_ids=SCENARIO_IDS,
                          metadata_seed_bindings_sha256=METADATA_SEED_BINDINGS_SHA256)
    values = run_pricing(values, FAMILY_FUNCTIONS)
    targets = {name: _json_value(runtime.at(values, variable_id, []))
               for name, variable_id in layout["target_bindings"].items()}
    calculated_variables = {}
    for item in layout["formula_cells"]:
        variable_id = item["variable_id"]
        record = calculated_variables.setdefault(variable_id, {
            "shape": layout["variable_shapes"][variable_id], "demanded_values": []})
        record["demanded_values"].append({
            "indices": item["indices"],
            "value": _json_value(runtime.at(values, variable_id, item["indices"])),
        })
    for record in calculated_variables.values():
        record["demanded_values"].sort(key=lambda item: tuple(item["indices"]))
    cells = {item["address"]: _json_value(runtime.at(values, item["variable_id"], item["indices"]))
             for item in layout["formula_cells"]}
    external_inputs = {binding["variable_id"]: _json_value(values[binding["variable_id"]])
                       for binding in layout.get("external_bindings", [])}
    age_axis = None
    external_capture = layout.get("external_capture")
    age_record = external_capture.get("age_axis") if isinstance(external_capture, dict) else None
    if isinstance(age_record, dict) and age_record.get("metadata_group"):
        age_axis = _json_value(values["__adapter_metadata__"][age_record["metadata_group"]])
    return {"schema_version": "step4.modular_model_result.v1", "status": "pass",
            "source_sha256": layout["source_sha256"], "targets": targets,
            "calculated_variables": calculated_variables, "cells": cells,
            "external_inputs": external_inputs, "external_age_axis": age_axis,
            "calculated_variable_count": len(calculated_variables),
            "calculated_value_count": sum(len(item["demanded_values"])
                                           for item in calculated_variables.values()),
            "formula_count": len(cells), "cell_count": len(cells)}

def main():
    parser = argparse.ArgumentParser(description="Run the standalone logical-variable model")
    parser.add_argument("--out", type=Path, default=Path("model_result.json"))
    args = parser.parse_args()
    result = run_model()
    args.out.write_text(json.dumps(result, indent=2, sort_keys=True) + "\\n", encoding="utf-8")

if __name__ == "__main__":
    main()
'''
    return (template.replace("__SCENARIO_IDS__", repr(scenario_ids))
            .replace("__SCENARIO_GUARD_SHA256__", repr(scenario_guard_sha256))
            .replace("__METADATA_SEED_BINDINGS_SHA256__", repr(metadata_seed_bindings_sha256)))


def build_modular_bundle_files(
    *,
    semantic_plan: dict[str, Any],
    source_profile: dict[str, Any],
    trace: dict[str, Any],
    raw_source_values: dict[str, Any],
    source_path: Path,
    semantic_plan_path: Path,
    source_profile_path: Path,
    semantic_map_path: Path,
    revision_dir: Path,
    target_names: list[str],
    path_coverage: dict[str, Any],
    source_sha256: str,
    design_sha256: str,
    semantic_plan_sha256: str,
    source_profile_sha256: str,
    trace_sha256: str,
    trace_path: Path,
    trace_source: str,
    scenario_ids: list[str],
    external_capture: dict[str, Any] | None = None,
    external_capture_artifact_sha256: str | None = None,
    scenario_guard: dict[str, Any] | None = None,
    implementation_plan_record: dict[str, Any] | None = None,
    implementation_plan_path: Path | None = None,
    implementation_plan_sha256: str | None = None,
) -> dict[str, Any]:
    """Compile, package, and report a source-bound logical-variable bundle."""
    coverage = validate_family_coverage(semantic_plan, source_profile, trace)
    if semantic_plan.get("source", {}).get("source_sha256") != source_sha256:
        raise CalculationBlocked("semantic plan source hash differs from approved workbook")
    if not isinstance(raw_source_values, dict):
        raise CalculationBlocked("raw source-value ledger is not an object")
    source_values = {_address_key(key): value for key, value in raw_source_values.items()}
    scenario_guard = {_address_key(address): value for address, value in (scenario_guard or {}).items()}
    adapter_groups = semantic_plan.get("adapter_metadata_groups", {})
    if not isinstance(adapter_groups, dict):
        raise CalculationBlocked("adapter metadata groups are malformed")
    metadata_addresses = {_address_key(address) for group in adapter_groups.values()
                          for address in group.get("addresses", [])}
    all_source_values = _augment_literal_sources(source_path, source_sha256, semantic_plan, trace,
                                                 source_values, scenario_guard)
    changed_scenario = [address for address, expected in scenario_guard.items()
                        if all_source_values.get(address) != expected]
    if changed_scenario:
        raise CalculationBlocked("source inputs differ from the approved saved scenario: "
                                 + ", ".join(sorted(changed_scenario)[:8]))
    index, variables = _source_index(semantic_plan)
    raw_bindings = _raw_bindings(semantic_plan, index, variables)
    metadata_seed_bindings = _metadata_seed_bindings(semantic_plan, index, variables)
    declared_metadata_seeds = semantic_plan.get("metadata_seed_bindings")
    if declared_metadata_seeds is not None and declared_metadata_seeds != metadata_seed_bindings:
        raise CalculationBlocked("metadata seed bindings differ from the current source-role declarations")
    raw_binding_addresses = {_address_key(item["address"]) for item in raw_bindings}
    # Step 3 approves this saved scenario only. Bind every declared raw business
    # coordinate so changed selectors and configuration tables cannot activate
    # producer branches that were excluded from the reviewed model.
    scenario_guard = {_address_key(address): value for address, value in scenario_guard.items()}
    for binding in raw_bindings:
        address = _address_key(binding["address"])
        if address not in all_source_values:
            raise CalculationBlocked(f"saved-scenario raw input is missing: {address}")
        scenario_guard[address] = all_source_values[address]
    source_literals = {_address_key(address) for address in semantic_plan.get("adapter_source_literals", [])
                       if isinstance(address, str)}
    if metadata_addresses & raw_binding_addresses:
        raise CalculationBlocked("adapter metadata overlaps a declared raw business coordinate")
    unclassified = set(all_source_values) - raw_binding_addresses - metadata_addresses - source_literals
    if unclassified:
        raise CalculationBlocked("source values are outside the approved logical inputs and metadata: "
                                  + ", ".join(sorted(unclassified)[:8]))
    missing_raw = raw_binding_addresses - set(all_source_values)
    if missing_raw:
        raise CalculationBlocked("approved raw business coordinates are missing: "
                                  + ", ".join(sorted(missing_raw)[:8]))
    source_values = {address: value for address, value in all_source_values.items()
                     if address in raw_binding_addresses or address in source_literals}
    raw_addresses = set(source_values)
    metadata_payload_groups = {}
    for group_id, group in sorted(adapter_groups.items()):
        addresses = [_address_key(address) for address in group.get("addresses", [])]
        values = [all_source_values[address] for address in addresses]
        metadata_payload_groups[group_id] = {"addresses": addresses, "values": values,
                                             "sha256": _canonical_hash(values)}
    formula_cells = _formula_cell_bindings(semantic_plan, trace, index)
    names = _named_bindings(semantic_plan, index, raw_addresses)
    external_bindings = _external_bindings(semantic_plan, external_capture,
                                           external_capture_artifact_sha256, index)
    external_values = external_capture.get("external_values", {}) if external_capture else {}
    external_value_hash = external_capture.get("external_values_sha256") if external_capture else None
    external_capture_hash = external_capture.get("external_capture_sha256") if external_capture else None
    external_age_axis = None
    if external_capture is not None:
        age_axis = external_capture.get("age_axis")
        if (not isinstance(age_axis, dict) or not isinstance(age_axis.get("source_range"), str)
                or not isinstance(age_axis.get("values"), list)):
            raise CalculationBlocked("external capture has no source-bound age axis")
        age_addresses = {_address_key(address) for address in _addresses(age_axis["source_range"])}
        matching_groups = [group_id for group_id, group in adapter_groups.items()
                           if {_address_key(address) for address in group.get("addresses", [])} == age_addresses]
        if len(matching_groups) != 1:
            raise CalculationBlocked("external age axis does not map to one adapter metadata group")
        external_age_axis = {"source_range": age_axis["source_range"],
                             "values_sha256": age_axis.get("value_sha256"),
                             "metadata_group": matching_groups[0]}
    implementation_binding = None
    if any(value is not None for value in
           (implementation_plan_record, implementation_plan_path, implementation_plan_sha256)):
        if (implementation_plan_record is None or implementation_plan_path is None
                or not isinstance(implementation_plan_sha256, str)
                or not implementation_plan_path.is_file()
                or hash_file(implementation_plan_path) != implementation_plan_sha256):
            raise CalculationBlocked("modular bundle implementation preflight binding is incomplete or stale")
        implementation_binding = {"path": str(implementation_plan_path),
                                  "sha256": implementation_plan_sha256,
                                  "revision": implementation_plan_record.get("revision")}
    targets = _target_bindings(target_names, semantic_plan)
    def resolve_name(name: str, sheet: str, owner: str) -> str:
        if name.casefold() not in names:
            raise CalculationBlocked(f"named reference is not mapped: {name}")
        return f"runtime.lookup_name(values, {name.casefold()!r})"

    compiled = compile_mapped_families(semantic_plan, trace, resolve_name)
    driver = render_grouped_driver(semantic_plan, compiled)
    if compiled["compiled_member_count"] != coverage["ordinary_formula_count"]:
        raise CalculationBlocked("compiled ordinary formula count differs from source coverage")
    metadata_group_hashes = {
        group_id: {"addresses": record["addresses"], "length": len(record["values"]),
                   "sha256": record["sha256"]}
        for group_id, record in metadata_payload_groups.items()
    }
    layout = {
        "source_sha256": source_sha256,
        "variable_shapes": {key: value.get("shape", []) for key, value in sorted(variables.items())},
        "raw_bindings": raw_bindings,
        "metadata_seed_bindings": metadata_seed_bindings,
        "adapter_metadata_groups": metadata_group_hashes,
        "external_bindings": external_bindings,
        "external_capture": ({"artifact_sha256": external_capture_artifact_sha256,
                               "capture_sha256": external_capture_hash,
                               "values_sha256": external_value_hash,
                               "age_axis": external_age_axis}
                              if external_capture is not None else None),
        "implementation_plan": implementation_binding,
        "scenario_guard": {"scenario_ids": scenario_ids, "source_values": scenario_guard},
        "formula_cells": formula_cells,
        "named_bindings": names,
        "target_bindings": targets,
    }
    equation_families = semantic_plan.get("equation_families", [])
    equation_group_count = (len(equation_families) if isinstance(equation_families, (dict, list)) else 0)
    family_manifest = {
        "schema_version": "step4.family_manifest.v1",
        "source_family_count": compiled["source_family_count"],
        "semantic_equation_group_count": equation_group_count,
        "compiled_variant_count": compiled["function_count"],
        "ordinary_formula_count": coverage["ordinary_formula_count"],
        "array_follower_count": coverage["array_follower_count"],
        "array_instance_count": coverage["array_instance_count"],
        "active_formula_member_count": coverage["active_formula_member_count"],
        "unmapped_count": 0,
        "source_family_to_function": compiled["source_family_to_function"],
        "source_family_variants": compiled["source_family_variants"],
        "member_function_exceptions": compiled["member_function_exceptions"],
        "families": semantic_plan.get("source_mappings", []),
        "arrays": semantic_plan.get("array_mappings", []),
    }
    bundle = revision_dir / "bundle"
    bundle.mkdir(parents=True)
    runtime_source = Path(__file__).with_name("modular_runtime.py").read_bytes()
    (bundle / "modular_runtime.py").write_bytes(runtime_source)
    (bundle / "input_adapter.py").write_text(_adapter_source(), encoding="utf-8")
    (bundle / "input_layout.json").write_text(
        json.dumps(layout, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    (bundle / "source_values.json").write_text(
        json.dumps(source_values, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    if metadata_payload_groups:
        metadata_payload = {"schema_version": "step4.adapter_metadata.v1",
                            "source_sha256": source_sha256,
                            "origin": "formula_mode_source", "formula_cache_inputs": False,
                            "groups": metadata_payload_groups}
        (bundle / "adapter_metadata.json").write_text(
            json.dumps(metadata_payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
            encoding="utf-8")
    if external_capture is not None:
        external_payload = {
            "schema_version": "step4.external_values.v1",
            "source_sha256": source_sha256,
            "capture_artifact_sha256": external_capture_artifact_sha256,
            "capture_sha256": external_capture_hash,
            "values_sha256": external_value_hash,
            "values": external_values,
            "boundaries": external_capture.get("boundaries", []),
            "formula_cut_members": external_capture.get("formula_cut_members", []),
            "formula_cut_formula_set_sha256": external_capture.get("formula_cut_formula_set_sha256"),
            "age_axis": external_capture.get("age_axis"),
        }
        (bundle / "external_values.json").write_text(
            json.dumps(external_payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    (bundle / "source_map.json").write_text(
        json.dumps({"semantic_plan_sha256": semantic_plan_sha256,
                    "formula_cells": formula_cells,
                    "external_bindings": external_bindings,
                    "adapter_metadata_cells": semantic_plan.get("adapter_metadata_cells", []),
                    "source_mappings": semantic_plan.get("source_mappings", []),
                    "array_mappings": semantic_plan.get("array_mappings", [])},
                   indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    family_source = "import modular_runtime as runtime\n\n" + "\n\n".join(
        compiled["functions"][name] for name in sorted(compiled["functions"]))
    family_source += "\n\nFAMILY_FUNCTIONS = {\n" + "\n".join(
        f"    {name!r}: {name}," for name in sorted(compiled["functions"])) + "\n}\n"
    (bundle / "formula_families.py").write_text(family_source, encoding="utf-8")
    pricing_source = ("from formula_families import FAMILY_FUNCTIONS\n\n" + driver +
                      "\ndef run(values):\n    return run_pricing(values, FAMILY_FUNCTIONS)\n")
    (bundle / "pricing.py").write_text(pricing_source, encoding="utf-8")
    metadata_seed_bindings_sha256 = _canonical_hash(metadata_seed_bindings)
    (bundle / "model.py").write_text(
        _model_source(scenario_ids, _canonical_hash(scenario_guard),
                      metadata_seed_bindings_sha256), encoding="utf-8")
    (bundle / "family_manifest.json").write_text(
        json.dumps(family_manifest, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    files = {path.name: hash_file(path) for path in sorted(bundle.iterdir()) if path.is_file()}
    raw_business_input_count = sum(item.get("role") == "raw" for item in variables.values())
    external_business_input_count = sum(item.get("role") == "external" for item in variables.values())
    derived_variable_count = sum(item.get("role") == "derived" for item in variables.values())
    manifest = {
        "schema_version": "step4.modular_model.v1", "execution_kind": "modular",
        "source_sha256": source_sha256,
        "design_sha256": design_sha256, "semantic_plan_sha256": semantic_plan_sha256,
        "source_profile_sha256": source_profile_sha256, "trace_sha256": trace_sha256,
        "trace_path": str(trace_path), "trace_source": trace_source, "scenario_ids": scenario_ids,
        "target_names": target_names,
        "formula_count": coverage["active_formula_member_count"],
        "array_member_count": coverage["array_follower_count"],
        "formula_addresses": [item["address"] for item in formula_cells],
        "raw_input_addresses": sorted(source_values),
        "external_input_addresses": sorted(address for binding in external_bindings
                                             for address in binding["addresses"]),
        "active_formula_addresses": [item["address"] for item in formula_cells],
        "ordinary_formula_count": coverage["ordinary_formula_count"],
        "array_follower_count": coverage["array_follower_count"],
        "array_instance_count": coverage["array_instance_count"],
        "active_formula_member_count": coverage["active_formula_member_count"],
        "source_family_count": compiled["source_family_count"],
        "semantic_equation_group_count": family_manifest["semantic_equation_group_count"],
        "compiled_variant_count": compiled["function_count"],
        "business_input_count": raw_business_input_count + external_business_input_count,
        "raw_business_input_count": raw_business_input_count,
        "external_input_count": sum(len(binding["addresses"]) for binding in external_bindings),
        "external_business_input_count": external_business_input_count,
        "model_variable_count": len(variables), "derived_variable_count": derived_variable_count,
        "raw_source_coordinate_count": len(raw_binding_addresses),
        "source_metadata_coordinate_count": len(metadata_addresses),
        "metadata_seed_binding_count": len(metadata_seed_bindings),
        "raw_input_count": len(raw_binding_addresses),
        "external_capture_sha256": external_capture_hash,
        "external_capture_artifact_sha256": external_capture_artifact_sha256,
        "external_values_sha256": external_value_hash,
        "external_input_ranges": ([item["source_range"] for item in external_capture.get("boundaries", [])]
                                  if external_capture else []),
        "external_age_axis": (external_age_axis if external_capture else None),
        "external_formula_cut_member_count": (external_capture.get("formula_cut_member_count")
                                               if external_capture else 0),
        "external_formula_cut_formula_set_sha256": (external_capture.get("formula_cut_formula_set_sha256")
                                                       if external_capture else None),
        "external_age_axis_sha256": (external_capture.get("age_axis", {}).get("value_sha256")
                                      if external_capture else None),
        "implementation_plan_sha256": implementation_plan_sha256,
        "scenario_guard_sha256": _canonical_hash(scenario_guard),
        "scenario_guard_coordinate_count": len(scenario_guard),
        "unmapped_count": 0,
        "files": files, "runtime_dependencies": ["Python standard library only"],
        "source_workbook_at_runtime": False, "project_package_at_runtime": False,
        "formula_cache_inputs": False,
        "scope": "approved saved scenario and source-bound logical plan",
    }
    (bundle / "model_manifest.json").write_text(
        json.dumps(manifest, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    files["model_manifest.json"] = hash_file(bundle / "model_manifest.json")
    return {"bundle": str(bundle), "bundle_hashes": files, "manifest": manifest,
            "coverage": coverage, "trace": trace, "trace_path": trace_path,
            "trace_hash": trace_sha256, "trace_source": trace_source,
            "source_hash": source_sha256, "source_path": str(source_path),
            "semantic_plan_path": str(semantic_plan_path),
            "source_profile_path": str(source_profile_path), "semantic_map_path": str(semantic_map_path),
            "target_names": target_names,
            "formula_count": coverage["active_formula_member_count"],
            "business_input_count": raw_business_input_count + external_business_input_count,
            "raw_business_input_count": raw_business_input_count,
            "external_business_input_count": external_business_input_count,
            "model_variable_count": len(variables), "derived_variable_count": derived_variable_count,
            "raw_source_coordinate_count": len(raw_binding_addresses),
            "source_metadata_coordinate_count": len(metadata_addresses),
            "raw_input_count": len(raw_binding_addresses), "array_member_count": coverage["array_follower_count"],
            "external_input_count": sum(len(binding["addresses"]) for binding in external_bindings),
            "function_count": compiled["function_count"], "source_family_count": compiled["source_family_count"],
            "path_coverage": path_coverage}
