"""Static Step 3 drafts; this module never evaluates workbook formulas."""

from __future__ import annotations

from collections import Counter, defaultdict
from bisect import bisect_left, bisect_right
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.cell import range_boundaries

from excel_to_act.graph.builder import RegexFormulaGraphBuilder
from excel_to_act.schemas import Step2Index, WorkbookInventory
from excel_to_act.steps.step2.workflow import _validate_saved_index

_ANALYSIS = "analysis.json"
_FIELDS = "fields.json"
_DEPENDENCIES = "dependencies.json"
_PLAN = "execution_plan.json"
_STAGE_FILES = {"fields": _FIELDS, "dependencies": _DEPENDENCIES, "plan": _PLAN,
                "semantic_plan": "semantic_plan.json", "semantic_check": "semantic_check.json"}
_A1 = re.compile(r"^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$", re.IGNORECASE)


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode("utf-8")


def _hash_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_bytes(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_bytes(payload)
    os.replace(temporary, path)


def _write_json(path: Path, value: Any) -> str:
    payload = _json_bytes(value)
    _write_bytes(path, payload)
    return _hash_bytes(payload)


def _blocked(tool: str, exc: Exception) -> dict[str, Any]:
    return {
        "tool": tool,
        "status": "blocked",
        "diagnostics": [{"code": "step3_invalid_input", "severity": "error", "message": str(exc)}],
    }


def _tool_result(tool: str, analysis_dir: Path, **values: Any) -> dict[str, Any]:
    return {"tool": tool, "status": "pass", "analysis": str(analysis_dir), **values}


def _resolved_child(root: Path, relative: str, label: str) -> Path:
    candidate = Path(relative)
    if candidate.is_absolute():
        raise ValueError(f"{label} must be relative to the checked Step 1 root")
    resolved = (root / candidate).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError as exc:
        raise ValueError(f"{label} resolves outside the checked Step 1 root") from exc
    return resolved


def _a1_cell(address: str | None) -> tuple[int, int] | None:
    if not isinstance(address, str):
        return None
    match = _A1.fullmatch(address.replace("$", ""))
    if not match:
        return None
    try:
        from openpyxl.utils.cell import column_index_from_string

        return int(match.group(2)), column_index_from_string(match.group(1))
    except ValueError:
        return None


def _split_qualified(value: str, fallback_sheet: str | None = None) -> tuple[str | None, str]:
    if "!" not in value:
        return fallback_sheet, value
    sheet, address = value.rsplit("!", 1)
    sheet = sheet.strip()
    if len(sheet) > 1 and sheet[0] == sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, address


def _physical_shape(address: str, fallback_sheet: str | None = None) -> tuple[int | None, list[int] | None]:
    _sheet, local = _split_qualified(address, fallback_sheet)
    try:
        min_col, min_row, max_col, max_row = range_boundaries(local.replace("$", ""))
    except (TypeError, ValueError):
        return None, None
    if None in (min_col, min_row, max_col, max_row):
        return None, None
    rows, columns = max_row - min_row + 1, max_col - min_col + 1
    dimensions = 0 if rows == columns == 1 else 1 if rows == 1 or columns == 1 else 2
    return dimensions, [rows, columns]


def _scope_info(scope_path: Path | None, source_sha: str, sheet_names: set[str]) -> dict[str, Any]:
    if scope_path is None:
        return {"path": None, "sha256": None, "source_run_id": None, "ignored_sheets": []}
    scope_path = scope_path.expanduser().resolve()
    value = _read_json(scope_path)
    if not isinstance(value, dict) or value.get("schema_version") != "analysis.scope.v1":
        raise ValueError("scope must use schema_version analysis.scope.v1")
    if value.get("workbook_sha256") != source_sha:
        raise ValueError("scope workbook_sha256 does not match the selected Step 2 source")
    ignored = value.get("ignored_sheets", [])
    if not isinstance(ignored, list) or any(not isinstance(name, str) or not name for name in ignored):
        raise ValueError("scope ignored_sheets must be a list of non-empty sheet names")
    if len({name.casefold() for name in ignored}) != len(ignored):
        raise ValueError("scope ignored_sheets contains duplicate names")
    unknown = sorted(name for name in ignored if name.casefold() not in {item.casefold() for item in sheet_names})
    if unknown:
        raise ValueError(f"scope names sheets absent from this workbook: {', '.join(unknown)}")
    return {
        "path": str(scope_path),
        "sha256": _hash_file(scope_path),
        "source_run_id": value.get("source_run_id"),
        "ignored_sheets": sorted(ignored, key=str.casefold),
    }


def prepare_analysis(
    index_path: Path,
    step1_root: Path,
    out_dir: Path,
    *,
    source_id: str | None = None,
    scope_path: Path | None = None,
    workflow_path: Path | None = None,
) -> dict[str, Any]:
    """Validate Step 2, bind one source and optional source-bound scope."""
    try:
        index_path = index_path.expanduser().resolve()
        root = step1_root.expanduser().resolve()
        out_dir = out_dir.expanduser().resolve()
        workflow_root = None
        workflow_manifest = None
        if workflow_path is not None:
            from excel_to_act.steps import conversion_workflow

            workflow_root, workflow_manifest = conversion_workflow.load_workflow(workflow_path)
            conversion_workflow.require_stage_approved(workflow_root, 2)
        checked = _validate_saved_index(index_path, root)
        if checked.get("status") != "pass":
            messages = "; ".join(item.get("message", item.get("code", "invalid")) for item in checked.get("diagnostics", []))
            raise ValueError(f"Step 2 index validation failed: {messages or checked.get('status')}")
        index_value = _read_json(index_path)
        index = Step2Index.model_validate(index_value)
        if Path(index.step1_root).expanduser().resolve() != root:
            raise ValueError("--step1-root does not match the Step 1 root recorded in index.json")
        eligible = [entry for entry in index.entries if entry.ready_for_next_step and entry.status not in {"blocked", "error", "fail"}]
        if source_id is None:
            if len(eligible) != 1:
                if not eligible:
                    raise ValueError("Step 2 index has no eligible source")
                raise ValueError("Step 2 index has multiple eligible sources; pass --source-id")
            entry = eligible[0]
        else:
            matches = [item for item in eligible if item.source_id == source_id]
            if len(matches) != 1:
                raise ValueError(f"source_id {source_id!r} is not one eligible entry in index.json")
            entry = matches[0]
        inventory_refs = [ref for ref in entry.artifacts if ref.name == "inventory.json" or Path(ref.path).name == "inventory.json"]
        manifest_refs = [ref for ref in entry.artifacts if ref.name == "workbook_manifest.json" or Path(ref.path).name == "workbook_manifest.json"]
        if len(inventory_refs) != 1 or len(manifest_refs) != 1:
            raise ValueError("selected source must reference exactly one inventory.json and workbook_manifest.json")
        inventory_path = _resolved_child(root, inventory_refs[0].path, "inventory artifact path")
        manifest_path = _resolved_child(root, manifest_refs[0].path, "manifest artifact path")
        if _hash_file(inventory_path) != inventory_refs[0].sha256:
            raise ValueError("inventory.json checksum does not match the checked Step 2 index")
        manifest = _read_json(manifest_path)
        sheet_names = {item.get("name") for item in manifest.get("sheets", []) if isinstance(item, dict) and isinstance(item.get("name"), str)}
        scope = _scope_info(scope_path, entry.source_sha256 or "", sheet_names)
        binding = {
            "index_path": str(index_path),
            "index_sha256": _hash_file(index_path),
            "step1_root": str(root),
            "source": {
                "source_id": entry.source_id,
                "source_path": entry.source_path,
                "source_sha256": entry.source_sha256,
                "run_id": entry.run_id,
                "run_path": entry.run_path,
            },
            "artifacts": {
                "inventory_path": inventory_refs[0].path,
                "inventory_sha256": inventory_refs[0].sha256,
                "manifest_path": manifest_refs[0].path,
                "manifest_sha256": manifest_refs[0].sha256,
            },
            "scope": scope,
        }
        if workflow_root is not None and workflow_manifest is not None:
            workflow_source = workflow_manifest.get("source", {})
            if (workflow_source.get("source_id") != entry.source_id
                    or workflow_source.get("run_id") != entry.run_id
                    or workflow_source.get("workbook_sha256") != entry.source_sha256):
                raise ValueError("selected Step 2 source does not match the approved workflow source")
            binding["workflow_path"] = str(workflow_root)
            workflow_manifest["input_boundary_required"] = True
            workflow_manifest["input_boundary_policy"] = {
                "schema_version": "step3.input_boundary.policy.v1",
                "required_before_semantic_design": True,
                "catalog_confirmation_releases_stage": 3,
                "catalog_confirmation_releases_step4": False,
            }
            conversion_workflow._write_manifest(workflow_root, workflow_manifest)
        manifest_value = {
            "schema_version": "step3.analysis.v1",
            "binding": binding,
            "binding_sha256": _hash_bytes(_json_bytes(binding)),
            "stage_hashes": {},
        }
        analysis_file = out_dir / _ANALYSIS
        if analysis_file.exists():
            existing = _read_json(analysis_file)
            if not isinstance(existing, dict) or existing.get("binding_sha256") != manifest_value["binding_sha256"]:
                raise ValueError("analysis directory already belongs to another source or scope; choose a new --out directory")
        else:
            _write_json(analysis_file, manifest_value)
        return _tool_result("step3.prepare", out_dir, source_id=entry.source_id, run_id=entry.run_id,
                            scope=scope, index_validation=checked,
                            workflow=str(workflow_root) if workflow_root else None,
                            input_boundary_required=workflow_root is not None)
    except Exception as exc:
        return _blocked("step3.prepare", exc)


def _load_context(analysis_dir: Path) -> tuple[Path, dict[str, Any], dict[str, Any], Path, Path]:
    analysis_dir = analysis_dir.expanduser().resolve()
    manifest_path = analysis_dir / _ANALYSIS
    manifest = _read_json(manifest_path)
    if not isinstance(manifest, dict) or manifest.get("schema_version") != "step3.analysis.v1":
        raise ValueError("analysis.json is missing or has an unsupported schema")
    binding = manifest.get("binding")
    if not isinstance(binding, dict) or manifest.get("binding_sha256") != _hash_bytes(_json_bytes(binding)):
        raise ValueError("analysis.json binding checksum mismatch")
    root = Path(binding["step1_root"]).resolve()
    index_path = Path(binding["index_path"]).resolve()
    if _hash_file(index_path) != binding["index_sha256"]:
        raise ValueError("Step 2 index changed after step3 prepare; prepare again")
    inventory_path = _resolved_child(root, binding["artifacts"]["inventory_path"], "inventory artifact path")
    if _hash_file(inventory_path) != binding["artifacts"]["inventory_sha256"]:
        raise ValueError("inventory.json changed after step3 prepare; prepare again")
    manifest_artifact_path = _resolved_child(root, binding["artifacts"]["manifest_path"], "manifest artifact path")
    if _hash_file(manifest_artifact_path) != binding["artifacts"]["manifest_sha256"]:
        raise ValueError("workbook_manifest.json changed after step3 prepare; prepare again")
    scope = binding["scope"]
    if scope.get("path") and _hash_file(Path(scope["path"])) != scope.get("sha256"):
        raise ValueError("scope changed after step3 prepare; prepare again")
    return analysis_dir, manifest, binding, root, inventory_path


def _read_stage(analysis_dir: Path, manifest: dict[str, Any], stage: str) -> tuple[dict[str, Any], str]:
    name = _STAGE_FILES[stage]
    path = analysis_dir / name
    expected = manifest.get("stage_hashes", {}).get(stage)
    if not isinstance(expected, str) or not path.is_file():
        raise ValueError(f"step3 {stage} must be run before this command")
    digest = _hash_file(path)
    if digest != expected:
        raise ValueError(f"{name} checksum mismatch; rebuild this stage before continuing")
    value = _read_json(path)
    if not isinstance(value, dict) or value.get("binding_sha256") != manifest["binding_sha256"]:
        raise ValueError(f"{name} belongs to a different analysis binding")
    return value, digest


def _record_stage(analysis_dir: Path, manifest: dict[str, Any], stage: str, artifact: dict[str, Any]) -> str:
    digest = _write_json(analysis_dir / _STAGE_FILES[stage], artifact)
    manifest.setdefault("stage_hashes", {})[stage] = digest
    _write_json(analysis_dir / _ANALYSIS, manifest)
    return digest


def _load_inventory(path: Path) -> dict[str, Any]:
    data = _read_json(path)
    if not isinstance(data, dict) or data.get("artifact_type") != "workbook_inventory":
        raise ValueError("inventory artifact is not a workbook_inventory")
    if not isinstance(data.get("sheets"), list):
        raise ValueError("inventory sheets must be a list")
    return data


def _name_descriptors(inventory: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[tuple[str, str], list[str]]]:
    descriptors: list[dict[str, Any]] = []
    direct: dict[tuple[str, str], list[str]] = defaultdict(list)
    for definition in inventory.get("workbook_ranges", []):
        if not isinstance(definition, dict) or definition.get("kind") != "defined_name":
            continue
        loc = definition.get("source_location") or {}
        sheet, address = _split_qualified(str(definition.get("address", "")), loc.get("sheet_name"))
        address = address.replace("$", "")
        dimensions, _shape = _physical_shape(address, sheet)
        descriptor = {
            "name": definition.get("name"),
            "scope": (definition.get("metadata") or {}).get("scope", "workbook"),
            "sheet": sheet,
            "address": definition.get("address"),
            "physical_dimensions": dimensions,
            "logical_axes": None,
        }
        descriptors.append(descriptor)
        cell = _a1_cell(address)
        if cell and sheet and definition.get("name"):
            direct[(sheet.casefold(), address.upper())].append(str(definition["name"]))
    return descriptors, direct


def _formula_profile(formula: str, origin: str) -> tuple[str, list[str]]:
    try:
        family = Translator(formula, origin=origin).translate_formula("A1")
    except (ValueError, TypeError, IndexError, TranslatorError, TokenizerError):
        # ponytail: formulas that cannot shift to A1 keep their text; adjacent-copy checks still group safe runs.
        family = formula
    operations: set[str] = set()
    try:
        for token in Tokenizer(formula).items:
            if token.type == "FUNC" and token.subtype == "OPEN":
                operations.add(token.value[:-1].rsplit(".", 1)[-1].upper())
            elif token.type.startswith("OPERATOR"):
                operations.add(token.value)
    except (TokenizerError, IndexError):
        operations.add("unparsed_formula")
    return family, sorted(operations)


def _same_copied_formula(previous: dict[str, Any], current: dict[str, Any]) -> bool:
    try:
        return Translator(previous["formula"], origin=previous["address"]).translate_formula(current["address"]) == current["formula"]
    except (ValueError, TypeError, IndexError, TranslatorError, TokenizerError):
        return previous["formula"] == current["formula"]


def _literal_key(cell: dict[str, Any]) -> tuple[Any, ...]:
    value = cell.get("value")
    value_kind = "bool" if isinstance(value, bool) else "number" if isinstance(value, (int, float)) else "text" if isinstance(value, str) else "empty"
    return value_kind, cell.get("data_type"), cell.get("style_id"), cell.get("number_format"), cell.get("_step3_table_band")


def _make_field(sheet: str, members: list[dict[str, Any]], direct_names: dict[tuple[str, str], list[str]],
                array_formula: dict[str, Any] | None = None) -> dict[str, Any]:
    members = sorted(members, key=lambda item: (item["row"], item["column"], item["address"]))
    cells = [item["source_identity"] for item in members]
    field_id = "field-" + hashlib.sha256((sheet + "\0" + "\0".join(sorted(cells))).encode("utf-8")).hexdigest()[:16]
    ordinary_formula_members = [item for item in members
                                if item["kind"] == "formula"
                                and (item.get("_step3_array_formula_member") or {}).get("role") != "follower"]
    formulas = [item["formula"] for item in ordinary_formula_members]
    array_followers = [item for item in members
                       if (item.get("_step3_array_formula_member") or {}).get("role") == "follower"]
    role = "calculated" if formulas or array_followers else "source"
    orientation = "row" if len(members) > 1 and len({item["row"] for item in members}) == 1 else "column" if len(members) > 1 else None
    family, operations = _formula_profile(formulas[0], members[0]["address"]) if formulas else (None, [])
    names = sorted({name for item in members for name in direct_names.get((sheet.casefold(), item["address"].upper()), [])}, key=str.casefold)
    data_types = sorted({str(item.get("data_type") or "unknown") for item in members})
    member_records = [{"cell_id": item["source_identity"], "address": item["address"]} for item in members]
    if array_formula is not None:
        for member_record, item in zip(member_records, members, strict=True):
            array_member = item.get("_step3_array_formula_member") or {}
            member_record.update({
                "source_kind": item.get("kind"),
                "formula_present": isinstance(item.get("formula"), str) and bool(item["formula"]),
                "formula": item.get("formula"),
                "cached_value_present": item.get("cached_value_present"),
                "cached_value_available": item.get("cached_value_available"),
                "array_formula_member": array_member,
            })
    physical_dimensions = array_formula.get("physical_dimensions") if array_formula else None
    if array_formula:
        physical_shape = f"{physical_dimensions}D" if physical_dimensions in {1, 2} else "0D"
        axis_candidates = ([{"axis": "row", "status": "unconfirmed"},
                           {"axis": "column", "status": "unconfirmed"}]
                          if physical_dimensions == 2 else
                          [{"axis": orientation, "status": "unconfirmed"}] if orientation else [])
    else:
        physical_shape = "1D" if len(members) > 1 else "0D"
        axis_candidates = [{"axis": orientation, "status": "unconfirmed"}] if orientation else []
    array_follower_count = len(array_followers)
    calculated_member_count = len(ordinary_formula_members) + array_follower_count
    return {
        "field_id": field_id,
        "sheet": sheet,
        "role": role,
        "value_kind": "formula" if formulas else ("boolean" if isinstance(members[0].get("value"), bool) else "number" if isinstance(members[0].get("value"), (int, float)) else "text" if isinstance(members[0].get("value"), str) else "unknown"),
        "physical_shape": physical_shape,
        "member_count": len(members),
        "calculated_member_count": calculated_member_count,
        "formula_member_count": len(ordinary_formula_members),
        "array_formula_follower_count": array_follower_count,
        "members": member_records,
        "address_extent": array_formula["address"] if array_formula else (f"{members[0]['address']}:{members[-1]['address']}" if len(members) > 1 else members[0]["address"]),
        "axis_candidates": axis_candidates,
        "logical_axes": None,
        "names": names,
        "data_types": data_types,
        "formula_family": family,
        "formula_examples": formulas[:1] + (formulas[-1:] if len(formulas) > 1 else []),
        "operation_tags": operations,
        "cached_formula_cells": sum(1 for item in ordinary_formula_members if item.get("cached_value_available") is True),
        "array_formula": array_formula,
        "semantic_role": "calculated_value" if role == "calculated" else "unconfirmed_source_value",
    }


def _array_formula_specs(inventory: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    """Read array output bounds only from the canonical formula anchor metadata."""
    result: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for sheet in inventory["sheets"]:
        sheet_name = str(sheet["name"])
        for cell in sheet.get("cells", []):
            attributes = cell.get("raw_formula_attributes") or {}
            if not isinstance(attributes, dict) or str(attributes.get("t", "")).casefold() != "array":
                continue
            reference = attributes.get("ref")
            formula = cell.get("formula")
            if not isinstance(reference, str) or not reference or not isinstance(formula, str) or not formula.startswith("="):
                raise ValueError(f"array formula anchor {sheet_name}!{cell.get('address')} lacks source ref or formula text")
            try:
                min_col, min_row, max_col, max_row = range_boundaries(reference.replace("$", ""))
            except (TypeError, ValueError) as exc:
                raise ValueError(f"array formula anchor {sheet_name}!{cell.get('address')} has an invalid ref {reference!r}") from exc
            if None in (min_col, min_row, max_col, max_row):
                raise ValueError(f"array formula anchor {sheet_name}!{cell.get('address')} has unbounded ref {reference!r}")
            if (cell.get("row"), cell.get("column")) != (min_row, min_col):
                raise ValueError(f"array formula anchor {sheet_name}!{cell.get('address')} is not the top-left of {reference}")
            normalized_ref = reference.replace("$", "")
            dimensions, shape = _physical_shape(normalized_ref, sheet_name)
            spec = {"sheet": sheet_name, "anchor": cell["address"], "address": normalized_ref,
                    "formula": formula, "min_row": min_row, "min_column": min_col,
                    "max_row": max_row, "max_column": max_col,
                    "physical_dimensions": dimensions, "physical_shape": shape,
                    "source_identity": cell.get("source_identity") or (cell.get("source_location") or {}).get("source_identity")}
            for existing in result[sheet_name.casefold()]:
                intersects = not (max_col < existing["min_column"] or min_col > existing["max_column"]
                                  or max_row < existing["min_row"] or min_row > existing["max_row"])
                if intersects:
                    raise ValueError(f"overlapping array formula ranges on {sheet_name}: {normalized_ref} and {existing['address']}")
            result[sheet_name.casefold()].append(spec)
    return result


def _build_fields(inventory: dict[str, Any], ignored_sheets: list[str], binding_sha256: str) -> dict[str, Any]:
    descriptors, direct_names = _name_descriptors(inventory)
    array_specs = _array_formula_specs(inventory)
    array_descriptors: list[dict[str, Any]] = []
    tables_by_sheet: dict[str, list[dict[str, Any]]] = defaultdict(list)
    tables: list[dict[str, Any]] = []
    for table_sheet in inventory["sheets"]:
        for item in table_sheet.get("ranges", []):
            if item.get("kind") == "table":
                tables_by_sheet[str(table_sheet["name"]).casefold()].append(item)
                tables.append(_range_descriptor(item, "table"))

    def table_band(sheet_name: str, cell: dict[str, Any]) -> str | None:
        for table in tables_by_sheet.get(sheet_name.casefold(), []):
            metadata = table.get("metadata") or {}
            header_rows, total_rows = metadata.get("header_row_count"), metadata.get("totals_row_count")
            try:
                min_col, min_row, max_col, max_row = range_boundaries(str(table.get("address", "")).replace("$", ""))
            except ValueError:
                continue
            if not (min_col <= cell["column"] <= max_col and min_row <= cell["row"] <= max_row):
                continue
            known = all(isinstance(count, int) and not isinstance(count, bool) and count >= 0 for count in (header_rows, total_rows))
            if not known:
                # ponytail: old inventory artifacts omit these counts, so table cells stay ungrouped.
                return f"unknown:{cell['address']}"
            if cell["row"] < min_row + header_rows:
                # Native table labels describe distinct columns; do not merge adjacent headers into one field.
                return f"header:{cell['address']}"
            if total_rows and cell["row"] > max_row - total_rows:
                return "totals"
            return "data"
        return None

    ignored = {name.casefold() for name in ignored_sheets}
    fields: list[dict[str, Any]] = []
    retained_cells = 0
    excluded_cells = 0
    for sheet in inventory["sheets"]:
        name = str(sheet["name"])
        sheet_array_specs = array_specs.get(name.casefold(), [])
        cells: list[dict[str, Any]] = []
        for cell in sheet.get("cells", []):
            if cell.get("kind") not in {"value", "formula"}:
                continue
            if name.casefold() in ignored:
                excluded_cells += 1
                continue
            location = cell.get("source_location") or {}
            cell_id = cell.get("source_identity") or location.get("source_identity") or f"{name}!{cell['address']}"
            copied = {**cell, "source_identity": str(cell_id), "_step3_table_band": table_band(name, cell)}
            for array_spec in sheet_array_specs:
                if (array_spec["min_row"] <= cell["row"] <= array_spec["max_row"]
                        and array_spec["min_column"] <= cell["column"] <= array_spec["max_column"]):
                    copied["_step3_array_formula_member"] = {
                        "anchor": array_spec["anchor"], "address": array_spec["address"],
                        "role": "anchor" if cell["address"].upper() == array_spec["anchor"].upper() else "follower",
                        "row_offset": cell["row"] - array_spec["min_row"],
                        "column_offset": cell["column"] - array_spec["min_column"],
                    }
                    break
            cells.append(copied)
        if name.casefold() in ignored:
            continue
        retained_cells += len(cells)
        by_address = {item["address"].upper(): item for item in cells}
        assigned: set[str] = set()
        named = {address for sheet_name, address in direct_names if sheet_name == name.casefold() and address in by_address}
        groups: list[list[dict[str, Any]]] = []
        array_group_specs: dict[tuple[str, ...], dict[str, Any]] = {}

        # Array output cells remain calculated members even when the source package stores
        # only a literal result at a follower coordinate. The anchor attributes, not values,
        # establish the formula range; absent cells are never invented here.
        for array_spec in sheet_array_specs:
            members = [cell for cell in cells
                       if array_spec["min_row"] <= cell["row"] <= array_spec["max_row"]
                       and array_spec["min_column"] <= cell["column"] <= array_spec["max_column"]]
            if not members:
                continue
            groups.append(members)
            member_addresses = tuple(sorted(item["address"].upper() for item in members))
            array_group_specs[member_addresses] = {
                "anchor": array_spec["anchor"], "address": array_spec["address"],
                "formula": array_spec["formula"], "physical_dimensions": array_spec["physical_dimensions"],
                "physical_shape": array_spec["physical_shape"],
                "shape": array_spec["physical_shape"], "member_count": len(members),
                "follower_count": sum(item["_step3_array_formula_member"]["role"] == "follower" for item in members),
            }
            assigned.update(member_addresses)
            array_descriptors.append({"sheet": name, "anchor": array_spec["anchor"],
                                      "address": array_spec["address"], "formula": array_spec["formula"],
                                      "physical_dimensions": array_spec["physical_dimensions"],
                                      "physical_shape": array_spec["physical_shape"],
                                      "stored_member_count": len(members),
                                      "stored_follower_count": sum(item["_step3_array_formula_member"]["role"] == "follower" for item in members)})

        # Named single cells stay scalar so adjacent inputs retain separate identities.
        for address in sorted(named - assigned):
            cell = by_address[address]
            groups.append([cell])
            assigned.add(address)

        formula_cells = [item for item in cells if item["kind"] == "formula" and item["address"].upper() not in assigned]
        for direction in ("column", "row"):
            buckets: dict[int, list[dict[str, Any]]] = defaultdict(list)
            for cell in formula_cells:
                if cell["address"].upper() in assigned:
                    continue
                buckets[cell["column"] if direction == "column" else cell["row"]].append(cell)
            for bucket in sorted(buckets):
                ordered = sorted(buckets[bucket], key=lambda item: item["row"] if direction == "column" else item["column"])
                run: list[dict[str, Any]] = []
                for cell in ordered:
                    coordinate = cell["row"] if direction == "column" else cell["column"]
                    previous_coordinate = (run[-1]["row"] if direction == "column" else run[-1]["column"]) if run else None
                    contiguous = (bool(run) and coordinate == previous_coordinate + 1
                                  and run[-1].get("_step3_table_band") == cell.get("_step3_table_band")
                                  and _same_copied_formula(run[-1], cell))
                    if not contiguous and run:
                        if len(run) > 1:
                            groups.append(run)
                            assigned.update(item["address"].upper() for item in run)
                        run = []
                    run.append(cell)
                if len(run) > 1:
                    groups.append(run)
                    assigned.update(item["address"].upper() for item in run)

        literal_cells = [item for item in cells if item["kind"] == "value" and item["address"].upper() not in assigned]
        for direction in ("column", "row"):
            buckets = defaultdict(list)
            for cell in literal_cells:
                if cell["address"].upper() in assigned:
                    continue
                buckets[cell["column"] if direction == "column" else cell["row"]].append(cell)
            for bucket in sorted(buckets):
                ordered = sorted(buckets[bucket], key=lambda item: item["row"] if direction == "column" else item["column"])
                run = []
                for cell in ordered:
                    coordinate = cell["row"] if direction == "column" else cell["column"]
                    prior = (run[-1]["row"] if direction == "column" else run[-1]["column"]) if run else None
                    contiguous = bool(run) and coordinate == prior + 1 and _literal_key(run[-1]) == _literal_key(cell)
                    if not contiguous and run:
                        if len(run) > 1:
                            groups.append(run)
                            assigned.update(item["address"].upper() for item in run)
                        run = []
                    run.append(cell)
                if len(run) > 1:
                    groups.append(run)
                    assigned.update(item["address"].upper() for item in run)
        groups.extend([[item] for item in cells if item["address"].upper() not in assigned])

        for members in groups:
            member_key = tuple(sorted(item["address"].upper() for item in members))
            field = _make_field(name, members, direct_names, array_group_specs.get(member_key))
            fields.append(field)

    fields.sort(key=lambda item: (item["sheet"].casefold(), item["members"][0]["address"], item["field_id"]))
    descriptors.sort(key=lambda item: (str(item.get("scope", "")).casefold(), str(item.get("name", "")).casefold(), str(item.get("address", ""))))
    tables.sort(key=lambda item: (str(item.get("sheet", "")).casefold(), str(item.get("name", "")).casefold(), str(item.get("address", ""))))
    summary = {
        "retained_stored_cells": retained_cells,
        "excluded_stored_cells": excluded_cells,
        "primary_fields": len(fields),
        "source_fields": sum(field["role"] == "source" for field in fields),
        "calculated_fields": sum(field["role"] == "calculated" for field in fields),
        "formula_cells": sum(field["member_count"] for field in fields if field["role"] == "calculated"),
        "ordinary_formula_cells": sum(field["formula_member_count"] for field in fields),
        "array_formula_followers": sum(field["array_formula_follower_count"] for field in fields),
        "array_formula_instances": len(array_descriptors),
        "cached_formula_cells": sum(field["cached_formula_cells"] for field in fields),
        "defined_name_descriptors": len(descriptors),
        "table_descriptors": len(tables),
        "fields_by_sheet": dict(sorted(Counter(field["sheet"] for field in fields).items(), key=lambda item: item[0].casefold())),
    }
    return {
        "schema_version": "step3.fields.v1",
        "binding_sha256": binding_sha256,
        "coverage": {"retained_stored_cells": retained_cells},
        "fields": fields,
        "descriptors": {"defined_names": descriptors, "tables": tables,
                        "array_formulas": sorted(array_descriptors, key=lambda item: (item["sheet"].casefold(), item["anchor"]))},
        "summary": summary,
    }


def _range_descriptor(item: dict[str, Any], kind: str) -> dict[str, Any]:
    address = str(item.get("address", ""))
    loc = item.get("source_location") or {}
    dimensions, shape = _physical_shape(address, loc.get("sheet_name"))
    descriptor = {
        "kind": kind,
        "name": item.get("name"),
        "sheet": loc.get("sheet_name"),
        "address": address,
        "physical_dimensions": dimensions,
        "physical_shape": shape,
        "logical_axes": None,
    }
    if kind == "table":
        metadata = item.get("metadata") or {}
        descriptor["row_boundaries_known"] = all(
            isinstance(metadata.get(key), int) and not isinstance(metadata.get(key), bool) and metadata[key] >= 0
            for key in ("header_row_count", "totals_row_count")
        )
        for key in ("header_row_count", "totals_row_count"):
            if key in metadata:
                descriptor[key] = metadata[key]
    return descriptor


def build_fields(analysis_dir: Path) -> dict[str, Any]:
    try:
        analysis_dir, manifest, binding, _root, inventory_path = _load_context(analysis_dir)
        inventory = _load_inventory(inventory_path)
        if inventory.get("workbook_sha256") != binding["source"]["source_sha256"]:
            raise ValueError("inventory workbook_sha256 does not match the selected source")
        artifact = _build_fields(inventory, binding["scope"]["ignored_sheets"], manifest["binding_sha256"])
        digest = _record_stage(analysis_dir, manifest, "fields", artifact)
        report = ["# Step 3 Field Draft", "", f"Source: `{binding['source']['source_path']}` · `{binding['source']['source_id']}`", "",
                  f"Stored cells: {artifact['summary']['retained_stored_cells']} · fields: {artifact['summary']['primary_fields']} · source: {artifact['summary']['source_fields']} · calculated: {artifact['summary']['calculated_fields']}",
                  "", "All logical axes are unconfirmed. The examples below are cell membership and formula-family evidence only.", ""]
        shown = 0
        for field in artifact["fields"]:
            if shown >= 30:
                break
            addresses = ", ".join(member["address"] for member in field["members"][:3])
            if field["member_count"] > 3:
                addresses += f", … ({field['member_count']} cells)"
            report.append(f"- `{field['sheet']}!{addresses}` · {field['role']} · {field['physical_shape']} · names: {', '.join(field['names']) or 'none'}")
            shown += 1
        _write_bytes(analysis_dir / "fields.md", ("\n".join(report) + "\n").encode("utf-8"))
        return _tool_result("step3.fields", analysis_dir, output=_FIELDS, sha256=digest, summary=artifact["summary"])
    except Exception as exc:
        return _blocked("step3.fields", exc)


def _field_index(fields_artifact: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[tuple[str, str], str]]:
    fields = {item["field_id"]: item for item in fields_artifact.get("fields", [])}
    cells: dict[tuple[str, str], str] = {}
    for field_id, item in fields.items():
        for member in item.get("members", []):
            cells[(item["sheet"].casefold(), member["address"].upper())] = field_id
    return fields, cells


def _category(description: str) -> str:
    text = description.casefold()
    if "dynamic reference" in text or "indirect" in text or "offset" in text:
        return "dynamic_reference"
    if "#ref!" in text or "broken" in text:
        return "broken_reference"
    if "tokenize" in text or "parenthes" in text or "empty and has no operands" in text:
        return "unsupported_formula"
    return "unresolved_reference"


def _resolve_node(
    node: Any,
    source_sheet: str,
    source_row: int,
    inventory: dict[str, Any],
    cell_maps: dict[str, dict[int, dict[int, dict[str, Any]]]],
    row_orders: dict[str, list[int]],
    cell_fields: dict[tuple[str, str], str],
    ignored: set[str],
) -> dict[str, Any]:
    if node.kind == "external":
        return {"status": "unresolved", "boundary_type": "external_reference", "target_sheet": None, "target_fields": [], "evidence": node.label}
    target_name = str(node.metadata.get("name", "")) if node.kind == "name" else ""
    if node.kind == "name":
        scope = str(node.metadata.get("scope", "workbook"))
        definition = next((item for item in inventory.get("workbook_ranges", [])
                           if item.get("kind") == "defined_name"
                           and str(item.get("name", "")).casefold() == target_name.casefold()
                           and str((item.get("metadata") or {}).get("scope", "workbook")).casefold() == scope.casefold()), None)
        if definition is None:
            return {"status": "unresolved", "boundary_type": "defined_name_unresolved", "target_sheet": None, "target_fields": [], "evidence": node.label}
        loc = definition.get("source_location") or {}
        sheet, address = _split_qualified(str(definition.get("address", "")), loc.get("sheet_name"))
        if not sheet:
            return {"status": "unresolved", "boundary_type": "defined_name_not_a_range", "target_sheet": None, "target_fields": [], "evidence": node.label}
        return _resolve_address(sheet, address, node.label, source_row, cell_maps, row_orders, cell_fields, ignored)
    if node.metadata.get("structured_reference"):
        table = next((item for sheet in inventory.get("sheets", []) for item in sheet.get("ranges", [])
                      if item.get("kind") == "table" and str(item.get("name", "")).casefold() == str(node.metadata.get("table", "")).casefold()), None)
        if table is None:
            return {"status": "unresolved", "boundary_type": "structured_reference_unresolved", "target_sheet": None, "target_fields": [], "evidence": node.label}
        table_sheet = (table.get("source_location") or {}).get("sheet_name")
        address = table.get("address")
        raw_reference = str(node.metadata.get("structured_reference", ""))
        table_name = str(node.metadata.get("table", ""))
        column = node.metadata.get("column")
        headers = node.metadata.get("table_columns") or []
        current_row = node.metadata.get("current_row") is True
        raw_lower = raw_reference.casefold()
        prefix = f"{table_name}[".casefold()
        raw_selector = raw_reference[len(prefix):-1] if raw_lower.startswith(prefix) and raw_reference.endswith("]") else ""
        selector_key = raw_selector.casefold()
        column_verified = bool(
            isinstance(column, str)
            and column
            and isinstance(headers, list)
            and any(str(header).casefold() == column.casefold() for header in headers)
        )
        is_all = selector_key == "#all" and not current_row and column is None
        is_data = selector_key == "#data" and not current_row and column is None
        is_current_column = (
            current_row and column_verified and selector_key == f"@{column}".casefold()
        )
        is_plain_column = (
            not current_row and column_verified and selector_key == str(column).casefold()
        )
        if not (is_all or is_data or is_current_column or is_plain_column):
            return {"status": "unresolved", "boundary_type": "structured_reference_unsupported", "target_sheet": table_sheet,
                    "target_fields": [], "evidence": raw_reference or node.label, "selector": raw_selector or raw_reference}
        try:
            min_col, min_row, max_col, max_row = range_boundaries(address)
            if column and column_verified:
                min_col += next(index for index, header in enumerate(headers) if str(header).casefold() == str(column).casefold())
                max_col = min_col
            if is_all:
                # #All explicitly includes every physical table row, including headers and totals.
                pass
            else:
                metadata = table.get("metadata") or {}
                missing = [key for key in ("header_row_count", "totals_row_count")
                           if not isinstance(metadata.get(key), int) or isinstance(metadata.get(key), bool) or metadata[key] < 0]
                if missing:
                    return {"status": "unresolved", "boundary_type": "structured_reference_boundary_unknown", "target_sheet": table_sheet,
                            "target_fields": [], "evidence": node.label, "missing_metadata": missing}
                data_min_row = min_row + metadata["header_row_count"]
                data_max_row = max_row - metadata["totals_row_count"]
                if is_current_column:  # Geometric membership is known; logical row meaning remains unconfirmed.
                    if not data_min_row <= source_row <= data_max_row:
                        return {"status": "unresolved", "boundary_type": "structured_reference_current_row_outside_data", "target_sheet": table_sheet,
                                "target_fields": [], "evidence": node.label}
                    min_row = max(data_min_row, source_row)
                    max_row = min(data_max_row, source_row)
                else:
                    min_row, max_row = data_min_row, data_max_row
                if max_row < min_row:
                    return {"status": "empty", "boundary_type": "empty_table_data", "target_sheet": table_sheet,
                            "target_fields": [], "evidence": node.label}
            address = f"{_column_letter(min_col)}{min_row}:{_column_letter(max_col)}{max_row}"
            return _resolve_address(table_sheet, address, node.label, source_row, cell_maps, row_orders, cell_fields, ignored)
        except (ValueError, StopIteration, TypeError):
            return {"status": "unresolved", "boundary_type": "structured_reference_unresolved", "target_sheet": table_sheet, "target_fields": [], "evidence": node.label}
    if not node.source_location or not node.source_location.sheet_name:
        return {"status": "unresolved", "boundary_type": "range_location_missing", "target_sheet": None, "target_fields": [], "evidence": node.label}
    return _resolve_address(node.source_location.sheet_name, node.source_location.address or "", node.label, source_row, cell_maps, row_orders, cell_fields, ignored)


def _column_letter(column: int) -> str:
    from openpyxl.utils.cell import get_column_letter

    return get_column_letter(column)


def _resolve_address(
    sheet: str | None,
    address: str,
    evidence: str,
    source_row: int,
    cell_maps: dict[str, dict[int, dict[int, dict[str, Any]]]],
    row_orders: dict[str, list[int]],
    cell_fields: dict[tuple[str, str], str],
    ignored: set[str],
) -> dict[str, Any]:
    if not sheet:
        return {"status": "unresolved", "boundary_type": "range_sheet_missing", "target_sheet": None, "target_fields": [], "evidence": evidence}
    try:
        min_col, min_row, max_col, max_row = range_boundaries(address.replace("$", ""))
    except ValueError:
        return {"status": "unresolved", "boundary_type": "unsupported_range", "target_sheet": sheet, "target_fields": [], "evidence": evidence}
    if any(bound is None for bound in (min_col, min_row, max_col, max_row)):
        return {"status": "unresolved", "boundary_type": "unsupported_unbounded_range", "target_sheet": sheet,
                "target_fields": [], "evidence": evidence, "address": address}
    sheet_key = sheet.casefold()
    indexed_rows = cell_maps.get(sheet_key, {})
    found: list[tuple[int, int, dict[str, Any]]] = []
    if min_row == max_row and min_col == max_col:
        cell = indexed_rows.get(min_row, {}).get(min_col)
        if cell is not None:
            found.append((min_row, min_col, cell))
    else:
        rows = row_orders.get(sheet_key, [])
        for row in rows[bisect_left(rows, min_row):bisect_right(rows, max_row)]:
            for column, cell in indexed_rows[row].items():
                if min_col <= column <= max_col:
                    found.append((row, column, cell))
    field_ids = sorted({cell_fields[(sheet_key, cell["address"].upper())] for _, _, cell in found if (sheet_key, cell["address"].upper()) in cell_fields})
    area = (max_col - min_col + 1) * (max_row - min_row + 1)
    if sheet_key in ignored:
        cached = [{"address": cell["address"], "kind": cell.get("kind"), "value": cell.get("value"), "cached_value": cell.get("cached_value"), "cached_value_available": cell.get("cached_value_available", False)} for _, _, cell in found[:8]]
        return {"status": "boundary", "boundary_type": "excluded_sheet", "target_sheet": sheet, "target_fields": [], "evidence": evidence,
                "range_area": area, "stored_cell_count": len(found), "cached_formula_count": sum(cell.get("kind") == "formula" and cell.get("cached_value_available") is True for _, _, cell in found), "stored_samples": cached}
    if area == 1 and not found:
        return {"status": "empty", "boundary_type": "empty_cell", "target_sheet": sheet, "target_fields": [], "evidence": evidence,
                "empty_cells": 1, "empty_examples": [address.replace("$", "").upper()]}
    if not found:
        return {"status": "empty", "boundary_type": "empty_range", "target_sheet": sheet, "target_fields": [], "evidence": evidence,
                "range_area": area, "empty_cells": area, "empty_examples": [f"{sheet}!{address}"]}
    empty_count = max(0, area - len(found))
    empty_examples = []
    if empty_count:
        for row, col in ((min_row, min_col), (min_row, max_col), (max_row, min_col), (max_row, max_col)):
            if col not in indexed_rows.get(row, {}):
                empty_examples.append(f"{_column_letter(col)}{row}")
            if len(empty_examples) == 3:
                break
    return {"status": "resolved", "boundary_type": "range" if area > 1 else "cell", "target_sheet": sheet,
            "target_fields": field_ids, "evidence": evidence, "range_area": area, "stored_cell_count": len(found),
            "empty_cells": empty_count, "empty_examples": empty_examples[:3]}


def _build_dependencies(inventory: dict[str, Any], fields_artifact: dict[str, Any], ignored_sheets: list[str], binding_sha256: str, fields_sha256: str) -> dict[str, Any]:
    ignored = {item.casefold() for item in ignored_sheets}
    fields, cell_fields = _field_index(fields_artifact)
    cell_maps: dict[str, dict[int, dict[int, dict[str, Any]]]] = {}
    row_orders: dict[str, list[int]] = {}
    for sheet in inventory["sheets"]:
        name = str(sheet["name"])
        rows: dict[int, dict[int, dict[str, Any]]] = defaultdict(dict)
        for cell in sheet.get("cells", []):
            rows[int(cell["row"])][int(cell["column"])] = cell
        cell_maps[name.casefold()] = rows
        row_orders[name.casefold()] = sorted(rows)
    # Keep excluded sheet names and range descriptors for boundary resolution, but skip their formulas.
    graph_payload = dict(inventory)
    graph_payload["sheets"] = [
        ({**sheet, "cells": sheet.get("cells", [])} if sheet["name"].casefold() not in ignored else {**sheet, "cells": []})
        for sheet in inventory["sheets"]
    ]
    graph = RegexFormulaGraphBuilder().build(WorkbookInventory.model_validate(graph_payload))
    nodes = {node.id: node for node in graph.nodes}
    grouped: dict[tuple[Any, ...], dict[str, Any]] = {}
    unresolved: dict[tuple[str, str], dict[str, Any]] = {}
    internal: dict[tuple[Any, ...], dict[str, Any]] = {}
    formula_to_field = {(field["sheet"].casefold(), member["address"].upper()): field_id
                        for field_id, field in fields.items() for member in field["members"] if field["role"] == "calculated"}
    for edge in graph.edges:
        source_location = edge.source_location
        if not source_location or not source_location.sheet_name or not source_location.address:
            continue
        source_sheet = source_location.sheet_name
        source_address = source_location.address.upper()
        source_cell = _a1_cell(source_address)
        consumer = formula_to_field.get((source_sheet.casefold(), source_address))
        node = nodes.get(edge.target)
        if consumer is None or node is None or source_cell is None:
            continue
        resolved = _resolve_node(node, source_sheet, source_cell[0], inventory, cell_maps, row_orders, cell_fields, ignored)
        target_fields = resolved.pop("target_fields")
        offsets = []
        if node.kind == "cell" and node.source_location and node.source_location.sheet_name and node.source_location.address:
            target_cell = _a1_cell(node.source_location.address)
            if target_cell and target_cell[0] <= 1_048_576 and target_cell[1] <= 16_384:
                offsets.append({"row": target_cell[0] - source_cell[0], "column": target_cell[1] - source_cell[1]})
        if resolved["status"] == "resolved" and not target_fields:
            resolved["status"] = "empty"
            resolved["boundary_type"] = "empty_range"
        elif resolved["status"] == "resolved" and resolved.get("empty_cells", 0) > 0:
            # Preserve sparse-range evidence alongside its ordinary prerequisite edges.
            key = (consumer, "boundary", "partial_empty_range", resolved.get("target_sheet"), resolved.get("evidence"), tuple((item["row"], item["column"]) for item in offsets))
            item = grouped.setdefault(key, {"consumer": consumer, "prerequisite": None, "status": "boundary", "boundary_type": "partial_empty_range",
                                            "target_sheet": resolved.get("target_sheet"), "evidence": resolved.get("evidence"), "relative_offsets": offsets,
                                            "occurrences": 0, "source_formula_cells": [],
                                            "target_details": {name: resolved[name] for name in ("range_area", "stored_cell_count", "empty_cells", "empty_examples") if name in resolved}})
            item["occurrences"] += 1
            if len(item["source_formula_cells"]) < 3 and source_address not in item["source_formula_cells"]:
                item["source_formula_cells"].append(source_address)
        if not target_fields:
            key = (consumer, resolved["status"], resolved["boundary_type"], resolved.get("target_sheet"), resolved.get("evidence"), tuple((item["row"], item["column"]) for item in offsets))
            item = grouped.setdefault(key, {"consumer": consumer, "prerequisite": None, "status": resolved["status"], "boundary_type": resolved["boundary_type"],
                                            "target_sheet": resolved.get("target_sheet"), "evidence": resolved.get("evidence"), "relative_offsets": offsets,
                                            "occurrences": 0, "source_formula_cells": [], "target_details": {k: v for k, v in resolved.items() if k not in {"status", "boundary_type", "target_sheet", "evidence"}}})
            item["occurrences"] += 1
            if len(item["source_formula_cells"]) < 3 and source_address not in item["source_formula_cells"]:
                item["source_formula_cells"].append(source_address)
            if resolved["status"] == "unresolved":
                if resolved["boundary_type"] == "structured_reference_boundary_unknown":
                    category = "table_boundary_unknown"
                elif resolved["boundary_type"] == "external_reference":
                    category = "external_reference"
                else:
                    category = "unresolved_reference"
                unresolved_item = unresolved.setdefault((consumer, category), {"consumer": consumer, "category": category, "occurrences": 0, "evidence_samples": []})
                unresolved_item["occurrences"] += 1
                if len(unresolved_item["evidence_samples"]) < 3 and resolved.get("evidence") not in unresolved_item["evidence_samples"]:
                    unresolved_item["evidence_samples"].append(resolved.get("evidence"))
            continue
        for prerequisite in target_fields:
            same_field = consumer == prerequisite
            consumer_members = fields[consumer].get("members", [])
            single_member_source = len(consumer_members) == 1 and consumer_members[0].get("address", "").upper() == source_address
            direct_same_cell = (resolved.get("target_sheet", "").casefold() == source_sheet.casefold()
                                and resolved.get("evidence", "").rsplit("!", 1)[-1].replace("$", "").upper() == source_address)
            same_cell = same_field and (single_member_source or (len(target_fields) == 1 and direct_same_cell))
            key = (consumer, prerequisite, tuple((item["row"], item["column"]) for item in offsets), same_cell)
            item = internal.setdefault(key, {"consumer": consumer, "prerequisite": prerequisite, "relationship": "prerequisite_before_consumer",
                                             "relative_offsets": offsets, "same_field": same_field, "same_cell": same_cell,
                                             "evidence": resolved.get("evidence"), "occurrences": 0, "source_formula_cells": [],
                                             "target_status": resolved["status"], "target_sheet": resolved.get("target_sheet")})
            item["occurrences"] += 1
            if len(item["source_formula_cells"]) < 3 and source_address not in item["source_formula_cells"]:
                item["source_formula_cells"].append(source_address)
            if same_field:
                item["recurrence_candidate"] = not same_cell

    for unsupported in graph.unsupported_features:
        loc = unsupported.source_location
        if not loc or not loc.sheet_name or not loc.address:
            continue
        field_id = formula_to_field.get((loc.sheet_name.casefold(), loc.address.upper()))
        if not field_id:
            continue
        category = _category(unsupported.description)
        item = unresolved.setdefault((field_id, category), {"consumer": field_id, "category": category, "occurrences": 0, "evidence_samples": []})
        item["occurrences"] += 1
        if len(item["evidence_samples"]) < 3 and unsupported.description not in item["evidence_samples"]:
            item["evidence_samples"].append(unsupported.description)

    edges = sorted((item for item in internal.values() if not item["same_field"]), key=lambda item: (item["consumer"], item["prerequisite"], str(item["evidence"])))
    internal_dependencies = sorted((item for item in internal.values() if item["same_field"]), key=lambda item: (item["consumer"], str(item["evidence"])))
    boundaries = sorted(grouped.values(), key=lambda item: (item["consumer"], item["status"], item["boundary_type"], str(item["evidence"])))
    unresolved_items = sorted(unresolved.values(), key=lambda item: (item["consumer"], item["category"]))
    summary = {"field_edges": len(edges), "reference_boundaries": len(boundaries), "unresolved_groups": len(unresolved_items),
               "internal_dependencies": len(internal_dependencies), "dynamic_reference_occurrences": sum(item["occurrences"] for item in unresolved_items if item["category"] == "dynamic_reference"),
               "excluded_sheet_boundaries": sum(item["occurrences"] for item in boundaries if item["boundary_type"] == "excluded_sheet"),
               "empty_reference_occurrences": sum(item["occurrences"] for item in boundaries
                                                   if item["status"] == "empty" or item["boundary_type"] == "partial_empty_range"),
               "unresolved_categories": dict(sorted(Counter(item["category"] for item in unresolved_items).items()))}
    return {"schema_version": "step3.dependencies.v1", "binding_sha256": binding_sha256, "fields_sha256": fields_sha256,
            "direction": "prerequisite_before_consumer", "edges": edges, "internal_dependencies": internal_dependencies,
            "boundaries": boundaries, "unresolved": unresolved_items, "summary": summary}


def build_dependencies(analysis_dir: Path) -> dict[str, Any]:
    try:
        analysis_dir, manifest, binding, _root, inventory_path = _load_context(analysis_dir)
        fields_artifact, fields_sha256 = _read_stage(analysis_dir, manifest, "fields")
        inventory = _load_inventory(inventory_path)
        artifact = _build_dependencies(inventory, fields_artifact, binding["scope"]["ignored_sheets"], manifest["binding_sha256"], fields_sha256)
        digest = _record_stage(analysis_dir, manifest, "dependencies", artifact)
        return _tool_result("step3.dependencies", analysis_dir, output=_DEPENDENCIES, sha256=digest, summary=artifact["summary"])
    except Exception as exc:
        return _blocked("step3.dependencies", exc)


def _scc(nodes: list[str], adjacency: dict[str, set[str]]) -> list[list[str]]:
    visited: set[str] = set()
    finish: list[str] = []
    for start in nodes:
        if start in visited:
            continue
        visited.add(start)
        stack: list[tuple[str, Any]] = [(start, iter(sorted(adjacency.get(start, ()))))]
        while stack:
            node, targets = stack[-1]
            try:
                target = next(targets)
            except StopIteration:
                stack.pop()
                finish.append(node)
                continue
            if target not in visited:
                visited.add(target)
                stack.append((target, iter(sorted(adjacency.get(target, ())))))
    reverse: dict[str, set[str]] = {node: set() for node in nodes}
    for source, targets in adjacency.items():
        for target in targets:
            reverse.setdefault(target, set()).add(source)
    visited.clear()
    components: list[list[str]] = []
    for start in reversed(finish):
        if start in visited:
            continue
        component: list[str] = []
        stack = [start]
        visited.add(start)
        while stack:
            node = stack.pop()
            component.append(node)
            for target in sorted(reverse.get(node, ()), reverse=True):
                if target not in visited:
                    visited.add(target)
                    stack.append(target)
        components.append(sorted(component))
    return components


def _build_plan(fields_artifact: dict[str, Any], dependencies: dict[str, Any], binding_sha256: str, fields_sha256: str, dependencies_sha256: str) -> dict[str, Any]:
    fields = {item["field_id"]: item for item in fields_artifact.get("fields", [])}
    ids = sorted(fields)
    adjacency: dict[str, set[str]] = {field_id: set() for field_id in ids}
    for edge in dependencies.get("edges", []):
        consumer, prerequisite = edge.get("consumer"), edge.get("prerequisite")
        if consumer in fields and prerequisite in fields:
            adjacency[consumer].add(prerequisite)
    for item in dependencies.get("internal_dependencies", []):
        if item.get("same_cell") and item.get("consumer") in fields:
            adjacency[item["consumer"]].add(item["prerequisite"])
    components = _scc(ids, adjacency)
    component_of = {field_id: position for position, component in enumerate(components) for field_id in component}
    prerequisites: dict[int, set[int]] = {position: set() for position in range(len(components))}
    for consumer, targets in adjacency.items():
        for prerequisite in targets:
            left, right = component_of[consumer], component_of[prerequisite]
            if left != right:
                prerequisites[left].add(right)
    ordered: list[int] = []
    remaining = set(prerequisites)
    while remaining:
        ready = sorted((item for item in remaining if not (prerequisites[item] & remaining)), key=lambda item: components[item])
        if not ready:
            raise ValueError("internal error: condensed dependency graph contains a cycle")
        ordered.extend(ready)
        remaining.difference_update(ready)
    order_by_component = {component: position for position, component in enumerate(ordered, 1)}
    blocks: list[dict[str, Any]] = []
    block_of_field: dict[str, str] = {}
    for order, component_index in enumerate(ordered, 1):
        members = components[component_index]
        block_id = f"block-{order:05d}"
        for field_id in members:
            block_of_field[field_id] = block_id
        cyclic = len(members) > 1 or any(field_id in adjacency[field_id] for field_id in members)
        blocks.append({"block_id": block_id, "order": order, "fields": members,
                       "requires": sorted(f"block-{order_by_component[dep]:05d}" for dep in prerequisites[component_index]),
                       "cycle_candidate": cyclic,
                       "status": "unresolved_cycle_candidate" if cyclic else "dependency_ordered"})

    recurrence_fields = sorted({item["consumer"] for item in dependencies.get("internal_dependencies", []) if item.get("recurrence_candidate")})
    cycle_blocks = [item["block_id"] for item in blocks if item["cycle_candidate"]]
    unresolved_fields = {item["consumer"] for item in dependencies.get("unresolved", [])}
    vector_buckets: dict[tuple[Any, ...], list[str]] = defaultdict(list)
    for field_id, item in fields.items():
        if item["role"] == "calculated" and item["physical_shape"] == "1D" and item.get("formula_family"):
            axis = item.get("axis_candidates", [{}])[0].get("axis")
            vector_buckets[(item["sheet"], item["formula_family"], item["member_count"], axis)].append(field_id)
    mutual = {(item["consumer"], item["prerequisite"]) for item in dependencies.get("edges", [])}
    vector_groups = []
    for key, members in sorted(vector_buckets.items(), key=lambda item: (str(item[0][0]).casefold(), str(item[0][1]), item[0][2], str(item[0][3]))):
        members = sorted(members)
        member_set = set(members)
        if len(members) < 2 or any(source in member_set and target in member_set for source, target in mutual):
            continue
        vector_groups.append({"group_id": f"vector-{len(vector_groups) + 1:04d}", "fields": members,
                              "physical_axis_candidate": key[3], "logical_axis_confirmed": False,
                              "compatible_physical_shape": True, "mutual_dependency": False,
                              "executable": False, "reason": "logical axes and numerical behavior are unconfirmed"})
    missing = []
    if recurrence_fields:
        missing.append({"evidence": "recurrence semantics", "fields": recurrence_fields})
    if cycle_blocks:
        missing.append({"evidence": "cycle resolution and iteration rules", "blocks": cycle_blocks})
    if unresolved_fields:
        missing.append({"evidence": "dynamic, broken, or unsupported formula references", "fields": sorted(unresolved_fields)})
    if vector_groups:
        missing.append({"evidence": "confirmed logical axes and CPU/Excel reconciliation", "vector_groups": [item["group_id"] for item in vector_groups]})
    return {"schema_version": "step3.plan.v1", "binding_sha256": binding_sha256, "fields_sha256": fields_sha256, "dependencies_sha256": dependencies_sha256,
            "direction": "prerequisite_before_consumer", "blocks": blocks, "block_of_field": block_of_field,
            "recurrence_candidates": recurrence_fields, "cycle_candidate_blocks": cycle_blocks,
            "vector_proposals": vector_groups, "missing_evidence": missing,
            "readiness": {"draft_only": True, "runtime_verified": False, "generation_ready": False, "gpu_executable": False},
            "summary": {"blocks": len(blocks), "cycle_candidate_blocks": len(cycle_blocks), "recurrence_fields": len(recurrence_fields),
                        "vector_proposals": len(vector_groups), "missing_evidence_items": len(missing)}}


def build_plan(analysis_dir: Path) -> dict[str, Any]:
    try:
        analysis_dir, manifest, _binding, _root, _inventory_path = _load_context(analysis_dir)
        fields_artifact, fields_sha256 = _read_stage(analysis_dir, manifest, "fields")
        dependencies, dependencies_sha256 = _read_stage(analysis_dir, manifest, "dependencies")
        if dependencies.get("fields_sha256") != fields_sha256:
            raise ValueError("dependencies.json is stale against fields.json; rebuild dependencies")
        artifact = _build_plan(fields_artifact, dependencies, manifest["binding_sha256"], fields_sha256, dependencies_sha256)
        digest = _record_stage(analysis_dir, manifest, "plan", artifact)
        report = ["# Step 3 Execution Plan Draft", "", "Dependency order follows prerequisite before consumer. Blocks describe static structure; they do not prove numerical execution order or runtime readiness.",
                  "", f"Blocks: {artifact['summary']['blocks']} · cycle candidates: {artifact['summary']['cycle_candidate_blocks']} · recurrence candidates: {artifact['summary']['recurrence_fields']} · vector proposals: {artifact['summary']['vector_proposals']}", "", "## Initial blocks", ""]
        report.extend(f"- {block['order']}. `{block['block_id']}` · {len(block['fields'])} field(s) · {block['status']}" for block in artifact["blocks"][:30])
        report.extend(["", "## Missing evidence", ""])
        report.extend(f"- {item['evidence']}" for item in artifact["missing_evidence"])
        if not artifact["missing_evidence"]:
            report.append("- Confirm field semantics and logical axes before implementation.")
        _write_bytes(analysis_dir / "execution_plan.md", ("\n".join(report) + "\n").encode("utf-8"))
        return _tool_result("step3.plan", analysis_dir, output=_PLAN, sha256=digest, summary=artifact["summary"], readiness=artifact["readiness"])
    except Exception as exc:
        return _blocked("step3.plan", exc)


def _validate_contract(fields_artifact: dict[str, Any], dependencies: dict[str, Any], plan: dict[str, Any], inventory: dict[str, Any], ignored_sheets: list[str], fields_sha256: str, dependencies_sha256: str) -> list[str]:
    errors: list[str] = []
    fields, cell_fields = _field_index(fields_artifact)
    ignored = {item.casefold() for item in ignored_sheets}
    expected: set[str] = set()
    for sheet in inventory.get("sheets", []):
        if sheet["name"].casefold() in ignored:
            continue
        for cell in sheet.get("cells", []):
            if cell.get("kind") in {"value", "formula"}:
                location = cell.get("source_location") or {}
                expected.add(str(cell.get("source_identity") or location.get("source_identity") or f"{sheet['name']}!{cell['address']}"))
    actual = [member.get("cell_id") for item in fields.values() for member in item.get("members", [])]
    if len(actual) != len(set(actual)):
        errors.append("stored source cells appear in more than one primary field")
    if set(actual) != expected:
        errors.append(f"primary field coverage differs from inventory ({len(expected - set(actual))} missing, {len(set(actual) - expected)} extra)")
    for edge in dependencies.get("edges", []):
        if edge.get("consumer") not in fields or edge.get("prerequisite") not in fields:
            errors.append("dependency edge refers to an unknown primary field")
            break
    if dependencies.get("fields_sha256") != fields_sha256:
        errors.append("dependencies.json predecessor checksum does not match fields.json")
    if plan.get("fields_sha256") != fields_sha256 or plan.get("dependencies_sha256") != dependencies_sha256:
        errors.append("execution_plan.json has stale prerequisite checksums")
    blocks = plan.get("blocks", [])
    block_by_field: dict[str, tuple[int, str]] = {}
    for block in blocks:
        for field_id in block.get("fields", []):
            if field_id in block_by_field:
                errors.append("a primary field appears in multiple execution blocks")
            block_by_field[field_id] = (int(block.get("order", 0)), block["block_id"])
    if set(block_by_field) != set(fields):
        errors.append("execution blocks do not cover every primary field exactly once")
    for edge in dependencies.get("edges", []):
        consumer, prerequisite = edge.get("consumer"), edge.get("prerequisite")
        if consumer in block_by_field and prerequisite in block_by_field:
            consumer_order = block_by_field[consumer][0]
            prerequisite_order = block_by_field[prerequisite][0]
            if block_by_field[consumer][1] != block_by_field[prerequisite][1] and prerequisite_order >= consumer_order:
                errors.append("execution plan places a prerequisite after its consumer")
                break
    return errors


def _handoff_markdown(binding: dict[str, Any], fields: dict[str, Any], dependencies: dict[str, Any], plan: dict[str, Any], validation: dict[str, Any], model_spec: dict[str, Any] | None = None) -> str:
    scope = binding["scope"]
    lines = [
        "# Step 3 Draft Handoff",
        "",
        f"**Source:** `{binding['source']['source_path']}` · source `{binding['source']['source_id']}` · run `{binding['source']['run_id']}`",
        f"**Workbook SHA-256:** `{binding['source']['source_sha256']}`",
        f"**Scope exclusions:** {', '.join(scope['ignored_sheets']) or 'none'}",
        f"**Structural validation:** `{validation['status']}`",
        "",
        "This is a static review draft. Formula caches are recorded only as stored scenario boundaries; no formula execution or numerical equivalence check was performed.",
        "",
        "## Draft counts",
        "",
        f"- Stored cells retained: {fields['summary']['retained_stored_cells']}",
        f"- Primary fields: {fields['summary']['primary_fields']} ({fields['summary']['source_fields']} source, {fields['summary']['calculated_fields']} calculated)",
        f"- Dependency field edges: {dependencies['summary']['field_edges']}; unresolved groups: {dependencies['summary']['unresolved_groups']}; boundaries: {dependencies['summary']['reference_boundaries']}",
        f"- Dependency blocks: {plan['summary']['blocks']}; recurrence fields: {plan['summary']['recurrence_fields']}; vector proposals: {plan['summary']['vector_proposals']}",
        "",
        "## Readiness",
        "",
        "- Runtime verified: no",
        "- Generation ready: no",
        "- GPU executable: no",
        "",
        "## Missing evidence",
        "",
    ]
    lines.extend(f"- {item['evidence']}" for item in plan.get("missing_evidence", []))
    if not plan.get("missing_evidence"):
        lines.append("- Confirm field semantics and logical axes before implementation.")
    if model_spec:
        lines.extend(["", "## Semantic draft", "", "Validated semantic specification: `model_spec.json` and `model_spec.md`."])
    lines.extend(["", "See `fields.json`, `dependencies.json`, `execution_plan.json`, and `validation.json` for machine-readable detail.", ""])
    return "\n".join(lines)


def validate_analysis(analysis_dir: Path, spec_path: Path | None = None) -> dict[str, Any]:
    try:
        analysis_dir, manifest, binding, root, inventory_path = _load_context(analysis_dir)
        checked = _validate_saved_index(Path(binding["index_path"]), root)
        if checked.get("status") != "pass":
            raise ValueError("Step 2 index or referenced artifacts no longer validate")
        fields_artifact, fields_sha256 = _read_stage(analysis_dir, manifest, "fields")
        dependencies, dependencies_sha256 = _read_stage(analysis_dir, manifest, "dependencies")
        plan, plan_sha256 = _read_stage(analysis_dir, manifest, "plan")
        if dependencies.get("fields_sha256") != fields_sha256:
            raise ValueError("dependencies.json is stale against fields.json; rebuild dependencies")
        if plan.get("fields_sha256") != fields_sha256 or plan.get("dependencies_sha256") != dependencies_sha256:
            raise ValueError("execution_plan.json is stale; rebuild the plan")
        inventory = _load_inventory(inventory_path)
        if inventory.get("workbook_sha256") != binding["source"]["source_sha256"]:
            raise ValueError("inventory workbook_sha256 does not match the selected source")
        errors = _validate_contract(fields_artifact, dependencies, plan, inventory, binding["scope"]["ignored_sheets"], fields_sha256, dependencies_sha256)
        validation = {"schema_version": "step3.validation.v1", "binding_sha256": manifest["binding_sha256"],
                      "fields_sha256": fields_sha256, "dependencies_sha256": dependencies_sha256, "plan_sha256": plan_sha256,
                      "status": "pass" if not errors else "blocked", "errors": errors,
                      "readiness": {"draft_only": True, "runtime_verified": False, "generation_ready": False, "gpu_executable": False},
                      "step2_validation": checked}
        validation_sha256 = _write_json(analysis_dir / "validation.json", validation)
        preserved_spec = None
        try:
            previous_handoff = _read_json(analysis_dir / "handoff.json")
            spec_ref = previous_handoff.get("model_spec") if isinstance(previous_handoff, dict) else None
            if (previous_handoff.get("binding_sha256") == manifest["binding_sha256"]
                    and isinstance(spec_ref, dict)
                    and spec_ref.get("fields_sha256") == fields_sha256
                    and spec_ref.get("dependencies_sha256") == dependencies_sha256):
                spec_json = analysis_dir / "model_spec.json"
                spec_markdown = analysis_dir / "model_spec.md"
                spec_value = _read_json(spec_json)
                if (_hash_file(spec_json) == spec_ref.get("json_sha256")
                        and _hash_file(spec_markdown) == spec_ref.get("markdown_sha256")
                        and spec_value.get("binding_sha256") == manifest["binding_sha256"]
                        and spec_value.get("stage_hashes") == {"fields": fields_sha256, "dependencies": dependencies_sha256}):
                    preserved_spec = spec_ref
        except (OSError, ValueError, AttributeError):
            preserved_spec = None
        handoff = {"schema_version": "step3.handoff.v1", "binding_sha256": manifest["binding_sha256"],
                   "source": binding["source"], "scope": binding["scope"], "status": validation["status"],
                   "fields_sha256": fields_sha256, "dependencies_sha256": dependencies_sha256, "plan_sha256": plan_sha256,
                   "validation_sha256": validation_sha256, "readiness": validation["readiness"],
                   "counts": {"fields": fields_artifact["summary"], "dependencies": dependencies["summary"], "plan": plan["summary"]}}
        if preserved_spec:
            handoff["model_spec"] = preserved_spec
        _write_json(analysis_dir / "handoff.json", handoff)
        _write_bytes(analysis_dir / "handoff.md", _handoff_markdown(binding, fields_artifact, dependencies, plan, validation, preserved_spec).encode("utf-8"))
        result = _tool_result("step3.validate", analysis_dir, output="handoff.md", validation=validation, handoff=handoff)
        if spec_path is not None and not errors:
            from excel_to_act.steps.step3.exploration import validate_model_spec

            semantic = validate_model_spec(analysis_dir, spec_path)
            result["model_spec"] = semantic
            if semantic["status"] != "pass":
                result["status"] = "blocked"
        if errors:
            result["status"] = "blocked"
        return result
    except Exception as exc:
        return _blocked("step3.validate", exc)


def tool_catalog() -> dict[str, Any]:
    from excel_to_act.steps.step3.tools import tool_catalog as catalog

    return catalog()
