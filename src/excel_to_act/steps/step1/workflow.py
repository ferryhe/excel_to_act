"""Human Step1 conversion, independent checks, tool loop, and promotion gate."""

from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from collections import Counter
from datetime import UTC, datetime
from pathlib import Path, PurePosixPath
from typing import Any
from uuid import uuid4

from openpyxl.utils.cell import coordinate_to_tuple

from excel_to_act.graph.builder import RegexFormulaGraphBuilder
from excel_to_act.ingest.ooxml_package import OPAQUE_MARKERS
from excel_to_act.ingest.openpyxl_reader import OpenpyxlWorkbookReader
from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor
from excel_to_act.steps.step1.source_scan import SourceScanError, object_identity, scan_step1_source
from excel_to_act.schemas import (
    CellInventory,
    CellKind,
    CoverageSummary,
    RangeInventory,
    SheetInventory,
    SourceLocation,
    WorkbookInventory,
    WorkbookManifest,
)

_SUPPORTED = {".xlsx", ".xlsm"}
_OPAQUE_PARTS = set(OPAQUE_MARKERS)
_ARTIFACT_FILES = (
    "source.json",
    "workbook_manifest.json",
    "inventory.json",
    "source_facts.json",
    "logical_objects.json",
    "package_parts.json",
)
_SOURCE_METRIC_FIELDS = (
    "logical_objects_total",
    "logical_objects_accounted",
    "traceability_ratio",
    "parsed_objects_total",
    "parsed_objects",
    "parsed_coverage_ratio",
    "package_parts_total",
    "parsed_package_parts",
    "parsed_package_parts_total",
    "parsed_package_parts_ratio",
    "package_parts_preserved",
    "package_preservation_ratio",
    "opaque_parts",
    "opaque_rate",
    "supported_facts_total",
    "supported_facts_exact",
    "supported_fidelity_ratio",
    "fidelity_deviations",
)
_SUPPORTED_OBJECT_KINDS = {
    "sheet",
    "cell",
    "defined_name",
    "table",
    "merged_range",
    "data_table",
    "form_control",
    "row_layout",
    "column_layout",
    "freeze_panes",
    "data_validation",
    "conditional_formatting",
    "comment",
    "hyperlink",
    "sheet_protection",
}


def _coverage_scope() -> dict[str, Any]:
    return {
        "logical_object_kinds": sorted(_SUPPORTED_OBJECT_KINDS),
        "cell_unit": "One worksheet <c> element containing a formula, value, or inline-string element; typed empty inline strings count, style-only blank cells do not.",
        "package_part_unit": "One record per non-directory OOXML ZIP member, with source name, size, SHA-256, content type, and copied-byte verification.",
        "parsed_logical_object_ratio": "parsed_coverage_ratio: supported source declarations with exact inventory identities divided by supported logical-object declarations; independent of package-part counts.",
        "parsed_package_parts_ratio": "parsed_package_parts_ratio: non-opaque names in the fresh source scan's parsed_parts set divided by all non-directory OOXML ZIP members; opaque members never count as parsed. This measures package-part parsing, not universal Excel semantic coverage.",
        "opaque_package_parts_ratio": "opaque_rate: opaque source package members divided by all non-directory OOXML ZIP members; opaque members are preserved, never counted as parsed.",
        "opaque_policy": "Opaque is package-part-level: bytes are preserved and checksummed; opaque members never count in parsed_package_parts_ratio. Supported logical objects keep their separate ledger.",
        "scope_limit": "This is exact accounting for the listed supported logical kinds, not universal semantic coverage of every Excel feature. Other structures remain explicitly opaque when the scanner cannot parse them.",
    }


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True, default=str) + "\n").encode("utf-8")


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_json_bytes(value))


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


class _AttemptHistoryError(ValueError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def _read_attempt_history(run_dir: Path) -> dict[str, Any]:
    path = run_dir / "attempt_history.json"
    try:
        history = _read_json(path)
    except (OSError, ValueError) as exc:
        raise _AttemptHistoryError(
            "attempt_history_unreadable",
            f"Cannot read attempt_history.json: {type(exc).__name__}: {exc}",
        ) from exc
    if not isinstance(history, dict):
        raise _AttemptHistoryError("attempt_history_invalid", "attempt_history.json must contain a JSON object.")

    attempts = history.get("attempts", [])
    attempt_limit = history.get("attempt_limit", 3)
    no_progress_stopped = history.get("no_progress_stopped", False)
    if not isinstance(attempts, list):
        raise _AttemptHistoryError("attempt_history_invalid", "attempt_history.json attempts must be a list.")
    if isinstance(attempt_limit, bool) or not isinstance(attempt_limit, int) or not 1 <= attempt_limit <= 3:
        raise _AttemptHistoryError("attempt_history_invalid", "attempt_history.json attempt_limit must be an integer from 1 to 3.")
    if not isinstance(no_progress_stopped, bool):
        raise _AttemptHistoryError("attempt_history_invalid", "attempt_history.json no_progress_stopped must be a boolean.")
    for index, attempt in enumerate(attempts):
        if not isinstance(attempt, dict):
            raise _AttemptHistoryError("attempt_history_invalid", f"attempt_history.json attempts[{index}] must be an object.")
        no_progress_count = attempt.get("no_progress_count", 0)
        if isinstance(no_progress_count, bool) or not isinstance(no_progress_count, int) or no_progress_count < 0:
            raise _AttemptHistoryError(
                "attempt_history_invalid",
                f"attempt_history.json attempts[{index}].no_progress_count must be a non-negative integer.",
            )
    return history


def _hash_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _unavailable_source_metrics() -> dict[str, None]:
    return dict.fromkeys(_SOURCE_METRIC_FIELDS)


def _ratio(value: float, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not 0 <= value <= 1:
        raise ValueError(f"{name} must be between 0 and 1")
    return float(value)


def _duplicate_keys(values: list[str]) -> list[str]:
    return sorted(value for value, count in Counter(values).items() if count > 1)


def _relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def _run_source(run_dir: Path) -> dict[str, Any]:
    source = _read_json(run_dir / "source.json")
    if not isinstance(source, dict):
        raise ValueError("source.json must contain a JSON object")
    required = {"schema_version", "batch_id", "run_id", "source_path", "relative_path", "sha256", "suffix", "thresholds", "allow_opaque", "attempt_limit"}
    missing = required - source.keys()
    if missing:
        raise ValueError(f"source.json is missing required fields: {', '.join(sorted(missing))}")
    if not isinstance(source.get("thresholds"), dict):
        raise ValueError("source.json thresholds must be an object")
    threshold_defaults = {"traceability": 1.0, "fidelity": 1.0, "parsed_package_min": 0.0, "opaque_max": 1.0}
    for name, default in threshold_defaults.items():
        value = source["thresholds"].get(name, default)
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not 0 <= value <= 1:
            raise ValueError(f"source.json {name} threshold must be a number between 0 and 1")
    if source.get("sha256") is not None and not isinstance(source.get("sha256"), str):
        raise ValueError("source.json sha256 must be a string or null")
    for name in ("batch_id", "run_id", "source_path", "relative_path", "suffix"):
        if not isinstance(source.get(name), str):
            raise ValueError(f"source.json {name} must be a string")
    if not isinstance(source.get("allow_opaque"), bool):
        raise ValueError("source.json allow_opaque must be a boolean")
    if isinstance(source.get("attempt_limit"), bool) or not isinstance(source.get("attempt_limit"), int):
        raise ValueError("source.json attempt_limit must be an integer")
    return source


def _safe_run_source(run_dir: Path) -> dict[str, Any]:
    try:
        return _run_source(run_dir)
    except (OSError, ValueError, json.JSONDecodeError):
        batch_id = run_dir.parents[3].name if len(run_dir.parents) > 3 else None
        return {"run_id": run_dir.name, "batch_id": batch_id, "source_path": None, "relative_path": None, "sha256": None}


def _run_error(tool: str, run_dir: Path, message: str) -> dict[str, Any]:
    return {
        "tool": tool,
        "status": "error",
        "source": None,
        "run_id": None,
        "artifacts": [],
        "metrics_state": "unavailable",
        "metrics": {},
        "diagnostics": [_diagnostic("run_unavailable", "error", f"{message}: {run_dir}")],
        "retryable": False,
        "next_tool": None,
    }


def _source_object_key(kind: str, details: dict[str, Any]) -> tuple[str, str, str]:
    sheet = str(details.get("sheet", ""))
    if kind == "defined_name":
        key = f"{details.get('scope', 'workbook')}|{details.get('name', '')}"
    elif kind == "table":
        key = str(details.get("name", ""))
    elif kind == "data_table":
        key = str(details.get("corner_cell", details.get("address", "")))
    elif kind == "form_control":
        key = str(details.get("control_id", details.get("control_type", "")))
    elif kind in {"sheet_protection"}:
        key = ""
    else:
        key = str(details.get("address", ""))
    return (kind, sheet, key)


def _record_source_identity(record: Any, identity: str) -> None:
    record.source_identity = identity
    record.source_location.source_identity = identity


def _add_fallback_range(inventory: WorkbookInventory, source_obj: dict[str, Any], sheet_indexes: dict[str, int]) -> RangeInventory:
    kind = source_obj["kind"]
    details = source_obj.get("details", {})
    sheet = details.get("sheet")
    address = str(details.get("address", details.get("corner_cell", "")))
    loc = SourceLocation(
        workbook_path=source_obj["source_location"].get("workbook_path"),
        sheet_name=sheet,
        sheet_index=sheet_indexes.get(str(sheet)) if sheet else None,
        address=address or None if sheet else None,
        object_type=kind,
        object_id=source_obj["identity"],
        ooxml_part=source_obj.get("part"),
        source_identity=source_obj["identity"],
    )
    item = RangeInventory(
        source_location=loc,
        name=details.get("name"),
        address=address,
        kind=kind,
        metadata={"source_details": details},
        source_identity=source_obj["identity"],
    )
    if sheet is None:
        inventory.workbook_ranges.append(item)
    else:
        target = next((s for s in inventory.sheets if s.name == sheet), None)
        if target is None:
            target = SheetInventory(
                source_location=SourceLocation(
                    workbook_path=loc.workbook_path,
                    sheet_name=str(sheet),
                    sheet_index=sheet_indexes.get(str(sheet)),
                    object_type="sheet",
                    object_id=str(sheet),
                    ooxml_part=source_obj.get("part"),
                ),
                name=str(sheet),
                index=sheet_indexes.get(str(sheet), len(inventory.sheets)),
                max_row=0,
                max_column=0,
            )
            inventory.sheets.append(target)
        target.ranges.append(item)
    return item


def _inventory_record_entries(inventory: WorkbookInventory) -> list[tuple[Any, tuple[str, tuple[str, int, str | None] | None]]]:
    records: list[tuple[Any, tuple[str, tuple[str, int, str | None] | None]]] = []
    for sheet in inventory.sheets:
        records.append((sheet, ("sheets", None)))
        parent = (sheet.name, sheet.index, sheet.source_identity)
        records.extend((record, ("cells", parent)) for record in sheet.cells)
        records.extend((record, ("ranges", parent)) for record in sheet.ranges)
        records.extend((record, ("layout_objects", parent)) for record in sheet.layout_objects)
    records.extend((record, ("workbook_ranges", None)) for record in inventory.workbook_ranges)
    return records


def _inventory_records(inventory: WorkbookInventory) -> list[Any]:
    return [record for record, _container in _inventory_record_entries(inventory)]


def _inventory_ownership(inventory: WorkbookInventory) -> dict[str, tuple[str, tuple[str, int, str | None] | None]]:
    return {
        record.source_identity: container
        for record, container in _inventory_record_entries(inventory)
        if isinstance(record.source_identity, str) and record.source_identity
    }


def _unidentified_inventory_projection(inventory: WorkbookInventory) -> Counter[tuple[Any, bytes]]:
    return Counter(
        (container, _json_bytes(record.model_dump(mode="json")))
        for record, container in _inventory_record_entries(inventory)
        if not record.source_identity
    )


def _inventory_ids(inventory: WorkbookInventory) -> set[str]:
    return {record.source_identity for record in _inventory_records(inventory) if getattr(record, "source_identity", None)}


def _record_matches(record: RangeInventory, source_obj: dict[str, Any]) -> bool:
    kind = source_obj["kind"]
    details = source_obj.get("details", {})
    if record.kind != kind:
        return False
    if kind == "defined_name":
        return record.name == details.get("name") and record.metadata.get("scope", "workbook") == details.get("scope", "workbook")
    if record.source_location.sheet_name != details.get("sheet"):
        return False
    if kind == "table":
        return record.name == details.get("name")
    if kind == "data_table":
        return record.metadata.get("corner_cell") == details.get("corner_cell") or record.address == details.get("address")
    if kind == "form_control":
        return record.metadata.get("control_id") == details.get("control_id")
    if kind == "sheet_protection":
        return True
    return str(record.address or "") == str(details.get("address", ""))


def _decorate_inventory(inventory: WorkbookInventory, scan: dict[str, Any]) -> WorkbookInventory:
    """Attach exact identities and raw-cell transforms to the legacy inventory."""

    sheet_indexes = {sheet["name"]: sheet["index"] for sheet in scan["sheets"]}
    by_name = {sheet.name: sheet for sheet in inventory.sheets}
    for source_sheet in scan["sheets"]:
        sheet = by_name.get(source_sheet["name"])
        if sheet is None:
            loc = SourceLocation(
                workbook_path=inventory.sheets[0].source_location.workbook_path if inventory.sheets else None,
                sheet_name=source_sheet["name"],
                sheet_index=source_sheet["index"],
                object_type="sheet",
                object_id=source_sheet["name"],
                ooxml_part=source_sheet["part"],
            )
            sheet = SheetInventory(
                source_location=loc,
                name=source_sheet["name"],
                index=source_sheet["index"],
                max_row=0,
                max_column=0,
                state=source_sheet["state"],
            )
            inventory.sheets.append(sheet)
            by_name[sheet.name] = sheet
        sheet.source_location.ooxml_part = source_sheet["part"]
        sheet.source_identity = source_sheet["identity"]
        sheet.source_location.source_identity = source_sheet["identity"]

    cells_by_key = {(sheet.name, cell.address): cell for sheet in inventory.sheets for cell in sheet.cells}
    scan_cells = {(cell["sheet"], cell["address"]): cell for cell in scan["cells"]}
    for key, raw in scan_cells.items():
        sheet = by_name[key[0]]
        cell = cells_by_key.get(key)
        if cell is None:
            row, column = coordinate_to_tuple(raw["address"])
            cell = CellInventory(
                source_location=SourceLocation(
                    workbook_path=inventory.sheets[0].source_location.workbook_path,
                    sheet_name=key[0],
                    sheet_index=sheet.index,
                    address=key[1],
                    object_type="cell",
                    object_id=f"{key[0]}!{key[1]}",
                    ooxml_part=raw["sheet_part"],
                ),
                address=key[1],
                row=row,
                column=column,
                kind=CellKind.formula if raw["formula_present"] else CellKind.literal,
            )
            sheet.cells.append(cell)
            cells_by_key[key] = cell
        cell.source_identity = raw["identity"]
        cell.source_location.source_identity = raw["identity"]
        cell.source_location.ooxml_part = raw["sheet_part"]
        cell.kind = CellKind.formula if raw["formula_present"] else CellKind.literal
        cell.value = raw["normalized_value"]
        cell.formula = raw["normalized_formula"]
        cell.cached_value = raw["normalized_cached_value"]
        cell.cached_value_available = raw["normalized_cached_value_available"]
        cell.data_type = "f" if raw["formula_present"] else {"s": "s", "inlineStr": "s", "str": "s", "b": "b", "e": "e", "d": "d"}.get(raw["cell_type"], "n")
        cell.number_format = raw["number_format"]
        cell.style_id = raw["style_id"]
        cell.raw_value_text = raw["raw_value_text"]
        cell.raw_formula_text = raw["raw_formula_text"]
        cell.raw_formula_attributes = raw["raw_formula_attributes"]
        cell.formula_present = raw["formula_present"]
        cell.ooxml_cell_type = raw["cell_type"]
        cell.cached_text = raw["cached_text"]
        cell.cached_text_present = raw["cached_text_present"]
        cell.workbook_date_system = raw["workbook_date_system"]
        cell.date_serial_text = raw["date_serial_text"]

    ranges = [item for sheet in inventory.sheets for item in [*sheet.ranges, *sheet.layout_objects]] + inventory.workbook_ranges
    used: set[int] = set()
    for source_obj in scan["objects"]:
        if source_obj["kind"] == "sheet":
            continue
        if source_obj["kind"] == "cell":
            cell = cells_by_key.get((source_obj.get("details", {}).get("sheet"), source_obj.get("details", {}).get("address")))
            if cell:
                cell.source_identity = source_obj["identity"]
            continue
        matched = next((item for item in ranges if id(item) not in used and _record_matches(item, source_obj)), None)
        if matched is None:
            matched = _add_fallback_range(inventory, source_obj, sheet_indexes)
            ranges.append(matched)
        used.add(id(matched))
        matched.source_identity = source_obj["identity"]
        matched.source_location.source_identity = source_obj["identity"]
        matched.source_location.ooxml_part = source_obj.get("part")
        matched.source_location.object_id = source_obj["identity"]

    # Step1 coverage counts only declared logical objects; diagnostics and parts
    # have their own ledgers and are deliberately not added to this denominator.
    ids = _inventory_ids(inventory)
    inventory.coverage = CoverageSummary(
        recognized_inventory_objects=len(ids),
        unsupported_or_opaque_objects=0,
        discovered_workbook_objects=len(scan["objects"]),
    )
    return inventory


def _build_conversion(source_path: Path, scan: dict[str, Any]) -> tuple[WorkbookManifest, WorkbookInventory]:
    manifest = OpenpyxlWorkbookReader().read_manifest(source_path)
    part_facts = scan["parts"]
    for part in manifest.package_parts:
        source_part = part_facts.get(part.name)
        if source_part is not None:
            part.opaque = bool(source_part.get("opaque"))
    inventory = OpenpyxlInventoryExtractor().extract(source_path, manifest)
    return manifest, _decorate_inventory(inventory, scan)


def _opaque(name: str, metadata: dict[str, Any] | None = None) -> str | None:
    if metadata is not None:
        return (metadata.get("opaque_reason") or "unparsed package part") if metadata.get("opaque") else None
    return next((description for marker, description in OPAQUE_MARKERS.items() if marker in name), None)


def _preserve_parts(source: Path, run_dir: Path, scan: dict[str, Any]) -> list[dict[str, Any]]:
    records = []
    with zipfile.ZipFile(source) as zf:
        for name, metadata in scan["parts"].items():
            payload = zf.read(name)
            if _hash_bytes(payload) != metadata["sha256"]:
                raise ValueError(f"source package part changed while reading: {name}")
            safe_name = f"{_hash_bytes(name.encode('utf-8'))[:16]}-{PurePosixPath(name).name or 'part'}"
            part_path = run_dir / "source_parts" / safe_name
            part_path.parent.mkdir(parents=True, exist_ok=True)
            part_path.write_bytes(payload)
            records.append(
                {
                    "name": name,
                    "source_identity": object_identity("package_part", name),
                    "size": metadata["size"],
                    "sha256": metadata["sha256"],
                    "content_type": metadata["content_type"],
                    "opaque": bool(metadata.get("opaque")),
                    "opaque_reason": metadata.get("opaque_reason"),
                    "source_location": {"ooxml_part": name, "object_type": "package_part", "object_id": name},
                    "preserved_path": _relative(part_path, run_dir),
                }
            )
    return records


def _write_base(source_path: Path, run_dir: Path, scan: dict[str, Any]) -> tuple[WorkbookManifest, WorkbookInventory, list[dict[str, Any]]]:
    manifest, inventory = _build_conversion(source_path, scan)
    _write_json(run_dir / "workbook_manifest.json", manifest.model_dump(mode="json"))
    _write_json(run_dir / "inventory.json", inventory.model_dump(mode="json"))
    _write_json(run_dir / "source_facts.json", {"date_system": scan["date_system"], "cells": scan["cells"]})
    ids = _inventory_ids(inventory)
    logical = [
        {**item, "accounted": item["identity"] in ids, "opaque": False}
        for item in scan["objects"]
    ]
    _write_json(run_dir / "logical_objects.json", logical)
    parts = _preserve_parts(source_path, run_dir, scan)
    _write_json(run_dir / "package_parts.json", parts)
    evidence = run_dir / "source_evidence" / source_path.name
    evidence.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source_path, evidence)
    if _hash_file(evidence) != _hash_file(source_path):
        raise ValueError("source evidence copy checksum mismatch")
    return manifest, inventory, parts


def discover_inputs(input_dir: Path, out_dir: Path) -> list[tuple[Path, str, str]]:
    """Return every regular file in deterministic relative-path order."""

    root = input_dir.expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"input directory does not exist: {root}")
    output = out_dir.expanduser().resolve()
    excludes_output = output == root or root in output.parents
    found: list[tuple[Path, str, str]] = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        resolved = path.resolve()
        if excludes_output and (resolved == output or output in resolved.parents):
            continue
        relative = path.relative_to(root).as_posix()
        found.append((path, relative, path.suffix.lower()))
    return sorted(found, key=lambda row: (row[1].casefold(), row[1]))


def _diagnostic(code: str, severity: str, message: str, source_location: dict[str, Any] | None = None) -> dict[str, Any]:
    return {"code": code, "severity": severity, "message": message, "source_location": source_location}


def _artifact_refs(run_dir: Path) -> list[dict[str, Any]]:
    refs = []
    for name in (*_ARTIFACT_FILES, "quality.json"):
        path = run_dir / name
        if path.is_file():
            refs.append({"name": name, "path": name, "sha256": _hash_file(path)})
    return refs


def _inventory_record_map(inventory: WorkbookInventory) -> dict[str, Any]:
    return {record.source_identity: record for record in _inventory_records(inventory) if getattr(record, "source_identity", None)}


def _candidate_summary(run_dir: Path) -> dict[str, Any]:
    try:
        return _read_json(run_dir / "quality.json")
    except (OSError, ValueError, json.JSONDecodeError):
        return {}


def _valid_final_link(run_dir: Path, quality: dict[str, Any]) -> str | None:
    """Keep a prior final link only while the candidate still passes its gate."""

    if not quality.get("ready_for_next_step"):
        return None
    try:
        previous = _read_json(run_dir / "handoff.json")
        final_path = previous.get("final_output")
        if not final_path:
            return None
        relative = PurePosixPath(str(final_path))
        if relative.is_absolute() or ".." in relative.parts:
            return None
        output_root = run_dir.parents[5].resolve()
        final_dir = (output_root / Path(*relative.parts)).resolve()
        final_dir.relative_to(output_root)
        promotion = _read_json(final_dir / "promotion.json")
        if (
            promotion.get("source_sha256") != quality.get("source_sha256")
            or promotion.get("run_id") != quality.get("run_id")
            or promotion.get("quality_status") != quality.get("status")
        ):
            return None
        promoted_handoff = _read_json(final_dir / "handoff.json")
        if promoted_handoff.get("final_output") != relative.as_posix():
            return None
        return relative.as_posix()
    except (OSError, ValueError, IndexError, json.JSONDecodeError):
        return None


def _check_run(run_dir: Path, *, write: bool = True) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    diagnostics: list[dict[str, Any]] = []
    metrics: dict[str, Any] = {}
    metrics_state = "unavailable"
    source: dict[str, Any] = {}
    fresh_scan: dict[str, Any] | None = None
    fresh_manifest: WorkbookManifest | None = None
    fresh_inventory: WorkbookInventory | None = None
    hard_blockers: list[str] = []
    fidelity_deviations: list[dict[str, Any]] = []
    attempt_history: dict[str, Any] | None = None
    attempt_history_error: _AttemptHistoryError | None = None
    try:
        attempt_history = _read_attempt_history(run_dir)
    except _AttemptHistoryError as exc:
        attempt_history_error = exc
        hard_blockers.append("recovery history is unreadable or invalid")
        diagnostics.append(_diagnostic(exc.code, "error", f"{exc}; stop recovery without changing attempt_history.json."))

    try:
        source = _run_source(run_dir)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        diagnostics.append(_diagnostic("source_record_unreadable", "error", f"Cannot read source.json: {exc}"))
        hard_blockers.append("source record is unreadable")

    original = Path(source.get("source_path", "")) if source.get("source_path") else None
    if original is None or not original.is_file():
        diagnostics.append(_diagnostic("source_unavailable", "error", "The original source path is unavailable."))
        hard_blockers.append("original source is unavailable")
    elif not source.get("sha256"):
        diagnostics.append(_diagnostic("source_hash_unavailable", "error", "The original source hash was not captured; rerun conversion while the source is readable."))
        hard_blockers.append("original source hash is unavailable")
    else:
        try:
            source_hash = _hash_file(original)
        except OSError as exc:
            diagnostics.append(_diagnostic("source_hash_unreadable", "error", f"Cannot read original source for hashing: {type(exc).__name__}: {exc}"))
            hard_blockers.append("original source cannot be read for hashing")
            source_hash = None
        if source_hash is not None and source_hash != source.get("sha256"):
            diagnostics.append(_diagnostic("source_changed", "error", "The source hash changed after conversion."))
            hard_blockers.append("source hash changed after conversion")
        elif source_hash is not None:
            try:
                fresh_scan = scan_step1_source(original)
                fresh_manifest, fresh_inventory = _build_conversion(original, fresh_scan)
            except (OSError, ValueError, zipfile.BadZipFile, SourceScanError, Exception) as exc:
                diagnostics.append(_diagnostic("source_xml_unreadable", "error", f"Source scan failed: {type(exc).__name__}: {exc}"))
                hard_blockers.append("source XML could not be read")

    candidate_inventory: WorkbookInventory | None = None
    try:
        candidate_inventory = WorkbookInventory.model_validate(_read_json(run_dir / "inventory.json"))
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        diagnostics.append(_diagnostic("inventory_unreadable", "error", f"Cannot read inventory.json: {exc}"))
        hard_blockers.append("candidate inventory is missing or invalid")
    else:
        inventory_identities = [record.source_identity for record in _inventory_records(candidate_inventory) if isinstance(record.source_identity, str) and record.source_identity]
        duplicate_identities = _duplicate_keys(inventory_identities)
        if duplicate_identities:
            diagnostics.append(_diagnostic("duplicate_inventory_identity", "error", f"inventory.json repeats {len(duplicate_identities)} source identity value(s); rerun inventory extraction."))
            hard_blockers.append("candidate inventory contains duplicate source identities")
            candidate_inventory = None

    candidate_manifest: WorkbookManifest | None = None
    try:
        candidate_manifest = WorkbookManifest.model_validate(_read_json(run_dir / "workbook_manifest.json"))
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        diagnostics.append(_diagnostic("manifest_unreadable", "error", f"Cannot read workbook_manifest.json: {exc}"))
        hard_blockers.append("candidate manifest is missing or invalid")

    candidate_facts: dict[str, Any] | None = None
    try:
        decoded_facts = _read_json(run_dir / "source_facts.json")
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        diagnostics.append(_diagnostic("source_facts_unreadable", "error", f"Cannot read source_facts.json: {exc}"))
        hard_blockers.append("candidate source facts are missing or invalid")
    else:
        facts_cells = decoded_facts.get("cells") if isinstance(decoded_facts, dict) else None
        if (
            not isinstance(decoded_facts, dict)
            or decoded_facts.get("date_system") not in {"1900", "1904"}
            or not isinstance(facts_cells, list)
            or any(not isinstance(cell, dict) or not isinstance(cell.get("identity"), str) or not cell.get("identity") for cell in facts_cells)
        ):
            diagnostics.append(_diagnostic("source_facts_invalid", "error", "source_facts.json must be an object with a 1900/1904 date_system and a cells list of records with string identities."))
            hard_blockers.append("candidate source facts have an invalid structure")
        else:
            duplicate_identities = _duplicate_keys([cell["identity"] for cell in facts_cells])
            if duplicate_identities:
                diagnostics.append(_diagnostic("duplicate_source_fact_identity", "error", f"source_facts.json repeats {len(duplicate_identities)} cell identity value(s); refresh source facts."))
                hard_blockers.append("candidate source facts contain duplicate cell identities")
            else:
                candidate_facts = decoded_facts

    candidate_logical: list[dict[str, Any]] | None = None
    try:
        decoded_logical = _read_json(run_dir / "logical_objects.json")
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        diagnostics.append(_diagnostic("logical_ledger_unreadable", "error", f"Cannot read logical_objects.json: {exc}"))
        hard_blockers.append("candidate logical ledger is missing or invalid")
    else:
        if not isinstance(decoded_logical, list) or any(not isinstance(obj, dict) or not isinstance(obj.get("identity"), str) or not obj.get("identity") for obj in decoded_logical):
            diagnostics.append(_diagnostic("logical_ledger_invalid", "error", "logical_objects.json must be a list of records with string identities."))
            hard_blockers.append("candidate logical ledger has an invalid structure")
        else:
            duplicate_identities = _duplicate_keys([obj["identity"] for obj in decoded_logical])
            if duplicate_identities:
                diagnostics.append(_diagnostic("duplicate_logical_object_identity", "error", f"logical_objects.json repeats {len(duplicate_identities)} object identity value(s); refresh the logical ledger."))
                hard_blockers.append("candidate logical ledger contains duplicate object identities")
            else:
                candidate_logical = decoded_logical

    candidate_parts: list[dict[str, Any]] | None = None
    try:
        decoded_parts = _read_json(run_dir / "package_parts.json")
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        diagnostics.append(_diagnostic("package_ledger_unreadable", "error", f"Cannot read package_parts.json: {exc}"))
        hard_blockers.append("candidate package ledger is missing or invalid")
    else:
        if not isinstance(decoded_parts, list) or any(not isinstance(part, dict) or not isinstance(part.get("name"), str) or not part.get("name") for part in decoded_parts):
            diagnostics.append(_diagnostic("package_ledger_invalid", "error", "package_parts.json must be a list of records with string names."))
            hard_blockers.append("candidate package ledger has an invalid structure")
        else:
            duplicate_names = _duplicate_keys([part["name"] for part in decoded_parts])
            if duplicate_names:
                diagnostics.append(_diagnostic("duplicate_package_part_name", "error", f"package_parts.json repeats {len(duplicate_names)} part name value(s); preserve package parts again."))
                hard_blockers.append("candidate package ledger contains duplicate part names")
            else:
                candidate_parts = decoded_parts

    if fresh_scan is not None and fresh_manifest is not None and fresh_inventory is not None:
        expected_cells = {cell["identity"]: cell for cell in fresh_scan["cells"]}
        expected_objects = {obj["identity"]: obj for obj in fresh_scan["objects"]}
        expected_parts = fresh_scan["parts"]
        parsed_source_parts = {
            name
            for name in fresh_scan.get("parsed_parts", [])
            if name in expected_parts and not _opaque(name, expected_parts[name])
        }

        expected_ownership = _inventory_ownership(fresh_inventory)
        actual_ownership = _inventory_ownership(candidate_inventory) if candidate_inventory else {}
        wrong_parent_ids = [
            identity
            for identity in expected_ownership.keys() & actual_ownership.keys()
            if expected_ownership[identity] != actual_ownership[identity]
        ]
        if wrong_parent_ids:
            hard_blockers.append("candidate inventory records are attached to the wrong worksheet or collection")
            diagnostics.append(
                _diagnostic(
                    "inventory_parent_mismatch",
                    "error",
                    f"{len(wrong_parent_ids)} source-identified record(s) have the wrong worksheet or collection owner; rerun inventory extraction.",
                )
            )

        expected_unidentified = _unidentified_inventory_projection(fresh_inventory)
        actual_unidentified = _unidentified_inventory_projection(candidate_inventory) if candidate_inventory else Counter()
        missing_unidentified = sum((expected_unidentified - actual_unidentified).values())
        extra_unidentified = sum((actual_unidentified - expected_unidentified).values())
        if missing_unidentified or extra_unidentified:
            hard_blockers.append("identity-less inventory projections differ from the source-derived projection")
            diagnostics.append(
                _diagnostic(
                    "inventory_projection_mismatch",
                    "error",
                    f"Identity-less inventory projections differ (missing={missing_unidentified}, extra={extra_unidentified}); rerun inventory extraction.",
                )
            )

        if candidate_manifest is None or candidate_manifest.sha256 != source.get("sha256") or _json_bytes(candidate_manifest.model_dump(mode="json")) != _json_bytes(fresh_manifest.model_dump(mode="json")):
            hard_blockers.append("candidate manifest differs from the current source")
            diagnostics.append(_diagnostic("manifest_mismatch", "error", "workbook_manifest.json is missing, changed, or describes another source."))

        if candidate_facts is not None:
            actual_cells = {cell.get("identity"): cell for cell in candidate_facts.get("cells", [])}
            raw_missing = set(expected_cells) - set(actual_cells)
            raw_extra = set(actual_cells) - set(expected_cells)
            raw_changed = [identity for identity in set(expected_cells) & set(actual_cells) if _json_bytes(expected_cells[identity]) != _json_bytes(actual_cells[identity])]
            if raw_missing or raw_extra or raw_changed or _json_bytes(candidate_facts.get("date_system")) != _json_bytes(fresh_scan["date_system"]):
                hard_blockers.append("raw source facts do not match the current workbook")
                diagnostics.append(
                    _diagnostic(
                        "raw_source_facts_mismatch",
                        "error",
                        f"Raw source facts differ (missing={len(raw_missing)}, extra={len(raw_extra)}, changed={len(raw_changed)}).",
                    )
                )

        fresh_ids = _inventory_ids(fresh_inventory)
        actual_ids = _inventory_ids(candidate_inventory) if candidate_inventory else set()
        missing_ids = fresh_ids - actual_ids
        extra_ids = actual_ids - fresh_ids
        logical_map = {obj.get("identity"): obj for obj in (candidate_logical or [])}
        logical_ids = {identity for identity in logical_map if identity}
        missing_ledger = set(expected_objects) - logical_ids
        extra_ledger = logical_ids - set(expected_objects)
        altered_ledger = [
            identity
            for identity in set(expected_objects) & logical_ids
            if _json_bytes({k: v for k, v in logical_map[identity].items() if k not in {"accounted", "opaque"}})
            != _json_bytes(expected_objects[identity])
        ]
        if missing_ids or extra_ids:
            hard_blockers.append("supported logical objects are missing or extra in inventory.json")
            diagnostics.append(
                _diagnostic(
                    "logical_objects_mismatch",
                    "error",
                    f"Logical object accounting differs (missing output={len(missing_ids)}, extra output={len(extra_ids)}).",
                )
            )
        if missing_ledger or extra_ledger or altered_ledger:
            hard_blockers.append("logical object ledger differs from fresh source declarations")
            diagnostics.append(_diagnostic("logical_ledger_mismatch", "error", f"Logical ledger differs (missing={len(missing_ledger)}, extra={len(extra_ledger)}, changed={len(altered_ledger)})."))
        expected_accounting = {identity: identity in actual_ids for identity in expected_objects}
        bad_ledger_flags = [
            identity
            for identity, obj in logical_map.items()
            if identity in expected_objects
            and (obj.get("opaque") is not False or obj.get("accounted") is not expected_accounting[identity])
        ]
        if bad_ledger_flags:
            hard_blockers.append("logical objects cannot be relabeled opaque or assigned stale accounting flags")
            diagnostics.append(_diagnostic("logical_accounting_flags_mismatch", "error", f"Logical ledger has {len(bad_ledger_flags)} altered opaque/accounted flag(s)."))

        # Compare normalized projections against direct raw transforms for cells;
        # then compare remaining records against the freshly sourced object projection.
        actual_records = _inventory_record_map(candidate_inventory) if candidate_inventory else {}
        expected_records = _inventory_record_map(fresh_inventory)
        for identity, expected in expected_records.items():
            actual = actual_records.get(identity)
            if actual is None:
                fidelity_deviations.append(
                    {
                        "identity": identity,
                        "source_location": expected.source_location.model_dump(mode="json"),
                        "fields": {"record": {"expected": expected.model_dump(mode="json"), "actual": None}},
                    }
                )
                continue
            expected_dump = expected.model_dump(mode="json")
            actual_dump = actual.model_dump(mode="json")
            if isinstance(expected, SheetInventory) and isinstance(actual, SheetInventory):
                for key in ("cells", "ranges", "layout_objects"):
                    expected_dump.pop(key, None)
                    actual_dump.pop(key, None)
            if _json_bytes(expected_dump) != _json_bytes(actual_dump):
                expected_cell = expected_cells.get(identity)
                if expected_cell:
                    keys = (
                        "source_location",
                        "source_identity",
                        "address",
                        "row",
                        "column",
                        "kind",
                        "value",
                        "formula",
                        "cached_value",
                        "cached_value_available",
                        "raw_value_text",
                        "raw_formula_text",
                        "raw_formula_attributes",
                        "formula_present",
                        "ooxml_cell_type",
                        "cached_text",
                        "cached_text_present",
                        "workbook_date_system",
                        "date_serial_text",
                        "style_id",
                        "number_format",
                        "data_type",
                    )
                    differences = {
                        key: {"expected": getattr(expected, key), "actual": getattr(actual, key)}
                        for key in keys
                        if _json_bytes(expected_dump.get(key)) != _json_bytes(actual_dump.get(key))
                    }
                    if not differences:
                        differences["record"] = {"expected": expected.model_dump(mode="json"), "actual": actual.model_dump(mode="json")}
                    if differences:
                        fidelity_deviations.append({"identity": identity, "source_location": expected.source_location.model_dump(mode="json"), "fields": differences})
                else:
                    fidelity_deviations.append(
                        {
                            "identity": identity,
                            "source_location": expected.source_location.model_dump(mode="json"),
                            "fields": {"record": {"expected": expected.model_dump(mode="json"), "actual": actual.model_dump(mode="json")}},
                        }
                    )

        preserved_count = 0
        if candidate_parts is not None:
            actual_part_map = {part.get("name"): part for part in candidate_parts}
            expected_names = set(expected_parts)
            actual_names = set(actual_part_map)
            bad_parts: list[str] = []
            for name in expected_names:
                part = actual_part_map.get(name)
                if part is None:
                    bad_parts.append(name)
                    continue
                expected_part = expected_parts[name]
                part_identity = object_identity("package_part", name)
                if (
                    part.get("sha256") != expected_part["sha256"]
                    or _json_bytes(part.get("size")) != _json_bytes(expected_part["size"])
                    or part.get("content_type") != expected_part["content_type"]
                    or part.get("source_identity") != part_identity
                    or _json_bytes(part.get("source_location"))
                    != _json_bytes({"ooxml_part": name, "object_type": "package_part", "object_id": name})
                    or part.get("opaque") is not expected_part["opaque"]
                    or part.get("opaque_reason") != expected_part["opaque_reason"]
                ):
                    bad_parts.append(name)
                    continue
                preserved = run_dir / str(part.get("preserved_path", ""))
                try:
                    preserved.resolve().relative_to(run_dir.resolve())
                    if not preserved.is_file() or _hash_file(preserved) != expected_parts[name]["sha256"]:
                        bad_parts.append(name)
                    else:
                        preserved_count += 1
                except (OSError, ValueError):
                    bad_parts.append(name)
            if expected_names != actual_names or bad_parts:
                hard_blockers.append("package part accounting or preserved bytes differ")
                diagnostics.append(_diagnostic("package_parts_mismatch", "error", f"Package parts missing/extra/changed: {len(expected_names ^ actual_names) + len(bad_parts)}."))

        source_evidence = run_dir / "source_evidence" / Path(source.get("source_path", "source.xlsx")).name
        if not source_evidence.is_file() or _hash_file(source_evidence) != source.get("sha256"):
            hard_blockers.append("original source evidence copy is missing or changed")
            diagnostics.append(_diagnostic("source_evidence_mismatch", "error", "Original source evidence does not match the source hash."))

        metrics.update(
            {
                "logical_objects_total": len(expected_objects),
                "logical_objects_accounted": len(expected_objects) - len(missing_ids),
                "traceability_ratio": (len(expected_objects) - len(missing_ids)) / len(expected_objects) if expected_objects else 0.0,
                "parsed_objects_total": len(expected_objects),
                "parsed_objects": len(expected_objects) - len(missing_ids),
                "parsed_coverage_ratio": (len(expected_objects) - len(missing_ids)) / len(expected_objects) if expected_objects else 0.0,
                "package_parts_total": len(expected_parts),
                "parsed_package_parts": sum(name in parsed_source_parts for name in expected_parts),
                "parsed_package_parts_total": len(expected_parts),
                "parsed_package_parts_ratio": sum(name in parsed_source_parts for name in expected_parts) / len(expected_parts) if expected_parts else 0.0,
                "package_parts_preserved": preserved_count if candidate_parts is not None else 0,
                "package_preservation_ratio": preserved_count / len(expected_parts) if expected_parts else 0.0,
                "opaque_parts": sum(1 for name, info in expected_parts.items() if _opaque(name, info)),
                "opaque_rate": sum(1 for name, info in expected_parts.items() if _opaque(name, info)) / len(expected_parts) if expected_parts else 0.0,
                "supported_facts_total": len(expected_records),
                "supported_facts_exact": len(expected_records) - len(fidelity_deviations),
                "supported_fidelity_ratio": (len(expected_records) - len(fidelity_deviations)) / len(expected_records) if expected_records else 0.0,
                "fidelity_deviations": fidelity_deviations,
            }
        )
        metrics_state = "measured"
    else:
        metrics.update(_unavailable_source_metrics())

    if metrics_state == "measured" and metrics.get("logical_objects_total") == 0:
        hard_blockers.append("source contains no readable supported logical objects")
        diagnostics.append(_diagnostic("empty_or_unreadable_corpus", "error", "No supported logical objects were found; 0/0 is not a passing result."))
    if metrics_state == "measured" and metrics.get("package_parts_total") == 0:
        hard_blockers.append("source contains no readable package parts")
        diagnostics.append(_diagnostic("empty_or_unreadable_package", "error", "No non-directory OOXML package parts were found; 0/0 is not a passing result."))

    stored_thresholds = source.get("thresholds", {})
    thresholds = {
        "traceability": stored_thresholds.get("traceability", 1.0),
        "fidelity": stored_thresholds.get("fidelity", 1.0),
        "parsed_package_min": stored_thresholds.get("parsed_package_min", 0.0),
        "opaque_max": stored_thresholds.get("opaque_max", 1.0),
    }
    trace_min = float(thresholds["traceability"])
    fidelity_min = float(thresholds["fidelity"])
    parsed_package_min = float(thresholds["parsed_package_min"])
    opaque_max = float(thresholds["opaque_max"])
    trace_ok = metrics["traceability_ratio"] >= trace_min if metrics_state == "measured" else None
    fidelity_ok = metrics["supported_fidelity_ratio"] >= fidelity_min if metrics_state == "measured" else None
    parsed_package_ok = metrics["parsed_package_parts_ratio"] >= parsed_package_min if metrics_state == "measured" else None
    opaque_rate_ok = metrics["opaque_rate"] <= opaque_max if metrics_state == "measured" else None
    if fidelity_deviations:
        diagnostics.append(
            _diagnostic(
                "normalized_fidelity_mismatch",
                "error" if not fidelity_ok else "warning",
                f"{len(fidelity_deviations)} supported normalized fact(s) differ from the current source transformation.",
                fidelity_deviations[0].get("source_location"),
            )
        )
    if trace_ok is False:
        hard_blockers.append("traceability ratio is below its configured threshold")
    if fidelity_ok is False:
        hard_blockers.append("supported-fact fidelity ratio is below its configured threshold")
    if parsed_package_ok is False:
        hard_blockers.append("parsed package-part ratio is below its configured threshold")
    if opaque_rate_ok is False:
        hard_blockers.append("opaque package-part rate is above its configured maximum")

    threshold_errors = [
        _diagnostic("quality_threshold_not_met", "error", text)
        for text, ok in (
            ("traceability", trace_ok),
            ("fidelity", fidelity_ok),
        )
        if ok is False
    ]
    if parsed_package_ok is False:
        threshold_errors.append(
            _diagnostic(
                "parsed_package_threshold_not_met",
                "error",
                f"Parsed package parts {metrics.get('parsed_package_parts', 0)}/{metrics.get('package_parts_total', 0)} are below the configured minimum {parsed_package_min:.1%}.",
            )
        )
    if opaque_rate_ok is False:
        threshold_errors.append(
            _diagnostic(
                "opaque_rate_threshold_not_met",
                "error",
                f"Opaque package part rate {metrics.get('opaque_rate', 0):.1%} exceeds the configured maximum {opaque_max:.1%}.",
            )
        )
    diagnostics.extend(threshold_errors)
    opaque_allowed = bool(source.get("allow_opaque", True))
    if metrics.get("opaque_parts", 0) and not opaque_allowed:
        hard_blockers.append("opaque package parts are not allowed by this run's policy")
        diagnostics.append(_diagnostic("opaque_parts_blocked", "error", f"{metrics['opaque_parts']} opaque package part(s) require specialist extraction or explicit acceptance."))
    elif metrics.get("opaque_parts", 0):
        diagnostics.append(_diagnostic("opaque_parts_preserved", "warning", f"{metrics['opaque_parts']} opaque package part(s) were preserved byte-for-byte and remain unparsed."))

    blockers = list(dict.fromkeys(hard_blockers))
    if blockers:
        status = "fail"
        ready = False
    elif metrics.get("opaque_parts", 0) or fidelity_deviations:
        status = "partial"
        ready = True
    else:
        status = "pass"
        ready = True
    next_tool = None
    if not ready:
        fact_repair_codes = {"logical_ledger_mismatch", "logical_ledger_unreadable", "logical_ledger_invalid", "duplicate_logical_object_identity", "raw_source_facts_mismatch", "source_facts_unreadable", "source_facts_invalid", "duplicate_source_fact_identity"}
        inventory_repair_codes = {
            "normalized_fidelity_mismatch",
            "logical_objects_mismatch",
            "inventory_unreadable",
            "duplicate_inventory_identity",
            "inventory_parent_mismatch",
            "inventory_projection_mismatch",
        }
        package_repair_codes = {"package_parts_mismatch", "package_ledger_unreadable", "package_ledger_invalid", "duplicate_package_part_name"}
        retryable_codes = fact_repair_codes | inventory_repair_codes | package_repair_codes | {"manifest_unreadable", "manifest_mismatch"}
        code = next((d["code"] for d in diagnostics if d["code"] in retryable_codes), None)
        if code in package_repair_codes:
            next_tool = {"name": "package.preserve", "arguments": {"run": str(run_dir)}}
        elif code in fact_repair_codes:
            next_tool = {"name": "source_facts.refresh", "arguments": {"run": str(run_dir)}}
        elif code in inventory_repair_codes:
            next_tool = {"name": "inventory.extract", "arguments": {"run": str(run_dir)}}
        elif code in {"manifest_unreadable", "manifest_mismatch"}:
            next_tool = {"name": "manifest.read", "arguments": {"run": str(run_dir)}}

    recovery_stopped = False
    stop_reason = None
    remaining_tool = None
    if attempt_history_error is not None:
        recovery_stopped = True
        stop_reason = attempt_history_error.code
        remaining_tool = next_tool
        next_tool = None
    elif status == "fail" and next_tool is not None and attempt_history is not None:
        attempts = attempt_history["attempts"]
        attempt_limit = attempt_history["attempt_limit"]
        if attempt_history.get("no_progress_stopped", False):
            stop_reason = "no_progress_stop"
        elif len(attempts) >= attempt_limit:
            stop_reason = "attempt_limit_reached"
        if stop_reason:
            recovery_stopped = True
            remaining_tool = next_tool
            next_tool = None
            diagnostics.append(
                _diagnostic(
                    stop_reason,
                    "error",
                    "Two consecutive corrective attempts did not improve the candidate; stop recovery."
                    if stop_reason == "no_progress_stop"
                    else "The configured corrective-attempt limit is exhausted; inspect the remaining suggestion manually or use a specialist."
                )
            )

    result = {
        "schema_version": "step1.v1",
        "source_sha256": source.get("sha256"),
        "run_id": source.get("run_id"),
        "coverage_scope": _coverage_scope(),
        "status": status,
        "ready_for_next_step": ready,
        "next_step": "step2_index" if ready else None,
        "thresholds": thresholds,
        "metrics_state": metrics_state,
        "metrics": metrics,
        "blockers": blockers,
        "diagnostics": diagnostics,
        "next_tool": next_tool,
        "retryable": next_tool is not None,
        "recovery_stopped": recovery_stopped,
        "stop_reason": stop_reason,
        "remaining_tool": remaining_tool,
        "checked_at": datetime.now(UTC).isoformat(),
    }
    if write and (run_dir / "source.json").exists():
        _write_json(run_dir / "quality.json", result)
        _write_handoff(run_dir, result, final_path=_valid_final_link(run_dir, result))
    return result


def _write_handoff(
    run_dir: Path,
    quality: dict[str, Any],
    *,
    final_path: str | None = None,
    output_root_override: Path | None = None,
) -> None:
    source = _safe_run_source(run_dir)
    refs = _artifact_refs(run_dir)
    try:
        output_root = output_root_override or run_dir.parents[5]
        run_path = _relative(run_dir, output_root)
    except (IndexError, ValueError):
        output_root = run_dir.parent
        run_path = run_dir.name
    next_actions = (
        [{**quality["next_tool"], "arguments": dict(quality["next_tool"].get("arguments", {}))}]
        if quality.get("next_tool")
        else ([] if quality.get("ready_for_next_step") else [{"name": "source.review", "arguments": {"run": str(run_dir)}}])
    )
    for action in next_actions:
        if action.get("arguments", {}).get("run") == str(run_dir):
            action["arguments"]["run"] = run_path
    final = {
        "schema_version": "step1.v1",
        "step": "step1_conversion",
        "coverage_scope": _coverage_scope(),
        "run_path": run_path,
        "artifact_paths_relative_to": run_path,
        "next_step": quality.get("next_step"),
        "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
        "run_id": source.get("run_id"),
        "batch_id": source.get("batch_id"),
        "status": quality.get("status"),
        "ready_for_next_step": quality.get("ready_for_next_step", False),
        "metrics_state": quality.get("metrics_state", "unavailable"),
        "metrics": quality.get("metrics", {}),
        "thresholds": quality.get("thresholds", {}),
        "artifacts": refs,
        "final_output": final_path,
        "recovery_stopped": quality.get("recovery_stopped", False),
        "stop_reason": quality.get("stop_reason"),
        "remaining_tool": quality.get("remaining_tool"),
        "blockers": quality.get("blockers", []),
        "diagnostics": quality.get("diagnostics", []),
        "next_actions": next_actions,
    }
    _write_json(run_dir / "handoff.json", final)
    metrics = quality.get("metrics", {})

    def display_metric(value: Any, spec: str = "") -> str:
        return "unavailable" if value is None else format(value, spec)

    lines = [
        "# Step1 conversion handoff",
        "",
        f"**Input:** `{source.get('relative_path')}` · **SHA-256:** `{source.get('sha256')}`",
        f"**Run:** `{source.get('run_id')}` · **Status:** **{quality.get('status')}** · **Ready for Step2:** **{quality.get('ready_for_next_step')}**",
        "",
        "## Coverage scope",
        "",
        f"- Logical object kinds: {', '.join(_coverage_scope()['logical_object_kinds'])}.",
        f"- Cell unit: {_coverage_scope()['cell_unit']}",
        f"- Package-part unit: {_coverage_scope()['package_part_unit']}",
        f"- Logical parsed ratio: {_coverage_scope()['parsed_logical_object_ratio']}",
        f"- Package parsed ratio: {_coverage_scope()['parsed_package_parts_ratio']}",
        f"- Opaque rate: {_coverage_scope()['opaque_package_parts_ratio']}",
        f"- Opaque policy: {_coverage_scope()['opaque_policy']}",
        f"- Scope limit: {_coverage_scope()['scope_limit']}",
        "",
        "## Quality metrics",
        "",
        f"- Current source measurements: {quality.get('metrics_state', 'unavailable')}.",
        f"- Traceable logical objects: {display_metric(metrics.get('logical_objects_accounted'))}/{display_metric(metrics.get('logical_objects_total'))} ({display_metric(metrics.get('traceability_ratio'), '.1%')})",
        f"- Preserved package parts: {display_metric(metrics.get('package_parts_preserved'))}/{display_metric(metrics.get('package_parts_total'))} ({display_metric(metrics.get('package_preservation_ratio'), '.1%')})",
        f"- Parsed package parts: {display_metric(metrics.get('parsed_package_parts'))}/{display_metric(metrics.get('package_parts_total'))} ({display_metric(metrics.get('parsed_package_parts_ratio'), '.1%')}); minimum {quality.get('thresholds', {}).get('parsed_package_min', 0):.1%}",
        f"- Opaque package parts: {display_metric(metrics.get('opaque_parts'))}/{display_metric(metrics.get('package_parts_total'))} ({display_metric(metrics.get('opaque_rate'), '.1%')}); maximum {quality.get('thresholds', {}).get('opaque_max', 1):.1%}",
        f"- Exact supported-fact fidelity: {display_metric(metrics.get('supported_facts_exact'))}/{display_metric(metrics.get('supported_facts_total'))} ({display_metric(metrics.get('supported_fidelity_ratio'), '.1%')}); minimum {quality.get('thresholds', {}).get('fidelity', 1):.1%}",
        f"- Traceability minimum: {quality.get('thresholds', {}).get('traceability', 1):.1%}; package preservation is always required at 100%.",
        f"- Parsed logical-object ratio: {display_metric(metrics.get('parsed_coverage_ratio'), '.1%')} (separate from package-part ratios).",
        f"- Opaque policy enabled: {source.get('allow_opaque', True)}; all package bytes still require exact preservation.",
        "",
    ]
    if quality.get("blockers"):
        lines.extend(["## Blockers", "", *[f"- {item}" for item in quality["blockers"]], ""])
    if quality.get("recovery_stopped"):
        pending = quality.get("remaining_tool") or {}
        lines.extend(["## Recovery stopped", "", f"- Reason: `{quality.get('stop_reason')}`.", f"- Remaining suggested tool: `{pending.get('name')}` (do not retry automatically).", ""])
    if quality.get("metrics", {}).get("opaque_parts"):
        lines.extend(["## Opaque parts", "", f"{quality['metrics']['opaque_parts']} original package part(s) were preserved byte-for-byte and were not parsed.", ""])
    if final_path:
        lines.extend(["## Final converted output", "", f"`{final_path}`", ""])
    lines.extend(["## Candidate artifacts", "", *[f"- `{ref['path']}` · SHA-256 `{ref['sha256']}`" for ref in refs], ""])
    if final["next_actions"]:
        lines.extend(["## Next action", "", *[f"- `{action['name']}` with `{json.dumps(action['arguments'], ensure_ascii=False)}`" for action in final["next_actions"]], ""])
    (run_dir / "handoff.md").write_text("\n".join(lines), encoding="utf-8")


def _record_source_error(run_dir: Path, metadata: dict[str, Any], message: str, code: str) -> dict[str, Any]:
    run_dir.mkdir(parents=True, exist_ok=True)
    _write_json(run_dir / "source.json", metadata)
    _write_json(run_dir / "attempt_history.json", {"attempt_limit": 3, "attempts": []})
    result = {
        "schema_version": "step1.v1",
        "source_sha256": metadata.get("sha256"),
        "run_id": metadata.get("run_id"),
        "coverage_scope": _coverage_scope(),
        "status": "fail",
        "ready_for_next_step": False,
        "next_step": None,
        "thresholds": metadata.get("thresholds", {}),
        "metrics_state": "unavailable",
        "metrics": _unavailable_source_metrics(),
        "blockers": [message],
        "diagnostics": [_diagnostic(code, "error", message)],
        "next_tool": None,
        "retryable": False,
        "checked_at": datetime.now(UTC).isoformat(),
    }
    _write_json(run_dir / "quality.json", result)
    _write_handoff(run_dir, result)
    return result


def _batch_handoff(batch_dir: Path, entries: list[dict[str, Any]], discovery: dict[str, Any] | None = None) -> dict[str, Any]:
    failed = [entry for entry in entries if entry.get("status") == "fail"]
    partial = [entry for entry in entries if entry.get("status") == "partial"]
    status = "fail" if failed else "partial" if partial else "pass"
    result = {
        "schema_version": "step1.batch.v1",
        "batch_id": batch_dir.name,
        "coverage_scope": _coverage_scope(),
        "status": status,
        "input_count": len(entries),
        "passed": sum(entry.get("status") == "pass" for entry in entries),
        "partial": len(partial),
        "failed": len(failed),
        "discovery": discovery or {},
        "entries": entries,
    }
    _write_json(batch_dir / "batch_handoff.json", result)
    lines = [f"# Step1 batch handoff · {batch_dir.name}", "", f"**Status:** {status} · **Inputs:** {len(entries)} · **Pass:** {result['passed']} · **Partial:** {len(partial)} · **Fail:** {len(failed)}", "", "| Input | SHA-256 | Status | Metrics | Candidate/handoff |", "| --- | --- | --- | --- | --- |"]
    for entry in entries:
        lines.append(f"| `{entry.get('relative_path')}` | `{entry.get('sha256')}` | {entry.get('status')} | {entry.get('metrics_state', 'unavailable')} | `{entry.get('handoff_path', '')}` |")
    (batch_dir / "batch_handoff.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    _write_json(batch_dir / "batch.json", {"batch_id": batch_dir.name, "discovery": discovery or {}, "entries": entries})
    return result


def convert_directory(
    input_dir: Path,
    out_dir: Path,
    *,
    traceability_min: float = 1.0,
    fidelity_min: float = 1.0,
    parsed_package_min: float = 0.0,
    opaque_max: float = 1.0,
    allow_opaque: bool = True,
) -> dict[str, Any]:
    try:
        traceability_min = _ratio(traceability_min, "traceability_min")
        fidelity_min = _ratio(fidelity_min, "fidelity_min")
        parsed_package_min = _ratio(parsed_package_min, "parsed_package_min")
        opaque_max = _ratio(opaque_max, "opaque_max")
        if not isinstance(allow_opaque, bool):
            raise ValueError("allow_opaque must be a boolean")
    except ValueError as exc:
        return {
            "tool": "step1.convert",
            "status": "error",
            "source": {"input_directory": str(input_dir)},
            "run_id": None,
            "artifacts": [],
            "metrics": {"input_count": 0, "pass": 0, "partial": 0, "fail": 0},
            "diagnostics": [_diagnostic("invalid_step1_options", "error", str(exc))],
            "entries": [],
            "retryable": False,
            "next_tool": None,
        }
    input_dir = input_dir.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    if input_dir == out_dir:
        return {
            "tool": "step1.convert",
            "status": "fail",
            "source": {"input_directory": str(input_dir)},
            "run_id": None,
            "artifacts": [],
            "metrics": {"input_count": 0, "pass": 0, "partial": 0, "fail": 1, "output_subtree_excluded": False},
            "diagnostics": [_diagnostic("output_is_input", "error", "--out cannot be the raw input directory; choose a child or sibling output directory.")],
            "entries": [],
            "retryable": False,
            "next_tool": None,
        }
    output_subtree_excluded = input_dir in out_dir.parents
    batch_id = f"{datetime.now(UTC):%Y%m%dT%H%M%S}-{uuid4().hex[:8]}"
    batch_dir = out_dir / "batches" / batch_id
    batch_dir.mkdir(parents=True, exist_ok=False)
    entries: list[dict[str, Any]] = []
    threshold_values = {
        "traceability": traceability_min,
        "fidelity": fidelity_min,
        "parsed_package_min": parsed_package_min,
        "opaque_max": opaque_max,
    }
    discovered = discover_inputs(input_dir, out_dir)
    if not discovered:
        entry = {"relative_path": "", "sha256": None, "status": "fail", "metrics_state": "unavailable", "metrics": _unavailable_source_metrics(), "diagnostics": [_diagnostic("empty_input_directory", "error", "No files were found in the input directory.")]}
        entries.append(entry)
    for source_path, relative_path, suffix in discovered:
        try:
            digest = _hash_file(source_path)
            hash_error = None
        except OSError as exc:
            digest = None
            hash_error = exc
        entry_key = hashlib.sha256(relative_path.encode("utf-8")).hexdigest()[:12]
        run_id = f"{datetime.now(UTC):%Y%m%dT%H%M%S}-{uuid4().hex[:8]}"
        run_dir = batch_dir / "sources" / (digest or "unreadable") / entry_key / run_id
        source_meta = {
            "schema_version": "step1.v1",
            "batch_id": batch_id,
            "run_id": run_id,
            "source_path": str(source_path),
            "relative_path": relative_path,
            "sha256": digest,
            "suffix": suffix,
            "thresholds": threshold_values,
            "allow_opaque": allow_opaque,
            "attempt_limit": 3,
        }
        if hash_error is not None:
            reason = f"Cannot read input file for hashing: {type(hash_error).__name__}: {hash_error}"
            result = _record_source_error(run_dir, source_meta, reason, "source_hash_failed")
        elif suffix not in _SUPPORTED:
            reason = f"Unsupported input type {suffix or '(no extension)'}; Step1 accepts .xlsx and .xlsm."
            result = _record_source_error(run_dir, source_meta, reason, "unsupported_input_type")
        else:
            try:
                scan = scan_step1_source(source_path)
                run_dir.mkdir(parents=True, exist_ok=False)
                _write_json(run_dir / "source.json", source_meta)
                _write_base(source_path, run_dir, scan)
                _write_json(run_dir / "attempt_history.json", {"attempt_limit": 3, "attempts": []})
                result = _check_run(run_dir)
            except (OSError, ValueError, zipfile.BadZipFile, SourceScanError, Exception) as exc:
                reason = f"Source read or conversion failed: {type(exc).__name__}: {exc}"
                result = _record_source_error(run_dir, source_meta, reason, "source_read_failed")
        entries.append(
            {
                "relative_path": relative_path,
                "sha256": digest,
                "run_id": run_id,
                "run_path": _relative(run_dir, out_dir),
                "handoff_path": _relative(run_dir / "handoff.json", out_dir),
                "status": result["status"],
                "metrics_state": result.get("metrics_state", "unavailable"),
                "metrics": result["metrics"],
                "diagnostics": result["diagnostics"],
            }
        )
    discovery = {
        "input_directory": str(input_dir),
        "recursive": True,
        "ordering": "relative_path_casefold_then_original",
        "output_subtree_excluded": output_subtree_excluded,
        "excluded_output_path": out_dir.relative_to(input_dir).as_posix() if output_subtree_excluded else None,
    }
    batch = _batch_handoff(batch_dir, entries, discovery)
    return {
        "tool": "step1.convert",
        "status": batch["status"],
        "coverage_scope": _coverage_scope(),
        "source": {"input_directory": str(input_dir), "entry_count": len(entries), "output_subtree_excluded": output_subtree_excluded},
        "run_id": batch_id,
        "artifacts": [{"name": "batch_handoff.json", "path": _relative(batch_dir / "batch_handoff.json", out_dir), "sha256": _hash_file(batch_dir / "batch_handoff.json")}, {"name": "batch_handoff.md", "path": _relative(batch_dir / "batch_handoff.md", out_dir), "sha256": _hash_file(batch_dir / "batch_handoff.md")}],
        "metrics": {"input_count": len(entries), "pass": batch["passed"], "partial": batch["partial"], "fail": batch["failed"], "output_subtree_excluded": output_subtree_excluded},
        "diagnostics": [diagnostic for entry in entries for diagnostic in entry.get("diagnostics", [])],
        "entries": entries,
        "retryable": any(entry.get("status") in {"pass", "partial"} for entry in entries),
        "next_tool": next(
            (
                {"name": "step1.check", "arguments": {"run": str(out_dir / entry["run_path"])}}
                for entry in entries
                if "run_path" in entry and entry.get("status") in {"pass", "partial"}
            ),
            None,
        ),
    }


_TOOLS: list[dict[str, Any]] = [
    {"name": "step1.convert", "command": "step1 convert INPUT_DIR --out OUTPUT_DIR [quality options]", "inputs": {"input_dir": "recursive directory", "out": "output root", "traceability_min": "number in [0,1], default 1; denominator is supported logical objects", "fidelity_min": "number in [0,1], default 1; denominator is supported inventory facts", "parsed_package_min": "number in [0,1], default 0; denominator is all non-directory OOXML parts", "opaque_max": "number in [0,1], default 1; denominator is all non-directory OOXML parts", "allow_opaque": "boolean, default true; false blocks every opaque part"}, "outputs": ["per-source candidate files", "package part copies", "separate logical and package ratios", "batch_handoff.json", "batch_handoff.md"], "next": ["step1.check", "step1.finalize"]},
    {"name": "step1.check", "command": "step1 check --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["fresh quality.json", "handoff.json", "handoff.md", "metrics", "diagnostics"], "next": ["inventory.extract", "source_facts.refresh", "package.preserve", "step1.finalize"]},
    {"name": "step1.coverage", "command": "step1 coverage --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["logical object and package part ledgers", "logical parsed_coverage_ratio", "package parsed_package_parts_ratio and opaque_rate", "separate denominators and declared coverage_scope"], "next": ["inventory.extract", "package.preserve"]},
    {"name": "step1.fidelity", "command": "step1 fidelity --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["raw source facts", "normalized projection comparisons", "source locations"], "next": ["inventory.extract", "source_facts.refresh"]},
    {"name": "manifest.read", "command": "step1 tool manifest.read --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted workbook_manifest.json"], "next": ["inventory.extract", "step1.check"]},
    {"name": "inventory.extract", "command": "step1 tool inventory.extract --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted inventory.json with normalized values and raw OOXML fields"], "next": ["step1.check"]},
    {"name": "cached_values.read", "command": "step1 tool cached_values.read --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted cached values and availability flags in inventory.json"], "next": ["step1.fidelity"]},
    {"name": "data_tables.read", "command": "step1 tool data_tables.read --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted data-table ranges and source formula attributes"], "next": ["step1.coverage"]},
    {"name": "form_controls.read", "command": "step1 tool form_controls.read --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted legacy form-control facts and links"], "next": ["step1.coverage"]},
    {"name": "vba.extract", "command": "step1 tool vba.extract --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted VBA modules or explicit opaque/skipped diagnostics"], "next": ["step1.check", "graph.build"]},
    {"name": "vba.cell_links", "command": "step1 tool vba.cell_links --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted VBA-to-cell graph edges"], "next": ["graph.build"]},
    {"name": "graph.build", "command": "step1 tool graph.build --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted dependency_graph.json"], "next": ["step1.check"]},
    {"name": "classify.rules", "command": "step1 tool classify.rules --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted module_classification.json (later-stage heuristic aid)"], "next": ["confirmation.build"]},
    {"name": "confirmation.build", "command": "step1 tool confirmation.build --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted confirmation_template.json (later-stage questions)"], "next": ["step1.finalize"]},
    {"name": "source_facts.refresh", "command": "step1 tool source_facts.refresh --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted fresh source_facts.json and logical_objects.json"], "next": ["step1.check"]},
    {"name": "package.preserve", "command": "step1 tool package.preserve --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["persisted package_parts.json and exact part bytes"], "next": ["step1.coverage"]},
    {"name": "storage.validate", "command": "step1 tool storage.validate --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["candidate artifact parse and checksum summary"], "next": ["step1.check"]},
    {"name": "report.handoff", "command": "step1 tool report.handoff --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["fresh human and machine handoff"], "next": ["step1.finalize"]},
    {"name": "step1.auto_recover", "command": "step1 auto-recover --run RUN_DIR --max-attempts 3", "inputs": {"run": "per-source run directory", "max_attempts": "integer from 1 to 3"}, "outputs": ["bounded persisted recovery history and fresh checks"], "next": ["step1.check", "step1.finalize"]},
    {"name": "step1.finalize", "command": "step1 finalize --run RUN_DIR", "inputs": {"run": "per-source run directory"}, "outputs": ["fresh quality report and unique promoted final folder when accepted"], "next": ["step2_index"]},
]


def tool_catalog() -> dict[str, Any]:
    return {
        "tool": "step1.tools",
        "status": "ok",
        "source": None,
        "run_id": None,
        "coverage_scope": _coverage_scope(),
        "metrics_contract": {
            "state_field": "metrics_state",
            "states": {"measured": "A fresh source scan supplied the current denominators and ratios.", "unavailable": "A fresh source scan could not be validated; source-derived totals, counts, ratios, and deviations are null."},
            "zero_denominator": "Numeric zero is used only when a fresh scan measured an empty denominator; empty corpora still fail.",
            "gate": "Unavailable metrics never satisfy thresholds, including thresholds set to zero; source/read blockers prevent finalization.",
        },
        "artifacts": [],
        "metrics": {"tool_count": len(_TOOLS)},
        "diagnostics": [],
        "retryable": False,
        "next_tool": None,
        "tools": _TOOLS,
    }


def _load_run_inventory(run_dir: Path) -> WorkbookInventory:
    return WorkbookInventory.model_validate(_read_json(run_dir / "inventory.json"))


def _refresh_from_source(run_dir: Path, component: str) -> dict[str, Any]:
    source = _run_source(run_dir)
    path = Path(source["source_path"])
    if not path.is_file() or _hash_file(path) != source["sha256"]:
        raise ValueError("source is missing or changed; candidate refresh is unsafe")
    scan = scan_step1_source(path)
    manifest, inventory = _build_conversion(path, scan)
    if component in {"inventory.extract", "cached_values.read", "data_tables.read", "form_controls.read", "manifest.read"}:
        _write_json(run_dir / "workbook_manifest.json", manifest.model_dump(mode="json"))
        _write_json(run_dir / "inventory.json", inventory.model_dump(mode="json"))
    if component == "source_facts.refresh":
        _write_json(run_dir / "source_facts.json", {"date_system": scan["date_system"], "cells": scan["cells"]})
        ids = _inventory_ids(inventory)
        _write_json(run_dir / "logical_objects.json", [{**item, "accounted": item["identity"] in ids, "opaque": False} for item in scan["objects"]])
    if component == "package.preserve":
        parts = _preserve_parts(path, run_dir, scan)
        _write_json(run_dir / "package_parts.json", parts)
        evidence = run_dir / "source_evidence" / path.name
        evidence.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, evidence)
    return {"source": {"path": source["relative_path"], "sha256": source["sha256"]}, "metrics": {"objects": len(scan["objects"]), "cells": len(scan["cells"]), "parts": len(scan["parts"])}, "artifacts": [component]}


def _tool_action(name: str, run_dir: Path) -> dict[str, Any]:
    if name in {"inventory.extract", "cached_values.read", "data_tables.read", "form_controls.read", "manifest.read", "source_facts.refresh", "package.preserve"}:
        return _refresh_from_source(run_dir, name)
    if name == "storage.validate":
        checked = []
        source = _run_source(run_dir)
        manifest = WorkbookManifest.model_validate(_read_json(run_dir / "workbook_manifest.json"))
        inventory = WorkbookInventory.model_validate(_read_json(run_dir / "inventory.json"))
        if manifest.sha256 != source.get("sha256") or inventory.workbook_sha256 != source.get("sha256"):
            raise ValueError("manifest or inventory source hash does not match source.json")
        for artifact in _ARTIFACT_FILES:
            path = run_dir / artifact
            if not path.is_file():
                raise ValueError(f"candidate artifact is missing: {artifact}")
            _read_json(path)
            checked.append({"name": artifact, "sha256": _hash_file(path)})
        return {"artifacts": checked, "metrics": {"validated_artifacts": len(checked), "inventory_records": len(_inventory_records(inventory))}}
    if name == "report.handoff":
        quality = _check_run(run_dir)
        return {"artifacts": ["quality.json", "handoff.json", "handoff.md"], "metrics_state": quality["metrics_state"], "metrics": quality["metrics"], "status": quality["status"]}
    if name == "graph.build":
        graph = RegexFormulaGraphBuilder().build(_load_run_inventory(run_dir))
        _write_json(run_dir / "dependency_graph.json", graph.model_dump(mode="json"))
        return {"artifacts": ["dependency_graph.json"], "metrics": {"nodes": len(graph.nodes), "edges": len(graph.edges)}}
    if name == "vba.extract":
        from excel_to_act.ingest.vba import extract_vba_project

        inventory = _load_run_inventory(run_dir)
        project = extract_vba_project(Path(_run_source(run_dir)["source_path"]))
        inventory.vba_modules = list(project.modules)
        _write_json(run_dir / "inventory.json", inventory.model_dump(mode="json"))
        return {"artifacts": ["inventory.json"], "metrics": {"vba_modules": len(project.modules), "available": project.available}, "diagnostics": [project.error] if project.error else []}
    if name == "vba.cell_links":
        from excel_to_act.ingest.vba import extract_vba_project
        from excel_to_act.inventory.vba_links import build_vba_edges, extract_vba_cell_links
        from excel_to_act.schemas import FormulaGraph

        inventory = _load_run_inventory(run_dir)
        source_path = Path(_run_source(run_dir)["source_path"])
        project = extract_vba_project(source_path)
        graph = FormulaGraph.model_validate(_read_json(run_dir / "dependency_graph.json")) if (run_dir / "dependency_graph.json").exists() else RegexFormulaGraphBuilder().build(inventory)
        refs = extract_vba_cell_links(project)
        nodes, edges = build_vba_edges(refs, workbook_path=str(source_path), known_sheets={sheet.name for sheet in inventory.sheets})
        known = {node.id for node in graph.nodes}
        graph.nodes.extend(node for node in nodes if node.id not in known)
        graph.edges.extend(edges)
        _write_json(run_dir / "dependency_graph.json", graph.model_dump(mode="json"))
        return {"artifacts": ["dependency_graph.json"], "metrics": {"vba_edges": len(edges)}}
    if name == "classify.rules":
        from excel_to_act.classify.classifier import RuleBasedClassifier
        from excel_to_act.schemas import FormulaGraph

        inventory = _load_run_inventory(run_dir)
        graph_path = run_dir / "dependency_graph.json"
        graph = FormulaGraph.model_validate(_read_json(graph_path)) if graph_path.exists() else RegexFormulaGraphBuilder().build(inventory)
        classification = RuleBasedClassifier().classify(inventory, graph)
        _write_json(run_dir / "module_classification.json", classification.model_dump(mode="json"))
        return {"artifacts": ["module_classification.json"], "metrics": {"items": len(classification.items)}}
    if name == "confirmation.build":
        from excel_to_act.confirm.templates import ConfirmationTemplateBuilder
        from excel_to_act.schemas import ModuleClassification

        classification_path = run_dir / "module_classification.json"
        if not classification_path.exists():
            _tool_action("classify.rules", run_dir)
        classification = ModuleClassification.model_validate(_read_json(classification_path))
        template = ConfirmationTemplateBuilder().build(classification)
        _write_json(run_dir / "confirmation_template.json", template.model_dump(mode="json"))
        return {"artifacts": ["confirmation_template.json"], "metrics": {"questions": len(template.questions)}}
    raise ValueError(f"unknown Step1 tool: {name}")


def _signature(quality: dict[str, Any]) -> str:
    metrics = quality.get("metrics", {})
    projection = {
        key: metrics.get(key)
        for key in (
            "traceability_ratio",
            "parsed_coverage_ratio",
            "parsed_package_parts_ratio",
            "opaque_rate",
            "package_preservation_ratio",
            "supported_fidelity_ratio",
            "logical_objects_accounted",
            "logical_objects_total",
            "package_parts_total",
            "parsed_package_parts",
            "opaque_parts",
        )
    }
    return _hash_bytes(_json_bytes({"status": quality.get("status"), "metrics_state": quality.get("metrics_state", "unavailable"), "blockers": quality.get("blockers", []), "metrics": projection, "thresholds": quality.get("thresholds", {})}))


def execute_tool(name: str, run_dir: Path) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    if not (run_dir / "source.json").is_file():
        return _run_error(name, run_dir, "Step1 run directory is missing source.json")
    if name in {"step1.check", "step1.coverage", "step1.fidelity"}:
        quality = _check_run(run_dir)
        source = _safe_run_source(run_dir)
        metric_names = {
            "step1.check": None,
            "step1.coverage": ("logical_objects_total", "logical_objects_accounted", "traceability_ratio", "parsed_objects_total", "parsed_objects", "parsed_coverage_ratio", "package_parts_total", "parsed_package_parts", "parsed_package_parts_total", "parsed_package_parts_ratio", "opaque_parts", "opaque_rate", "package_parts_preserved", "package_preservation_ratio"),
            "step1.fidelity": ("supported_facts_total", "supported_facts_exact", "supported_fidelity_ratio", "fidelity_deviations"),
        }[name]
        metrics = quality["metrics"] if metric_names is None else {key: quality["metrics"].get(key) for key in metric_names}
        return {
            "tool": name,
            "status": quality["status"],
            "metrics_state": quality["metrics_state"],
            "coverage_scope": quality["coverage_scope"],
            "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
            "run_id": quality.get("run_id"),
            "artifacts": _artifact_refs(run_dir) + [{"name": "handoff.json", "path": "handoff.json", "sha256": _hash_file(run_dir / "handoff.json")}],
            "metrics": metrics,
            "thresholds": quality.get("thresholds", {}),
            "diagnostics": quality["diagnostics"],
            "retryable": quality["retryable"],
            "next_tool": quality["next_tool"],
            "recovery_stopped": quality.get("recovery_stopped", False),
            "stop_reason": quality.get("stop_reason"),
            "remaining_tool": quality.get("remaining_tool"),
        }
    known = {tool["name"] for tool in _TOOLS}
    if name not in known or name in {"step1.convert", "step1.finalize"}:
        return {"tool": name, "status": "error", "source": None, "run_id": None, "artifacts": [], "metrics_state": "unavailable", "metrics": {}, "diagnostics": [_diagnostic("unknown_tool", "error", f"Tool cannot be invoked for a run: {name}")], "retryable": False, "next_tool": None}

    quality_before = _check_run(run_dir)
    try:
        history = _read_attempt_history(run_dir)
    except _AttemptHistoryError:
        history = None
        quality_before = _check_run(run_dir)
    recovery_tools = {"inventory.extract", "cached_values.read", "data_tables.read", "form_controls.read", "source_facts.refresh", "package.preserve"}
    source = _safe_run_source(run_dir)
    if name in recovery_tools and quality_before.get("recovery_stopped"):
        return {
            "tool": name,
            "status": "blocked",
            "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
            "run_id": source.get("run_id"),
            "artifacts": _artifact_refs(run_dir),
            "metrics_state": quality_before["metrics_state"],
            "metrics": quality_before["metrics"],
            "thresholds": quality_before.get("thresholds", {}),
            "diagnostics": quality_before["diagnostics"],
            "retryable": False,
            "next_tool": None,
            "recovery_stopped": True,
            "stop_reason": quality_before.get("stop_reason"),
            "remaining_tool": quality_before.get("remaining_tool"),
            "attempt": len(history.get("attempts", [])) if history is not None else None,
            "attempt_limit": history.get("attempt_limit") if history is not None else None,
        }
    suggested_tool = (quality_before.get("next_tool") or {}).get("name")
    recovery_action = name in recovery_tools and quality_before.get("status") == "fail" and suggested_tool == name
    if recovery_action:
        if history is None:
            quality_before = _check_run(run_dir)
            source = _safe_run_source(run_dir)
            return {
                "tool": name,
                "status": "blocked",
                "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
                "run_id": source.get("run_id"),
                "artifacts": _artifact_refs(run_dir),
                "metrics_state": quality_before["metrics_state"],
                "metrics": quality_before["metrics"],
                "thresholds": quality_before.get("thresholds", {}),
                "diagnostics": quality_before["diagnostics"],
                "retryable": False,
                "next_tool": None,
                "recovery_stopped": True,
                "stop_reason": quality_before.get("stop_reason"),
                "remaining_tool": quality_before.get("remaining_tool"),
                "attempt": None,
                "attempt_limit": None,
            }
        if len(history["attempts"]) >= history["attempt_limit"]:
            diagnostic = _diagnostic("attempt_limit_reached", "error", "The bounded recovery limit of three mutating tool attempts has been reached.")
            return {"tool": name, "status": "blocked", "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")}, "run_id": source.get("run_id"), "artifacts": [], "metrics_state": quality_before["metrics_state"], "metrics": quality_before["metrics"], "thresholds": quality_before.get("thresholds", {}), "diagnostics": [diagnostic], "retryable": False, "next_tool": None, "recovery_stopped": True, "stop_reason": "attempt_limit_reached", "remaining_tool": quality_before.get("next_tool")}
    before_signature = _signature(quality_before)
    try:
        action_result = _tool_action(name, run_dir)
        action_error = None
    except Exception as exc:
        action_result = {}
        action_error = f"{type(exc).__name__}: {exc}"
    quality_after = _check_run(run_dir)
    after_signature = _signature(quality_after)
    if recovery_action:
        attempts = history["attempts"]
        previous_no_progress = attempts[-1].get("no_progress_count", 0) if attempts else 0
        improved = before_signature != after_signature and (
            quality_after["status"] != "fail" or len(quality_after.get("blockers", [])) < len(quality_before.get("blockers", []))
        )
        no_progress = 0 if improved else previous_no_progress + 1
        attempts.append({"tool": name, "before_signature": before_signature, "after_signature": after_signature, "improved": improved, "no_progress_count": no_progress, "at": datetime.now(UTC).isoformat(), "error": action_error})
        history["no_progress_stopped"] = no_progress >= 2
        _write_json(run_dir / "attempt_history.json", history)
        quality_after = _check_run(run_dir)
    diagnostics = quality_after["diagnostics"]
    if action_error:
        diagnostics = [_diagnostic("tool_failed", "error", action_error), *diagnostics]
    return {
        "tool": name,
        "status": "error" if action_error else "blocked" if quality_after.get("recovery_stopped") else quality_after["status"],
        "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
        "run_id": source.get("run_id"),
        "artifacts": _artifact_refs(run_dir),
        "metrics_state": quality_after["metrics_state"],
        "metrics": quality_after["metrics"],
        "thresholds": quality_after.get("thresholds", {}),
        "diagnostics": diagnostics,
        "retryable": bool(quality_after.get("retryable") and not action_error),
        "next_tool": quality_after.get("next_tool"),
        "recovery_stopped": quality_after.get("recovery_stopped", False),
        "stop_reason": quality_after.get("stop_reason"),
        "remaining_tool": quality_after.get("remaining_tool"),
        "action": action_result,
        "attempt": len(history["attempts"]) if history is not None else None,
        "attempt_limit": history.get("attempt_limit") if history is not None else None,
    }


def auto_recover(run_dir: Path, *, max_attempts: int = 3) -> dict[str, Any]:
    if not 1 <= max_attempts <= 3:
        raise ValueError("max_attempts must be between 1 and 3")
    run_dir = run_dir.expanduser().resolve()
    if not (run_dir / "source.json").is_file():
        return _run_error("step1.auto_recover", run_dir, "Step1 run directory is missing source.json")
    quality = _check_run(run_dir)
    try:
        history = _read_attempt_history(run_dir)
    except _AttemptHistoryError:
        history = None
        quality = _check_run(run_dir)
    actions = []
    if history is not None and not quality.get("recovery_stopped"):
        history["attempt_limit"] = min(history["attempt_limit"], max_attempts)
        _write_json(run_dir / "attempt_history.json", history)
        quality = _check_run(run_dir)
    while quality.get("status") == "fail" and quality.get("retryable") and quality.get("next_tool") and not quality.get("recovery_stopped") and len(actions) < max_attempts:
        action_name = quality["next_tool"]["name"]
        result = execute_tool(action_name, run_dir)
        actions.append(result)
        quality = _check_run(run_dir)
        if result.get("status") == "error":
            break
    source = _safe_run_source(run_dir)
    return {
        "tool": "step1.auto_recover",
        "status": quality.get("status"),
        "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
        "run_id": source.get("run_id"),
        "artifacts": _artifact_refs(run_dir),
        "metrics_state": quality.get("metrics_state", "unavailable"),
        "metrics": quality.get("metrics", {}),
        "thresholds": quality.get("thresholds", {}),
        "diagnostics": quality.get("diagnostics", []),
        "retryable": bool(quality.get("retryable")),
        "next_tool": quality.get("next_tool"),
        "recovery_stopped": quality.get("recovery_stopped", False),
        "stop_reason": quality.get("stop_reason"),
        "remaining_tool": quality.get("remaining_tool"),
        "actions": actions,
        "attempts_used": len(actions),
        "attempt_limit": history.get("attempt_limit") if history is not None else None,
        "attempts_total": len(history["attempts"]) if history is not None else None,
    }


def finalize_run(run_dir: Path) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    if not (run_dir / "source.json").is_file():
        return _run_error("step1.finalize", run_dir, "Step1 run directory is missing source.json")
    quality = _check_run(run_dir)
    source = _safe_run_source(run_dir)
    output_root = run_dir.parents[5]
    batch_dir = run_dir.parents[3]
    relative_run = Path("final") / "batches" / str(source.get("batch_id", batch_dir.name)) / "sources" / str(source.get("sha256")) / run_dir.parent.name / run_dir.name / uuid4().hex[:8]
    promotion = output_root / relative_run
    final_relative: str | None = None
    if quality.get("ready_for_next_step"):
        promotion.parent.mkdir(parents=True, exist_ok=True)
        final_relative = relative_run.as_posix()
        _write_handoff(run_dir, quality, final_path=final_relative)
        shutil.copytree(run_dir, promotion)
        _write_handoff(promotion, quality, final_path=final_relative, output_root_override=output_root)
        final_files = {
            file.relative_to(promotion).as_posix(): _hash_file(file)
            for file in promotion.rglob("*")
            if file.is_file()
        }
        _write_json(promotion / "promotion.json", {"source_sha256": source["sha256"], "run_id": source["run_id"], "batch_id": source["batch_id"], "quality_status": quality["status"], "checks": "fresh source and candidate checks completed immediately before promotion", "files": final_files})
    if final_relative is None:
        _write_handoff(run_dir, quality)
    batch_path = batch_dir / "batch.json"
    if batch_path.exists():
        batch_data = _read_json(batch_path)
        for entry in batch_data.get("entries", []):
            if entry.get("run_id") == source.get("run_id"):
                entry["status"] = quality.get("status")
                entry["metrics_state"] = quality.get("metrics_state", "unavailable")
                entry["metrics"] = quality.get("metrics", {})
                entry["final_path"] = final_relative
                entry["handoff_path"] = (run_dir / "handoff.json").relative_to(output_root).as_posix()
        _batch_handoff(batch_dir, batch_data.get("entries", []), batch_data.get("discovery", {}))
    return {
        "tool": "step1.finalize",
        "status": quality.get("status"),
        "source": {"path": source.get("relative_path"), "sha256": source.get("sha256")},
        "run_id": source.get("run_id"),
        "artifacts": [{"name": "handoff.json", "path": _relative(run_dir / "handoff.json", output_root), "sha256": _hash_file(run_dir / "handoff.json")}, {"name": "handoff.md", "path": _relative(run_dir / "handoff.md", output_root), "sha256": _hash_file(run_dir / "handoff.md")}],
        "metrics_state": quality.get("metrics_state", "unavailable"),
        "metrics": quality.get("metrics", {}),
        "thresholds": quality.get("thresholds", {}),
        "diagnostics": quality.get("diagnostics", []),
        "retryable": bool(quality.get("retryable")),
        "next_tool": quality.get("next_tool"),
        "recovery_stopped": quality.get("recovery_stopped", False),
        "stop_reason": quality.get("stop_reason"),
        "remaining_tool": quality.get("remaining_tool"),
        "final_output": final_relative,
        "ready_for_next_step": quality.get("ready_for_next_step", False),
    }


def agent_definition() -> str:
    from importlib.resources import files

    return files("excel_to_act.steps.step1").joinpath("agent.md").read_text(encoding="utf-8")
