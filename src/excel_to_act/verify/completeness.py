"""Independent identity accounting for legacy Phase 1 inspect output."""

from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from excel_to_act.schemas import (
    CompletenessCheck,
    CompletenessReport,
    CompletenessStatus,
    FormulaGraph,
    ModuleClassification,
    ObjectCoverage,
    SourceLocation,
    UnsupportedSeverity,
    WorkbookInventory,
    WorkbookManifest,
)
from excel_to_act.steps.step1.source_scan import object_identity, scan_step1_source


def _inventory_objects(inventory: WorkbookInventory) -> list[tuple[Any, str, str | None]]:
    objects: list[tuple[Any, str, str | None]] = []
    for sheet in inventory.sheets:
        objects.append((sheet, "sheet", sheet.name))
        objects.extend((cell, "cell", sheet.name) for cell in sheet.cells)
        objects.extend((item, item.kind, sheet.name) for item in (*sheet.ranges, *sheet.layout_objects))
    objects.extend(
        (item, item.kind, None if item.kind == "defined_name" and item.metadata.get("scope", "workbook") == "workbook" else item.metadata.get("scope"))
        for item in inventory.workbook_ranges
    )
    return objects


def _source_location(path: str, source: dict[str, Any]) -> SourceLocation:
    details = source["details"]
    sheet = details.get("sheet") or (details.get("name") if source["kind"] == "sheet" else None)
    if source["kind"] == "defined_name" and details.get("scope") != "workbook":
        sheet = details.get("scope")
    return SourceLocation(
        workbook_path=path,
        sheet_name=sheet,
        address=details.get("address") if sheet else None,
        object_type=source["kind"],
        object_id=details.get("name") or details.get("address") or source["identity"],
        ooxml_part=source["part"],
        source_identity=source["identity"],
    )


def _formula_check(inventory: WorkbookInventory, graph: FormulaGraph) -> CompletenessCheck:
    outgoing = {edge.source for edge in graph.edges}
    resolved = {node.id for node in graph.nodes if node.kind == "cell" and node.metadata.get("references_resolved") is True}
    recorded = {feature.source_location.object_id for feature in graph.unsupported_features if feature.source_location and feature.source_location.object_id}
    formula_cells = [f"{sheet.name}!{cell.address}" for sheet in inventory.sheets for cell in sheet.cells if cell.formula]
    unlinked = [ref for ref in formula_cells if f"cell:{ref}" not in outgoing and f"cell:{ref}" not in resolved and ref not in recorded]
    return CompletenessCheck(
        name="formulas_linked", passed=not unlinked, severity=UnsupportedSeverity.warning,
        expected=len(formula_cells), actual=len(formula_cells) - len(unlinked),
        detail="formula cells without an edge: " + str(unlinked[:10]) if unlinked else "formula cells linked or recorded as unparseable",
    )


def verify_completeness(
    manifest: WorkbookManifest,
    inventory: WorkbookInventory,
    graph: FormulaGraph,
    classification: ModuleClassification,
) -> CompletenessReport:
    """Compare output identities with fresh OOXML declarations and package members."""

    del classification  # Classification is downstream evidence, not an object ledger.
    path = manifest.workbook_path
    formula_check = _formula_check(inventory, graph)
    try:
        scan = scan_step1_source(Path(path))
    except Exception as exc:
        detail = f"Source XML scan failed: {type(exc).__name__}: {exc}"
        return CompletenessReport(
            workbook_sha256=manifest.sha256,
            status=CompletenessStatus.fail,
            checks=[formula_check, CompletenessCheck(name="source_scan", passed=False, severity=UnsupportedSeverity.error, detail=detail)],
            blocking_reasons=[detail],
        )

    source_by_id = {obj["identity"]: obj for obj in scan["objects"]}
    expected: dict[tuple[str | None, str], set[str]] = defaultdict(set)
    actual: dict[tuple[str | None, str], list[str]] = defaultdict(list)
    for source in scan["objects"]:
        loc = _source_location(path, source)
        expected[(loc.sheet_name, source["kind"])].add(source["identity"])
    for record, kind, sheet in _inventory_objects(inventory):
        identity = record.source_identity
        if identity:
            actual[(sheet, kind)].append(identity)

    rows: list[ObjectCoverage] = []
    gaps: list[SourceLocation] = []
    checks: list[CompletenessCheck] = []
    for sheet, kind in sorted(expected.keys() | actual.keys(), key=lambda key: (key[0] or "", key[1])):
        actual_ids = set(actual[(sheet, kind)])
        missing = sorted(expected[(sheet, kind)] - actual_ids)
        unexpected = sorted(actual_ids - expected[(sheet, kind)])
        duplicate = sorted(identity for identity, count in Counter(actual[(sheet, kind)]).items() if count > 1)
        actual_count = len(actual[(sheet, kind)])
        rows.append(ObjectCoverage(
            sheet_name=sheet, kind=kind,
            expected=len(expected[(sheet, kind)]), actual=actual_count,
            missing_identities=missing, unexpected_identities=unexpected, duplicate_identities=duplicate,
        ))
        gaps.extend(_source_location(path, source_by_id[identity]) for identity in missing)
        if missing or unexpected or duplicate:
            checks.append(CompletenessCheck(
                name=f"{kind}_identities", passed=False, severity=UnsupportedSeverity.error,
                expected=len(expected[(sheet, kind)]), actual=actual_count,
                detail=f"{sheet or 'workbook'} {kind}: missing={missing}, unexpected={unexpected}, duplicate={duplicate}",
            ))

    source_parts = {name: fact for name, fact in scan["parts"].items()}
    manifest_parts = {part.name: part for part in manifest.package_parts}
    missing_parts = sorted(source_parts.keys() - manifest_parts.keys())
    unexpected_parts = sorted(manifest_parts.keys() - source_parts.keys())
    wrong_opaque = sorted(
        name for name in source_parts.keys() & manifest_parts.keys()
        if source_parts[name]["opaque"] != manifest_parts[name].opaque
    )
    duplicate_parts = sorted(name for name, count in Counter(part.name for part in manifest.package_parts).items() if count > 1)
    package_row = ObjectCoverage(
        kind="package_part", expected=len(source_parts), actual=len(manifest.package_parts),
        missing_identities=[object_identity("package_part", name) for name in missing_parts],
        unexpected_identities=[object_identity("package_part", name) for name in unexpected_parts],
        duplicate_identities=[object_identity("package_part", name) for name in duplicate_parts],
        opaque_expected=sum(bool(part["opaque"]) for part in source_parts.values()),
        opaque_actual=sum(part.opaque for part in manifest_parts.values()),
    )
    rows.append(package_row)
    gaps.extend(SourceLocation(
        workbook_path=path, ooxml_part=name, object_type="package_part", object_id=name,
        source_identity=object_identity("package_part", name),
    ) for name in (*missing_parts, *wrong_opaque))
    checks.extend([
        CompletenessCheck(
            name="sheets_accounted", passed=not any(row.kind == "sheet" and (row.missing_identities or row.unexpected_identities or row.duplicate_identities) for row in rows),
            severity=UnsupportedSeverity.error,
            expected=sum(row.expected for row in rows if row.kind == "sheet"),
            actual=sum(row.actual for row in rows if row.kind == "sheet"),
            detail="worksheet identities match source declarations",
        ),
        CompletenessCheck(
            name="cells_accounted", passed=not any(row.kind == "cell" and (row.missing_identities or row.unexpected_identities or row.duplicate_identities) for row in rows),
            severity=UnsupportedSeverity.error,
            expected=sum(row.expected for row in rows if row.kind == "cell"),
            actual=sum(row.actual for row in rows if row.kind == "cell"),
            detail="cell identities match source declarations",
        ),
        CompletenessCheck(
            name="content_parts_accounted", passed=not (missing_parts or unexpected_parts or wrong_opaque or duplicate_parts),
            severity=UnsupportedSeverity.error,
            expected=len(source_parts), actual=len(manifest.package_parts),
            detail=f"package parts: missing={missing_parts}, unexpected={unexpected_parts}, duplicate={duplicate_parts}, opaque mismatch={wrong_opaque}",
        ),
    ])

    checks.append(formula_check)

    recognized = len({identity for ids in actual.values() for identity in ids})
    discovered = len(source_by_id)
    arithmetic_ok = recognized == discovered and inventory.coverage.recognized_inventory_objects == recognized and inventory.coverage.discovered_workbook_objects == discovered and inventory.coverage.unsupported_or_opaque_objects == 0
    checks.append(CompletenessCheck(
        name="coverage_arithmetic", passed=arithmetic_ok, severity=UnsupportedSeverity.error,
        expected=discovered, actual=recognized,
        detail=f"logical objects: source={discovered}, actual={recognized}, inventory coverage={inventory.coverage.model_dump()}; package parts are a separate ledger",
    ))
    status = CompletenessStatus.ok
    if any(not check.passed and check.severity == UnsupportedSeverity.error for check in checks):
        status = CompletenessStatus.fail
    elif any(not check.passed and check.severity == UnsupportedSeverity.warning for check in checks):
        status = CompletenessStatus.warn
    return CompletenessReport(
        workbook_sha256=manifest.sha256,
        status=status,
        checks=checks,
        recognized_inventory_objects=recognized,
        unsupported_or_opaque_objects=0,
        discovered_workbook_objects=discovered,
        object_coverage=rows,
        gaps=gaps,
        blocking_reasons=[f"{check.name}: {check.detail}" for check in checks if not check.passed and check.severity == UnsupportedSeverity.error],
    )
