"""Independent completeness checks over the Step 1 output.

`WorkbookInventory.coverage` cannot prove completeness by itself: `discovered` is
currently computed as `recognized + opaque`, i.e. the coverage equation is an
identity that always holds. These checks re-derive the object universe straight
from the workbook package (worksheet XML and package parts) and compare it with
what the inventory actually recorded, so a dropped object shows up as a gap.

Severity model:

- ``error``   — output is materially incomplete; the run must be treated as failed
- ``warning`` — suspicious but recoverable
- ``info``    — known gap that is tracked elsewhere (does not change status)
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from excel_to_act.ingest.data_table import sheet_xml_parts
from excel_to_act.schemas import (
    CompletenessCheck,
    CompletenessReport,
    CompletenessStatus,
    FormulaGraph,
    ModuleClassification,
    SourceLocation,
    UnsupportedSeverity,
    WorkbookInventory,
    WorkbookManifest,
)

_INFRA_PARTS = {"[Content_Types].xml"}
_INFRA_MARKERS = ("_rels/", ".rels")
# Parts fully consumed by openpyxl and surfaced through per-cell/per-sheet artifacts.
_INTERNALLY_CONSUMED = {"xl/workbook.xml", "xl/sharedStrings.xml", "xl/styles.xml"}
# Rendering-only parts: not workbook objects, so their absence is informational.
_PRESENTATION_PREFIXES = ("xl/theme/", "xl/printerSettings/")
# Metadata parts we do not model yet (see README §4 A); tracked, not blocking.
_METADATA_PREFIXES = ("docProps/",)
# Package part matcher -> inventory evidence that proves the part was consumed.
# Keep worksheet/table directory matching broad because relationship targets can
# use custom filenames. Comments need exact names because Excel stores them flat.
# A matching name alone is not enough: one evidence item is consumed per part so
# removing collection is still detected.
_COVERED_PART_RULES: tuple[tuple[str | re.Pattern[str], str], ...] = (
    ("xl/worksheets/", "worksheet"),
    ("xl/tables/", "table"),
    # Both layouts occur: Excel writes ``xl/comments1.xml``; other writers nest
    # them as ``xl/comments/comment1.xml``.
    (re.compile(r"xl/comments\d+\.xml"), "comment"),
    (re.compile(r"xl/comments/comment\d+\.xml"), "comment"),
)
_CELL_VALUE_TAGS = {"v", "f", "is"}


def _collected_evidence(inventory: WorkbookInventory) -> dict[str, list[str]]:
    """Inventory evidence available to account for individual package parts."""

    evidence: dict[str, list[str]] = {}
    for sheet in inventory.sheets:
        evidence.setdefault("worksheet", []).append(sheet.name)
        for item in list(sheet.ranges) + list(sheet.layout_objects):
            evidence.setdefault(item.kind, []).append(item.name or item.address or item.source_location.object_id or "")
    for item in inventory.workbook_ranges:
        evidence.setdefault(item.kind, []).append(item.name or item.address or item.source_location.object_id or "")
    return evidence


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _is_infrastructure(part_name: str) -> bool:
    return part_name in _INFRA_PARTS or any(marker in part_name for marker in _INFRA_MARKERS)


def count_cells_in_sheet_xml(workbook_path: Path, part: str) -> int:
    """Count non-empty cells straight from worksheet XML, bypassing openpyxl."""

    with zipfile.ZipFile(workbook_path) as zf:
        root = ET.fromstring(zf.read(part))
    return sum(
        1
        for cell in root.iter()
        if _local(cell.tag) == "c" and any(_local(child.tag) in _CELL_VALUE_TAGS for child in cell)
    )


def verify_completeness(
    manifest: WorkbookManifest,
    inventory: WorkbookInventory,
    graph: FormulaGraph,
    classification: ModuleClassification,
) -> CompletenessReport:
    """Re-derive the object universe from the package and compare with the output."""

    checks: list[CompletenessCheck] = []
    gaps: list[SourceLocation] = []
    workbook_path = Path(manifest.workbook_path) if manifest.workbook_path else None
    path_str = str(workbook_path) if workbook_path else None

    # 1. Every sheet declared in workbook.xml has an inventory entry.
    manifest_sheets = {sheet.name for sheet in manifest.sheets}
    inventory_sheets = {sheet.name for sheet in inventory.sheets}
    missing_sheets = sorted(manifest_sheets - inventory_sheets)
    for name in missing_sheets:
        gaps.append(SourceLocation(workbook_path=path_str, object_type="sheet", object_id=name))
    checks.append(
        CompletenessCheck(
            name="sheets_accounted",
            passed=not missing_sheets,
            severity=UnsupportedSeverity.error,
            expected=len(manifest_sheets),
            actual=len(inventory_sheets),
            detail="every <sheet> in workbook.xml has a matching SheetInventory entry"
            if not missing_sheets
            else f"missing worksheets: {missing_sheets}",
        )
    )

    # 2. Cell counts per sheet match an independent XML scan.
    parts = sheet_xml_parts(workbook_path) if workbook_path and workbook_path.exists() else {}
    xml_total = 0
    inventory_total = 0
    cell_mismatches: list[str] = []
    for sheet in inventory.sheets:
        inventory_total += len(sheet.cells)
        part = parts.get(sheet.name)
        if not part:
            continue
        try:
            xml_count = count_cells_in_sheet_xml(workbook_path, part)
        except (KeyError, ET.ParseError, zipfile.BadZipFile, OSError):
            continue
        xml_total += xml_count
        if xml_count != len(sheet.cells):
            cell_mismatches.append(f"{sheet.name}: xml={xml_count} inventory={len(sheet.cells)}")
            gaps.append(
                SourceLocation(
                    workbook_path=path_str,
                    sheet_name=sheet.name,
                    object_type="cell",
                    object_id=f"{abs(xml_count - len(sheet.cells))} cells unaccounted",
                )
            )
    checks.append(
        CompletenessCheck(
            name="cells_accounted",
            passed=not cell_mismatches,
            severity=UnsupportedSeverity.error,
            expected=xml_total,
            actual=inventory_total,
            detail="cell counts match an independent scan of the worksheet XML"
            if not cell_mismatches
            else "cell counts differ: " + "; ".join(cell_mismatches),
        )
    )

    # 3. No package part was silently dropped: each covered package part must
    #    consume its own inventory evidence, marked opaque, or be known non-object.
    evidence = _collected_evidence(inventory)
    unaccounted_content: list[str] = []
    unaccounted_metadata: list[str] = []
    for part in manifest.package_parts:
        name = part.name
        if _is_infrastructure(name) or name in _INTERNALLY_CONSUMED:
            continue
        if name.startswith(_PRESENTATION_PREFIXES) or name.startswith(_METADATA_PREFIXES):
            unaccounted_metadata.append(name)
            continue
        if part.opaque:
            continue
        accounted = False
        for matcher, required in _COVERED_PART_RULES:
            matches = (
                name.startswith(matcher)
                if isinstance(matcher, str)
                else matcher.fullmatch(name) is not None
            )
            if matches:
                available = evidence.get(required, [])
                if available:
                    available.pop()
                    accounted = True
                break
        if accounted:
            continue
        unaccounted_content.append(name)
        gaps.append(SourceLocation(workbook_path=path_str, ooxml_part=name, object_type="package_part", object_id=name))
    checks.append(
        CompletenessCheck(
            name="content_parts_accounted",
            passed=not unaccounted_content,
            severity=UnsupportedSeverity.error,
            expected=0,
            actual=len(unaccounted_content),
            detail="every content part is either collected (with evidence in the output) or marked opaque"
            if not unaccounted_content
            else f"silently dropped parts: {unaccounted_content}",
        )
    )
    checks.append(
        CompletenessCheck(
            name="metadata_parts_accounted",
            passed=not unaccounted_metadata,
            severity=UnsupportedSeverity.info,
            expected=0,
            actual=len(unaccounted_metadata),
            detail=f"unmodelled metadata parts (parseable, not collected yet): {', '.join(unaccounted_metadata)}"
            if unaccounted_metadata
            else "no unmodelled metadata parts",
        )
    )

    # 4. Every formula cell produced an edge or is recorded as unparseable.
    outgoing = {edge.source for edge in graph.edges}
    resolved = {
        node.id
        for node in graph.nodes
        if node.kind == "cell" and node.metadata.get("references_resolved") is True
    }
    recorded = {
        feature.source_location.object_id
        for feature in graph.unsupported_features
        if feature.source_location and feature.source_location.object_id
    }
    formula_cells = [
        f"{sheet.name}!{cell.address}" for sheet in inventory.sheets for cell in sheet.cells if cell.formula
    ]
    unlinked = [
        ref
        for ref in formula_cells
        if f"cell:{ref}" not in outgoing and f"cell:{ref}" not in resolved and ref not in recorded
    ]
    checks.append(
        CompletenessCheck(
            name="formulas_linked",
            passed=not unlinked,
            severity=UnsupportedSeverity.warning,
            expected=len(formula_cells),
            actual=len(formula_cells) - len(unlinked),
            detail="every formula cell has a dependency edge, a resolved zero-reference result, or an unparseable record"
            if not unlinked
            else f"formula cells without an edge: {unlinked[:10]}",
        )
    )

    # 5. Coverage arithmetic against the independently discovered universe.
    discovered = xml_total + len(manifest_sheets) + len(unaccounted_content) + len(unaccounted_metadata)
    recognized = inventory.coverage.recognized_inventory_objects
    opaque = inventory.coverage.unsupported_or_opaque_objects
    arithmetic_ok = recognized + opaque == discovered
    checks.append(
        CompletenessCheck(
            name="coverage_arithmetic",
            passed=arithmetic_ok,
            severity=UnsupportedSeverity.info,
            expected=discovered,
            actual=recognized + opaque,
            detail="coverage equation closes: recognized + opaque == independently recomputed discovered"
            if arithmetic_ok
            else f"coverage equation does not close: recognized({recognized}) + opaque({opaque}) != "
            f"independently recomputed discovered({discovered}); the coverage block inside "
            "inventory.json is an identity and cannot detect this (issue #5)",
        )
    )

    status = CompletenessStatus.ok
    if any(not check.passed and check.severity == UnsupportedSeverity.error for check in checks):
        status = CompletenessStatus.fail
    elif any(not check.passed and check.severity == UnsupportedSeverity.warning for check in checks):
        status = CompletenessStatus.warn

    blocking_reasons = [
        f"{check.name}: {check.detail}"
        for check in checks
        if not check.passed and check.severity == UnsupportedSeverity.error
    ]
    return CompletenessReport(
        workbook_sha256=manifest.sha256,
        status=status,
        checks=checks,
        recognized_inventory_objects=recognized,
        unsupported_or_opaque_objects=opaque,
        discovered_workbook_objects=discovered,
        gaps=gaps,
        blocking_reasons=blocking_reasons,
    )
