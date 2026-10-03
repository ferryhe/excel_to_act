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
# part prefix -> inventory evidence that proves the part was actually consumed.
# The prefix alone is not enough: a part is only accounted for when the output
# contains the corresponding objects, otherwise removing collection would go
# unnoticed.
_COVERED_PART_RULES: tuple[tuple[str, str], ...] = (
    ("xl/worksheets/", "cells"),
    ("xl/tables/", "table"),
    ("xl/comments/", "comment"),
)
_CELL_VALUE_TAGS = {"v", "f", "is"}


def _collected_evidence(inventory: WorkbookInventory) -> set[str]:
    """Evidence kinds present in the inventory (cells, table, comment, ...)."""

    evidence: set[str] = set()
    for sheet in inventory.sheets:
        if sheet.cells:
            evidence.add("cells")
        for item in list(sheet.ranges) + list(sheet.layout_objects):
            evidence.add(item.kind)
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
            detail="workbook.xml 里每个 <sheet> 都有对应 SheetInventory"
            if not missing_sheets
            else f"缺失的工作表: {missing_sheets}",
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
            detail="各表单元格数与 worksheet XML 独立扫描一致"
            if not cell_mismatches
            else "单元格数不一致: " + "; ".join(cell_mismatches),
        )
    )

    # 3. No package part was silently dropped: every part is either collected
    #    (with evidence in the output), marked opaque, or a known non-object.
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
        for prefix, required in _COVERED_PART_RULES:
            if name.startswith(prefix):
                accounted = required in evidence
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
            detail="每个内容部件都已采集（有产出证据）或已标 opaque"
            if not unaccounted_content
            else f"静默丢弃的部件: {unaccounted_content}",
        )
    )
    checks.append(
        CompletenessCheck(
            name="metadata_parts_accounted",
            passed=not unaccounted_metadata,
            severity=UnsupportedSeverity.info,
            expected=0,
            actual=len(unaccounted_metadata),
            detail=f"未建模的元数据部件（可解析，暂未采集）: {', '.join(unaccounted_metadata)}"
            if unaccounted_metadata
            else "没有未建模的元数据部件",
        )
    )

    # 4. Every formula cell produced an edge or is recorded as unparseable.
    outgoing = {edge.source for edge in graph.edges}
    recorded = {
        feature.source_location.object_id
        for feature in graph.unsupported_features
        if feature.source_location and feature.source_location.object_id
    }
    formula_cells = [
        f"{sheet.name}!{cell.address}" for sheet in inventory.sheets for cell in sheet.cells if cell.formula
    ]
    unlinked = [ref for ref in formula_cells if f"cell:{ref}" not in outgoing and ref not in recorded]
    checks.append(
        CompletenessCheck(
            name="formulas_linked",
            passed=not unlinked,
            severity=UnsupportedSeverity.warning,
            expected=len(formula_cells),
            actual=len(formula_cells) - len(unlinked),
            detail="每个公式格都有依赖边或有 unparseable 记录"
            if not unlinked
            else f"无依赖边的公式格: {unlinked[:10]}",
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
            detail="覆盖等式闭合：recognized + opaque == 独立重算的 discovered"
            if arithmetic_ok
            else f"覆盖等式未闭合：recognized({recognized}) + opaque({opaque}) ≠ 独立重算 discovered({discovered})；"
            "inventory 自带的 coverage 是恒等式，探测不到这个差异（issue #5）",
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
