"""Handoff artifact: the contract a downstream agent reads to continue.

Step 1 ends here. Everything Step 2+ needs — where the artifacts are, what they
contain, what could not be parsed, and what blocks progress — is in one file so
no downstream agent has to guess or re-scan the whole output directory.
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path

from excel_to_act.schemas import (
    CompletenessReport,
    ConfirmationTemplate,
    CellKind,
    FormulaGraph,
    Handoff,
    HandoffArtifactRef,
    ModuleClassification,
    RunMetadata,
    UnsupportedSeverity,
    WorkbookInventory,
    WorkbookManifest,
    ValidationReport,
)

_KIND_BY_FILE = {
    "workbook_manifest.json": "manifest",
    "inventory.json": "inventory",
    "dependency_graph.json": "graph",
    "module_classification.json": "classification",
    "confirmation_template.json": "confirmation",
    "completeness.json": "completeness",
    "validation_report.json": "validation",
    "handoff.json": "handoff",
    "run_metadata.json": "metadata",
    "artifact_index.json": "index",
}

# Labels used by the human-facing markdown; keep them self-explanatory.
_KIND_LABEL = {
    "manifest": "package parts",
    "inventory": "cells and ranges",
    "graph": "dependency edges",
    "classification": "module classification",
    "confirmation": "open questions",
    "completeness": "completeness checks",
    "validation": "numerical comparisons",
    "metadata": "run metadata",
    "index": "run index",
    "other": "other",
}

_SUMMARY_LABEL = {
    "sheets": "Worksheets",
    "cells": "Cells",
    "formula_cells": "— of which formulas",
    "cached_values": "— with cached value",
    "defined_names": "Defined names",
    "tables": "Excel tables",
    "merged_ranges": "Merged ranges",
    "data_tables": "What-if data tables",
    "form_controls": "Form controls",
    "vba_modules": "VBA modules",
    "graph_nodes": "Graph nodes",
    "graph_edges": "Graph edges",
    "vba_edges": "— of which VBA",
}


def build_summary(inventory: WorkbookInventory, graph: FormulaGraph) -> dict[str, int]:
    """Counts a human (and Step 2's index) needs at a glance."""

    cells = [cell for sheet in inventory.sheets for cell in sheet.cells]
    kinds = Counter(
        item.kind
        for sheet in inventory.sheets
        for item in list(sheet.ranges) + list(sheet.layout_objects)
    )
    kinds.update(item.kind for item in inventory.workbook_ranges)
    return {
        "sheets": len(inventory.sheets),
        "cells": len(cells),
        "formula_cells": sum(1 for cell in cells if cell.kind == CellKind.formula),
        "cached_values": sum(1 for cell in cells if cell.cached_value_available),
        "defined_names": kinds.get("defined_name", 0),
        "tables": kinds.get("table", 0),
        "merged_ranges": kinds.get("merged_range", 0),
        "data_tables": kinds.get("data_table", 0),
        "form_controls": kinds.get("form_control", 0),
        "vba_modules": len(inventory.vba_modules),
        "graph_nodes": len(graph.nodes),
        "graph_edges": len(graph.edges),
        "vba_edges": sum(1 for edge in graph.edges if edge.relationship == "vba_ref"),
    }


def _count_for(
    name: str,
    manifest: WorkbookManifest,
    inventory: WorkbookInventory,
    graph: FormulaGraph,
    classification: ModuleClassification,
    confirmation: ConfirmationTemplate,
    completeness: CompletenessReport,
    validation: ValidationReport | None,
) -> int | None:
    if name == "workbook_manifest.json":
        return len(manifest.package_parts)
    if name == "inventory.json":
        return sum(len(sheet.cells) for sheet in inventory.sheets)
    if name == "dependency_graph.json":
        return len(graph.edges)
    if name == "module_classification.json":
        return len(classification.items)
    if name == "confirmation_template.json":
        return len(confirmation.questions)
    if name == "completeness.json":
        return len(completeness.checks)
    if name == "validation_report.json" and validation is not None:
        return validation.coverage.compared_cells
    return None


def build_handoff(
    manifest: WorkbookManifest,
    inventory: WorkbookInventory,
    graph: FormulaGraph,
    classification: ModuleClassification,
    confirmation: ConfirmationTemplate,
    completeness: CompletenessReport,
    metadata: RunMetadata,
    validation: ValidationReport | None = None,
) -> Handoff:
    artifacts = [
        HandoffArtifactRef(
            name=artifact.name,
            path=artifact.path,
            kind=_KIND_BY_FILE.get(artifact.name, "other"),
            count=_count_for(artifact.name, manifest, inventory, graph, classification, confirmation, completeness, validation),
            sha256=artifact.sha256,
        )
        for artifact in metadata.artifacts
        if artifact.name not in {"run_metadata.json", "artifact_index.json"}
    ]

    # The package ledger is the source of truth for unresolved parts. Diagnostics
    # may repeat a part or describe a warning without representing another object.
    all_features = inventory.unsupported_features + graph.unsupported_features
    opaque_parts = {part.name: part for part in manifest.package_parts if part.opaque}
    opaque_objects = {
        feature.source_location.source_identity: feature
        for feature in all_features
        if feature.opaque and feature.source_location.source_identity
        and feature.source_location.ooxml_part not in opaque_parts
    }
    opaque_summary = [
        {"feature_type": feature_type, "count": count}
        for feature_type, count in Counter(
            [part.opaque_reason or "unresolved package part" for part in opaque_parts.values()]
            + [feature.feature_type for feature in opaque_objects.values()]
        ).most_common()
    ]

    blockers = list(completeness.blocking_reasons)
    for feature in all_features:
        if feature.severity == UnsupportedSeverity.error and feature.description not in blockers:
            blockers.append(feature.description)

    # Blockers keep the check name so a reader can find the exact row in
    # completeness.json; warnings and next actions are prose only.
    warnings = [
        check.detail
        for check in completeness.checks
        if not check.passed and check.severity == UnsupportedSeverity.warning and check.detail
    ]
    for feature in all_features:
        if feature.severity == UnsupportedSeverity.warning and feature.description not in warnings:
            warnings.append(feature.description)
    next_actions: list[str] = [
        check.detail
        for check in completeness.checks
        if not check.passed and check.severity == UnsupportedSeverity.info and check.detail
    ]
    opaque_total = sum(entry["count"] for entry in opaque_summary)
    if opaque_total:
        next_actions.append(
            f"{opaque_total} unresolved (opaque) objects: decide whether to parse them "
            "or record them as accepted"
        )
    if confirmation.questions:
        next_actions.append(f"Answer {len(confirmation.questions)} confirmation question(s) before Step 3")
    if completeness.status == "fail":
        next_actions.append("Resolve the blockers above before any downstream consumer trusts inventory.json")
    if validation is not None and validation.status != "pass":
        next_actions.append(
            f"Numerical validation is {validation.status}; review validation_report.json before model generation"
        )
    next_actions.append("Step 2: build the agent-facing index from the artifacts above")

    return Handoff(
        workbook_sha256=manifest.sha256,
        workbook_path=manifest.workbook_path,
        run_id=metadata.run_id,
        status=completeness.status,
        numerical_status=validation.status if validation is not None else "not_run",
        summary=build_summary(inventory, graph),
        artifacts=artifacts,
        coverage=inventory.coverage,
        opaque_summary=opaque_summary,
        blockers=blockers,
        warnings=warnings,
        next_actions=next_actions,
    )


def render_handoff_markdown(handoff: Handoff) -> str:
    """Human-readable twin of ``handoff.json``.

    Deliberately short: a person should be able to read the top of this file and
    know whether the decomposition can be trusted, without opening any JSON.
    """

    workbook = Path(handoff.workbook_path).name if handoff.workbook_path else "(unknown)"
    artifact_names = {artifact.name for artifact in handoff.artifacts}
    detail = (
        "Full detail lives in `inventory.json` / `dependency_graph.json`."
        if {"inventory.json", "dependency_graph.json"} <= artifact_names
        else "Diagnostic detail lives in `workbook_manifest.json` / `completeness.json`."
    )
    lines = [
        f"# Handoff · {handoff.step} → {handoff.next_step}",
        "",
        f"**{workbook}** · structural **{handoff.status}** · numerical **{handoff.numerical_status}** · run `{handoff.run_id}`",
        "",
        f"> Human summary; machines read `handoff.json` in the same directory. {detail}",
        "",
    ]

    if handoff.summary:
        lines += ["## At a glance", "", "| Item | Count |", "| --- | --- |"]
        lines += [
            f"| {_SUMMARY_LABEL.get(key, key)} | {value} |" for key, value in handoff.summary.items()
        ]
        lines.append("")

    if handoff.blockers:
        lines += ["## Blockers (must fix first)", ""] + [f"- {item}" for item in handoff.blockers] + [""]
    if handoff.warnings:
        lines += ["## Warnings", ""] + [f"- {item}" for item in handoff.warnings] + [""]

    lines += ["## Artifacts", "", "| File | Contents | Items |", "| --- | --- | --- |"]
    lines += [
        f"| `{artifact.name}` | {_KIND_LABEL.get(artifact.kind, artifact.kind)} | "
        f"{artifact.count if artifact.count is not None else '-'} |"
        for artifact in handoff.artifacts
        if artifact.kind not in {"metadata", "index"}
    ]
    lines.append("")

    if handoff.opaque_summary:
        lines += ["## Unresolved (opaque)", "", "| Type | Count |", "| --- | --- |"]
        lines += [f"| {entry['feature_type']} | {entry['count']} |" for entry in handoff.opaque_summary]
        lines.append("")
    if handoff.next_actions:
        lines += ["## Next steps", ""] + [f"- {item}" for item in handoff.next_actions] + [""]

    return "\n".join(lines)
