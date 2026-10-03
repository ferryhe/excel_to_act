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
    FormulaGraph,
    Handoff,
    HandoffArtifactRef,
    ModuleClassification,
    RunMetadata,
    UnsupportedSeverity,
    WorkbookInventory,
    WorkbookManifest,
)

_KIND_BY_FILE = {
    "workbook_manifest.json": "manifest",
    "inventory.json": "inventory",
    "dependency_graph.json": "graph",
    "module_classification.json": "classification",
    "confirmation_template.json": "confirmation",
    "completeness.json": "completeness",
    "handoff.json": "handoff",
    "run_metadata.json": "metadata",
    "artifact_index.json": "index",
}

# Labels used by the human-facing markdown; keep them self-explanatory.
_KIND_LABEL = {
    "manifest": "包部件清单",
    "inventory": "单元格与范围清单",
    "graph": "依赖边",
    "classification": "模块分类",
    "confirmation": "待确认问题",
    "completeness": "完备性检查项",
    "metadata": "运行元数据",
    "index": "历史运行索引",
    "other": "其他",
}

_SUMMARY_LABEL = {
    "sheets": "工作表",
    "cells": "单元格",
    "formula_cells": "其中公式格",
    "cached_values": "有缓存值",
    "defined_names": "已定义名称",
    "tables": "Excel 表",
    "merged_ranges": "合并区",
    "data_tables": "模拟运算表",
    "form_controls": "表单控件",
    "vba_modules": "VBA 模块",
    "graph_nodes": "依赖图节点",
    "graph_edges": "依赖图边",
    "vba_edges": "其中 VBA 边",
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
        "formula_cells": sum(1 for cell in cells if cell.formula),
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
    return None


def build_handoff(
    manifest: WorkbookManifest,
    inventory: WorkbookInventory,
    graph: FormulaGraph,
    classification: ModuleClassification,
    confirmation: ConfirmationTemplate,
    completeness: CompletenessReport,
    metadata: RunMetadata,
) -> Handoff:
    artifacts = [
        HandoffArtifactRef(
            name=artifact.name,
            path=artifact.path,
            kind=_KIND_BY_FILE.get(artifact.name, "other"),
            count=_count_for(artifact.name, manifest, inventory, graph, classification, confirmation, completeness),
            sha256=artifact.sha256,
        )
        for artifact in metadata.artifacts
    ]

    all_features = manifest.unsupported_features + inventory.unsupported_features + graph.unsupported_features
    # Only genuinely unparseable parts count as opaque; warnings such as
    # "missing_cached_values" are recoverable and must not inflate this number.
    opaque_features = [feature for feature in all_features if feature.opaque]
    opaque_summary = [
        {"feature_type": feature_type, "count": count}
        for feature_type, count in Counter(feature.feature_type for feature in opaque_features).most_common()
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
        next_actions.append(f"{opaque_total} 个未解析对象（opaque）：决定是补解析，还是记录为「接受」")
    if confirmation.questions:
        next_actions.append(f"Step 3 之前需回答 {len(confirmation.questions)} 个确认问题")
    if completeness.status == "fail":
        next_actions.append("先解决上面的阻塞项，再让任何下游消费 inventory.json")
    next_actions.append("Step 2：基于上述产物建立 agent 使用的索引")

    return Handoff(
        workbook_sha256=manifest.sha256,
        workbook_path=manifest.workbook_path,
        run_id=metadata.run_id,
        status=completeness.status,
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
    lines = [
        f"# Handoff · {handoff.step} → {handoff.next_step}",
        "",
        f"**{workbook}** · 状态 **{handoff.status}** · run `{handoff.run_id}`",
        "",
        "> 这是给人看的摘要；机器读同目录 `handoff.json`，完整明细在 `inventory.json` / `dependency_graph.json`。",
        "",
    ]

    if handoff.summary:
        lines += ["## 一眼看懂", "", "| 项 | 数量 |", "| --- | --- |"]
        lines += [
            f"| {_SUMMARY_LABEL.get(key, key)} | {value} |" for key, value in handoff.summary.items()
        ]
        lines.append("")

    if handoff.blockers:
        lines += ["## ⛔ 阻塞（必须先处理）", ""] + [f"- {item}" for item in handoff.blockers] + [""]
    if handoff.warnings:
        lines += ["## ⚠️ 警告", ""] + [f"- {item}" for item in handoff.warnings] + [""]

    lines += ["## 产出", "", "| 文件 | 内容 | 条数 |", "| --- | --- | --- |"]
    lines += [
        f"| `{artifact.name}` | {_KIND_LABEL.get(artifact.kind, artifact.kind)} | "
        f"{artifact.count if artifact.count is not None else '-'} |"
        for artifact in handoff.artifacts
        if artifact.kind not in {"metadata", "index"}
    ]
    lines.append("")

    if handoff.opaque_summary:
        lines += ["## 未解析（opaque）", "", "| 类型 | 数量 |", "| --- | --- |"]
        lines += [f"| {entry['feature_type']} | {entry['count']} |" for entry in handoff.opaque_summary]
        lines.append("")
    if handoff.next_actions:
        lines += ["## 下一步", ""] + [f"- {item}" for item in handoff.next_actions] + [""]

    return "\n".join(lines)
