"""Excel-native module classifier overlay."""

from __future__ import annotations

from excel_to_act.classify.rules import classify_cell, classify_range
from excel_to_act.schemas import (
    FormulaGraph,
    ModuleCategory,
    ModuleClassification,
    ModuleClassificationItem,
    SourceLocation,
    WorkbookInventory,
)


class RuleBasedClassifier:
    name = "rule_based_classifier"

    def classify(self, inventory: WorkbookInventory, graph: FormulaGraph) -> ModuleClassification:
        items: list[ModuleClassificationItem] = []
        for sheet in inventory.sheets:
            for cell in sheet.cells:
                category, confidence, reasons, hints = classify_cell(cell)
                items.append(ModuleClassificationItem(id=f"cell:{sheet.name}!{cell.address}", category=category, confidence=confidence, reasons=reasons, source_location=cell.source_location, source_artifact_refs=["inventory", "dependency_graph"], actuarial_hints=hints))
            for rng in [*sheet.ranges, *sheet.layout_objects]:
                category, confidence, reasons = classify_range(rng)
                items.append(ModuleClassificationItem(id=f"{rng.kind}:{sheet.name}!{rng.address}", category=category, confidence=confidence, reasons=reasons, source_location=rng.source_location, source_artifact_refs=["inventory"]))
        for rng in inventory.workbook_ranges:
            category, confidence, reasons = classify_range(rng)
            items.append(ModuleClassificationItem(id=f"{rng.kind}:{rng.name or rng.address}", category=category, confidence=confidence, reasons=reasons, source_location=rng.source_location, source_artifact_refs=["inventory"]))
        for i, feature in enumerate([*inventory.unsupported_features, *graph.unsupported_features]):
            items.append(ModuleClassificationItem(id=f"unsupported:{i}:{feature.feature_type}", category=ModuleCategory.unsupported_opaque, confidence=1.0, reasons=[feature.description], source_location=feature.source_location, source_artifact_refs=["inventory", "dependency_graph"]))
        if not items:
            loc = SourceLocation(workbook_path="unknown", object_type="workbook")
            items.append(ModuleClassificationItem(id="workbook:empty", category=ModuleCategory.other, confidence=0.1, reasons=["empty workbook catch-all"], source_location=loc))
        return ModuleClassification(items=items, unsupported_features=[*inventory.unsupported_features, *graph.unsupported_features])
