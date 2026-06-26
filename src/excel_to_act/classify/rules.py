"""Rule helpers for Excel-native module classification."""

from __future__ import annotations

from excel_to_act.schemas import ActuarialHint, CellInventory, ModuleCategory, RangeInventory


def hints_for_text(text: str | None) -> list[ActuarialHint]:
    if not text:
        return []
    lower = text.lower()
    hints: list[ActuarialHint] = []
    if any(token in lower for token in ["assumption", "mortality", "lapse", "expense"]):
        hints.append(ActuarialHint.assumption)
    if any(token in lower for token in ["rate", "table"]):
        hints.append(ActuarialHint.rate_table)
    if any(token in lower for token in ["cashflow", "cash flow", "premium", "claim"]):
        hints.append(ActuarialHint.cashflow)
    if any(token in lower for token in ["projection", "duration"]):
        hints.append(ActuarialHint.projection)
    if any(token in lower for token in ["output", "result"]):
        hints.append(ActuarialHint.output)
    return hints


def classify_cell(cell: CellInventory) -> tuple[ModuleCategory, float, list[str], list[ActuarialHint]]:
    text = str(cell.value or cell.formula or "")
    hints = hints_for_text(text)
    if cell.formula:
        if any(fn in cell.formula.upper() for fn in ["VLOOKUP", "XLOOKUP", "HLOOKUP", "INDEX", "MATCH"]):
            return ModuleCategory.lookup_block, 0.8, ["formula uses lookup function"], hints
        return ModuleCategory.formula_block, 0.75, ["cell contains formula; dependencies parsed separately"], hints
    if cell.number_format and cell.number_format != "General":
        return ModuleCategory.input, 0.55, ["literal cell with explicit number format"], hints
    if isinstance(cell.value, str):
        return ModuleCategory.presentation, 0.6, ["text label or presentation content"], hints
    return ModuleCategory.input, 0.5, ["literal value cell"], hints


def classify_range(rng: RangeInventory) -> tuple[ModuleCategory, float, list[str]]:
    if rng.kind == "table":
        return ModuleCategory.data_table, 0.85, ["Excel table object"]
    if rng.kind in {"merged_range", "row_layout", "column_layout", "freeze_panes", "conditional_formatting"}:
        return ModuleCategory.presentation, 0.6, [f"layout/presentation object: {rng.kind}"]
    if rng.kind in {"data_validation"}:
        return ModuleCategory.input, 0.65, ["data validation constrains input"]
    return ModuleCategory.other, 0.4, [f"catch-all classification for {rng.kind}"]
