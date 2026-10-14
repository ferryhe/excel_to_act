from __future__ import annotations

from typing import Any

from excel_to_act.steps.conversion_workflow import review_tool_entries


def tool_catalog() -> dict[str, Any]:
    return {"schema_version": "step5.tools.v1", "tools": [
        {"name": "step5.agent", "command": "step5 agent", "outputs": ["Stage 5 review contract"]},
        {"name": "step5.tools", "command": "step5 tools", "outputs": ["Stage 5 commands and limits"]},
        {"name": "step5.oracle", "command": "step5 oracle --workflow DIR [--validation VARIABLE_VALIDATION.json] [--target NAME] [--range SHEET!A1:B2]", "outputs": ["Stage 5 native evidence (does not advance the stage)"]},
        {"name": "step5.validate", "command": "step5 validate --workflow DIR [--inputs REQUEST.json]", "outputs": ["isolated generated-code execution evidence"]},
        {"name": "step5.reconcile", "command": "step5 reconcile --workflow DIR --validation FILE --oracle FILE", "outputs": ["validation_report.json/.md"]},
        {"name": "step5.finalize-cases", "command": "step5 finalize-cases --workflow DIR --case-report FILE (repeat 32 times)", "outputs": ["formal exact-case validation_report.json/.md"]},
        *review_tool_entries(5),
    ], "limits": ["Native Excel capture requires Windows and Microsoft Excel.",
                   "Step 5 uses only generated code and a fresh native Excel oracle; it does not calculate with formula caches.",
                   "Variable cases require --inputs on validate, then --validation on oracle, then reconcile with that exact validation and oracle; case receipts bind the input digest without changing the saved-scenario Stage 5 checkpoint.",
                   "Optional target and range captures may include exploratory evidence outside the plan; capture does not expand the approved comparison scope or advance Stage 5.",
                   "Reconciliation compares only bound design targets and active formula addresses present in both native and generated results. It reports captured selectors separately from exact compared active formula addresses."]}
