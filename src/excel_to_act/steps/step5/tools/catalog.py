from __future__ import annotations

from typing import Any


def tool_catalog() -> dict[str, Any]:
    return {"schema_version": "step5.tools.v1", "tools": [
        {"name": "step5.agent", "command": "step5 agent", "outputs": ["Stage 5 review contract"]},
        {"name": "step5.tools", "command": "step5 tools", "outputs": ["Stage 5 commands and limits"]},
        {"name": "step5.oracle", "command": "step5 oracle --workflow DIR [--target NAME] [--range SHEET!A1:B2]", "outputs": ["Stage 5 native evidence (does not advance the stage)"]},
        {"name": "step5.validate", "command": "step5 validate --workflow DIR", "outputs": ["isolated generated-code execution evidence"]},
        {"name": "step5.reconcile", "command": "step5 reconcile --workflow DIR --validation FILE --oracle FILE", "outputs": ["validation_report.json/.md"]},
        {"name": "workflow.typesafe", "command": "workflow typesafe --workflow DIR --stage 5 --request REQUEST.json", "outputs": ["live, current-artifact-bound TypeSafe review evidence"]},
        {"name": "workflow.confirm", "command": "workflow confirm --workflow DIR --stage 5 --reviewer agent|human|typesafe --decision approve|reject --message TEXT [--evidence REVIEW.json]", "outputs": ["review receipt"]},
        {"name": "workflow.reject", "command": "workflow reject --workflow DIR --stage 5 --reviewer agent|human|typesafe --return-to N --message TEXT [--evidence REVIEW.json]", "outputs": ["rejection receipt"]},
        {"name": "workflow.status", "command": "workflow status --workflow DIR", "outputs": ["freshness, approvals, and next permitted stage"]},
    ], "limits": ["Native Excel capture requires Windows and Microsoft Excel.",
                   "Step 5 uses only generated code and a fresh native Excel oracle; it does not calculate with formula caches.",
                   "Only the approved saved scenario and planned comparison range are in scope."]}
