from __future__ import annotations

from typing import Any


def tool_catalog() -> dict[str, Any]:
    return {"schema_version": "step6.tools.v2", "tools": [
        {"name": "step6.agent", "command": "step6 agent", "outputs": ["Stage 6 delivery contract"]},
        {"name": "step6.tools", "command": "step6 tools", "outputs": ["Stage 6 commands"]},
        {"name": "workflow.attach", "command": "workflow attach --parent PARENT_DIR --out NEW_DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256",
         "outputs": ["separate seven-stage workflow pinned to approved parent Stages 1–5 and the separately verified comparison"]},
        {"name": "workflow.bind-comparison", "command": "workflow bind-comparison --workflow DIR --comparison COMPARISON.json --comparison-sha256 APPROVED_SHA256",
         "outputs": ["separately verified exact comparison bound after native Stage 5 approval"]},
        {"name": "step6.candidate", "command": "step6 candidate --workflow DIR --version VERSION --request REQUEST.json --comparison COMPARISON.json",
         "native_exact_cases_comparison": "current approved Step 5 validation_report.json",
         "outputs": ["immutable package candidate and exact JSON/Markdown review pair"]},
        {"name": "workflow.confirm", "command": "workflow confirm --workflow DIR --stage 6 --reviewer agent|human --decision approve --message TEXT",
         "outputs": ["current-pair review receipt; actual human response required"]},
        {"name": "workflow.reject", "command": "workflow reject --workflow DIR --stage 6 --reviewer agent|human --return-to 1|2|3|4|5|6 --message TEXT",
         "attached_command": "workflow reject --workflow DIR --stage 6 --reviewer agent|human --return-to 6 --message TEXT",
         "availability": "native workflows may return to Stages 1–6; attached workflows may return only to Stage 6; inherited stages are read-only",
         "outputs": ["rejection receipt and invalidated downstream decisions"]},
        {"name": "step6.deliver", "command": "step6 deliver --workflow DIR",
         "outputs": ["effective internal delivery only after genuine acceptance and exact readback"]},
        {"name": "step6.revoke", "command": "step6 revoke --workflow DIR --reason TEXT",
         "outputs": ["inactive current registration; immutable history retained"]},
        {"name": "workflow.status", "command": "workflow status --workflow DIR",
         "outputs": ["current approvals, delivery readback, next permitted stage"]},
    ], "limits": ["Step 6 requires current production Stages 1–5 and exact scenario comparison.",
                   "Candidate smoke is not publication acceptance.",
                   "Only internal delivery is supported; supported-domain calls do not need new approval."]}
