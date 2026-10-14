from __future__ import annotations

from typing import Any


def tool_catalog() -> dict[str, Any]:
    return {"schema_version": "step7.tools.v1", "tools": [
        {"name": "step7.agent", "command": "step7 agent", "outputs": ["Stage 7 review contract"]},
        {"name": "step7.skill", "command": "step7 skill", "outputs": ["Packaged Step 7 report skill"]},
        {"name": "step7.template", "command": "step7 template", "outputs": ["Human-readable report scaffold"]},
        {"name": "step7.tools", "command": "step7 tools", "outputs": ["Stage 7 commands and limits"]},
        {"name": "step7.report", "command": "step7 report --workflow DIR", "outputs": ["conversion_report.json/.md"]},
        {"name": "workflow.confirm", "command": "workflow confirm --workflow DIR --stage 7 --reviewer agent|human --decision approve --message TEXT", "outputs": ["current-pair review receipt"]},
        {"name": "workflow.reject", "command": "workflow reject --workflow DIR --stage 7 --reviewer agent|human --return-to 1|2|3|4|5|6|7 --message TEXT",
         "attached_command": "workflow reject --workflow DIR --stage 7 --reviewer agent|human --return-to 6|7 --message TEXT",
         "availability": "native workflows may return to Stages 1–7; attached workflows may return only to Stages 6–7; inherited stages are read-only",
         "outputs": ["rejection receipt"]},
        {"name": "workflow.status", "command": "workflow status --workflow DIR", "outputs": ["freshness, approvals, delivery and next permitted stage"]},
    ], "limits": ["Stage 7 requires current approved Stages 1–6 and exact delivery readback.",
                   "The final report summarizes only evidence in accepted artifacts.",
                   "Creating the report never approves it or activates production."]}
