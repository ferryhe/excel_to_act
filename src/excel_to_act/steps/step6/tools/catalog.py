from __future__ import annotations

from typing import Any


def tool_catalog() -> dict[str, Any]:
    return {"schema_version": "step6.tools.v1", "tools": [
        {"name": "step6.agent", "command": "step6 agent", "outputs": ["Stage 6 review contract"]},
          {"name": "step6.skill", "command": "step6 skill", "outputs": ["Packaged Step 6 report skill"]},
          {"name": "step6.template", "command": "step6 template", "outputs": ["Human-readable report scaffold"]},
        {"name": "step6.tools", "command": "step6 tools", "outputs": ["Stage 6 commands and limits"]},
        {"name": "step6.report", "command": "step6 report --workflow DIR", "outputs": ["conversion_report.json/.md"]},
        {"name": "workflow.typesafe", "command": "workflow typesafe --workflow DIR --stage 6 --request REQUEST.json", "outputs": ["live, current-artifact-bound TypeSafe review evidence"]},
        {"name": "workflow.confirm", "command": "workflow confirm --workflow DIR --stage 6 --reviewer agent|human|typesafe --decision approve --message TEXT [--evidence REVIEW.json]", "outputs": ["approval receipt; use workflow.reject to reject"]},
        {"name": "workflow.reject", "command": "workflow reject --workflow DIR --stage 6 --reviewer agent|human|typesafe --return-to 1|2|3|4|5|6 --message TEXT [--evidence REVIEW.json]", "outputs": ["rejection receipt routed to the current or an earlier stage"]},
        {"name": "workflow.status", "command": "workflow status --workflow DIR", "outputs": ["freshness, approvals, and next permitted stage"]},
    ], "limits": ["Stage 6 requires confirmed Steps 1–5.",
                   "The final report summarizes only evidence in the accepted artifacts.",
                   "A saved-scenario result does not establish all-configuration or full-workbook closure."]}
