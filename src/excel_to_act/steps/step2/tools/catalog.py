from __future__ import annotations

from typing import Any

from excel_to_act.steps.conversion_workflow import review_tool_entries


def tool_catalog() -> dict[str, Any]:
    from excel_to_act.steps.step2.workflow import _TOOLS
    """Return the initial machine-readable Step 2 actions."""
    catalog_tools = [*_TOOLS, *review_tool_entries(2)]
    return {
        "tool": "step2.tools",
        "status": "ok",
        "source": None,
        "run_id": None,
        "artifacts": [],
        "metrics": {"tool_count": len(catalog_tools)},
        "diagnostics": [],
        "retryable": False,
        "next_tool": None,
        "tools": catalog_tools,
        "review_workflow": [
            "workflow status --workflow DIR",
            "workflow confirm --workflow DIR --stage 2 --reviewer agent|human --decision approve|reject --message TEXT [--return-to N]",
            "workflow reject --workflow DIR --stage 2 --reviewer agent|human --return-to 1|2 --message TEXT",
            "The human receipt may be recorded only after an explicit user response; do not infer or fabricate approval.",
        ],
    }
