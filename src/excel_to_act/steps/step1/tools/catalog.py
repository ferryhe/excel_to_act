from __future__ import annotations

from typing import Any

from excel_to_act.steps.conversion_workflow import review_tool_entries


def tool_catalog() -> dict[str, Any]:
    from excel_to_act.steps.step1.workflow import _TOOLS, _coverage_scope
    catalog_tools = [*_TOOLS, *review_tool_entries(1)]
    return {
        "tool": "step1.tools",
        "status": "ok",
        "source": None,
        "run_id": None,
        "coverage_scope": _coverage_scope(),
        "standalone_control_commands": [
            {
                "family": family,
                "identify": f"step1 {family} identify WORKBOOK",
                "convert": f"step1 {family} convert WORKBOOK --out DIR [--dry-run]",
                "evaluate": f"step1 {family} evaluate WORKBOOK --out DIR",
                "behavior": "Static source inspection and handoff; macros are never executed. Dry run creates and changes no output.",
            }
            for family in ("checkbox", "activex", "vba")
        ],
        "metrics_contract": {
            "state_field": "metrics_state",
            "states": {"measured": "A fresh source scan supplied the current denominators and ratios.", "unavailable": "A fresh source scan could not be validated; source-derived totals, counts, ratios, and deviations are null."},
            "zero_denominator": "Numeric zero is used only when a fresh scan measured an empty denominator; empty corpora still fail.",
            "gate": "Unavailable metrics never satisfy thresholds, including thresholds set to zero; source/read blockers prevent finalization.",
        },
        "artifacts": [],
        "metrics": {"tool_count": len(catalog_tools)},
        "diagnostics": [],
        "retryable": False,
        "next_tool": None,
        "tools": catalog_tools,
        "review_workflow": [
            "workflow status --workflow DIR",
            "workflow confirm --workflow DIR --stage 1 --reviewer agent|human --decision approve --message TEXT",
            "workflow reject --workflow DIR --stage 1 --reviewer agent|human --return-to 1 --message TEXT",
            "The human receipt may be recorded only after an explicit user response; do not infer or fabricate approval.",
        ],
    }
