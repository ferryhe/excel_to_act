from __future__ import annotations

from typing import Any

from excel_to_act.steps.conversion_workflow import review_tool_entries


def tool_catalog() -> dict[str, Any]:
    """Return Step 2 recovery tools and implemented reader commands."""
    from excel_to_act.steps.step2.workflow import _TOOLS
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
        "reader_commands": [{
            "name": "step2.prepare",
            "command": "step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]",
            "inputs": {"index": "Validated native Step 2 index", "step1_root": "Step 1 output root", "out": "Prepared reading package", "scope": "Optional source/run-bound analysis scope"},
            "outputs": ["manifest.json", "scope.json", "dependency_audit.json", "retained_dependency_cells.json", "paired English handoffs", "source views", "dependency graphs"],
            "next": ["step2.query"],
        }, {
            "name": "step2.query",
            "command": "step2 query --manifest MANIFEST --source-id SOURCE --kind KIND [--target TARGET] [--sheet SHEET] [--range A1_RANGE] [--budget N] [--cursor CURSOR] [--out PACKET]",
            "inputs": {"manifest": "#30 reading manifest", "source_id": "Exact source ID from the manifest", "kind": "overview, sheet, cell, range, name, control, vba, or feature", "target": "Selector-specific target", "sheet": "Worksheet context", "range": "A1 range", "budget": "Positive estimated-token page budget", "cursor": "Source/selector-bound continuation cursor"},
            "outputs": ["source-bound evidence packet", "canonical view references and complete delivered records", "selector summary, pagination, selected view/byte metrics, diagnostics"],
            "next": ["step2.trace", "views.validate"],
        }, {
            "name": "step2.trace",
            "command": "step2 trace --manifest MANIFEST --source-id SOURCE --kind cell|range|name --target TARGET [--sheet SHEET] --direction upstream|downstream|both [--max-depth N] [--max-nodes N] [--max-edges N] [--out PACKET]",
            "inputs": {"manifest": "Prepared reading manifest", "source_id": "Required exact source ID", "direction": "upstream dependencies follows formula-cell -> input; downstream reverses; both combines", "limits": "Defaults depth 8, nodes 100, edges 200; ceilings 100/10000/20000"},
            "outputs": ["canonical evidence packet with static trace", "source formula edges and provenance-marked structural derivations", "frontier, truncated, unresolved, scope boundaries, next actions and reuse counters"],
            "next": ["views.validate"],
        }],
        "review_workflow": [
            "workflow status --workflow DIR",
            "workflow confirm --workflow DIR --stage 2 --reviewer agent|human --decision approve --message TEXT",
            "workflow reject --workflow DIR --stage 2 --reviewer agent|human --return-to 1|2 --message TEXT",
            "The human receipt may be recorded only after an explicit user response; do not infer or fabricate approval.",
        ],
    }
