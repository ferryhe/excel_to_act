from __future__ import annotations

from typing import Any


def tool_catalog() -> dict[str, Any]:
    return {"schema_version": "step4.tools.v1", "tools": [
        {"name": "step4.agent", "command": "step4 agent", "outputs": ["Stage 4 generation contract"]},
        {"name": "step4.tools", "command": "step4 tools", "outputs": ["Stage 4 commands and limits"]},
        {"name": "step4.capture_external", "command": "step4 capture-external --workflow DIR", "outputs": ["catalog-derived capture for confirmed external vectors and source-bound external_inputs.json/.md"]},
        {"name": "step4.discover", "command": "step4 discover --workflow DIR", "outputs": ["source-bound active-trace discovery for approved targets"]},
        {"name": "step4.plan", "command": "step4 plan --workflow DIR --implementation INPUT.json", "outputs": ["fresh-trace projection", "compiler/schedule preflight", "implementation_plan.json/.md"]},
        {"name": "step4.generate", "command": "step4 generate --workflow DIR [--trace TRACE.json]", "outputs": ["standalone Python bundle", "bounded smoke evidence", "generation_report.json/.md"]},
        {"name": "workflow.typesafe", "command": "workflow typesafe --workflow DIR --stage 4 --request REQUEST.json", "outputs": ["live, current-artifact-bound TypeSafe review evidence"]},
        {"name": "workflow.confirm", "command": "workflow confirm --workflow DIR --stage 4 --reviewer agent|human|typesafe --decision approve|reject --message TEXT [--evidence REVIEW.json]", "outputs": ["review receipt"]},
        {"name": "workflow.reject", "command": "workflow reject --workflow DIR --stage 4 --reviewer agent|human|typesafe --return-to N --message TEXT [--evidence REVIEW.json]", "outputs": ["rejection receipt"]},
        {"name": "workflow.status", "command": "workflow status --workflow DIR", "outputs": ["freshness, approvals, and next permitted stage"]},
    ], "limits": ["External capture, discovery, implementation planning, and generation require current Stage 3 approval.",
                   "Modular generation requires a passing current implementation preflight; preflight does not approve or advance Stage 4.",
                   "External capture derives vector IDs, finite ranges, shapes, and the shared key axis from the confirmed catalog and semantic cut; callers cannot override ranges.",
                   "The current GP design has five external 106-value rate vectors on one 106-key age axis (530 captured coordinates). Their upstream CI producer formulas remain excluded; the values are recaptured in native Excel for each Stage 5 comparison.",
                   "The current primary-input reconciliation expects seven approved Main cells. Another workbook needs its own reviewed catalog, semantic plan, and Stage 4 adapter; this workflow is not a general workbook converter.",
                   "Discovery uses formula text and literal source values; it does not call native Excel or read formula caches.",
                   "The captured vectors are a separate formula-derived external input ledger, never raw source values or cache values.",
                   "Only targets and scenarios in the approved Stage 3 design are discovered and emitted.",
                   "Generated code uses standard-library helpers and separate raw and external input ledgers.",
                   "Unsupported active formulas, unresolved references, cycles, or unbound inputs block generation.",
                   "This is not a general Excel compiler or GPU implementation."]}
