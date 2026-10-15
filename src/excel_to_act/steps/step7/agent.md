# Step 7 Agent Contract: Final Conversion Report

Step 7 summarizes the accepted source, index, design, generated implementation, validation/reconciliation, and internal delivery records. It does not calculate formulas or upgrade a saved-scenario result into all-configuration coverage.

For the reusable final-report writing and review scaffold, see the [Excel-to-Act Step 7 skill](excel-to-act-step7/SKILL.md), or print it with `excel-to-act step7 skill`. Print its Markdown scaffold with `excel-to-act step7 template`. The CLI renderer and its JSON evidence remain authoritative for generated reports.

The Markdown report is the human handover. Start with the result table in accepted target order and the scoped conclusion, then review purpose and scope, source inputs and assumptions, calculation order, conversion approach, validation, limitations, handover instructions, and acceptance state. It should use accepted target names, ordering, units or basis, scenario literals, input-catalog totals and source-tab counts, field-group shapes and time alignment, recurrence notes, execution order, and shared-module responsibilities. The module flow summarizes accepted design data; it is not implementation proof by itself.

If the accepted design links an input-boundary catalog, confirm that the recorded digest matches and that Step 7 binds the catalog files in its own input-file list. If the catalog is missing or stale, the report must say those catalog facts are unavailable. Do not infer missing source counts, axes, inputs, or business meanings.

Keep current result values and scoped conclusions in the report body. Explain the recorded standalone-Python and native-workbook validation methods, comparison denominators, partial source extraction, and actual unresolved business interpretations when those facts are present. Do not put raw JSON/list dumps or SHA strings in the main body. Detailed input literals and locations, hashes, review receipts, design options, preserved formula notes, and design questions belong in appendices. Preserve question text and dispositions in JSON; capture and native-comparison evidence may answer an execution prerequisite while preserving the question record. State execution conclusions from measured evidence.

Include only Step 5 route IDs supported by matching native comparisons. Keep condition-only checks separate from numerical paths; an inactive feedback condition does not mean a feedback solver exists. This handover does not claim actuarial certification or compliance with a professional standard.

Write new instructions, report narration, and generated artifact prose in English. Preserve source values, formulas, identifiers, and immutable provenance verbatim. Only the repository-root `README.md` and `README.zh-CN.md` are approved bilingual documentation exceptions.

Read `step7 tools` and confirm that Stages 1–6 are current and approved and Step 6 delivery readback is valid. Run `step7 report --workflow DIR`. Review `conversion_report.json/.md` for source-quality limits, Step 3 planning history, Step 4 generated path IDs, and only the Step 5 route IDs supported by native comparisons. Review the saved-scenario target values, code bundle/run command, exclusions, and open questions. Keep condition-only checks separate from numerical paths; an inactive feedback condition does not mean a feedback solver exists. A partial slice must not be called a complete general-purpose model; all-configuration coverage stays unknown unless supported by evidence.

Record the Agent decision only after checking the exact report pair and references. The report remains immutable after acceptance; read Agent and human decisions from the linked `workflow.json` acceptance ledger. The Step 7 human decision must follow an explicit user response. If attached upstream evidence changes, create a separate attachment to a current approved parent; do not rewrite the pinned parent. Native workflow revisions follow the return and reapproval route.

## Shared workflow and decision order

Treat `input/` as the source Excel area and `output/` as the home for each new source-bound workflow and its reports. Start a new conversion from the current workbook and user request; never carry case facts or approvals from another workbook or run.

Read `stepN tools` and choose the existing actions that answer this case's open questions. Explore from the source and current bound evidence, inspect each action's checks and outputs, then prepare the paired machine-readable JSON and reader-facing Markdown handoff. Review all files named in the handoff and confirm their source and upstream bindings before asking for a decision. Do not continue when a required check fails or an input is stale.

Before entering a stage, run `workflow status` and verify that the stage is currently permitted and every required upstream decision and artifact is current. A successful tool check, report creation, draft CLI action, or clarification does not itself approve a stage or advance the workflow. Use standalone source-exploration APIs for discovery; record workflow approvals through the source-bound ledger.

The current review order is Agent first, then the actual human. `workflow confirm` records approvals only; use the stage-specific `workflow.reject` command and its explicit `--return-to` route to reject. Record the Agent decision against the exact current JSON/Markdown pair with `workflow confirm`; present that pair and its practical limits to the user; record a human approval only after an explicit response. TypeSafe delegation and live TypeSafe review commands are disabled. The Step 3 input-boundary checkpoint always requires an actual human decision.

Ask the user when target, scope, input kind, axis, units, or business interpretation is materially unresolved. Keep unknown facts marked as unknown; do not turn planning into measured execution or claim checks that did not run.

When a reviewer rejects a report at Stage 7, route native workflows to the responsible Stage 1–7 with `workflow reject --workflow DIR --stage 7 --reviewer agent|human --return-to 1|2|3|4|5|6|7 --message TEXT`. Attached workflows can return only to Step 6 or 7 with `--return-to 6|7`; their inherited Stages 1–5 are read-only. Keep the rejected artifact and decision in the ledger. Revise affected reports, rerun their checks, refresh later reports that depend on them, and obtain fresh Agent and actual-human reviews before advancing. Inspect the ledger with `workflow status` after each decision and before each transition.
