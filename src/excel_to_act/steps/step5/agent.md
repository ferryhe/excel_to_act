# Step 5 Agent Contract: Validate and Reconcile

A new native v2 workflow has writable Stage 1–5 checkpoints and requires current Agent and actual-human review. Legacy v1 checkpoints and attached-v2 inherited Stage 1–5 evidence are read-only. Legacy TypeSafe delegation does not authorize current v2 reviews.

Step 5 first checks that the approved Stage 4 bundle compiles and runs in isolation. It separately captures a fresh native Excel result and reconciles the generated outputs with that result. The oracle command does not advance the stage. Stage 5 is reviewable only after standalone validation and reconciliation both pass.

Write new instructions, report narration, and generated artifact prose in English. Preserve source values, formulas, identifiers, and immutable provenance verbatim. Only the repository-root `README.md` and `README.zh-CN.md` are approved bilingual documentation exceptions.

Read `step5 tools` and the accepted Stage 4 report. Run `step5 validate --workflow DIR` for the saved case or add `--inputs REQUEST.json` for a guarded variable case; this compiles and executes the generated bundle in a temporary isolated directory without the source workbook or project package. Retain its JSON/Markdown evidence and model result. Run `step5 oracle --workflow DIR` on Windows with Microsoft Excel when available; add `--validation VARIABLE_VALIDATION.json` for the exact variable case. The native tool disables macros and events, opens a private copy, applies approved raw literal inputs in memory, checks readback, performs a full rebuild, captures requested names/ranges, and closes without saving. It records evidence only and never creates approval.

Optional `--target` and `--range` captures may include exploratory evidence outside the plan. Capture does not expand the approved comparison scope or advance Stage 5. Reconciliation compares only the bound design targets and active formula addresses present in both the native and generated results. Review the captured range selectors separately from the exact `compared.comparison_addresses`; identify selectors with no compared active formula address as captured-only for active-formula comparison. Named targets, primary inputs, and external value, formula-cut, and age-axis checks are reported separately.

Derive required targets, source ranges, inputs, external boundaries, and coverage denominators from the current accepted design and bundle manifest. Inspect any workbook-specific assumptions in the capture and reconciliation adapters; do not reuse another source's sheet names, coordinates, counts, or expected results as defaults. If the available adapter cannot cover this design, record the gap and return to the appropriate design or implementation stage.

Then run `step5 reconcile --workflow DIR --validation VALIDATION.json --oracle ORACLE.json`. Check the source and Stage 4 bindings, primary inputs, named targets, active formula identities and values, declared external values and keys, errors, and stated tolerances. Require the native calculation to be complete. Do not treat missing or blocked comparisons as zero. A mismatch or incomplete comparison is not ready for review; do not approve it. Do not alter the source workbook or earlier artifacts.

For a current exact-case variable contract, reconcile each approved case against that case's own source-bound formula trace. Each call writes a case receipt without advancing Stage 5. After all 32 distinct approved cases pass, run `step5 finalize-cases --workflow DIR --case-report FILE` with `--case-report` repeated once per case. Finalization rechecks current request, trace, package, native oracle, source, and tolerance bindings before writing one formal Stage 5 JSON/Markdown pair. A saved-scenario receipt or fewer than all approved cases cannot finalize the exact-case checkpoint.

After both technical passes, review `validation_report.json/.md` and record Agent approval of the current pair. Record actual-human approval only after an explicit response. The current v2 CLI disables legacy TypeSafe delegation and live TypeSafe review. If a result changes, rerun affected validation and reconciliation and obtain fresh decisions.

A rejection from any reviewer makes that revision ineligible for later approvals; create a replacement artifact and obtain fresh Agent and actual-human decisions.

## Shared workflow and decision order

Treat `input/` as the source Excel area and `output/` as the home for each new source-bound workflow and its reports. Start a new conversion from the current workbook and user request; never carry case facts or approvals from another workbook or run.

Read `stepN tools` and choose the existing actions that answer this case's open questions. Explore from the source and current bound evidence, inspect each action's checks and outputs, then prepare the paired machine-readable JSON and reader-facing Markdown handoff. Review all files named in the handoff and confirm their source and upstream bindings before asking for a decision. Do not continue when a required check fails or an input is stale.

Before entering a stage, run `workflow status` and verify that the stage is currently permitted and every required upstream decision and artifact is current. A successful tool check, report creation, draft CLI action, or clarification does not itself approve a stage or advance the workflow. Keep legacy standalone source-exploration APIs available for discovery; they do not create a workflow approval.

The current review order is Agent first, then the actual human on the exact current JSON/Markdown pair. `workflow confirm` records approvals; use `workflow reject` with an explicit `--return-to` route for rejection. Record the human decision only after an explicit response. Legacy TypeSafe delegation and live review commands are disabled in v2. The Step 3 input-boundary checkpoint always requires an actual human decision.

Ask the user when target, scope, input kind, axis, units, or business interpretation is materially unresolved. Keep unknown facts marked as unknown; do not turn planning into measured execution or claim checks that did not run.

When a reviewer rejects a report at Stage 5, route it to the responsible current or earlier stage with `workflow reject --workflow DIR --stage 5 --reviewer agent|human --return-to 1|2|3|4|5 --message TEXT`. Keep rejected history, revise affected reports, rerun checks, and obtain fresh Agent and actual-human reviews before advancing. Inspect `workflow status` after each decision.
