# Step 5 Agent Contract: Validate and Reconcile

Step 5 first checks that the approved Stage 4 bundle compiles and runs in isolation. It separately captures a fresh native Excel result and reconciles the generated outputs with that result. The oracle command does not advance the stage. Stage 5 is reviewable only after standalone validation and reconciliation both pass.

Write new instructions, report narration, and generated artifact prose in English. Preserve source values, formulas, identifiers, and immutable provenance verbatim. Only the repository-root `README.md` and `README.zh-CN.md` are approved bilingual documentation exceptions.

Read `step5 tools` and the accepted Stage 4 report. Run `step5 validate --workflow DIR`; this compiles and executes the generated bundle in a temporary isolated directory without the source workbook or project package. Retain its JSON/Markdown evidence and model result. Run `step5 oracle --workflow DIR` on Windows with Microsoft Excel when available. The native tool disables macros and events, opens a private read-only copy, performs a full rebuild, captures requested names/ranges, and closes without saving. It records evidence only and never creates approval.

Derive required targets, source ranges, inputs, external boundaries, and coverage denominators from the current accepted design and bundle manifest. Inspect any workbook-specific assumptions in the capture and reconciliation adapters; do not reuse another source's sheet names, coordinates, counts, or expected results as defaults. If the available adapter cannot cover this design, record the gap and return to the appropriate design or implementation stage.

Then run `step5 reconcile --workflow DIR --validation VALIDATION.json --oracle ORACLE.json`. Check the source and Stage 4 bindings, primary inputs, named targets, active formula identities and values, declared external values and keys, errors, and stated tolerances. Require the native calculation to be complete. Do not treat missing or blocked comparisons as zero. A mismatch or incomplete comparison is not ready for review; do not approve it. Do not alter the source workbook or earlier artifacts.

After both required passes, review `validation_report.json/.md` and record the Agent decision against the current report pair. By default, the human decision must follow an explicit user response. If this workflow has an active TypeSafe delegation, submit a current-bound request with `workflow typesafe --workflow DIR --stage 5 --request REQUEST.json`, then confirm using its generated evidence file; hand-written packets do not count as live reviews. If a result changes, rerun the affected validation and reconciliation actions and obtain fresh confirmations.

A new explicit actual-human approval under a valid active TypeSafe delegation restores the required pair to Agent+human only for that exact current, unrejected JSON/Markdown pair and revision; it does not change the authorization or other stages. A rejection from any reviewer makes that revision ineligible for later approvals; create a replacement artifact and obtain new decisions.

## Shared workflow and decision order

Treat `input/` as the source Excel area and `output/` as the home for each new source-bound workflow and its reports. Start a new conversion from the current workbook and user request; never carry case facts or approvals from another workbook or run.

Read `stepN tools` and choose the existing actions that answer this case's open questions. Explore from the source and current bound evidence, inspect each action's checks and outputs, then prepare the paired machine-readable JSON and reader-facing Markdown handoff. Review all files named in the handoff and confirm their source and upstream bindings before asking for a decision. Do not continue when a required check fails or an input is stale.

Before entering a stage, run `workflow status` and verify that the stage is currently permitted and every required upstream decision and artifact is current. A successful tool check, report creation, draft CLI action, or clarification does not itself approve a stage or advance the workflow. Keep legacy standalone source-exploration APIs available for discovery; they do not create a workflow approval.

The default review order is Agent first, then the actual human. Record the Agent decision against the exact current JSON/Markdown pair with `workflow confirm`; present that pair and its practical limits to the user; record a human approval only after an explicit response. New TypeSafe approvals also require a current matching Agent approval. Use TypeSafe only under a valid, explicitly scoped authorization already registered for this workflow; never infer or transfer delegation. The Step 3 input-boundary checkpoint always requires an actual human decision.

Ask the user when target, scope, input kind, axis, units, or business interpretation is materially unresolved. Keep unknown facts marked as unknown; do not turn planning into measured execution or claim checks that did not run.

When a reviewer rejects a report, route it to the current or an earlier responsible stage with `workflow reject --return-to`. Keep the rejected artifact and decision in the ledger. Revise affected reports, rerun their checks, refresh later reports that depend on them, and obtain fresh Agent and human or explicitly delegated reviews before advancing. Inspect the ledger with `workflow status` after each decision and before each transition.
