---
name: excel-to-act-step6
description: "Use for current Step 6 internal tool delivery; historical v1 Step 6 reports are read-only."
---

# Excel-to-Act Step 6

For a current seven-stage workflow, read `excel-to-act step6 agent`, `excel-to-act step6 tools`, and `excel-to-act workflow status --workflow DIR`. After current approved Stages 1–5 and a separately bound exact comparison, use `excel-to-act step6 candidate --workflow DIR --version VERSION --request REQUEST.json --comparison COMPARISON.json`. Review its immutable exact JSON/Markdown pair, package inventory, interface, dependencies, fixed assets, internal version and permission, and isolated smoke. A passing smoke is not acceptance.

Record Agent review, then obtain an actual human decision on the exact pair. Old Step 6 approvals or TypeSafe delegation cannot authorize this delivery. Only after genuine current approval, use `excel-to-act step6 deliver --workflow DIR` to deliver the same package and verify exact delivery/readback. A failed readback cannot leave an effective new delivery. Use `excel-to-act step7 report --workflow DIR` for the current final report only after Step 6 remains approved and that exact delivery/readback is valid; report creation is not approval. This route is internal only.

## Historical v1 report reference (read-only)

Existing six-stage Step 6 final reports remain readable with their original meaning. The current CLI is read-only for v1, and its `step6 report` command rejects both v1 and v2. The following report standard is retained for interpreting historical reports; use the copied `excel-to-act-step7` skill for current final reports.

Historical reports used the canonical CLI renderer's bound JSON evidence and Markdown report. This retained reporting and review scaffold is not a second schema or renderer.

## Use the workflow evidence

When reading a saved v1 report, inspect its recorded reviewer configuration and delegation with their original stage scope, reviewer pair, authorization, and threshold. Use its accepted Stage 1–5 artifacts and recorded freshness state. Do not carry a reviewer pair, threshold, scenario, target, input count, worksheet, or approval into v2.

The historical `excel-to-act step6 report --workflow DIR` created a report revision without approving it. That write route is now disabled. Do not regenerate a historical report or record decisions merely because this skill is read or used as a reference. This skill grants no approval, delegation, or external-action authority.

The retained `assets/conversion_report.md` is the historical Markdown scaffold. Use `excel-to-act step7 template` for current report coverage. Keep the canonical renderer's JSON/Markdown pair and schema authoritative; do not replace their evidence with a manually completed template.

## Report and review standard

Make the report a concise English handover that leads with measured requested results and separately labels diagnostic intermediates or additional checks. Explain intended use and scope, accepted scalar/vector/table inputs and their source locations, keyed axes and time alignment, field groups, module order, recurrences, and the result equation when those facts are present in accepted evidence.

Explain the conversion approach and the actual recorded validation method, comparison coverage, denominators, and tolerances. Include only route evidence supported by the accepted generation and matching runtime comparison. Keep inactive condition checks separate from implemented numerical routes and any solver capability. Distinguish design-time questions and options from measured execution results.

State material extraction limits, comparison differences, and unresolved business interpretation in plain terms. Keep unknowns unknown. Do not infer missing source facts, resolve business meaning from numerical agreement, broaden saved-case evidence to other scenarios, or claim actuarial certification, professional-standard compliance, production deployment, or unsupported performance.

Use input-catalog facts only after the accepted source binding and every recorded catalog digest match. If a required source digest or artifact is absent, invalid, or stale, report those facts as unavailable. Preserve source values, identifiers, formulas, and provenance exactly; write new instructions and narrative in English.

Historical reports include practical bundle and run links and point readers to their acceptance ledger for later Agent, TypeSafe, or human decisions. A saved report remains immutable after decisions are recorded; do not edit it to add approval state. Current Step 7 decisions follow the current Agent and actual-human review contract.

For the fuller section checklist and example table structure, read [the conversion report scaffold](assets/conversion_report.md). The CLI renderer remains the source of the final report structure and evidence.
