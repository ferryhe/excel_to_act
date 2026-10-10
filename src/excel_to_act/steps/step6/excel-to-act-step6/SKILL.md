---
name: excel-to-act-step6
description: "Use when creating, finalizing, or reviewing the final human-readable Step 6 model-conversion report from accepted Excel-to-Act workflow evidence; not for general report writing."
---

# Excel-to-Act Step 6

Use this skill when a user asks for the final model-conversion handover report for an Excel-to-Act workflow, or asks you to review that report. The canonical CLI renderer produces the JSON evidence and Markdown report as a bound pair; this skill supplies the reporting and review scaffold, not a second schema or renderer.

## Use the workflow evidence

Read `excel-to-act step6 agent`, `excel-to-act step6 tools`, and `excel-to-act workflow status --workflow DIR` before acting. Inspect the current workflow's reviewer configuration and active delegation record; follow its exact stage scope, reviewer pair, authorization, and threshold. Use the accepted Stage 1–5 artifacts and current freshness state. Do not assume a delegation exists or carry a reviewer pair, threshold, scenario, target, input count, worksheet, or approval over from another workflow.

When report creation is requested and the current workflow permits it, use `excel-to-act step6 report --workflow DIR`. It writes a new report revision; it does not approve or reject that revision. Do not regenerate reports or record decisions merely because this skill is being installed, read, or used as a reference. This skill does not grant approval, delegation, or external-action authority.

Use `excel-to-act step6 template` to print the Markdown scaffold when drafting or checking report coverage. Treat the template as a writing aid. Keep the canonical renderer's JSON/Markdown pair and schema authoritative; do not replace their evidence with a manually completed template.

## Report and review standard

Make the report a concise English handover that leads with measured requested results and separately labels diagnostic intermediates or additional checks. Explain intended use and scope, accepted scalar/vector/table inputs and their source locations, keyed axes and time alignment, field groups, module order, recurrences, and the result equation when those facts are present in accepted evidence.

Explain the conversion approach and the actual recorded validation method, comparison coverage, denominators, and tolerances. Include only route evidence supported by the accepted generation and matching runtime comparison. Keep inactive condition checks separate from implemented numerical routes and any solver capability. Distinguish design-time questions and options from measured execution results.

State material extraction limits, comparison differences, and unresolved business interpretation in plain terms. Keep unknowns unknown. Do not infer missing source facts, resolve business meaning from numerical agreement, broaden saved-case evidence to other scenarios, or claim actuarial certification, professional-standard compliance, production deployment, or unsupported performance.

Use input-catalog facts only after the accepted source binding and every recorded catalog digest match. If a required source digest or artifact is absent, invalid, or stale, report those facts as unavailable. Preserve source values, identifiers, formulas, and provenance exactly; write new instructions and narrative in English.

Include practical bundle and run links, then point readers to the current workflow acceptance ledger for later Agent, TypeSafe, or human decisions. The report is generated ready for review and remains immutable after decisions are recorded; do not edit it to add approval state. Follow the current Step 6 review contract for any actual decision or return/revision.

For the fuller section checklist and example table structure, read [the conversion report scaffold](assets/conversion_report.md). The CLI renderer remains the source of the final report structure and evidence.
