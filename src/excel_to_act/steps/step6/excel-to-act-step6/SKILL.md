---
name: excel-to-act-step6
description: "Prepare, review, deliver, or revoke the exact Step 6 internal tool package in an Excel-to-Act seven-stage conversion."
---

# Excel-to-Act Step 6

Read `excel-to-act step6 agent`, `excel-to-act step6 tools`, and `excel-to-act workflow status --workflow DIR`. Confirm that Stages 1–5 are current and approved and the comparison matches the exact request. For an approved exact-case Step 5 report, use that report as the comparison for a covered request; other comparisons require an independently verified binding. Attachment accepts a production parent with manifest schema `conversion.workflow.v1`, stage keys `1`–`6`, and approved Stages 1–5.

Use `excel-to-act step6 candidate --workflow DIR --version VERSION --request REQUEST.json --comparison COMPARISON.json`. Review its immutable exact JSON/Markdown pair, package inventory, interface, dependencies, fixed assets, internal version and permission, and isolated smoke. A passing smoke is not acceptance.

## Review and deliver

Record Agent review, then obtain an actual human decision on the exact pair. TypeSafe delegation and live review commands are disabled. Only after current approval, use `excel-to-act step6 deliver --workflow DIR` to deliver the same package and verify exact delivery/readback. A failed readback cannot leave an effective delivery. Check `workflow status` for the effective internal version.

Use `workflow reject --workflow DIR --stage 6 --reviewer agent|human --return-to STAGE --message TEXT` for rejection. Native workflows can return to Stages 1–6. Attached workflows have read-only inherited Stages 1–5 and can return only to Stage 6; correct parent evidence through a separate attachment. Preserve package files and review records.

Use `excel-to-act step6 revoke --workflow DIR --reason TEXT` to deactivate an internal delivery while retaining its package and audit trail.

Use Step 7 (`excel-to-act step7 report --workflow DIR`) for the final report after Step 6 remains approved and exact delivery/readback is valid. Report creation requires its own Agent and human review. This skill provides internal delivery only; reading or installing it grants no approval or external-action authority.
