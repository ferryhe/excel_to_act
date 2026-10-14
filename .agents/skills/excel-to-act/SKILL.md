---
name: excel-to-act
description: "Convert .xlsx/.xlsm models into standalone Python tools with the Excel-to-Act seven-stage workflow, or continue, validate, internally deliver, or report on an existing conversion. Not for general spreadsheet editing or development of this repository."
---

# Excel to Act

Use the installed `excel-to-act` CLI as the producer and acceptance ledger. This skill coordinates its existing tools; it does not supply a second compiler, report renderer, or approval system.

## Check the runtime and source

Python 3.11+ is required. Check `python -m pip show excel-to-act` and `excel-to-act --help`. A distributed skill includes its matching wheel in `runtime/` beside this file; install that wheel into the selected Python environment if the runtime is missing or differs. In a source checkout, install that checkout with `python -m pip install -e .`. Do not substitute an unrelated installed version or fetch an unpinned latest version. The release's `release.json` identifies the version and file hashes. Optional VBA inspection uses the `vba` extra; native Excel capture additionally needs Windows and Microsoft Excel. This workflow needs no OpenAI API key.

Keep source workbooks under a user-selected input directory and each conversion's evidence in a separate output directory. Do not overwrite the original workbook. Ask for materially unresolved targets, input boundaries, axes, units, or business meaning before depending on them. Source facts and approvals belong to this conversion; do not copy them from examples or earlier workbooks.

For an existing workflow, first run `excel-to-act workflow status --workflow DIR`. Read the current stage's `excel-to-act stepN agent` and `excel-to-act stepN tools`; these packaged contracts define the commands, supported formula/input shapes, prerequisite artifacts, and current review route. The CLI's current source bindings and limits take precedence over a remembered workflow.

## Route the requested work

| Stage | Tools and handoff |
| --- | --- |
| 1. Extract | `step1 convert INPUT_DIR --out OUTPUT_DIR`, `check`, and `finalize`; resolve the returned candidate `run_path` and final `final_output` relative to that output root. `step1 report --run FINAL_RUN --out NEW_WORKFLOW` starts a native seven-stage workflow. |
| 2. Index | `step2 index`, `validate`, and `report`; preserve source identity and saved artifact references. |
| 3. Confirm and design | Explore with `prepare`, `fields`, `dependencies`, bounded `query`/`trace`, `source-trace`, `profile`, `plan`, and `check`. Review `input-catalog` separately before the final `report`; export the accepted `interface` when required. |
| 4. Generate | Follow the accepted capture plan with `capture-external` when needed, then `discover`, `plan --implementation FILE`, and `generate`. Unsupported or unbound designs return to Step 3. |
| 5. Reconcile | `validate` runs the standalone bundle; `oracle` captures a fresh native Excel result; `reconcile` compares matching evidence. A smoke run is not numerical equivalence. |
| 6. Internal delivery | After current upstream acceptance and a separately verified comparison binding, prepare `candidate`. Deliver the identical package with `deliver` only after its current reviews; check effective delivery and readback. `revoke` retains history. |
| 7. Report | After current Step 6 acceptance and valid delivery/readback, use `step7 report`. `step7 skill` and `step7 template` expose the maintained reporting scaffold; keep the canonical JSON/Markdown pair. |

For the first source checkpoint, pass the promoted directory returned by `step1 finalize` as `final_output`, resolved under the Step 1 output root. That directory contains `promotion.json`; the original candidate `run_path` is not a valid `step1 report --run` input.

For the implemented `exact_cases` route, use `step3 produce-exact-cases` to verify already captured native evidence and calculate one source trace per request. It does not call Excel or create an approval. Each case needs its own profile and semantic plan/check. In Step 5, use `validate --inputs REQUEST`, then `oracle --validation THAT_VALIDATION`, then `reconcile` for that same request. Only `finalize-cases` with all 32 distinct passing approved case receipts produces the formal Step 5 checkpoint. Do not replace this with one saved-scenario pass or share another request's external data or comparison.

For migration, `workflow attach` creates a separate seven-stage root pinned to an approved legacy parent and independently verified comparison digest. Legacy v1 and attached inherited Stages 1–5 are read-only. Their old Step 6 report or delegation is not a new internal-delivery approval.

## Preserve acceptance and evidence

- A technical `pass`, generated report, or successful package invocation is evidence for review, not approval. Read `workflow status` before transitions.
- Review the exact current JSON/Markdown pair as Agent first. Record a human receipt only for an explicit human response on that same pair. The input boundary always needs an actual human decision; current v2 also requires actual-human decisions at its other stage gates. Do not synthesize, infer, or reuse approvals, including from test fixtures.
- Use `workflow confirm` and `workflow reject --return-to STAGE` for decisions. Retain rejected revisions and historical packages. Source, artifact, input, or upstream revisions can invalidate downstream acceptance; refresh affected evidence rather than editing an approved report.
- Formula caches are never model inputs, and extraction does not execute VBA. Keep declared raw inputs, formula-derived external data, source metadata, and calculation results distinct.
- Describe only measured scenario coverage, errors, tolerances, and limits. Do not infer business correctness, all-configuration support, a feedback solver, GPU support, or actuarial certification from numerical agreement.
- This skill authorizes no public upload, production activation, external distribution, or release. Step 6 concerns the separately approved internal model package. Installing or preparing this skill must not start a conversion, write a workflow decision, or deliver a model.
