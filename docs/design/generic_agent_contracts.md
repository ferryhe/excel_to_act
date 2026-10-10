# Generic Agent contracts and source-specific evidence

The six packaged `agent.md` files define reusable actions, evidence requirements, deliverables and review gates. Each one is paired with a stage-local `tools/` catalog package that backs the CLI's `stepN tools` output; `agent.md` is available through `stepN agent`. The common flow starts with the source workbook in `input/`, writes a new source-bound workflow under `output/`, explores with tools selected for the current case, checks the evidence, and prepares paired JSON and human-readable Markdown handoffs. Each new conversion supplies its own source, target order, scenario, catalog, axes, equation families, external boundaries, coverage and implementation map. Do not place workbook facts or one user's case-specific choices in these instructions.

Before entering each stage, inspect `workflow status` and confirm that the stage is currently permitted and its upstream evidence and decisions are current. A draft command, report-creation success, or clarification does not itself approve or advance a stage. Keep the legacy standalone source-exploration APIs available; they help investigate a workbook without creating workflow approval.

The default review order is Agent first and actual human second on the exact current JSON/Markdown pair. A new human or TypeSafe approval cannot be recorded without a matching current Agent approval. Use TypeSafe only under an already registered, explicitly scoped authorization; a prior workflow does not authorize a new source or run. The Step 3 input-boundary checkpoint always requires actual human review. The Agent asks the user about material uncertainty in targets, scope, input kind, axes, units or business interpretation; a clarification is not a final stage decision. Rejections preserve history, return to the responsible current or earlier stage, and require revised affected reports plus fresh reviews.

| Information | Where it belongs |
| --- | --- |
| Workflow actions, provenance, error handling, stage deliverables and review gates | Packaged stage Agent contracts |
| Targets, source locations, selector values, input groups, vector lengths, table axes and exclusions | Current source-bound catalog and design artifacts |
| Source exceptions, selected feedback conditions and expected numerical results | Current case evidence and reports |
| Accepted scenario guards, source-to-variable mappings and executable equations | The generated bundle and its manifest |
| Actual formula subset, supported capture shapes and adapter restrictions | Tool capability contracts and implementation; assess compatibility for each new source |
| Illustrative source facts or manual acceptance instructions | A clearly named case study, not reusable Agent defaults |

## Corrections from the contract audit

All six Agent files were inspected. Source-specific prose was found in Steps 3 and 4. Steps 1, 2, 5 and 6 did not contain the current workbook's coordinates or numerical defaults; Step 5 received an explicit compatibility and coverage check.

| Earlier coupling | Reusable rule now |
| --- | --- |
| A particular deferred benefit and six selector coordinates | Record each excluded feedback condition from the current scenario; cite its selectors and dependent producer, preserve dependency evidence, and require reviewed redesign for activation |
| Named summary/projection cells | Partition the current source into logical scalars and series; keep a seed and recurrence segments in one variable |
| Fixed projection lengths and time-zero location | Derive starting/terminal keys and lengths independently for every variable and align by logical keys |
| Three particular equation quirks | Preserve any unusual source reference, nonlinear term and period-specific error behavior; review corrections separately |
| A particular path count and rejected bundle | Derive path status and coverage from current evidence; rejected implementations and approvals remain historical |
| Five external curves and their age-axis length | Derive IDs, counts, shapes and keys from the confirmed catalog; enforce the implemented capture shape restrictions |
| One worksheet's primary-input layout | Check source-specific adapter compatibility before generation/reconciliation; missing support is a blocker |

The extracted source facts are retained in the [Pricing case study](../examples/pricing_conversion_case.md), existing source-bound reports and their unchanged approval ledger. The reusable contracts no longer prescribe those facts for another workbook.

Schema identifiers and manifest-listed filenames are tool interfaces, not business defaults. Existing identifiers such as `gp.source_candidate_trace.v1` remain unchanged to preserve file contracts. Their names do not authorize a GP target, a fixed set of sheets or a reused scenario. Renaming interfaces or generalizing a source-specific executable adapter would be a separate implementation change.

## Verification and limits

The audit checks all six installed `stepN agent` outputs against their packaged files, scans for the removed case names/coordinates/counts, runs existing contract tests, and checks authored language and links. The source workbook, accepted report pairs, executable bundle and workflow decisions are preserved. See the local [audit record](../../output/agent_contract_genericity_20261009/audit-r1.md).

This change makes the instructions generic. It does not establish that every Excel can already complete Steps 4–5: supported formulas, external capture and source-specific adapters still require compatibility checks and potentially implementation work. No new native Excel run or alternate-scenario certification is implied.
