# Pricing GP calculation: saved scenario

Python independently reproduces `GP = 2.7283284514524744` from the original `Pricing.xlsm` formulas and literal inputs. A fresh native Excel rebuild agrees exactly on GP and its three root intermediates. This is a verified saved-scenario calculation prototype; it is not general code generation, full-workbook equivalence, or GPU execution.

## Bound result and inputs

Source: `input/Pricing.xlsm`; SHA-256 `dc1de98b5e27fa85e233e7ccfa56bc21f571b2bb1686ab13b86a43a9cdfe7c7c`.

Target: defined name `GP`, `Premium!J1`, formula `=PVFB/(AnnuityDue-PVLoading)`.

Scenario: IssueAge 0, Sex Male, BenefitTerm 30, PremPayPeriod 10, SumAssured 1000, PricingIntRate 0.035, CIDecrement Yes; no overrides. Other selected inputs are read from the same hash-bound workbook. The raw model value's business units remain pending confirmation.

| Quantity | Python | Fresh Excel |
| --- | ---: | ---: |
| AnnuityDue | 8.574676381865864 | 8.574676381865864 |
| PVLoading | 3.161998467798486 | 3.161998467798486 |
| PVFB | 14.767563151498457 | 14.767563151498457 |
| GP | 2.7283284514524744 | 2.7283284514524744 |

The annuity range is `Premium!T9:T18`; loading is `CD10:CD19`; benefit costs sum `CA10:CA115`. The CLI compares 7,763 reached Premium formulas, including 2,009 annual state, timing, and cost values. There are no mismatches at absolute and relative tolerances of `1e-12`.

## Reusable CLI flow

```powershell
excel-to-act step3 oracle --workbook input/Pricing.xlsm `
    --target AnnuityDue --target PVLoading --target PVFB --target GP `
    --range "Premium!B9:CD115" --out output/pricing/oracle
if ($LASTEXITCODE -ne 0) { throw "Native Excel oracle failed; read diagnostics." }
excel-to-act step3 gp --workbook input/Pricing.xlsm `
    --baseline output/pricing/oracle/excel_oracle.json `
    --abs-tol 1e-12 --rel-tol 1e-12 --out output/pricing/gp
if ($LASTEXITCODE -ne 0) { throw "GP calculation or reconciliation failed; read diagnostics." }
```

`oracle` requires Windows and Microsoft Excel. It opens a private byte-identical copy read-only, disables macros and events, runs CalculateFullRebuild, captures requested names/ranges and Excel settings, then closes without saving. The numerical `gp` command requires a successful `gp.excel_oracle.v1` result from `step3.oracle`, explicit matching source/copy hashes, and all seven saved primary inputs. Failed or incomplete baselines are blocked before replacing existing results. Formula-result caches are never calculation inputs. Unselected supported IF branches remain unevaluated; active unsupported functions/references and actual cycles block rather than yielding a fabricated value.

A changed workbook needs a fresh baseline and a fresh result/closure assessment. The GP command currently uses the `Premium!J1` target and the Main/Premium comparison profile. The source formula subset is bounded; a different layout or unsupported feature requires further analysis and implementation. The static `prepare → fields → dependencies → plan → query/trace → validate --spec` workflow remains available for that investigation.

## Deliverables for agents and humans

Actual run directories:

| Deliverable | Agent-readable artifact | Human-readable artifact / purpose |
| --- | --- | --- |
| Native oracle | `output/step3_gp_oracle_20261008/excel_oracle.json` | `excel_oracle.md`: engine/settings and requested values |
| Calculation and reconciliation | `output/step3_gp_runtime_20261008/gp_result.json` | `gp_result.md`: result, intermediates, compared values, tolerance, and any mismatch |
| Active path | `output/step3_gp_runtime_20261008/active_trace.json` | `active_trace.md`: bounded examples; full cell/formula/branch/name/lookup details remain in JSON |
| Computation specification | `output/step3_gp_calculation_20261008/gp_computation_spec.json` | `gp_computation_spec.md`: source/calculated scalar/vector/matrix groups, logical axes, ordering, shared work and scope |
| Independent evidence | `runtime_audit.json`, `native_semantics_cases.json`, `python_semantics_verification.json`, `inputs_before.json`, `inputs_after.json` in the specification directory | Trace order/endpoint checks, 11 native behavior comparisons, and preservation evidence |

The active trace contains 34,296 reached cells: 11,198 literal source cells, 22,910 ordinary formula cells and 188 array-formula members. Its 106,613 dependency records include selected/scanned range metadata and explicit array-member-to-anchor edges. `discovery_order` and `completion_order` are separate; checked cell prerequisites complete before their consumers. The source/scanned range metadata endpoints are not additional calculated cells.

The specification's semantic groups are agent interpretations grounded in labels, formulas, and the runtime trace. They propose future vector/matrix implementation boundaries; the current evaluator does not emit NumPy/GPU kernels. The native comparison directly checks reached Premium interfaces and annual results, not every internal upstream cell separately.

## Required path decisions

Seven numerical business routes are verified for this scenario: annuity, loading, death, CI, waiver, survival, and medical. Waiver/survival/medical cost vectors are zero under the current choices, while source-required incidence operands remain evaluated. The eighth previously registered path is the GP feedback condition: Level100/None columns are selected, so the GP-dependent ROP/WaiverOfPrem columns are inactive. A feedback solver is not implemented or verified.

The numerical route ratio is 7/7 within the saved scenario; the registered-path runtime count is 7/8 when the inactive feedback unit remains in the denominator. These are known-only measures. Coverage of all possible configurations and branch combinations is unknown.

The run reaches selected final-rate bridges in four sheets previously excluded from the static draft: CalculationOfBE_CI, CalculationOfBE_MinorCI, CalculationOfBE_ModerateCI, and CalculationOfBE_SpecialCI. Only their needed male outputs, keys and prerequisites are read. The prior scope and source-bound artifacts are unchanged. Source formulas retain the unusual previous-H reference in the I recurrence and the two I factors in CI cost. The trace records 150 source IFERROR catches of division-by-zero; those fallback results agree with native Excel.

## Verification and boundary

Both actual CLI commands exited successfully. Eleven native lookup/error/branch cases agree with the independent evaluator; five invalid-baseline CLI cases are rejected without replacing prior results. The original workbook and all 90 protected prior files retain their hashes. The final full suite has **249 passed** with 19 existing dependency deprecation warnings. Ruff and diff checks pass. The wheel contains byte-identical calculator/native-script resources and all three agent contracts. A fresh independent Sol/high review is **PASS**. Complete hashes and gate evidence are in `output/step3_gp_calculation_20261008/validation.json` and `validation.md`.

Static citation validation remains separate from numerical verification. Other configurations, actual GP feedback, actuarial correctness or units beyond reproducing the workbook, array code generation, Julia/C++ backends, and GPU correctness/performance remain outside this increment. Copilot is user-disabled. No commit, PR, publication, or merge was performed.
