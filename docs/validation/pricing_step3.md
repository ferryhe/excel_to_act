# Pricing Step 3 Run

This run completed `prepare → fields → dependencies → plan → validate` against the checked Pricing source and produced a reusable draft of the structural analysis CLI. Structural validation passed; logical axes, dynamic references, and numeric execution still need evidence.

## How the Step 3 goal is served by the tools

Step 3 aims to produce a model specification that can guide implementation: its inputs and assumptions, the meaning of calculated fields, the axes along which they are calculated, dependencies and execution order, calculations that can be grouped, and how final outputs reconcile to Excel.

The current CLI provides the structural foundation for that specification. An agent still needs to read headers, formulas, names, and macros to organize business meaning, units, axes, boundary conditions, and dynamic selection rules. None of the `logical_axes` are confirmed yet; passing structural validation does not prove that the business specification is complete.

| CLI / tool used | How it serves the goal | Current artifact and Pricing example |
| --- | --- | --- |
| `tools` / tool catalog | Identifies available actions and their limits | Commands, prerequisites, and deliverables |
| `prepare` / Step 2 validation and source binding | Identifies which Excel file and scope are being interpreted | `analysis.json`; binds the current Pricing hash and 45 analyzed sheets |
| `fields` / cell grouping, copied-formula recognition, name mapping | Builds field candidates and their raw sources | `fields.json`; Main input scalars, GP/NLP calculated scalars, 106-cell calculated columns and a Qtable two-dimensional region in Premium |
| `dependencies` / static reference resolution and field mapping | Builds calculation chains and identifies where a chain stops | `dependencies.json`; GP depends on PVFB, AnnuityDue, and PVLoading; dynamic references remain explicit |
| `plan` / dependency ordering and candidate grouping | Gives an implementation order and identifies execution rules still needed | `execution_plan.json`; 4,150 blocks, 159 recurrence-candidate fields, 9 cycle-candidate blocks, and 23 vector-candidate groups |
| `validate` / source, coverage, and stage-contract checks | Keeps the draft internally consistent and hands off unknowns | `validation.json`, `handoff.json/md`; structural pass, while numeric and generation readiness remain false |

## One Pricing calculation chain

The table below is based on the currently bound `inventory.json`, `fields.json`, and static dependencies. Formulas and headers are source facts; items needing business confirmation are listed separately.

| Layer | Excel source facts | What can be stated now |
| --- | --- | --- |
| Inputs | `Main!C3/C4/C6/C7/C8` correspond to IssueAge, Sex, BenefitTerm, PremPayPeriod, and SumAssured; saved values are `0 / Male / 30 / 10 / 1000` | Five input scalars can be represented separately; valid values and units still need to be documented |
| Age-related column | `Premium!C8` is `Attain Age / Insured`; `C10 = IssueAge+B10-1`, `C115 = IssueAge+B115-1` | There is evidence for an attained-age axis; the meaning of column B, beginning/end-of-period alignment, and termination conditions still need confirmation from other formulas |
| Incidence-related column | `Premium!AA7` is `SelfDriveAccDeathInc`, AA8 is labeled `SelfDrive / AccDeath / Incidence`, and `AA10:AA115` contains copied formulas | It looks up Qtable using Sex and a dynamic column name, applies dynamic names `Fix` and `Mult`, then takes `MIN(...,1)`; when column C exceeds 105 it takes the previous row's AA value. Dynamic names, applicable inputs, and boundary rules need an explicit specification |
| Death-cost vector operation | `BV8` is `Death Benefit Cost`; `BV10 = SUMPRODUCT(Y10:AH10,AV10:BE10)*U10`; `U8` is `Discount MOP` | Each row takes a dot product of ten incidence-related values and ten corresponding death-benefit values, then multiplies by a discount-related value; the composition of the benefit values and MOP timing convention need verification |
| Cost aggregation | `Premium!CA8` is `Total Cost`; `CA10 = SUM(BV10:BZ10)`; `H3 / PVFB = SUM(CA10:CA115)` | The CA row values aggregate into PVFB; the meaning, units, and discounting relationship of BV:BZ need further tracing |
| Premium-related aggregation | `H1 / AnnuityDue = SUM(T9:OFFSET(T9,PremPayPeriod-1,0))`; `H2 / PVLoading = SUM(CD10:OFFSET(CD10,PremPayPeriod-1,0))` | Both aggregation lengths depend on PremPayPeriod. With the saved value of 10, the formulas select T9:T18 and CD10:CD19; the timing convention for their different starting rows needs verification |
| Outputs | `Premium!J1 / GP = PVFB/(AnnuityDue-PVLoading)`; `J2 / NLP = PVFB/AnnuityDue` | The direct formulas and prerequisite fields for these scalars are clear; their full input closure is still limited by upstream dynamic and boundary rules |

The agent can now trace GP/NLP from inputs through lookup/copied columns, row costs, and aggregations. The entire chain cannot yet be claimed as a numerically equivalent program. Some semantic evidence already exists in Excel headers and formulas; the next step is to extract and organize that evidence, then ask questions that truly require a business decision or runtime verification.

The field dependencies include the checked static path `AA10:AA115 → BV10:BV115 → CA10:CA115 → PVFB → GP`. The headers for Y:AH and AV:BE correspond, in order, to ten event categories: disease, accident, self-drive, public transport, air, train, natural disaster, commercial vehicle, elevator, and holidays. Combined with `SUMPRODUCT`, this supports candidate calculation matrices of `106 rows × 10 event categories`, whose row-wise dot products produce the BV vector. This grouping is supported by evidence in Excel; the time/age convention for rows, column mapping, and numeric examples still need confirmation, and the current CLI has not yet produced this semantic specification.

The existing `fields.md` shows only sample fields, and `execution_plan.md` mainly shows block identifiers; the JSON contains full membership, dependencies, and candidate groups. These are enough for agent navigation and structural checks, but a readable full Pricing business specification still needs to be added. The scenario/age loop, input-write order, and output mapping in `PremiumTable.CmdPremium_Click` remain in the original business-process plan and have not been incorporated into the full specification or runtime validation.

## Source and scope

- Source: `input/Pricing.xlsm`, SHA-256 `dc1de98b5e27fa85e233e7ccfa56bc21f571b2bb1686ab13b86a43a9cdfe7c7c`.
- Step 1 run: `20261008T120940-8282f510`; Step 2 source ID: `695cf70322dddcb11534`.
- Index: `output/real_steps_20261008/step2/index.json`; preparation revalidated all 62 artifact references.
- The four-sheet exclusion decision matched the source hash and was reused; 45 sheets were analyzed. Cross-sheet boundaries were recalculated, without treating an old dependency snapshot as new evidence.
- The original workbook, Step 1 artifacts and handoff, and Step 2 index/reports/state—67 files total—were checked and remained unchanged.

## Actual results

| Item | Result |
| --- | --- |
| Retained / excluded stored cells | 155,908 / 1,856; every retained cell belongs to exactly one primary field |
| Primary fields | 4,324: 3,078 source-value fields and 1,246 calculated fields |
| Calculated-field shapes | 230 0D and 1,016 1D; shapes describe physical structure, while logical axes remain unconfirmed |
| Retained formulas and caches | 99,948 formula cells; saved caches remain derived results and are not raw inputs |
| Workbook-wide named-range descriptors | 550: 499 0D, 19 1D, 31 2D, and 1 unresolved; a two-dimensional descriptor can include different fields |
| Inter-field / intra-field dependencies | 105,910 / 159 |
| Reference boundaries | 3,776, including 647 partial-empty-range boundaries; 54,091 references involved empty cells, not 54,091 distinct empty cells |
| Excluded-sheet boundaries | 864 static references, retaining the source-scope limitation |
| Unresolved references | 140 groups: 139 dynamic and 1 broken; 53,109 dynamic diagnostic occurrences, not distinct cells |
| Execution plan | 4,150 dependency blocks; 9 cycle-candidate blocks; 159 fields with recurrence candidates; 23 vector-candidate groups |
| Validation | `status: pass`; `runtime_verified / generation_ready / gpu_executable` are all `false` |

Source-value fields include text, labels, blanks, and numbers, so the 3,078 source-value fields cannot all be described as numeric inputs. A two-dimensional region descriptor does not automatically combine adjacent business columns into one two-dimensional calculated field.

Paths inspected directly:

| Object | Result |
| --- | --- |
| `Main!C3/C4/C6/C7/C8` | `IssueAge`, `Sex`, `BenefitTerm`, `PremPayPeriod`, and `SumAssured` remain separate source-value scalars |
| `Premium!J1` / `GP` | Calculated scalar; depends on `PVFB`, `AnnuityDue`, and `PVLoading` |
| `Premium!J2` / `NLP` | Calculated scalar; depends on `PVFB` and `AnnuityDue` |
| `Premium!AA10:AA115` and similar | 106 copied-formula cells form a physical one-dimensional field; lookup, dynamic references, and previous-row references are retained without claiming a confirmed time axis |
| `Qtable!C5:BT111` | Two-dimensional region descriptor for the `Qtable` name |
| `PremiumTable!A2:G8` | Two-dimensional region descriptor for `ValueTable`; it does not imply that all its columns belong to the same variable |

## Commands and artifacts

```powershell
$analysis = "output/step3_pricing_20261008/analysis"
excel-to-act step3 tools
excel-to-act step3 prepare --index output/real_steps_20261008/step2/index.json --step1-root output/real_steps_20261008/step1 --out $analysis --scope output/control_modules_20261006_a9a9d3/analysis_scope.json
if ($LASTEXITCODE -ne 0) { throw "Step 3 prepare failed; read diagnostics." }
foreach ($stage in @("fields", "dependencies", "plan", "validate")) {
    excel-to-act step3 $stage --analysis $analysis
    if ($LASTEXITCODE -ne 0) { throw "Step 3 $stage failed; read diagnostics." }
}
```

The output directory contains `analysis.json`, `fields.json` / `fields.md`, `dependencies.json`, `execution_plan.json` / `execution_plan.md`, `validation.json`, and `handoff.json` / `handoff.md`. Machine-readable files retain full field and dependency data; readable reports provide an overview. Execution logs and step-by-step responses are in the parent directory's `logs/` and `cli_execution.json`.

The same commands also completed against a two-sheet `Reusable.xlsx` with 35 stored cells: 12 fields, 6 field dependencies, 1 dynamic-reference group, and 1 recurrence candidate. It came from a different workbook and used no Pricing-specific paths or name rules. The nine stage/report files had identical SHA-256 hashes after repeated runs. Running it with the Pricing scope file returned `blocked`, exit code 1, and did not write an analysis manifest.

After repeating Pricing preparation, its nine files were unchanged; rerunning the fields, plan, and validation stages on the final code left all nine files byte-identical again. Checksums and records are in `prepare_resume.json` and `determinism.json`; original-source checks are in `input_preservation.json`.

## Development checks

- `python -m pytest -q --disable-warnings`: 230 passed, with 19 pre-existing dependency deprecation warnings, in 60.88 seconds.
- Ruff passed for Step 3, the CLI, graph builder, inventory extractor, and added tests.
- `git diff --check`: passed.
- Independent `gpt-6-sol / high` review: PASS; report at `output/step3_pricing_20261008/review_8.md`.
- The final code was run again through the four analysis stages on the different workbook and Pricing validation: all passed, artifacts remained byte-identical, and 67 source files remained unchanged; recorded in `final_verification.json`.
- Structured table references were checked for headers, data rows, totals rows, current-row syntax, disabled headers, and older artifacts missing metadata. Column references use native table column names; when older artifacts lack column names or required row boundaries, the reference remains unresolved until re-extraction, finalization, and indexing provide the new information.
- Native table headers are retained as individual fields; named scalars that reference themselves through a cell, name, or multi-target range are classified as cycle candidates, while recurrence candidates in copied formulas across rows remain intact.
- Whole-row/whole-column named ranges and bare `[@Column]` syntax remain explicitly unresolved; these unknowns do not block the remaining structural analysis. Finite ranges and current-row column references qualified by a table name use the supported rules.
- Copilot is disabled per user preference; this work was not committed, published, or merged.

## Remaining work

| Item to resolve | Impact |
| --- | --- |
| Logical axes, offset meanings, and initial conditions for 159 fields | Determines recurrence direction and how much history to retain |
| Actual loop paths and solution rules for 9 cycle-candidate blocks | Distinguishes a cycle caused by grouping from actual simultaneous iteration |
| Target rules or source for 140 dynamic/broken-reference groups | Completes target-field closure; saved caches cannot replace cross-scenario calculations |
| Axis confirmation and CPU/Excel reconciliation for 23 vector-candidate groups | Determines whether to combine calculations, then how to select a GPU execution method |

This run did not execute macros, recalculate formulas, generate a Python numeric model, or make GPU performance claims. The business-process specification for `PremiumTable.CmdPremium_Click` remains a later semantic increment.
