# GP modular implementation route

Status: implementation architecture for the approved saved-scenario GP. This document proposes code and contract changes; it does not change the source workbook, existing workflow artifacts, or runtime code.

Source identity: `input/Pricing.xlsm`, SHA-256 `dc1de98b5e27fa85e233e7ccfa56bc21f571b2bb1686ab13b86a43a9cdfe7c7c`. Keep the implementation bound to this hash and the existing saved case. The current active trace contains 22,910 ordinary formula cells and 188 array-formula follower cells, for 23,098 formula cells total. A static profile groups the ordinary formulas into 263 syntactic candidates; the candidates still need reviewed semantic mappings. The 188 array followers are eight active array results across four selected CI tables, using two `TRANSPOSE` shapes per table.

## Recommendation

Build a source-bound family compiler with a deliberately authored pricing orchestrator. Do not write a hand-copied formula implementation for the whole closure, and do not retain the current per-cell formula interpreter as the output model.

The family compiler parses each active source formula at generation time and applies the approved semantic variable map. The current map groups the 263 syntactic source patterns into 72 semantic equation groups; source-only probing found 178 distinct logical formula-body variants within those groups. Emit one direct Python function per distinct reviewed variant, deduplicated across repeated rows, and retain the named group and source-family membership as provenance. Do not force distinct variants into one function merely to meet the group count, and do not emit a per-cell dispatcher. Functions use logical variable names and axes, never Excel-cell lookups. `pricing.py` owns readable module order and the explicit Premium and MultipleCI recurrence loops. This keeps the actuarial projection logic visible while covering every active formula.

The alternatives do not meet both constraints as well:

| Route | Assessment |
| --- | --- |
| Handwritten fixed-GP template | Readable for a few core equations, but it would need hand-authored implementations and source-coordinate mappings for all 263 candidate families, eight array shapes, lookup branches, and 23,098 outputs. It is not the smallest reliable route to complete trace coverage and invites formula omissions. Retain handwritten orchestration and core recurrence layout, not a hand-copied implementation of every rate and lookup formula. |
| Source-family compiler plus explicit orchestration | Recommended. It reuses the current formula parser and source-only trace, maps all active members to reviewed variables/families, emits one specialized function per distinct reviewed formula-body variant, and keeps loops/reductions in ordinary Python. The current target is 72 semantic groups and 178 distinct variants, not 23,098 functions. It is source-bound and limited to the approved workbook and semantic plan. |
| Current per-cell emitter/runtime | Rejected product shape. `generator.py` emits a `FORMULAS` lambda for each formula address; the bundled `Runtime` then dispatches by cell address. That reproduces the workbook graph, not compact calculation logic. Moving those lambdas to another module or compressing their names would not fix it. |

The compiler is a generation-time facility, not a generic Excel engine in the standalone package. Unsupported active syntax, an unresolved source reference, an unmapped family member, or a shape mismatch blocks generation. There is no per-cell fallback.

## Contracts and generated bundle

Keep the source coordinate boundary in one adapter. The pricing modules use the variable IDs, values, and axes from the accepted Step 3 semantic plan. Cell addresses may appear in source values, source maps, formula fingerprints, reconciliation output, and exception messages; they must not be variable names used by the pricing API.

```text
bundle/
  model.py                 # standalone CLI and result writer
  input_adapter.py         # source-address literals -> named raw input structures
  formula_families.py      # compiled direct functions, one per distinct reviewed variant
  pricing.py               # readable table, rate, recurrence, cashflow, and root flow
  runtime.py                # small Excel-value/error/lookup/transpose primitives
  source_values.json        # literal/blank/error/date inputs only; source-address ledger
  source_map.json           # source coordinates <-> logical variables and axes
  family_manifest.json      # family fingerprints, members, segments, coverage
  model_manifest.json       # source/design/trace hashes and bundle-file hashes
```

Keep `source_values.json` address-keyed for exact provenance and Stage 5 input checking. `input_adapter.py` uses `source_map.json` to form named scalars, vectors, and tables. It must reject a formula address in the raw ledger. Formula caches and results from previous models never enter the bundle. Formula-derived `Qtable`, CI rates, coefficient vectors, bridges, and Premium values are recomputed.

The source map records each scalar or range rule with `variable_id`, role/kind, logical axes and values, source extent, demanded indices, seed/segments, and a coordinate mapping. It also records each active formula member's `family_id`, equation segment, mapped variable/index, and formula fingerprint; array followers additionally map to an anchor and 2D offset. This cell-level ledger is only for completeness and reconciliation. Store regular extents/axis transforms compactly and exceptions explicitly rather than generating per-cell code or per-cell runtime formulas.

The formula fingerprint binds a canonical parsed expression to the semantic mapping: operators/functions/literals; references rewritten to `variable_id` plus axis offset or table lookup; relevant defined-name definitions; equation segment; and, for arrays, anchor shape and orientation. The member ledger binds source address and exact source formula hash. Also bind the original workbook, Step 3 semantic plan, and fresh Step 4 trace hashes. A changed formula, name, axis mapping, or array shape invalidates the matching mapping. Formula text that is merely translated or syntactically similar does not become a semantic family without the reviewed plan.

The generated result should expose `targets`, named calculated `variables` needed for review, and `cells` emitted by the source-map adapter. `cells` is a diagnostic/reconciliation view keyed by source address, not the model interface. For this trace it must cover the exact active formula-member set: 22,910 ordinary formulas plus 188 mapped array followers. The manifest should report `ordinary_formula_count`, `array_follower_count`, their union `active_formula_member_count`, `unique_semantic_family_count`, `unique_compiled_variant_count`, `unique_array_family_count`, `array_instance_count`, and `unmapped_count` (which must be zero). The current target is 72 semantic groups and 178 distinct compiled formula-body variants. Report source array instances/followers separately from the much smaller group and variant counts.

## Variables, dimensions, and execution

Use the latest design's logical roles and dimensions. In particular, `Premium!H1:H3` and `J1` are scalar roots; `H10:H115`, `I10:I115`, and `J10:J115` are separate 106-value policy-year projections even though they share columns with those roots. `B10:B115` is a raw policy-year-key vector. CI rate data is age/category, while MultipleCI is age/state. Do not flatten these axes into one universal time index.

| Contract group | Runtime shape and boundary | Execution |
| --- | --- | --- |
| Raw inputs | Named scenario scalars and selectors, literal lookup tables/headers, literal policy-year keys, blanks, and error/date values. `source_values.json` holds only these source facts. | Build once through the source adapter. Do not substitute calculated formulas or caches. |
| Selected CI rates and bridges | Four age-by-24-category matrices (`106x24`), four 106-value age bridges, and computed table/coefficient vectors. | Compute lookup-table formula families and exact selected male tables first. Category and age-table rows are independent after their prerequisites. |
| MultipleCI | Four derived ratio aliases, four derived payment weights, 14 derived state multipliers; 14 state vectors with row-8 zero seeds; Health with row-8 seed 1; 105 age outputs `B9:B113` and `T9:T113`. State/Health updates are needed only for rows 9:112 (104 updates); T is needed for rows 9:113. | For each age update, calculate W/X as demanded, snapshot the prior 14 states and Health, compute all 14 next states from that unchanged snapshot, then update Health. Compute T through age row 113 from its prior state row and current X; row 113 has no W or state update. Preserve O9's exact incoming terms `H8*RatioCI3 + I8*RatioCI3 + L8*RatioCI1`. The terminal Q state, Death R, and diagnostic age columns are excluded. |
| Premium states and decrements | H/I/J each have 106 values, policy years 1:106. K/L/O/P/Q have 106 demanded values through row 115; M/N have 105 demanded values through row 114. T is 107 values for years 0:106 (`T9=1` is index 0); U/V each have 106 values for years 1:106. F and CC/CD are needed only for the first 10 rows. | Seed H/I/J at row 10 with 1 and T year zero with 1. For each policy year, calculate the H/I/J state from the previous-year snapshot, then the current-year K/L/O/P/Q values; calculate M/N only through year 105. Preserve I's additive prior-H reference and source-local IFERROR rules. Do not compute excluded G/R/S/CB or unused row-115 M/N. |
| Incidence/benefit and costs | Death `106x10`, CI `106x5`, waiver `106x6`, survival `106x3`, medical `106x2`; BV:BZ and CA are five 106-value cost vectors plus total. | Calculate all demanded source operands even when the selected benefit is zero. Preserve source branches and errors. BY10 has no IFERROR; BY11:BY115 uses IFERROR blank. BW retains its `I*I` operand. |
| Roots | AnnuityDue, PVLoading, PVFB, and GP are separate calculated scalars. `CC/CD` are only needed for the first 10 loading rows. | Resolve the source OFFSETs from `PremPayPeriod` (saved case: T9:T18 and CD10:CD19), sum CA10:CA115, then calculate `GP=PVFB/(AnnuityDue-PVLoading)`. |

`pricing.py` uses named objects such as `raw.policy_year`, `rates.ci_by_age_category`, `multiple_ci.state_by_age`, `premium.h_state`, `premium.i_state`, `premium.j_state`, `timing.t_by_policy_year`, and `costs.pvfb`. Keep ambiguous state names source-based until business meanings are confirmed. For the 14-state grid, use a 105-by-14 logical array and the source header labels as metadata; do not expose `C8`, `D8`, etc. as model fields.

The planned order is: load literals; compute prerequisite lookup/coefficient families; compute selected CI rate matrices and their row totals; calculate the four final CI bridges and derived selected rate columns; project the 14-state age recurrence and T outputs; project annual Premium states/decrements; compute incidence/benefit matrices; calculate T/U/V, costs and loading; reduce the roots. Bridges consume the selected rate totals and cannot precede them. The plan's family dependency graph supplies the acyclic order. Treat explicit lagged age and policy projections as recurrence groups: take a prior-row snapshot and update each group atomically. Independent table rows and categories can be vectorized; recurrence time remains sequential. A same-time cycle without an approved recurrence group blocks.

The two recurrence drivers should be as plain as this shape-level pseudocode; each named next-state function is emitted once from its reviewed source family and source fingerprint:

```python
for age_index in range(105):
    x_by_age[age_index] = ci_rate(age_index)  # X is demanded through source row 113

for age_index in range(104):
    previous_states = state_by_age[age_index].copy()  # 14 source states, no in-place reads
    previous_health = health_by_age[age_index]
    w = death_rate(age_index)
    x = x_by_age[age_index]
    next_states = [
        next_state_01(previous_states, w, x, coeffs),
        next_state_02(previous_states, w, x, coeffs),
        next_state_03(previous_states, w, x, coeffs),
        next_state_04(previous_states, w, x, coeffs),
        next_state_05(previous_states, w, x, coeffs),
        next_state_06(previous_states, w, x, coeffs),
        next_state_07(previous_states, w, x, coeffs),
        next_state_08(previous_states, w, x, coeffs),
        next_state_09(previous_states, w, x, coeffs),
        next_state_10(previous_states, w, x, coeffs),
        next_state_11(previous_states, w, x, coeffs),
        next_state_12(previous_states, w, x, coeffs),
        next_state_13(previous_states, w, x, coeffs),
        next_state_14(previous_states, w, x, coeffs),
    ]
    state_by_age[age_index + 1] = next_states
    health_by_age[age_index + 1] = previous_health * max(1 - w - x, 0)

for age_index in range(105):
    ci_output[age_index] = source_ci_incidence(
        age_index, state_by_age[age_index], health_by_age[age_index], x_by_age[age_index], coeffs)

for year_index in range(106):
    if year_index:
        previous = snapshot_premium(year_index - 1)
        h_state[year_index] = source_iferror(lambda: max(
            previous.h + (-previous.k - previous.l + previous.o - previous.p)
            / previous.j * previous.h - previous.m, 0), 0)
        i_state[year_index] = source_iferror(lambda: max(
            previous.h + (-previous.k - previous.l + previous.o - previous.p)
            / previous.j * previous.i - previous.n, 0), 0)  # source adds prior H, multiplies prior I
        j_state[year_index] = previous.q_state
    k_l_o_p_q = current_decrements(year_index, h_state, i_state, j_state, rates)
    if year_index < 105:
        m_n = current_minor_moderate_decrements(year_index, h_state, i_state, rates)
```

`state_01` through `state_14` map in source order to the labels in `MultipleCI_Markov!C7:P7`; labels remain metadata unless their business meanings are confirmed. The source formula mapping supplies each of the 14 actual state equations; it is not a generic transition rule. Each `next_state_XX` receives the prior row and the named ratio/multiplier/lookup values selected by that equation. In particular, `next_state_13`/the source O equation retains the unusual O9 terms documented above. `source_ci_incidence` preserves the source `IF(B>=106,1,...)` branch, which is inactive on this saved age 0–104 path. After the policy loop, compute `T[0]=1`, T years 1:106, U/V years 1:106, costs, and the four scalar roots. The model must not conflate year-index 0 with the year-zero seed of T.

## Family compilation and source semantics

Reuse the source-only Step 4 discovery and the existing Step 3 `_Parser`/formula evaluator's parsing and value rules. Reuse the existing `TRANSPOSE`, `MATCH`, exact/approximate `VLOOKUP`, `OFFSET`, `INDIRECT`, error, and branch semantics as compatibility helpers where reached. Do not copy the Step 3 evaluator wholesale into the standalone bundle: it is an address-based interpreter and has an intentionally broader responsibility.

At generation time, lower a representative formula AST through the approved family map. A reference becomes a named scalar, a named variable at a logical index/lag, a table row/column access, or a declared dynamic lookup. Emit direct Python for that family, such as `premium_h_next(previous, decrements)` or `ci_component_rate(age, category, tables)`. Helper calls may implement Excel-compatible error propagation and lookup semantics, but there is no `env.get("Sheet!A1")`, formula-token dispatch, or per-cell formula cache in the standalone runtime.

Group only members whose normalized formula structure, axis offsets, literals, dependencies, equation segment, and output shape match the approved family. The source profile's 263 candidates are a starting inventory, not an instruction to emit exactly 263 functions. The semantic plan may split a syntactic candidate when axes or behavior differ, and a semantic group may contain several distinct formula-body variants. Deduplicate identical reviewed bodies across source members and groups, but preserve variant differences. The current acceptance target is 72 semantic groups and 178 distinct direct functions; function count follows distinct logical formula variants, not repeated rows or 23,098 source addresses. Each active address must map exactly once to a variable/index or declared array slot and one source-family member; no active address may be silently excluded.

For the eight selected `TRANSPOSE` array ranges, compile once per unique semantic array family (two formula shapes/patterns across eight source instances), preserve the source orientation, then map each active follower to its row/column offset in the resulting age/category table. The trace counts 188 active follower cells; do not emit 188 functions. The generator verifies source anchor, `array_ref`, parsed result shape, and complete follower mapping. Reuse the existing array shape checks as the low-level model, but make shape and logical axes explicit in the semantic plan.

Dynamic names and lookup semantics stay source-bound. Compile `MATCH` modes and `VLOOKUP` approximate/exact modes from the formulas; do not replace them with nearest-key defaults. Keep `IF`, `IFERROR`, and `IFNA` lazy so inactive branches do not invent dependencies and active errors are handled at the formula that owns the fallback. Keep `OFFSET` lengths driven by the raw `PremPayPeriod`; do not bake the saved output ranges into inputs. If an active `INDIRECT`, structured reference, or formula feature cannot be mapped to a logical input/family, stop with a source address, fingerprint, and missing mapping.

## Step 3–6 change surface

| Stage | Contract changes | Required evidence and tests |
| --- | --- | --- |
| Step 3 | Add the semantic plan/check to the accepted Stage 3 design: variable role/kind/axes, full extent and demanded indices, seeds/segments/lags, equation-family IDs/fingerprints, modules, source-member mapping, and array shapes. Bind JSON/Markdown hashes with the design. Static check reports candidate coverage; it does not claim to prove the current target closure. | Extend `steps/step3/design.py`, `workflow.py`, CLI/tool catalog and `tests/test_step3_workflow.py`. Test required fields, axis and boundary coverage, family fingerprint staleness, duplicate/unmapped candidates, and explicit array shapes. Keep current `FormulaEvaluator` useful for source queries/validation, not as the generated model. |
| Step 4 | Keep `discover` source-only: formula text plus literal values, no Excel and no formula caches. Add semantic family/shape checks to generation; bind fresh trace to the approved plan and current workbook hash. Emit the modular bundle, source map, family manifest, direct family functions and readable recurrence orchestrator. Block on any unassigned active ordinary formula or array member. Smoke run all targets and source-map outputs in isolation. | Refactor `steps/step4/generator.py`; reuse parser/array discovery from `discovery.py` and small helpers from `runtime.py`. Extend `tests/test_step4_generation.py`: translated rows share one family function; different axes/lags split appropriately; 188-style followers map to shaped arrays; mutations to source, defined names, family hashes, or shape block; output function count is stable as members increase; standalone bundle contains no project/workbook imports or derived raw inputs. Update Step 4 contract/tool catalog. |
| Step 5 | Allow a hash-checked multi-module bundle and validate every bundle file, source map, family manifest, and raw-input ledger. Keep the isolated run. Reconcile target roots and every mapped active formula value, including the non-Premium bridge/lookup formulas, not only `Premium!B9:CD115`. Keep error comparisons exact and absent comparisons blocking. | Update `steps/step5/workflow.py` and add `tests/test_step5_workflow.py`. Request 15 per-sheet active-formula bounding ranges: current trace boxes total 34,576 cells. With the existing 8,667-cell `Premium!B9:CD115` baseline, the combined 43,243 cells remain under the current 50,000-cell oracle limit. Test full formula address-set equality, values/errors, array followers, raw primary inputs, stale hash rejection, and no workbook in standalone execution. |
| Step 6 | Report logical module/family counts, 22,910 ordinary plus 188 array-follower coverage, mapped active formula count, demanded shapes, and the source-bound saved-scenario limit. Rename or supplement the current `active_premium_formula_cells` metric with all-source active formula counts. Never imply all-configuration or full-workbook support. | Update `steps/step6/workflow.py`, `tests/test_step6_report.py`, and Step 6 contract/tool catalog. Test that the reported family/formula counts come from accepted Step 4/5 evidence, not the Step 3 candidate profile. |

Update `interfaces/cli.py`, each changed stage's `agent.md` and tool catalog, and the CLI descriptions in `README.md`/`README.zh-CN.md`. Keep the actual command names where possible; `step4 generate` should consume only the semantic plan bound to the approved Stage 3 artifact, never an alternate unbound path.

## Acceptance and implementation tier

Generation is ready only when the fresh trace has zero unmapped formula members, all 22,910 ordinary formulas and 188 array followers are represented in the source-map coverage ledger, array shapes agree, and no active family or source input is missing. The bundle must run without the workbook or project package, use no formula cache values, and expose a compact family count rather than one function per formula cell.

Stage 5 must reconcile the four roots (`AnnuityDue=8.574676381865864`, `PVLoading=3.161998467798486`, `PVFB=14.767563151498457`, `GP=2.7283284514524744`) plus every source-mapped active formula result against a fresh full-rebuild Microsoft Excel oracle. Require the exact active formula/member set, exact error outcomes, source/plan/trace/bundle hashes, and the agreed numeric tolerance. The four root values are comparison results only; they cannot be model inputs. Stage 6 must identify this as one saved-scenario execution and keep other configurations, active feedback, GPU, and full-workbook conversion outside scope.

Implementation tier: **complex, `gpt-6-sol / high`, one persistent worker**. Evidence: 263 source-pattern candidates across 15 sheets, 8 transposed array shapes, source-specific dynamic lookup/error behavior, a 14-state age recurrence, H/I/J coupled policy recurrences, and changes to all four downstream stage contracts. The state equations, data axes, and boundary rows are now explicit; the hard work is careful mapping and integration, not an unresolved algorithm requiring the extreme tier. Keep one worker through fixes and fresh review; do not escalate because of the raw cell count alone.
