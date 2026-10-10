## What to deliver

Deliver `excel-to-act step2 trace`: bounded exploration of static upstream/downstream dependencies from a cell, range, or defined name, so an agent can see current paths, boundaries, and missing evidence. Validate the complete prepare → query → trace → citation-validation chain with source-backed fixtures and a workflow run scoped to its own evidence.

This is the third of three sequential implementation items and is recommended as one PR. It depends on the preparation package/graph from [#30 prepare](https://github.com/ferryhe/excel_to_act/issues/30) and the evidence packet, scope guard, and validator from [#31 query](https://github.com/ferryhe/excel_to_act/issues/31); do not rebuild the parser, query, or packet validator.

## CLI contract

```text
excel-to-act step2 trace --manifest MANIFEST --source-id SOURCE --kind KIND --target TARGET [--sheet SHEET] --direction upstream|downstream|both [--max-depth N] [--max-nodes N] [--max-edges N] [--out PACKET]
```

- `source-id` is always required; initial kinds are cell / range / name. Support finite default limits so every invocation is bounded.
- Upstream means target dependencies; downstream means consumers. Existing formula edges are formula cell → referenced input, so traversal direction must be converted explicitly.
- Reuse the prepared graph and lookup data; do not reparse the workbook or rebuild the graph for each trace.
- Expand supported, unambiguous absolute-A1 name destinations and stored range members so formula members can be traced further. Mark added parsing/membership links as source-backed structural derivations; do not present them as source formula edges.
- Reverse tracing must detect consumers that reference the target cell through a name or containing range; checking only exact-cell node edges cannot justify a no-consumer result.
- Constant names may be terminal. Keep relative/dynamic/external/unsupported multi-area references as unresolved boundaries. VBA literal references may be supporting evidence but cannot establish read/write direction, runtime order, or a complete call graph; `Range(Var1)` is not automatically resolved.
- Use a visited set and depth/node/edge limits to guarantee termination. Output frontier, truncated, unresolved, scope boundaries, and next action; a partial traversal must not be presented as dependency closure.
- Reuse item 2 scope guards: show boundaries crossing into excluded sheets and any permitted read-only data, but do not expand their internal formulas, even when an excluded cached cell contains formula text; do not modify source facts.
- stdout/output packets follow the delivered evidence/source contract. Integrity or invalid-selection failures return structured diagnostics and nonzero exit codes. Tracing does not count as an indexing recovery attempt.

## Deliverables

1. A runnable, discoverable trace CLI with parameter, boundary, and error contracts.
2. A minimal static resolver and bounded traversal using existing graph/reference helpers; reuse absolute-A1/scoped-name resolution. Do not build a formula evaluator or full VBA AST.
3. Trace evidence packets reuse item 2 canonical source records and claim validation; add derivation relations, frontier, limits, unresolved boundaries, and read-only scope labels. Preserve cross-sheet targets instead of converting them into false local edges.
4. Complete exploration instructions in the Step 2 agent / tool catalog / README: prepare → overview → small queries → bounded trace → source validation → Step 3 interpretation → fresh subagent review. Investigate unresolved/low-confidence/disputed items with specific evidence or a bounded TypeSafe/Jev latest judgment; do not treat a typed judgment as source, numeric, or runtime PASS.
5. A replayable CLI fixture walkthrough and a source-bound run record covering available checkbox, ActiveX/VBA entry points, and formula output paths; produce a static workflow specification and explicit unresolved items.
6. Trace regressions, cross-sheet and name/range relation tests, and full CLI integration checks; commands/artifacts must be replayable by a later agent using the instructions.

## Scope and dependencies

- The primary deliverable is a static exploration CLI; source-grounded exploration reports may be produced and reviewed. Do not build a full automated Step 3 semantic system.
- Do not run VBA, edit the original Excel file, implement recalculation/numeric oracles, or generate Python. Numeric reconciliation remains part of #7 and is not a prerequisite for this item.
- Do not expand the parser or reopen excluded-source repairs as part of this bounded trace work. Report supported paths and limitations accurately.
- Keep full macro read/write paths and dynamic addresses unresolved until proven. Baselines, input variations, and macro runtime cases are later behavioral-validation evidence, not substitutes for PASS of this static CLI item.
- Do not add a general graph framework, database, agent controller, or abstraction without a current consumer.

## Checks

- [ ] Small-fixture upstream / downstream / both directions agree with real formula-edge direction; cross-sheet targets retain original identity and location.
- [ ] Name destinations and formula members inside ranges can be traced further; reverse trace finds consumers that depend on a target cell through a name/containing range.
- [ ] Scoped names and same-name ambiguity do not guess; constant, dynamic, external, relative, and unsupported multi-area results are explicit.
- [ ] Cycles and depth/node/edge limits terminate; frontier and truncated results are correct. Hitting a limit or unresolved boundary must not claim all dependencies are known.
- [ ] Stop internal traversal at the approved read-only dependency scope; do not return excluded-sheet data/code outside that scope or expand the active scope without a newly confirmed decision.
- [ ] Trace the selected source's current declared outputs to supported targets, reporting unresolved paths instead of requiring every unknown path to be resolved.
- [ ] Replay source-backed fixture selectors for any available control, event, module, name and range records before trace and citation validation.
- [ ] Source-grounded reports distinguish “declared/static interpretation” from “executed”; cached formula values and module/event declarations are not new-scenario calculation or execution evidence.
- [ ] The complete-chain instrumentation records raw extraction, inventory parse, graph build, and selected-view loads, proving repeated query/trace reuses prepared results; ordinary reads do not consume the recovery budget.
- [ ] Give the exploration specification/evidence to a fresh independent reviewer; accept only reproducible findings directly mapped to this item's CLI/data contract, fix them and review again; related tests, Ruff, and CI pass.
