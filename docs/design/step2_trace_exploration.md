# Step 2 static exploration specification

Prepare the source once, inspect overview, then query small source-bound facts. Trace selected cells/ranges/names with explicit direction and finite bounds. Validate all citations against the delivered query/trace packet before interpreting a workflow in Step 3. A fresh reviewer receives the source/run/scope identity, commands, packets, report, and check results.

## Reproducible fixture

```powershell
python -m pytest tests/test_step2_trace.py -q
python -m pytest tests/test_step2_query.py tests/test_step2_prepare.py tests/test_graph_builder.py -q
python -m compileall src
python -m ruff check .
python -m pytest -q
excel-to-act --help
excel-to-act step2 trace --help
```

The trace fixture reuses the query fixture's canonical file/source/record contract. It covers upstream/downstream/both, quoted cross-sheet identity, name/range consumers, range formula members, scoped ambiguity, constants and unsupported names, cycles, each bound, approved and forbidden excluded cells, unchanged source facts, zero indexing recovery consumption, citation validity and tampering. The query fixtures replay checkbox/ActiveX/module evidence and citation validation.

## Pricing replay

Use the existing validated Step 1 artifacts as read-only inputs. Write every new output to a separate evidence directory. `SOURCE` is the exact `sources[].source_id` in the prepared manifest; the workbook SHA alone is not a substitute. All commands below run without workbook editing, VBA execution or recalculation.

```text
excel-to-act step2 prepare --index INDEX --step1-root STEP1_ROOT --scope analysis_scope.json --out READING
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind overview --out overview.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind control --sheet CalculationOfBE_OtherDisease --target "Check Box 1" --out checkbox.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind cell --sheet CalculationOfBE_OtherDisease --target F14 --out checkbox-cell.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind control --target PremiumTable.CmdPremium --out activex.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind vba --target Sheet9 --budget 20000 --out module.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind name --target ValueTable --out value-table.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind name --target Variables --out variables.json
excel-to-act step2 query --manifest READING/manifest.json --source-id SOURCE --kind name --target IssueAge --out issue-age.json
excel-to-act step2 trace --manifest READING/manifest.json --source-id SOURCE --kind name --target GP --direction upstream --out gp.json
excel-to-act step2 trace --manifest READING/manifest.json --source-id SOURCE --kind name --target NLP --direction upstream --out nlp.json
excel-to-act views validate --views gp.json --output gp-claims.json
```

For query pagination, continue the returned cursor with the same selector until the needed complete source record is delivered. Increase the budget when a record is oversized. For trace truncation, start from a specific frontier target with explicit bounds. Neither omitted records nor unresolved paths may be cited as established facts. The real local run records exact commands, outputs, counters and unresolved paths in the Issue's external evidence directory.

## Interpretation and evidence limits

- Existing formula edges are source declarations. Name destinations and stored range members are labeled structural derivations with graph/manifest/view/record provenance. They are never fabricated local formula references. Missing stored cells are explicit boundaries.
- Scoped names select a local declaration before a workbook declaration. Without context, multiple declarations require selection. Constant names terminate; relative, dynamic, external and unsupported multi-area destinations remain unresolved. Current-row structured references require caller context and are unresolved in this version.
- Excluded worksheet nodes are scope boundaries. Approved name destinations, stored range members and incoming edges from retained formula sources remain traversable within bounds. Approved dependency cell records may retain formula text as immutable source evidence. Excluded-sheet formulas are never expanded; data outside the approved range and worksheet code are never delivered. The 265 excluded invalid links remain outside the work queue.
- `OtherDisease Check Box 1 → F14` is a declared binding. `PremiumTable.CmdPremium → Sheet9.CmdPremium_Click` is a declared event association. Module exports and literal VBA references do not establish execution, direction, order or a full call graph. `Range(Var1)` remains unresolved.
- The approved Pricing dependency snapshot's 1,364 formula caches are saved source results, not a new scenario recalculation. Static exploration does not establish numerical equivalence, macro behavior or runtime success. `closure: not_proven` always preserves this distinction, even when the supported static traversal finishes.
- Counters distinguish raw extraction, inventory deserialization, graph builds, prepared graph loads and selected view loads. Repeat query/trace calls may reverify/reload prepared files; they never reextract the workbook or rebuild the formula graph. Reading never changes indexing recovery state.

## Fresh independent review

Replay the fixture and inspect the real report/packets against the ten Issue acceptance criteria. Trace original formula direction and cross-sheet identities; check reverse name/range consumers, ambiguity, each bound, cycles, partial status, scope, canonical records and tamper failures. Read the unchanged index state evidence and actual instrumentation. Do not infer semantic, numeric or runtime PASS from citation PASS.

Report only realistically reproducible CLI/workflow/data-contract/error-handling failures directly tied to this Issue. CI/check failures are merge-gate evidence. Do not expand scope to formula evaluation, VBA AST/execution, opaque parts, Step 3 automation, the excluded invalid links, or general hardening. A concrete disagreement may be sent with competing evidence to a bounded TypeSafe/Jev latest Judge; its typed decision is not source or runtime evidence and cannot replace a fresh reviewer PASS or required CI.
