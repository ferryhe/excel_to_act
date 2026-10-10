# Step3 evidence exploration and semantic draft

Goal: make the fact selection and dependency tracing used by the Agent repeatable, then use those tools to create a source-cited Pricing semantic draft for AI and people. Keep accepted source artifacts and the existing structural stages unchanged.

Pre-implementation assessment: complex. The change joins source/scoped-name selection, bounded graph traversal, stage integrity and canonical evidence validation across the Step3 workflow, exports, CLI and tests. Existing `_load_context`, `_read_stage`, inventory and dependency artifacts supply the foundation; standard-library JSON and graph traversal suffice. Keep the selected `gpt-6-sol / high` worker for implementation and repairs, then obtain a fresh `gpt-6-sol / high` acceptance review. Copilot is user-disabled. No publication is authorized.

## Small CLI surface

- `step3 query --analysis DIR --target TARGET [--sheet SHEET] [--offset N] [--limit N] [--out DIR]`: select an exact name, qualified finite cell/range or field ID. Return bounded complete source records, field summaries, source/run/stage binding, pagination and unknowns. A `vba:MODULE` target may expose the existing canonical inventory module record for static workflow analysis; it does not execute code or infer worksheet ownership. Ambiguity needs explicit context. Excluded worksheet logic stays outside direct cell queries.
- `step3 trace --analysis DIR --target TARGET [--sheet SHEET] [--direction upstream|downstream] [--max-depth N] [--max-fields N] [--out DIR]`: traverse checked field dependencies with deliberate prerequisite/consumer direction. Return readable nodes, edge evidence, internal dependencies, unresolved/excluded boundaries and explicit depth/size frontier. Trace does not resolve dynamic formulas or VBA execution.
- `step3 validate --analysis DIR [--spec INPUT.json]`: preserve existing structural validation. An optional Agent-authored spec cites declared evidence packet paths and byte hashes, current binding and stage hashes. Validate only supplied evidence IDs, reconstruct their source/structural facts against canonical artifacts, and render one validated context as `model_spec.json` and `model_spec.md`. Interpretations remain `inferred`; missing decisions remain `pending`; source facts must match checked evidence. No readiness claim becomes true.

Query/trace without `--out` return machine JSON. With `--out`, write paired `query.json/md` or `trace.json/md` from the same packet. Re-running identical input/selection produces identical files. Never write into bound source roots or overwrite structural-stage inputs with packet output. Invalid source/stage/spec/evidence must fail clearly; an invalid semantic input must not replace the last valid outputs.

Use a small fixed semantic section set: field roles, logical axes, calculation groups, outputs, macro workflow, open questions and validation plan. Preserve useful structured details and evidence references. A natural-language inference is not an exact fact claim; validating citations does not prove its business meaning or numerical behavior.

## Pricing exploration and deliverables

1. Capture Main inputs; Premium year/age axes, row-9/row-10 initial conditions and term indicators.
2. Trace GP/NLP and explain the cost aggregates and fee/annuity timing.
3. Capture the ten incidence columns and ten matching death-benefit columns; record candidate matrices and row-wise cost reduction.
4. Follow dynamic naming prefixes to current declared names and configuration cells. Preserve current-value resolution separately from general scenario validity.
5. Inspect state recurrences and EOP/MOP/BOP formulas, including survival/state weights and lookup matching conventions.
6. Read the retained PremiumTable macro, scenario table, ordered input writes and output mapping. Describe static behavior without claiming runtime execution.
7. Produce paired evidence packets, a paired semantic specification, a prioritized exploration backlog and an Excel/CPU reconciliation case plan. List exactly which evidence is still missing.

Actual runs live under `output/step3_exploration_20261008/`; human documentation links the paired artifacts and current CLI. Pricing and a different workbook must use the same commands, with no Pricing-specific parser rules.

## Acceptance checks

- Wrong/stale source or stage, ambiguous selectors and out-of-scope cells fail explicitly.
- Finite range/field paging is stable and declares omitted records; absent stored cells are not invented.
- Upstream/downstream and multi-start range traces preserve edge direction, stop on limits/cycles, and retain unknown/excluded frontier.
- Tampered packets, changed facts with updated packet hashes, unknown evidence IDs and stale spec bindings are rejected before replacing valid outputs.
- Fact, inferred and pending statuses are kept distinct. Semantic validation cannot enable runtime/generation/GPU readiness.
- JSON/Markdown pairs share content and provenance. Actual Pricing evidence supports every factual conclusion; unresolved interpretations remain explicit.
- Existing structural CLI and full tests continue to pass; a fresh independent reviewer checks realistic acceptance-scoped defects only.
