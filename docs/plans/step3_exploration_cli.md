# Step 3 source evidence CLI

The query and trace tools let an Agent inspect checked source records and bounded static dependencies for a selected workflow. Use them to build a source-cited design while keeping accepted upstream artifacts unchanged. The commands navigate evidence; they do not calculate formulas, run VBA, or prove numerical behavior.


## Small CLI surface

- `step3 query --analysis DIR --target TARGET [--sheet SHEET] [--offset N] [--limit N] [--out DIR]`: select an exact name, qualified finite cell/range or field ID. Return bounded complete source records, field summaries, source/run/stage binding, pagination and unknowns. A `vba:MODULE` target may expose the existing canonical inventory module record for static workflow analysis; it does not execute code or infer worksheet ownership. Ambiguity needs explicit context. Excluded worksheet logic stays outside direct cell queries.
- `step3 trace --analysis DIR --target TARGET [--sheet SHEET] [--direction upstream|downstream] [--max-depth N] [--max-fields N] [--out DIR]`: traverse checked field dependencies with deliberate prerequisite/consumer direction. Return readable nodes, edge evidence, internal dependencies, unresolved/excluded boundaries and explicit depth/size frontier. Trace does not resolve dynamic formulas or VBA execution.
- `step3 validate --analysis DIR [--spec INPUT.json]`: preserve existing structural validation. An optional Agent-authored spec cites declared evidence packet paths and byte hashes, current binding and stage hashes. Validate only supplied evidence IDs, reconstruct their source/structural facts against canonical artifacts, and render one validated context as `model_spec.json` and `model_spec.md`. Interpretations remain `inferred`; missing decisions remain `pending`; source facts must match checked evidence. No readiness claim becomes true.

Query/trace without `--out` return machine JSON. With `--out`, write paired `query.json/md` or `trace.json/md` from the same packet. Re-running identical input/selection produces identical files. Never write into bound source roots or overwrite structural-stage inputs with packet output. Invalid source/stage/spec/evidence must fail clearly; an invalid semantic input must not replace the last valid outputs.

Use a small fixed semantic section set: field roles, logical axes, calculation groups, outputs, macro workflow, open questions and validation plan. Preserve useful structured details and evidence references. A natural-language inference is not an exact fact claim; validating citations does not prove its business meaning or numerical behavior.

## Targeted exploration and deliverables

1. Bind the requested output selectors, order, scenario, scope and current input boundary.
2. Query the target and relevant source records using explicit selectors and worksheet context when needed.
3. Trace supported upstream or downstream paths and record scope boundaries, frontiers, dynamic references and unresolved evidence.
4. Inspect declared names, tables, controls or VBA records only when they relate to the selected workflow; static association does not prove execution.
5. Prepare source-cited query/trace packets, a paired semantic design, open questions and a scoped validation plan.
6. Keep each run's artifacts under its own source-bound workflow and local `output/` directory. A later source uses the same command contracts with its own targets and evidence.

## Acceptance checks

- Wrong/stale source or stage, ambiguous selectors and out-of-scope cells fail explicitly.
- Finite range/field paging is stable and declares omitted records; absent stored cells are not invented.
- Upstream/downstream and multi-start range traces preserve edge direction, stop on limits/cycles, and retain unknown/excluded frontier.
- Tampered packets, changed facts with updated packet hashes, unknown evidence IDs and stale spec bindings are rejected before replacing valid outputs.
- Fact, inferred and pending statuses are kept distinct. Semantic validation cannot enable runtime/generation/GPU readiness.
- JSON/Markdown pairs share content and provenance. Current source evidence supports every factual conclusion; unresolved interpretations remain explicit.
- Existing structural CLI and full tests continue to pass; a fresh independent reviewer checks realistic acceptance-scoped defects only.
