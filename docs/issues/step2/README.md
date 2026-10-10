# Step 2 issue plan

Issues are ordered by dependency. Drafts preserve the published issue bodies and numbered acceptance criteria.

## [#18 feat(step2): index Step 1 handoffs](https://github.com/ferryhe/excel_to_act/issues/18)

- Dependency order: None.
- Draft: [step2-18.md](step2-18.md)

## [#19 feat(step2): validate indexed artifact references](https://github.com/ferryhe/excel_to_act/issues/19)

- Dependency order: #18.
- Draft: [step2-19.md](step2-19.md)

## [#20 feat(step2): resume indexing with bounded recovery](https://github.com/ferryhe/excel_to_act/issues/20)

- Dependency order: #18, #19.
- Draft: [step2-20.md](step2-20.md)

## [#21 feat(step2): package the Step 2 agent contract](https://github.com/ferryhe/excel_to_act/issues/21)

- Dependency order: #18, #19, #20.
- Draft: [step2-21.md](step2-21.md)

## Progressive exploration CLI

These three issues extend the completed Step 2 foundation above. Implement them in dependency order, with one PR per issue. Each issue defines its CLI contract, deliverables, exclusions, and acceptance checks. The reviewed proposal is [progressive_excel_exploration.md](../../design/progressive_excel_exploration.md).

### [#30 feat(cli): step2 prepare — Generate a Reusable Reading Package and Dual-Format Handoffs](https://github.com/ferryhe/excel_to_act/issues/30)

- Deliverables: reusable manifest, scoped view files and graph, dependency audit/snapshot, paired English human and AI handoffs.
- Dependency order: completed foundation; verify availability of the locally reviewed Step 1 control/VBA changes before implementation.
- Draft: [step2-prepare.md](step2-prepare.md)

### [#31 feat(cli): step2 query — Read Scoped Evidence and Validate Agent Citations](https://github.com/ferryhe/excel_to_act/issues/31)

- Deliverables: eight query selectors, bounded evidence packets, scope enforcement, source citation validation, and agent reading instructions.
- Dependency order: #30.
- Draft: [step2-query.md](step2-query.md)

### [#32 feat(cli): step2 trace — Trace Dependencies with Bounds and Complete a Real Exploration Workflow](https://github.com/ferryhe/excel_to_act/issues/32)

- Deliverables: bounded static dependency traversal, name/range resolution, replayable source-backed CLI fixture, and independently reviewed exploration evidence.
- Dependency order: #30, #31.
- Draft: [step2-trace.md](step2-trace.md)
