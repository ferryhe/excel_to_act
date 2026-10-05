**CLI:** `excel-to-act step2 agent`; final `excel-to-act step2 tools` catalogue

## Goal

Make the Step 2 workflow usable from an installed package without requiring repository-only instructions or repeated rediscovery of the tool sequence.

## Scope

- Write `src/excel_to_act/steps/step2/agent.md` with the normal workflow, status interpretation, progressive tool disclosure, retry bounds, and stop conditions. Explain how the index hands off to a later analysis agent: use it to select and progressively load relevant Step 1 artifacts, then choose analysis tools and loop on their results.
- Add `excel-to-act step2 agent` to print the packaged contract.
- Complete `step2 tools` metadata for all delivered Step 2 actions, including exact inputs, outputs, and next tools.
- Update the Step 2 README and root English/Chinese READMEs to show implemented status, commands, and output paths.
- Update `docs/issues/step2/README.md` and the four issue drafts with the published GitHub issue links and dependency order.
- Verify the installed package includes `agent.md` and command help reflects the delivered interface.

## Files

- `src/excel_to_act/steps/step2/agent.md`
- `src/excel_to_act/steps/step2/README.md`
- `src/excel_to_act/steps/step2/workflow.py`
- `src/excel_to_act/interfaces/cli.py`
- `pyproject.toml`
- `README.md` and `README.zh-CN.md`
- `docs/issues/step2/README.md` and `docs/issues/step2/step2-*.md`

## Acceptance criteria

1. `step2 agent` prints the packaged workflow contract after installation.
2. `pyproject.toml` includes `steps/step2/agent.md` as package data, and an installed-package check confirms that `step2 agent` can read and print it.
3. `step2 tools` lists every available action and no unavailable command.
4. The agent contract follows `step2.index` → `step2.validate` → bounded repair/resume and preserves all batch entries.
5. The contract tells the downstream analysis agent to use index references and quality metadata as navigation/context, load detailed artifacts only as needed, and keep analysis conclusions separate from Step 2 indexing.
6. Both root READMEs distinguish the new batch index from the legacy per-workbook `artifact_index.json`.
7. The issue plan links all four published GitHub issues in dependency order and retains each issue's numbered acceptance criteria.

## Dependencies and boundary

- Depends on #18 through #20 so the packaged instructions match working commands and response fields.
- Legacy `inspect` handoff support remains a separately scoped compatibility decision.
