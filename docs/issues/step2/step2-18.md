**CLI:** `excel-to-act step2 index`; initial `excel-to-act step2 tools`

## Goal

Build a deterministic, agent-facing index from the current Step 1 machine handoffs. Accept one current `step1.v1` source handoff or one `step1.batch.v1` batch handoff and retain every batch entry, including partial and failed inputs. The index is the launch point for a later analysis agent: it supplies stable source/artifact references and quality context so that agent can select and inspect relevant Step 1 artifacts without rescanning the batch. Do not use the legacy `Handoff` Pydantic model as the current Step 1 contract: its schema version and artifact fields differ.

## Scope

- Add typed Step 2 index, source-entry, and artifact-reference contracts.
- Resolve Step 1 paths using the actual layered handoff contract and explicit `--step1-root`.
- Accept `--handoff` as an absolute path or a current-working-directory-relative path, but require it to resolve inside `--step1-root`. Normalize all indexed artifact paths relative to that root.
- Write `index.json` and `INDEX.md` under the requested `--out` directory, defaulting to `output/step2_index/` to match the root README. Use a distinct `--out` directory to retain indexes for multiple batches.
- Record source SHA-256, relative path, run ID, readiness/status, available metrics, artifact names/paths/checksums, diagnostics, and next actions.
- Return a structured result with status, metrics, diagnostics, artifacts, `retryable`, and `next_tool` fields.
- Add an initial tool catalogue containing the resolver and index builder actions.

## Files

- `src/excel_to_act/steps/step2/workflow.py`
- `src/excel_to_act/steps/step2/__init__.py`
- `src/excel_to_act/schemas/artifacts.py` and `schemas/__init__.py`
- `src/excel_to_act/interfaces/cli.py`
- `src/excel_to_act/steps/step2/README.md`
- `tests/test_step2_index.py` and `tests/test_step2_cli.py`

## Acceptance criteria

1. A single-source handoff and a batch handoff both build an index.
2. Tests use the current runtime `step1.v1` and `step1.batch.v1` shapes; the legacy `inspect` handoff remains unsupported.
3. Every batch entry is represented with its original status and source identity; failed entries are not dropped.
4. The index provides enough stable references and status/quality context for a later analysis agent to choose artifacts for agentic analysis; it does not duplicate all artifact contents or make semantic conclusions.
5. For a batch, resolve each entry's `handoff_path` relative to `--step1-root`, load its source handoff, then resolve artifact `path` using `artifact_paths_relative_to` or `run_path`. For a single-source handoff, use the same source-relative artifact base. Preserve `final_output` as separate metadata; it does not rebase candidate artifact references.
6. A top-level `--handoff` outside `--step1-root` is rejected with a blocking diagnostic.
7. Invalid root handoffs or missing required per-source handoffs return a clear blocking diagnostic; entries without a handoff remain present with their original evidence.
8. Repeating the command with unchanged handoff and artifact inputs produces the same index content.
9. `step2 tools` lists the initial tools with commands, inputs, outputs, and next actions.
10. The CLI exits successfully for a valid or usable partial index and nonzero when the handoff cannot be indexed safely.
11. The CLI's first release performs basic build-time input checks. Full saved-index integrity validation and validation from `step2 index` are added in STEP2-02.

## Dependencies and boundaries

- Depends on the existing `step1.v1` and `step1.batch.v1` contracts.
- STEP2-02 adds the standalone integrity check. This issue may perform basic build-time validation needed to produce a coherent index.
- Support for the legacy `inspect` handoff, semantic classification, and automatic interpretation of opaque content are separate work.
