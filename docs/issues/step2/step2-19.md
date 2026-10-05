**CLI:** `excel-to-act step2 validate --index PATH --step1-root DIR`

## Goal

Provide a repeatable integrity check for a saved Step 2 index and the Step 1 artifacts it references.

## Scope

- Check that every referenced path exists under the supplied Step 1 output root.
- Recompute SHA-256 and compare it with the handoff reference.
- Validate artifact JSON against a known schema when one exists; check source hash and run ID when the artifact carries those fields.
- Report missing, malformed, or mismatched references with the source entry, artifact name, and available source location.
- Return structured status, diagnostics, metrics, `retryable`, and a specific `next_tool` where a known repair is available.
- Wire the same validator into `step2 index` after index construction, so the normal command follows resolve → build → validate; retain `step2 validate` as the standalone recheck command.
- Add the validator to the `step2 tools` catalogue. The generic action dispatcher is delivered in STEP2-03.

## Files

- `src/excel_to_act/steps/step2/workflow.py`
- `src/excel_to_act/interfaces/cli.py`
- `src/excel_to_act/steps/step2/README.md`
- `tests/test_step2_index.py` and `tests/test_step2_cli.py`

## Acceptance criteria

1. The CLI detects a missing artifact, changed checksum, invalid JSON/schema, and a supported source/run identity mismatch.
2. Diagnostics identify the exact source entry and artifact.
3. A valid index passes; errors that prevent safe downstream use return a nonzero exit code.
4. `step2 index` uses the shared validator and returns the same validation result shape as `step2 validate`.
5. A diagnostic only recommends a repair when an existing Step 1 or Step 2 tool can address it; opaque content remains unresolved.

## Dependency and boundary

- Depends on #18's index contract and reference format.
- Does not change Step 1 artifacts or claim that checksum validation proves semantic correctness.
