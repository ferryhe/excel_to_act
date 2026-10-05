# Step 2: Artifact Index

Step 2 reads current Step 1 handoffs (`step1.v1` or `step1.batch.v1`) and writes a navigation and quality index. It keeps every batch entry, including failed entries and entries whose per-source handoff is unavailable. It stores source identity, status, metrics, diagnostics, next actions, and references; artifact contents are not copied and no semantic conclusions are made.

```text
excel-to-act step2 tools
excel-to-act step2 index --handoff PATH --step1-root DIR [--out DIR]
```

`step2 tools` lists the initial resolver and index-builder actions. `step2 index` writes deterministic `index.json` and `INDEX.md` to `output/step2_index/` by default. Use `--out` to choose another directory. A valid input with one or more entries exits successfully with index status `partial`; an invalid root handoff, invalid Step 1 root, or empty batch is blocked and exits nonzero. Missing per-source handoffs remain in a usable batch index with error diagnostics.

`--handoff` may be absolute or relative to the current directory or `--step1-root`, but must resolve inside the Step 1 root. Batch entry `handoff_path` values resolve from that root. Artifact paths resolve from `artifact_paths_relative_to`, falling back to `run_path`; a single source handoff uses the same bases. References are stored relative to the Step 1 root. `final_output` is kept as metadata and does not change artifact path resolution.

The first release checks handoff shape and path inputs while building. `validation_status` remains `not_run`; artifact existence, checksum, and full saved-index integrity validation belong to STEP2-02.
