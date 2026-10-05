# Step 2: Artifact Index

Step 2 reads current Step 1 handoffs (`step1.v1` or `step1.batch.v1`) and writes a navigation and quality index. It keeps every batch entry, including failed entries and entries whose per-source handoff is unavailable. It stores source identity, status, metrics, diagnostics, next actions, and references; artifact contents are not copied and no semantic conclusions are made.

```text
excel-to-act step2 tools
excel-to-act step2 index --handoff PATH --step1-root DIR [--out DIR]
excel-to-act step2 validate --index PATH --step1-root DIR
```

`step2 tools` lists the resolver, index builder, and saved-index validator. `step2 index` writes deterministic `index.json` and `INDEX.md` to `output/step2_index/` by default, validates its references, and exits nonzero when integrity checks fail. The index's `status` describes the Step 1 entries; `validation_status` describes saved-index integrity. `step2 validate` checks the input handoff reference and each artifact path, checksum, JSON, and known schema. Unknown JSON content is parsed but left schema-opaque. Diagnostics name the source entry and artifact and include a source location when the artifact carries one.

`--handoff` may be absolute or relative to the current directory or `--step1-root`, but must resolve inside the Step 1 root. Batch entry `handoff_path` values resolve from that root. Artifact paths resolve from `artifact_paths_relative_to`, falling back to `run_path`; a single source handoff uses the same bases. References are stored relative to the Step 1 root. `final_output` is kept as metadata and does not change artifact path resolution.

Paths are resolved under `--step1-root`; absolute and escaping paths are blocked. Known JSON artifacts use the existing Step 1 contract checks or Pydantic models. The validator recommends a `next_tool` only when the Step 1 catalogue names a producer for that artifact.
