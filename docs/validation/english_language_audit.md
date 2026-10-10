# English-language audit and conversion ledger

## Policy and scope

Keep repository-root `README.md` and `README.zh-CN.md` as the bilingual entry points. Nested README files follow the English artifact policy. All other authored instructions, documentation, code comments, report narration, analysis descriptions and future deliverables use English. Source values, formulas, identifiers and immutable provenance remain exact data; translating them would invalidate the project's fidelity and evidence contracts. Unicode test payloads remain logically unchanged and can be written with escapes in English source code.

The audit covers tracked/untracked project text, ignored local analysis deliverables and generators, local notes, and Graphify working artifacts. Build outputs and caches are regenerated or checked against their sources. Binary workbooks and preserved package/source evidence are not prose to translate. JSON is checked after decoding as well as for literal non-English characters, so escaped prose cannot hide a remaining language defect.

Pre-implementation assessment: complex. Source instructions and their resource assertions are small changes, but existing generated specifications, saved command results and handoff hashes must stay consistent while canonical workbook facts and source-bound decisions remain exact. Keep the persistent Sol/high worker for source/document changes; the controller handles artifact narration, generator translation and regeneration. Obtain a fresh independent Sol/high acceptance review. Copilot is user-disabled; no publication is authorized.

## Initial source differences, before edits

| Order | File | Difference | Required action |
| ---: | --- | --- | --- |
| 1 | `src/excel_to_act/steps/step3/agent.md` | Chinese contract prose | Translate the whole contract; preserve the output order, path/error boundaries, coverage definitions and readiness limits; add the English-artifact policy |
| 2 | `tests/test_step3_workflow.py` | Chinese resource assertions | Match the English contract without changing the tested behavior |
| 3 | `docs/step3_cli.md` | Chinese guide formerly named `docs/step3_cli.zh-CN.md` | Translate and rename to `docs/step3_cli.md`; update all project links |
| 4 | `docs/plans/step3_result_driven.md` | Chinese workflow/tool plan | Translate without changing supported/proposed distinctions or coverage scope |
| 5 | `docs/validation/pricing_step3.md` | Chinese source-backed analysis | Translate narration; retain exact source records, counts, formulas and limits |
| 6 | `docs/validation/pricing_step3_exploration.md` | Chinese exploration report | Translate narration and preserve the provenance and known/inferred/pending distinctions |
| 7 | `docs/validation/pricing_gp_result_driven.md` | Chinese GP options/report | Translate; preserve known-only ratios and zero implementation/runtime coverage |
| 8 | `docs/issues/step2/step2-prepare.md` | Chinese proposed issue body | Translate all requirements, acceptance cases and limits |
| 9 | `docs/issues/step2/step2-query.md` | Chinese proposed issue body | Translate all requirements, acceptance cases and limits |
| 10 | `docs/issues/step2/step2-trace.md` | Chinese proposed issue body | Translate all requirements, acceptance cases and limits |
| 11 | `docs/design/progressive_excel_exploration.md` | Already English when individually audited | Preserve its bytes and record the unchanged hash |
| 12 | `docs/design/fidelity_rules.md` | Literal Unicode example inside English prose | Describe the code points in English while preserving the example's meaning |
| 13 | `tests/test_step1.py` | Unicode fixture strings, not Chinese instructions | Use escapes and verify that runtime constants/fixture bytes stay unchanged |

Step1 and Step2 agent contracts are already English. Add the same artifact-language rule to prevent future narration from reverting to another language. README language stays intact; only affected links/titles need updates.

## Local-artifact differences

The initial literal-character scan returned 275 text-file matches before classifying README, source/provenance, authored analysis and local notes. Detailed files, decoded JSON pointers, counts and source-data exclusions are recorded under `output/english_language_audit/`. This count is not the number of files that may safely be translated.

Process each eligible file separately and record its before/after hash and remaining findings. Translate generators before rebuilding active specifications and paired reports. Regenerate handoff/model-spec references with existing validators instead of inventing checksums or changing canonical facts. Historical imported evidence remains verbatim and is listed explicitly rather than silently counted as an English deliverable.

## Completion gates

- No non-English authored prose remains outside README, including decoded JSON, generators, agent contracts and local reports.
- Source/provenance exceptions are individually identified; source files and canonical fact objects remain unchanged.
- English resource/CLI behavior and wheel packaging match; the Unicode tests still exercise the same payloads.
- Both current analysis/spec pairs and handoff hashes validate after regeneration; source identities, numeric values, formulas, path IDs, coverage counts and readiness remain unchanged.
- The guide rename has no stale project references. Bilingual README content remains configured as before.
- Tests appropriate to the affected callers pass, followed by a fresh independent acceptance review.

## Final results

English conversion is complete. [The file-by-file difference table](../../output/english_language_audit/file_changes.md) records 86 individual files audited, translated, renamed, regenerated or updated. This count includes unchanged audit entries and README link updates; it is not a claim that 86 Chinese files changed. Detailed records are in [source translation](../../output/english_language_audit/source_translation.json), [Step 3 artifacts](../../output/english_language_audit/step3_artifact_translation.json), [local notes](../../output/english_language_audit/local_notes_translation.json) and [historical narration](../../output/english_language_audit/historical_translation.json).

The frozen [residual scan](../../output/english_language_audit/final_residual_report.md) checked 6,244 candidate files, including ignored local deliverables, filenames, literal Han, decoded JSON keys/values and escaped strings. It found **zero authored or unclassified residuals**. The 228 remaining Han-bearing records have individual exception reasons: 185 source/provenance records, 33 source-bound evidence records, 8 fetched issue snapshots and 2 source field maps. The only bilingual prose exceptions are the two repository-root README files. Three Chinese issue-link labels in the nested Step 2 README were translated.

Source exceptions remain explicit in [the residual inventory](../../output/english_language_audit/final_residual_inventory.json). Both 21 MB field maps were parsed and contain only the same original workbook sheet identifier: 145 field values and one summary key per file. Copied VBA facts remain exact. The original scope decision remains hash-bound to existing handoffs and analyses; [its English context and binding proof](../../output/english_language_audit/scope_provenance.json) explain the retained record without asserting an authenticated human quotation. [Issue excerpt proof](../../output/english_language_audit/quoted_issue_provenance.json) matches all 12 captured titles/body values to fetched snapshots. [Rich-text proof](../../output/english_language_audit/source_value_provenance.json) maps the diagnostic value to original shared-string index 1036 at worksheet part `sheet1.xml`, cell `B6`.

Validation is recorded in [validation.json](../../output/english_language_audit/validation.json):

- Full suite: 243 passed; 19 existing oletools/pyparsing warnings. Ruff and `git diff --check` passed. The focused source checks also passed, and Unicode test string constants retain their pre-edit digest.
- All three CLI agent outputs match their English source after normalizing platform newlines; all three packaged wheel resources match source bytes exactly.
- Exploration, Pricing and GP verification passed. All 67 source files and the three broad structural files retain their baseline hashes; the GP copy retains the four core analysis files.
- Of 11 previously snapshotted broad-analysis files, 8 remain byte-identical. Only normalized model-spec JSON/Markdown and its dependent handoff JSON changed. The original snapshot was preserved.
- Both specification pairs and handoff references validate. GP still has 8 registered known-only path units and 4 shared modules, with resolved/implemented/runtime-verified coverage at 0/8 and readiness false. No formula, numeric fact or coverage meaning was changed.
- Twelve Graphify memory filenames now have English slugs with original timestamps. Twelve issue draft bodies match their Markdown, and reflection outputs were regenerated from English memories.

The first mixed-mode Pricing verification normalized an existing difference in the handoff Markdown footer between spec attachment and ordinary validation. Repeating the complete verification on the normalized files passed; the source and model facts stayed exact. The observed failed attempt and subsequent result are retained in the validation record.

Fresh independent Sol/high acceptance [review 2](../../output/english_language_audit/review_2.md) is **PASS**, with no acceptance-scoped defects. It independently checked all 86 after-hashes and the 228 explicit source/provenance exclusions. The earlier review's audit-report consistency finding was corrected before this fresh review. No commit, push, PR, external publication, Excel recalculation, macro execution, numerical implementation or GPU execution was performed.

Some generated artifacts and reflection outputs had no whole-file pre-edit hash captured. Their ledgers state this limitation and record available original value hashes, sizes and current hashes without inventing a baseline. Large canonical source JSON remains explicitly marked as raw-scanned rather than fully decoded; its original-language data is not claimed to be English.
