# verify

Completeness verification: prove the Step 1 output accounts for the input workbook.

`verify_completeness()` re-derives the object universe straight from the package
(worksheet XML cell counts + package part enumeration) instead of trusting
`WorkbookInventory.coverage`, whose `discovered` is an identity and therefore
cannot fail. A gap is reported with its `SourceLocation` whenever something
present in the workbook is missing from the output.

Status escalates to `fail` only on `error`-severity checks, which makes the CLI
exit non-zero; `warning` and `info` are recorded for the handoff's next actions.
