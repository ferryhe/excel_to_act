# Agent reading contract

An Agent may cite only a record present in its supplied `WorkbookView`, including a `region` view when a narrower recorded range is useful. Region views exist only for valid rectangular `RangeInventory` A1 addresses and contain inventory records fully inside that range; they do not imply that unrecorded cells or layout objects exist. Every claim carries `view_id`, `source_run_id`, `source_schema_version`, `record_id`, and the record's complete `source_location` object. A fact claim's `value` must equal the record's complete `facts` object. Unknown IDs, source run/schema mismatches, missing or altered locations, and changed facts fail validation.

Claim `kind` is one of:

- `fact`: source-exact; include `value` equal to the record facts.
- `derivation`: include a non-empty `derivation` description; this is not a source fact.
- `unverified`: set `verified` to `false`; this is not a source fact.
- `opaque`: use only for an opaque source record, set `reported_opaque` to `true`, and omit `value`. The validator returns it in `opaque_reports`; an Agent must not guess its contents.

Example shared source fields:

```json
{
  "view_id": "<view_id>",
  "source_run_id": "<source_run_id>",
  "source_schema_version": "phase1.v1",
  "record_id": "<record_id>",
  "source_location": {
    "workbook_path": "model.xlsx",
    "sheet_name": "Inputs",
    "sheet_index": 0,
    "address": "B4",
    "object_type": "cell",
    "object_id": "Inputs!B4",
    "ooxml_part": "xl/worksheets/sheet1.xml",
    "source_identity": "<source identity>"
  }
}
```

The runnable examples in `tests/test_views.py` include three valid claims (fact, derivation, and unverified), an opaque report without a guessed value, and invalid cases for missing `source_location`, unknown `view_id`, wrong run, and missing record. The view compiler retains full cross-sheet graph targets and unresolved-reference diagnostics. Consumers must not turn cross-sheet targets into local edges or infer opaque details.

The #11 layer boundary needed here is small: L1 supplies deterministic source facts plus identity and location; a downstream reader may cite those facts or label its own derivation/unverified conclusion. This contract does not define semantic interpretation or numerical verification.
