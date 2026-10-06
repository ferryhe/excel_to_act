# schemas

Runtime Pydantic artifact contracts for Phase 1 live here.

The top-level `schemas/` directory is reserved for exported JSON Schema files and contract documentation generated from these models.

`schemas/workbook_view.schema.json` is the exported schema for `excel_to_act.schemas.WorkbookView`; `tests/test_views.py` checks it against the runtime model schema.
