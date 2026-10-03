# inventory

Complete workbook inventory: sheets, cells, ranges, names, tables, validation, conditional formatting, comments, hyperlinks, protection, layout, and cached formula values.

Cached values come from a second `data_only=True` pass (`ingest/cached_values.py`). A cell whose value Excel never stored gets `cached_value=None` and `cached_value_available=False`; when no formula cell in the workbook has a stored result, extraction emits a single `missing_cached_values` warning instead of failing.
