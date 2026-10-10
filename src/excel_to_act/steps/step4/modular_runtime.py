"""Small standalone Excel-value helpers for source-bound family output."""

from __future__ import annotations

from dataclasses import dataclass
import math
from typing import Any, Callable, Iterable


class CalculationBlocked(RuntimeError):
    """A source-bound model reached an input or feature it cannot represent."""


@dataclass(frozen=True)
class ExcelError(Exception):
    code: str

    def __post_init__(self) -> None:
        Exception.__init__(self, self.code)


@dataclass(frozen=True)
class MissingSourceValue:
    """A declared source/model slot that was not populated by the active plan."""

    address: str


@dataclass(frozen=True)
class TableView:
    rows: tuple[tuple[Any, ...], ...]

    def column(self, index: int) -> tuple[Any, ...]:
        if index < 1 or not self.rows or index > len(self.rows[0]):
            raise ExcelError("#REF!")
        return tuple(row[index - 1] for row in self.rows)


@dataclass(frozen=True)
class ColumnOverlayTable:
    """A source table whose derived columns are requested only when selected."""

    row_count: int
    column_count: int
    columns: dict[int, Any]

    def column(self, index: int) -> list[Any]:
        if not isinstance(index, int) or isinstance(index, bool) or not 1 <= index <= self.column_count:
            raise ExcelError("#REF!")
        if index not in self.columns:
            raise CalculationBlocked(f"table column {index} is unavailable in the approved saved-case scope")
        values = _vector(force(self.columns[index]))
        if len(values) != self.row_count:
            raise CalculationBlocked(
                f"table column {index} has {len(values)} values; expected {self.row_count}")
        return values

    def cell(self, index: int, row: int) -> Any:
        """Read one selected table cell without forcing unrelated return rows."""
        if (not isinstance(index, int) or isinstance(index, bool)
                or not 1 <= index <= self.column_count):
            raise ExcelError("#REF!")
        if index not in self.columns:
            raise CalculationBlocked(f"table column {index} is unavailable in the approved saved-case scope")
        if not isinstance(row, int) or isinstance(row, bool) or not 0 <= row < self.row_count:
            raise ExcelError("#REF!")
        column = self.columns[index]
        if isinstance(column, (list, tuple)):
            if len(column) != self.row_count:
                raise CalculationBlocked(
                    f"table column {index} has {len(column)} values; expected {self.row_count}")
            return literal(force(column[row]))
        # Keep support for existing whole-column providers. Source adapters use
        # per-cell providers below so unselected header/missing cells stay lazy.
        return self.column(index)[row]


@dataclass(frozen=True)
class LogicalReference:
    """A source-mapped view used only by Excel reference helpers."""

    data: Any
    start: tuple[int, ...]
    shape: tuple[int, ...]
    # A logical vector can originate from either a worksheet row or column.
    # Retain the source range orientation for operations such as TRANSPOSE.
    source_shape: tuple[int, int] | None = None


def force(value: Any) -> Any:
    return value() if callable(value) else value


def literal(value: Any) -> Any:
    if isinstance(value, MissingSourceValue):
        raise CalculationBlocked(f"active source or formula value is missing: {value.address}")
    if isinstance(value, dict) and "excel_error" in value:
        return ExcelError(str(value["excel_error"]))
    return value


def logical_reference(data: Any, start: tuple[int, ...], shape: tuple[int, ...],
                      source_shape: tuple[int, int] | None = None) -> LogicalReference:
    return LogicalReference(data, start, shape, source_shape)


def table_view(rows: list[list[Any]]) -> TableView:
    """Materialize a small cross-variable table from already mapped scalar cells."""
    return TableView(tuple(tuple(reference_value(item) for item in row) for row in rows))


def _nested_at(value: Any, indices: tuple[int, ...]) -> Any:
    for index in indices:
        value = value[index]
    return value


def _reference_values(reference: LogicalReference) -> list[Any]:
    if not reference.shape:
        return [literal(_nested_at(reference.data, reference.start))]
    if len(reference.shape) == 1:
        start = reference.start[0]
        return [literal(reference.data[start + index]) for index in range(reference.shape[0])]
    if len(reference.shape) == 2:
        start_row, start_col = reference.start
        rows, cols = reference.shape
        return [literal(reference.data[row][col])
                for row in range(start_row, start_row + rows)
                for col in range(start_col, start_col + cols)]
    raise CalculationBlocked("logical references support at most two axes")


def reference_value(value: Any) -> Any:
    if not isinstance(value, LogicalReference):
        return literal(value)
    values = _reference_values(value)
    return values[0] if not value.shape else values


def range_values(values: Iterable[Any]) -> list[Any]:
    return [reference_value(value) for value in values]


def range_between(start: Callable[[], Any], end: Callable[[], Any]) -> Any:
    left, right = force(start), force(end)
    if not isinstance(left, LogicalReference) or not isinstance(right, LogicalReference):
        return ExcelError("#VALUE!")
    if left.data is not right.data or len(left.start) != len(right.start) or len(left.start) not in {1, 2}:
        return ExcelError("#REF!")
    if len(left.start) == 1:
        low, high = sorted((left.start[0], right.start[0]))
        size = high - low + 1
        return LogicalReference(left.data, (low,), (size,), (size, 1))
    low_row, high_row = sorted((left.start[0], right.start[0]))
    low_col, high_col = sorted((left.start[1], right.start[1]))
    shape = (high_row - low_row + 1, high_col - low_col + 1)
    return LogicalReference(left.data, (low_row, low_col),
                            shape, shape)


def offset(reference: Callable[[], Any], rows: Callable[[], Any], columns: Callable[[], Any],
           height: Callable[[], Any] | None = None, width: Callable[[], Any] | None = None) -> Any:
    base = force(reference)
    if not isinstance(base, LogicalReference):
        return ExcelError("#VALUE!")
    try:
        row_delta = int(_number(force(rows)))
        col_delta = int(_number(force(columns)))
        if len(base.start) == 1:
            if col_delta:
                return ExcelError("#REF!")
            start = (base.start[0] + row_delta,)
            shape = (int(_number(force(height))) if height is not None else
                     base.shape[0] if base.shape else 1,)
        else:
            start = (base.start[0] + row_delta, base.start[1] + col_delta)
            shape = (int(_number(force(height))) if height is not None else base.shape[0],
                     int(_number(force(width))) if width is not None else base.shape[1])
        if any(size < 1 for size in shape):
            return ExcelError("#VALUE!")
        if len(shape) == 1 and base.source_shape is not None and base.source_shape[0] == 1:
            source_shape = (1, shape[0])
        else:
            source_shape = ((shape[0], 1) if len(shape) == 1 else (shape[0], shape[1]))
        result = LogicalReference(base.data, start, shape, source_shape)
        _reference_values(result)
        return result
    except (ExcelError, CalculationBlocked, IndexError, TypeError):
        return ExcelError("#REF!")


def lookup_name(values: dict[str, Any], name: str) -> Any:
    names = values.get("__names__", {})
    if not isinstance(name, str) or not isinstance(names, dict):
        raise CalculationBlocked(f"dynamic source name is not mapped: {name!r}")
    binding = names.get(name.casefold())
    if not isinstance(binding, dict):
        raise CalculationBlocked(f"dynamic source name is not mapped: {name!r}")
    kind = binding.get("kind")
    if kind == "variable":
        variable_id = binding.get("variable_id")
        if not isinstance(variable_id, str) or variable_id not in values:
            raise CalculationBlocked(f"dynamic source name has no logical variable: {name!r}")
        return values[variable_id]
    if kind == "coordinate":
        variable_id, indices = binding.get("variable_id"), binding.get("indices")
        if not isinstance(variable_id, str) or not isinstance(indices, list):
            raise CalculationBlocked(f"dynamic source name has an invalid logical coordinate: {name!r}")
        return at(values, variable_id, indices)
    if kind == "source_literal":
        address = binding.get("address")
        if not isinstance(address, str):
            raise CalculationBlocked(f"dynamic source name has no declared source address: {name!r}")
        source_values = values.get("__source_values__", {})
        return literal(source_values.get(address, MissingSourceValue(address)))
    if kind == "vector":
        return [_binding_item(values, item) for item in binding.get("items", [])]
    if kind == "table":
        columns = binding.get("columns", {})
        if not isinstance(columns, dict):
            raise CalculationBlocked(f"dynamic source table has invalid columns: {name!r}")
        return ColumnOverlayTable(
            row_count=binding.get("row_count", 0),
            column_count=binding.get("column_count", 0),
            columns={int(column): [
                (lambda item=item: _binding_item(values, item)) for item in items]
                     for column, items in columns.items()},
        )
    raise CalculationBlocked(f"dynamic source name is unsupported in the saved scope: {name!r}")


def source_literal(values: dict[str, Any], address: str) -> Any:
    """Read one source-declared metadata literal without adding a model variable."""
    source_values = values.get("__source_values__", {})
    if not isinstance(address, str) or not isinstance(source_values, dict):
        raise CalculationBlocked("source metadata literal has an invalid address ledger")
    return literal(source_values.get(address, MissingSourceValue(address)))


def adapter_metadata(values: dict[str, Any], group_id: str, indices: tuple[int, ...]) -> Any:
    """Read a source-bound metadata slot by logical group/index, never by cell address."""
    groups = values.get("__adapter_metadata__", {})
    if not isinstance(group_id, str) or not isinstance(groups, dict) or group_id not in groups:
        raise CalculationBlocked(f"adapter metadata group is missing: {group_id!r}")
    value = groups[group_id]
    for index in indices:
        if not isinstance(index, int) or isinstance(index, bool):
            raise CalculationBlocked(f"adapter metadata index is invalid: {index!r}")
        try:
            value = value[index]
        except (IndexError, KeyError, TypeError) as exc:
            raise CalculationBlocked(f"adapter metadata index is outside {group_id}: {index}") from exc
    return literal(value)


def adapter_metadata_reference(values: dict[str, Any], group_id: str, start: tuple[int, ...],
                               shape: tuple[int, ...],
                               source_shape: tuple[int, int] | None = None) -> LogicalReference:
    groups = values.get("__adapter_metadata__", {})
    if not isinstance(groups, dict) or group_id not in groups:
        raise CalculationBlocked(f"adapter metadata group is missing: {group_id!r}")
    return logical_reference(groups[group_id], start, shape, source_shape)


def _binding_item(values: dict[str, Any], item: dict[str, Any]) -> Any:
    metadata_group = item.get("adapter_metadata_group")
    if isinstance(metadata_group, str):
        indices = item.get("indices", [])
        if not isinstance(indices, list):
            raise CalculationBlocked(f"named table metadata binding is invalid: {metadata_group}")
        return adapter_metadata(values, metadata_group, tuple(indices))
    variable_id = item.get("variable_id")
    if isinstance(variable_id, str):
        indices = item.get("indices", [])
        if not isinstance(indices, list) or variable_id not in values:
            raise CalculationBlocked(f"named table variable binding is invalid: {variable_id}")
        return at(values, variable_id, indices)
    address = item.get("address") or item.get("missing_address")
    if not isinstance(address, str):
        raise CalculationBlocked("named table cell binding has no logical variable or source address")
    source_values = values.get("__source_values__", {})
    return literal(source_values.get(address, MissingSourceValue(address)))


def at(values: dict[str, Any], variable_id: str, indices: Iterable[int]) -> Any:
    if variable_id not in values:
        raise CalculationBlocked(f"logical variable is missing: {variable_id}")
    value = values[variable_id]
    for index in indices:
        if not isinstance(index, int) or isinstance(index, bool):
            raise CalculationBlocked(f"logical index is invalid for {variable_id}: {index!r}")
        try:
            value = value[index]
        except (IndexError, KeyError, TypeError) as exc:
            raise CalculationBlocked(f"logical index is outside {variable_id}: {index}") from exc
    return literal(value)


def _error(value: Any) -> bool:
    return isinstance(reference_value(value), ExcelError)


def _number(value: Any) -> float:
    value = reference_value(value)
    if _error(value):
        raise value
    if value is None or value == "":
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    raise ExcelError("#VALUE!")


def _truth(value: Any) -> bool:
    value = reference_value(value)
    if _error(value):
        raise value
    if value is None or value == "":
        return False
    if isinstance(value, str):
        return value.casefold() not in {"false", "0", ""}
    return bool(value)


def _first_error(values: Iterable[Any]) -> ExcelError | None:
    return next((value for value in values if _error(value)), None)


def _flatten(value: Any) -> list[Any]:
    value = force(value)
    if isinstance(value, LogicalReference):
        return _reference_values(value)
    if isinstance(value, TableView):
        return [item for row in value.rows for item in row]
    if isinstance(value, (list, tuple)):
        result: list[Any] = []
        for item in value:
            result.extend(_flatten(item))
        return result
    return [literal(value)]


def binary(operator: str, left: Callable[[], Any], right: Callable[[], Any]) -> Any:
    a = reference_value(force(left))
    if _error(a):
        return a
    b = reference_value(force(right))
    if _error(b):
        return b
    if operator == "&":
        return ("" if a is None else str(a)) + ("" if b is None else str(b))
    try:
        if operator in {"=", "<>", "<", ">", "<=", ">="}:
            # A referenced blank cell compares as zero in numeric formulas.
            if a is None and b is None:
                a, b = 0, 0
            elif a is None and isinstance(b, (int, float, bool)):
                a = 0
            elif b is None and isinstance(a, (int, float, bool)):
                b = 0
            if isinstance(a, (int, float, bool)) and isinstance(b, (int, float, bool)):
                a, b = float(a), float(b)
            return {"=": lambda: a == b, "<>": lambda: a != b,
                    "<": lambda: a < b, ">": lambda: a > b,
                    "<=": lambda: a <= b, ">=": lambda: a >= b}[operator]()
        x, y = _number(a), _number(b)
        if operator == "+":
            return x + y
        if operator == "-":
            return x - y
        if operator == "*":
            return x * y
        if operator == "/":
            return ExcelError("#DIV/0!") if y == 0 else x / y
        if operator == "^":
            try:
                return x ** y
            except (OverflowError, ValueError):
                return ExcelError("#NUM!")
    except ExcelError as exc:
        return exc
    return ExcelError("#VALUE!")


def unary(operator: str, value: Callable[[], Any]) -> Any:
    value = reference_value(force(value))
    if _error(value):
        return value
    try:
        number = _number(value)
        if operator == "+":
            return number
        if operator == "-":
            return -number
        if operator == "%":
            return number / 100
    except ExcelError as exc:
        return exc
    return ExcelError("#VALUE!")


def excel_if(test: Callable[[], Any], yes: Callable[[], Any], no: Callable[[], Any]) -> Any:
    try:
        return force(yes) if _truth(force(test)) else force(no)
    except ExcelError as exc:
        return exc


def excel_iferror(value: Callable[[], Any], fallback: Callable[[], Any]) -> Any:
    try:
        result = force(value)
        return force(fallback) if _error(result) else result
    except ExcelError:
        return force(fallback)


def excel_ifna(value: Callable[[], Any], fallback: Callable[[], Any]) -> Any:
    try:
        result = force(value)
        return force(fallback) if _error(result) and result.code == "#N/A" else result
    except ExcelError as exc:
        return force(fallback) if exc.code == "#N/A" else exc


def minimum(*values: Callable[[], Any]) -> Any:
    flattened = [item for value in values for item in _flatten(value)]
    error = _first_error(flattened)
    if error:
        return error
    numbers = [_number(item) for item in flattened if isinstance(item, (int, float, bool))]
    return min(numbers) if numbers else 0


def maximum(*values: Callable[[], Any]) -> Any:
    flattened = [item for value in values for item in _flatten(value)]
    error = _first_error(flattened)
    if error:
        return error
    numbers = [_number(item) for item in flattened if isinstance(item, (int, float, bool))]
    return max(numbers) if numbers else 0


def excel_sum(*values: Callable[[], Any]) -> Any:
    flattened = [item for value in values for item in _flatten(value)]
    error = _first_error(flattened)
    if error:
        return error
    try:
        return sum(_number(item) for item in flattened if isinstance(item, (int, float, bool)))
    except ExcelError as exc:
        return exc


def _vector(value: Any) -> list[Any]:
    return _flatten(value)


def sumproduct(*values: Callable[[], Any]) -> Any:
    vectors = [_vector(force(value)) for value in values]
    if not vectors:
        return 0
    if any(len(vector) != len(vectors[0]) for vector in vectors):
        return ExcelError("#VALUE!")
    total = 0.0
    for row in zip(*vectors):
        error = _first_error(row)
        if error:
            return error
        try:
            product = math.prod(_number(value) for value in row)
        except ExcelError as exc:
            return exc
        total += product
    return total


def excel_and(*values: Callable[[], Any]) -> Any:
    for value in values:
        try:
            if not _truth(force(value)):
                return False
        except ExcelError as exc:
            return exc
    return True


def excel_or(*values: Callable[[], Any]) -> Any:
    saw_error: ExcelError | None = None
    for value in values:
        try:
            if _truth(force(value)):
                return True
        except ExcelError as exc:
            saw_error = saw_error or exc
    return saw_error or False


def _equal(left: Any, right: Any) -> bool:
    left, right = reference_value(left), reference_value(right)
    if isinstance(left, str) and isinstance(right, str):
        return left.casefold() == right.casefold()
    if isinstance(left, (int, float)) and not isinstance(left, bool) and isinstance(right, (int, float)) and not isinstance(right, bool):
        return float(left) == float(right)
    return left == right


def match(lookup: Callable[[], Any], values: Callable[[], Any], mode: Callable[[], Any] | None = None) -> Any:
    key = reference_value(force(lookup))
    source = _vector(force(values))
    try:
        match_mode = int(_number(force(mode))) if mode is not None else 1
    except ExcelError as exc:
        return exc
    error = _first_error([key, *source])
    if error:
        return error
    if match_mode == 0:
        for index, value in enumerate(source, 1):
            if _equal(value, key):
                return index
        return ExcelError("#N/A")
    candidates = [(index, value) for index, value in enumerate(source, 1)
                  if isinstance(value, (int, float)) and not isinstance(value, bool)]
    if isinstance(key, (int, float)) and not isinstance(key, bool):
        choices = [(index, value) for index, value in candidates if value <= key] if match_mode == 1 else [
            (index, value) for index, value in candidates if value >= key]
        if not choices:
            return ExcelError("#N/A")
        return (max if match_mode == 1 else min)(choices, key=lambda pair: pair[1])[0]
    return ExcelError("#N/A")


def _table_rows(table: Any) -> tuple[tuple[Any, ...], ...]:
    table = literal(table)
    if _error(table):
        raise table
    if isinstance(table, TableView):
        return table.rows
    if isinstance(table, (list, tuple)) and table and isinstance(table[0], (list, tuple)):
        return tuple(tuple(row) for row in table)


def _lookup_row(keys: list[Any], key: Any, approximate: bool) -> int | ExcelError:
    if not keys:
        return ExcelError("#N/A")
    if not approximate:
        selected = next((index for index, candidate in enumerate(keys) if _equal(candidate, key)), -1)
        return selected if selected >= 0 else ExcelError("#N/A")
    if isinstance(key, (int, float)) and not isinstance(key, bool):
        numeric = [(index, candidate) for index, candidate in enumerate(keys)
                   if isinstance(candidate, (int, float)) and not isinstance(candidate, bool)]
        choices = [(index, candidate) for index, candidate in numeric if candidate <= key]
        if not choices:
            return ExcelError("#N/A")
        return max(choices, key=lambda pair: pair[1])[0]
    return ExcelError("#N/A")
    raise ExcelError("#VALUE!")


def vlookup(lookup: Callable[[], Any], table: Callable[[], Any], column: Callable[[], Any],
            approximate: Callable[[], Any] | None = None) -> Any:
    key = reference_value(force(lookup))
    table_value = literal(force(table))
    try:
        col = int(_number(force(column)))
        approx = _truth(force(approximate)) if approximate is not None else True
    except ExcelError as exc:
        return exc
    error = _first_error([key, table_value])
    if error:
        return error
    if isinstance(table_value, ColumnOverlayTable):
        try:
            keys = table_value.column(1)
            selected = _lookup_row(keys, key, approx)
            if isinstance(selected, ExcelError):
                return selected
            result = reference_value(table_value.cell(col, selected))
            return 0 if result is None or result == "" else result
        except ExcelError as exc:
            return exc
    try:
        rows = _table_rows(table_value)
    except ExcelError as exc:
        return exc
    if not rows or col < 1 or col > len(rows[0]):
        return ExcelError("#REF!")
    keys = [row[0] for row in rows]
    selected = _lookup_row(keys, key, approx)
    if isinstance(selected, ExcelError):
        return selected
    result = reference_value(rows[selected][col - 1])
    return 0 if result is None or result == "" else result


def transpose(value: Callable[[], Any]) -> Any:
    rows = force(value)
    if isinstance(rows, ExcelError):
        return rows
    if isinstance(rows, ColumnOverlayTable):
        raise CalculationBlocked("TRANSPOSE cannot materialize an unresolved lazy table overlay")
    if isinstance(rows, LogicalReference):
        # Excel returns zero when TRANSPOSE materializes a truly blank cell
        # reference. Keep source_values unchanged; this coercion belongs only
        # at the reference-to-array formula boundary.
        flat = [0 if item is None else item for item in _reference_values(rows)]
        if not rows.shape:
            normalized = [[flat[0]]]
        else:
            source_shape = rows.source_shape
            if source_shape is None:
                source_shape = ((len(flat), 1) if len(rows.shape) == 1 else
                                (rows.shape[0], rows.shape[1]))
            source_rows, source_columns = source_shape
            if (source_rows < 1 or source_columns < 1
                    or source_rows * source_columns != len(flat)):
                raise CalculationBlocked("logical reference source orientation does not match its values")
            normalized = [flat[row * source_columns:(row + 1) * source_columns]
                          for row in range(source_rows)]
    elif isinstance(rows, TableView):
        # A TableView contains worksheet references, just like a logical
        # reference. Excel materializes blank references as zero in TRANSPOSE.
        normalized = [[0 if item is None else item for item in row] for row in rows.rows]
    elif isinstance(rows, tuple):
        normalized = [list(row) if isinstance(row, (tuple, list)) else [row] for row in rows]
    elif not isinstance(rows, list):
        normalized = [[rows]]
    else:
        normalized = [list(row) if isinstance(row, (list, tuple)) else [row] for row in rows]
    if not normalized:
        return []
    if any(len(row) != len(normalized[0]) for row in normalized):
        return ExcelError("#VALUE!")
    return [list(column) for column in zip(*normalized)]


def indirect(value: Callable[[], Any], resolve: Callable[[str], Any]) -> Any:
    name = force(value)
    if _error(name):
        return name
    if not isinstance(name, str):
        return ExcelError("#REF!")
    try:
        return resolve(name)
    except CalculationBlocked:
        raise
    except Exception:
        return ExcelError("#REF!")


def array(value: Any) -> list[Any]:
    return [force(item) for item in value]


def unsupported_reference(reference: str) -> Any:
    raise CalculationBlocked(f"active formula reference is not mapped: {reference}")


def unsupported_name(name: str) -> Any:
    raise CalculationBlocked(f"active formula name is not mapped: {name}")


def unsupported_range() -> Any:
    raise CalculationBlocked("active dynamic range was not lowered by the approved semantic plan")
