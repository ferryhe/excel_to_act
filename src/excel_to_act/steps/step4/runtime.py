"""Standalone runtime copied into each generated model bundle.

This module deliberately depends only on the Python standard library. The saved
model supplies expressions, raw literals, names, and the bounded formula map.
"""

from __future__ import annotations

from dataclasses import dataclass
import datetime as _datetime
import json
import math
from pathlib import Path
import re
from typing import Any, Callable


class CalculationBlocked(ValueError):
    pass


@dataclass(frozen=True)
class ExcelError:
    code: str


@dataclass(frozen=True)
class CellRange:
    sheet: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def single(self) -> bool:
        return self.min_row == self.max_row and self.min_col == self.max_col

    @property
    def address(self) -> str:
        first = f"{column_letter(self.min_col)}{self.min_row}"
        last = f"{column_letter(self.max_col)}{self.max_row}"
        return f"{self.sheet}!{first}" if first == last else f"{self.sheet}!{first}:{last}"


def column_number(value: str) -> int:
    total = 0
    for character in value.upper():
        total = total * 26 + ord(character) - 64
    return total


def column_letter(value: int) -> str:
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _split_address(value: str) -> tuple[str | None, str]:
    if "!" not in value:
        return None, value
    sheet, address = value.rsplit("!", 1)
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, address


def _parse_address(value: str, default_sheet: str | None = None) -> CellRange:
    if "[" in value or "]" in value:
        raise CalculationBlocked(f"external or unsupported reference: {value}")
    sheet, address = _split_address(value)
    sheet = sheet or default_sheet
    if not sheet:
        raise CalculationBlocked(f"reference has no worksheet context: {value}")
    parts = address.replace("$", "").split(":")
    if len(parts) not in {1, 2}:
        raise CalculationBlocked(f"unsupported cell reference: {value}")
    cell_pattern = re.compile(r"^([A-Z]{1,3})([1-9][0-9]*)$", re.IGNORECASE)
    first = cell_pattern.fullmatch(parts[0])
    last = cell_pattern.fullmatch(parts[-1])
    if not first or not last:
        raise CalculationBlocked(f"unsupported unbounded or invalid reference: {value}")
    min_col, min_row = column_number(first.group(1)), int(first.group(2))
    max_col, max_row = column_number(last.group(1)), int(last.group(2))
    if min_col > max_col or min_row > max_row:
        raise CalculationBlocked(f"reversed reference is unsupported: {value}")
    return CellRange(sheet, min_row, min_col, max_row, max_col)


def _number(value: Any, *, text_as_zero: bool = False) -> float:
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        if value == "" or text_as_zero:
            return 0.0
        try:
            return float(value)
        except ValueError as exc:
            raise CalculationBlocked(f"unsupported text-to-number conversion: {value!r}") from exc
    raise CalculationBlocked(f"unsupported numeric value type: {type(value).__name__}")


def _truthy(value: Any) -> bool:
    if isinstance(value, ExcelError):
        return False
    if value is None:
        return False
    return value != "" and bool(value)


def _compare(left: Any, right: Any) -> int:
    if isinstance(left, ExcelError) or isinstance(right, ExcelError):
        if isinstance(left, ExcelError) and isinstance(right, ExcelError) and left.code == right.code:
            return 0
        raise CalculationBlocked("Excel error encountered in lookup comparison")
    if left is None:
        left = 0
    if right is None:
        right = 0
    if isinstance(left, str) and isinstance(right, str):
        a, b = left.casefold(), right.casefold()
    elif isinstance(left, (int, float, bool)) and isinstance(right, (int, float, bool)):
        a, b = float(left), float(right)
    elif isinstance(left, str) and isinstance(right, (int, float, bool)):
        return -1
    elif isinstance(right, str) and isinstance(left, (int, float, bool)):
        return 1
    elif type(left) is type(right):
        a, b = str(left).casefold(), str(right).casefold()
    else:
        rank = {int: 0, float: 0, bool: 0, str: 1, _datetime.date: 2, _datetime.datetime: 2}
        left_rank, right_rank = rank.get(type(left), 3), rank.get(type(right), 3)
        if left_rank != right_rank:
            return (left_rank > right_rank) - (left_rank < right_rank)
        a, b = str(left).casefold(), str(right).casefold()
    return (a > b) - (a < b)


def _binary(operator: str, left: Any, right: Any) -> Any:
    if isinstance(left, ExcelError):
        return left
    if isinstance(right, ExcelError):
        return right
    if operator == "&":
        return ("" if left is None else str(left)) + ("" if right is None else str(right))
    if operator in {"=", "<>", "<", ">", "<=", ">="}:
        compare = _compare(left, right)
        return {"=": compare == 0, "<>": compare != 0, "<": compare < 0,
                ">": compare > 0, "<=": compare <= 0, ">=": compare >= 0}[operator]
    try:
        first, second = _number(left), _number(right)
    except CalculationBlocked:
        return ExcelError("#VALUE!")
    if operator == "+":
        return first + second
    if operator == "-":
        return first - second
    if operator == "*":
        return first * second
    if operator == "/":
        return ExcelError("#DIV/0!") if second == 0 else first / second
    if operator == "^":
        try:
            result = first**second
        except (OverflowError, ValueError, ZeroDivisionError):
            return ExcelError("#NUM!")
        return result if math.isfinite(result) else ExcelError("#NUM!")
    raise CalculationBlocked(f"unsupported operator {operator!r}")


def _json_value(value: Any) -> Any:
    if isinstance(value, ExcelError):
        return {"excel_error": value.code}
    if isinstance(value, CellRange):
        return {"range": value.address}
    if isinstance(value, (_datetime.date, _datetime.datetime)):
        return {"excel_datetime": value.isoformat(), "date_only": isinstance(value, _datetime.date) and not isinstance(value, _datetime.datetime)}
    if isinstance(value, list):
        return [[_json_value(item) for item in row] if isinstance(row, list) else _json_value(row) for row in value]
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


class Runtime:
    def __init__(self, raw_values: dict[str, Any], formulas: dict[str, Callable[["Runtime"], Any]],
                 names: dict[str, Any], array_members: dict[str, tuple[str, int, int]],
                 known_cells: set[str], active_addresses: list[str],
                 array_shapes: dict[str, tuple[int, int]] | None = None):
        self.raw_values = raw_values
        self.formulas = formulas
        self.names = names
        self.array_members = array_members
        self.array_shapes = array_shapes or {}
        self.known_cells = known_cells
        self.active_addresses = active_addresses
        self.cache: dict[str, Any] = {}
        self.stack: list[str] = []
        self.name_stack: list[str] = []
        self.array_anchors: dict[str, list[list[Any]]] = {}
        self.evaluation_order: list[str] = []

    @staticmethod
    def _key(value: str) -> str:
        sheet, address = _split_address(value)
        if not sheet:
            raise CalculationBlocked(f"cell address must include worksheet: {value}")
        return f"{sheet.casefold()}!{address.replace('$', '').upper()}"

    def ref(self, sheet: str, min_row: int, min_col: int, max_row: int, max_col: int) -> CellRange:
        return CellRange(sheet, min_row, min_col, max_row, max_col)

    def literal(self, value: Any) -> Any:
        if isinstance(value, dict) and "excel_error" in value:
            return ExcelError(value["excel_error"])
        if isinstance(value, dict) and "excel_datetime" in value:
            parsed = _datetime.datetime.fromisoformat(value["excel_datetime"])
            return parsed.date() if value.get("date_only") else parsed
        return value

    def unsupported(self, token: str) -> Any:
        raise CalculationBlocked(f"unsupported or external reference syntax {token!r}")

    def name(self, name: str, sheet: str) -> Any:
        local_key = f"{sheet.casefold()}|{name.casefold()}"
        key = local_key if local_key in self.names else f"workbook|{name.casefold()}"
        record = self.names.get(key)
        if record is None:
            return ExcelError("#NAME?")
        if key in self.name_stack:
            raise CalculationBlocked(f"defined-name cycle at {sheet}!{name}")
        if "address" in record:
            return _parse_address(record["address"], sheet)
        self.name_stack.append(key)
        try:
            return record["expression"](self, sheet)
        finally:
            self.name_stack.pop()

    def range_between(self, left: Callable[[], Any], right: Callable[[], Any], owner: str) -> Any:
        first, last = left(), right()
        if isinstance(first, ExcelError):
            return first
        if isinstance(last, ExcelError):
            return last
        if not isinstance(first, CellRange) or not first.single or not isinstance(last, CellRange):
            raise CalculationBlocked(f"unsupported range expression in {owner}")
        if first.sheet.casefold() != last.sheet.casefold():
            raise CalculationBlocked(f"cross-sheet range expression is unsupported in {owner}")
        return CellRange(first.sheet, first.min_row, first.min_col, last.max_row, last.max_col)

    def cell(self, sheet: str, row: int, col: int) -> Any:
        address = f"{sheet}!{column_letter(col)}{row}"
        return self.get(address)

    def get(self, address: str) -> Any:
        key = self._key(address)
        if key in self.cache:
            return self.cache[key]
        member = self.array_members.get(key)
        if member is not None:
            anchor, row_offset, col_offset = member
            matrix = self.array_anchors.get(anchor)
            if matrix is None:
                self.get(anchor)
                matrix = self.array_anchors.get(anchor)
            if matrix is None:
                raise CalculationBlocked(f"array anchor {anchor} did not return a matrix")
            try:
                result = matrix[row_offset][col_offset]
            except IndexError as exc:
                raise CalculationBlocked(f"array output does not cover {address}") from exc
            self.cache[key] = result
            self.evaluation_order.append(key)
            return result
        if key not in self.known_cells:
            raise CalculationBlocked(f"cell is outside the source-bound active trace: {address}")
        formula = self.formulas.get(key)
        if formula is None:
            if key not in self.raw_values:
                raise CalculationBlocked(f"raw input is not available for source cell {address}")
            value = self.literal(self.raw_values[key])
            self.cache[key] = value
            return value
        if key in self.stack:
            start = self.stack.index(key)
            raise CalculationBlocked("active formula cycle: " + " -> ".join(self.stack[start:] + [key]))
        self.stack.append(key)
        try:
            result = formula(self)
            shape = self.array_shapes.get(key)
            if shape is not None:
                rows, columns = shape
                matrix = result if isinstance(result, list) else [[result]]
                if matrix and not isinstance(matrix[0], list):
                    matrix = [matrix]
                if len(matrix) != rows or any(len(row) != columns for row in matrix):
                    raise CalculationBlocked(f"array formula at {key} returned a shape other than {rows}x{columns}")
                self.array_anchors[key] = matrix
                result = matrix[0][0]
            elif isinstance(result, CellRange):
                if not result.single:
                    raise CalculationBlocked(f"formula at {key} returned a multi-cell reference as a scalar")
                result = self.cell(result.sheet, result.min_row, result.min_col)
            self.cache[key] = result
            self.evaluation_order.append(key)
            return result
        finally:
            self.stack.pop()

    def _scalar(self, value: Any) -> Any:
        if isinstance(value, CellRange):
            if not value.single:
                raise CalculationBlocked(f"multi-cell reference used as scalar: {value.address}")
            return self.cell(value.sheet, value.min_row, value.min_col)
        return value

    def _matrix(self, value: Any) -> list[list[Any]]:
        if isinstance(value, CellRange):
            return [[self.cell(value.sheet, row, col)
                     for col in range(value.min_col, value.max_col + 1)]
                    for row in range(value.min_row, value.max_row + 1)]
        if isinstance(value, list):
            return value
        return [[self._scalar(value)]]

    def _flatten(self, value: Any) -> list[Any]:
        if isinstance(value, CellRange):
            return [item for row in self._matrix(value) for item in row]
        if isinstance(value, list):
            return [item for row in value for item in (row if isinstance(row, list) else [row])]
        return [value]

    def binary(self, operator: str, left: Callable[[], Any], right: Callable[[], Any]) -> Any:
        first, second = left(), right()
        first_multi = isinstance(first, CellRange) and not first.single or isinstance(first, list)
        second_multi = isinstance(second, CellRange) and not second.single or isinstance(second, list)
        if not first_multi and not second_multi:
            return _binary(operator, self._scalar(first), self._scalar(second))
        first_matrix, second_matrix = self._matrix(first), self._matrix(second)
        shape_a = (len(first_matrix), len(first_matrix[0]) if first_matrix else 0)
        shape_b = (len(second_matrix), len(second_matrix[0]) if second_matrix else 0)
        if first_multi and second_multi and shape_a != shape_b:
            return ExcelError("#VALUE!")
        rows = shape_a[0] if first_multi else shape_b[0]
        cols = shape_a[1] if first_multi else shape_b[1]
        return [[_binary(operator,
                         first_matrix[row][col] if first_multi else first_matrix[0][0],
                         second_matrix[row][col] if second_multi else second_matrix[0][0])
                 for col in range(cols)] for row in range(rows)]

    def unary(self, operator: str, operand: Callable[[], Any]) -> Any:
        value = self._scalar(operand())
        if isinstance(value, ExcelError):
            return value
        if operator == "%":
            return _number(value) / 100
        number = _number(value)
        return number if operator == "+" else -number

    def call(self, name: str, args: list[Callable[[], Any]], sheet: str, owner: str) -> Any:
        def scalar(index: int) -> Any:
            return self._scalar(args[index]())

        if name == "IF":
            if len(args) not in {2, 3}:
                raise CalculationBlocked(f"IF requires 2 or 3 arguments at {owner}")
            condition = scalar(0)
            if isinstance(condition, ExcelError):
                return condition
            selected = 1 if _truthy(condition) else (2 if len(args) == 3 else None)
            return False if selected is None else args[selected]()
        if name == "IFERROR":
            if len(args) != 2:
                raise CalculationBlocked(f"IFERROR requires 2 arguments at {owner}")
            value = args[0]()
            return args[1]() if isinstance(value, ExcelError) else value
        if name in {"IFNA", "_XLFN.IFNA"}:
            if len(args) != 2:
                raise CalculationBlocked(f"IFNA requires 2 arguments at {owner}")
            value = args[0]()
            return args[1]() if isinstance(value, ExcelError) and value.code == "#N/A" else value
        values: list[Any]
        if name in {"AND", "OR"}:
            values = [arg() for arg in args]
            error = next((value for value in values if isinstance(value, ExcelError)), None)
            if error:
                return error
            truth = [_truthy(self._scalar(value)) for value in values]
            return all(truth) if name == "AND" else any(truth)
        if name in {"SUM", "MIN", "MAX"}:
            values = [arg() for arg in args]
            numbers: list[float] = []
            for value in values:
                for item in self._flatten(value):
                    if isinstance(item, ExcelError):
                        return item
                    if isinstance(item, bool):
                        if not isinstance(value, CellRange):
                            numbers.append(float(item))
                    elif isinstance(item, (int, float)):
                        numbers.append(float(item))
            if name == "SUM":
                return sum(numbers)
            if not numbers:
                return 0.0
            return min(numbers) if name == "MIN" else max(numbers)
        if name == "SUMPRODUCT":
            arrays = [self._matrix(arg()) for arg in args]
            if not arrays:
                return 0.0
            shape = (len(arrays[0]), len(arrays[0][0]) if arrays[0] else 0)
            if any((len(item), len(item[0]) if item else 0) != shape for item in arrays):
                return ExcelError("#VALUE!")
            total = 0.0
            for row in range(shape[0]):
                for col in range(shape[1]):
                    product = 1.0
                    for array in arrays:
                        item = array[row][col]
                        if isinstance(item, ExcelError):
                            return item
                        product *= _number(item, text_as_zero=True)
                    total += product
            return total
        if name == "TRANSPOSE":
            if len(args) != 1:
                raise CalculationBlocked(f"TRANSPOSE requires one argument at {owner}")
            matrix = self._matrix(args[0]())
            return [list(row) for row in zip(*matrix)]
        if name == "MATCH":
            if len(args) not in {2, 3}:
                raise CalculationBlocked(f"MATCH requires 2 or 3 arguments at {owner}")
            needle, values_range = scalar(0), args[1]()
            if isinstance(needle, ExcelError):
                return needle
            if not isinstance(values_range, CellRange):
                return ExcelError("#N/A")
            mode_value = scalar(2) if len(args) == 3 else 1
            if isinstance(mode_value, ExcelError):
                return mode_value
            mode = int(_number(mode_value))
            values = self._flatten(values_range)
            if mode == 0:
                match = next((index for index, value in enumerate(values) if _compare(value, needle) == 0), None)
            elif mode == 1:
                if any(_compare(values[i], values[i + 1]) > 0 for i in range(len(values) - 1)):
                    raise CalculationBlocked(f"MATCH approximate range is not ascending at {owner}")
                match = next((i for i in range(len(values) - 1, -1, -1) if _compare(values[i], needle) <= 0), None)
            elif mode == -1:
                if any(_compare(values[i], values[i + 1]) < 0 for i in range(len(values) - 1)):
                    raise CalculationBlocked(f"MATCH reverse approximate range is not descending at {owner}")
                match = next((i for i, value in enumerate(values) if _compare(value, needle) >= 0), None)
            else:
                return ExcelError("#N/A")
            return match + 1 if match is not None else ExcelError("#N/A")
        if name == "VLOOKUP":
            if len(args) not in {3, 4}:
                raise CalculationBlocked(f"VLOOKUP requires 3 or 4 arguments at {owner}")
            needle, table, col_value = scalar(0), args[1](), scalar(2)
            exact = not _truthy(scalar(3)) if len(args) == 4 else False
            if isinstance(needle, ExcelError):
                return needle
            if not isinstance(table, CellRange) or not isinstance(col_value, (int, float)):
                return ExcelError("#VALUE!")
            column = int(col_value)
            width = table.max_col - table.min_col + 1
            if column < 1:
                return ExcelError("#VALUE!")
            if column > width:
                return ExcelError("#REF!")
            keys = [self.cell(table.sheet, row, table.min_col) for row in range(table.min_row, table.max_row + 1)]
            if exact:
                match = next((i for i, value in enumerate(keys) if _compare(value, needle) == 0), None)
            else:
                first_candidate = 1 if (isinstance(needle, (int, float)) and keys and isinstance(keys[0], str)
                                        and len(keys) > 1 and all(isinstance(key, (int, float)) for key in keys[1:])) else 0
                searchable = keys[first_candidate:]
                if any(_compare(searchable[i], searchable[i + 1]) > 0 for i in range(len(searchable) - 1)):
                    raise CalculationBlocked(f"VLOOKUP approximate range is not ascending at {owner}")
                found = next((i for i in range(len(searchable) - 1, -1, -1)
                              if _compare(searchable[i], needle) <= 0), None)
                match = found + first_candidate if found is not None else None
            if match is None:
                return ExcelError("#N/A")
            value = self.cell(table.sheet, table.min_row + match, table.min_col + column - 1)
            return 0 if value is None else value
        if name == "OFFSET":
            if len(args) not in {3, 4, 5}:
                raise CalculationBlocked(f"OFFSET requires 3 to 5 arguments at {owner}")
            base = args[0]()
            if isinstance(base, ExcelError):
                return base
            if not isinstance(base, CellRange):
                return ExcelError("#VALUE!")
            values = [scalar(i) for i in range(1, len(args))]
            if any(isinstance(item, ExcelError) for item in values):
                return next(item for item in values if isinstance(item, ExcelError))
            row_offset, col_offset = int(_number(values[0])), int(_number(values[1]))
            height = int(_number(values[2])) if len(values) >= 3 else base.max_row - base.min_row + 1
            width = int(_number(values[3])) if len(values) >= 4 else base.max_col - base.min_col + 1
            start_row, start_col = base.min_row + row_offset, base.min_col + col_offset
            end_row, end_col = start_row + height - 1, start_col + width - 1
            if height < 1 or width < 1 or start_row < 1 or start_col < 1 or end_row > 1_048_576 or end_col > 16_384:
                return ExcelError("#REF!")
            return CellRange(base.sheet, start_row, start_col, end_row, end_col)
        if name == "INDIRECT":
            if len(args) not in {1, 2}:
                raise CalculationBlocked(f"INDIRECT requires one or two arguments at {owner}")
            text = scalar(0)
            if isinstance(text, ExcelError):
                return text
            a1 = scalar(1) if len(args) == 2 else True
            if isinstance(a1, ExcelError):
                return a1
            if not _truthy(a1):
                raise CalculationBlocked(f"INDIRECT R1C1 is unsupported at {owner}")
            if not isinstance(text, str):
                return ExcelError("#REF!")
            try:
                return _parse_address(text, sheet)
            except CalculationBlocked:
                if "!" in text or "[" in text or "]" in text:
                    raise CalculationBlocked(f"unsupported or external INDIRECT reference {text!r} at {owner}")
                return self.name(text, sheet)
        raise CalculationBlocked(f"unsupported active function {name} at {owner}")

    def run(self, target_addresses: dict[str, str], output_path: Path) -> dict[str, Any]:
        values = {address: _json_value(self.get(address)) for address in self.active_addresses}
        named = {}
        for name in target_addresses:
            resolved = self.name(name, "Main")
            if isinstance(resolved, CellRange):
                if resolved.single:
                    value = self.get(resolved.address)
                else:
                    value = [[self.get(f"{resolved.sheet}!{column_letter(col)}{row}")
                              for col in range(resolved.min_col, resolved.max_col + 1)]
                             for row in range(resolved.min_row, resolved.max_row + 1)]
            else:
                value = resolved
            named[name] = _json_value(value)
        result = {"schema_version": "step4.generated_model_result.v1", "status": "pass",
                  "targets": named, "cells": values,
                  "evaluation_order": self.evaluation_order,
                  "cell_count": len(values), "formula_count": len(self.formulas)}
        output_path.write_text(json.dumps(result, indent=2, sort_keys=True), encoding="utf-8")
        return result
