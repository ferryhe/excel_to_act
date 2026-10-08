"""Compare saved formula caches with an optional, independent recalculation."""

from __future__ import annotations

import math
import re
from datetime import date, datetime, time
from numbers import Real
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format, is_datetime
from openpyxl.utils.datetime import CALENDAR_MAC_1904, from_excel

from excel_to_act.schemas import CellKind, WorkbookInventory
from excel_to_act.schemas.artifacts import ValidationCoverage, ValidationDifference, ValidationReport

_RESULT_KEY = re.compile(r"^'\[[^]]+\](.*)'!(\$?[A-Z]+\$?\d+)$", re.IGNORECASE)


def _recalculate(workbook_path: Path) -> dict[tuple[str, str], object]:
    from formulas import ExcelModel  # optional extra; never imported by the default installation
    from formulas.tokens.operand import XlError

    results: dict[tuple[str, str], object] = {}
    for key, result in ExcelModel().loads(str(workbook_path)).finish().calculate().items():
        match = _RESULT_KEY.fullmatch(key)
        if match is None:
            continue
        value = result.value
        if getattr(value, "shape", None) != (1, 1):
            continue
        scalar = value[0][0]
        if hasattr(scalar, "item"):
            scalar = scalar.item()
        if not isinstance(scalar, str | int | float | bool):
            continue
        if isinstance(scalar, str) and not isinstance(scalar, XlError):
            scalar = str(scalar)
        results[(match[1].replace("''", "'").casefold(), match[2].replace("$", "").upper())] = scalar
    return results


def _unsupported_name_error(formula: str | None) -> bool:
    if not formula:
        return True
    try:
        from formulas import Parser, get_functions
        from formulas.functions import not_implemented
        from formulas.tokens.function import Function

        functions = get_functions()
        tokens, _ = Parser().ast(formula)
        return any(isinstance(token, Function) and
                   functions.get(token.name.upper(), not_implemented) is not_implemented
                   for token in tokens)
    except Exception:
        return True


def _equal(baseline: object, actual: object, *, cached_type: str | None, date_format: bool, epoch: datetime,
           tolerance: float) -> bool:
    if isinstance(baseline, bool) or isinstance(actual, bool):
        return type(baseline) is type(actual) and baseline == actual
    if date_format and cached_type == "n":
        if not isinstance(actual, Real):
            return False
        expected = None
        try:
            expected = time.fromisoformat(str(baseline))
        except ValueError:
            try:
                expected = datetime.fromisoformat(str(baseline))
            except ValueError:
                pass
        if expected is not None:
            try:
                observed = from_excel(actual, epoch=epoch)
                if isinstance(expected, time):
                    return isinstance(observed, time) and abs((
                        datetime.combine(date.min, expected) - datetime.combine(date.min, observed)
                    ).total_seconds()) <= tolerance * 86400
                if isinstance(observed, date) and not isinstance(observed, datetime):
                    observed = datetime.combine(observed, datetime.min.time())
                return abs((expected - observed).total_seconds()) <= tolerance * 86400
            except (TypeError, ValueError, OverflowError):
                return False
    if isinstance(baseline, str) and isinstance(actual, str):
        return baseline == str(actual) and (cached_type == "e") == (type(actual) is not str)
    if isinstance(baseline, Real) and isinstance(actual, Real):
        return math.isfinite(baseline) and math.isfinite(actual) and math.isclose(
            baseline, actual, rel_tol=0, abs_tol=tolerance
        )
    return type(baseline) is type(actual) and baseline == actual


def validate_workbook(workbook_path: Path, inventory: WorkbookInventory,
                      absolute_tolerance: float = 1e-9) -> ValidationReport:
    """A pass means compared caches match; cache freshness remains unknown."""

    if not math.isfinite(absolute_tolerance) or absolute_tolerance < 0:
        raise ValueError("absolute_tolerance must be finite and nonnegative")
    formulas = [(sheet.name, cell) for sheet in inventory.sheets for cell in sheet.cells
                if cell.kind == CellKind.formula]
    missing = [f"{sheet}!{cell.address}" for sheet, cell in formulas
               if not cell.cached_value_available]
    coverage = ValidationCoverage(formula_cells=len(formulas),
                                  cached_cells=len(formulas) - len(missing),
                                  recalculated_cells=0, compared_cells=0, missing_cache=missing)
    report = ValidationReport(workbook_sha256=inventory.workbook_sha256,
                              absolute_tolerance=absolute_tolerance, coverage=coverage,
                              diagnostics=["Saved cache freshness cannot be established from the workbook."])
    if not formulas:
        report.diagnostics.append("No formula cells to compare.")
        return report
    try:
        actual = _recalculate(workbook_path)
    except ImportError:
        report.diagnostics.append("formulas backend unavailable; numerical validation was not run.")
        return report
    except Exception as exc:
        report.status = "incomplete"
        report.coverage.unsupported_formula = [f"{sheet}!{cell.address}" for sheet, cell in formulas]
        report.diagnostics.append(f"formulas recalculation failed: {type(exc).__name__}: {exc}")
        return report

    date_cells = any(is_date_format(cell.number_format or "") for _, cell in formulas)
    epoch = datetime(1899, 12, 30)
    if date_cells:
        workbook = load_workbook(workbook_path, read_only=True)
        try:
            epoch = workbook.epoch
        finally:
            workbook.close()
    for sheet, cell in formulas:
        key = (sheet.casefold(), cell.address.upper())
        if key not in actual:
            coverage.unsupported_formula.append(f"{sheet}!{cell.address}")
            continue
        if isinstance(actual[key], str) and str(actual[key]) == "#NAME?" and _unsupported_name_error(cell.formula):
            coverage.unsupported_formula.append(f"{sheet}!{cell.address}")
            continue
        if (epoch == CALENDAR_MAC_1904 and cell.cached_value_available and
                cell.ooxml_cell_type == "n" and is_datetime(cell.number_format or "") in {"date", "datetime"} and
                isinstance(actual[key], Real) and not isinstance(actual[key], bool)):
            coverage.unsupported_formula.append(f"{sheet}!{cell.address}")
            report.diagnostics.append(
                f"Cannot compare {sheet}!{cell.address}: formulas returns an untyped numeric serial for a 1904 date; "
                "its date system cannot be verified."
            )
            continue
        coverage.recalculated_cells += 1
        if not cell.cached_value_available:
            continue
        coverage.compared_cells += 1
        observed = actual[key]
        baseline = cell.cached_value
        if cell.ooxml_cell_type == "n" and cell.date_serial_text is not None:
            try:
                baseline = float(cell.date_serial_text)
            except ValueError:
                baseline = cell.date_serial_text
        if not _equal(baseline, observed,
                      cached_type=cell.ooxml_cell_type,
                      date_format=is_date_format(cell.number_format or ""),
                      epoch=epoch, tolerance=absolute_tolerance):
            difference = ValidationDifference(sheet=sheet, address=cell.address,
                baseline=baseline, actual=str(observed) if isinstance(observed, str) else observed,
                tolerance=absolute_tolerance,
                reason="value or type differs")
            report.differences.append(difference)
    if coverage.unsupported_formula:
        report.diagnostics.append("The listed formulas could not be compared with supported scalar semantics.")
    if coverage.missing_cache:
        report.diagnostics.append("The listed formulas have no saved baseline cache.")
    report.status = ("fail" if report.differences else "incomplete" if
                     coverage.missing_cache or coverage.unsupported_formula else "pass")
    return report
