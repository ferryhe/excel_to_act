"""Exercise the selected optional formula oracle on the Issue #7 fixture."""

from __future__ import annotations

import math
import tempfile
from pathlib import Path

from formulas import ExcelModel
from openpyxl import Workbook


def main() -> None:
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "oracle-fixture.xlsx"
        workbook = Workbook()
        inputs = workbook.active
        inputs.title = "Inputs"
        inputs["A1"] = 2
        inputs["A2"] = 3
        calc = workbook.create_sheet("Calc")
        calc["B1"] = "=Inputs!A1+Inputs!A2*4"
        workbook.save(path)

        result = ExcelModel().loads(str(path)).finish().calculate()
        key = next((key for key in result if key.endswith("]CALC'!B1")), None)
        if key is None:
            raise AssertionError(f"oracle result omitted Calc!B1: {list(result)}")
        value = result[key].value[0][0]
        if not math.isclose(float(value), 14.0, rel_tol=0, abs_tol=1e-9):
            raise AssertionError(f"Calc!B1 expected 14.0, got {value!r}")
    print("formulas oracle calculated Inputs!A1 + Inputs!A2 * 4 as 14.0.")


if __name__ == "__main__":
    main()
