from __future__ import annotations

import json
from datetime import datetime, time
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.utils.exceptions import InvalidFileException
from typer.testing import CliRunner

from excel_to_act.ingest.openpyxl_reader import OpenpyxlWorkbookReader
from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor
from excel_to_act.interfaces.cli import app
from excel_to_act.schemas import ValidationReport
from excel_to_act.store.local_store import LocalArtifactStore
from excel_to_act.verify import numerical


class _XlError(str):
    """Core-test stand-in for the optional backend's typed error string."""


def fixture(path: Path, *, missing: bool = False,
            date_cache: tuple[str, str] = ("n", "45292"), epoch: datetime | None = None) -> Path:
    workbook = Workbook()
    if epoch is not None:
        workbook.epoch = epoch
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"], inputs["A2"] = 2, 3
    calc = workbook.create_sheet("Calc")
    calc["B1"] = "=Inputs!A1+Inputs!A2*4"
    calc["B2"] = "=DATE(2024,1,1)"
    calc["B2"].number_format = "yyyy-mm-dd"
    calc["B3"] = "=TRUE()"
    calc["B4"] = "=1/0"
    calc["B4"].number_format = "yyyy-mm-dd"
    workbook.save(path)
    if missing:
        return path
    replacements = {"B1": ("n", "14"), "B2": date_cache,
                    "B3": ("b", "1"), "B4": ("e", "#DIV/0!")}
    staged = path.with_suffix(".staged")
    with ZipFile(path) as source, ZipFile(staged, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                root = ET.fromstring(data)
                ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                for cell in root.iter(f"{ns}c"):
                    if cell.get("r") in replacements:
                        kind, value = replacements[cell.get("r")]
                        cell.set("t", kind)
                        cell.find(f"{ns}v").text = value
                data = ET.tostring(root)
            target.writestr(item, data)
    staged.replace(path)
    return path


def inventory(path: Path):
    manifest = OpenpyxlWorkbookReader().read_manifest(path)
    return OpenpyxlInventoryExtractor().extract(path, manifest)


def test_time_only_cache_comparison() -> None:
    options = {"cached_type": "n", "date_format": True, "epoch": datetime(1899, 12, 30)}
    assert numerical._equal("12:00:00", 0.5, tolerance=1e-9, **options)
    assert numerical._equal(time(12), 0.5, tolerance=1e-9, **options)
    assert numerical._equal("12:00:01", 0.5, tolerance=2 / 86400, **options)
    assert not numerical._equal("12:00:01", 0.5, tolerance=0.5 / 86400, **options)
    assert not numerical._equal("13:00:00", 0.5, tolerance=1e-9, **options)
    assert numerical._equal("2024-01-01T00:00:00", 45292, tolerance=1e-9, **options)
    assert numerical._equal("12:00:00", 0.5, tolerance=1e-9,
                            **{**options, "epoch": CALENDAR_MAC_1904})
    assert not numerical._equal("2024-01-01T00:00:00", "2024-01-01T00:00:00",
                                tolerance=1e-9, **options)


def test_date_formatted_text_compares_as_text() -> None:
    options = {"cached_type": "str", "date_format": True,
               "epoch": datetime(1899, 12, 30), "tolerance": 1e-9}
    assert numerical._equal("N/A", "N/A", **options)
    assert not numerical._equal("N/A", "different", **options)
    assert not numerical._equal("N/A", 0.5, **options)
    assert not numerical._equal("2024-01-01", 45292, **options)
    assert numerical._equal("2024-01-01", 45292, **{**options, "cached_type": "n"})


def test_excel_error_and_same_content_text_have_distinct_types() -> None:
    options = {"date_format": False, "epoch": datetime(1899, 12, 30), "tolerance": 1e-9}
    error = _XlError("#DIV/0!")
    assert numerical._equal("#DIV/0!", error, cached_type="e", **options)
    assert numerical._equal("#DIV/0!", "#DIV/0!", cached_type="str", **options)
    assert not numerical._equal("#DIV/0!", error, cached_type="str", **options)
    assert not numerical._equal("#DIV/0!", "#DIV/0!", cached_type="e", **options)


def test_string_date_cache_is_not_a_numeric_date(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "string-date.xlsx", date_cache=("str", "2024-01-01"))
    facts = inventory(path)
    assert facts.sheets[1].cells[1].ooxml_cell_type == "str"
    monkeypatch.setattr(numerical, "_recalculate", lambda _: {
        ("calc", "B1"): 14, ("calc", "B2"): 45292,
        ("calc", "B3"): True, ("calc", "B4"): _XlError("#DIV/0!"),
    })
    report = numerical.validate_workbook(path, facts)
    assert report.status == "fail" and report.coverage.compared_cells == 4
    assert [(d.address, d.baseline, d.actual) for d in report.differences] == [
        ("B2", "2024-01-01", 45292)
    ]


def test_numeric_date_cache_preserves_submillisecond_difference(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "precise-date.xlsx", date_cache=("n", "45292.000000003"))
    facts = inventory(path)
    assert facts.sheets[1].cells[1].cached_value == "2024-01-01T00:00:00"
    assert facts.sheets[1].cells[1].date_serial_text == "45292.000000003"
    monkeypatch.setattr(numerical, "_recalculate", lambda _: {
        ("calc", "B1"): 14, ("calc", "B2"): 45292,
        ("calc", "B3"): True, ("calc", "B4"): _XlError("#DIV/0!"),
    })
    report = numerical.validate_workbook(path, facts)
    assert report.status == "fail" and report.coverage.compared_cells == 4
    assert [(d.address, d.baseline, d.actual, d.tolerance) for d in report.differences] == [
        ("B2", 45292.000000003, 45292, 1e-9)
    ]


def test_numeric_date_cache_rejects_text_result(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "numeric-date.xlsx")
    facts = inventory(path)
    monkeypatch.setattr(numerical, "_recalculate", lambda _: {
        ("calc", "B1"): 14, ("calc", "B2"): "2024-01-01",
        ("calc", "B3"): True, ("calc", "B4"): _XlError("#DIV/0!"),
    })
    report = numerical.validate_workbook(path, facts)
    assert report.status == "fail" and report.coverage.compared_cells == 4
    assert [(d.address, d.baseline, d.actual) for d in report.differences] == [
        ("B2", 45292.0, "2024-01-01")
    ]


def test_1904_date_without_typed_backend_serial_is_incomplete(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "date1904.xlsx", date_cache=("n", "43830"), epoch=CALENDAR_MAC_1904)
    facts = inventory(path)
    assert facts.sheets[1].cells[1].date_serial_text == "43830"
    monkeypatch.setattr(numerical, "_recalculate", lambda _: {
        ("calc", "B1"): 14, ("calc", "B2"): 45292,
        ("calc", "B3"): True, ("calc", "B4"): _XlError("#DIV/0!"),
    })
    report = numerical.validate_workbook(path, facts)
    assert report.status == "incomplete" and not report.differences
    assert report.coverage.unsupported_formula == ["Calc!B2"]
    assert report.coverage.compared_cells == report.coverage.recalculated_cells == 3
    assert any("untyped numeric serial" in message and "1904" in message for message in report.diagnostics)


def test_comparison_types_and_difference(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "cached.xlsx")
    facts = inventory(path)
    assert [cell.cached_value for cell in facts.sheets[1].cells] == [
        14, "2024-01-01T00:00:00", True, "#DIV/0!"
    ]
    assert [cell.ooxml_cell_type for cell in facts.sheets[1].cells] == ["n", "n", "b", "e"]
    actual = {("calc", "B1"): 14, ("calc", "B2"): 45292,
              ("calc", "B3"): True, ("calc", "B4"): _XlError("#DIV/0!")}
    monkeypatch.setattr(numerical, "_recalculate", lambda _: actual)
    report = numerical.validate_workbook(path, facts)
    assert report.status == "pass" and report.coverage.compared_cells == 4
    assert json.loads(Path("schemas/validation_report.schema.json").read_text()) == ValidationReport.model_json_schema()
    assert report.baseline_source == "saved_workbook_cache"
    assert report.actual_source == "formulas" and report.cache_freshness == "unknown"

    actual[("calc", "B1")] = 15
    actual[("calc", "B2")] = True  # a date-formatted cell must not coerce a bool to a serial
    actual[("calc", "B3")] = 1  # bool must not compare as the number one
    actual[("calc", "B4")] = _XlError("#VALUE!")
    report = numerical.validate_workbook(path, facts, absolute_tolerance=0.01)
    assert report.status == "fail"
    assert [(d.sheet, d.address, d.baseline, d.actual, d.tolerance) for d in report.differences] == [
        ("Calc", "B1", 14, 15, 0.01),
        ("Calc", "B2", 45292.0, True, 0.01),
        ("Calc", "B3", True, 1, 0.01),
        ("Calc", "B4", "#DIV/0!", "#VALUE!", 0.01),
    ]

    actual.update({("calc", "B1"): 14, ("calc", "B2"): 45292,
                   ("calc", "B3"): True, ("calc", "B4"): _XlError("#NAME?")})
    facts.sheets[1].cells[-1].formula = "=BOGUS(1)"
    unsupported = numerical.validate_workbook(path, facts)
    assert unsupported.status == "incomplete"
    assert unsupported.coverage.unsupported_formula == ["Calc!B4"]
    assert unsupported.coverage.compared_cells == 3


def test_coverage_states_and_store_handoff(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "missing.xlsx", missing=True)
    facts = inventory(path)
    monkeypatch.setattr(numerical, "_recalculate", lambda _: {("calc", "B1"): 14})
    incomplete = numerical.validate_workbook(path, facts)
    assert incomplete.status == "incomplete"
    assert incomplete.coverage.formula_cells == 4
    assert incomplete.coverage.cached_cells == incomplete.coverage.compared_cells == 0
    assert len(incomplete.coverage.missing_cache) == 4
    assert len(incomplete.coverage.unsupported_formula) == 3

    def unavailable(_):
        raise ImportError("formulas")

    monkeypatch.setattr(numerical, "_recalculate", unavailable)
    assert numerical.validate_workbook(path, facts).status == "not_run"
    out = tmp_path / "artifacts"
    result = CliRunner().invoke(app, ["inspect", str(path), "--out", str(out)])
    assert result.exit_code == 0, result.output
    report = ValidationReport.model_validate_json((out / "validation_report.json").read_bytes())
    handoff = json.loads((out / "handoff.json").read_text())
    assert report.status == handoff["numerical_status"] == "not_run"
    assert handoff["status"] in {"pass", "warn"}
    metadata = json.loads((out / "run_metadata.json").read_text())
    store = LocalArtifactStore(out)
    saved = store.read_run(metadata["workbook_sha256"], metadata["run_id"])
    reference = next(item for item in saved.artifacts if item.name == "validation_report.json")
    assert store.read_artifact(reference) == report
    assert "numerical **not_run**" in (out / "handoff.md").read_text()


def test_failed_inspect_clears_previous_validation_alias(tmp_path: Path, monkeypatch) -> None:
    path = fixture(tmp_path / "cached.xlsx")
    out = tmp_path / "artifacts"
    monkeypatch.setattr(numerical, "_recalculate", lambda _: {
        ("calc", "B1"): 14, ("calc", "B2"): 45292,
        ("calc", "B3"): True, ("calc", "B4"): _XlError("#DIV/0!"),
    })
    assert CliRunner().invoke(app, ["inspect", str(path), "--out", str(out)]).exit_code == 0
    assert json.loads((out / "validation_report.json").read_text())["status"] == "pass"
    previous = json.loads((out / "run_metadata.json").read_text())

    def fail_extract(*_):
        raise InvalidFileException("simulated inventory read failure")

    monkeypatch.setattr(OpenpyxlInventoryExtractor, "extract", fail_extract)
    result = CliRunner().invoke(app, ["inspect", str(path), "--out", str(out)])
    assert result.exit_code == 1
    handoff = json.loads((out / "handoff.json").read_text())
    assert (handoff["status"], handoff["numerical_status"]) == ("fail", "not_run")
    latest = json.loads((out / "run_metadata.json").read_text())
    store = LocalArtifactStore(out)
    failed_run = store.read_run(latest["workbook_sha256"], latest["run_id"])
    assert all(item.name != "validation_report.json" for item in failed_run.artifacts)
    assert not (store.run_dir(latest["workbook_sha256"], latest["run_id"]) / "validation_report.json").exists()
    assert not (out / "validation_report.json").exists()
    assert (store.run_dir(previous["workbook_sha256"], previous["run_id"]) / "validation_report.json").exists()
