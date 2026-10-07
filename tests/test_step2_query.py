from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
from typing import Any

import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.schemas import SourceLocation, ViewRecord, WorkbookView
from excel_to_act.schemas.step2_prepare import (
    DefinedNameLookup,
    ReadingFileRef,
    ReadingLookup,
    ReadingSource,
    Step2ReadingManifest,
)
from excel_to_act.steps.step2.query import QueryFailure, query, validate_evidence_packet
from excel_to_act.views import _tokens


SOURCE_ID = "source-a"
SOURCE_SHA = "source-sha"
RUN_ID = "run-a"


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode()


def _loc(sheet: str | None, address: str | None, object_type: str = "cell") -> SourceLocation:
    return SourceLocation(
        workbook_path="model.xlsm", sheet_name=sheet, address=address,
        object_type=object_type, object_id=f"{sheet}!{address}" if sheet and address else object_type,
        ooxml_part="xl/vbaProject.bin" if sheet is None else None, source_identity=SOURCE_ID,
    )


def _record(
    record_id: str, sheet: str | None, address: str | None, facts: dict[str, Any],
    kind: str = "cell", object_type: str | None = None,
) -> ViewRecord:
    return ViewRecord(record_id=record_id, record_type=kind,
                      source_location=_loc(sheet, address, object_type or kind), facts=facts)


def _view(view_id: str, scope: str, sheet: str | None, records: list[ViewRecord], source_file: Path) -> WorkbookView:
    return WorkbookView(
        view_id=view_id, source_id=SOURCE_ID, source_sha256=SOURCE_SHA, source_run_id=RUN_ID,
        source_schema_version="phase1.v1", scope=scope, sheet_name=sheet,
        records=records, chunks=[], opaque_report=[r.record_id for r in records if r.facts.get("opaque") is True],
    )


@pytest.fixture
def package(tmp_path: Path) -> dict[str, Any]:
    reading = tmp_path / "reading"
    step1 = tmp_path / "step1"
    reading.mkdir()
    step1.mkdir()

    checkbox = {
        "schema_version": "phase1.v1", "artifact_type": "checkbox_bindings",
        "workbook_sha256": SOURCE_SHA, "status": "complete", "diagnostics": [],
        "bindings": [
            {"sheet": "Main", "sheet_part": "xl/worksheets/sheet1.xml", "shape_id": "1",
             "control_name": "Check Box 1", "linked_cell_raw": "Main!A2", "linked_cell": "Main!A2",
             "linked_sheet": "Main", "linked_address": "A2", "binding_status": "resolved", "sources": []},
            {"sheet": "Calc", "sheet_part": "xl/worksheets/sheet2.xml", "shape_id": "2",
             "control_name": "Check Box 1", "linked_cell_raw": "Calc!A1", "linked_cell": "Calc!A1",
             "linked_sheet": "Calc", "linked_address": "A1", "binding_status": "resolved", "sources": []},
        ],
    }
    active_x = {
        "schema_version": "phase1.v1", "artifact_type": "activex_events",
        "workbook_sha256": SOURCE_SHA, "status": "complete", "diagnostics": [],
        "controls": [
            {"sheet": "Main", "sheet_part": "xl/worksheets/sheet1.xml", "sheet_code_name": "SheetMain",
             "shape_id": "3", "control_name": "CmdPremium", "part": "xl/activeX/activeX1.xml",
             "binary_part": "xl/activeX/activeX1.bin", "event_procedures": ["CmdPremium_Click"],
             "binding_status": "resolved"},
            {"sheet": "Calc", "sheet_part": "xl/worksheets/sheet2.xml", "sheet_code_name": "SheetCalc",
             "shape_id": "4", "control_name": "CmdExcluded", "part": "xl/activeX/activeX2.xml",
             "binary_part": "xl/activeX/activeX2.bin", "event_procedures": ["CmdExcluded_Click"],
             "binding_status": "resolved"},
        ],
    }
    step1_files: dict[str, bytes] = {
        "checkbox_bindings.json": _json_bytes(checkbox),
        "activex_events.json": _json_bytes(active_x),
        "vba_sources/ModA.bas": b"Sub Run()\nEnd Sub\n",
        "vba_sources/Calc.bas": b"Sub CalcRun()\nEnd Sub\n",
        "vba_sources/SheetUnknown.cls": b"Class Unknown\n",
        "vba_sources/SheetCalc.cls": b"Class SheetCalc\n",
    }
    source_refs: dict[str, ReadingFileRef] = {}
    for name, data in step1_files.items():
        path = step1 / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        source_refs[Path(name).name] = ReadingFileRef(root="step1", path=name, sha256=hashlib.sha256(data).hexdigest())
    vba_handoff = {
        "schema_version": "phase1.v1", "artifact_type": "vba_handoff", "workbook_sha256": SOURCE_SHA,
        "available": True, "status": "complete", "diagnostics": [],
        "modules": [
            {"name": "ModA", "kind": "StdModule", "procedures": ["Run"], "source_file": "vba_sources/ModA.bas",
             "sha256": hashlib.sha256(step1_files["vba_sources/ModA.bas"]).hexdigest()},
            {"name": "Calc", "kind": "StdModule", "procedures": ["CalcRun"], "source_file": "vba_sources/Calc.bas",
             "sha256": hashlib.sha256(step1_files["vba_sources/Calc.bas"]).hexdigest()},
            {"name": "SheetUnknown", "kind": "ClassModule", "procedures": [], "source_file": "vba_sources/SheetUnknown.cls",
             "sha256": hashlib.sha256(step1_files["vba_sources/SheetUnknown.cls"]).hexdigest()},
            {"name": "SheetCalc", "kind": "ClassModule", "procedures": ["CmdExcluded_Click"], "source_file": "vba_sources/SheetCalc.cls",
             "sha256": hashlib.sha256(step1_files["vba_sources/SheetCalc.cls"]).hexdigest()},
        ],
    }
    vba_bytes = _json_bytes(vba_handoff)
    (step1 / "vba_handoff.json").write_bytes(vba_bytes)
    source_refs["vba_handoff.json"] = ReadingFileRef(root="step1", path="vba_handoff.json", sha256=hashlib.sha256(vba_bytes).hexdigest())

    checkbox_ref = {"root": "step1", "path": source_refs["checkbox_bindings.json"].path,
                    "sha256": source_refs["checkbox_bindings.json"].sha256}
    activex_ref = {"root": "step1", "path": source_refs["activex_events.json"].path,
                   "sha256": source_refs["activex_events.json"].sha256}
    main_records = [
        _record("m-a1", "Main", "A1", {"address": "A1", "kind": "value", "value": 0, "formula": None,
                                         "cached_value": None, "cached_value_available": False}),
        _record("m-a2", "Main", "A2", {"address": "A2", "kind": "formula", "value": None, "formula": "=1+1",
                                         "cached_value": None, "cached_value_available": False}),
        _record("m-range", "Main", "A1:B2", {"address": "A1:B2", "kind": "table", "metadata": {}}, "range"),
        _record("m-check", "Main", None, {"record": checkbox["bindings"][0], "artifact_ref": checkbox_ref}, "checkbox_binding"),
        _record("m-activex", "Main", None, {"record": active_x["controls"][0], "artifact_ref": activex_ref}, "activex_event"),
        _record("m-feature", "Main", None, {"feature_type": "external_link", "description": "external link", "opaque": True}, "unsupported"),
    ]
    calc_records = [
        _record("c-o7", "Calc", "O7", {"address": "O7", "kind": "value", "value": 1,
                                         "formula": None, "cached_value": None, "cached_value_available": False}),
        _record("c-form-control", "Calc", "O7", {"kind": "form_control", "metadata": {
            "linked_cell": "Calc!O7", "linked_cell_raw": "Calc!O7", "linked_sheet": "Calc", "linked_address": "O7",
        }}, "range", object_type="form_control"),
        _record("c-p8", "Calc", "P8", {"address": "P8", "kind": "value", "value": "x" * 700,
                                         "formula": None, "cached_value": None, "cached_value_available": False}),
        _record("c-q9", "Calc", "Q9", {"address": "Q9", "kind": "value", "value": 2,
                                         "formula": None, "cached_value": None, "cached_value_available": False}),
        _record("c-a1", "Calc", "A1", {"address": "A1", "kind": "value", "value": 3,
                                         "formula": None, "cached_value": None, "cached_value_available": False}),
    ]
    workbook_records = [
        _record("w-mod-a", None, None, {"name": "ModA", "kind": "StdModule", "code": "Sub Run()\nEnd Sub\n", "procedures": ["Run"]}, "vba_module"),
        _record("w-mod-calc", None, None, {"name": "Calc", "kind": "StdModule", "code": "Sub CalcRun()\nEnd Sub\n", "procedures": ["CalcRun"]}, "vba_module"),
        _record("w-sheet-unknown", None, None, {"name": "SheetUnknown", "kind": "ClassModule", "code": "Class Unknown\n", "procedures": []}, "vba_module"),
        _record("w-sheet-calc", None, None, {"name": "SheetCalc", "kind": "ClassModule", "code": "Class SheetCalc\n", "procedures": ["CmdExcluded_Click"]}, "vba_module"),
        _record("w-feature", None, None, {"feature_type": "workbook_feature", "description": "feature", "opaque": True}, "unsupported"),
    ]
    view_values = [
        _view("v-book", "workbook", None, workbook_records, reading),
        _view("v-main", "sheet", "Main", main_records, reading),
        _view("v-calc", "sheet", "Calc", calc_records, reading),
    ]
    view_refs = []
    outputs: dict[str, str] = {}
    for view in view_values:
        relative = f"sources/{SOURCE_ID}/views/{view.view_id}.json"
        data = _json_bytes(view.model_dump(mode="json"))
        (reading / relative).parent.mkdir(parents=True, exist_ok=True)
        (reading / relative).write_bytes(data)
        digest = hashlib.sha256(data).hexdigest()
        view_refs.append(ReadingFileRef(root="reading", path=relative, sha256=digest, bytes=len(data),
                                        view_id=view.view_id, scope=view.scope, sheet_name=view.sheet_name))
        outputs[relative] = digest
    scope_bytes = b"{}\n"
    (reading / "scope.json").write_bytes(scope_bytes)
    outputs["scope.json"] = hashlib.sha256(scope_bytes).hexdigest()
    workbook_ref, main_ref, calc_ref = view_refs
    names = [
        DefinedNameLookup(name="DupName", scope="workbook", node_id="name:global", address="A1",
                          source_location=_loc("Calc", "A1", "defined_name")),
        DefinedNameLookup(name="DupName", scope="Main", node_id="name:local", address="A2",
                          source_location=_loc("Main", "A2", "defined_name")),
        DefinedNameLookup(name="CalcName", scope="workbook", node_id="name:calc", address="O7:Q9",
                          source_location=_loc("Calc", "O7:Q9", "defined_name")),
        DefinedNameLookup(name="OutsideName", scope="workbook", node_id="name:outside", address="A1",
                          source_location=_loc("Calc", "A1", "defined_name")),
    ]
    source = ReadingSource(
        source_id=SOURCE_ID, source_path="model.xlsm", source_sha256=SOURCE_SHA, run_id=RUN_ID,
        status="partial", ready_for_next_step=True, artifacts=source_refs, views=view_refs,
        retained_sheets=["Main"], excluded_sheets=["Calc"],
        allowed_dependency_ranges=[{"sheet_name": "Calc", "address": "O7:Q112"}],
        scope_sha256="scope-sha", lookup=ReadingLookup(workbook_view=workbook_ref,
            sheet_views={"Main": main_ref, "Calc": calc_ref}, defined_names=names),
    )
    manifest = Step2ReadingManifest(
        compiler_version="test", revision_id="revision-a", input_fingerprint="fingerprint-a", options={},
        roots={"reading": ".", "step1": str(step1), "step2_index": str(tmp_path / "index")},
        source_count=1, index=ReadingFileRef(root="step2_index", path="index.json", sha256="0" * 64),
        scope=ReadingFileRef(root="reading", path="scope.json", sha256=outputs["scope.json"], bytes=len(scope_bytes)),
        sources=[source], outputs=outputs,
    )
    manifest_path = reading / "manifest.json"
    manifest_path.write_bytes(_json_bytes(manifest.model_dump(mode="json")))
    state = tmp_path / "state.json"
    state.write_bytes(b'{"attempts":["unchanged"]}\n')
    return {"manifest": manifest_path, "source_id": SOURCE_ID, "reading": reading, "step1": step1,
            "source": source, "views": view_values, "view_refs": view_refs, "state": state}


def _query(package: dict[str, Any], kind: str, **kwargs: Any) -> dict[str, Any]:
    return query(package["manifest"], package["source_id"], kind, **kwargs)


def _claim(view: dict[str, Any], record: dict[str, Any]) -> dict[str, Any]:
    return {"view_id": view["view_id"], "source_run_id": view["source_run_id"],
            "source_schema_version": view["source_schema_version"], "record_id": record["record_id"],
            "source_location": record["source_location"], "kind": "fact", "value": record["facts"]}


def test_all_eight_selectors_and_source_identity(package: dict[str, Any]) -> None:
    assert _query(package, "overview")["summary"]["source"]["source_id"] == SOURCE_ID
    assert _query(package, "sheet", sheet="Main")["summary"]["record_count"] == 6
    zero = _query(package, "cell", sheet="Main", target="A1")["views"][0]["records"][0]
    assert zero["facts"]["value"] == 0 and zero["facts"]["cached_value_available"] is False
    formula = _query(package, "cell", sheet="Main", target="A2")["views"][0]["records"][0]
    assert formula["facts"]["formula"] == "=1+1" and formula["facts"]["cached_value"] is None
    assert _query(package, "range", sheet="Main", address="A1:B2")["pagination"]["delivered_count"] == 3
    assert _query(package, "name", target="DupName", sheet="Main")["summary"]["declaration"]["scope"] == "Main"
    control = _query(package, "control", target="Main.Check Box 1")
    assert control["views"][0]["records"][0]["facts"]["record"] == {
        "sheet": "Main", "sheet_part": "xl/worksheets/sheet1.xml", "shape_id": "1",
        "control_name": "Check Box 1", "linked_cell_raw": "Main!A2", "linked_cell": "Main!A2",
        "linked_sheet": "Main", "linked_address": "A2", "binding_status": "resolved", "sources": [],
    }
    module = _query(package, "vba", target="ModA")
    assert module["views"][0]["records"][0]["facts"]["code"] == "Sub Run()\nEnd Sub\n"
    assert module["source_file_refs"][0]["path"] == "vba_sources/ModA.bas"
    same_named_standard_module = _query(package, "vba", target="Calc")
    assert same_named_standard_module["status"] == "ok"
    assert same_named_standard_module["summary"]["owner_sheet"] is None
    assert same_named_standard_module["views"][0]["records"][0]["facts"]["code"] == "Sub CalcRun()\nEnd Sub\n"
    assert same_named_standard_module["source_file_refs"][0]["path"] == "vba_sources/Calc.bas"
    module_view = same_named_standard_module["views"][0]
    assert validate_evidence_packet(same_named_standard_module, {"claims": [_claim(module_view, module_view["records"][0])]})["valid"] is True
    assert _query(package, "feature", target="external_link", sheet="Main")["views"][0]["records"][0]["facts"]["feature_type"] == "external_link"
    with pytest.raises(QueryFailure, match="explicit source ID"):
        query(package["manifest"], "wrong", "overview")
    with pytest.raises(QueryFailure) as missing:
        _query(package, "cell", sheet="Main", target="A99")
    assert missing.value.status == "not_found" and missing.value.diagnostic["code"] == "cell_missing"


def test_scope_guards_metadata_and_name_ambiguity(package: dict[str, Any]) -> None:
    ambiguous = _query(package, "name", target="DupName")
    assert ambiguous["status"] == "needs_selection" and ambiguous["summary"]["needs_selection"] is True
    assert ambiguous["diagnostics"][0]["severity"] == "error"
    excluded_name = next(item for item in ambiguous["summary"]["candidates"] if item.get("requires_scope"))
    assert "address" not in excluded_name and "source_location" not in excluded_name
    ambiguous_control = _query(package, "control", target="Check Box 1")
    assert ambiguous_control["status"] == "needs_selection"
    excluded_candidate = next(item for item in ambiguous_control["summary"]["candidates"] if item["sheet"] == "Calc")
    assert excluded_candidate["record"].get("linked_address") is None
    excluded_sheet = _query(package, "sheet", sheet="Calc")
    assert excluded_sheet["status"] == "excluded" and excluded_sheet["views"] == []
    excluded_feature = _query(package, "feature", target="anything", sheet="Calc")
    assert excluded_feature["status"] == "excluded" and excluded_feature["views"] == []
    allowed = _query(package, "cell", sheet="Calc", target="O7")
    assert allowed["views"][0]["records"][0]["record_id"] == "c-o7"
    range_packet = _query(package, "range", sheet="Calc", address="O7", budget=10000)
    assert [record["record_id"] for record in range_packet["views"][0]["records"]] == ["c-o7"]
    range_view = range_packet["views"][0]
    assert validate_evidence_packet(range_packet, {"claims": [_claim(range_view, range_view["records"][0])]})["valid"] is True
    name_packet = _query(package, "name", target="CalcName", budget=10000)
    assert "c-form-control" not in [record["record_id"] for record in name_packet["views"][0]["records"]]
    control_record = package["views"][2].records[1].model_dump(mode="json")
    forged_control = copy.deepcopy(range_packet)
    forged_control["views"][0]["records"].append(control_record)
    forged_control["delivered_record_ids"]["v-calc"].append("c-form-control")
    forged_control["pagination"]["delivered"].append({"view_id": "v-calc", "record_id": "c-form-control"})
    forged_control["pagination"]["delivered_count"] += 1
    forged_control["pagination"]["estimated_tokens"] += _tokens(ViewRecord.model_validate(control_record))
    assert validate_evidence_packet(forged_control, {"claims": [_claim(forged_control["views"][0], control_record)]})["valid"] is False
    for kind, kwargs in (("cell", {"sheet": "Calc", "target": "A1"}),
                         ("range", {"sheet": "Calc", "address": "O7:S112"}),
                         ("name", {"target": "OutsideName"})):
        with pytest.raises(QueryFailure) as error:
            _query(package, kind, **kwargs)
        assert error.value.status == "out_of_scope"
        assert error.value.diagnostic["requested_selector"]
        assert error.value.diagnostic["allowed_targets"] == [{"sheet": "Calc", "range": "O7:Q112"}]
    excluded_control = _query(package, "control", target="Calc.Check Box 1")
    assert excluded_control["status"] == "excluded" and excluded_control["views"] == []
    excluded_activex = _query(package, "control", target="Calc.CmdExcluded")
    assert excluded_activex["status"] == "excluded" and excluded_activex["views"] == []
    assert excluded_activex["summary"]["control"]["excluded"] is True
    excluded_vba = _query(package, "vba", target="SheetCalc")
    assert excluded_vba["status"] == "excluded" and excluded_vba["views"] == []
    assert excluded_vba["summary"]["module"]["owner_sheet"] == "Calc"
    unknown_vba = _query(package, "vba", target="SheetUnknown")
    assert unknown_vba["status"] == "needs_scope_resolution" and unknown_vba["views"] == []
    assert unknown_vba["diagnostics"][0]["code"] == "vba_ownership_unknown"
    assert unknown_vba["diagnostics"][0]["severity"] == "error"
    assert unknown_vba["diagnostics"][0]["selector"]["target"] == "sheetunknown"
    cli_result = CliRunner().invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                          "--source-id", SOURCE_ID, "--kind", "vba", "--target", "SheetUnknown"])
    assert cli_result.exit_code == 1
    assert json.loads(cli_result.stdout)["diagnostics"][0]["code"] == "vba_ownership_unknown"


def test_paging_keeps_complete_records_and_resumes_after_oversized_record(package: dict[str, Any]) -> None:
    records = package["views"][2].records[:3]
    budget = _tokens(records[0])
    first = _query(package, "range", sheet="Calc", address="O7:Q112", budget=budget)
    assert first["status"] == "truncated"
    assert [item["record_id"] for item in first["pagination"]["delivered"]] == ["c-o7"]
    assert first["pagination"]["oversized"]["record_id"] == "c-p8"
    assert first["pagination"]["omitted_count"] == 2
    assert first["views"][0]["records"][0]["facts"]["value"] == 1
    second = _query(package, "range", sheet="Calc", address="O7:Q112", budget=2000,
                    cursor=first["pagination"]["next_cursor"])
    assert [item["record_id"] for item in second["pagination"]["delivered"]] == ["c-p8", "c-q9"]
    assert second["status"] == "ok" and second["pagination"]["next_cursor"] is None
    oversized = _query(package, "range", sheet="Calc", address="O7:Q112", budget=1)
    assert oversized["status"] == "oversized" and oversized["pagination"]["delivered_count"] == 0
    assert oversized["diagnostics"][0]["severity"] == "error"
    assert oversized["diagnostics"][0]["canonical_ref"]["view_id"] == "v-calc"
    with pytest.raises(QueryFailure) as bad_cursor:
        _query(package, "range", sheet="Calc", address="O7:Q112", budget=2000, cursor="bad")
    assert bad_cursor.value.status == "integrity_failed"
    with pytest.raises(QueryFailure) as changed_selector:
        _query(package, "range", sheet="Calc", address="O7:Q111", budget=2000,
               cursor=first["pagination"]["next_cursor"])
    assert changed_selector.value.status == "integrity_failed"
    manifest_bytes = package["manifest"].read_bytes()
    try:
        changed_manifest = json.loads(manifest_bytes)
        changed_manifest["revision_id"] = "revision-b"
        package["manifest"].write_bytes(_json_bytes(changed_manifest))
        with pytest.raises(QueryFailure) as changed_revision:
            _query(package, "range", sheet="Calc", address="O7:Q112", budget=2000,
                   cursor=first["pagination"]["next_cursor"])
        assert changed_revision.value.status == "integrity_failed"
    finally:
        package["manifest"].write_bytes(manifest_bytes)


def test_packet_validates_only_exact_delivered_records_and_legacy_cli_still_works(package: dict[str, Any], tmp_path: Path) -> None:
    packet = _query(package, "sheet", sheet="Main", budget=200)
    view = packet["views"][0]
    record = view["records"][0]
    good_output = {"claims": [_claim(view, record)]}
    assert validate_evidence_packet(packet, good_output)["valid"] is True

    missing_record = package["views"][1].records[-1]
    unavailable_claim = {"claims": [_claim(view, {**missing_record.model_dump(mode="json"), "record_id": missing_record.record_id})]}
    unavailable_claim["claims"][0]["view_id"] = view["view_id"]
    assert validate_evidence_packet(packet, unavailable_claim)["valid"] is False

    altered = copy.deepcopy(packet)
    altered["views"][0]["records"][0]["facts"]["value"] = False
    assert validate_evidence_packet(altered, good_output)["valid"] is False
    bad_ref = copy.deepcopy(packet)
    bad_ref["canonical_view_refs"][0]["sha256"] = "0" * 64
    assert validate_evidence_packet(bad_ref, good_output)["valid"] is False
    bad_selector = copy.deepcopy(packet)
    bad_selector["selector"]["target"] = "A2"
    assert validate_evidence_packet(bad_selector, good_output)["valid"] is False
    bad_page = copy.deepcopy(packet)
    bad_page["pagination"]["delivered"] = []
    assert validate_evidence_packet(bad_page, good_output)["valid"] is False
    skipped_prefix = copy.deepcopy(packet)
    skipped_prefix["pagination"]["start_index"] = 1
    assert validate_evidence_packet(skipped_prefix, good_output)["valid"] is False
    bad_manifest = copy.deepcopy(packet)
    bad_manifest["manifest"]["sha256"] = "0" * 64
    assert validate_evidence_packet(bad_manifest, good_output)["valid"] is False
    view_path = package["reading"] / packet["canonical_view_refs"][0]["path"]
    original_view = view_path.read_bytes()
    try:
        view_path.write_bytes(original_view + b" ")
        assert validate_evidence_packet(packet, good_output)["valid"] is False
    finally:
        view_path.write_bytes(original_view)
    bad_run = copy.deepcopy(good_output)
    bad_run["claims"][0]["source_run_id"] = "wrong"
    assert validate_evidence_packet(packet, bad_run)["valid"] is False
    bad_location = copy.deepcopy(good_output)
    bad_location["claims"][0]["source_location"]["address"] = "A99"
    assert validate_evidence_packet(packet, bad_location)["valid"] is False

    packet_path = tmp_path / "packet.json"
    output_path = tmp_path / "claims.json"
    packet_path.write_bytes(_json_bytes(packet))
    output_path.write_bytes(_json_bytes(good_output))
    assert CliRunner().invoke(app, ["views", "validate", "--views", str(packet_path), "--output", str(output_path)]).exit_code == 0
    legacy_path = tmp_path / "views.json"
    legacy_path.write_bytes(_json_bytes([package["views"][1].model_dump(mode="json")]))
    assert CliRunner().invoke(app, ["views", "validate", "--views", str(legacy_path), "--output", str(output_path)]).exit_code == 0


def test_feature_without_target_round_trips_through_packet_validation(package: dict[str, Any], tmp_path: Path) -> None:
    packet_path = tmp_path / "feature-packet.json"
    query_result = CliRunner().invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                            "--source-id", SOURCE_ID, "--kind", "feature", "--sheet", "Main",
                                            "--out", str(packet_path)])
    assert query_result.exit_code == 0
    packet = json.loads(packet_path.read_text())
    assert packet["selector"]["target"] is None
    view = packet["views"][0]
    record = view["records"][0]
    claim = {"claims": [_claim(view, record)]}
    claims_path = tmp_path / "feature-claims.json"
    claims_path.write_bytes(_json_bytes(claim))
    validation = CliRunner().invoke(app, ["views", "validate", "--views", str(packet_path), "--output", str(claims_path)])
    result = json.loads(validation.stdout)
    assert validation.exit_code == 0 and result["valid"] is True and result["provenance_only"] is True


@pytest.mark.parametrize(
    ("artifact_name", "kind", "target"),
    [
        ("checkbox_bindings.json", "control", "Main.Check Box 1"),
        ("vba_handoff.json", "vba", "ModA"),
        ("activex_events.json", "vba", "SheetCalc"),
    ],
)
def test_missing_control_and_vba_artifacts_are_structured_cli_failures(
    package: dict[str, Any], artifact_name: str, kind: str, target: str,
) -> None:
    ref = package["source"].artifacts[artifact_name]
    (package["step1"] / ref.path).unlink()
    result = CliRunner().invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                     "--source-id", SOURCE_ID, "--kind", kind, "--target", target])
    assert result.exit_code == 1
    payload = json.loads(result.stdout)
    assert payload["status"] == "unavailable"
    assert payload["diagnostics"][0]["code"] == "evidence_file_unavailable"
    assert payload["diagnostics"][0]["root"] == ref.root
    assert payload["diagnostics"][0]["path"] == ref.path


def test_malformed_packet_cursor_returns_structured_validation_failure(package: dict[str, Any], tmp_path: Path) -> None:
    packet = _query(package, "sheet", sheet="Main", budget=200)
    view = packet["views"][0]
    record = view["records"][0]
    claim = {"claims": [_claim(view, record)]}
    assert packet["pagination"]["next_cursor"]
    bad_cursor = copy.deepcopy(packet)
    bad_cursor["pagination"]["next_cursor"] = "W10"
    result = validate_evidence_packet(bad_cursor, claim)
    assert result["valid"] is False
    assert result["diagnostics"][0]["code"] == "pagination_cursor_invalid"

    packet_path = tmp_path / "malformed-cursor-packet.json"
    output_path = tmp_path / "malformed-cursor-claims.json"
    packet_path.write_bytes(_json_bytes(bad_cursor))
    output_path.write_bytes(_json_bytes(claim))
    cli_result = CliRunner().invoke(app, ["views", "validate", "--views", str(packet_path), "--output", str(output_path)])
    assert cli_result.exit_code == 1
    assert json.loads(cli_result.stdout)["diagnostics"][0]["code"] == "pagination_cursor_invalid"


def test_missing_reading_root_returns_structured_query_and_validation_failures(
    package: dict[str, Any], tmp_path: Path,
) -> None:
    packet = _query(package, "cell", sheet="Main", target="A1")
    view = packet["views"][0]
    claim = {"claims": [_claim(view, view["records"][0])]}
    packet_path = tmp_path / "packet-missing-root.json"
    claims_path = tmp_path / "claims-missing-root.json"
    packet_path.write_bytes(_json_bytes(packet))
    claims_path.write_bytes(_json_bytes(claim))

    original_manifest = package["manifest"].read_bytes()
    broken_manifest = json.loads(original_manifest)
    broken_manifest["roots"].pop("reading")
    package["manifest"].write_bytes(_json_bytes(broken_manifest))
    try:
        query_result = CliRunner().invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                                "--source-id", SOURCE_ID, "--kind", "overview"])
        assert query_result.exit_code == 1
        query_payload = json.loads(query_result.stdout)
        assert query_payload["diagnostics"][0]["code"] == "manifest_root_missing"
        assert query_payload["diagnostics"][0]["root"] == "reading"

        direct_validation = validate_evidence_packet(packet, claim)
        assert direct_validation["valid"] is False
        assert direct_validation["diagnostics"][0]["code"] == "manifest_root_missing"
        validation_result = CliRunner().invoke(app, ["views", "validate", "--views", str(packet_path), "--output", str(claims_path)])
        assert validation_result.exit_code == 1
        assert json.loads(validation_result.stdout)["diagnostics"][0]["code"] == "manifest_root_missing"
    finally:
        package["manifest"].write_bytes(original_manifest)


def test_query_and_validation_do_not_touch_recovery_state_or_rebuild(package: dict[str, Any], monkeypatch: pytest.MonkeyPatch) -> None:
    before = package["state"].read_bytes()

    def forbidden(*args: Any, **kwargs: Any) -> None:
        raise AssertionError("query dispatched preparation or recovery work")

    from excel_to_act import views as views_module
    from excel_to_act.steps.step2 import prepare as prepare_module
    from excel_to_act.steps.step2 import workflow as workflow_module

    monkeypatch.setattr(views_module.RegexFormulaGraphBuilder, "build", forbidden)
    monkeypatch.setattr(views_module, "compile_views", forbidden)
    monkeypatch.setattr(prepare_module, "prepare", forbidden)
    monkeypatch.setattr(workflow_module, "build_index", forbidden)
    monkeypatch.setattr(workflow_module, "execute_tool", forbidden)
    packet = _query(package, "cell", sheet="Main", target="A1")
    result = validate_evidence_packet(packet, {"claims": [_claim(packet["views"][0], packet["views"][0]["records"][0])]})
    assert result["valid"] is True
    assert packet["metrics"]["raw_workbook_parses"] == 0
    assert packet["metrics"]["inventory_deserializations"] == 0
    assert packet["metrics"]["graph_builds"] == 0
    assert packet["metrics"]["combined_views_loaded"] is False
    assert package["state"].read_bytes() == before


def test_cli_requires_source_id_and_advertises_query(package: dict[str, Any], tmp_path: Path) -> None:
    runner = CliRunner()
    missing = runner.invoke(app, ["step2", "query", "--manifest", str(package["manifest"]), "--kind", "overview"])
    assert missing.exit_code != 0
    wrong = runner.invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                "--source-id", "wrong", "--kind", "overview"])
    assert wrong.exit_code == 1 and json.loads(wrong.stdout)["status"] == "unavailable"
    result_path = tmp_path / "packet.json"
    query_result = runner.invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                       "--source-id", SOURCE_ID, "--kind", "cell", "--sheet", "Main",
                                       "--target", "A1", "--out", str(result_path)])
    assert query_result.exit_code == 0 and json.loads(result_path.read_text())["views"][0]["records"][0]["facts"]["value"] == 0
    oversized_result = runner.invoke(app, ["step2", "query", "--manifest", str(package["manifest"]),
                                           "--source-id", SOURCE_ID, "--kind", "range", "--sheet", "Calc",
                                           "--range", "O7:Q112", "--budget", "1"])
    assert oversized_result.exit_code == 1 and json.loads(oversized_result.stdout)["status"] == "oversized"
    catalogue = json.loads(runner.invoke(app, ["step2", "tools"]).stdout)
    reader = next(item for item in catalogue["reader_commands"] if item["name"] == "step2.query")
    assert reader["inputs"] and reader["outputs"] and reader["next"] == ["views.validate"]
