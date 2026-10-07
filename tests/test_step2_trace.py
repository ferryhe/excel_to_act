from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path

import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.schemas import FormulaGraph, GraphEdge, GraphNode, GraphNodeKind
from excel_to_act.schemas.step2_prepare import DefinedNameLookup, ReadingFileRef
from excel_to_act.steps.step2.query import QueryFailure, validate_evidence_packet
from test_step2_query import _claim, _json_bytes, _loc, _record, package  # noqa: F401


@pytest.fixture
def traced(package):  # noqa: F811 - imported shared pytest fixture
    source = package["source"]
    names = [
        ("Output", "workbook", "$A$2", "Main"),
        ("Input", "workbook", "$A$1", "O'Brien"),
        ("Local", "workbook", "$A$1", "Main"),
        ("Local", "Main", "$A$2", "Main"),
        ("Constant", "workbook", "42", None),
        ("Relative", "workbook", "A1", "Main"),
        ("Dynamic", "workbook", 'OFFSET(Main!$A$1,0,0)', None),
        ("External", "workbook", "'[other.xlsx]Main'!$A$1", None),
        ("ApprovedRange", "workbook", "$O$7:$Q$9", "Calc"),
        ("Multi", "workbook", "$A$1", "Main"),
        ("Multi", "workbook", "$A$2", "Main"),
    ]
    source.lookup.defined_names = [DefinedNameLookup(
        name=name, scope=scope, address=address,
        node_id=f"name:{scope}!{name}", source_location=_loc(sheet, address if sheet else None, "defined_name"),
    ) for name, scope, address, sheet in names]
    main = package["views"][1]
    main.records += [
        _record("m-b1", "Main", "B1", {"formula": "=Input", "value": None}),
        _record("m-b2", "Main", "B2", {"formula": "=SUM('O''Brien'!A1:A2)", "value": None}),
        _record("m-c1", "Main", "C1", {"formula": "=Calc!O7+Calc!A1", "value": None}),
    ]
    main.records[1].facts["formula"] = "=B1"
    calc = package["views"][2]
    calc.records[0].facts["formula"] = "=A1"
    cells = ["Main!A2", "Main!B1", "Main!B2", "Main!C1", "O'Brien!A1", "Calc!O7", "Calc!A1"]
    nodes = [GraphNode(id=f"cell:{label}", kind=GraphNodeKind.cell, label=label,
                       source_location=_loc(*label.rsplit("!", 1))) for label in cells]
    nodes += [GraphNode(id="range:O'Brien!A1:A2", kind=GraphNodeKind.range,
                        label="O'Brien!A1:A2", source_location=_loc("O'Brien", "A1:A2"))]
    nodes += [GraphNode(id=item.node_id, kind=GraphNodeKind.name, label=item.name,
                        source_location=item.source_location,
                        metadata={"name": item.name, "scope": item.scope, "address": item.address})
              for item in source.lookup.defined_names if item.name != "Multi"]
    pairs = [("cell:Main!A2", "cell:Main!B1"), ("cell:Main!B1", "name:workbook!Input"),
             ("cell:Main!B2", "range:O'Brien!A1:A2"), ("cell:Main!C1", "cell:Calc!O7"),
             ("cell:Main!C1", "cell:Calc!A1"), ("cell:Calc!O7", "cell:Calc!A1")]
    graph = FormulaGraph(nodes=nodes, edges=[GraphEdge(source=a, target=b) for a, b in pairs])
    manifest = json.loads(package["manifest"].read_bytes())
    for view, ref in zip(package["views"], source.views):
        data = _json_bytes(view.model_dump(mode="json"))
        (package["reading"] / ref.path).write_bytes(data)
        ref.sha256, ref.bytes = hashlib.sha256(data).hexdigest(), len(data)
        manifest["outputs"][ref.path] = ref.sha256
    relative = "graph.json"
    data = _json_bytes(graph.model_dump(mode="json"))
    (package["reading"] / relative).write_bytes(data)
    source.dependency_graph = ReadingFileRef(root="reading", path=relative,
                                            sha256=hashlib.sha256(data).hexdigest(), bytes=len(data))
    manifest["outputs"][relative] = source.dependency_graph.sha256
    manifest["sources"] = [source.model_dump(mode="json")]
    package["manifest"].write_bytes(_json_bytes(manifest))
    return package


def _trace(traced, kind="cell", target="Main!A2", direction="upstream", **kwargs):
    from excel_to_act.steps.step2.trace import trace
    return trace(traced["manifest"], traced["source_id"], kind, target,
                 direction=direction, **kwargs)


def _ids(packet):
    return {node["id"] for node in packet["trace"]["nodes"]}


def test_direction_names_ranges_and_cross_sheet(traced):
    upstream = _trace(traced, kind="name", target="Output")
    assert {"name:workbook!Output", "cell:Main!A2", "cell:Main!B1", "cell:O'Brien!A1"} <= _ids(upstream)
    assert "cell:Main!B2" not in _ids(upstream)
    assert any(e["relationship"] == "name_destination" and e["label"] == "structural_derivation"
               for e in upstream["trace"]["edges"])
    down = _trace(traced, target="'O''Brien'!A1", direction="downstream")
    assert {"cell:Main!A2", "cell:Main!B1", "cell:Main!B2", "range:O'Brien!A1:A2"} <= _ids(down)
    both = _trace(traced, target="Main!B1", direction="both")
    assert {"cell:Main!A2", "cell:O'Brien!A1"} <= _ids(both)
    ranged = _trace(traced, kind="range", target="Main!A1:B2")
    assert "cell:O'Brien!A1" in _ids(ranged)
    assert any(e["relationship"] == "stored_range_member" for e in ranged["trace"]["edges"])


def test_name_scope_and_explicit_unresolved_boundaries(traced):
    ambiguous = _trace(traced, kind="name", target="Local")
    assert ambiguous["status"] == "needs_selection"
    local = _trace(traced, kind="name", target="Local", sheet="Main")
    assert "name:Main!Local" in _ids(local)
    constant = _trace(traced, kind="name", target="Constant")
    assert constant["trace"]["terminals"][0]["reason"] == "constant_name"
    for target in ("Relative", "Dynamic", "External", "Multi"):
        packet = _trace(traced, kind="name", target=target)
        assert packet["trace"]["unresolved"]
        assert packet["trace"]["closure"] == "not_proven"


@pytest.mark.parametrize("limits", [{"max_depth": 0}, {"max_nodes": 1}, {"max_edges": 1}])
def test_bounds_frontier_and_cycle(traced, limits):
    limited = _trace(traced, **limits)
    assert limited["trace"]["truncated"] is True
    assert limited["trace"]["frontier"] and limited["trace"]["next_actions"]
    assert limited["trace"]["closure"] == "not_proven"
    graph_ref = traced["source"].dependency_graph
    path = traced["reading"] / graph_ref.path
    graph = json.loads(path.read_bytes())
    graph["edges"].append({"source": "cell:O'Brien!A1", "target": "cell:Main!A2"})
    data = _json_bytes(graph)
    path.write_bytes(data)
    manifest = json.loads(traced["manifest"].read_bytes())
    ref = manifest["sources"][0]["dependency_graph"]
    ref["sha256"], ref["bytes"] = hashlib.sha256(data).hexdigest(), len(data)
    manifest["outputs"][graph_ref.path] = ref["sha256"]
    traced["manifest"].write_bytes(_json_bytes(manifest))
    assert len(_ids(_trace(traced))) == 4


def test_scope_stops_formula_and_preserves_source(traced):
    before = traced["state"].read_bytes()
    packet = _trace(traced, target="Main!C1")
    assert packet["trace"]["scope_boundaries"]
    assert not any(e["source"] == "cell:Calc!O7" for e in packet["trace"]["edges"])
    records = [r for v in packet["views"] for r in v["records"]]
    assert next(r for r in records if r["record_id"] == "c-o7")["facts"]["formula"] == "=A1"
    assert not any(r["record_id"] == "c-a1" for r in records)
    assert packet["metrics"]["graph_builds"] == packet["metrics"]["raw_workbook_parses"] == 0
    assert packet["metrics"]["inventory_deserializations"] == 0
    assert traced["state"].read_bytes() == before
    with pytest.raises(QueryFailure) as failed:
        _trace(traced, target="Calc!A1")
    assert failed.value.status == "out_of_scope"


def test_trace_citations_and_tampering(traced):
    packet = _trace(traced)
    view = packet["views"][0]
    output = {"claims": [_claim(view, view["records"][0])]}
    assert validate_evidence_packet(packet, output)["valid"] is True
    for field in ("edges", "nodes", "limits", "frontier"):
        tampered = copy.deepcopy(packet)
        tampered["trace"][field] = {} if field == "limits" else []
        if tampered == packet:
            tampered["trace"][field] = [{"fake": True}]
        assert validate_evidence_packet(tampered, output)["valid"] is False
    packet["views"][0]["records"][0]["facts"]["value"] = "tampered"
    assert validate_evidence_packet(packet, output)["valid"] is False


def test_approved_excluded_named_range_delivers_only_approved_members(traced):
    packet = _trace(traced, kind="name", target="ApprovedRange", max_nodes=5)
    assert {"name:workbook!ApprovedRange", "range:Calc!O7:Q9", "cell:Calc!O7",
            "cell:Calc!P8", "cell:Calc!Q9"} == _ids(packet)
    assert any(edge["relationship"] == "name_destination" for edge in packet["trace"]["edges"])
    assert sum(edge["relationship"] == "stored_range_member" for edge in packet["trace"]["edges"]) == 3
    records = [record for view in packet["views"] for record in view["records"]]
    assert {record["record_id"] for record in records} == {"c-o7", "c-p8", "c-q9"}
    assert next(record for record in records if record["record_id"] == "c-o7")["facts"]["formula"] == "=A1"
    assert "cell:Calc!A1" not in _ids(packet)
    assert not any(edge["label"] == "source_formula_edge" for edge in packet["trace"]["edges"])
    assert validate_evidence_packet(packet, {"claims": []})["valid"] is True


def test_approved_excluded_cell_finds_retained_consumer(traced):
    packet = _trace(traced, target="Calc!O7", direction="downstream")
    assert "cell:Main!C1" in _ids(packet)
    assert any(edge["source"] == "cell:Main!C1" and edge["target"] == "cell:Calc!O7"
               and edge["label"] == "source_formula_edge" for edge in packet["trace"]["edges"])
    assert "cell:Calc!A1" not in _ids(packet)
    assert not any(edge["source"] == "cell:Calc!O7" and edge["label"] == "source_formula_edge"
                   for edge in packet["trace"]["edges"])
    record_ids = {record["record_id"] for view in packet["views"] for record in view["records"]}
    assert {"m-c1", "c-o7"} <= record_ids
    assert not {"c-a1", "c-form-control"} & record_ids
    assert validate_evidence_packet(packet, {"claims": []})["valid"] is True


@pytest.mark.parametrize("field", ["target", "sheet"])
def test_cli_validation_rejects_malformed_trace_selector(traced, tmp_path: Path, field):
    packet = _trace(traced)
    packet["selector"][field] = 123
    packet_path = tmp_path / "malformed.json"
    packet_path.write_bytes(_json_bytes(packet))
    output = tmp_path / "claims.json"
    output.write_text('{"claims": []}', encoding="utf-8")
    result = CliRunner().invoke(app, ["views", "validate", "--views", str(packet_path),
                                    "--output", str(output)])
    assert result.exit_code == 1
    response = json.loads(result.stdout)
    assert response["valid"] is False and response["diagnostics"]
    assert response["diagnostics"][0]["severity"] == "error"
    assert isinstance(result.exception, SystemExit)
    assert "Traceback" not in result.stdout


def test_cli_discovery_contract_and_errors(traced, tmp_path: Path):
    runner = CliRunner()
    assert "trace" in runner.invoke(app, ["step2", "--help"]).stdout
    out = tmp_path / "packet.json"
    args = ["step2", "trace", "--manifest", str(traced["manifest"]), "--source-id", traced["source_id"],
            "--kind", "cell", "--target", "Main!A2", "--direction", "upstream", "--out", str(out)]
    result = runner.invoke(app, args)
    assert result.exit_code == 0, result.stdout
    packet = json.loads(out.read_bytes())
    assert packet["trace"]["direction"] == "upstream"
    claims = tmp_path / "claims.json"
    claims.write_text('{"claims": []}', encoding="utf-8")
    assert runner.invoke(app, ["views", "validate", "--views", str(out), "--output", str(claims)]).exit_code == 0
    invalid_limit = runner.invoke(app, args + ["--max-depth", "-1"])
    assert invalid_limit.exit_code == 1
    assert json.loads(invalid_limit.stdout)["diagnostics"][0]["code"] == "trace_limit_invalid"
    invalid = runner.invoke(app, args[:-2] + ["--direction", "wrong"])
    assert invalid.exit_code == 1 and json.loads(invalid.stdout)["diagnostics"]
    assert runner.invoke(app, ["step2", "trace", "--manifest", str(traced["manifest"]),
                               "--kind", "cell", "--target", "Main!A2", "--direction", "upstream"]).exit_code != 0


def test_instrumented_extraction_prepare_query_trace_reuse(tmp_path: Path, monkeypatch):
    from excel_to_act.graph.builder import RegexFormulaGraphBuilder
    from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor
    from excel_to_act.steps.step2 import prepare as prepare_module
    from excel_to_act.steps.step2 import query as query_module
    from excel_to_act.steps.step2 import trace as trace_module
    from test_step2_prepare import _prepare_args, _setup

    counters = {"raw_extractions": 0, "graph_builds": 0, "selected_view_loads": 0}
    extract = OpenpyxlInventoryExtractor.extract
    build = RegexFormulaGraphBuilder.build
    load = query_module._load_views

    def extract_count(*args, **kwargs):
        counters["raw_extractions"] += 1
        return extract(*args, **kwargs)

    def build_count(*args, **kwargs):
        counters["graph_builds"] += 1
        return build(*args, **kwargs)

    def load_count(refs, *args, **kwargs):
        counters["selected_view_loads"] += len(refs)
        return load(refs, *args, **kwargs)

    monkeypatch.setattr(OpenpyxlInventoryExtractor, "extract", extract_count)
    monkeypatch.setattr(RegexFormulaGraphBuilder, "build", build_count)
    monkeypatch.setattr(query_module, "_load_views", load_count)
    monkeypatch.setattr(trace_module, "_load_views", load_count)
    root, index, scope, _, _ = _setup(tmp_path, defined_names=True)
    assert counters["raw_extractions"] > 0
    before_prepare = dict(counters)
    out = tmp_path / "reading"
    prepared = CliRunner().invoke(app, _prepare_args(root, index, out, scope))
    assert prepared.exit_code == 0, prepared.stdout
    metrics = json.loads(prepared.stdout)["metrics"]
    assert metrics["inventory_model_loads"] == 1
    assert counters["raw_extractions"] == before_prepare["raw_extractions"]
    assert counters["graph_builds"] == before_prepare["graph_builds"] + 1
    # Check reads against actual invoked APIs, independently of returned metrics.
    state_before = (index.parent / "state.json").read_bytes()
    before_reads = dict(counters)
    manifest = out / "manifest.json"
    source_id = json.loads(manifest.read_bytes())["sources"][0]["source_id"]
    for _ in range(2):
        query_module.query(manifest, source_id, "cell", target="Main!A1")
        trace_module.trace(manifest, source_id, "cell", "Main!A1", direction="upstream")
    assert counters["raw_extractions"] == before_reads["raw_extractions"]
    assert counters["graph_builds"] == before_reads["graph_builds"]
    assert counters["selected_view_loads"] == before_reads["selected_view_loads"] + 4
    assert (index.parent / "state.json").read_bytes() == state_before
    # Inventory parsing is forbidden during repeated reads.
    def inventory_forbidden(*args, **kwargs):
        raise AssertionError("read command deserialized inventory")
    monkeypatch.setattr(prepare_module.WorkbookInventory, "model_validate", inventory_forbidden)
    trace_module.trace(manifest, source_id, "cell", "Main!A1", direction="upstream")


def test_prepared_multi_area_with_graph_and_missing_id_stays_unresolved(tmp_path: Path):
    from excel_to_act.steps.step2.prepare import prepare
    from excel_to_act.steps.step2.trace import trace
    from test_step2_prepare import _setup

    root, index, _, _, _ = _setup(tmp_path, multi_area_name=True)
    out = tmp_path / "reading"
    assert prepare(index, root, out)["status"] == "prepared"
    manifest = out / "manifest.json"
    source = json.loads(manifest.read_bytes())["sources"][0]
    declarations = [item for item in source["lookup"]["defined_names"] if item["name"] == "Both"]
    assert [(item["address"], item["node_id"]) for item in declarations] == [
        ("$A$1", "name:Both"), ("$B$1", None),
    ]
    packet = trace(manifest, source["source_id"], "cell", "Main!A2", direction="upstream")
    assert _ids(packet) == {"cell:Main!A2", "name:Both"}
    assert any(item["node_id"] == "name:Both" and item["reason"] == "ambiguous_or_multi_area_name"
               for item in packet["trace"]["unresolved"])
    assert not any(edge["relationship"] == "name_destination" for edge in packet["trace"]["edges"])
    assert validate_evidence_packet(packet, {"claims": []})["valid"] is True
    selected = trace(manifest, source["source_id"], "name", "Both", direction="upstream")
    assert selected["status"] == "needs_selection"
    assert set(selected["trace"]["roots"]) == {"name:Both", "name:workbook!Both"}
    assert selected["trace"]["unresolved"] and not selected["trace"]["edges"]
    for address in ("A1", "B1"):
        reverse = trace(manifest, source["source_id"], "cell", f"Calc!{address}", direction="downstream")
        assert not any(edge["relationship"] == "name_destination" for edge in reverse["trace"]["edges"])
        assert "cell:Main!A2" not in _ids(reverse)
