from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName

from excel_to_act.steps.conversion_workflow import (
    append_stage_artifact,
    create_workflow,
    load_workflow,
    record_decision,
)
from excel_to_act.steps.step4 import generate_model
from excel_to_act.steps.step4.discovery import (
    _approved_stored_empty_addresses,
    _cell_key,
    _record_external_override,
)
from excel_to_act.steps.step5.workflow import _copy_bundle_files, _verify_python


def _approve(workflow_dir: Path, stage: int) -> None:
    for reviewer in ("agent", "human"):
        record_decision(workflow_dir, stage, reviewer, "approve", f"Synthetic fixture review for stage {stage}.")


def _approved_gp_workflow(tmp_path: Path) -> tuple[Path, Path]:
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    source_path = source_dir / "model.xlsx"
    workbook = Workbook()
    main = workbook.active
    main.title = "Main"
    main["A1"] = 3
    main["A2"] = 4
    premium = workbook.create_sheet("Premium")
    formula = "=Main!A1*2"
    premium["J1"] = formula
    workbook.defined_names.add(DefinedName("GP", attr_text="'Premium'!$J$1"))
    workbook.save(source_path)
    workbook.close()

    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    workflow_dir = tmp_path / "workflow"
    create_workflow(workflow_dir, {"workbook_sha256": source_hash, "run_id": "fixture-run",
                                   "source_id": "fixture-source", "workbook_path": str(source_path)})
    _, manifest = load_workflow(workflow_dir)
    manifest["mode"] = "synthetic_fixture"
    from excel_to_act.steps.conversion_workflow import _write_manifest
    _write_manifest(workflow_dir, manifest)

    stage1 = append_stage_artifact(workflow_dir, 1, "checkpoint.json",
                                   {"status": "ready_for_review", "source_sha256": source_hash},
                                   "Synthetic Stage 1 source checkpoint.", input_files={"source": source_path})
    _approve(workflow_dir, 1)
    stage2 = append_stage_artifact(workflow_dir, 2, "checkpoint.json",
                                   {"status": "ready_for_review", "source_sha256": source_hash},
                                   "Synthetic Stage 2 index checkpoint.", input_stages={1: stage1["artifact"]})
    _approve(workflow_dir, 2)
    stage3_payload = {
        "schema_version": "step3.analysis_design.v1",
        "status": "ready_for_review",
        "source": {"workbook_path": str(source_path), "workbook_sha256": source_hash,
                   "source_id": "fixture-source", "run_id": "fixture-run"},
        "design": {"targets": [{"target_id": "T_GP", "selector": "GP", "result_order": 0, "shape": []}],
                    "scenarios": [{"scenario_id": "saved-fixture"}], "validation_plan": [],
                    "coverage": {"known_path_ids": ["R_ONE"], "planned_path_ids": ["R_ONE"]},
                    "options": [{"selected_for_draft": True, "numerical_paths": ["R_ONE"]}],
                    "paths": [{"path_id": "R_ONE", "source_formula": formula}],
                    "historical_evidence": []},
    }
    append_stage_artifact(workflow_dir, 3, "analysis_design.json", stage3_payload,
                          "Synthetic approved-design fixture.", input_stages={2: stage2["artifact"]})
    _approve(workflow_dir, 3)
    return workflow_dir, source_path


def test_stored_empty_adapter_reads_need_candidate_and_exact_scope_evidence(tmp_path: Path) -> None:
    source_path = tmp_path / "stored-empty.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    sheet["C3"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
    workbook.save(source_path)
    workbook.close()
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    scope = {"allowed_regions": [{"source_ref": "Main!C3", "rationale": "fixture"},
                                 {"source_ref": "Main!C3:C5", "rationale": "range is not a blank waiver"}]}
    trace = {"unknowns": [
        {"category": "unrepresented_source_coordinate", "evidence": "main!C3"},
        {"category": "unsupported_formula", "evidence": "Main!C3"},
        {"category": "unrepresented_source_coordinate", "evidence": "Main!C4"},
    ]}

    assert _approved_stored_empty_addresses(source_path, source_hash, scope, trace) == {"main!C3"}

    changed_scope = {"allowed_regions": [{"source_ref": "Main!C4", "rationale": "fixture"}]}
    with pytest.raises(ValueError, match="stored, formula-free blank"):
        _approved_stored_empty_addresses(source_path, source_hash, changed_scope, trace)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    sheet["C3"] = 12
    workbook.save(source_path)
    workbook.close()
    changed_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    with pytest.raises(ValueError, match="stored, formula-free blank"):
        _approved_stored_empty_addresses(source_path, changed_hash, scope, trace)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    sheet["C3"] = "=1"
    workbook.save(source_path)
    workbook.close()
    formula_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    with pytest.raises(ValueError, match="stored, formula-free blank"):
        _approved_stored_empty_addresses(source_path, formula_hash, scope, trace)


def test_external_override_keys_match_casefolded_trace_membership() -> None:
    overrides: dict[str, dict[str, object]] = {}
    _record_external_override(overrides, "Qtable!W5", {"value": 0.25})

    assert set(overrides) == {"qtable!W5"}
    assert _cell_key("qTaBlE!w5") in overrides

    with pytest.raises(ValueError, match="duplicated"):
        _record_external_override(overrides, "qTaBlE!w5", {"value": 0.5})


def test_incomplete_accepted_design_cannot_create_legacy_bundle_or_advance_stage4(tmp_path: Path) -> None:
    workflow_dir, source_path = _approved_gp_workflow(tmp_path)
    before = (workflow_dir / "workflow.json").read_bytes()
    _, before_manifest = load_workflow(workflow_dir)
    before_revisions = list(before_manifest["stages"]["4"]["revisions"])
    stage3 = before_manifest["stages"]["3"]["revisions"][-1]
    stage3_json = workflow_dir / stage3["artifact"]["json"]
    stage3_md = workflow_dir / stage3["artifact"]["md"]
    stage3_hashes = (hashlib.sha256(stage3_json.read_bytes()).hexdigest(),
                     hashlib.sha256(stage3_md.read_bytes()).hexdigest())

    result = generate_model(workflow_dir, tmp_path / "missing-trace.json")

    assert result["status"] == "blocked"
    assert "modular bundles only" in result["reason"]
    assert "Return to Step 3" in result["reason"]
    assert "step4 plan" in result["reason"]
    assert (workflow_dir / "workflow.json").read_bytes() == before
    _, after_manifest = load_workflow(workflow_dir)
    assert after_manifest["stages"]["4"]["revisions"] == before_revisions
    assert (hashlib.sha256(stage3_json.read_bytes()).hexdigest(),
            hashlib.sha256(stage3_md.read_bytes()).hexdigest()) == stage3_hashes
    assert not (workflow_dir / "stage4" / "revision-0001" / "generation_report.json").exists()
    assert not (workflow_dir / "stage4" / "bundles" / "revision-0001" / "bundle").exists()
    assert source_path.is_file()


@pytest.mark.parametrize(
    "payload",
    [
        pytest.param({}, id="missing-design"),
        pytest.param({"design": None}, id="null-design"),
        pytest.param(
            {"design": {"semantic_mapping_required": True},
             "analysis": {"artifacts": {"semantic_plan": None, "semantic_check": None}}},
            id="null-semantic-bindings",
        ),
        pytest.param(
            {"design": {"semantic_mapping_required": True},
             "analysis": {"artifacts": {"semantic_check": {"path": "check.json", "sha256": "check"}}}},
            id="missing-semantic-plan",
        ),
        pytest.param(
            {"design": {"semantic_mapping_required": True},
             "analysis": {"artifacts": {"semantic_plan": {"path": "plan.json", "sha256": "plan"}}}},
            id="missing-semantic-check",
        ),
    ],
)
def test_missing_modular_bindings_return_actionable_step3_guidance(payload: dict[str, object]) -> None:
    from excel_to_act.steps.step4.generator import _require_modular_design

    with pytest.raises(ValueError) as error:
        _require_modular_design(payload)

    assert "modular bundles only" in str(error.value)
    assert "Return to Step 3" in str(error.value)
    assert "step4 plan" in str(error.value)



def test_generation_reader_shows_ordered_non_scalar_values_and_unknowns_without_mutation() -> None:
    import json
    from excel_to_act.steps.step4.generator import _generation_markdown

    design = {
        "targets": [
            {"selector": "Scalar", "result_order": 2, "shape": [], "result_kind": None, "units": None},
            {"selector": "Vector", "result_order": 0, "shape": [3], "result_kind": "vector", "units": "USD",
             "axes": [{"name": "duration", "role": "time", "keys": [0, False, "None"]}]},
            {"selector": "Matrix", "result_order": 1, "shape": [2, 2], "result_kind": "table", "units": "count",
             "axes": []},
        ],
        "open_questions": [
            {"question_id": "Q_CAPTURE", "question": "Capture is required before execution.",
             "status": "resolved by recorded capture", "disposition": "execution prerequisite"},
            {"question_id": "Q_MEANING", "question": "What does the selected axis mean?",
             "status": None, "disposition": None},
        ],
    }
    payload = {
        "source_sha256": "source-fixture", "design_sha256": "design-fixture",
        "active_trace": {"source": "synthetic fixture", "sha256": "trace-fixture"},
        "formula_count": 0, "raw_input_count": 0, "array_member_count": 0,
        "array_instance_count": 0, "compiled_function_count": 0, "source_family_count": 0,
        "semantic_equation_group_count": None,
        "business_input_count": None, "raw_business_input_count": 0,
        "external_business_input_count": False,
        "smoke": {"status": None, "formula_count": 0,
                  "targets": {"Matrix": [[1, 2], [3, 4]], "Scalar": False,
                              "Vector": [0, False, None], "Diagnostic": [7, 8]}},
        "path_coverage": {"implemented_path_ids": [], "numerical_path_ids": [],
                          "implementation_evidence": [], "condition_only_path_ids": []},
        "bundle": {"path": "bundle"},
    }
    before = json.dumps(payload, sort_keys=True)
    report = _generation_markdown(payload, design=design, machine_link="machine.json")
    front = report.split("## Detailed generation evidence", 1)[0]

    assert front.index("**Vector**") < front.index("**Matrix**") < front.index("**Scalar**")
    assert "position [0] = 0" in front and "position [1] = False" in front
    assert "position [2] = Not recorded" in front
    assert "position [0, 0] = 1" in front and "position [1, 1] = 4" in front
    assert "shape []; units Not recorded" in front
    assert "Generated smoke result: `False`" in front
    assert "keys: position [0] = 0<br>position [1] = False<br>position [2] = None" in front
    assert "business input objects: Not recorded (0 raw, False formula-derived external)" in front
    assert "**Q_CAPTURE**" in front and "execution prerequisite" in front
    assert "resolved by recorded capture" in front and "**Q_MEANING**" in front
    assert "Recorded answer: Not recorded" in front and "status: Not recorded" in front
    assert "[generation_report.json](<machine.json>)" in front
    assert "[[1, 2], [3, 4]]" not in front
    assert json.dumps(payload, sort_keys=True) == before

def test_step5_validates_and_copies_every_local_bundle_module(tmp_path: Path) -> None:
    bundle = tmp_path / "bundle"
    bundle.mkdir()
    contents = {
        "model.py": "from pricing import run\nfrom modular_runtime import Runtime\n",
        "pricing.py": "from formula_families import compute\nfrom input_adapter import inputs\n",
        "formula_families.py": "def compute(value):\n    return value\n",
        "input_adapter.py": "def inputs():\n    return {}\n",
        "modular_runtime.py": "class Runtime:\n    pass\n",
        "source_values.json": "{}\n",
        "input_layout.json": '{"external_bindings": []}\n',
        "source_map.json": "{}\n",
        "family_manifest.json": "{}\n",
    }
    file_hashes = {}
    for name, source in contents.items():
        path = bundle / name
        path.write_text(source, encoding="utf-8")
        file_hashes[name] = hashlib.sha256(path.read_bytes()).hexdigest()
    manifest = {"schema_version": "step4.modular_model.v1", "execution_kind": "modular",
                "files": file_hashes, "raw_input_addresses": [], "formula_addresses": ["Premium!J1"],
                "external_input_addresses": [], "external_input_count": 0, "formula_cache_inputs": False}

    assert _verify_python(bundle, manifest) == []
    destination = tmp_path / "isolated"
    copied = _copy_bundle_files(bundle, file_hashes, destination)
    assert set(copied) == set(contents)
    assert (destination / "formula_families.py").is_file()
    assert (destination / "source_map.json").read_text(encoding="utf-8") == "{}\n"

    bad_source = "import excel_to_act\n"
    bad_path = bundle / "formula_families.py"
    bad_path.write_text(bad_source, encoding="utf-8")
    file_hashes["formula_families.py"] = hashlib.sha256(bad_path.read_bytes()).hexdigest()
    assert any("imports unsupported module excel_to_act" in failure
               for failure in _verify_python(bundle, {**manifest, "files": file_hashes}))
