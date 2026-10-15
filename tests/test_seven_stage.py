"""Synthetic fixtures for the version and delivery gates; no real acceptance is created."""

from __future__ import annotations

import json
from importlib.resources import files
import os
from pathlib import Path
import shutil
import subprocess

import pytest
from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps import conversion_workflow as workflow
from excel_to_act.steps import seven_stage
from excel_to_act.steps.step6.workflow import _report_markdown as legacy_markdown
from excel_to_act.steps.step6.workflow import create_conversion_report as legacy_report
from excel_to_act.steps.step7.workflow import _report_markdown as step7_markdown
from excel_to_act.steps.step7.workflow import _exact_case_details
from excel_to_act.steps.step7.workflow import _publication_details
from excel_to_act.steps.step7.workflow import create_conversion_report as create_step7_report


def _approve(root: Path, stage: int) -> None:
    for reviewer in ("agent", "human"):
        workflow.record_decision(root, stage, reviewer, "approve", "Synthetic test review.")


def _parent(tmp_path: Path, *, marked_fixture: bool = False) -> Path:
    root = tmp_path / "legacy"
    source = tmp_path / "synthetic-source.xlsx"
    source.write_bytes(b"synthetic source")
    source_record = {"workbook_path": str(source), "workbook_sha256": workflow.hash_file(source),
                     "run_id": "synthetic-run", "source_id": "synthetic-source"}
    if marked_fixture:
        source_record["fixture"] = True
    workflow.create_workflow(root, source_record)
    previous = None
    for stage in range(1, 6):
        revision = workflow.append_stage_artifact(
            root, stage, f"stage{stage}.json", {"status": "ready_for_review"},
            f"Synthetic Stage {stage}.",
            input_stages={stage - 1: previous["artifact"]} if previous else None,
        )
        _approve(root, stage)
        previous = revision
    return root


def _attach(root: Path, parent: Path) -> dict:
    comparison = root.parent / "synthetic-attachment-comparison.json"
    comparison.write_bytes(workflow.json_bytes({"test_fixture": True}))
    return seven_stage.create_attached_workflow(root, parent, comparison,
                                                workflow.hash_file(comparison))


@pytest.mark.parametrize("stage", [0, -1, 8])
def test_attached_stage_status_rejects_out_of_range_stage(tmp_path: Path, stage: int) -> None:
    parent = _parent(tmp_path)
    root = tmp_path / "v2"
    _attach(root, parent)
    with pytest.raises(ValueError, match="unsupported workflow stage"):
        workflow.stage_status(*workflow.load_workflow(root), stage)


def test_stage_status_preserves_unknown_schema_error(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="unsupported schema"):
        workflow.stage_status(tmp_path, {"schema_version": "conversion.workflow.future"}, 0)


def test_attached_and_legacy_valid_stage_status_ranges(tmp_path: Path) -> None:
    parent = _parent(tmp_path)
    root = tmp_path / "v2"
    _attach(root, parent)
    for stage in workflow.STAGES:
        assert workflow.stage_status(*workflow.load_workflow(parent), stage)["stage"] == stage
    for stage in workflow.V2_STAGES:
        state = workflow.stage_status(*workflow.load_workflow(root), stage)
        assert state["stage"] == stage
        if stage <= 5:
            assert state["status"] == "approved"
    (tmp_path / "synthetic-source.xlsx").unlink()
    assert workflow.stage_status(*workflow.load_workflow(root), 1)["status"] == "stale"
    native = {"schema_version": "conversion.workflow.v2", "origin": "native",
              "stages": {str(stage): {"current_revision": 0, "revisions": []}
                         for stage in workflow.V2_STAGES}}
    for stage in workflow.V2_STAGES:
        assert workflow.stage_status(tmp_path, native, stage) == {
            "stage": stage, "status": "pending", "revision": None,
            "reasons": ["artifact not produced"],
        }


def test_v1_is_read_only_and_unknown_versions_or_stage_ranges_reject(tmp_path: Path) -> None:
    root = _parent(tmp_path)
    before = (root / "workflow.json").read_bytes()
    status = CliRunner().invoke(app, ["workflow", "status", "--workflow", str(root)])
    assert status.exit_code == 0
    assert len(json.loads(status.stdout)["stages"]) == 6
    assert json.loads(status.stdout)["stages"][5]["meaning"] == "saved final report"
    for command in (["workflow", "confirm", "--workflow", str(root), "--stage", "5",
                     "--reviewer", "agent", "--decision", "approve", "--message", "Synthetic"],
                    ["step6", "report", "--workflow", str(root)]):
        result = CliRunner().invoke(app, command)
        assert result.exit_code != 0
    assert (root / "workflow.json").read_bytes() == before

    data = json.loads(before)
    for version in (None, "conversion.workflow.future"):
        data["schema_version"] = version
        (root / "workflow.json").write_text(json.dumps(data), encoding="utf-8")
        with pytest.raises(ValueError, match="unsupported schema"):
            workflow.load_workflow(root)
    data["schema_version"] = "conversion.workflow.v1"
    data["stages"]["7"] = {"current_revision": 0, "revisions": []}
    (root / "workflow.json").write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(ValueError, match="invalid stage range"):
        workflow.load_workflow(root)


def test_legacy_python_write_entries_leave_history_untouched(tmp_path: Path) -> None:
    root = tmp_path / "legacy"
    root.mkdir()
    manifest = {"schema_version": "conversion.workflow.v1", "mode": "production",
                "source": {"workbook_sha256": "a" * 64},
                "stages": {str(i): {"current_revision": 0, "revisions": []}
                           for i in range(1, 7)}}
    (root / "workflow.json").write_bytes(workflow.json_bytes(manifest))
    before = {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()}
    calls = (
        lambda: workflow.bind_source_id(root, "source-id"),
        lambda: workflow.append_stage_artifact(root, 1, "stage1.json", {}, "draft"),
        lambda: workflow.record_decision(root, 1, "agent", "approve", "review"),
        lambda: workflow.register_review_delegation(root, tmp_path / "authorization.json"),
        lambda: workflow.submit_typesafe_review(root, 3, tmp_path / "request.json"),
    )
    for call in calls:
        with pytest.raises(ValueError, match="read-only"):
            call()
    assert "read-only" in legacy_report(root)["reason"]
    assert {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()} == before
    assert not (tmp_path / "new-legacy").exists()


def test_attached_v2_keeps_legacy_report_receipts_out_and_pins_parent(tmp_path: Path) -> None:
    parent = _parent(tmp_path)
    workflow.append_stage_artifact(parent, 6, "conversion_report.json",
                                   {"status": "ready_for_review"}, "Old final report.")
    _approve(parent, 6)
    v2 = tmp_path / "v2"
    _attach(v2, parent)
    status = workflow.workflow_status(v2)
    assert status["schema_version"] == "conversion.workflow_status.v2"
    assert len(status["stages"]) == 7
    assert status["stages"][5]["status"] == "pending"
    assert status["next_permitted_stage"] == 6
    assert status["active_delivery"] is None
    assert workflow.current_revision(workflow.load_workflow(v2)[1], 6) is None
    assert workflow.current_revision(workflow.load_workflow(parent)[1], 6)["decisions"]

    parent_manifest = parent / "workflow.json"
    parent_manifest.write_bytes(parent_manifest.read_bytes() + b" ")
    stale = workflow.workflow_status(v2)
    assert all(item["status"] == "stale" for item in stale["stages"][:5])
    assert stale["next_permitted_stage"] is None
    with pytest.raises(ValueError, match="pinned parent workflow changed"):
        seven_stage.require_parent(v2)


def test_attach_pins_separately_supplied_comparison_identity(tmp_path: Path) -> None:
    parent = _parent(tmp_path)
    comparison = tmp_path / "synthetic-attachment-comparison.json"
    comparison.write_bytes(workflow.json_bytes({"test_fixture": True}))
    digest = workflow.hash_file(comparison)
    root = tmp_path / "v2"
    with pytest.raises(ValueError, match="pinned native comparison"):
        seven_stage.create_attached_workflow(root, parent, comparison, "0" * 64)
    assert not root.exists()
    missing = CliRunner().invoke(app, ["workflow", "attach", "--parent", str(parent),
                                      "--out", str(root), "--comparison", str(comparison)])
    assert missing.exit_code != 0 and not root.exists()
    result = CliRunner().invoke(app, ["workflow", "attach", "--parent", str(parent),
                                     "--out", str(root), "--comparison", str(comparison),
                                     "--comparison-sha256", digest])
    assert result.exit_code == 0, result.output
    _, manifest = workflow.load_workflow(root)
    assert manifest["comparison"] == {"path": str(comparison.resolve()), "sha256": digest}
    before = (root / "workflow.json").read_bytes()
    with pytest.raises(ValueError, match="new and separate"):
        seven_stage.create_attached_workflow(root, parent, comparison, digest)
    assert (root / "workflow.json").read_bytes() == before
    comparison.write_bytes(comparison.read_bytes() + b" ")
    with pytest.raises(ValueError, match="pinned native comparison"):
        seven_stage.require_parent(root)
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_synthetic_parent_cannot_attach_or_register(tmp_path: Path) -> None:
    parent = _parent(tmp_path, marked_fixture=True)
    with pytest.raises(ValueError, match="synthetic fixture"):
        _attach(tmp_path / "v2", parent)
    _, manifest = workflow.load_workflow(parent)
    manifest["mode"] = "synthetic_fixture"
    workflow._write_manifest(parent, manifest)
    with pytest.raises(ValueError, match="synthetic workflow"):
        _attach(tmp_path / "v2", parent)


def test_missing_current_step5_human_approval_blocks_attach(tmp_path: Path) -> None:
    parent = _parent(tmp_path)
    _, manifest = workflow.load_workflow(parent)
    manifest["stages"]["5"]["revisions"][-1]["decisions"] = [
        item for item in manifest["stages"]["5"]["revisions"][-1]["decisions"]
        if item["reviewer"] == "agent"]
    workflow._write_manifest(parent, manifest)
    with pytest.raises(ValueError, match="pinned Stage 5 is awaiting_confirmation"):
        _attach(tmp_path / "v2", parent)


def test_legacy_delegation_cannot_be_registered_on_v2(tmp_path: Path) -> None:
    parent = _parent(tmp_path)
    v2 = tmp_path / "v2"
    _attach(v2, parent)
    authorization = tmp_path / "synthetic-old-auth.json"
    authorization.write_text('{"schema_version":"workflow.review_delegation.authorization.v1"}',
                             encoding="utf-8")
    with pytest.raises(ValueError, match="TypeSafe delegation is disabled"):
        workflow.register_review_delegation(v2, authorization)
    result = CliRunner().invoke(app, ["workflow", "delegate", "--workflow", str(v2),
                                      "--authorization", str(authorization)])
    assert result.exit_code == 1 and "TypeSafe delegation is disabled" in result.output


def test_step6_pending_agent_only_rejected_and_forged_candidate_block(tmp_path: Path) -> None:
    parent = _parent(tmp_path)
    v2 = tmp_path / "v2"
    _attach(v2, parent)
    with pytest.raises(ValueError, match="not genuinely accepted"):
        seven_stage.require_delivery(v2)
    revision = workflow.append_stage_artifact(
        v2, 6, "delivery_candidate.json", {"schema_version": "step6.delivery_candidate.v1",
                                           "status": "ready_for_review", "simulation": True},
        "Synthetic forged candidate.", input_stages={5: workflow.current_revision(
            workflow.load_workflow(parent)[1], 5)["artifact"]},
    )
    state = workflow.stage_status(*workflow.load_workflow(v2), 6)
    assert state["status"] == "stale"
    with pytest.raises(ValueError, match="cannot review stale"):
        workflow.record_decision(v2, 6, "agent", "approve", "Synthetic Agent approval.")
    assert workflow.workflow_status(v2)["active_delivery"] is None
    assert workflow.current_revision(workflow.load_workflow(v2)[1], 6) == revision


def test_exact_scenario_comparison_rejects_other_request(tmp_path: Path) -> None:
    request = tmp_path / "request.json"
    request.write_text('{"value": 2}', encoding="utf-8")
    comparison = {"cases": [{"case": "first", "request_sha256": "0" * 64,
                              "status": "pass", "public_result_sha256": "1" * 64}]}
    with pytest.raises(ValueError, match="exact request scenario"):
        seven_stage._comparison_case(comparison, request)
    with pytest.raises(ValueError, match="valid scenario list"):
        seven_stage._comparison_case({"cases": ["malformed"]}, request)


def test_step7_resources_are_copied_and_report_keeps_legacy_sections() -> None:
    package = files("excel_to_act.steps.step7").joinpath("excel-to-act-step7")
    assert "Step 7" in package.joinpath("SKILL.md").read_text(encoding="utf-8")
    assert "## Executive summary" in package.joinpath("assets", "conversion_report.md").read_text(encoding="utf-8")
    for command in ("agent", "skill", "template", "tools"):
        result = CliRunner().invoke(app, ["step7", command])
        assert result.exit_code == 0, result.stdout
    old = legacy_markdown({"source": {}, "design_summary": {}, "review_policy": {}})
    synthetic_delivery = {"version": "v1", "permission": "internal",
                          "path": "/synthetic", "package_sha256": "hash",
                          "receipt_id": "receipt",
                          "readback": {"result_sha256": "hash", "files": {"model.py": "a" * 64}}}
    publication = _publication_details(
        {"interface_sha256": "b" * 64, "comparison_sha256": "c" * 64},
        Path("/synthetic/comparison.json"),
        {"cases": [{} for _ in range(32)],
         "sample": {"count": 32, "ages": list(range(4)), "sex_tokens": [0, 1],
                    "benefit_terms": [0, 1], "payment_periods": [0, 1]}}, synthetic_delivery)
    new = step7_markdown({"source": {}, "design_summary": {}, "review_policy": {},
                          "delivery": synthetic_delivery, "publication": publication})
    for section in ("## Executive summary", "## Purpose/intended use/scope",
                    "## Source model and inputs/assumptions", "## Calculation design and time/module order",
                    "## Conversion approach", "## Validation and reconciliation results",
                    "## Differences/limitations/open decisions", "## Handover/run instructions",
                    "## Acceptance", "### A. Detailed evidence and artifact hashes",
                    "### B. Review history", "### C. Prior option and question history",
                    "### D. Input-boundary source records", "### E. Report organization references"):
        assert section in old and section in new
    assert "## Internal package delivery" in new
    assert publication["interface_sha256"] in new
    assert publication["comparison"]["sha256"] in new
    assert "32-case comparison" in new
    assert "Compared sample scope:" in new
    assert "model_manifest.json" in new
    assert publication["comparison"]["sample"]["count"] == 32
    assert publication["asset_ledger"]["files"] == synthetic_delivery["readback"]["files"]


def test_step7_exact_case_report_has_no_saved_sample_claim() -> None:
    delivery = {"version": "synthetic-v1", "permission": "internal", "path": "/synthetic",
                "package_sha256": "a" * 64, "receipt_id": "synthetic",
                "readback": {"result_sha256": "b" * 64, "files": {"model.py": "c" * 64}}}
    comparison = {"schema_version": "step5.validation_report.v1", "case_count": 2,
                  "cases": [{"case_id": "synthetic-1"}, {"case_id": "synthetic-2"}],
                  "compared": {"exact_cases": 2, "named_targets": 2, "required_targets": 2,
                               "active_formula_members": 10, "required_active_formula_members": 10},
                  "mismatch_count": 0, "tolerances": {"absolute": 1e-12, "relative": 1e-12}}
    publication = _publication_details(
        {"interface_sha256": "d" * 64, "comparison_sha256": "e" * 64},
        Path("/synthetic/comparison.json"), comparison, delivery)
    report = {"source": {}, "design_summary": {
                  "scope": {"boundary": "Two synthetic cases."},
                  "targets": [{"target_id": "GP", "selector": "GP", "result_order": 0, "shape": []}]},
              "review_policy": {}, "delivery": delivery, "publication": publication,
              "case_results": [
                  {"case_id": "synthetic-1", "target_results": {"GP": {"model": 1.0, "excel": 1.0,
                                                                      "matched": True}},
                   "formula_member_count": 5,
                   "report": {"path": "/synthetic/report1.json", "sha256": "1" * 64},
                   "native_oracle": {"path": "/synthetic/oracle1.json", "sha256": "2" * 64},
                   "model_result": {"path": "/synthetic/result1.json", "sha256": "3" * 64}},
                  {"case_id": "synthetic-2", "target_results": {"GP": {"model": 2.0, "excel": 2.0,
                                                                      "matched": True}},
                   "formula_member_count": 5,
                   "report": {"path": "/synthetic/report2.json", "sha256": "4" * 64},
                   "native_oracle": {"path": "/synthetic/oracle2.json", "sha256": "5" * 64},
                   "model_result": {"path": "/synthetic/result2.json", "sha256": "6" * 64}}],
              "validation_summary": comparison["compared"],
              "tolerances": comparison["tolerances"], "mismatch_count": 0,
              "readiness": {"approved_exact_cases": True, "all_configurations": False,
                            "full_workbook_conversion": False},
              "quality": {"status": "partial", "metrics": {"opaque_parts": 539,
                          "control_modules": {"checkbox_invalid": 265, "activex_controls": 8,
                                              "vba_modules": 50}}}}
    markdown = step7_markdown(report)
    assert publication["comparison"]["case_count"] == 2
    assert "sample" not in publication["comparison"]
    assert "2 accepted exact request cases" in markdown
    assert "GP / synthetic-1" in markdown and "GP / synthetic-2" in markdown
    assert "saved case only" not in markdown and "accepted saved scenario" not in markdown
    assert "Comparison tolerances: absolute 1e-12; relative 1e-12" in markdown
    assert "265 invalid checkbox" in markdown and "539 opaque" in markdown


def test_step7_labels_stage3_exclusion_as_design_time_with_approved_stage4_and_5() -> None:
    exclusion = ("No assertion of GP unit/basis, active formula closure, "
                 "current Stage4/5 acceptance or publication approval.")
    report = {"design_summary": {"scope": {"excluded": [exclusion]}},
              "stages": {"4": {"revision": 6, "status": "approved"},
                         "5": {"revision": 1, "status": "approved"}}}

    markdown = step7_markdown(report)

    assert "**Exclusions recorded in the accepted Stage 3 design:**" in markdown
    assert exclusion in markdown
    assert "**Excluded from this result:**" not in markdown
    assert "| Step 4 | 6 | approved |" in markdown
    assert "| Step 5 | 1 | approved |" in markdown
    assert report["design_summary"]["scope"]["excluded"] == [exclusion]


@pytest.mark.parametrize("fault", [None, "wrong_trace", "duplicate_case", "wrong_axes",
                                   "stale_report", "wrong_source", "missing_target",
                                   "extra_target", "malformed_target", "nonnumeric_target",
                                   "wrong_target_count"])
@pytest.mark.parametrize("target_names", [("GP",), ("GP", "PVFB"),
                                         ("PVFB", "GP"), ("PVFB",)])
def test_step7_exact_case_lineage_and_counts_fail_closed(
        tmp_path: Path, fault: str | None, target_names: tuple[str, ...]) -> None:
    source_sha = "a" * 64
    design_cases = []
    cases = []
    for index in (1, 2):
        case_id = f"synthetic-{index}"
        trace_sha = str(index) * 64
        request = tmp_path / f"{case_id}-request.json"
        request.write_text(json.dumps({"inputs": {"IssueAge": {"axes": [],
                                                               "values": 0 if index == 1 else 30}}}),
                           encoding="utf-8")
        request_sha = workflow.hash_file(request)
        files = {}
        for kind in ("native_oracle", "model_result", "validation"):
            path = tmp_path / f"{case_id}-{kind}.json"
            path.write_text("{}", encoding="utf-8")
            files[kind] = {"path": str(path), "sha256": workflow.hash_file(path)}
        case_report = {"status": "ready_for_review", "exact_case": {"case_id": case_id},
                       "mismatch_count": 0, "target_results": {
                           name: {"model": float(index), "excel": float(index), "matched": True}
                           for name in target_names},
                       "oracle": {"sha256": files["native_oracle"]["sha256"],
                                  "source_copy_sha256": source_sha,
                                  "engine": "Microsoft Excel", "macros_executed": False,
                                  "engine_settings": {"version": "16.0"}}}
        path = tmp_path / f"{case_id}-report.json"
        path.write_text(json.dumps(case_report), encoding="utf-8")
        files["report"] = {"path": str(path), "sha256": workflow.hash_file(path)}
        cases.append({"case_id": case_id, "trace_sha256": trace_sha,
                      "request_sha256": request_sha, "formula_member_count": 5, **files})
        design_cases.append({"trace": {"sha256": trace_sha},
                             "public_request": {"path": str(request), "sha256": request_sha}})
    stage3 = {"source": {"workbook_sha256": source_sha},
              "design": {"scope": {"sample_axes": {"IssueAge": [0, 30]}},
                         "targets": [{"target_id": name, "selector": name, "result_order": order,
                                      "shape": []} for order, name in enumerate(target_names)],
                         "tool_interface": {"support_domain": {"status": "exact_cases",
                                                               "cases": design_cases}}}}
    stage5 = {"execution_mode": "exact_cases", "source_sha256": source_sha,
              "case_count": 2, "cases": cases, "mismatch_count": 0,
              "compared": {"exact_cases": 2, "named_targets": 2 * len(target_names),
                           "required_targets": 2 * len(target_names),
                           "active_formula_members": 10, "required_active_formula_members": 10}}
    if fault == "wrong_trace":
        cases[1]["trace_sha256"] = "f" * 64
    elif fault == "duplicate_case":
        cases[1]["case_id"] = cases[0]["case_id"]
    elif fault == "wrong_axes":
        stage3["design"]["scope"]["sample_axes"]["IssueAge"] = [0]
    elif fault == "stale_report":
        Path(cases[0]["report"]["path"]).write_text("{}", encoding="utf-8")
    elif fault == "wrong_source":
        stage5["source_sha256"] = "f" * 64
    elif fault == "wrong_target_count":
        stage5["compared"]["named_targets"] -= 1
    elif fault in {"missing_target", "extra_target", "malformed_target", "nonnumeric_target"}:
        path = Path(cases[0]["report"]["path"])
        report = json.loads(path.read_text(encoding="utf-8"))
        results = report["target_results"]
        if fault == "missing_target":
            del results[target_names[0]]
        elif fault == "extra_target":
            results["UNDECLARED"] = {"model": 1.0, "excel": 1.0, "matched": True}
        elif fault == "malformed_target":
            del results[target_names[0]]["model"]
        else:
            results[target_names[0]]["excel"] = "not-a-number"
        path.write_text(json.dumps(report), encoding="utf-8")
        cases[0]["report"]["sha256"] = workflow.hash_file(path)
    if fault is None:
        scope, results, method = _exact_case_details(stage3, stage5, {})
        assert scope == {"case_count": 2, "axes": {"IssueAge": [0, 30]}}
        assert len(results) == 2 and method["formula_member_count"] == 10
        assert all(list(item["target_results"]) == list(target_names) for item in results)
        markdown = step7_markdown({"design_summary": {"targets": stage3["design"]["targets"]},
                                   "case_results": results})
        for name in target_names:
            assert f"{name} / synthetic-1" in markdown
        if len(target_names) > 1:
            assert markdown.index(f"{target_names[0]} / synthetic-1") < \
                markdown.index(f"{target_names[1]} / synthetic-1")
    else:
        with pytest.raises(ValueError, match="exact|stale|axes|target|evidence"):
            _exact_case_details(stage3, stage5, {})


def test_synthetic_step7_handover_runs_delivered_entry_without_changing_version(tmp_path: Path) -> None:
    source = tmp_path / "source"
    source.mkdir()
    (source / "model.py").write_text(
        "import argparse, json\nfrom pathlib import Path\n"
        "p=argparse.ArgumentParser();p.add_argument('--inputs');p.add_argument('--out');a=p.parse_args()\n"
        "Path(a.out).write_text(json.dumps(json.loads(Path(a.inputs).read_text())))\n",
        encoding="utf-8",
    )
    package = tmp_path / "delivered" / "v1"
    shutil.copytree(source, package)
    ledger = {"model.py": workflow.hash_file(package / "model.py")}
    report = {"source": {}, "stages": {}, "generation_summary": {
        "bundle": {"path": str(source), "files": ledger},
        "run_command": "python model.py --inputs REQUEST.json --out model_result.json"},
        "delivery": {"path": str(package), "version": "v1", "permission": "internal",
                     "package_sha256": "synthetic", "receipt_id": "synthetic",
                     "readback": {"files": ledger, "result_sha256": "synthetic"}},
        "publication": {"interface_sha256": "synthetic",
                        "comparison": {"case_count": 2, "path": str(tmp_path / "comparison.json"),
                                       "sha256": "synthetic", "sample": {}},
                        "asset_ledger": {"path": str(package / "model_manifest.json"), "files": ledger}}}
    rendered = step7_markdown(report)
    assert "Generated bundle:" in rendered
    assert "From the generated bundle directory" not in rendered
    assert "From outside the delivered package directory, run:" in rendered
    command = rendered.split("From outside the delivered package directory, run:", 1)[1]
    command = command.split("```sh\n", 1)[1].split("\n```", 1)[0]
    assert f'"{package / "model.py"}"' in command
    assert str(source / "model.py") not in command
    assert seven_stage._package_files(package, ledger) == ledger

    for value in (2, 3):
        outside = tmp_path / f"call{value}"
        outside.mkdir()
        (outside / "REQUEST.json").write_text(json.dumps({"value": value}), encoding="utf-8")
        environment = os.environ.copy()
        environment.pop("PYTHONPATH", None)
        result = subprocess.run(command, cwd=outside, env=environment,
                                shell=True, capture_output=True, text=True, timeout=30, check=False)
        assert result.returncode == 0, result.stderr
        assert json.loads((outside / "model_result.json").read_text()) == {"value": value}
        assert seven_stage._package_files(package, ledger) == ledger
    assert not (package / "model_result.json").exists()
    assert report["delivery"]["version"] == "v1"


def test_isolated_entry_keeps_sibling_modules_without_project_path(tmp_path: Path) -> None:
    package = tmp_path / "synthetic-package"
    package.mkdir()
    (package / "helper.py").write_text("VALUE = 3\n", encoding="utf-8")
    (package / "model.py").write_text(
        "import json\nfrom pathlib import Path\nfrom helper import VALUE\n"
        "import argparse\np=argparse.ArgumentParser()\np.add_argument('--inputs')\n"
        "p.add_argument('--out')\na=p.parse_args()\n"
        "Path(a.out).write_text(json.dumps({'value': json.loads(Path(a.inputs).read_text())['value'] + VALUE}))\n",
        encoding="utf-8",
    )
    request = tmp_path / "synthetic-request.json"
    request.write_text('{"value": 2}', encoding="utf-8")
    record = seven_stage._run_entry(package, request, tmp_path / "outside" / "result.json")
    assert record["exit_code"] == 0
    assert record["command"][1:3] == ["-E", "-s"]
    assert json.loads((tmp_path / "outside" / "result.json").read_text()) == {"value": 5}


def _synthetic_candidate(tmp_path: Path, *, create_revision: bool = True) -> tuple[Path, Path, Path]:
    """A local fake package and fake upstream reports, never business evidence."""
    tmp_path.mkdir(parents=True, exist_ok=True)
    parent = _parent(tmp_path, marked_fixture=True)
    template = parent / "stage4" / "synthetic-bundle"
    template.mkdir(parents=True)
    (template / "model.py").write_text(
        "import argparse, json\nfrom pathlib import Path\n"
        "p=argparse.ArgumentParser();p.add_argument('--inputs');p.add_argument('--out');a=p.parse_args()\n"
        "v=json.loads(Path(a.inputs).read_text())['value']\n"
        "Path(a.out).write_text(json.dumps({'value':v*2}))\n", encoding="utf-8")
    source_sha = workflow.load_workflow(parent)[1]["source"]["workbook_sha256"]
    (template / "model_manifest.json").write_bytes(workflow.json_bytes({
        "schema_version": "step4.modular_model.v1", "source_sha256": source_sha,
        "files": {"model.py": workflow.hash_file(template / "model.py")},
        "variable_input_contract_sha256": "synthetic-interface"}))
    ledger = {name: workflow.hash_file(template / name) for name in ("model.py", "model_manifest.json")}
    stage3_revision = workflow.current_revision(workflow.load_workflow(parent)[1], 3)
    stage4_revision = workflow.append_stage_artifact(parent, 4, "generation_report.json", {
        "schema_version": "step4.generation_report.v1", "status": "ready_for_review",
        "bundle": {"path": str(template), "files": ledger},
        "runtime_dependencies": "Python standard library only",
        "variable_input_contract_sha256": "synthetic-interface"},
        "Synthetic generation report.", input_stages={3: stage3_revision["artifact"]})
    _approve(parent, 4)
    stage4 = parent / stage4_revision["artifact"]["json"]
    stage5_revision = workflow.append_stage_artifact(parent, 5, "validation_report.json", {
        "schema_version": "step5.validation_report.v1", "status": "ready_for_review",
        "source_sha256": source_sha, "stage4_artifact_sha256": workflow.hash_file(stage4),
        "mismatch_count": 0, "oracle": {"engine": "Microsoft Excel"},
        "formula_cache_inputs": False, "standalone_project_imports": False},
        "Synthetic validation report.", input_stages={4: stage4_revision["artifact"]})
    _approve(parent, 5)
    root = tmp_path / "v2"
    old_root, old = workflow.load_workflow(parent)
    workflow._write_manifest(root, {
        "schema_version": "conversion.workflow.v2", "mode": "synthetic_fixture",
        "source": old["source"], "parent": {"root": str(old_root),
                                        "manifest_sha256": workflow.hash_file(old_root / "workflow.json")},
        "stages": {str(stage): {"current_revision": 0, "revisions": []}
                   for stage in range(1, 8)}, "delivery": None, "delivery_history": [],
    })
    package = root / "stage6" / "packages" / "synthetic-v1" if create_revision else template
    if create_revision:
        package.mkdir(parents=True)
        for name in ledger:
            (package / name).write_bytes((template / name).read_bytes())
    request = tmp_path / "synthetic-request.json"
    request.write_text('{"value": 2}', encoding="utf-8")
    other = tmp_path / "synthetic-other-request.json"
    other.write_text('{"value": 3}', encoding="utf-8")
    result = (root / "stage6" / "smoke" / "synthetic-v1" / "result.json"
              if create_revision else tmp_path / "synthetic-expected-result.json")
    smoke = seven_stage._run_entry(package, request, result)
    comparison = tmp_path / "synthetic-comparison.json"
    comparison.write_bytes(workflow.json_bytes({
        "status": "pass", "mismatch_count": 0, "source_sha256": source_sha,
        "stage4_report_sha256": workflow.hash_file(stage4),
        "stage4_bundle_manifest_sha256": workflow.hash_file(package / "model_manifest.json"),
        "formula_cache_inputs": False, "public_pass_count": 1,
        "cases": [{"case": "synthetic-first",
        "request_sha256": workflow.hash_file(request), "status": "pass",
        "public_result_sha256": workflow.hash_file(result)}]}))
    _, attached = workflow.load_workflow(root)
    attached["comparison"] = {"path": str(comparison.resolve()),
                              "sha256": workflow.hash_file(comparison)}
    workflow._write_manifest(root, attached)
    if not create_revision:
        return root, request, other
    stage5 = parent / stage5_revision["artifact"]["json"]
    candidate = {"schema_version": "step6.delivery_candidate.v1", "tool": "step6.candidate",
        "status": "ready_for_review", "simulation": True,
        "version": "synthetic-v1", "permission": "internal", "entry": "model.py",
        "parent": workflow.load_workflow(root)[1]["parent"], "package_path": str(package),
        "package_sha256": workflow.hash_bytes(workflow.json_bytes(ledger)), "files": ledger,
        "runtime_dependencies": "Python standard library only",
        "interface_sha256": "synthetic-interface", "comparison_case": "synthetic-first",
        "request_sha256": workflow.hash_file(request),
        "comparison_sha256": workflow.hash_file(comparison), "smoke": smoke,
        "stage4_artifact_sha256": workflow.hash_file(stage4),
        "stage5_artifact_sha256": workflow.hash_file(stage5)}
    workflow.append_stage_artifact(root, 6, "delivery_candidate.json", candidate,
        "Synthetic test candidate; no real business approval.",
        input_files={"stage4_report": stage4, "stage5_report": stage5,
                     "native_comparison": comparison, "smoke_request": request,
                     "smoke_result": result,
                     **{f"package:{name}": package / name for name in ledger}},
        input_stages={5: stage5_revision["artifact"]})
    return root, request, other


@pytest.mark.parametrize("review_state", ("pending", "accepted", "rejected", "revoked"))
def test_synthetic_candidate_retry_respects_rejection_and_revocation(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch, review_state: str) -> None:
    root, request, _other = _synthetic_candidate(tmp_path)
    parent, upstream = workflow.load_workflow(tmp_path / "legacy")
    comparison = Path(workflow.load_workflow(root)[1]["comparison"]["path"])
    # This test exercises retry bookkeeping, not the synthetic upstream evidence.
    monkeypatch.setattr(seven_stage, "require_parent", lambda *_args: (parent, upstream))
    monkeypatch.setattr(seven_stage, "require_upstream", lambda *_args: (parent, upstream))
    monkeypatch.setattr(seven_stage, "_require_stage5", lambda *_args: None)
    monkeypatch.setattr(seven_stage, "_require_comparison", lambda *_args: None)
    monkeypatch.setattr(seven_stage, "candidate_staleness", lambda *_args: [])

    if review_state == "rejected":
        workflow.record_decision(root, 6, "agent", "reject", "Synthetic rejection.", return_to=6)
    elif review_state in {"accepted", "revoked"}:
        _approve(root, 6)
    if review_state == "revoked":
        _, manifest = workflow.load_workflow(root)
        human = workflow.stage_status(root, manifest, 6)["human"]
        manifest["delivery_history"].append({
            "schema_version": "step6.delivery_revocation.v1",
            "version": "synthetic-v1", "receipt_id": human["receipt_id"],
            "reason": "Synthetic revocation."})
        workflow._write_manifest(root, manifest)

    expected_status = {"pending": "awaiting_confirmation", "accepted": "simulated_approved",
                       "rejected": "rejected", "revoked": "simulated_approved"}
    assert workflow.stage_status(*workflow.load_workflow(root), 6)["status"] == expected_status[review_state]
    before = (root / "workflow.json").read_bytes()
    package = root / "stage6" / "packages" / "synthetic-v1"
    package_before = {path.name: path.read_bytes() for path in package.iterdir()}
    if review_state in {"rejected", "revoked"}:
        try:
            result = seven_stage.create_candidate(root, "synthetic-v1", request, comparison)
        except ValueError as exc:
            assert "new version" in str(exc)
            assert review_state in str(exc)
        else:
            pytest.fail(f"{review_state} candidate returned false success: {result}")
    else:
        result = seven_stage.create_candidate(root, "synthetic-v1", request, comparison)
        assert result["status"] == "pass" and result["idempotent"] is True
        assert result["revision"] == 1
    assert (root / "workflow.json").read_bytes() == before
    assert {path.name: path.read_bytes() for path in package.iterdir()} == package_before
    assert workflow.load_workflow(root)[1]["delivery"] is None


def test_synthetic_source_extras_are_not_copied_into_candidate(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root, request, _other = _synthetic_candidate(tmp_path, create_revision=False)
    parent, upstream = workflow.load_workflow(tmp_path / "legacy")
    stage4_revision = workflow.current_revision(upstream, 4)
    stage4 = workflow.read_json(parent / stage4_revision["artifact"]["json"])
    bundle = Path(stage4["bundle"]["path"])
    (bundle / "__pycache__").mkdir()
    (bundle / "__pycache__" / "model.pyc").write_bytes(b"synthetic cache")
    (bundle / "unlisted").mkdir()
    (bundle / "unlisted" / "note.txt").write_text("synthetic extra", encoding="utf-8")
    source_before = {path.relative_to(parent): path.read_bytes()
                     for path in parent.rglob("*") if path.is_file()}
    root_before = (root / "workflow.json").read_bytes()
    comparison = Path(workflow.load_workflow(root)[1]["comparison"]["path"])

    monkeypatch.setattr(seven_stage, "require_upstream", lambda *_args: (parent, upstream))
    monkeypatch.setattr(seven_stage, "require_stage_approved", lambda _root, stage:
                        (upstream, workflow.current_revision(upstream, stage)))
    monkeypatch.setattr(seven_stage, "_stage4", lambda _root:
                        (upstream, stage4_revision, stage4, bundle))
    monkeypatch.setattr(seven_stage, "_require_stage5", lambda *_args: None)
    monkeypatch.setattr(seven_stage, "_require_comparison", lambda *_args: None)

    class SyntheticStop(Exception):
        pass

    def stop_before_recording(*_args, **_kwargs):
        raise SyntheticStop

    monkeypatch.setattr(seven_stage, "append_stage_artifact", stop_before_recording)
    with pytest.raises(SyntheticStop):
        seven_stage.create_candidate(root, "synthetic-source-extra", request, comparison)

    staged = root / "stage6" / "packages" / "synthetic-source-extra"
    ledger = stage4["bundle"]["files"]
    assert seven_stage._package_files(staged, ledger) == ledger
    assert sorted(path.relative_to(staged).as_posix() for path in staged.rglob("*") if path.is_file()) == sorted(ledger)
    assert {path.relative_to(parent): path.read_bytes()
            for path in parent.rglob("*") if path.is_file()} == source_before
    assert (root / "workflow.json").read_bytes() == root_before
    assert workflow.current_revision(workflow.load_workflow(root)[1], 6) is None


def test_synthetic_candidate_cannot_be_reviewed_delivered_or_reported(tmp_path: Path) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    assert workflow.load_workflow(root)[1]["mode"] == "synthetic_fixture"
    assert workflow.stage_status(*workflow.load_workflow(root), 6)["status"] == "stale"
    with pytest.raises(ValueError, match="cannot review stale"):
        workflow.record_decision(root, 6, "agent", "approve", "Synthetic Agent review.")
    with pytest.raises(ValueError, match="synthetic workflow"):
        seven_stage.deliver(root)
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"


def test_synthetic_candidate_stays_stale_when_comparison_pin_changes(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    revision = workflow.current_revision(manifest, 6)
    candidate_path = root / revision["artifact"]["json"]
    candidate = workflow.read_json(candidate_path)
    candidate.pop("simulation")
    candidate_path.write_bytes(workflow.json_bytes(candidate))
    manifest["comparison"]["sha256"] = "0" * 64
    monkeypatch.setattr(seven_stage, "_require_stage5", lambda *_args: None)
    reasons = seven_stage.candidate_staleness(root, manifest, revision)
    assert any("pinned native comparison" in reason for reason in reasons)
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"


def test_native_exact_case_stage5_can_supply_candidate_comparison(tmp_path: Path,
                                                                  monkeypatch: pytest.MonkeyPatch) -> None:
    request = tmp_path / "request.json"
    request.write_text('{"synthetic": true}', encoding="utf-8")
    result = tmp_path / "result.json"
    result.write_text('{"synthetic": true}', encoding="utf-8")
    stage5 = {"schema_version": "step5.validation_report.v1", "execution_mode": "exact_cases",
              "cases": [{"case_id": "synthetic-case", "request_sha256": workflow.hash_file(request),
                         "model_result": {"path": str(result), "sha256": workflow.hash_file(result)}}]}
    report = tmp_path / "stage5.json"
    report.write_bytes(workflow.json_bytes(stage5))
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"synthetic source")
    manifest = {"schema_version": "conversion.workflow.v2", "origin": "native", "mode": "production",
                "source": {"workbook_path": str(source),
                                               "workbook_sha256": workflow.hash_file(source)}}
    monkeypatch.setattr(seven_stage, "require_stage_approved", lambda _root, _stage:
                        (tmp_path, {"artifact": {"json": report.name}}))
    assert seven_stage.require_upstream(tmp_path, manifest)[0] == tmp_path
    assert seven_stage._comparison_case(stage5, request) == {
        "case": "synthetic-case", "public_result_sha256": workflow.hash_file(result)}
    monkeypatch.setattr(seven_stage, "current_revision", lambda _manifest, _stage:
                        {"artifact": {"json_sha256": workflow.hash_file(report)}})
    assert seven_stage._candidate_comparison(manifest, report, report, stage5, "synthetic-source",
                                              "synthetic-stage4", tmp_path / "bundle.json") == workflow.hash_file(report)
    with pytest.raises(ValueError, match="current approved exact-case"):
        seven_stage._candidate_comparison(manifest, result, report, stage5, "synthetic-source",
                                          "synthetic-stage4", tmp_path / "bundle.json")
    other = tmp_path / "other.json"
    other.write_text('{"synthetic": false}', encoding="utf-8")
    with pytest.raises(ValueError, match="exact request"):
        seven_stage._comparison_case(stage5, other)


def test_two_supported_synthetic_calls_keep_staged_package_identity(tmp_path: Path) -> None:
    root, request, other = _synthetic_candidate(tmp_path)
    package = root / "stage6" / "packages" / "synthetic-v1"
    before = {p.name: workflow.hash_file(p) for p in package.iterdir()}
    first = seven_stage._run_entry(package, request, tmp_path / "call-one" / "result.json")
    second = seven_stage._run_entry(package, other, tmp_path / "call-two" / "result.json")
    assert first["exit_code"] == second["exit_code"] == 0
    assert first["result_sha256"] != second["result_sha256"]
    assert {p.name: workflow.hash_file(p) for p in package.iterdir()} == before
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_synthetic_smoke_executes_verified_short_copy(tmp_path: Path) -> None:
    root, request, _other = _synthetic_candidate(tmp_path)
    candidate = workflow.read_json(root / "stage6" / "revision-0001" / "delivery_candidate.json")
    package = Path(candidate["package_path"])
    before = seven_stage._package_files(package, candidate["files"])
    invocation = seven_stage._run_entry(package, request, tmp_path / "readback" / "result.json",
                                        candidate["files"])
    assert invocation["exit_code"] == 0
    assert invocation["package_copy_verified"] is True
    assert invocation["command"][4] == str(package / "model.py")
    assert invocation["executed_command"][4] != invocation["command"][4]
    assert seven_stage._package_files(package, candidate["files"]) == before
    with pytest.raises(ValueError, match="changed or is missing|differs"):
        seven_stage._run_entry(package, request, tmp_path / "bad" / "result.json",
                               {**candidate["files"], "model.py": "0" * 64})


def test_revoked_delivery_stays_inactive_if_old_pointer_is_restored(tmp_path: Path,
                                                                    monkeypatch: pytest.MonkeyPatch) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    revision = workflow.current_revision(manifest, 6)
    candidate = workflow.read_json(root / revision["artifact"]["json"])
    record = {"schema_version": "step6.delivery_record.v1", "receipt_id": "synthetic-receipt",
              "candidate_artifact": revision["artifact"],
              "package_sha256": candidate["package_sha256"],
              "version": candidate["version"], "permission": candidate["permission"]}
    manifest["delivery"] = record  # Simulated stale pointer restored from saved state.
    manifest["delivery_history"] = [record, {
        "schema_version": "step6.delivery_revocation.v1", "receipt_id": record["receipt_id"],
        "version": record["version"], "reason": "Synthetic rollback."}]
    workflow._write_manifest(root, manifest)
    monkeypatch.setattr(seven_stage, "require_parent", lambda *_args: (root, manifest))
    monkeypatch.setattr(seven_stage, "stage_status", lambda *_args: {
        "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-receipt"}})
    with pytest.raises(ValueError, match="revoked"):
        seven_stage.require_delivery(root)
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"
    monkeypatch.undo()
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"
    assert workflow.load_workflow(root)[1]["delivery_history"][-1]["reason"] == "Synthetic rollback."


@pytest.mark.parametrize("fault", ["missing_readback", "changed_readback", "changed_package", "added_package"])
def test_synthetic_delivery_faults_block_effective_pointer(tmp_path: Path,
                                                          monkeypatch: pytest.MonkeyPatch,
                                                          fault: str) -> None:
    root, request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    revision = workflow.current_revision(manifest, 6)
    candidate = workflow.read_json(root / revision["artifact"]["json"])
    package = root / "deliveries" / candidate["version"]
    shutil.copytree(Path(candidate["package_path"]), package)
    result = root / "deliveries" / "readback" / "synthetic-result.json"
    invocation = seven_stage._run_entry(package, request, result)
    record = {"schema_version": "step6.delivery_record.v1", "receipt_id": "synthetic-receipt",
              "candidate_artifact": revision["artifact"], "package_sha256": candidate["package_sha256"],
              "version": candidate["version"], "permission": candidate["permission"],
              "path": str(package), "readback": {
                  "files": candidate["files"], "entry": candidate["entry"],
                  "request_sha256": candidate["request_sha256"], "result_path": str(result),
                  "result_sha256": invocation["result_sha256"], "invocation": invocation}}
    manifest["delivery"] = record
    manifest["delivery_history"] = [record]
    workflow._write_manifest(root, manifest)
    if fault == "missing_readback":
        result.unlink()
    elif fault == "changed_readback":
        result.write_text('{"value":999}', encoding="utf-8")
    else:
        if fault == "added_package":
            (package / "json.py").write_text("raise RuntimeError('shadowed json')\n", encoding="utf-8")
        else:
            (package / "model.py").write_text("changed", encoding="utf-8")
    monkeypatch.setattr(seven_stage, "require_parent", lambda *_args: (root, manifest))
    monkeypatch.setattr(seven_stage, "stage_status", lambda *_args: {
        "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-receipt"}})
    with pytest.raises(ValueError, match="changed|missing|undeclared"):
        seven_stage.require_delivery(root)
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"
    monkeypatch.undo()
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_tampered_synthetic_candidate_and_cross_scenario_proof_block(tmp_path: Path) -> None:
    root, request, other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    revision = workflow.current_revision(manifest, 6)
    candidate_path = root / revision["artifact"]["json"]
    original = candidate_path.read_bytes()
    candidate_path.write_bytes(original + b" ")
    assert workflow._revision_staleness(root, manifest, 6, revision)
    candidate_path.write_bytes(original)
    candidate = json.loads(original)
    entry = Path(candidate["package_path"]) / "model.py"
    entry.write_text(entry.read_text(encoding="utf-8") + "\n# changed\n", encoding="utf-8")
    assert workflow._revision_staleness(root, manifest, 6, revision)
    compared = json.loads(Path(revision["input_files"]["native_comparison"]["path"]).read_text())
    with pytest.raises(ValueError, match="exact request scenario"):
        seven_stage._comparison_case(compared, other)
    assert workflow.hash_file(request) != workflow.hash_file(other)
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_interrupted_synthetic_delivery_keeps_history_without_pointer(tmp_path: Path) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    pending = root / "deliveries" / "synthetic-v1.pending"
    pending.mkdir(parents=True)
    (pending / "model.py").write_text("changed", encoding="utf-8")
    with pytest.raises(ValueError, match="synthetic workflow"):
        seven_stage.deliver(root)
    assert pending.is_dir()
    assert workflow.load_workflow(root)[1]["delivery"] is None


def test_synthetic_copy_retry_and_idempotence_are_never_real_registration(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    pending = root / "deliveries" / ".pending" / "synthetic-v1"
    pending.mkdir(parents=True)
    staged = root / "stage6" / "packages" / "synthetic-v1"
    for name in ("model.py", "model_manifest.json"):
        (pending / name).write_bytes((staged / name).read_bytes())
    with monkeypatch.context() as gate:
        gate.setattr(seven_stage, "require_parent", lambda *_args: (root, manifest))
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-receipt"}})
        first = seven_stage.deliver(root)
        assert first["idempotent"] is False and not pending.exists()
        assert seven_stage.deliver(root)["idempotent"] is True
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"

    second_root, _request, _other = _synthetic_candidate(tmp_path / "changed")
    _, second_manifest = workflow.load_workflow(second_root)
    changed = second_root / "deliveries" / ".pending" / "synthetic-v1"
    changed.mkdir(parents=True)
    (changed / "model.py").write_text("changed", encoding="utf-8")
    with monkeypatch.context() as gate:
        gate.setattr(seven_stage, "require_parent", lambda *_args: (second_root, second_manifest))
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-receipt"}})
        second = seven_stage.deliver(second_root)
        assert second["delivery"]["version"] == "synthetic-v1"
        assert second["delivery"]["readback"]["files"] == workflow.read_json(
            second_root / workflow.current_revision(second_manifest, 6)["artifact"]["json"])["files"]
        assert not changed.exists()
    assert workflow.workflow_status(second_root)["active_delivery"] is None


def test_synthetic_final_version_collision_preserves_files_and_history(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    candidate = workflow.read_json(root / workflow.current_revision(manifest, 6)["artifact"]["json"])
    final = root / "deliveries" / candidate["version"]
    shutil.copytree(Path(candidate["package_path"]), final)
    (final / "model.py").write_bytes(b"synthetic older package")
    before_files = {path.relative_to(final): path.read_bytes() for path in final.rglob("*") if path.is_file()}
    before_manifest = (root / "workflow.json").read_bytes()
    with monkeypatch.context() as gate:
        gate.setattr(seven_stage, "require_parent", lambda *_args: (root, manifest))
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-receipt"}})
        with pytest.raises(ValueError, match="generated bundle file changed|existing delivery version differs"):
            seven_stage.deliver(root)
    assert {path.relative_to(final): path.read_bytes() for path in final.rglob("*") if path.is_file()} == before_files
    assert (root / "workflow.json").read_bytes() == before_manifest
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_synthetic_version_suffix_cannot_move_older_delivery(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root, request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    original = workflow.read_json(root / workflow.current_revision(manifest, 6)["artifact"]["json"])

    def new_version(version: str) -> None:
        source = Path(original["package_path"])
        package = root / "stage6" / "packages" / version
        shutil.copytree(source, package)
        candidate = {**original, "version": version, "package_path": str(package)}
        workflow.append_stage_artifact(root, 6, "delivery_candidate.json", candidate,
                                       "Synthetic test candidate; no real acceptance.",
                                       input_files={"smoke_request": request})

    with monkeypatch.context() as gate:
        gate.setattr(seven_stage, "require_parent", lambda *_args: (root, manifest))
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-one"}})
        new_version("v1.pending")
        first = seven_stage.deliver(root)["delivery"]
        older = root / "deliveries" / "v1.pending"
        older_files = {p.name: p.read_bytes() for p in older.iterdir() if p.is_file()}
        new_version("v1")
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-two"}})
        second = seven_stage.deliver(root)["delivery"]
        assert second["version"] == "v1"
        assert older.is_dir()
        assert {p.name: p.read_bytes() for p in older.iterdir() if p.is_file()} == older_files
        assert first in workflow.load_workflow(root)[1]["delivery_history"]
        new_version("readback")
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-three"}})
        assert seven_stage.deliver(root)["delivery"]["version"] == "readback"
        assert older.is_dir() and (root / "deliveries" / "v1").is_dir()
        assert (root / "deliveries" / ".readback" / "readback").is_dir()
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"


def test_synthetic_interrupted_partial_delivery_retries_without_pointer(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    original_copy = seven_stage.shutil.copy2
    attempts = 0

    def interrupted_copy(source, target):
        nonlocal attempts
        attempts += 1
        if attempts == 2:
            Path(target).write_bytes(Path(source).read_bytes()[:5])
            raise OSError("synthetic interrupted copy")
        return original_copy(source, target)

    with monkeypatch.context() as gate:
        gate.setattr(seven_stage, "require_parent", lambda *_args: (root, manifest))
        gate.setattr(seven_stage, "stage_status", lambda *_args: {
            "status": "approved", "human": {"decision": "approve", "receipt_id": "synthetic-receipt"}})
        gate.setattr(seven_stage.shutil, "copy2", interrupted_copy)
        with pytest.raises(OSError, match="synthetic interrupted copy"):
            seven_stage.deliver(root)
        assert workflow.load_workflow(root)[1]["delivery"] is None
        pending = root / "deliveries" / ".pending" / "synthetic-v1"
        assert pending.is_dir()
        assert any(path.is_file() and path.stat().st_size == 5 for path in pending.rglob("*"))
        unrelated = root / "deliveries" / ".pending" / "another-version" / "keep.txt"
        unrelated.parent.mkdir(parents=True)
        unrelated.write_bytes(b"untouched")
        gate.setattr(seven_stage.shutil, "copy2", original_copy)
        assert seven_stage.deliver(root)["delivery"]["version"] == "synthetic-v1"
        assert unrelated.read_bytes() == b"untouched"
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_synthetic_candidate_copy_resumes_only_exact_partial_files(tmp_path: Path) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    _, manifest = workflow.load_workflow(root)
    candidate = workflow.read_json(root / workflow.current_revision(manifest, 6)["artifact"]["json"])
    source = Path(candidate["package_path"])
    ledger = candidate["files"]
    partial = root / "stage6" / "packages" / "synthetic-retry"
    partial.mkdir()
    (partial / "model.py").write_bytes((source / "model.py").read_bytes())
    seven_stage._resume_package_copy(source, partial, ledger)
    assert seven_stage._package_files(partial, ledger) == ledger
    (partial / "model.py").write_bytes((source / "model.py").read_bytes()[:5])
    seven_stage._resume_package_copy(source, partial, ledger)
    assert seven_stage._package_files(partial, ledger) == ledger
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_isolated_entry_keeps_exact_package_inventory_after_sibling_import(tmp_path: Path) -> None:
    package = tmp_path / "synthetic-package"
    package.mkdir()
    (package / "helper.py").write_text("VALUE = 2\n", encoding="utf-8")
    (package / "model.py").write_text(
        "import json\nfrom pathlib import Path\nfrom helper import VALUE\n"
        "Path(__import__('sys').argv[-1]).write_text(json.dumps({'value': VALUE}))\n",
        encoding="utf-8")
    ledger = {name: workflow.hash_file(package / name) for name in ("model.py", "helper.py")}
    request = tmp_path / "synthetic-request.json"
    request.write_text("{}", encoding="utf-8")
    record = seven_stage._run_entry(package, request, tmp_path / "result" / "result.json")
    assert record["command"][1:4] == ["-E", "-s", "-B"]
    assert seven_stage._package_files(package, ledger) == ledger
    assert not (package / "__pycache__").exists()


def test_synthetic_parent_cannot_create_a_publication_candidate(tmp_path: Path) -> None:
    root, request, _other = _synthetic_candidate(tmp_path, create_revision=False)
    comparison = tmp_path / "synthetic-comparison.json"
    with pytest.raises(ValueError, match="synthetic workflow"):
        seven_stage.create_candidate(root, "synthetic-v1", request, comparison)
    assert workflow.current_revision(workflow.load_workflow(root)[1], 6) is None
    assert workflow.workflow_status(root)["active_delivery"] is None


def test_parent_mutation_keeps_synthetic_candidate_stale(tmp_path: Path) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    parent_manifest = tmp_path / "legacy" / "workflow.json"
    parent_manifest.write_bytes(parent_manifest.read_bytes() + b" ")
    assert workflow.workflow_status(root)["active_delivery"] is None
    assert create_step7_report(root)["status"] == "blocked"


def test_v2_rejection_route_is_bounded_to_its_own_stage(tmp_path: Path,
                                                            monkeypatch: pytest.MonkeyPatch) -> None:
    root, _request, _other = _synthetic_candidate(tmp_path)
    monkeypatch.setattr(workflow, "_revision_staleness", lambda *_args: [])
    monkeypatch.setattr(seven_stage, "candidate_staleness", lambda *_args: [])
    with pytest.raises(ValueError, match="route to one of: 6"):
        workflow.record_decision(root, 6, "agent", "reject", "Synthetic rejection.", return_to=5)
    assert workflow.load_workflow(root)[1]["delivery"] is None
