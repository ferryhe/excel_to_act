from __future__ import annotations

import json
from importlib.resources import files
from pathlib import Path

from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app
from excel_to_act.steps.conversion_workflow import append_stage_artifact, create_workflow, load_workflow, stage_status


def test_stages_one_to_three_publish_review_commands_and_agent_contracts() -> None:
    runner = CliRunner()
    source_root = Path(__file__).resolve().parents[1] / "src" / "excel_to_act" / "steps"

    for stage in (1, 2, 3):
        catalog_result = runner.invoke(app, [f"step{stage}", "tools"])
        assert catalog_result.exit_code == 0, catalog_result.stdout
        catalog = json.loads(catalog_result.stdout)
        tools_by_name = {entry["name"]: entry for entry in catalog["tools"]}

        assert {"workflow.status", "workflow.confirm", "workflow.reject"} <= tools_by_name.keys()
        assert tools_by_name["workflow.status"]["command"] == "workflow status --workflow DIR"
        assert f"--stage {stage}" in tools_by_name["workflow.confirm"]["command"]
        assert f"--stage {stage}" in tools_by_name["workflow.reject"]["command"]
        assert "--decision approve " in tools_by_name["workflow.confirm"]["command"]
        assert "reject" not in tools_by_name["workflow.confirm"]["command"]
        assert "--return-to " + "|".join(str(value) for value in range(1, stage + 1)) in tools_by_name["workflow.reject"]["command"]
        assert "--decision approve|reject" not in "\n".join(catalog.get("review_workflow", []))

        agent = (source_root / f"step{stage}" / "agent.md").read_text(encoding="utf-8")
        assert "workflow status --workflow " in agent
        assert "workflow confirm --workflow " in agent
        assert f"--stage {stage} --reviewer agent|human" in agent
        assert "workflow reject --workflow " in agent

    step3_agent = (source_root / "step3" / "agent.md").read_text(encoding="utf-8")
    assert "step4 discover --workflow DIR" in step3_agent


def test_all_six_catalogs_expose_approval_only_confirm_and_routed_rejection() -> None:
    runner = CliRunner()
    for stage in range(1, 7):
        result = runner.invoke(app, [f"step{stage}", "tools"])
        assert result.exit_code == 0, result.stdout
        catalog = json.loads(result.stdout)
        tools_by_name = {entry["name"]: entry for entry in catalog["tools"]}
        confirm = tools_by_name["workflow.confirm"]["command"]
        reject = tools_by_name["workflow.reject"]["command"]
        routes = "|".join(str(value) for value in range(1, stage + 1))
        assert "--decision approve " in confirm
        assert "reject" not in confirm
        assert f"--return-to {routes}" in reject


def test_synthetic_stage_one_human_rejection_uses_advertised_route(tmp_path: Path) -> None:
    root = tmp_path / "workflow"
    create_workflow(root, {"workbook_sha256": "a" * 64, "run_id": "synthetic-run"})
    append_stage_artifact(root, 1, "checkpoint.json", {"status": "ready_for_review"}, "Synthetic checkpoint.")

    result = CliRunner().invoke(app, [
        "workflow", "reject", "--workflow", str(root), "--stage", "1", "--reviewer", "human",
        "--return-to", "1", "--message", "Synthetic user requested revision.",
    ])

    assert result.exit_code == 0, result.stdout
    record = json.loads(result.stdout)
    assert record["status"] == "pass"
    _workflow_root, manifest = load_workflow(root)
    state = stage_status(root, manifest, 1)
    assert state["status"] == "rejected"
    assert manifest["stages"]["1"]["revisions"][0]["decisions"][-1]["return_to"] == 1


def test_later_step_agent_contracts_describe_scoped_human_restoration() -> None:
    source_root = Path(__file__).resolve().parents[1] / "src" / "excel_to_act" / "steps"
    for stage in (3, 4, 5, 6):
        agent = (source_root / f"step{stage}" / "agent.md").read_text(encoding="utf-8")
        assert "current, unrejected JSON/Markdown pair and revision" in agent
        assert "A rejection from any reviewer makes that revision ineligible for later approvals" in agent
    assert "input-boundary catalog always requires actual human review" in (source_root / "step3" / "agent.md").read_text(encoding="utf-8")


def test_step4_catalog_describes_catalog_derived_external_vectors_and_current_limits() -> None:
    result = CliRunner().invoke(app, ["step4", "tools"])
    assert result.exit_code == 0, result.stdout
    catalog = json.loads(result.stdout)
    limits = "\n".join(catalog["limits"])
    assert "five external 106-value rate vectors" in limits
    assert "530 captured coordinates" in limits
    assert "confirmed catalog" in limits and "callers cannot override ranges" in limits
    assert "seven approved Main cells" in limits
    assert "upstream CI producer formulas remain excluded" in limits


def test_step5_catalog_and_agent_distinguish_capture_from_compared_scope() -> None:
    runner = CliRunner()
    result = runner.invoke(app, ["step5", "tools"])
    assert result.exit_code == 0, result.stdout
    catalog = json.loads(result.stdout)
    limits = "\n".join(catalog["limits"])
    assert "exploratory evidence outside the plan" in limits
    assert "does not expand the approved comparison scope" in limits
    assert "exact compared active formula addresses" in limits

    source_root = Path(__file__).resolve().parents[1] / "src" / "excel_to_act" / "steps"
    agent = (source_root / "step5" / "agent.md").read_text(encoding="utf-8")
    assert "Optional `--target` and `--range` captures may include exploratory evidence outside the plan" in agent
    assert "compared.comparison_addresses" in agent
    assert "captured-only for active-formula comparison" in agent


def test_step6_skill_and_template_are_packaged_cli_resources() -> None:
    package = files("excel_to_act.steps.step6").joinpath("excel-to-act-step6")
    skill_text = package.joinpath("SKILL.md").read_text(encoding="utf-8")
    template_text = package.joinpath("assets", "conversion_report.md").read_text(encoding="utf-8")

    runner = CliRunner()
    skill_result = runner.invoke(app, ["step6", "skill"])
    template_result = runner.invoke(app, ["step6", "template"])
    catalog_result = runner.invoke(app, ["step6", "tools"])

    assert skill_result.exit_code == 0, skill_result.stdout
    assert template_result.exit_code == 0, template_result.stdout
    assert catalog_result.exit_code == 0, catalog_result.stdout
    assert skill_result.stdout.rstrip() == skill_text.rstrip()
    assert template_result.stdout.rstrip() == template_text.rstrip()

    catalog = json.loads(catalog_result.stdout)
    commands = {entry["name"]: entry["command"] for entry in catalog["tools"]}
    assert commands["step6.skill"] == "step6 skill"
    assert commands["step6.template"] == "step6 template"
