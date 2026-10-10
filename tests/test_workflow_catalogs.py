from __future__ import annotations

import json
from importlib.resources import files
from pathlib import Path

from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app


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

        agent = (source_root / f"step{stage}" / "agent.md").read_text(encoding="utf-8")
        assert "workflow status --workflow " in agent
        assert "workflow confirm --workflow " in agent
        assert f"--stage {stage} --reviewer agent|human" in agent
        assert "workflow reject --workflow " in agent

    step3_agent = (source_root / "step3" / "agent.md").read_text(encoding="utf-8")
    assert "step4 discover --workflow DIR" in step3_agent


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
