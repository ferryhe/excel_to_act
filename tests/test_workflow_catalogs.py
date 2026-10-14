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


def test_legacy_catalogs_and_new_delivery_catalog_expose_their_scopes() -> None:
    runner = CliRunner()
    for stage in range(1, 6):
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
        assert "native v2 Stage 1–5 only" in tools_by_name["workflow.confirm"]["availability"]
        assert "attached v2 inherited stages and legacy v1 are read-only" in tools_by_name["workflow.reject"]["availability"]
        assert "typesafe" not in confirm and "typesafe" not in reject
        if stage >= 3:
            assert "disabled" in tools_by_name["workflow.typesafe"]["availability"]
        if stage == 3:
            assert "disabled" in tools_by_name["workflow.delegate"]["availability"]
    delivery = json.loads(runner.invoke(app, ["step6", "tools"]).stdout)
    commands = {entry["name"]: entry["command"] for entry in delivery["tools"]}
    assert "step6 candidate" in commands["step6.candidate"]
    assert "--stage 6" in commands["workflow.confirm"]
    assert "step6 deliver" in commands["step6.deliver"]
    delivery_reject = next(item for item in delivery["tools"] if item["name"] == "workflow.reject")
    assert "--stage 6 --reviewer agent|human" in delivery_reject["command"]
    assert "--return-to 1|2|3|4|5|6" in delivery_reject["command"]
    assert "--return-to 6 --message" in delivery_reject["attached_command"]

    final = json.loads(runner.invoke(app, ["step7", "tools"]).stdout)
    reject = next(item for item in final["tools"] if item["name"] == "workflow.reject")
    assert "--return-to 1|2|3|4|5|6|7" in reject["command"]
    assert "--return-to 6|7" in reject["attached_command"]
    assert "legacy v1" in reject["availability"]

    source_root = Path(__file__).resolve().parents[1] / "src" / "excel_to_act" / "steps"
    assert "--return-to 1|2|3|4|5|6" in (source_root / "step6" / "agent.md").read_text(encoding="utf-8")
    assert "--return-to 1|2|3|4|5|6|7" in (source_root / "step7" / "agent.md").read_text(encoding="utf-8")


def test_workflow_guide_describes_native_and_attached_review_routes() -> None:
    guide = (Path(__file__).resolve().parents[1] / "docs" / "workflow.md").read_text(encoding="utf-8")
    assert "step1 report --run RUN_DIR --out WORKFLOW_DIR" in guide
    assert "Native v2 uses `workflow confirm` and `workflow reject`" in guide
    assert "Attached v2 keeps inherited v1 Stage 1–5 receipts as read-only provenance" in guide
    assert "`workflow delegate` and `workflow typesafe` reject writes" in guide
    assert "Step 7-only report edits return to Stage 7" in guide
    assert "attached v2 cannot return to inherited Stages 1–5" in guide
    assert "report defects to Step 6" not in guide
    for stage in ("3b", "4", "5"):
        row = next(line for line in guide.splitlines() if line.startswith(f"| {stage} —"))
        assert "Agent" in row and "actual human" in row
        assert "delegation" not in row.lower()


def test_synthetic_stage_one_human_rejection_uses_advertised_route(tmp_path: Path) -> None:
    root = tmp_path / "workflow"
    create_workflow(root, {"workbook_sha256": "a" * 64, "run_id": "synthetic-run"})
    append_stage_artifact(root, 1, "checkpoint.json", {"status": "ready_for_review"}, "Synthetic checkpoint.")

    result = CliRunner().invoke(app, [
        "workflow", "reject", "--workflow", str(root), "--stage", "1", "--reviewer", "human",
        "--return-to", "1", "--message", "Synthetic user requested revision.",
    ])

    assert result.exit_code == 1
    assert "read-only" in result.stdout
    from excel_to_act.steps.conversion_workflow import record_decision
    record_decision(root, 1, "human", "reject", "Synthetic user requested revision.", return_to=1)
    _workflow_root, manifest = load_workflow(root)
    state = stage_status(root, manifest, 1)
    assert state["status"] == "rejected"
    assert manifest["stages"]["1"]["revisions"][0]["decisions"][-1]["return_to"] == 1


def test_later_step_agent_contracts_describe_current_human_route() -> None:
    source_root = Path(__file__).resolve().parents[1] / "src" / "excel_to_act" / "steps"
    for stage in (3, 4, 5):
        agent = (source_root / f"step{stage}" / "agent.md").read_text(encoding="utf-8")
        assert "actual-human" in agent or "actual human" in agent
        assert "Legacy TypeSafe delegation" in agent or "legacy TypeSafe delegation" in agent
        assert "A rejection from any reviewer makes that revision ineligible for later approvals" in agent
    assert "input-boundary catalog always requires actual human review" in (source_root / "step3" / "agent.md").read_text(encoding="utf-8")
    assert "actual human decision on this exact pair" in (source_root / "step6" / "agent.md").read_text(encoding="utf-8")


def test_emitted_step3_and_step7_agents_describe_current_seven_stage_review_route() -> None:
    runner = CliRunner()
    step3 = runner.invoke(app, ["step3", "agent"])
    step7 = runner.invoke(app, ["step7", "agent"])

    assert step3.exit_code == step7.exit_code == 0
    assert "Stage 6 packages a candidate for review" in step3.stdout
    assert "internally delivers the exact package only after Agent and explicit human approval" in step3.stdout
    assert "Stage 7 creates the final report" in step3.stdout
    assert "Stage 6 summarizes" not in step3.stdout
    assert "actual human" in step7.stdout
    assert "live TypeSafe review commands are disabled in v2" in step7.stdout
    assert "Use TypeSafe only" not in step7.stdout
    assert "explicitly delegated reviews" not in step7.stdout


def test_current_guides_show_only_the_executable_human_review_route() -> None:
    root = Path(__file__).resolve().parents[1]
    english = (root / "README.md").read_text(encoding="utf-8")
    chinese = (root / "README.zh-CN.md").read_text(encoding="utf-8")
    guide = (root / "docs" / "design" / "generic_agent_contracts.md").read_text(encoding="utf-8")

    assert "Agent → human or scoped delegation" not in english
    assert "Agent → 人工或范围匹配的委托" not in chinese
    assert "seven packaged `agent.md` files" in guide
    assert "live TypeSafe review commands are disabled in v2" in guide
    assert "Use TypeSafe only" not in guide


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


def test_step7_skill_and_template_are_packaged_cli_resources() -> None:
    package = files("excel_to_act.steps.step7").joinpath("excel-to-act-step7")
    skill_text = package.joinpath("SKILL.md").read_text(encoding="utf-8")
    template_text = package.joinpath("assets", "conversion_report.md").read_text(encoding="utf-8")

    runner = CliRunner()
    skill_result = runner.invoke(app, ["step7", "skill"])
    template_result = runner.invoke(app, ["step7", "template"])
    catalog_result = runner.invoke(app, ["step7", "tools"])

    assert skill_result.exit_code == 0, skill_result.stdout
    assert template_result.exit_code == 0, template_result.stdout
    assert catalog_result.exit_code == 0, catalog_result.stdout
    assert skill_result.stdout.rstrip() == skill_text.rstrip()
    assert template_result.stdout.rstrip() == template_text.rstrip()
    assert "## Internal package delivery" in template_text
    for field in ("acceptance receipt", "version", "permission", "package", "entry", "readback"):
        assert field in template_text.lower()
    assert "Agent, then actual human" in template_text
    assert "active delegation" not in template_text

    catalog = json.loads(catalog_result.stdout)
    commands = {entry["name"]: entry["command"] for entry in catalog["tools"]}
    assert commands["step7.skill"] == "step7 skill"
    assert commands["step7.template"] == "step7 template"


def test_step6_packaged_skill_routes_current_delivery_and_legacy_report() -> None:
    package = files("excel_to_act.steps.step6").joinpath("excel-to-act-step6")
    metadata = package.joinpath("agents", "openai.yaml").read_text(encoding="utf-8")
    skill = package.joinpath("SKILL.md").read_text(encoding="utf-8")

    assert "Step 6 internal tool delivery" in metadata
    assert "Step 6 model conversion report" not in metadata
    assert "step6 candidate" in skill and "step6 deliver" in skill
    assert "exact JSON/Markdown pair" in skill and "actual human" in skill
    assert "delivery/readback" in skill and "step7 report" in skill
    assert "Historical v1" in skill and "read-only" in skill
    assert "Step 7" in skill and "final report" in skill
