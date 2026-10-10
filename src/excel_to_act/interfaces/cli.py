"""CLI entrypoint for Phase 1 workflows."""

from __future__ import annotations

from pathlib import Path
import json
from importlib.resources import files

import typer

from excel_to_act.orchestrator.phase1 import Phase1Orchestrator
from excel_to_act.ingest.control_artifacts import ARTIFACT_FILES, build_single_control_artifact, clear_declared_vba_sources, evaluate_control_output, vba_handoff_markdown
from excel_to_act.schemas import VbaHandoff
from excel_to_act.steps.step2.workflow import build_index, execute_tool as execute_step2_tool, validate_saved_index
from excel_to_act.steps.step2.tools import tool_catalog as step2_tool_catalog
from excel_to_act.steps.step2.prepare import prepare as prepare_step2
from excel_to_act.steps.step2.query import QueryFailure, query as query_step2, validate_evidence_packet
from excel_to_act.steps.step2.trace import trace as trace_step2
from excel_to_act.steps.step3 import build_dependencies, build_fields, build_plan, prepare_analysis, query as step3_query, trace as step3_trace, validate_analysis
from excel_to_act.steps.step3.tools import tool_catalog as step3_tool_catalog
from excel_to_act.steps.step3.semantic import build_semantic_plan, validate_semantic_plan
from excel_to_act.steps.step3.profiling import (
    build_source_candidate_trace,
    profile_source_candidates,
    profile_source_families,
)
from excel_to_act.steps.step3.input_boundary import create_input_boundary
from excel_to_act.steps.conversion_workflow import record_decision, register_review_delegation, submit_typesafe_review, workflow_status
from excel_to_act.steps.step1.checkpoint import create_checkpoint as create_step1_checkpoint
from excel_to_act.steps.step2.checkpoint import create_checkpoint as create_step2_checkpoint
from excel_to_act.steps.step3.design import create_design_report
from excel_to_act.steps.step4 import (
    capture_external_inputs, create_implementation_plan, discover_active_trace, generate_model,
)
from excel_to_act.steps.step5 import capture_oracle, reconcile, validate_generated
from excel_to_act.steps.step5.tools import tool_catalog as step5_tool_catalog
from excel_to_act.steps.step6 import create_conversion_report
from excel_to_act.steps.step6.tools import tool_catalog as step6_tool_catalog
from excel_to_act.steps.step4.tools import tool_catalog as step4_tool_catalog
from excel_to_act.steps.step1.workflow import agent_definition, auto_recover, convert_directory, execute_tool, finalize_run
from excel_to_act.steps.step1.tools import tool_catalog
from excel_to_act.schemas import WorkbookView
from excel_to_act.views import compile_views, serialize_views, validate_agent_output

app = typer.Typer(help="Excel to actuarial model decomposition toolkit")
step1_app = typer.Typer(help="Human Step1 raw-directory conversion and checked handoff")
checkbox_app = typer.Typer(help="Identify, convert, and evaluate worksheet checkbox bindings")
activex_app = typer.Typer(help="Identify, convert, and evaluate ActiveX event declarations")
vba_app = typer.Typer(help="Identify, export, and evaluate VBA source handoff")
step2_app = typer.Typer(help="Index Step 1 handoffs for downstream agents")
step3_app = typer.Typer(help="Build static Step 3 analysis and design evidence")
step4_app = typer.Typer(help="Generate standalone Python from an approved Step 3 design")
step5_app = typer.Typer(help="Validate generated Python and reconcile it with native Excel")
step6_app = typer.Typer(help="Write the final report from accepted conversion checkpoints")
workflow_app = typer.Typer(help="Review and track the six conversion checkpoints")
views_app = typer.Typer(help="Compile deterministic source-addressable workbook views")
app.add_typer(step1_app, name="step1")
step1_app.add_typer(checkbox_app, name="checkbox")
step1_app.add_typer(activex_app, name="activex")
step1_app.add_typer(vba_app, name="vba")
app.add_typer(step2_app, name="step2")
app.add_typer(step3_app, name="step3")
app.add_typer(step4_app, name="step4")
app.add_typer(step5_app, name="step5")
app.add_typer(step6_app, name="step6")
app.add_typer(workflow_app, name="workflow")
app.add_typer(views_app, name="views")


@app.callback()
def main() -> None:
    """Excel to actuarial model decomposition toolkit."""


@app.command()
def inspect(
    workbook: Path = typer.Argument(..., help=".xlsx/.xlsm workbook to inspect"),
    out: Path = typer.Option(Path("artifacts"), "--out", "-o", help="Output directory for Phase 1 JSON artifacts"),
) -> None:
    """Inspect a workbook and produce Phase 1 artifacts."""

    metadata = Phase1Orchestrator().run(workbook, out)
    typer.echo(f"Wrote Phase 1 artifacts to {out}")
    typer.echo(f"Run ID: {metadata.run_id}")
    status = metadata.completeness_status or "unknown"
    handoff = next((a.path for a in metadata.artifacts if a.name == "handoff.md"), None)
    if handoff:
        typer.echo(f"Handoff: {handoff}")
    typer.echo(f"Completeness: {status}")
    if status == "fail":
        completeness = next((a.path for a in metadata.artifacts if a.name == "completeness.json"), "completeness.json")
        try:
            blockers = json.loads(Path(completeness).read_text(encoding="utf-8")).get("blocking_reasons", [])
            reason = blockers[0] if blockers else "Input could not be read or the output is incomplete."
        except (OSError, ValueError, IndexError, TypeError):
            reason = "Input could not be read or the output is incomplete."
        typer.echo(f"Reason: {reason}", err=True)
        typer.echo(f"Diagnostic: {completeness}", err=True)
        raise typer.Exit(code=1)


def _emit_json(result: dict) -> None:
    typer.echo(json.dumps(result, ensure_ascii=False, indent=2, default=str))


def _control_models(kind: str, workbook: Path):
    return build_single_control_artifact(kind, workbook)


def _control_identify(kind: str, workbook: Path) -> None:
    artifact, _sources = _control_models(kind, workbook)
    _emit_json(artifact.model_dump(mode="json"))


def _control_convert(kind: str, workbook: Path, out: Path, dry_run: bool) -> None:
    artifact, sources = _control_models(kind, workbook)
    artifact_name = ARTIFACT_FILES[kind]
    outputs = [artifact_name]
    if kind == "vba":
        outputs.extend(sorted(sources))
        outputs.append("vba_handoff.md")
    plan = {"tool": f"step1.{kind}.convert", "status": artifact.status, "dry_run": dry_run, "workbook": str(workbook.expanduser().resolve()), "output": str(out.expanduser().resolve()), "records": len(getattr(artifact, "bindings", getattr(artifact, "controls", getattr(artifact, "modules", [])))), "files": outputs, "diagnostics": artifact.diagnostics}
    if not dry_run:
        out = out.expanduser().resolve()
        out.mkdir(parents=True, exist_ok=True)
        if kind == "vba":
            previous = None
            try:
                previous = VbaHandoff.model_validate_json((out / artifact_name).read_bytes())
            except (OSError, ValueError):
                pass
            clear_declared_vba_sources(out, previous)
        (out / artifact_name).write_bytes((json.dumps(artifact.model_dump(mode="json"), ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode("utf-8"))
        for relative, content in sources.items():
            path = out / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(content)
        if kind == "vba":
            (out / "vba_handoff.md").write_bytes(vba_handoff_markdown(artifact).encode("utf-8"))
        plan["written"] = True
    _emit_json(plan)


def _control_evaluate(kind: str, workbook: Path, out: Path) -> None:
    result = evaluate_control_output(kind, workbook, out)
    _emit_json(result)
    if result["status"] == "fail":
        raise typer.Exit(code=1)


@step2_app.command("agent")
def step2_agent() -> None:
    """Print the packaged Step 2 host-agent contract."""

    typer.echo(files("excel_to_act.steps.step2").joinpath("agent.md").read_text(encoding="utf-8"))


@step1_app.command("report")
def step1_report(
    run: Path = typer.Option(..., "--run", exists=True, file_okay=False, readable=True,
                             help="Finalized Step 1 per-source run directory"),
    out: Path = typer.Option(..., "--out", help="New six-stage workflow directory"),
) -> None:
    """Create a read-only checkpoint from finalized Step 1 artifacts."""
    result = create_step1_checkpoint(run, out)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step2_app.command("tools")
def step2_tools() -> None:
    """Print the initial machine-readable Step 2 action catalogue."""

    _emit_json(step2_tool_catalog())


@step2_app.command("prepare")
def step2_prepare(
    index: Path = typer.Option(..., "--index", help="Validated native Step 2 index.json"),
    step1_root: Path = typer.Option(..., "--step1-root", help="Step 1 output root for resolving indexed artifacts"),
    out: Path = typer.Option(..., "--out", help="Reading package output directory"),
    scope: Path | None = typer.Option(None, "--scope", help="Source/run-bound analysis.scope.v1 JSON"),
    resume: bool = typer.Option(False, "--resume", help="Reuse a matching package after verifying every saved output hash"),
    dry_run: bool = typer.Option(False, "--dry-run", help="Report capabilities and planned files without writing"),
) -> None:
    """Prepare validated Step 1 and Step 2 artifacts for progressive reading."""
    result = prepare_step2(index, step1_root, out, scope_path=scope, resume=resume, dry_run=dry_run)
    _emit_json(result)
    if result["status"] == "blocked":
        raise typer.Exit(code=1)


@step2_app.command("query")
def step2_query(
    manifest: Path = typer.Option(..., "--manifest", exists=True, dir_okay=False, readable=True, help="#30 prepared reading manifest.json"),
    source_id: str = typer.Option(..., "--source-id", help="Exact source ID from the manifest; never inferred"),
    kind: str = typer.Option(..., "--kind", help="overview, sheet, cell, range, name, control, vba, or feature"),
    target: str | None = typer.Option(None, "--target", help="Cell, name, control, module, feature, or sheet target"),
    sheet: str | None = typer.Option(None, "--sheet", help="Worksheet context for a selector"),
    range_address: str | None = typer.Option(None, "--range", help="A1 range for a range query"),
    budget: int = typer.Option(1200, "--budget", min=1, help="Estimated token budget; records are never split"),
    cursor: str | None = typer.Option(None, "--cursor", help="Continue the same source-bound selector page"),
    out: Path | None = typer.Option(None, "--out", help="Optional evidence packet path"),
) -> None:
    """Read a bounded page from a prepared Step 2 evidence package."""
    try:
        result = query_step2(manifest, source_id, kind, target=target, sheet=sheet,
                             address=range_address, budget=budget, cursor=cursor)
    except QueryFailure as exc:
        _emit_json({"tool": "step2.query", "status": exc.status, "diagnostics": [exc.diagnostic]})
        raise typer.Exit(code=1) from exc
    if out is not None:
        out = out.expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    _emit_json(result)
    if result.get("status") in {"integrity_failed", "unavailable", "unsupported", "not_found", "needs_selection", "out_of_scope", "needs_scope_resolution", "oversized"}:
        raise typer.Exit(code=1)


@step2_app.command("trace")
def step2_trace(
    manifest: Path = typer.Option(..., "--manifest", exists=True, dir_okay=False, readable=True),
    source_id: str = typer.Option(..., "--source-id", help="Exact manifest source ID; required"),
    kind: str = typer.Option(..., "--kind", help="cell, range, or name"),
    target: str = typer.Option(..., "--target", help="A1 cell/range or defined name"),
    direction: str = typer.Option(..., "--direction", help="upstream dependencies, downstream consumers, or both"),
    sheet: str | None = typer.Option(None, "--sheet", help="Worksheet or name scope"),
    max_depth: int = typer.Option(8, "--max-depth", help="Depth bound (0..100)"),
    max_nodes: int = typer.Option(100, "--max-nodes", help="Node bound (1..10000)"),
    max_edges: int = typer.Option(200, "--max-edges", help="Edge bound (1..20000)"),
    out: Path | None = typer.Option(None, "--out", help="Optional evidence packet path"),
) -> None:
    """Trace a bounded static dependency path using a prepared graph and source records."""
    try:
        result = trace_step2(manifest, source_id, kind, target, sheet=sheet, direction=direction,
                             max_depth=max_depth, max_nodes=max_nodes, max_edges=max_edges)
        if out is not None:
            out = out.expanduser().resolve()
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except (QueryFailure, OSError) as exc:
        diagnostic = exc.diagnostic if isinstance(exc, QueryFailure) else {"code": "trace_write_failed", "severity": "error", "message": str(exc)}
        _emit_json({"tool": "step2.trace", "status": exc.status if isinstance(exc, QueryFailure) else "unavailable", "diagnostics": [diagnostic]})
        raise typer.Exit(code=1) from exc
    _emit_json(result)
    if result["status"] == "needs_selection":
        raise typer.Exit(code=1)


@step2_app.command("index")
def step2_index(
    handoff: Path = typer.Option(..., "--handoff", help="Step 1 source or batch handoff JSON"),
    step1_root: Path = typer.Option(..., "--step1-root", help="Step 1 output root for resolving handoff and artifact paths"),
    out: Path = typer.Option(Path("output/step2_index"), "--out", "-o", help="Step 2 index output directory"),
    resume: bool = typer.Option(False, "--resume", help="Reuse validated entries with unchanged inputs and saved outputs"),
) -> None:
    """Build a deterministic index from a current Step 1 handoff."""

    result = build_index(handoff, step1_root, out, resume=resume)
    _emit_json(result)
    if result["status"] == "blocked":
        raise typer.Exit(code=1)


@step2_app.command("report")
def step2_report(
    index: Path = typer.Option(..., "--index", exists=True, dir_okay=False, readable=True,
                               help="Saved Step 2 index.json"),
    step1_root: Path = typer.Option(..., "--step1-root", exists=True, file_okay=False, readable=True,
                                    help="Step 1 output root for validating index references"),
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False,
                                  help="Workflow directory containing the approved Step 1 checkpoint"),
    source_id: str | None = typer.Option(None, "--source-id", help="Select one eligible indexed source"),
) -> None:
    """Create the approved-source Step 2 index checkpoint without rewriting it."""
    result = create_step2_checkpoint(index, step1_root, workflow, source_id=source_id)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@workflow_app.command("status")
def workflow_status_command(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
) -> None:
    """Show stage artifact freshness, decisions, and the next permitted stage."""
    try:
        _emit_json(workflow_status(workflow))
    except Exception as exc:
        _emit_json({"schema_version": "conversion.workflow_status.v1", "status": "blocked", "reason": str(exc)})
        raise typer.Exit(code=1) from exc


@workflow_app.command("delegate")
def workflow_delegate(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    authorization: Path = typer.Option(..., "--authorization", help="Human-authored workflow-bound delegation JSON"),
) -> None:
    """Register an explicit, source-bound TypeSafe review delegation."""
    try:
        _emit_json(register_review_delegation(workflow, authorization))
    except Exception as exc:
        _emit_json({"status": "blocked", "reason": str(exc)})
        raise typer.Exit(code=1) from exc


@workflow_app.command("typesafe")
def workflow_typesafe(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    stage: int = typer.Option(..., "--stage", min=3, max=6),
    request: Path = typer.Option(..., "--request", exists=True, dir_okay=False,
                                 help="Current-artifact-bound TypeSafe Choice request JSON"),
) -> None:
    """Submit a live TypeSafe review and record its response in the workflow."""
    try:
        _emit_json(submit_typesafe_review(workflow, stage, request))
    except Exception as exc:
        _emit_json({"status": "blocked", "reason": str(exc)})
        raise typer.Exit(code=1) from exc


@workflow_app.command("confirm")
def workflow_confirm(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    stage: int = typer.Option(..., "--stage", min=1, max=6),
    reviewer: str = typer.Option(..., "--reviewer", help="agent, human, or delegated typesafe for stages 3..6"),
    decision: str = typer.Option(..., "--decision", help="approve or reject"),
    message: str = typer.Option(..., "--message", help="Review evidence or the user's explicit response"),
    return_to: int | None = typer.Option(None, "--return-to", min=1, max=6,
                                         help="Required when rejecting; stage to revise"),
    evidence: Path | None = typer.Option(None, "--evidence", help="Required TypeSafe review record when reviewer is typesafe"),
) -> None:
    """Record a review receipt for the exact current stage artifact."""
    try:
        result = record_decision(workflow, stage, reviewer, decision, message,
                                 return_to=return_to, evidence_path=evidence)
        _emit_json({"status": "pass", "receipt": result, "workflow_status": workflow_status(workflow)})
    except Exception as exc:
        _emit_json({"status": "blocked", "reason": str(exc)})
        raise typer.Exit(code=1) from exc


@workflow_app.command("reject")
def workflow_reject(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    stage: int = typer.Option(..., "--stage", min=1, max=6),
    reviewer: str = typer.Option(..., "--reviewer", help="agent, human, or delegated typesafe for stages 3..6"),
    return_to: int = typer.Option(..., "--return-to", min=1, max=6),
    message: str = typer.Option(..., "--message", help="Evidence-backed reason or explicit user response"),
    evidence: Path | None = typer.Option(None, "--evidence", help="Required TypeSafe review record when reviewer is typesafe"),
) -> None:
    """Reject the current artifact and record the required revision stage."""
    try:
        result = record_decision(workflow, stage, reviewer, "reject", message,
                                 return_to=return_to, evidence_path=evidence)
        _emit_json({"status": "pass", "receipt": result, "workflow_status": workflow_status(workflow)})
    except Exception as exc:
        _emit_json({"status": "blocked", "reason": str(exc)})
        raise typer.Exit(code=1) from exc


@step2_app.command("validate")
def step2_validate(
    index: Path = typer.Option(..., "--index", help="Saved Step 2 index.json to validate"),
    step1_root: Path = typer.Option(..., "--step1-root", help="Step 1 output root for resolving referenced paths"),
) -> None:
    """Validate a saved Step 2 index and its referenced Step 1 artifacts."""

    result = validate_saved_index(index, step1_root)
    _emit_json(result)
    if result["status"] == "blocked":
        raise typer.Exit(code=1)


@step2_app.command("tool")
def step2_tool(
    name: str = typer.Argument(..., help="One name from step2 tools"),
    step1_root: Path = typer.Option(..., "--step1-root"),
    handoff: Path | None = typer.Option(None, "--handoff"),
    index: Path | None = typer.Option(None, "--index"),
    out: Path = typer.Option(Path("output/step2_index"), "--out", "-o"),
    resume: bool = typer.Option(False, "--resume"),
) -> None:
    """Dispatch exactly one selected Step 2 action."""
    result = execute_step2_tool(name, step1_root=step1_root, handoff_path=handoff,
                               index_path=index, out_dir=out, resume=resume)
    _emit_json(result)
    if result["status"] == "blocked":
        raise typer.Exit(code=1)


@step3_app.command("agent")
def step3_agent() -> None:
    """Print the packaged result-driven Step 3 host-agent contract."""

    typer.echo(files("excel_to_act.steps.step3").joinpath("agent.md").read_text(encoding="utf-8"))


@step3_app.command("tools")
def step3_tools() -> None:
    """Print the Step 3 command catalogue and static-analysis limits."""

    _emit_json(step3_tool_catalog())


@step3_app.command("prepare")
def step3_prepare(
    index: Path = typer.Option(..., "--index", help="Pass-validated Step 2 index.json"),
    step1_root: Path = typer.Option(..., "--step1-root", help="Step 1 artifact root recorded in the index"),
    out: Path = typer.Option(..., "--out", "-o", help="Step 3 analysis directory"),
    source_id: str | None = typer.Option(None, "--source-id", help="Select one eligible source from a multi-source index"),
    scope: Path | None = typer.Option(None, "--scope", help="Optional source-bound analysis_scope.json"),
    workflow: Path | None = typer.Option(None, "--workflow", help="Bind this analysis to a reviewed workflow and require an input boundary before semantic design"),
) -> None:
    """Validate Step 2 and bind one source, optional scope, and optional workflow."""

    result = prepare_analysis(index, step1_root, out, source_id=source_id, scope_path=scope,
                              workflow_path=workflow)
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("input-catalog")
def step3_input_catalog(
    analysis: Path = typer.Option(..., "--analysis", exists=True, file_okay=False,
                                  help="Workflow-bound prepared Step 3 analysis directory"),
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False,
                                  help="Workflow with approved Step 2 and the input-boundary policy"),
    targets: Path = typer.Option(..., "--targets", exists=True, dir_okay=False, readable=True,
                                 help="step3.input_targets.v1 JSON with source, scenario, and ordered targets; optional result_kind, units, and axes"),
    catalog: Path = typer.Option(..., "--catalog", exists=True, dir_okay=False, readable=True,
                                 help="Agent-authored step3.input_boundary.input.v1 JSON catalog"),
) -> None:
    """Validate grouped source inputs and submit the separate Step 3 boundary checkpoint."""
    result = create_input_boundary(analysis, workflow, targets, catalog)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step3_app.command("fields")
def step3_fields(analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory")) -> None:
    """Partition retained stored cells into conservative primary fields."""

    result = build_fields(analysis)
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("dependencies")
def step3_dependencies(analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory")) -> None:
    """Resolve supported references to field dependencies and boundaries."""

    result = build_dependencies(analysis)
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("profile")
def step3_profile(
    analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory"),
    trace: Path | None = typer.Option(None, "--trace", exists=True, dir_okay=False, readable=True,
                                      help="Source-bound active_trace.json from Step 4 discovery"),
    source_trace: Path | None = typer.Option(None, "--source-trace", exists=True, dir_okay=False, readable=True,
                                             help="Step 3 gp.source_candidate_trace.v1 JSON after input-catalog confirmation"),
    out: Path = typer.Option(..., "--out", help="Independent output directory for the reusable syntax profile"),
) -> None:
    """Group source formulas by translated syntax without evaluating them."""
    if (trace is None) == (source_trace is None):
        _emit_json({"tool": "step3.profile", "status": "blocked",
                    "diagnostics": [{"code": "profile_trace_choice_invalid", "severity": "error",
                                     "message": "provide exactly one of --trace or --source-trace"}]})
        raise typer.Exit(code=1)
    result = (profile_source_candidates(analysis, source_trace, out) if source_trace is not None
              else profile_source_families(analysis, trace, out))
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("source-trace")
def step3_source_trace(
    analysis: Path = typer.Option(..., "--analysis", exists=True, file_okay=False,
                                  help="Prepared analysis with current fields, dependencies, and structural plan"),
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False,
                                  help="Workflow with a confirmed Step 3 input-boundary catalog"),
    model_scope: Path | None = typer.Option(None, "--model-scope", exists=True, dir_okay=False, readable=True,
                                            help="Optional Agent-proposed source-coordinate fence JSON"),
    out: Path = typer.Option(..., "--out", help="Independent output directory for candidate trace JSON and Markdown"),
) -> None:
    """Trace static formula-reference candidates without Step 4 history or formula values."""
    result = build_source_candidate_trace(analysis, workflow, out, model_scope)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step3_app.command("plan")
def step3_plan(
    analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory"),
    semantic_map: Path | None = typer.Option(None, "--semantic-map", help="Optional source-bound GP semantic map input JSON"),
) -> None:
    """Build structural blocks and optionally bind a complete static semantic map."""

    result = build_plan(analysis)
    if result["status"] == "pass" and semantic_map is not None:
        result["semantic_mapping"] = build_semantic_plan(analysis, semantic_map)
        if result["semantic_mapping"]["status"] != "pass":
            result["status"] = "blocked"
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("query")
def step3_query_command(
    analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory"),
    target: str = typer.Option(..., "--target", help="Exact name, finite address/range, field ID, or vba:MODULE"),
    sheet: str | None = typer.Option(None, "--sheet", help="Worksheet context for unqualified addresses or local names"),
    offset: int = typer.Option(0, "--offset", help="Zero-based stable page offset"),
    limit: int = typer.Option(100, "--limit", help="Page size (maximum 500 stored records)"),
    out: Path | None = typer.Option(None, "--out", help="Write paired query.json and query.md to this directory"),
) -> None:
    """Select a bounded page of checked source facts and field summaries."""
    result = step3_query(analysis, target, sheet=sheet, offset=offset, limit=limit, out_dir=out)
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("trace")
def step3_trace_command(
    analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory"),
    target: str = typer.Option(..., "--target", help="Exact name, finite address/range, field ID, or formula field selector"),
    sheet: str | None = typer.Option(None, "--sheet", help="Worksheet context for unqualified addresses or local names"),
    direction: str = typer.Option("upstream", "--direction", help="Traverse prerequisites (upstream) or consumers (downstream)"),
    max_depth: int = typer.Option(4, "--max-depth", help="Maximum dependency hops (limit 50)"),
    max_fields: int = typer.Option(100, "--max-fields", help="Maximum visited fields (limit 500)"),
    out: Path | None = typer.Option(None, "--out", help="Write paired trace.json and trace.md to this directory"),
) -> None:
    """Trace checked dependencies with explicit unknown and truncation boundaries."""
    result = step3_trace(analysis, target, sheet=sheet, direction=direction, max_depth=max_depth, max_fields=max_fields, out_dir=out)
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("check")
def step3_check(
    analysis: Path = typer.Option(..., "--analysis", help="Prepared Step 3 analysis directory"),
    spec: Path | None = typer.Option(None, "--spec", help="Optional Agent-authored step3.model_spec.input.v1 JSON draft"),
    semantic_plan: Path | None = typer.Option(None, "--semantic-plan", help="Check this analysis directory's current semantic_plan.json"),
) -> None:
    """Check static source citations, cell coverage, dependencies, and draft block order."""

    semantic_result = None
    if semantic_plan is not None:
        semantic_result = validate_semantic_plan(analysis, semantic_plan)
        if semantic_result["status"] != "pass":
            _emit_json(semantic_result)
            raise typer.Exit(code=1)
    result = validate_analysis(analysis, spec_path=spec)
    if semantic_result is not None:
        result["semantic_check"] = semantic_result
    _emit_json(result)
    if result["status"] != "pass":
        raise typer.Exit(code=1)


@step3_app.command("report")
def step3_report(
    analysis: Path = typer.Option(..., "--analysis", exists=True, file_okay=False,
                                  help="Completed static Step 3 analysis directory"),
    design: Path = typer.Option(..., "--design", exists=True, dir_okay=False, readable=True,
                                help="Agent-authored step3.analysis_design.input.v1 JSON"),
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False,
                                  help="Workflow directory with approved Step 2 checkpoint"),
) -> None:
    """Create a reviewable source-bound analysis/design report; performs no formula calculation."""
    result = create_design_report(analysis, design, workflow)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step4_app.command("agent")
def step4_agent() -> None:
    """Print the packaged Step 4 generation contract."""
    typer.echo(files("excel_to_act.steps.step4").joinpath("agent.md").read_text(encoding="utf-8"))


@step4_app.command("tools")
def step4_tools() -> None:
    """Print Step 4 generation actions and limits."""
    _emit_json(step4_tool_catalog())


@step4_app.command("generate")
def step4_generate(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    trace: Path | None = typer.Option(None, "--trace", exists=True, dir_okay=False, readable=True,
                                      help="Normally omit; if supplied, must be the active Stage 4 discovery trace bound by the current implementation preflight"),
) -> None:
    """Generate a modular Python bundle after the current semantic and implementation preflights."""
    result = generate_model(workflow, trace)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step4_app.command("plan")
def step4_plan(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    implementation: Path = typer.Option(..., "--implementation", exists=True, dir_okay=False,
                                        readable=True, help="Source-bound step4.active_implementation.v1 JSON"),
) -> None:
    """Preflight the approved modular mapping and schedules without creating a code bundle."""
    result = create_implementation_plan(workflow, implementation)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step4_app.command("discover")
def step4_discover(workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False)) -> None:
    """Discover the approved target closure from formula text and raw source values; do not call Excel."""
    result = discover_active_trace(workflow)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step4_app.command("capture-external")
def step4_capture_external(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
) -> None:
    """Capture only formula-derived input vectors approved by the current Step 3 design."""
    result = capture_external_inputs(workflow)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step5_app.command("agent")
def step5_agent() -> None:
    """Print the packaged Step 5 validation and reconciliation contract."""
    typer.echo(files("excel_to_act.steps.step5").joinpath("agent.md").read_text(encoding="utf-8"))


@step5_app.command("tools")
def step5_tools() -> None:
    """Print Step 5 validation, oracle, and reconciliation actions."""
    _emit_json(step5_tool_catalog())


@step5_app.command("validate")
def step5_validate(workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False)) -> None:
    """Compile and run generated Python in isolation without the source workbook."""
    result = validate_generated(workflow)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step5_app.command("oracle")
def step5_oracle(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    target: list[str] = typer.Option([], "--target", help="Defined name to capture; may be repeated"),
    cell_range: list[str] = typer.Option([], "--range", help="Finite qualified range to capture; may be repeated"),
) -> None:
    """Capture native Excel results on a private macro-disabled copy; does not advance Step 5."""
    result = capture_oracle(workflow, targets=target or None, ranges=cell_range or None)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step5_app.command("reconcile")
def step5_reconcile(
    workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False),
    validation: Path = typer.Option(..., "--validation", exists=True, dir_okay=False, readable=True),
    oracle: Path = typer.Option(..., "--oracle", exists=True, dir_okay=False, readable=True),
    abs_tol: float = typer.Option(1e-12, "--abs-tol", min=0),
    rel_tol: float = typer.Option(1e-12, "--rel-tol", min=0),
) -> None:
    """Compare the standalone model with a current native Excel oracle and write the Stage 5 report."""
    result = reconcile(workflow, validation, oracle, abs_tol=abs_tol, rel_tol=rel_tol)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@step6_app.command("agent")
def step6_agent() -> None:
    """Print the packaged Step 6 final-report contract."""
    typer.echo(files("excel_to_act.steps.step6").joinpath("agent.md").read_text(encoding="utf-8"))


@step6_app.command("skill")
def step6_skill() -> None:
    """Print the packaged final-report skill instructions."""
    skill = files("excel_to_act.steps.step6").joinpath("excel-to-act-step6", "SKILL.md")
    typer.echo(skill.read_text(encoding="utf-8"))


@step6_app.command("template")
def step6_template() -> None:
    """Print the packaged human-readable report scaffold."""
    template = files("excel_to_act.steps.step6").joinpath(
        "excel-to-act-step6", "assets", "conversion_report.md"
    )
    typer.echo(template.read_text(encoding="utf-8"))


@step6_app.command("tools")
def step6_tools() -> None:
    """Print Step 6 report action and limits."""
    _emit_json(step6_tool_catalog())


@step6_app.command("report")
def step6_report(workflow: Path = typer.Option(..., "--workflow", exists=True, file_okay=False)) -> None:
    """Create the final report after Stages 1–5 have been confirmed."""
    result = create_conversion_report(workflow)
    _emit_json(result)
    if result.get("status") != "pass":
        raise typer.Exit(code=1)


@views_app.command("compile")
def views_compile(
    handoff: Path = typer.Option(..., "--handoff", help="Step 1 source or batch handoff JSON"),
    step1_root: Path = typer.Option(..., "--step1-root", help="Step 1 output root"),
    out: Path = typer.Option(Path("views.json"), "--out", "-o", help="Deterministic JSON output"),
    budget: int = typer.Option(1200, "--budget", min=1, help="Estimated token budget per chunk"),
) -> None:
    """Compile validated Step 1 artifacts into deterministic sheet views."""
    data = serialize_views(compile_views(handoff, step1_root, budget=budget))
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    typer.echo(f"Wrote views to {out}")


@views_app.command("validate")
def views_validate(
    views_file: Path = typer.Option(..., "--views", exists=True, dir_okay=False),
    output: Path = typer.Option(..., "--output", exists=True, dir_okay=False),
) -> None:
    """Validate claims against a legacy view list or a source-bound evidence packet."""
    try:
        view_data = json.loads(views_file.read_text(encoding="utf-8"))
        agent_output = json.loads(output.read_text(encoding="utf-8"))
        if isinstance(view_data, list):
            views = [WorkbookView.model_validate(item) for item in view_data]
            result = validate_agent_output(agent_output, views)
        else:
            result = validate_evidence_packet(view_data, agent_output)
    except (OSError, ValueError, TypeError) as exc:
        result = {"valid": False, "diagnostics": [{"code": "validation_input_invalid", "severity": "error", "message": str(exc)}]}
    _emit_json(result)
    if result.get("valid") is not True:
        raise typer.Exit(code=1)


@step1_app.command("tools")
def step1_tools() -> None:
    """Print the machine-readable Step1 tool catalogue."""

    _emit_json(tool_catalog())


@checkbox_app.command("identify")
def checkbox_identify(workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True)) -> None:
    _control_identify("checkbox", workbook)


@checkbox_app.command("convert")
def checkbox_convert(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True),
    out: Path = typer.Option(..., "--out"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    _control_convert("checkbox", workbook, out, dry_run)


@checkbox_app.command("evaluate")
def checkbox_evaluate(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True),
    out: Path = typer.Option(..., "--out"),
) -> None:
    _control_evaluate("checkbox", workbook, out)


@activex_app.command("identify")
def activex_identify(workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True)) -> None:
    _control_identify("activex", workbook)


@activex_app.command("convert")
def activex_convert(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True),
    out: Path = typer.Option(..., "--out"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    _control_convert("activex", workbook, out, dry_run)


@activex_app.command("evaluate")
def activex_evaluate(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True),
    out: Path = typer.Option(..., "--out"),
) -> None:
    _control_evaluate("activex", workbook, out)


@vba_app.command("identify")
def vba_identify(workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True)) -> None:
    _control_identify("vba", workbook)


@vba_app.command("convert")
def vba_convert(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True),
    out: Path = typer.Option(..., "--out"),
    dry_run: bool = typer.Option(False, "--dry-run"),
) -> None:
    _control_convert("vba", workbook, out, dry_run)


@vba_app.command("evaluate")
def vba_evaluate(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True),
    out: Path = typer.Option(..., "--out"),
) -> None:
    _control_evaluate("vba", workbook, out)


@step1_app.command("agent")
def step1_agent() -> None:
    """Print the packaged reusable host-agent definition."""

    typer.echo(agent_definition())


@step1_app.command("convert")
def step1_convert(
    input_dir: Path = typer.Argument(..., exists=True, file_okay=False, readable=True),
    out: Path = typer.Option(Path("step1-output"), "--out", "-o", help="Candidate/final output root"),
    traceability_min: float = typer.Option(1.0, "--traceability-min", help="Minimum logical-object accounting ratio [0,1]"),
    fidelity_min: float = typer.Option(1.0, "--fidelity-min", help="Minimum exact supported-fact fidelity ratio [0,1]"),
    parsed_package_min: float = typer.Option(0.0, "--parsed-package-min", help="Minimum parsed OOXML package-part ratio [0,1]"),
    opaque_max: float = typer.Option(1.0, "--opaque-max", help="Maximum opaque OOXML package-part ratio [0,1]"),
    allow_opaque: bool = typer.Option(True, "--allow-opaque/--no-allow-opaque", help="Keep preserved opaque parts as a partial handoff"),
) -> None:
    """Convert all workbook candidates in a directory and report one batch envelope."""

    result = convert_directory(
        input_dir,
        out,
        traceability_min=traceability_min,
        fidelity_min=fidelity_min,
        parsed_package_min=parsed_package_min,
        opaque_max=opaque_max,
        allow_opaque=allow_opaque,
    )
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)


@step1_app.command("check")
def step1_check(run: Path = typer.Option(..., "--run", help="Per-source run directory")) -> None:
    """Recompute coverage, fidelity, package preservation, and the handoff."""

    result = execute_tool("step1.check", run)
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)


@step1_app.command("coverage")
def step1_coverage(run: Path = typer.Option(..., "--run", help="Per-source run directory")) -> None:
    """Check independent logical-object and package-part ledgers."""

    result = execute_tool("step1.coverage", run)
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)


@step1_app.command("fidelity")
def step1_fidelity(run: Path = typer.Option(..., "--run", help="Per-source run directory")) -> None:
    """Check raw OOXML fields and normalized supported facts."""

    result = execute_tool("step1.fidelity", run)
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)


@step1_app.command("tool")
def step1_tool(
    name: str = typer.Argument(..., help="Tool name from `step1 tools`"),
    run: Path = typer.Option(..., "--run", help="Per-source run directory"),
) -> None:
    """Invoke one persisted extraction, analysis, storage, or report tool."""

    result = execute_tool(name, run)
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)


@step1_app.command("auto-recover")
def step1_auto_recover(
    run: Path = typer.Option(..., "--run", help="Per-source run directory"),
    max_attempts: int = typer.Option(3, "--max-attempts", min=1, max=3),
) -> None:
    """Run only suggested deterministic recovery tools, stopping on no progress."""

    result = auto_recover(run, max_attempts=max_attempts)
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)


@step1_app.command("finalize")
def step1_finalize(run: Path = typer.Option(..., "--run", help="Per-source run directory")) -> None:
    """Fresh-check a candidate and promote it only when its gate is ready."""

    result = finalize_run(run)
    _emit_json(result)
    if result["status"] in {"fail", "error", "blocked"}:
        raise typer.Exit(code=1)
