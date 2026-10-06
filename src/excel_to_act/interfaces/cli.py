"""CLI entrypoint for Phase 1 workflows."""

from __future__ import annotations

from pathlib import Path
import json
from importlib.resources import files

import typer

from excel_to_act.orchestrator.phase1 import Phase1Orchestrator
from excel_to_act.ingest.control_artifacts import ARTIFACT_FILES, build_single_control_artifact, clear_declared_vba_sources, evaluate_control_output, vba_handoff_markdown
from excel_to_act.schemas import VbaHandoff
from excel_to_act.steps.step2.workflow import build_index, execute_tool as execute_step2_tool, tool_catalog as step2_tool_catalog, validate_saved_index
from excel_to_act.steps.step1.workflow import agent_definition, auto_recover, convert_directory, execute_tool, finalize_run, tool_catalog
from excel_to_act.schemas import WorkbookView
from excel_to_act.views import compile_views, serialize_views, validate_agent_output

app = typer.Typer(help="Excel to actuarial model decomposition toolkit")
step1_app = typer.Typer(help="Human Step1 raw-directory conversion and checked handoff")
checkbox_app = typer.Typer(help="Identify, convert, and evaluate worksheet checkbox bindings")
activex_app = typer.Typer(help="Identify, convert, and evaluate ActiveX event declarations")
vba_app = typer.Typer(help="Identify, export, and evaluate VBA source handoff")
step2_app = typer.Typer(help="Index Step 1 handoffs for downstream agents")
views_app = typer.Typer(help="Compile deterministic source-addressable workbook views")
app.add_typer(step1_app, name="step1")
step1_app.add_typer(checkbox_app, name="checkbox")
step1_app.add_typer(activex_app, name="activex")
step1_app.add_typer(vba_app, name="vba")
app.add_typer(step2_app, name="step2")
app.add_typer(views_app, name="views")


@app.callback()
def main() -> None:
    """Excel to actuarial model decomposition toolkit."""


@app.command()
def inspect(
    workbook: Path = typer.Argument(..., exists=True, dir_okay=False, readable=True, help=".xlsx/.xlsm workbook to inspect"),
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
        typer.echo("Completeness check failed; see completeness.json for blocking gaps.", err=True)
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


@step2_app.command("tools")
def step2_tools() -> None:
    """Print the initial machine-readable Step 2 action catalogue."""

    _emit_json(step2_tool_catalog())


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
    """Validate a JSON Agent response against compiled views."""
    views = [WorkbookView.model_validate(item) for item in json.loads(views_file.read_text(encoding="utf-8"))]
    result = validate_agent_output(json.loads(output.read_text(encoding="utf-8")), views)
    _emit_json(result)


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
