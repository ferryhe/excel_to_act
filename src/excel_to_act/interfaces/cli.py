"""CLI entrypoint for Phase 1 workflows."""

from __future__ import annotations

from pathlib import Path
import json

import typer

from excel_to_act.orchestrator.phase1 import Phase1Orchestrator
from excel_to_act.steps.step2.workflow import build_index, tool_catalog as step2_tool_catalog
from excel_to_act.steps.step1.workflow import agent_definition, auto_recover, convert_directory, execute_tool, finalize_run, tool_catalog

app = typer.Typer(help="Excel to actuarial model decomposition toolkit")
step1_app = typer.Typer(help="Human Step1 raw-directory conversion and checked handoff")
step2_app = typer.Typer(help="Index Step 1 handoffs for downstream agents")
app.add_typer(step1_app, name="step1")
app.add_typer(step2_app, name="step2")


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


@step2_app.command("tools")
def step2_tools() -> None:
    """Print the initial machine-readable Step 2 action catalogue."""

    _emit_json(step2_tool_catalog())


@step2_app.command("index")
def step2_index(
    handoff: Path = typer.Option(..., "--handoff", help="Step 1 source or batch handoff JSON"),
    step1_root: Path = typer.Option(..., "--step1-root", help="Step 1 output root for resolving handoff and artifact paths"),
    out: Path = typer.Option(Path("output/step2_index"), "--out", "-o", help="Step 2 index output directory"),
) -> None:
    """Build a deterministic index from a current Step 1 handoff."""

    result = build_index(handoff, step1_root, out)
    _emit_json(result)
    if result["status"] == "blocked":
        raise typer.Exit(code=1)


@step1_app.command("tools")
def step1_tools() -> None:
    """Print the machine-readable Step1 tool catalogue."""

    _emit_json(tool_catalog())


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
