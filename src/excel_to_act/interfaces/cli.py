"""CLI entrypoint for Phase 1 workflows."""

from __future__ import annotations

from pathlib import Path

import typer

from excel_to_act.orchestrator.phase1 import Phase1Orchestrator

app = typer.Typer(help="Excel to actuarial model decomposition toolkit")


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
