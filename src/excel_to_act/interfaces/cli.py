"""CLI entrypoint placeholder for Phase 1 workflows."""

import typer

app = typer.Typer(help="Excel to actuarial model decomposition toolkit")


@app.command()
def inspect(workbook: str) -> None:
    """Inspect a workbook and produce Phase 1 artifacts.

    Implementation is planned in PR-02 onward; this command documents the intended public entrypoint.
    """
    typer.echo(f"Phase 1 inspect workflow is not implemented yet: {workbook}")
