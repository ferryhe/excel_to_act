from typer.testing import CliRunner

from excel_to_act.interfaces.cli import app


def test_cli_help() -> None:
    result = CliRunner().invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "Inspect a workbook and produce Phase 1 artifacts" in result.output
