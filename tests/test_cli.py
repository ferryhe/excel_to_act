from typer.testing import CliRunner
from typer.main import get_command

from excel_to_act.interfaces.cli import app


def test_cli_help() -> None:
    result = CliRunner().invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "Inspect a workbook and produce Phase 1 artifacts" in result.output


def test_step4_external_capture_help_describes_catalog_derived_inputs() -> None:
    command = get_command(app).commands["step4"].commands["capture-external"]
    assert "formula-derived input vectors approved by the current Step 3 design" in command.help
    assert "five" not in command.help.lower()
