"""Form controls: linked cell read from the sheet's legacy VML drawing."""

from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path

from openpyxl import Workbook

from excel_to_act.ingest.form_controls import read_form_controls, sheet_vml_parts
from excel_to_act.ingest.openpyxl_reader import OpenpyxlWorkbookReader
from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor

_VML = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:x="urn:schemas-microsoft-com:office:excel" '
    'xmlns:o="urn:schemas-microsoft-com:office:office">'
    '<v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute">'
    '<x:ClientData ObjectType="Checkbox">'
    "<x:Anchor>1, 2, 3, 4</x:Anchor>"
    "<x:FmlaLink>Inputs!$B$2</x:FmlaLink>"
    "<x:Checked>1</x:Checked>"
    "</x:ClientData>"
    "</v:shape>"
    '<v:shape id="_x0000_s1026" type="#_x0000_t201" style="position:absolute">'
    '<x:ClientData ObjectType="Note">'
    "<x:FmlaLink>A1</x:FmlaLink>"
    "</x:ClientData>"
    "</v:shape>"
    "</xml>"
)


def build_fixture(path: Path) -> Path:
    wb = Workbook()
    inputs = wb.active
    inputs.title = "Inputs"
    inputs["B2"] = 0.05
    controls = wb.create_sheet("Controls")
    controls["A1"] = "switch"
    wb.save(path)
    wb.close()
    return path


def inject_vml_drawing(path: Path, sheet_index: int, vml_xml: str) -> None:
    """Attach a legacy VML drawing (controls live here) to a sheet."""
    rels_part = str(posixpath.join("xl/worksheets/_rels", f"sheet{sheet_index + 1}.xml.rels"))
    vml_part = "xl/drawings/vmlDrawing1.vml"
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        payload = {name: zf.read(name) for name in names}
    payload[rels_part] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdVml" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
        'Target="../drawings/vmlDrawing1.vml"/></Relationships>'
    ).encode("utf-8")
    payload[vml_part] = vml_xml.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in payload.items():
            zf.writestr(name, data)


def _inventory(path: Path):
    manifest = OpenpyxlWorkbookReader().read_manifest(path)
    return OpenpyxlInventoryExtractor().extract(path, manifest)


def _controls(inventory):
    sheet = next(s for s in inventory.sheets if s.name == "Controls")
    return [r for r in sheet.ranges if r.kind == "form_control"]


def test_form_control_linked_cell_is_extracted(tmp_path: Path) -> None:
    fixture = build_fixture(tmp_path / "controls.xlsx")
    inject_vml_drawing(fixture, 1, _VML)

    controls = read_form_controls(fixture)["Controls"]
    # The Note-type ClientData must not be reported as a control.
    assert len(controls) == 1
    control = controls[0]
    assert control.control_type == "Checkbox"
    assert control.linked_cell == "Inputs!B2"
    assert control.properties["Checked"] == "1"

    inventory = _inventory(fixture)
    ranges = _controls(inventory)
    assert len(ranges) == 1
    assert ranges[0].metadata["linked_cell"] == "Inputs!B2"
    assert ranges[0].metadata["control_type"] == "Checkbox"


def test_sheet_without_vml_has_no_controls(tmp_path: Path) -> None:
    fixture = build_fixture(tmp_path / "none.xlsx")
    assert sheet_vml_parts(fixture) == {}
    assert read_form_controls(fixture) == {}
    inventory = _inventory(fixture)
    assert _controls(inventory) == []
    assert (
        inventory.coverage.recognized_inventory_objects + inventory.coverage.unsupported_or_opaque_objects
        == inventory.coverage.discovered_workbook_objects
    )
