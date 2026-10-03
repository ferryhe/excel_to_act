"""Legacy form controls (checkbox / spin / scroll / list / drop) from VML drawings.

Actuarial models use these controls as scenario switches: a checkbox or scroll
bar writes into a *linked cell* that the formulas then read. Missing that link
means losing the model's input surface.

In OOXML the link lives in the sheet's legacy VML drawing, not in a tidy
SpreadsheetML element:

    <x:ClientData ObjectType="Checkbox">
      <x:FmlaLink>Inputs!$B$2</x:FmlaLink>
      <x:Checked>1</x:Checked>
    </x:ClientData>

``xl/ctrlProps/*.xml`` does not carry the linked cell (ActiveX controls keep
their properties in ``xl/activeX/*.bin``, which stays opaque), so VML is the
authoritative source for legacy controls.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from excel_to_act.ingest.data_table import resolve_package_part, sheet_xml_parts

_VML_MARKER = "vmlDrawing"
# `Note` is a cell comment and `Pict` is a picture/embedded object; neither is a control.
_NON_CONTROL_TYPES = {"Note", "Pict"}
# A ClientData block without any of these is not bound to anything we can use.
_LINK_PROPERTIES = ("FmlaLink", "FmlaRange", "FmlaMacro")


@dataclass(frozen=True)
class FormControl:
    """One legacy form control bound to a worksheet cell."""

    control_id: str | None = None
    control_type: str | None = None
    linked_cell: str | None = None  # full reference, e.g. "Inputs!B2"
    linked_sheet: str | None = None
    linked_address: str | None = None  # bare A1, e.g. "B2"
    list_fill_range: str | None = None
    macro: str | None = None
    properties: dict[str, str] = field(default_factory=dict)


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _normalize_ref(value: str | None) -> str | None:
    """Turn ``Inputs!$B$2`` into ``Inputs!B2``."""

    if not value:
        return None
    return value.strip().replace("$", "").replace("'", "")


def _split_ref(ref: str | None) -> tuple[str | None, str | None]:
    if not ref:
        return None, None
    if "!" in ref:
        sheet, _, address = ref.rpartition("!")
        return sheet, address
    return None, ref


def sheet_vml_parts(workbook_path: Path) -> dict[str, list[str]]:
    """Return ``{sheet_name: [vml part names]}``; a sheet may carry several drawings."""

    workbook_path = workbook_path.expanduser().resolve()
    result: dict[str, list[str]] = {}
    with zipfile.ZipFile(workbook_path) as zf:
        names = set(zf.namelist())
        for sheet_name, part in sheet_xml_parts(workbook_path).items():
            rels_part = str(PurePosixPath(part).parent / "_rels" / f"{PurePosixPath(part).name}.rels")
            if rels_part not in names:
                continue
            try:
                rels = ET.fromstring(zf.read(rels_part))
            except ET.ParseError:
                continue
            for rel in rels:
                if rel.attrib.get("TargetMode") == "External":
                    continue
                target = rel.attrib.get("Target", "")
                is_vml = _VML_MARKER in target or rel.attrib.get("Type", "").endswith(f"/{_VML_MARKER}")
                if not is_vml:
                    continue
                resolved = resolve_package_part(str(PurePosixPath(part).parent), target)
                if resolved in names:
                    result.setdefault(sheet_name, []).append(resolved)
    return result


def read_form_controls(workbook_path: Path) -> dict[str, list[FormControl]]:
    """Return ``{sheet_name: [FormControl]}`` for every legacy control found.

    Sheets without a VML drawing are absent from the result. Malformed or
    missing parts yield no controls instead of raising, because this is an
    optional enrichment of the inventory.
    """

    workbook_path = workbook_path.expanduser().resolve()
    result: dict[str, list[FormControl]] = {}
    vml_parts = sheet_vml_parts(workbook_path)
    if not vml_parts:
        return result
    with zipfile.ZipFile(workbook_path) as zf:
        for sheet_name, parts in vml_parts.items():
            for part in parts:
                try:
                    root = ET.fromstring(zf.read(part))
                except (KeyError, ET.ParseError):
                    continue
                parents = {child: parent for parent in root.iter() for child in parent}
                for element in root.iter():
                    if _local(element.tag) != "ClientData":
                        continue
                    properties = {_local(child.tag): (child.text or "").strip() for child in element}
                    control_type = element.attrib.get("ObjectType") or properties.get("ObjectType") or None
                    if control_type in _NON_CONTROL_TYPES:
                        continue
                    if not any(properties.get(key) for key in _LINK_PROPERTIES):
                        continue
                    node = element
                    shape_id = None
                    while node is not None:
                        if _local(node.tag) == "shape" and node.attrib.get("id"):
                            shape_id = node.attrib.get("id")
                            break
                        node = parents.get(node)
                    linked_cell = _normalize_ref(properties.get("FmlaLink"))
                    linked_sheet, linked_address = _split_ref(linked_cell)
                    result.setdefault(sheet_name, []).append(
                        FormControl(
                            control_id=shape_id or control_type,
                            control_type=control_type,
                            linked_cell=linked_cell,
                            linked_sheet=linked_sheet,
                            linked_address=linked_address,
                            list_fill_range=_normalize_ref(properties.get("FmlaRange")),
                            macro=properties.get("FmlaMacro") or None,
                            properties=properties,
                        )
                    )
    return result
