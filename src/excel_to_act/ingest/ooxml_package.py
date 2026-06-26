"""OOXML package scanner for opaque/unsupported workbook parts."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from excel_to_act.schemas import PackagePart, SourceLocation, UnsupportedFeature

OPAQUE_MARKERS = {
    "vbaProject.bin": "macro project",
    "/pivot": "pivot table/cache",
    "/charts/": "chart",
    "/drawings/": "drawing",
    "/externalLinks/": "external link",
    "/connections": "connection",
    "/embeddings/": "embedded OLE object",
    "/media/": "media",
}


def _content_types(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(zf.read("[Content_Types].xml"))
    except Exception:
        return {}
    ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        if child.tag == ns + "Default":
            defaults[child.attrib.get("Extension", "")] = child.attrib.get("ContentType", "")
        elif child.tag == ns + "Override":
            overrides[child.attrib.get("PartName", "").lstrip("/")] = child.attrib.get("ContentType", "")
    result = dict(overrides)
    for info in zf.infolist():
        if info.filename not in result:
            result[info.filename] = defaults.get(Path(info.filename).suffix.lstrip("""."""), "")
    return result


def scan_ooxml_package(workbook_path: Path) -> tuple[list[PackagePart], list[UnsupportedFeature]]:
    """Return OOXML package parts and opaque diagnostics."""

    if workbook_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        loc = SourceLocation(workbook_path=str(workbook_path), object_type="workbook")
        return [], [UnsupportedFeature(feature_type="file_type", description=f"Unsupported file type: {workbook_path.suffix}", source_location=loc, severity="error")]
    parts: list[PackagePart] = []
    unsupported: list[UnsupportedFeature] = []
    with zipfile.ZipFile(workbook_path) as zf:
        content_types = _content_types(zf)
        for info in zf.infolist():
            name = info.filename
            marker = next((desc for token, desc in OPAQUE_MARKERS.items() if token in name), None)
            loc = SourceLocation(workbook_path=str(workbook_path), ooxml_part=name, object_type="ooxml_part", object_id=name)
            parts.append(PackagePart(name=name, content_type=content_types.get(name), size=info.file_size, opaque=marker is not None, source_location=loc))
            if marker is not None:
                unsupported.append(UnsupportedFeature(feature_type=marker, description=f"Opaque OOXML part recorded: {name}", source_location=loc, metadata={"part_name": name, "bytes": info.file_size}))
    return parts, unsupported
