"""openpyxl-backed workbook manifest reader."""

from __future__ import annotations

import hashlib
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from excel_to_act.ingest.ooxml_package import scan_ooxml_package
from excel_to_act.schemas import SheetManifest, SourceLocation, UnsupportedFeature, UnsupportedSeverity, WorkbookManifest
from excel_to_act.steps.step1.source_scan import SourceScanError, object_identity, scan_step1_source

WORKBOOK_READ_ERRORS = (
    OSError,
    EOFError,
    zipfile.BadZipFile,
    zipfile.LargeZipFile,
    zlib.error,
    RuntimeError,
    KeyError,
    ET.ParseError,
    InvalidFileException,
    SourceScanError,
    ValueError,
)


def workbook_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


class OpenpyxlWorkbookReader:
    name = "openpyxl"

    def read_manifest(self, workbook_path: Path) -> WorkbookManifest:
        workbook_path = workbook_path.expanduser().resolve()
        wb = None
        try:
            file_size = workbook_path.stat().st_size
            sha256 = workbook_sha256(workbook_path)
            if workbook_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                _, unsupported = scan_ooxml_package(workbook_path)
                return WorkbookManifest(
                    workbook_path=str(workbook_path),
                    file_name=workbook_path.name,
                    file_size=file_size,
                    sha256=sha256,
                    sheets=[],
                    unsupported_features=unsupported,
                )

            with workbook_path.open("rb") as fh:
                signature = fh.read(8)
            if signature == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                raise InvalidFileException(
                    "Input uses an OLE/CFB container, which may indicate a password-protected "
                    "Office workbook. Decrypt it and resave an unprotected .xlsx/.xlsm before inspecting."
                )

            # The Step 1 scanner already identifies unreadable required OOXML
            # parts; reuse it so inspect cannot turn a broken package into an
            # apparently empty workbook.
            source_scan = scan_step1_source(workbook_path)
            parts, unsupported = scan_ooxml_package(workbook_path)
            for part in parts:
                source_part = source_scan["parts"][part.name]
                part.opaque = source_part["opaque"]
                part.opaque_reason = source_part["opaque_reason"]
                part.source_location.source_identity = object_identity("package_part", part.name)
            for feature in unsupported:
                name = feature.source_location.ooxml_part
                if name in source_scan["parts"]:
                    feature.opaque = source_scan["parts"][name]["opaque"]
            try:
                wb = load_workbook(workbook_path, data_only=False, read_only=False, keep_vba=workbook_path.suffix.lower() == ".xlsm")
            except Exception as exc:
                raise InvalidFileException(
                    f"openpyxl could not read workbook: {type(exc).__name__}: {exc}"
                ) from exc
            sheets = [
                SheetManifest(
                    name=ws.title,
                    index=i,
                    state=ws.sheet_state,
                    max_row=ws.max_row or 0,
                    max_column=ws.max_column or 0,
                )
                for i, ws in enumerate(wb.worksheets)
            ]
            defined_names = getattr(wb, "defined_names", None)
            named_count = len(list(getattr(defined_names, "items", lambda: [])())) if defined_names is not None else 0
            named_count += sum(
                len(list(getattr(getattr(ws, "defined_names", None), "items", lambda: [])()))
                for ws in wb.worksheets
            )
            calc = getattr(wb, "calculation", None)
            calc_mode = getattr(calc, "calcMode", None) or getattr(calc, "mode", None)
            return WorkbookManifest(
                workbook_path=str(workbook_path),
                file_name=workbook_path.name,
                file_size=workbook_path.stat().st_size,
                sha256=workbook_sha256(workbook_path),
                is_macro_enabled=workbook_path.suffix.lower() == ".xlsm" or any("vbaProject.bin" in p.name for p in parts),
                sheets=sheets,
                named_ranges_count=named_count,
                calc_mode=calc_mode,
                package_parts=parts,
                unsupported_features=unsupported,
            )
        except WORKBOOK_READ_ERRORS as exc:
            try:
                file_size = workbook_path.stat().st_size
            except OSError:
                file_size = 0
            try:
                sha256 = workbook_sha256(workbook_path)
            except OSError:
                sha256 = ""
            location = SourceLocation(workbook_path=str(workbook_path), object_type="workbook")
            return WorkbookManifest(
                workbook_path=str(workbook_path),
                file_name=workbook_path.name,
                file_size=file_size,
                sha256=sha256,
                unsupported_features=[
                    UnsupportedFeature(
                        feature_type="input_read_error",
                        description=f"Could not read workbook: {type(exc).__name__}: {exc}",
                        source_location=location,
                        severity=UnsupportedSeverity.error,
                        opaque=False,
                    )
                ],
            )
        finally:
            if wb is not None:
                wb.close()
