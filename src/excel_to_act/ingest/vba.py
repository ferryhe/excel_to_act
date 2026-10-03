"""VBA macro project extraction.

VBA lives in ``xl/vbaProject.bin`` — an embedded OLE compound document. openpyxl's
``keep_vba=True`` only preserves that binary blob; the human-readable source is
read with `oletools` (olevba, BSD-3). oletools is an **optional** dependency so
the core never breaks when it is absent — callers get ``oletools_missing=True``
instead of an import error.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from excel_to_act.schemas import VbaModule

_VBA_PART = "vbaProject.bin"
_KIND_BY_EXT = {".bas": "StdModule", ".cls": "ClassModule", ".frm": "UserForm", ".frx": "UserForm"}
_PROC_RE = re.compile(r"(?:Public |Private |Friend )?(?:Sub|Function)\s+(\w+)")


@dataclass(frozen=True)
class VbaProject:
    """Result of reading a workbook's VBA project.

    ``available`` is True whenever the workbook physically contains a
    ``vbaProject.bin``, regardless of whether oletools could read it.
    """

    available: bool = False
    oletools_missing: bool = False
    modules: list[VbaModule] = field(default_factory=list)
    error: str | None = None


def _has_vba(workbook_path: Path) -> bool:
    try:
        with zipfile.ZipFile(workbook_path) as zf:
            return any(info.filename.endswith(_VBA_PART) for info in zf.infolist())
    except (zipfile.BadZipFile, OSError):
        return False


def extract_vba_project(workbook_path: Path) -> VbaProject:
    """Extract VBA module source and procedure names, or report why not."""

    if not _has_vba(workbook_path):
        return VbaProject(available=False)
    try:
        from oletools.olevba import VBA_Parser  # type: ignore
    except ImportError:
        return VbaProject(available=True, oletools_missing=True)

    modules: list[VbaModule] = []
    try:
        parser = VBA_Parser(str(workbook_path))
        try:
            for _filename, _stream_path, vba_filename, vba_code in parser.extract_macros():
                if not vba_filename:
                    continue
                stem, kind = vba_filename, "StdModule"
                for ext, candidate in _KIND_BY_EXT.items():
                    if stem.lower().endswith(ext):
                        kind = candidate
                        stem = stem[: -len(ext)]
                        break
                code = vba_code or ""
                modules.append(
                    VbaModule(
                        name=stem,
                        kind=kind,
                        code=code,
                        procedures=_PROC_RE.findall(code),
                    )
                )
        finally:
            parser.close()
    except Exception as exc:  # olevba can raise on exotic or corrupt streams
        return VbaProject(available=True, error=str(exc))
    return VbaProject(available=True, modules=modules)
