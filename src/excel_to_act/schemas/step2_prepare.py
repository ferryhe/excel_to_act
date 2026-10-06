"""Runtime contracts for the Step 2 prepared reading package."""

from __future__ import annotations

from pathlib import PurePosixPath
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from excel_to_act.schemas.artifacts import CellInventory, SourceLocation


class ReadingFileRef(BaseModel):
    root: Literal["reading", "step1", "step2_index"]
    path: str
    sha256: str
    bytes: int | None = Field(default=None, ge=0)
    view_id: str | None = None
    scope: Literal["workbook", "sheet", "region"] | None = None
    sheet_name: str | None = None

    @field_validator("path")
    @classmethod
    def root_relative_path(cls, value: str) -> str:
        path = PurePosixPath(value)
        if not value or "\\" in value or path.is_absolute() or ".." in path.parts:
            raise ValueError("reading references must use root-relative POSIX paths")
        return value


class ReadingSource(BaseModel):
    source_id: str
    source_path: str | None = None
    source_sha256: str | None = None
    run_id: str | None = None
    status: str
    ready_for_next_step: bool = False
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    artifacts: dict[str, ReadingFileRef] = Field(default_factory=dict)
    views: list[ReadingFileRef] = Field(default_factory=list)
    dependency_graph: ReadingFileRef | None = None
    retained_sheets: list[str] = Field(default_factory=list)
    excluded_sheets: list[str] = Field(default_factory=list)
    allowed_dependency_ranges: list[dict[str, str]] = Field(default_factory=list)
    allowed_read_paths: list[ReadingFileRef] = Field(default_factory=list)
    scope_sha256: str | None = None
    effective_scope: dict[str, Any] = Field(default_factory=dict)
    lookup: "ReadingLookup" = Field(default_factory=lambda: ReadingLookup())


class DefinedNameLookup(BaseModel):
    name: str
    scope: str
    node_id: str | None = None
    address: str
    source_location: SourceLocation


class ReadingLookup(BaseModel):
    workbook_view: ReadingFileRef | None = None
    sheet_views: dict[str, ReadingFileRef] = Field(default_factory=dict)
    region_views: list[ReadingFileRef] = Field(default_factory=list)
    defined_names: list[DefinedNameLookup] = Field(default_factory=list)


class Step2ReadingManifest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    schema_version: Literal["step2.reading.v1"] = "step2.reading.v1"
    artifact_type: Literal["reading_manifest"] = "reading_manifest"
    compiler_version: str
    revision_id: str
    input_fingerprint: str
    options: dict[str, Any]
    roots: dict[str, str]
    source_count: int = Field(ge=0)
    index: ReadingFileRef
    scope: ReadingFileRef
    sources: list[ReadingSource]
    outputs: dict[str, str]


class DependencyAuditSource(BaseModel):
    source_id: str
    source_sha256: str
    run_id: str
    scope_sha256: str
    supported_static_inbound_reference_count: int = Field(ge=0)
    retained_dependency_range_count: int = Field(ge=0)
    unresolved_static_reference_count: int = Field(ge=0)
    ranges: list[dict[str, str]] = Field(default_factory=list)
    references: list[dict[str, Any]] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)


class DependencySnapshotSource(BaseModel):
    source_sha256: str
    run_id: str
    scope_sha256: str
    ranges: list[dict[str, str]] = Field(default_factory=list)
    cells: list[CellInventory] = Field(default_factory=list)


class Step2DependencyAudit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    schema_version: Literal["step2.dependency_audit.v1"] = "step2.dependency_audit.v1"
    artifact_type: Literal["dependency_audit"] = "dependency_audit"
    sources: list[DependencyAuditSource] = Field(default_factory=list)
    ignored_legacy_sidecars: list[str] = Field(default_factory=list)


class Step2DependencySnapshot(BaseModel):
    model_config = ConfigDict(extra="forbid")
    schema_version: Literal["step2.dependency_snapshot.v1"] = "step2.dependency_snapshot.v1"
    artifact_type: Literal["retained_dependency_cells"] = "retained_dependency_cells"
    sources: list[DependencySnapshotSource] = Field(default_factory=list)
