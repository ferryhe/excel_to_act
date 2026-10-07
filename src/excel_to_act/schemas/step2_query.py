"""Contracts for source-bound Step 2 query evidence packets."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from excel_to_act.schemas.artifacts import WorkbookView
from excel_to_act.schemas.step2_prepare import ReadingFileRef


class EvidenceManifestRef(BaseModel):
    model_config = ConfigDict(extra="forbid")

    path: str
    reading_root: str
    sha256: str
    bytes: int = Field(ge=0)


class EvidenceSelector(BaseModel):
    model_config = ConfigDict(extra="forbid")

    kind: Literal["overview", "sheet", "cell", "range", "name", "control", "vba", "feature"]
    target: str | None = None
    sheet: str | None = None
    range: str | None = None


class EvidenceSource(BaseModel):
    model_config = ConfigDict(extra="forbid")

    source_id: str
    source_path: str | None = None
    source_sha256: str | None = None
    run_id: str | None = None
    status: str
    revision_id: str
    scope_sha256: str | None = None
    retained_sheets: list[str] = Field(default_factory=list)
    excluded_sheets: list[str] = Field(default_factory=list)
    allowed_dependency_ranges: list[dict[str, str]] = Field(default_factory=list)


class EvidencePacket(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["step2.evidence_packet.v1"] = "step2.evidence_packet.v1"
    artifact_type: Literal["step2_evidence_packet"] = "step2_evidence_packet"
    status: str
    manifest: EvidenceManifestRef
    source: EvidenceSource
    selector: EvidenceSelector
    canonical_view_refs: list[ReadingFileRef] = Field(default_factory=list)
    source_file_refs: list[ReadingFileRef] = Field(default_factory=list)
    views: list[WorkbookView] = Field(default_factory=list)
    delivered_record_ids: dict[str, list[str]] = Field(default_factory=dict)
    summary: dict[str, Any] = Field(default_factory=dict)
    pagination: dict[str, Any] = Field(default_factory=dict)
    metrics: dict[str, Any] = Field(default_factory=dict)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    trace: dict[str, Any] | None = None
