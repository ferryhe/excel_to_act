"""Pydantic artifact contracts for Phase 1 Excel decomposition."""

from __future__ import annotations

from datetime import UTC, datetime
from enum import Enum
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

SCHEMA_VERSION = "phase1.v1"


class Artifact(BaseModel):
    """Base model for all serialized Phase 1 artifacts."""

    model_config = ConfigDict(use_enum_values=True)

    schema_version: str = SCHEMA_VERSION


class SourceLocation(BaseModel):
    """Original workbook location for a discovered object."""

    workbook_path: str | None = None
    sheet_name: str | None = None
    sheet_index: int | None = Field(default=None, ge=0)
    address: str | None = None
    object_type: str
    object_id: str | None = None
    ooxml_part: str | None = None

    @model_validator(mode="after")
    def validate_location(self) -> "SourceLocation":
        if not (self.sheet_name or self.ooxml_part or self.workbook_path):
            raise ValueError("source location must include sheet_name, ooxml_part, or workbook_path")
        if self.address and not self.sheet_name:
            raise ValueError("cell/range address requires sheet_name")
        return self


class UnsupportedSeverity(str, Enum):
    info = "info"
    warning = "warning"
    error = "error"


class UnsupportedFeature(BaseModel):
    feature_type: str
    description: str
    source_location: SourceLocation
    severity: UnsupportedSeverity = UnsupportedSeverity.warning
    opaque: bool = True
    metadata: dict[str, Any] = Field(default_factory=dict)


class SheetManifest(BaseModel):
    name: str
    index: int = Field(ge=0)
    state: str = "visible"
    max_row: int = Field(ge=0)
    max_column: int = Field(ge=0)


class PackagePart(BaseModel):
    name: str
    content_type: str | None = None
    relationship_type: str | None = None
    size: int = Field(ge=0)
    opaque: bool = False
    source_location: SourceLocation


class WorkbookManifest(Artifact):
    artifact_type: Literal["workbook_manifest"] = "workbook_manifest"
    workbook_path: str
    file_name: str
    file_size: int = Field(ge=0)
    sha256: str
    is_macro_enabled: bool = False
    sheets: list[SheetManifest] = Field(default_factory=list)
    named_ranges_count: int = Field(ge=0, default=0)
    calc_mode: str | None = None
    package_parts: list[PackagePart] = Field(default_factory=list)
    unsupported_features: list[UnsupportedFeature] = Field(default_factory=list)


class CellKind(str, Enum):
    literal = "value"
    formula = "formula"
    blank = "blank"


class CellInventory(BaseModel):
    source_location: SourceLocation
    address: str
    row: int = Field(ge=1)
    column: int = Field(ge=1)
    kind: CellKind
    value: str | int | float | bool | None = None
    data_type: str | None = None
    number_format: str | None = None
    formula: str | None = None
    style_id: int | None = None


class RangeInventory(BaseModel):
    source_location: SourceLocation
    name: str | None = None
    address: str
    kind: str
    metadata: dict[str, Any] = Field(default_factory=dict)


class SheetInventory(BaseModel):
    source_location: SourceLocation
    name: str
    index: int = Field(ge=0)
    max_row: int = Field(ge=0)
    max_column: int = Field(ge=0)
    state: str = "visible"
    cells: list[CellInventory] = Field(default_factory=list)
    ranges: list[RangeInventory] = Field(default_factory=list)
    layout_objects: list[RangeInventory] = Field(default_factory=list)


class CoverageSummary(BaseModel):
    recognized_inventory_objects: int = Field(ge=0)
    unsupported_or_opaque_objects: int = Field(ge=0)
    discovered_workbook_objects: int = Field(ge=0)

    @model_validator(mode="after")
    def coverage_equation(self) -> "CoverageSummary":
        if self.recognized_inventory_objects + self.unsupported_or_opaque_objects != self.discovered_workbook_objects:
            raise ValueError("coverage equation violated")
        return self


class WorkbookInventory(Artifact):
    artifact_type: Literal["workbook_inventory"] = "workbook_inventory"
    workbook_sha256: str
    sheets: list[SheetInventory] = Field(default_factory=list)
    workbook_ranges: list[RangeInventory] = Field(default_factory=list)
    unsupported_features: list[UnsupportedFeature] = Field(default_factory=list)
    coverage: CoverageSummary


class GraphNodeKind(str, Enum):
    cell = "cell"
    range = "range"
    name = "name"
    external = "external"
    unsupported = "unsupported"


class GraphNode(BaseModel):
    id: str
    kind: GraphNodeKind
    label: str
    source_location: SourceLocation | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)


class GraphEdge(BaseModel):
    source: str
    target: str
    relationship: str = "references"
    formula: str | None = None
    source_location: SourceLocation | None = None


class FormulaGraph(Artifact):
    artifact_type: Literal["formula_graph"] = "formula_graph"
    nodes: list[GraphNode] = Field(default_factory=list)
    edges: list[GraphEdge] = Field(default_factory=list)
    unsupported_features: list[UnsupportedFeature] = Field(default_factory=list)


class ModuleCategory(str, Enum):
    input = "input"
    data_table = "data_table"
    formula_block = "formula_block"
    lookup_block = "lookup_block"
    output = "output"
    presentation = "presentation"
    external_dependency = "external_dependency"
    unsupported_opaque = "unsupported_opaque"
    other = "other"


class ActuarialHint(str, Enum):
    assumption = "assumption"
    rate_table = "rate_table"
    cashflow = "cashflow"
    projection = "projection"
    output = "output"
    unknown = "unknown"


class ModuleClassificationItem(BaseModel):
    id: str
    category: ModuleCategory
    confidence: float = Field(ge=0.0, le=1.0)
    reasons: list[str] = Field(default_factory=list)
    source_location: SourceLocation
    source_artifact_refs: list[str] = Field(default_factory=list)
    actuarial_hints: list[ActuarialHint] = Field(default_factory=list)


class ModuleClassification(Artifact):
    artifact_type: Literal["module_classification"] = "module_classification"
    items: list[ModuleClassificationItem] = Field(default_factory=list)
    unsupported_features: list[UnsupportedFeature] = Field(default_factory=list)


class ConfirmationQuestion(BaseModel):
    id: str
    prompt: str
    question_type: str
    source_location: SourceLocation | None = None
    options: list[str] = Field(default_factory=list)
    default: str | None = None
    rationale: str | None = None


class DecisionRecord(BaseModel):
    question_id: str
    decision: str
    reviewer: str | None = None
    timestamp: datetime = Field(default_factory=lambda: datetime.now(UTC))
    source_artifact_version: str = SCHEMA_VERSION


class ConfirmationTemplate(Artifact):
    artifact_type: Literal["confirmation_template"] = "confirmation_template"
    questions: list[ConfirmationQuestion] = Field(default_factory=list)
    decisions: list[DecisionRecord] = Field(default_factory=list)


class ArtifactMetadata(BaseModel):
    name: str
    path: str
    sha256: str
    bytes: int = Field(ge=0)
    schema_version: str = SCHEMA_VERSION


class RunMetadata(Artifact):
    artifact_type: Literal["run_metadata"] = "run_metadata"
    run_id: str
    workbook_sha256: str
    started_at: datetime
    completed_at: datetime | None = None
    artifacts: list[ArtifactMetadata] = Field(default_factory=list)

    @field_validator("run_id")
    @classmethod
    def run_id_not_empty(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("run_id must be non-empty")
        return value


def artifact_json_schema(model: type[BaseModel]) -> dict[str, Any]:
    return model.model_json_schema()


def artifact_file_name(model: Artifact | type[Artifact]) -> str:
    artifact_type = getattr(model, "artifact_type", None)
    if artifact_type is None and isinstance(model, type):
        artifact_type = model.model_fields["artifact_type"].default
    return f"{artifact_type}.json"


def as_path(value: str | Path) -> Path:
    return Path(value).expanduser().resolve()
