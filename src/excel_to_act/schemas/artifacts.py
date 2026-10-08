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
    source_identity: str | None = None

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
    opaque_reason: str | None = None
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
    source_identity: str | None = None
    address: str
    row: int = Field(ge=1)
    column: int = Field(ge=1)
    kind: CellKind
    value: str | int | float | bool | None = None
    data_type: str | None = None
    number_format: str | None = None
    formula: str | None = None
    style_id: int | None = None
    # Cached result Excel stored next to the formula (``<v>`` in the sheet XML).
    # ``None`` means "not stored", which is different from "stored as empty".
    cached_value: str | int | float | bool | None = None
    # True only for formula cells whose cached result was actually read.
    cached_value_available: bool = False
    # Exact OOXML values are additive to the legacy normalized projection.
    raw_value_text: str | None = None
    raw_formula_text: str | None = None
    raw_formula_attributes: dict[str, str] = Field(default_factory=dict)
    formula_present: bool | None = None
    ooxml_cell_type: str | None = None
    cached_text: str | None = None
    cached_text_present: bool | None = None
    workbook_date_system: str | None = None
    date_serial_text: str | None = None


class RangeInventory(BaseModel):
    source_location: SourceLocation
    name: str | None = None
    address: str
    kind: str
    metadata: dict[str, Any] = Field(default_factory=dict)
    source_identity: str | None = None


class SheetInventory(BaseModel):
    source_location: SourceLocation
    name: str
    index: int = Field(ge=0)
    max_row: int = Field(ge=0)
    max_column: int = Field(ge=0)
    state: str = "visible"
    source_identity: str | None = None
    cells: list[CellInventory] = Field(default_factory=list)
    ranges: list[RangeInventory] = Field(default_factory=list)
    layout_objects: list[RangeInventory] = Field(default_factory=list)


class CoverageSummary(BaseModel):
    recognized_inventory_objects: int = Field(ge=0)
    unsupported_or_opaque_objects: int = Field(ge=0)
    discovered_workbook_objects: int = Field(ge=0)

class VbaModule(BaseModel):
    """A VBA module, class, document, or form extracted from ``vbaProject.bin``."""

    name: str
    kind: str = "StdModule"  # StdModule | ClassModule | UserForm | Document
    code: str = ""
    procedures: list[str] = Field(default_factory=list)


class CheckboxBinding(BaseModel):
    sheet: str
    sheet_part: str
    shape_id: str | None = None
    control_name: str | None = None
    legacy_control_id: str | None = None
    linked_cell_raw: str | None = None
    linked_cell: str | None = None
    linked_sheet: str | None = None
    linked_address: str | None = None
    binding_status: Literal["resolved", "named", "dynamic", "invalid", "unresolved"]
    sources: list[str] = Field(default_factory=list)
    diagnostic: str | None = None


class CheckboxBindings(Artifact):
    artifact_type: Literal["checkbox_bindings"] = "checkbox_bindings"
    workbook_sha256: str
    status: Literal["complete", "partial", "not_applicable"]
    bindings: list[CheckboxBinding] = Field(default_factory=list)
    diagnostics: list[str] = Field(default_factory=list)


class ActiveXControl(BaseModel):
    sheet: str
    sheet_part: str
    sheet_code_name: str | None = None
    shape_id: str
    control_name: str
    class_id: str | None = None
    part: str
    binary_part: str | None = None
    event_procedures: list[str] = Field(default_factory=list)
    binding_status: Literal["resolved", "unresolved", "unavailable"]
    diagnostic: str | None = None


class ActiveXEvents(Artifact):
    artifact_type: Literal["activex_events"] = "activex_events"
    workbook_sha256: str
    status: Literal["complete", "partial", "not_applicable", "unavailable"]
    controls: list[ActiveXControl] = Field(default_factory=list)
    diagnostics: list[str] = Field(default_factory=list)


class VbaSource(BaseModel):
    name: str
    kind: str
    procedures: list[str] = Field(default_factory=list)
    source_file: str
    sha256: str


class VbaHandoff(Artifact):
    artifact_type: Literal["vba_handoff"] = "vba_handoff"
    workbook_sha256: str
    available: bool
    status: Literal["complete", "partial", "not_applicable", "unavailable"]
    modules: list[VbaSource] = Field(default_factory=list)
    diagnostics: list[str] = Field(default_factory=list)


class WorkbookInventory(Artifact):
    artifact_type: Literal["workbook_inventory"] = "workbook_inventory"
    workbook_sha256: str
    sheets: list[SheetInventory] = Field(default_factory=list)
    workbook_ranges: list[RangeInventory] = Field(default_factory=list)
    vba_modules: list[VbaModule] = Field(default_factory=list)
    unsupported_features: list[UnsupportedFeature] = Field(default_factory=list)
    coverage: CoverageSummary


class GraphNodeKind(str, Enum):
    cell = "cell"
    range = "range"
    name = "name"
    external = "external"
    unsupported = "unsupported"
    vba = "vba"


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
    confidence: float | None = None
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


class CompletenessStatus(str, Enum):
    # `pass` is a keyword; the serialized value stays "pass".
    ok = "pass"
    warn = "warn"
    fail = "fail"


class CompletenessCheck(BaseModel):
    """One independent assertion about whether the output accounts for the input."""

    name: str
    passed: bool
    severity: UnsupportedSeverity = UnsupportedSeverity.warning
    expected: int | None = None
    actual: int | None = None
    detail: str = ""


class ObjectCoverage(BaseModel):
    """Identity counts for one logical kind or the separate package-part ledger."""

    sheet_name: str | None = None
    kind: str
    expected: int = Field(ge=0)
    actual: int = Field(ge=0)
    missing_identities: list[str] = Field(default_factory=list)
    unexpected_identities: list[str] = Field(default_factory=list)
    duplicate_identities: list[str] = Field(default_factory=list)
    opaque_expected: int = Field(ge=0, default=0)
    opaque_actual: int = Field(ge=0, default=0)


class CompletenessReport(Artifact):
    artifact_type: Literal["completeness_report"] = "completeness_report"
    workbook_sha256: str
    run_id: str | None = None
    status: CompletenessStatus = CompletenessStatus.ok
    checks: list[CompletenessCheck] = Field(default_factory=list)
    # Counts are recomputed independently of WorkbookInventory.coverage, whose
    # `discovered` is currently an identity and therefore proves nothing.
    recognized_inventory_objects: int = Field(ge=0, default=0)
    unsupported_or_opaque_objects: int = Field(ge=0, default=0)
    discovered_workbook_objects: int = Field(ge=0, default=0)
    object_coverage: list[ObjectCoverage] = Field(default_factory=list)
    gaps: list[SourceLocation] = Field(default_factory=list)
    blocking_reasons: list[str] = Field(default_factory=list)


class HandoffArtifactRef(BaseModel):
    name: str
    path: str
    kind: str
    count: int | None = None
    sha256: str


class Handoff(Artifact):
    """The single artifact a downstream agent reads to continue the pipeline.

    Machine-readable summary of what Step 1 produced, what it could not parse,
    and what must happen before Step 3 may trust the output.
    """

    artifact_type: Literal["handoff"] = "handoff"
    step: str = "step1_decomposition"
    next_step: str = "step2_index"
    workbook_sha256: str
    workbook_path: str | None = None
    run_id: str | None = None
    produced_at: datetime = Field(default_factory=lambda: datetime.now(UTC))
    status: CompletenessStatus = CompletenessStatus.ok
    # One-glance counts (sheets / cells / formula_cells / edges / ...). Machine
    # readable so Step 2 can index it, and rendered as the TL;DR in handoff.md.
    summary: dict[str, int] = Field(default_factory=dict)
    artifacts: list[HandoffArtifactRef] = Field(default_factory=list)
    coverage: CoverageSummary | None = None
    opaque_summary: list[dict[str, Any]] = Field(default_factory=list)
    blockers: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    next_actions: list[str] = Field(default_factory=list)


class Step2ArtifactRef(BaseModel):
    name: str
    path: str
    sha256: str


class Step2IndexEntry(BaseModel):
    source_id: str
    source_path: str | None = None
    source_sha256: str | None = None
    run_id: str | None = None
    batch_id: str | None = None
    run_path: str | None = None
    handoff_path: str | None = None
    status: str
    ready_for_next_step: bool = False
    metrics_state: str = "unavailable"
    metrics: dict[str, Any] = Field(default_factory=dict)
    thresholds: dict[str, Any] = Field(default_factory=dict)
    artifacts: list[Step2ArtifactRef] = Field(default_factory=list)
    blockers: list[str] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    next_actions: list[dict[str, Any]] = Field(default_factory=list)
    final_output: str | None = None


class Step2Index(BaseModel):
    schema_version: Literal["step2.index.v1"] = "step2.index.v1"
    step1_root: str
    input_handoff_path: str
    input_handoff_sha256: str
    batch_id: str | None = None
    status: Literal["pass", "partial", "blocked"] = "partial"
    validation_status: Literal["not_run", "pass", "partial", "blocked"] = "not_run"
    entries: list[Step2IndexEntry] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    metrics: dict[str, int] = Field(default_factory=dict)


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
    completeness_status: str | None = None
    artifacts: list[ArtifactMetadata] = Field(default_factory=list)

    @field_validator("run_id")
    @classmethod
    def run_id_not_empty(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("run_id must be non-empty")
        return value


class ViewRecord(BaseModel):
    record_id: str
    record_type: str
    source_location: SourceLocation
    facts: dict[str, Any] = Field(default_factory=dict)


class ViewChunk(BaseModel):
    chunk_id: str
    record_ids: list[str]
    estimated_tokens: int = Field(ge=0)
    over_budget_record_ids: list[str] = Field(default_factory=list)


class WorkbookView(Artifact):
    artifact_type: Literal["workbook_view"] = "workbook_view"
    view_id: str
    source_id: str
    source_sha256: str
    source_run_id: str
    source_schema_version: str
    scope: Literal["workbook", "sheet", "region"]
    sheet_name: str | None = None
    region_address: str | None = None
    records: list[ViewRecord]
    chunks: list[ViewChunk]
    opaque_report: list[str] = Field(default_factory=list)


def artifact_json_schema(model: type[BaseModel]) -> dict[str, Any]:
    return model.model_json_schema()


def artifact_file_name(model: Artifact | type[Artifact]) -> str:
    artifact_type = getattr(model, "artifact_type", None)
    if artifact_type is None and isinstance(model, type):
        artifact_type = model.model_fields["artifact_type"].default
    return f"{artifact_type}.json"


def as_path(value: str | Path) -> Path:
    return Path(value).expanduser().resolve()
