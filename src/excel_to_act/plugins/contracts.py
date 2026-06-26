"""Typed plugin contracts for Phase 1 stages."""

from __future__ import annotations

from pathlib import Path
from typing import Protocol, runtime_checkable

from excel_to_act.schemas import (
    ConfirmationTemplate,
    FormulaGraph,
    ModuleClassification,
    RunMetadata,
    WorkbookInventory,
    WorkbookManifest,
)


@runtime_checkable
class WorkbookReader(Protocol):
    name: str

    def read_manifest(self, workbook_path: Path) -> WorkbookManifest: ...


@runtime_checkable
class InventoryExtractor(Protocol):
    name: str

    def extract(self, workbook_path: Path, manifest: WorkbookManifest) -> WorkbookInventory: ...


@runtime_checkable
class GraphBuilder(Protocol):
    name: str

    def build(self, inventory: WorkbookInventory) -> FormulaGraph: ...


@runtime_checkable
class Classifier(Protocol):
    name: str

    def classify(self, inventory: WorkbookInventory, graph: FormulaGraph) -> ModuleClassification: ...


@runtime_checkable
class ConfirmationBuilder(Protocol):
    name: str

    def build(self, classification: ModuleClassification) -> ConfirmationTemplate: ...


@runtime_checkable
class ArtifactStore(Protocol):
    name: str

    def write_run(
        self,
        manifest: WorkbookManifest,
        inventory: WorkbookInventory,
        graph: FormulaGraph,
        classification: ModuleClassification,
        confirmation: ConfirmationTemplate,
        metadata: RunMetadata,
    ) -> RunMetadata: ...
