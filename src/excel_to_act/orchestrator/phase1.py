"""Phase 1 orchestration pipeline."""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4

from excel_to_act.classify.classifier import RuleBasedClassifier
from excel_to_act.confirm.templates import ConfirmationTemplateBuilder
from excel_to_act.graph.builder import RegexFormulaGraphBuilder
from excel_to_act.ingest.openpyxl_reader import OpenpyxlWorkbookReader
from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor
from excel_to_act.schemas import (
    ConfirmationTemplate,
    FormulaGraph,
    ModuleClassification,
    RunMetadata,
    WorkbookInventory,
    WorkbookManifest,
)
from excel_to_act.store.local_store import LocalArtifactStore


class Phase1Artifacts(tuple):
    pass


class Phase1Orchestrator:
    """State-machine style orchestration for backend Phase 1 artifacts."""

    def __init__(self) -> None:
        self.reader = OpenpyxlWorkbookReader()
        self.extractor = OpenpyxlInventoryExtractor()
        self.graph_builder = RegexFormulaGraphBuilder()
        self.classifier = RuleBasedClassifier()
        self.confirmation_builder = ConfirmationTemplateBuilder()

    def run(self, workbook_path: Path, out_dir: Path) -> RunMetadata:
        started = datetime.now(UTC)
        manifest: WorkbookManifest = self.reader.read_manifest(workbook_path)
        inventory: WorkbookInventory = self.extractor.extract(workbook_path, manifest)
        graph: FormulaGraph = self.graph_builder.build(inventory)
        classification: ModuleClassification = self.classifier.classify(inventory, graph)
        confirmation: ConfirmationTemplate = self.confirmation_builder.build(classification)
        metadata = RunMetadata(run_id=f"{started:%Y%m%dT%H%M%S}-{uuid4().hex[:8]}", workbook_sha256=manifest.sha256, started_at=started)
        return LocalArtifactStore(out_dir).write_run(manifest, inventory, graph, classification, confirmation, metadata)
