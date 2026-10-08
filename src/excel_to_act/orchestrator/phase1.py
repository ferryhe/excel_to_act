"""Phase 1 orchestration pipeline."""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4

from openpyxl.utils.exceptions import InvalidFileException

from excel_to_act.classify.classifier import RuleBasedClassifier
from excel_to_act.confirm.templates import ConfirmationTemplateBuilder
from excel_to_act.graph.builder import RegexFormulaGraphBuilder
from excel_to_act.ingest.openpyxl_reader import OpenpyxlWorkbookReader
from excel_to_act.ingest.vba import extract_vba_project
from excel_to_act.inventory.extractor import OpenpyxlInventoryExtractor
from excel_to_act.inventory.vba_links import build_vba_edges, extract_vba_cell_links
from excel_to_act.report.handoff import build_handoff, render_handoff_markdown
from excel_to_act.schemas import (
    ConfirmationTemplate,
    CompletenessCheck,
    CompletenessReport,
    CompletenessStatus,
    CoverageSummary,
    FormulaGraph,
    ModuleClassification,
    RunMetadata,
    SourceLocation,
    UnsupportedFeature,
    UnsupportedSeverity,
    WorkbookInventory,
    WorkbookManifest,
)
from excel_to_act.store.local_store import LocalArtifactStore
from excel_to_act.verify.completeness import verify_completeness


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
        metadata = RunMetadata(run_id=f"{started:%Y%m%dT%H%M%S}-{uuid4().hex[:8]}", workbook_sha256=manifest.sha256, started_at=started)
        errors = [feature for feature in manifest.unsupported_features if feature.severity == UnsupportedSeverity.error]
        store = LocalArtifactStore(out_dir)
        if not errors:
            try:
                inventory: WorkbookInventory = self.extractor.extract(workbook_path, manifest)
            except InvalidFileException as exc:
                manifest.unsupported_features.append(
                    UnsupportedFeature(
                        feature_type="input_read_error",
                        description=f"Could not read workbook: {type(exc).__name__}: {exc}",
                        source_location=SourceLocation(
                            workbook_path=str(workbook_path.expanduser().resolve()),
                            object_type="workbook",
                        ),
                        severity=UnsupportedSeverity.error,
                        opaque=False,
                    )
                )
                errors = [manifest.unsupported_features[-1]]

        if errors:
            reason = "; ".join(feature.description for feature in errors)
            completeness = CompletenessReport(
                workbook_sha256=manifest.sha256,
                status=CompletenessStatus.fail,
                checks=[CompletenessCheck(name="input_readable", passed=False, severity=UnsupportedSeverity.error, detail=reason)],
                blocking_reasons=[reason],
            )
            metadata.completeness_status = completeness.status
            metadata = store.write_failed_run(manifest, completeness, metadata)
            inventory = WorkbookInventory(
                workbook_sha256=manifest.sha256,
                unsupported_features=manifest.unsupported_features,
                coverage=CoverageSummary(
                    recognized_inventory_objects=0,
                    unsupported_or_opaque_objects=0,
                    discovered_workbook_objects=0,
                ),
            )
            handoff = build_handoff(
                manifest, inventory, FormulaGraph(), ModuleClassification(), ConfirmationTemplate(), completeness, metadata
            )
            handoff.summary = {}
            handoff.coverage = None
            handoff.next_step = "rerun_inspect"
            handoff.next_actions = [
                "Resolve the input error and rerun inspect before proceeding.",
                "Not produced: inventory.json, dependency_graph.json, module_classification.json, confirmation_template.json.",
            ]
            run_dir = store.run_dir(metadata.workbook_sha256, metadata.run_id)
            return store.append_artifacts(
                metadata,
                [
                    store.write_json("handoff.json", handoff, run_dir),
                    store.write_text("handoff.md", render_handoff_markdown(handoff), run_dir),
                ],
            )

        graph: FormulaGraph = self.graph_builder.build(inventory)
        self._integrate_vba(workbook_path, inventory, graph)
        classification: ModuleClassification = self.classifier.classify(inventory, graph)
        confirmation: ConfirmationTemplate = self.confirmation_builder.build(classification)
        # Completeness is verified before anything is persisted, and the handoff is
        # written last so it can point at every artifact path.
        completeness = verify_completeness(manifest, inventory, graph, classification)
        metadata.completeness_status = completeness.status
        metadata = store.write_run(manifest, inventory, graph, classification, confirmation, metadata, completeness=completeness)
        run_dir = store.run_dir(metadata.workbook_sha256, metadata.run_id)
        handoff = build_handoff(manifest, inventory, graph, classification, confirmation, completeness, metadata)
        return store.append_artifacts(
            metadata,
            [
                store.write_json("handoff.json", handoff, run_dir),
                store.write_text("handoff.md", render_handoff_markdown(handoff), run_dir),
            ],
        )

    def _integrate_vba(self, workbook_path: Path, inventory: WorkbookInventory, graph: FormulaGraph) -> None:
        """Extract VBA and fold its cell references into the dependency graph.

        The VBA module list is stored on the inventory; references become edges
        with ``relationship="vba_ref"`` in the same node space as formula edges.
        """
        project = extract_vba_project(workbook_path)
        inventory.vba_modules = list(project.modules)
        if project.available and project.oletools_missing:
            inventory.unsupported_features.append(
                UnsupportedFeature(
                    feature_type="vba_extraction_skipped",
                    description="Workbook contains VBA but the optional `oletools` dependency is not installed; VBA source and references were not extracted.",
                    source_location=SourceLocation(workbook_path=str(workbook_path), object_type="workbook"),
                    severity=UnsupportedSeverity.warning,
                    opaque=True,
                )
            )
        if project.error:
            inventory.unsupported_features.append(
                UnsupportedFeature(
                    feature_type="vba_extraction_failed",
                    description=f"VBA source could not be parsed: {project.error}",
                    source_location=SourceLocation(workbook_path=str(workbook_path), object_type="workbook"),
                    severity=UnsupportedSeverity.warning,
                    opaque=True,
                )
            )
        if not project.modules:
            return
        refs = extract_vba_cell_links(project)
        nodes, edges = build_vba_edges(
            refs,
            workbook_path=str(workbook_path),
            known_sheets={sheet.name for sheet in inventory.sheets},
        )
        # Merge into the existing graph without duplicating nodes/edges the formula
        # builder already produced.
        seen_nodes = {node.id for node in graph.nodes}
        graph.nodes.extend(node for node in nodes if node.id not in seen_nodes)
        seen_edges = {(edge.source, edge.target, edge.relationship) for edge in graph.edges}
        for edge in edges:
            key = (edge.source, edge.target, edge.relationship)
            if key not in seen_edges:
                seen_edges.add(key)
                graph.edges.append(edge)
