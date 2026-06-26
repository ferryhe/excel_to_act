"""Local filesystem artifact store with hashes, run namespaces, and validation."""

from __future__ import annotations

import hashlib
import json
import shutil
from datetime import UTC, datetime
from pathlib import Path
from typing import TypeVar

from pydantic import BaseModel

from excel_to_act.schemas import (
    ArtifactMetadata,
    ConfirmationTemplate,
    FormulaGraph,
    ModuleClassification,
    RunMetadata,
    WorkbookInventory,
    WorkbookManifest,
)
from excel_to_act.schemas.artifacts import SCHEMA_VERSION


ARTIFACT_NAMES: list[tuple[str, str]] = [
    ("workbook_manifest", "workbook_manifest.json"),
    ("inventory", "inventory.json"),
    ("dependency_graph", "dependency_graph.json"),
    ("module_classification", "module_classification.json"),
    ("confirmation_template", "confirmation_template.json"),
]

_MODEL_BY_FILE: dict[str, type[BaseModel]] = {
    "workbook_manifest.json": WorkbookManifest,
    "inventory.json": WorkbookInventory,
    "dependency_graph.json": FormulaGraph,
    "module_classification.json": ModuleClassification,
    "confirmation_template.json": ConfirmationTemplate,
    "run_metadata.json": RunMetadata,
}

ArtifactModel = TypeVar("ArtifactModel", bound=BaseModel)


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


class LocalArtifactStore:
    """Persist Phase 1 artifacts under workbook-hash/run-id namespaces.

    Canonical artifacts are stored at:

        <out>/workbooks/<workbook_sha256>/<run_id>/<artifact>.json

    The latest run is also copied to ``<out>/<artifact>.json`` for the PR-11 CLI contract.
    Copies are convenience aliases only; the run metadata points at canonical paths so multiple
    runs of the same workbook are preserved.
    """

    name = "local_artifact_store"

    def __init__(self, out_dir: Path) -> None:
        self.out_dir = out_dir.expanduser().resolve()

    def workbook_dir(self, workbook_sha256: str) -> Path:
        return self.out_dir / "workbooks" / workbook_sha256

    def run_dir(self, workbook_sha256: str, run_id: str) -> Path:
        return self.workbook_dir(workbook_sha256) / run_id

    def write_json(self, name: str, artifact: BaseModel, run_dir: Path) -> ArtifactMetadata:
        run_dir.mkdir(parents=True, exist_ok=True)
        path = run_dir / name
        data = artifact.model_dump_json(indent=2).encode("utf-8") + b"\n"
        path.write_bytes(data)
        return ArtifactMetadata(
            name=name,
            path=str(path),
            sha256=_sha256_bytes(data),
            bytes=len(data),
            schema_version=getattr(artifact, "schema_version", SCHEMA_VERSION),
        )

    def read_json(self, path: str | Path, model: type[ArtifactModel]) -> ArtifactModel:
        data = Path(path).read_bytes()
        artifact = model.model_validate_json(data)
        schema_version = getattr(artifact, "schema_version", None)
        if schema_version != SCHEMA_VERSION:
            raise ValueError(f"unsupported schema_version for {path}: {schema_version!r}")
        return artifact

    def read_artifact(self, metadata: ArtifactMetadata) -> BaseModel:
        model = _MODEL_BY_FILE.get(metadata.name)
        if model is None:
            raise ValueError(f"unknown artifact type: {metadata.name}")
        data = Path(metadata.path).read_bytes()
        if _sha256_bytes(data) != metadata.sha256:
            raise ValueError(f"artifact checksum mismatch: {metadata.path}")
        artifact = model.model_validate_json(data)
        schema_version = getattr(artifact, "schema_version", SCHEMA_VERSION)
        if schema_version != metadata.schema_version:
            raise ValueError(
                f"artifact schema_version mismatch for {metadata.name}: "
                f"{schema_version!r} != {metadata.schema_version!r}"
            )
        return artifact

    def read_run(self, workbook_sha256: str, run_id: str) -> RunMetadata:
        metadata = self.read_json(self.run_dir(workbook_sha256, run_id) / "run_metadata.json", RunMetadata)
        for artifact in metadata.artifacts:
            if artifact.name in {"artifact_index.json", "run_metadata.json"}:
                continue
            self.read_artifact(artifact)
        return metadata

    def _write_index(self, metadata: RunMetadata, run_dir: Path) -> ArtifactMetadata:
        index = {
            "schema_version": SCHEMA_VERSION,
            "workbook_sha256": metadata.workbook_sha256,
            "latest_run_id": metadata.run_id,
            "runs": [],
        }
        index_path = self.workbook_dir(metadata.workbook_sha256) / "artifact_index.json"
        if index_path.exists():
            index = json.loads(index_path.read_text())
            index.setdefault("runs", [])
        run_entry = {
            "run_id": metadata.run_id,
            "run_dir": str(run_dir),
            "started_at": metadata.started_at.isoformat(),
            "completed_at": metadata.completed_at.isoformat() if metadata.completed_at else None,
            "artifacts": [a.model_dump() for a in metadata.artifacts],
        }
        index["schema_version"] = SCHEMA_VERSION
        index["workbook_sha256"] = metadata.workbook_sha256
        index["latest_run_id"] = metadata.run_id
        index["runs"] = [r for r in index["runs"] if r.get("run_id") != metadata.run_id]
        index["runs"].append(run_entry)
        index_bytes = (json.dumps(index, indent=2, default=str) + "\n").encode("utf-8")
        index_path.parent.mkdir(parents=True, exist_ok=True)
        index_path.write_bytes(index_bytes)
        root_index_path = self.out_dir / "artifact_index.json"
        root_index_path.parent.mkdir(parents=True, exist_ok=True)
        root_index_path.write_bytes(index_bytes)
        return ArtifactMetadata(
            name="artifact_index.json",
            path=str(index_path),
            sha256=_sha256_bytes(index_bytes),
            bytes=len(index_bytes),
        )

    def _copy_latest_aliases(self, metadata: RunMetadata) -> None:
        self.out_dir.mkdir(parents=True, exist_ok=True)
        for artifact in metadata.artifacts:
            if artifact.name == "artifact_index.json":
                continue
            shutil.copy2(artifact.path, self.out_dir / artifact.name)

    def write_run(
        self,
        manifest: WorkbookManifest,
        inventory: WorkbookInventory,
        graph: FormulaGraph,
        classification: ModuleClassification,
        confirmation: ConfirmationTemplate,
        metadata: RunMetadata,
    ) -> RunMetadata:
        run_dir = self.run_dir(metadata.workbook_sha256, metadata.run_id)
        written = [
            self.write_json("workbook_manifest.json", manifest, run_dir),
            self.write_json("inventory.json", inventory, run_dir),
            self.write_json("dependency_graph.json", graph, run_dir),
            self.write_json("module_classification.json", classification, run_dir),
            self.write_json("confirmation_template.json", confirmation, run_dir),
        ]
        metadata.artifacts = written
        metadata.completed_at = datetime.now(UTC)
        meta = self.write_json("run_metadata.json", metadata, run_dir)
        metadata.artifacts.append(meta)
        index_meta = self._write_index(metadata, run_dir)
        metadata.artifacts.append(index_meta)
        # Rewrite metadata after adding index metadata so canonical run metadata is complete.
        metadata.artifacts[-2] = self.write_json("run_metadata.json", metadata, run_dir)
        self._copy_latest_aliases(metadata)
        return metadata
