"""Build a navigation and quality index from current Step 1 handoffs."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

from excel_to_act.schemas import SCHEMA_VERSION, Step2ArtifactRef, Step2Index, Step2IndexEntry
from excel_to_act.store.local_store import _MODEL_BY_FILE

_INDEX_NAME = "index.json"
_SUMMARY_NAME = "INDEX.md"
_DEFAULT_OUT = Path("output/step2_index")


def _diagnostic(code: str, message: str, *, source_id: str | None = None) -> dict[str, Any]:
    return {"code": code, "severity": "error", "message": message, "source_id": source_id}


def _artifact_ref_diagnostic(
    code: str,
    message: str,
    source_id: str,
    ref: Any = None,
) -> dict[str, Any]:
    diagnostic = _diagnostic(code, message, source_id=source_id)
    diagnostic["category"] = "step2_artifact_reference"
    if isinstance(ref, dict):
        if isinstance(ref.get("name"), str):
            diagnostic["artifact"] = ref["name"]
        if isinstance(ref.get("path"), str):
            diagnostic["path"] = ref["path"]
    return diagnostic


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode("utf-8")


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _relative_to_root(root: Path, path: Path, label: str) -> str:
    try:
        return path.resolve().relative_to(root).as_posix()
    except ValueError as exc:
        raise ValueError(f"{label} resolves outside --step1-root") from exc


def _top_handoff_path(value: Path, root: Path) -> tuple[Path, str]:
    value = value.expanduser()
    candidates = [value.resolve()] if value.is_absolute() else [(Path.cwd() / value).resolve(), (root / value).resolve()]
    for candidate in candidates:
        if candidate.is_file():
            return candidate, _relative_to_root(root, candidate, "--handoff")
    candidate = candidates[-1]
    relative = _relative_to_root(root, candidate, "--handoff")
    raise ValueError(f"handoff file does not exist: {relative}")


def _root_relative(root: Path, value: str) -> tuple[Path, str]:
    relative = Path(value)
    if relative.is_absolute():
        raise ValueError("handoff_path must be relative to --step1-root")
    resolved = (root / relative).resolve()
    return resolved, _relative_to_root(root, resolved, "batch handoff_path")


def _artifact_path(root: Path, base: str | None, value: str) -> str:
    child = Path(value)
    if child.is_absolute():
        raise ValueError("artifact paths must be relative")
    resolved = (root / Path(base) / child if base else root / child).resolve()
    return _relative_to_root(root, resolved, "artifact path")


def _items(value: Any) -> list[dict[str, Any]]:
    return [item for item in value if isinstance(item, dict)] if isinstance(value, list) else []


def _text(value: Any) -> str | None:
    return value if isinstance(value, str) and value else None


def _validate_source_handoff(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict) or value.get("schema_version") != "step1.v1":
        raise ValueError("source handoff must use schema_version step1.v1")
    if not isinstance(value.get("source"), dict):
        raise ValueError("step1.v1 handoff must contain a source object")
    if not isinstance(value.get("status"), str):
        raise ValueError("step1.v1 handoff must contain a string status")
    if not isinstance(value.get("artifacts"), list):
        raise ValueError("step1.v1 handoff must contain an artifacts list")
    return value


def _handoff_identity_mismatch(entry: dict[str, Any], handoff: dict[str, Any]) -> str | None:
    source = handoff["source"]
    identities = (
        ("source path", _text(entry.get("relative_path")), _text(source.get("path"))),
        ("source SHA-256", _text(entry.get("sha256")), _text(source.get("sha256"))),
        ("run ID", _text(entry.get("run_id")), _text(handoff.get("run_id"))),
    )
    for label, batch_value, handoff_value in identities:
        if batch_value is not None and handoff_value is not None and batch_value != handoff_value:
            return f"Batch entry {label} {batch_value!r} does not match per-source handoff {handoff_value!r}."
    return None


def _source_id(path: str | None, digest: str | None, run_id: str | None, ordinal: int) -> str:
    identity: Any = [path, digest, run_id] if path or digest or run_id else ["entry", ordinal]
    raw = json.dumps(identity, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:20]


def _artifact_refs(
    root: Path,
    handoff: dict[str, Any],
    source_id: str,
    diagnostics: list[dict[str, Any]],
) -> list[Step2ArtifactRef]:
    refs = handoff.get("artifacts", [])
    if not isinstance(refs, list):
        diagnostics.append(_artifact_ref_diagnostic("artifact_refs_invalid", "Source handoff artifacts must be a list.", source_id))
        return []
    base = handoff.get("artifact_paths_relative_to") or handoff.get("run_path")
    if refs and not isinstance(base, str):
        diagnostics.extend(
            _artifact_ref_diagnostic(
                "artifact_base_missing",
                "Source handoff has artifacts but no artifact_paths_relative_to or run_path.",
                source_id,
                ref,
            )
            for ref in refs
        )
        return []

    result: list[Step2ArtifactRef] = []
    for ref in refs:
        if not isinstance(ref, dict):
            diagnostics.append(_artifact_ref_diagnostic("artifact_ref_invalid", "Artifact reference must be an object.", source_id))
            continue
        name, path, digest = ref.get("name"), ref.get("path"), ref.get("sha256")
        if not all(isinstance(item, str) and item for item in (name, path, digest)):
            diagnostics.append(_artifact_ref_diagnostic("artifact_ref_incomplete", "Artifact reference needs string name, path, and sha256 fields.", source_id, ref))
            continue
        try:
            normalized = _artifact_path(root, base, path)
        except ValueError as exc:
            diagnostics.append(_artifact_ref_diagnostic("artifact_path_invalid", str(exc), source_id, ref))
            continue
        result.append(Step2ArtifactRef(name=name, path=normalized, sha256=digest))
    return result


def _entry_from_handoff(
    root: Path,
    handoff: dict[str, Any],
    handoff_path: str,
    ordinal: int,
    fallback: dict[str, Any] | None = None,
) -> Step2IndexEntry:
    fallback = fallback or {}
    source = handoff.get("source") if isinstance(handoff.get("source"), dict) else {}
    source_path = _text(fallback.get("relative_path")) or _text(source.get("path"))
    source_sha = _text(fallback.get("sha256")) or _text(source.get("sha256")) or _text(handoff.get("source_sha256"))
    run_id = _text(fallback.get("run_id")) or _text(handoff.get("run_id"))
    source_id = _source_id(source_path, source_sha, run_id, ordinal)
    diagnostics = _items(handoff.get("diagnostics"))
    if not diagnostics:
        diagnostics = _items(fallback.get("diagnostics"))
    artifacts = _artifact_refs(root, handoff, source_id, diagnostics)
    for item in diagnostics:
        if item.get("category") == "step2_artifact_reference":
            item["source_path"] = source_path
    ready = handoff.get("ready_for_next_step", fallback.get("ready_for_next_step", False)) is True
    metrics = handoff.get("metrics") if isinstance(handoff.get("metrics"), dict) else fallback.get("metrics", {})
    thresholds = handoff.get("thresholds") if isinstance(handoff.get("thresholds"), dict) else {}
    blockers = handoff.get("blockers") if isinstance(handoff.get("blockers"), list) else []
    next_actions = handoff.get("next_actions") if isinstance(handoff.get("next_actions"), list) else []
    return Step2IndexEntry(
        source_id=source_id,
        source_path=source_path,
        source_sha256=source_sha,
        run_id=run_id,
        batch_id=_text(fallback.get("batch_id")) or _text(handoff.get("batch_id")),
        run_path=_text(handoff.get("run_path")) or _text(fallback.get("run_path")),
        handoff_path=handoff_path,
        status=str(fallback.get("status") or handoff.get("status") or "error"),
        ready_for_next_step=ready,
        metrics_state=str(handoff.get("metrics_state") or fallback.get("metrics_state") or "unavailable"),
        metrics=metrics if isinstance(metrics, dict) else {},
        thresholds=thresholds,
        artifacts=artifacts,
        blockers=[item for item in blockers if isinstance(item, str)],
        diagnostics=diagnostics,
        next_actions=_items(next_actions),
        final_output=_text(handoff.get("final_output")),
    )


def _missing_handoff_entry(
    item: Any,
    ordinal: int,
    path: str | None,
    detail: str,
    code: str = "source_handoff_unavailable",
) -> Step2IndexEntry:
    batch_entry = item if isinstance(item, dict) else {}
    source_path = _text(batch_entry.get("relative_path"))
    source_sha = _text(batch_entry.get("sha256"))
    run_id = _text(batch_entry.get("run_id"))
    source_id = _source_id(source_path, source_sha, run_id, ordinal)
    diagnostic = _diagnostic(code, detail, source_id=source_id)
    diagnostics = _items(batch_entry.get("diagnostics")) + [diagnostic]
    return Step2IndexEntry(
        source_id=source_id,
        source_path=source_path,
        source_sha256=source_sha,
        run_id=run_id,
        batch_id=_text(batch_entry.get("batch_id")),
        run_path=_text(batch_entry.get("run_path")),
        handoff_path=path,
        status=str(batch_entry.get("status") or "error"),
        ready_for_next_step=False,
        metrics_state=str(batch_entry.get("metrics_state") or "unavailable"),
        metrics=batch_entry.get("metrics") if isinstance(batch_entry.get("metrics"), dict) else {},
        blockers=[detail],
        diagnostics=diagnostics,
        next_actions=_items(batch_entry.get("next_actions")),
    )


def _resolve_batch_entry(root: Path, item: Any, ordinal: int) -> Step2IndexEntry:
    if not isinstance(item, dict):
        return _missing_handoff_entry(item, ordinal, None, "Batch entry must be an object.")
    raw_path = item.get("handoff_path")
    if not isinstance(raw_path, str) or not raw_path:
        return _missing_handoff_entry(item, ordinal, None, "Batch entry has no per-source handoff_path.")
    try:
        path, relative = _root_relative(root, raw_path)
        handoff = _read_json(path)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        try:
            display_path = _root_relative(root, raw_path)[1]
        except ValueError:
            display_path = None
        return _missing_handoff_entry(item, ordinal, display_path, f"Cannot load per-source handoff: {exc}")
    try:
        handoff = _validate_source_handoff(handoff)
    except ValueError as exc:
        return _missing_handoff_entry(
            item,
            ordinal,
            relative,
            f"Invalid per-source handoff: {exc}",
            code="source_handoff_invalid",
        )
    mismatch = _handoff_identity_mismatch(item, handoff)
    if mismatch:
        return _missing_handoff_entry(item, ordinal, relative, mismatch, code="source_handoff_mismatch")
    return _entry_from_handoff(root, handoff, relative, ordinal, item)


def _markdown(index: Step2Index) -> str:
    def cell(value: Any) -> str:
        return str(value).replace("|", r"\|").replace("`", r"\`").replace("\r", " ").replace("\n", " ")

    lines = [
        "# Step 2 Artifact Index",
        "",
        f"**Status:** `{index.status}`  ",
        f"**Batch:** `{index.batch_id or 'single source'}`  ",
        f"**Sources:** {len(index.entries)}  ",
        f"**Reference validation:** `{index.validation_status}`  ",
        f"**Step 1 root:** `{cell(index.step1_root)}`",
        "",
        "This index provides navigation and quality context. Load Step 1 artifacts as needed; it makes no semantic conclusions.",
        "",
        "## Sources",
        "",
        "| Source | SHA-256 | Status | Ready | Artifacts |",
        "| --- | --- | --- | --- | --- |",
    ]
    for entry in index.entries:
        artifacts = ", ".join(f"`{cell(ref.name)}`" for ref in entry.artifacts) or "—"
        lines.append(f"| `{cell(entry.source_path or entry.source_id)}` | `{cell(entry.source_sha256 or 'unavailable')}` | {cell(entry.status)} | {str(entry.ready_for_next_step).lower()} | {artifacts} |")
        lines.extend(["", f"### `{cell(entry.source_path or entry.source_id)}`", "", f"- Run: `{cell(entry.run_id or 'unavailable')}`"])
        if entry.metrics:
            lines.append(f"- Metrics: `{cell(json.dumps(entry.metrics, ensure_ascii=False, sort_keys=True))}`")
        lines.extend(f"- Blocker: {cell(blocker)}" for blocker in entry.blockers)
        lines.extend(f"- Diagnostic: {cell(item.get('code', 'unknown'))} — {cell(item.get('message', ''))}" for item in entry.diagnostics)
        lines.extend(f"- Artifact: `{cell(ref.name)}` · `{cell(ref.path)}` · SHA-256 `{cell(ref.sha256)}`" for ref in entry.artifacts)
    lines.extend(["", "## Index diagnostics", ""])
    lines.extend(f"- {cell(item.get('code', 'unknown'))}: {cell(item.get('message', ''))}" for item in index.diagnostics)
    if not index.diagnostics:
        lines.append("- None")
    return "\n".join(lines) + "\n"


def _write_index(out_dir: Path, index: Step2Index) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    index_path, summary_path = out_dir / _INDEX_NAME, out_dir / _SUMMARY_NAME
    index_path.write_bytes(_json_bytes(index.model_dump(mode="json")))
    summary_path.write_text(_markdown(index), encoding="utf-8")
    return index_path, summary_path


def _error_result(handoff: Path, code: str, message: str) -> dict[str, Any]:
    return {
        "tool": "step2.index",
        "status": "blocked",
        "source": {"handoff": str(handoff)},
        "run_id": None,
        "artifacts": [],
        "metrics": {},
        "diagnostics": [_diagnostic(code, message)],
        "retryable": False,
        "next_tool": None,
    }


_ARTIFACT_REBUILDERS = {
    "workbook_manifest.json": "manifest.read",
    "inventory.json": "inventory.extract",
    "dependency_graph.json": "graph.build",
    "module_classification.json": "classify.rules",
    "confirmation_template.json": "confirmation.build",
    "handoff.json": "report.handoff",
}


def _first_source_location(value: Any) -> dict[str, Any] | None:
    if isinstance(value, dict):
        location = value.get("source_location")
        if isinstance(location, dict):
            return location
        for item in value.values():
            found = _first_source_location(item)
            if found:
                return found
    elif isinstance(value, list):
        for item in value:
            found = _first_source_location(item)
            if found:
                return found
    return None


def validate_saved_index(index_path: Path, step1_root: Path) -> dict[str, Any]:
    """Check a saved index and every referenced artifact using Step 1 contracts."""
    root = step1_root.expanduser().resolve()
    result = {
        "tool": "step2.index.validate",
        "status": "pass",
        "source": {"index": str(index_path)},
        "run_id": None,
        "artifacts": [],
        "metrics": {"entries_checked": 0, "artifacts_checked": 0, "input_handoff_checked": 0},
        "diagnostics": [],
        "retryable": False,
        "next_tool": None,
    }

    def report(code: str, message: str, entry: Any, ref: Any, payload: Any = None) -> None:
        name = getattr(ref, "name", "unknown")
        candidate = _ARTIFACT_REBUILDERS.get(name) or _ARTIFACT_REBUILDERS.get(Path(getattr(ref, "path", "")).name)
        diagnostic = _diagnostic(code, message, source_id=getattr(entry, "source_id", None))
        diagnostic.update({
            "source_path": getattr(entry, "source_path", None),
            "artifact": name,
            "path": getattr(ref, "path", None),
        })
        location = _first_source_location(payload)
        if location:
            diagnostic["source_location"] = location
        if candidate and code in {"artifact_missing", "artifact_checksum_mismatch", "artifact_json_invalid", "artifact_schema_invalid", "artifact_identity_mismatch"}:
            diagnostic["next_tool"] = candidate
        result["diagnostics"].append(diagnostic)

    try:
        if not root.is_dir():
            raise ValueError(f"Step 1 output root does not exist: {root}")
        index = Step2Index.model_validate(_read_json(index_path.expanduser().resolve()))
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        result["status"] = "blocked"
        result["diagnostics"].append(_diagnostic("index_invalid", f"Cannot load saved Step 2 index: {exc}"))
        return result

    result["source"] = {"index": str(index_path.expanduser().resolve())}
    result["run_id"] = index.batch_id
    result["metrics"]["entries_checked"] = len(index.entries)

    for entry in index.entries:
        for item in entry.diagnostics:
            if item.get("category") == "step2_artifact_reference":
                result["diagnostics"].append({**item, "source_id": entry.source_id, "source_path": entry.source_path})

    def input_handoff_error(code: str, message: str) -> None:
        diagnostic = _diagnostic(code, message)
        diagnostic.update({"artifact": "input_handoff", "path": index.input_handoff_path})
        if len(index.entries) == 1:
            diagnostic.update({"source_id": index.entries[0].source_id, "source_path": index.entries[0].source_path})
        result["diagnostics"].append(diagnostic)

    input_handoff = Path(index.input_handoff_path)
    if input_handoff.is_absolute():
        input_handoff_error("input_handoff_path_invalid", "Index input_handoff_path must be relative to --step1-root.")
    else:
        resolved_handoff = (root / input_handoff).resolve()
        try:
            relative_handoff = resolved_handoff.relative_to(root)
        except ValueError:
            input_handoff_error("input_handoff_path_invalid", "Index input_handoff_path resolves outside --step1-root.")
        else:
            if not resolved_handoff.is_file():
                input_handoff_error("input_handoff_missing", f"Index input handoff does not exist: {relative_handoff.as_posix()}")
            elif _hash_file(resolved_handoff) != index.input_handoff_sha256:
                input_handoff_error("input_handoff_checksum_mismatch", f"SHA-256 does not match for input handoff {relative_handoff.as_posix()}.")
            else:
                try:
                    _read_json(resolved_handoff)
                except (OSError, ValueError, json.JSONDecodeError) as exc:
                    input_handoff_error("input_handoff_json_invalid", f"Invalid JSON in input handoff {relative_handoff.as_posix()}: {exc}")
                else:
                    result["metrics"]["input_handoff_checked"] = 1
    rebuilders: set[str] = set()
    for entry in index.entries:
        for ref in entry.artifacts:
            result["metrics"]["artifacts_checked"] += 1
            path = Path(ref.path)
            if path.is_absolute():
                report("artifact_path_invalid", "Artifact path must be relative to --step1-root.", entry, ref)
                continue
            resolved = (root / path).resolve()
            try:
                relative = resolved.relative_to(root)
            except ValueError:
                report("artifact_path_invalid", "Artifact path resolves outside --step1-root.", entry, ref)
                continue
            if not resolved.is_file():
                report("artifact_missing", f"Referenced artifact does not exist: {relative.as_posix()}", entry, ref)
                continue
            if _hash_file(resolved) != ref.sha256:
                payload = None
                if resolved.suffix.lower() == ".json":
                    try:
                        payload = _read_json(resolved)
                    except (OSError, ValueError, json.JSONDecodeError):
                        pass
                report("artifact_checksum_mismatch", f"SHA-256 does not match for {relative.as_posix()}.", entry, ref, payload)
                continue
            if resolved.suffix.lower() != ".json":
                continue
            try:
                payload = _read_json(resolved)
            except (OSError, ValueError, json.JSONDecodeError) as exc:
                report("artifact_json_invalid", f"Invalid JSON in {relative.as_posix()}: {exc}", entry, ref)
                continue
            artifact_name = Path(ref.name).name
            model = _MODEL_BY_FILE.get(artifact_name) or _MODEL_BY_FILE.get(resolved.name)
            handoff_reference = artifact_name == "handoff.json" or resolved.name == "handoff.json"
            current_handoff = handoff_reference and isinstance(payload, dict) and payload.get("schema_version") == "step1.v1"
            legacy_handoff = handoff_reference and isinstance(payload, dict) and payload.get("schema_version") == SCHEMA_VERSION and payload.get("artifact_type") == "handoff"
            if handoff_reference and current_handoff:
                try:
                    _validate_source_handoff(payload)
                except ValueError as exc:
                    report("artifact_schema_invalid", f"Invalid step1.v1 handoff schema: {exc}", entry, ref, payload)
                    continue
            elif handoff_reference and not legacy_handoff:
                report("artifact_schema_invalid", f"{ref.name} must use step1.v1 or the Phase 1 handoff contract.", entry, ref, payload)
                continue
            elif model:
                try:
                    model.model_validate(payload)
                except ValueError as exc:
                    report("artifact_schema_invalid", f"Invalid {ref.name} schema: {exc}", entry, ref, payload)
                    continue
            if isinstance(payload, dict):
                checks: list[tuple[Any, Any, str]] = []
                if artifact_name == "workbook_manifest.json" or resolved.name == "workbook_manifest.json":
                    checks.append((payload.get("sha256"), entry.source_sha256, "source SHA-256"))
                elif artifact_name == "source.json" or resolved.name == "source.json":
                    checks.extend([
                        (payload.get("sha256"), entry.source_sha256, "source SHA-256"),
                        (payload.get("relative_path"), entry.source_path, "source path"),
                        (payload.get("run_id"), entry.run_id, "run ID"),
                    ])
                elif artifact_name == "quality.json" or resolved.name == "quality.json":
                    checks.append((payload.get("source_sha256"), entry.source_sha256, "source SHA-256"))
                    checks.append((payload.get("run_id"), entry.run_id, "run ID"))
                elif current_handoff:
                    source = payload.get("source") if isinstance(payload.get("source"), dict) else {}
                    checks.extend([
                        (source.get("path"), entry.source_path, "source path"),
                        (source.get("sha256"), entry.source_sha256, "source SHA-256"),
                        (payload.get("run_id"), entry.run_id, "run ID"),
                    ])
                elif model:
                    if "workbook_sha256" in model.model_fields:
                        checks.append((payload.get("workbook_sha256"), entry.source_sha256, "source SHA-256"))
                    if "run_id" in model.model_fields:
                        checks.append((payload.get("run_id"), entry.run_id, "run ID"))
                for actual, expected, label in checks:
                    if isinstance(actual, str) and expected is not None and actual != expected:
                        report(
                            "artifact_identity_mismatch",
                            f"Artifact {label} {actual!r} does not match source entry {expected!r}.",
                            entry,
                            ref,
                            payload,
                        )

    for item in result["diagnostics"]:
        if item.get("next_tool"):
            rebuilders.add(item["next_tool"])
    if result["diagnostics"]:
        result["status"] = "blocked"
        if len(rebuilders) == 1 and all(item.get("next_tool") for item in result["diagnostics"]):
            result["next_tool"] = rebuilders.pop()
    return result


def build_index(handoff_path: Path, step1_root: Path, out_dir: Path = _DEFAULT_OUT) -> dict[str, Any]:
    """Index one ``step1.v1`` or ``step1.batch.v1`` handoff without copying artifacts."""
    root = step1_root.expanduser().resolve()
    if not root.is_dir():
        return _error_result(handoff_path, "step1_root_invalid", f"Step 1 output root does not exist: {root}")
    try:
        input_file, input_relative = _top_handoff_path(handoff_path, root)
        payload = _read_json(input_file)
        if not isinstance(payload, dict):
            raise ValueError("handoff must be a JSON object")
        schema = payload.get("schema_version")
        if schema == "step1.batch.v1":
            raw_entries = payload.get("entries")
            if not isinstance(raw_entries, list):
                raise ValueError("step1.batch.v1 handoff entries must be a list")
            entries = [_resolve_batch_entry(root, item, ordinal) for ordinal, item in enumerate(raw_entries, 1)]
        elif schema == "step1.v1":
            entries = [_entry_from_handoff(root, _validate_source_handoff(payload), input_relative, 1)]
        else:
            raise ValueError(f"unsupported handoff schema_version: {schema!r}")

        status = "partial" if entries else "blocked"
        entry_statuses = [entry.status for entry in entries]
        metrics = {
            "input_count": len(entries),
            "passed": sum(value == "pass" for value in entry_statuses),
            "partial": sum(value == "partial" for value in entry_statuses),
            "failed": sum(value in {"fail", "error", "blocked"} for value in entry_statuses),
            "ready": sum(entry.ready_for_next_step for entry in entries),
            "artifact_count": sum(len(entry.artifacts) for entry in entries),
        }
        index_diagnostics = [_diagnostic("batch_empty", "Batch handoff contains no source entries.")] if not entries else []
        index = Step2Index(
            step1_root=str(root),
            input_handoff_path=input_relative,
            input_handoff_sha256=_hash_file(input_file),
            batch_id=_text(payload.get("batch_id")),
            status=status,
            entries=entries,
            diagnostics=index_diagnostics,
            metrics=metrics,
        )
        out_dir = out_dir.expanduser().resolve()
        index_file, summary_file = _write_index(out_dir, index)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return _error_result(handoff_path, "handoff_invalid", f"Cannot build Step 2 index: {exc}")

    validation = validate_saved_index(index_file, root)
    index.validation_status = validation["status"]
    index.diagnostics.extend(item for item in validation["diagnostics"] if item.get("category") != "step2_artifact_reference")
    _write_index(out_dir, index)
    validation.update({
        "tool": "step2.index",
        "status": "blocked" if validation["status"] == "blocked" else status,
        "source": {"handoff": input_relative, "sha256": index.input_handoff_sha256},
        "artifacts": [
            {"name": _INDEX_NAME, "path": str(index_file), "sha256": _hash_file(index_file)},
            {"name": _SUMMARY_NAME, "path": str(summary_file), "sha256": _hash_file(summary_file)},
        ],
        "metrics": {**metrics, **validation["metrics"]},
        "diagnostics": index_diagnostics + [item for entry in entries for item in entry.diagnostics] + [
            item for item in validation["diagnostics"] if item.get("category") != "step2_artifact_reference"
        ],
    })
    return validation


_TOOLS: list[dict[str, Any]] = [
    {
        "name": "step2.handoff.resolve",
        "command": "step2 index --handoff PATH --step1-root DIR [--out DIR]",
        "inputs": {"handoff": "step1.v1 or step1.batch.v1 JSON", "step1_root": "Step 1 output root"},
        "outputs": ["normalized source entries", "handoff and artifact references", "input diagnostics"],
        "next": ["step2.index.build"],
    },
    {
        "name": "step2.index.build",
        "command": "step2 index --handoff PATH --step1-root DIR [--out DIR]",
        "inputs": {"entries": "resolved Step 1 source entries", "out": "index output directory"},
        "outputs": ["index.json", "INDEX.md", "per-entry quality context and diagnostics"],
        "next": [],
    },
    {
        "name": "step2.index.validate",
        "command": "step2 validate --index PATH --step1-root DIR",
        "inputs": {"index": "saved Step 2 index.json", "step1_root": "Step 1 output root"},
        "outputs": ["saved-index integrity status and diagnostics"],
        "next": [],
    },
]


def tool_catalog() -> dict[str, Any]:
    """Return the initial machine-readable Step 2 actions."""
    return {
        "tool": "step2.tools",
        "status": "ok",
        "source": None,
        "run_id": None,
        "artifacts": [],
        "metrics": {"tool_count": len(_TOOLS)},
        "diagnostics": [],
        "retryable": False,
        "next_tool": None,
        "tools": _TOOLS,
    }
