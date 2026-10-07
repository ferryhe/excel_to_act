"""Build a navigation and quality index from current Step 1 handoffs."""

from __future__ import annotations

import hashlib
import json
import os
from collections import Counter
from collections.abc import Callable
from pathlib import Path
from typing import Any

from excel_to_act.schemas import SCHEMA_VERSION, Step2ArtifactRef, Step2Index, Step2IndexEntry
from excel_to_act.steps.step1.workflow import _run_source
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


def _next_actions(value: Any, metrics: Any, diagnostics: list[dict[str, Any]]) -> list[dict[str, Any]]:
    actions = _items(value)
    if not any(action.get("name") == "source.review" or str(action.get("name", "")).startswith(("human.", "specialist.")) for action in actions) and (
        (isinstance(metrics, dict) and (metrics.get("opaque_parts") or metrics.get("opaque_rate")))
        or any("opaque" in str(item.get("code", "")) for item in diagnostics)
    ):
        actions.append({"name": "specialist.review", "reason": "Human or specialist review of preserved opaque content is required."})
    return actions


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


def _validate_batch_handoff(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict) or value.get("schema_version") != "step1.batch.v1":
        raise ValueError("batch handoff must use schema_version step1.batch.v1")
    if not isinstance(value.get("entries"), list):
        raise ValueError("step1.batch.v1 handoff entries must be a list")
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
    next_actions = _next_actions(handoff.get("next_actions"), metrics, diagnostics)
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
        next_actions=_next_actions(batch_entry.get("next_actions"), batch_entry.get("metrics"), diagnostics),
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


def _markdown(index: Step2Index, output_dir: Path | None = None) -> str:
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
        for ref in entry.artifacts:
            if ref.name in {"controls_handoff.md", "vba_handoff.md"}:
                target = str(Path(index.step1_root) / ref.path)
                if output_dir is not None:
                    try:
                        target = os.path.relpath(target, output_dir)
                    except ValueError:
                        pass
                lines.append(f"- Human control/VBA handoff: [{cell(ref.name)}](<{target.replace(chr(92), '/')}>).")
        if entry.metrics:
            lines.append(f"- Metrics: `{cell(json.dumps(entry.metrics, ensure_ascii=False, sort_keys=True))}`")
        lines.extend(f"- Blocker: {cell(blocker)}" for blocker in entry.blockers)
        lines.extend(f"- Diagnostic: {cell(item.get('code', 'unknown'))} — {cell(item.get('message', ''))}" for item in entry.diagnostics)
        lines.extend(f"- Next action: {cell(json.dumps(action, ensure_ascii=False, sort_keys=True))}" for action in entry.next_actions)
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
    summary_path.write_text(_markdown(index, out_dir.resolve()), encoding="utf-8")
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


def _validate_saved_index(
    index_path: Path, step1_root: Path, skip_entry_keys: set[str] | None = None,
    *, entry_keys: list[str] | None = None,
    on_entry_validated: Callable[[str, Step2IndexEntry, bool], None] | None = None,
    validated_artifacts: dict[tuple[str, str], tuple[Path, Any]] | None = None,
) -> dict[str, Any]:
    """Check a saved index and every referenced artifact using Step 1 contracts."""
    root = step1_root.expanduser().resolve()
    result = {
        "tool": "step2.index.validate",
        "status": "pass",
        "source": {"index": str(index_path)},
        "run_id": None,
        "artifacts": [],
        "metrics": {"entries_checked": 0, "artifacts_checked": 0, "input_handoff_checked": 0, "inventory_models_deserialized": 0},
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
    keys = entry_keys if entry_keys is not None else [str(position) for position in range(len(index.entries))]
    checked_entries = [(key, entry) for key, entry in zip(keys, index.entries) if key not in (skip_entry_keys or set())]
    result["metrics"]["entries_checked"] = len(checked_entries)

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
                    payload = _read_json(resolved_handoff)
                except (OSError, ValueError, json.JSONDecodeError) as exc:
                    input_handoff_error("input_handoff_json_invalid", f"Invalid JSON in input handoff {relative_handoff.as_posix()}: {exc}")
                else:
                    try:
                        schema = payload.get("schema_version") if isinstance(payload, dict) else None
                        if schema == "step1.v1":
                            _validate_source_handoff(payload)
                        elif schema == "step1.batch.v1":
                            _validate_batch_handoff(payload)
                        else:
                            raise ValueError(f"unsupported handoff schema_version: {schema!r}")
                    except ValueError as exc:
                        input_handoff_error("input_handoff_schema_invalid", f"Invalid Step 1 handoff schema: {exc}")
                    else:
                        result["metrics"]["input_handoff_checked"] = 1
    rebuilders: set[str] = set()
    input_valid = not result["diagnostics"]
    for key, entry in checked_entries:
        diagnostics_before = len(result["diagnostics"])
        for item in entry.diagnostics:
            if item.get("category") == "step2_artifact_reference":
                result["diagnostics"].append({**item, "source_id": entry.source_id, "source_path": entry.source_path})
        if entry.handoff_path and entry.handoff_path != index.input_handoff_path and not any(
            item.get("code") in {"source_handoff_unavailable", "source_handoff_invalid", "source_handoff_mismatch"}
            for item in entry.diagnostics
        ):
            handoff_ref = {"name": "source_handoff", "path": entry.handoff_path}

            def handoff_error(code: str, message: str) -> None:
                diagnostic = _diagnostic(code, message, source_id=entry.source_id)
                diagnostic.update({"source_path": entry.source_path, "artifact": handoff_ref["name"], "path": handoff_ref["path"]})
                result["diagnostics"].append(diagnostic)

            source_handoff = Path(entry.handoff_path)
            if source_handoff.is_absolute():
                handoff_error("source_handoff_path_invalid", "Per-source handoff path must be relative to --step1-root.")
            else:
                resolved_source_handoff = (root / source_handoff).resolve()
                try:
                    relative_source_handoff = resolved_source_handoff.relative_to(root)
                except ValueError:
                    handoff_error("source_handoff_path_invalid", "Per-source handoff path resolves outside --step1-root.")
                else:
                    if not resolved_source_handoff.is_file():
                        handoff_error("source_handoff_missing", f"Per-source handoff does not exist: {relative_source_handoff.as_posix()}")
                    else:
                        try:
                            source_payload = _validate_source_handoff(_read_json(resolved_source_handoff))
                        except (OSError, ValueError, json.JSONDecodeError) as exc:
                            handoff_error("source_handoff_schema_invalid", f"Cannot validate per-source handoff: {exc}")
                        else:
                            mismatch = _handoff_identity_mismatch({
                                "relative_path": entry.source_path,
                                "sha256": entry.source_sha256,
                                "run_id": entry.run_id,
                            }, source_payload)
                            if mismatch:
                                handoff_error("source_handoff_identity_mismatch", mismatch)
        for ref in entry.artifacts:
            artifact_diagnostics_before = len(result["diagnostics"])
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
            validated_payload = payload
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
                    if artifact_name == "inventory.json":
                        result["metrics"]["inventory_models_deserialized"] += 1
                    validated_payload = model.model_validate(payload)
                except ValueError as exc:
                    report("artifact_schema_invalid", f"Invalid {ref.name} schema: {exc}", entry, ref, payload)
                    continue
            if isinstance(payload, dict):
                checks: list[tuple[Any, Any, str]] = []
                if artifact_name == "workbook_manifest.json" or resolved.name == "workbook_manifest.json":
                    checks.append((payload.get("sha256"), entry.source_sha256, "source SHA-256"))
                elif artifact_name == "source.json" or resolved.name == "source.json":
                    if payload.get("schema_version") != "step1.v1":
                        report("artifact_schema_invalid", "source.json must use schema_version step1.v1.", entry, ref, payload)
                        continue
                    try:
                        _run_source(resolved.parent)
                    except (OSError, ValueError, json.JSONDecodeError) as exc:
                        report("artifact_schema_invalid", f"Invalid source.json schema: {exc}", entry, ref, payload)
                        continue

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
            if (validated_artifacts is not None
                    and len(result["diagnostics"]) == artifact_diagnostics_before):
                validated_artifacts[(entry.source_id, artifact_name)] = (resolved, validated_payload)
        if on_entry_validated is not None:
            on_entry_validated(key, entry, input_valid and len(result["diagnostics"]) == diagnostics_before)

    for item in result["diagnostics"]:
        if item.get("next_tool"):
            rebuilders.add(item["next_tool"])
    if result["diagnostics"]:
        result["status"] = "blocked"
        if len(rebuilders) == 1 and all(item.get("next_tool") for item in result["diagnostics"]):
            result["next_tool"] = rebuilders.pop()
    result["retryable"] = result["next_tool"] is not None
    return result


def _build_index(handoff_path: Path, step1_root: Path, out_dir: Path, state: dict[str, Any], resume: bool = False) -> dict[str, Any]:
    """Index one ``step1.v1`` or ``step1.batch.v1`` handoff without copying artifacts."""
    root = step1_root.expanduser().resolve()
    if not root.is_dir():
        return _error_result(handoff_path, "step1_root_invalid", f"Step 1 output root does not exist: {root}")
    try:
        input_file, input_relative = _top_handoff_path(handoff_path, root)
        payload, entries, entry_states, skipped = _load_entries(input_file, input_relative, root, out_dir, state, resume)

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

    def record_entry(entry: Step2IndexEntry, saved: dict[str, Any], valid: bool) -> None:
        saved.update({
            "output_sha256": _hash_value(entry.model_dump(mode="json")),
            "validation_status": "pass" if valid and not any(
                item.get("code") in {"source_handoff_unavailable", "source_handoff_invalid", "source_handoff_mismatch"}
                for item in entry.diagnostics
            ) else "blocked",
        })

    def checkpoint(key: str, entry: Step2IndexEntry, valid: bool) -> None:
        saved = entry_states[key]
        record_entry(entry, saved, valid)
        state["entries"][key] = saved
        # ponytail: rewrite state per entry; use a journal if large batches make this costly.
        _write_state(out_dir, state)

    validation = _validate_saved_index(index_file, root, skipped, entry_keys=list(entry_states), on_entry_validated=checkpoint)
    index.validation_status = validation["status"]
    index.diagnostics.extend(item for item in validation["diagnostics"] if item.get("category") != "step2_artifact_reference")
    _write_index(out_dir, index)
    input_valid = not any(item.get("artifact") == "input_handoff" for item in validation["diagnostics"])
    for entry, (key, saved) in zip(entries, entry_states.items()):
        if key in skipped:
            record_entry(entry, saved, input_valid)
    state.update({"entries": entry_states, "input_handoff_sha256": index.input_handoff_sha256,
                  "output_hashes": {_INDEX_NAME: _hash_file(index_file), _SUMMARY_NAME: _hash_file(summary_file)}})
    state["validation_input_sha256"] = _action_input("step2.index.validate", root, None, index_file)
    state["validation_result"] = dict(validation)
    validation.update({
        "tool": "step2.index",
        "status": "blocked" if validation["status"] == "blocked" else status,
        "source": {"handoff": input_relative, "sha256": index.input_handoff_sha256},
        "artifacts": [
            {"name": _INDEX_NAME, "path": str(index_file), "sha256": _hash_file(index_file)},
            {"name": _SUMMARY_NAME, "path": str(summary_file), "sha256": _hash_file(summary_file)},
        ],
        "metrics": {**metrics, **validation["metrics"], "entries_skipped": len(skipped), "entries_processed": len(entries) - len(skipped)},
        "diagnostics": index_diagnostics + [item for entry in entries for item in entry.diagnostics] + [
            item for item in validation["diagnostics"] if item.get("category") != "step2_artifact_reference"
        ],
    })
    return validation


def _hash_value(value: Any) -> str:
    return hashlib.sha256(_json_bytes(value)).hexdigest()


def _file_state(path: Path) -> str:
    try:
        return _hash_file(path)
    except OSError as exc:
        return f"unavailable:{type(exc).__name__}"


def _entry_fingerprint(root: Path, item: Any, handoff_relative: str) -> str:
    """Hash one source and its current references, independent of batch siblings."""
    evidence: list[Any] = [item]
    try:
        raw_path = item.get("handoff_path") if isinstance(item, dict) else None
        path, relative = _root_relative(root, raw_path or handoff_relative)
        evidence.append([relative, _file_state(path)])
        handoff = _read_json(path)
        if isinstance(handoff, dict):
            base = handoff.get("artifact_paths_relative_to") or handoff.get("run_path")
            for ref in _items(handoff.get("artifacts")):
                try:
                    artifact = _artifact_path(root, base, ref["path"])
                    evidence.append([artifact, _file_state(root / artifact)])
                except (KeyError, TypeError, ValueError) as exc:
                    evidence.append([ref, str(exc)])
    except (OSError, TypeError, ValueError) as exc:
        evidence.append(str(exc))
    return _hash_value(evidence)


def _input_items(path: Path) -> tuple[dict[str, Any], list[Any]]:
    payload = _read_json(path)
    schema = payload.get("schema_version") if isinstance(payload, dict) else None
    if schema == "step1.batch.v1":
        return _validate_batch_handoff(payload), payload["entries"]
    if schema == "step1.v1":
        return _validate_source_handoff(payload), [payload]
    raise ValueError(f"unsupported handoff schema_version: {schema!r}")


def _outputs_match(out_dir: Path, state: dict[str, Any]) -> bool:
    hashes = state.get("output_hashes", {})
    return all(name in hashes and _file_state(out_dir / name) == hashes[name] for name in (_INDEX_NAME, _SUMMARY_NAME))


def _load_entries(
    input_file: Path, input_relative: str, root: Path, out_dir: Path,
    state: dict[str, Any], resume: bool, *, cache_only: bool = False,
) -> tuple[dict[str, Any], list[Step2IndexEntry], dict[str, Any], set[str]]:
    payload, items = _input_items(input_file)
    cached: dict[str, Step2IndexEntry] = {}
    attempts = state.get("attempts", [])
    unfinished = bool(attempts and attempts[-1]["tool"] == "step2.index.build" and _interrupted(attempts[-1]))
    if resume:
        try:
            old = Step2Index.model_validate(_read_json(out_dir / _INDEX_NAME))
            if (old.step1_root == str(root) and old.input_handoff_path == input_relative
                    and (_outputs_match(out_dir, state) or (
                        unfinished and (out_dir / _SUMMARY_NAME).read_text(encoding="utf-8") == _markdown(old, out_dir.resolve())
                    ))):
                cached = {_hash_value(entry.model_dump(mode="json")): entry for entry in old.entries}
        except (OSError, ValueError):
            pass
    entries, saved_entries, skipped = [], {}, set()
    occurrences: dict[str, int] = {}
    for ordinal, item in enumerate(items, 1):
        if payload["schema_version"] == "step1.batch.v1":
            location = [_text(item.get(name)) for name in ("relative_path", "handoff_path")] if isinstance(item, dict) else [None, None]
            # Count duplicates within their identity, so unrelated siblings cannot shift keys.
            identity = _hash_value(["location", *location] if any(location) else ["unidentified", item])
            occurrences[identity] = occurrences.get(identity, 0) + 1
            key = _hash_value([identity, occurrences[identity]])
        else:
            key = _hash_value([input_relative, ordinal])
        fingerprint = _entry_fingerprint(root, item, input_relative)
        previous = state.get("entries", {}).get(key, {})
        entry = cached.get(previous.get("output_sha256"))
        if (entry is not None and previous.get("validation_status") == "pass"
                and previous.get("input_sha256") == fingerprint
                and previous.get("output_sha256") == _hash_value(entry.model_dump(mode="json"))):
            skipped.add(key)
        elif cache_only:
            return payload, [], {}, set()
        elif payload["schema_version"] == "step1.batch.v1":
            entry = _resolve_batch_entry(root, item, ordinal)
        else:
            entry = _entry_from_handoff(root, payload, input_relative, ordinal)
        entries.append(entry)
        saved_entries[key] = {"source_id": entry.source_id, "input_sha256": fingerprint}
    return payload, entries, saved_entries, skipped


def _read_state(out_dir: Path, root: Path) -> dict[str, Any]:
    path = out_dir / "state.json"
    if not path.exists():
        return {"schema_version": "step2.state.v1", "step1_root": str(root), "attempt_limit": 3,
                "attempts": [], "entries": {}, "output_hashes": {}}
    state = _read_json(path)
    if (not isinstance(state, dict) or state.get("schema_version") != "step2.state.v1"
            or state.get("step1_root") != str(root) or state.get("attempt_limit") != 3
            or not isinstance(state.get("attempts"), list)
            or not isinstance(state.get("entries"), dict) or not isinstance(state.get("output_hashes"), dict)
            or not isinstance(state.get("validated_entry_hashes", []), list)
            or any(not isinstance(value, str) for value in state.get("validated_entry_hashes", []))
            or not isinstance(state.get("validation_result", {}), dict)):
        raise ValueError("Invalid Step 2 state or a different Step 1 root; preserve state.json and use the matching index directory.")
    known = {tool["name"] for tool in _TOOLS}
    for attempt in state["attempts"]:
        if (not isinstance(attempt, dict) or attempt.get("tool") not in known
                or not isinstance(attempt.get("input_sha256"), str)
                or type(attempt.get("no_progress_count")) is not int
                or attempt["no_progress_count"] < 0 or not isinstance(attempt.get("result"), dict)
                or not isinstance(attempt.get("input_paths", {}), dict)
                or not isinstance(attempt["result"].get("diagnostics"), list)
                or any(not isinstance(item, dict) for item in attempt["result"]["diagnostics"])
                or not isinstance(attempt["result"].get("metrics"), dict)):
            raise ValueError("Invalid Step 2 attempt history; preserve state.json.")
    if any(not isinstance(value, dict) for value in state["entries"].values()):
        raise ValueError("Invalid Step 2 entry state; preserve state.json.")
    return state


def _action_input(name: str, root: Path, handoff: Path | None, index: Path | None) -> str:
    path = index if name == "step2.index.validate" else handoff
    evidence: list[Any] = [str(root)]
    if path is not None:
        try:
            if name == "step2.index.validate":
                path = path.expanduser().resolve()
                payload = Step2Index.model_validate(_read_json(path))
                input_handoff, _ = _root_relative(root, payload.input_handoff_path)
                evidence.extend([str(path), _file_state(path), _file_state(input_handoff)])
                for entry in payload.entries:
                    evidence.append(_entry_fingerprint(root, {"handoff_path": entry.handoff_path}, payload.input_handoff_path))
                    evidence.extend([ref.path, _file_state(_root_relative(root, ref.path)[0])] for ref in entry.artifacts)
            else:
                path, relative = _top_handoff_path(path, root)
                _, items = _input_items(path)
                evidence.extend([str(path), _file_state(path), *[_entry_fingerprint(root, item, relative) for item in items]])
        except (OSError, TypeError, ValueError) as exc:
            evidence.extend([str(path), _file_state(path), str(exc)])
    return _hash_value(evidence)


def _write_state(out_dir: Path, state: dict[str, Any]) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    temporary = out_dir / "state.json.tmp"
    temporary.write_bytes(_json_bytes(state))
    temporary.replace(out_dir / "state.json")


def _stopped(result: dict[str, Any], name: str, reason: str, attempts: list[Any]) -> dict[str, Any]:
    return {**result, "tool": name, "status": "blocked", "retryable": False, "next_tool": None,
            "recovery_stopped": True, "stop_reason": reason, "attempt": len(attempts), "attempt_limit": 3}


def _interrupted(attempt: dict[str, Any]) -> bool:
    return any(item.get("code") == "attempt_interrupted" for item in attempt["result"]["diagnostics"])


def _resolved_entry_evidence(result: dict[str, Any], keys: list[str]) -> dict[str, Any]:
    return {
        identity: {key: value for key, value in entry.items() if key != "diagnostics"}
        for identity, entry in zip(keys, result.get("entries", [])) if not any(
            item.get("code") in {"source_handoff_unavailable", "source_handoff_invalid", "source_handoff_mismatch"}
            for item in entry["diagnostics"]
        )
    }


def execute_tool(
    name: str, *, step1_root: Path, handoff_path: Path | None = None,
    index_path: Path | None = None, out_dir: Path = _DEFAULT_OUT, resume: bool = False,
) -> dict[str, Any]:
    """Dispatch one catalogue action; the host chooses all subsequent calls."""
    root = step1_root.expanduser().resolve()
    out_dir = (index_path.parent if name == "step2.index.validate" and index_path is not None else out_dir).expanduser().resolve()
    if name not in {tool["name"] for tool in _TOOLS}:
        return _error_result(handoff_path or index_path or out_dir, "unknown_tool", f"Unknown Step 2 tool: {name}") | {"tool": name}
    try:
        state = _read_state(out_dir, root)
    except (OSError, ValueError) as exc:
        return _stopped(_error_result(out_dir, "state_invalid", str(exc)), name, "state_invalid", [])
    attempts = state["attempts"]
    required = "index" if name == "step2.index.validate" else "handoff"
    if (index_path if required == "index" else handoff_path) is None:
        return _stopped(_error_result(out_dir, "handoff_invalid", f"{name} requires --{required}"), name, "non_retryable_failure", attempts)
    fingerprint = _action_input(name, root, handoff_path, index_path)
    base_fingerprint = fingerprint
    if (name == "step2.index.validate" and state.get("validation_input_sha256") == fingerprint
            and state.get("validation_result", {}).get("status") == "pass"):
        result = state["validation_result"]
        return {**result, "tool": name, "metrics": {**result["metrics"], "entries_checked": 0, "artifacts_checked": 0,
                "entries_skipped": result["metrics"].get("entries_checked", 0)},
                "attempt": len(attempts), "attempt_limit": 3, "recovery_stopped": False, "stop_reason": None}
    # An unchanged validated resume is a cache read, not another tool attempt.
    if resume and name == "step2.index.build" and handoff_path is not None:
        try:
            input_file, relative = _top_handoff_path(handoff_path, root)
            _, entries, _, skipped = _load_entries(input_file, relative, root, out_dir, state, True, cache_only=True)
            if (entries and len(skipped) == len(entries) and _outputs_match(out_dir, state)
                    and state.get("input_handoff_sha256") == _hash_file(input_file)):
                result = next(a["result"] for a in reversed(attempts) if a["tool"] == name and not _interrupted(a)
                              and {ref["name"]: ref["sha256"] for ref in a["result"].get("artifacts", [])} == state["output_hashes"])
                return {**result, "status": "partial", "retryable": False, "next_tool": None,
                        "metrics": {**result["metrics"], "entries_skipped": len(skipped), "entries_processed": 0, "entries_checked": 0, "artifacts_checked": 0},
                        "attempt": len(attempts), "attempt_limit": 3, "recovery_stopped": False, "stop_reason": None}
        except (OSError, ValueError, StopIteration):
            pass
        if state.get("output_hashes") and not _outputs_match(out_dir, state):
            fingerprint = _hash_value([fingerprint, {_INDEX_NAME: _file_state(out_dir / _INDEX_NAME), _SUMMARY_NAME: _file_state(out_dir / _SUMMARY_NAME)}])
    last = attempts[-1] if attempts else {}
    last_completed = next((a for a in reversed(attempts) if not _interrupted(a)), {})
    failed_input_unchanged = False
    if last.get("result", {}).get("stop_reason") == "non_retryable_failure":
        paths = last.get("input_paths", {})
        failed_input_unchanged = last.get("base_input_sha256") == _action_input(
            last["tool"], root,
            handoff_path if handoff_path is not None else Path(paths["handoff"]) if paths.get("handoff") else None,
            index_path if index_path is not None else Path(paths["index"]) if paths.get("index") else None,
        )
    reason = (
        "attempt_limit_reached" if len(attempts) >= 3 else
        "no_progress" if last_completed.get("no_progress_count", 0) >= 2 else
        "non_retryable_failure" if failed_input_unchanged else
        "repeated_tool_input" if any(a["tool"] == name and a["input_sha256"] == fingerprint and not (resume and _interrupted(a)) for a in attempts) else None
    )
    if reason:
        return _stopped(last["result"], name, reason, attempts)
    validated_before = {key: entry.get("output_sha256") for key, entry in state["entries"].items() if entry.get("validation_status") == "pass"}
    validation_before = Counter(state.get("validated_entry_hashes", []))
    attempt = {"tool": name, "input_sha256": fingerprint, "base_input_sha256": base_fingerprint,
               "input_paths": {"handoff": str((handoff_path if handoff_path.is_absolute() or handoff_path.is_file() else root / handoff_path).resolve()) if handoff_path else None, "index": str(index_path.resolve()) if index_path else None},
               "evidence_sha256": last_completed.get("evidence_sha256"),
               "improved": False, "no_progress_count": last_completed.get("no_progress_count", 0),
               "result": _error_result(out_dir, "attempt_interrupted", "Step 2 action started; no completed result was saved.")}
    # ponytail: host dispatches sequentially; add an output-directory lock if concurrent calls are required.
    attempts.append(attempt)
    try:
        _write_state(out_dir, state)
    except OSError as exc:
        attempts.pop()
        failure = _error_result(out_dir, "state_write_failed", f"Cannot reserve Step 2 attempt: {exc}")
        if last:
            failure = {**last["result"], "diagnostics": last["result"]["diagnostics"] + failure["diagnostics"]}
        return _stopped(failure, name, "non_retryable_failure", attempts)
    try:
        if name == "step2.index.validate":
            validated_hashes = []

            def validated_entry(_key: str, entry: Step2IndexEntry, valid: bool) -> None:
                if valid and not any(
                    item.get("code") in {"source_handoff_unavailable", "source_handoff_invalid", "source_handoff_mismatch"}
                    for item in entry.diagnostics
                ):
                    validated_hashes.append(_hash_value(entry.model_dump(mode="json")))

            result = _validate_saved_index(index_path, root, on_entry_validated=validated_entry)
        elif name == "step2.index.build":
            result = _build_index(handoff_path, root, out_dir, state, resume)
        else:
            input_file, relative = _top_handoff_path(handoff_path, root)
            _, entries, entry_states, _ = _load_entries(input_file, relative, root, out_dir, state, False)
            result = {"tool": name, "status": "partial", "source": {"handoff": relative}, "run_id": None,
                      "artifacts": [], "metrics": {"input_count": len(entries)},
                      "diagnostics": [d for entry in entries for d in entry.diagnostics],
                      "retryable": False, "next_tool": "step2.index.build", "entries": [entry.model_dump(mode="json") for entry in entries]}
    except Exception as exc:
        result = _error_result(handoff_path or index_path or out_dir, "handoff_invalid", str(exc))
    result["tool"] = name
    # Compare substantive evidence, excluding dispatch/checked counters and paths.
    errors = [item for item in result["diagnostics"] if item.get("severity") != "warning"]
    evidence = _hash_value({"diagnostics": errors, "entries": {
        key: {field: entry.get(field) for field in ("source_id", "output_sha256", "validation_status")}
        for key, entry in state.get("entries", {}).items()
    }})
    previous_errors = sum(item.get("severity") != "warning" for item in last_completed.get("result", {}).get("diagnostics", []))
    successful_progress = result["status"] != "blocked" and (not last_completed or evidence != last_completed.get("evidence_sha256"))
    if name == "step2.index.build":
        validated = {key: entry.get("output_sha256") for key, entry in state["entries"].items() if entry.get("validation_status") == "pass"}
        evidence = _hash_value(validated)
        successful_progress = any(validated_before.get(key) != digest for key, digest in validated.items())
        state["validated_entry_hashes"] = sorted(validated.values())
    elif name == "step2.index.validate":
        state["validated_entry_hashes"] = sorted(validated_hashes)
        evidence = _hash_value(state["validated_entry_hashes"])
        successful_progress = bool(Counter(validated_hashes) - validation_before)
    elif name == "step2.handoff.resolve":
        resolved = _resolved_entry_evidence(result, list(entry_states)) if "entries" in result else {}
        previous = last_completed.get("resolved_entries", {})
        attempt["resolved_entries"] = resolved
        evidence = _hash_value(resolved)
        successful_progress = any(previous.get(key) != entry for key, entry in resolved.items())
    improved = successful_progress or (
        bool(last_completed) and len(errors) < previous_errors
    )
    no_progress = 0 if improved else last_completed.get("no_progress_count", 0) + 1
    result.update({"attempt": len(attempts), "attempt_limit": 3, "recovery_stopped": False, "stop_reason": None})
    if result["status"] == "blocked" and not result.get("retryable"):
        result = _stopped(result, name, "non_retryable_failure", attempts)
    elif no_progress >= 2:
        result = _stopped(result, name, "no_progress", attempts)
    elif len(attempts) >= 3:
        result.update({"recovery_stopped": True, "stop_reason": "attempt_limit_reached", "retryable": False, "next_tool": None})
    attempt.update({"evidence_sha256": evidence, "improved": improved, "no_progress_count": no_progress, "result": result})
    if name == "step2.index.validate":
        state.update({"validation_input_sha256": fingerprint, "validation_result": result})
    _write_state(out_dir, state)
    return result


def build_index(handoff_path: Path, step1_root: Path, out_dir: Path = _DEFAULT_OUT, *, resume: bool = False) -> dict[str, Any]:
    return execute_tool("step2.index.build", step1_root=step1_root, handoff_path=handoff_path, out_dir=out_dir, resume=resume)


def validate_saved_index(index_path: Path, step1_root: Path) -> dict[str, Any]:
    return execute_tool("step2.index.validate", step1_root=step1_root, index_path=index_path)


_TOOLS: list[dict[str, Any]] = [
    {
        "name": "step2.handoff.resolve",
        "command": "step2 tool step2.handoff.resolve --handoff PATH --step1-root DIR [--out DIR]",
        "inputs": {"handoff": "step1.v1 or step1.batch.v1 JSON; absolute or relative path inside step1_root", "step1_root": "Step 1 output root", "out": "state directory (default output/step2_index)"},
        "outputs": ["normalized source entries", "handoff and artifact references", "input diagnostics", "state.json", "structured status/metrics/retryable/next_tool"],
        "next": ["step2.index.build"],
    },
    {
        "name": "step2.index.build",
        "command": "step2 tool step2.index.build --handoff PATH --step1-root DIR [--out DIR] [--resume]",
        "inputs": {"handoff": "step1.v1 or step1.batch.v1 JSON inside step1_root", "step1_root": "Step 1 output root", "out": "index output directory (default output/step2_index)", "resume": "reuse validated unchanged entries"},
        "outputs": ["index.json", "INDEX.md", "state.json", "per-entry quality context and diagnostics", "structured status/metrics/retryable/next_tool"],
        "next": ["step2.index.validate"],
    },
    {
        "name": "step2.index.validate",
        "command": "step2 tool step2.index.validate --index PATH --step1-root DIR",
        "inputs": {"index": "saved Step 2 index.json; its parent directory holds Step 2 state", "step1_root": "Step 1 output root"},
        "outputs": ["saved-index integrity status and diagnostics", "structured status/metrics/retryable/next_tool", "state.json beside the index; actual dispatches persist shared attempt/validation state"],
        "next": [],
    },
]


def tool_catalog() -> dict[str, Any]:
    """Return Step 2 recovery tools and implemented reader commands."""
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
        "reader_commands": [{
            "name": "step2.prepare",
            "command": "step2 prepare --index INDEX --step1-root ROOT --out READING_DIR [--scope SCOPE] [--resume] [--dry-run]",
            "outputs": ["manifest.json", "scope.json", "dependency_audit.json", "retained_dependency_cells.json", "paired English handoffs", "source views", "dependency graphs"],
        }],
    }
