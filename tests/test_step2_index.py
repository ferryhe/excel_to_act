import hashlib
import json
from pathlib import Path

import pytest

from excel_to_act.steps.step1.workflow import _write_handoff
from excel_to_act.steps.step2.workflow import build_index, execute_tool, tool_catalog, validate_saved_index


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value), encoding="utf-8")


def source_handoff(**updates: object) -> dict:
    handoff = {
        "schema_version": "step1.v1",
        "run_path": "runs/source-a",
        "artifact_paths_relative_to": "runs/source-a/artifacts",
        "source": {"path": "models/a.xlsx", "sha256": "source-hash"},
        "run_id": "run-a",
        "batch_id": "batch-a",
        "status": "pass",
        "ready_for_next_step": True,
        "metrics_state": "current",
        "metrics": {"traceability_ratio": 1.0},
        "thresholds": {"traceability": 1.0},
        "artifacts": [{"name": "inventory", "path": "inventory.json", "sha256": "artifact-hash"}],
        "final_output": "final/batches/batch-a/source-a",
        "diagnostics": [],
        "blockers": [],
        "next_actions": [],
    }
    handoff.update(updates)
    return handoff


def source_metadata() -> dict:
    return {
        "schema_version": "step1.v1",
        "batch_id": "batch-a",
        "run_id": "run-a",
        "source_path": "C:/models/a.xlsx",
        "relative_path": "models/a.xlsx",
        "sha256": None,
        "suffix": ".xlsx",
        "thresholds": {},
        "allow_opaque": True,
        "attempt_limit": 3,
    }


def test_null_attempt_diagnostic_returns_preserved_state_invalid(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, out)
    state_file = out / "state.json"
    state = json.loads(state_file.read_text())
    state["attempts"][0]["result"]["diagnostics"] = [None]
    write_json(state_file, state)
    before = state_file.read_bytes()
    result = build_index(handoff, root, out)
    assert result["status"] == "blocked"
    assert result["stop_reason"] == "state_invalid"
    assert result["diagnostics"][0]["code"] == "state_invalid"
    assert state_file.read_bytes() == before


@pytest.mark.parametrize("name", ["step2.index.build", "step2.handoff.resolve", "step2.index.validate"])
def test_corrected_required_argument_dispatches_without_spending_invalid_attempt(tmp_path: Path, name: str) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    result = execute_tool(name, step1_root=root, out_dir=out)
    assert result["status"] == "blocked"
    assert result["stop_reason"] == "non_retryable_failure"
    assert result["attempt"] == 0
    assert not (out / "state.json").exists()
    if name == "step2.index.validate":
        built = tmp_path / "built"
        build_index(handoff, root, built)
        index = out / "index.json"
        write_json(index, json.loads((built / "index.json").read_text()))
        corrected = execute_tool(name, step1_root=root, index_path=index)
    else:
        corrected = execute_tool(name, step1_root=root, handoff_path=handoff, out_dir=out)
    assert corrected["status"] in {"pass", "partial"}
    assert corrected["stop_reason"] is None
    assert corrected["attempt"] == 1
    assert len(json.loads((out / "state.json").read_text())["attempts"]) == 1


def test_corrected_handoff_path_retries_with_shared_budget(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    bad, valid = root / "bad.json", root / "valid.json"
    write_json(bad, {"schema_version": "unsupported"})
    write_json(valid, source_handoff(artifacts=[]))
    failed = build_index(bad, root, out)
    assert failed["stop_reason"] == "non_retryable_failure"
    assert build_index(bad, root, out)["attempt"] == 1
    corrected = build_index(valid, root, out)
    assert corrected["status"] == "partial"
    assert corrected["stop_reason"] is None
    assert corrected["attempt"] == 2
    history = json.loads((out / "state.json").read_text())["attempts"]
    assert len(history) == 2
    assert history[0]["result"] == failed


def test_reservation_output_file_returns_structured_failure_without_dispatch(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "occupied"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    out.write_text("occupied", encoding="utf-8")
    original, calls = workflow._build_index, []

    def build(*args, **kwargs):
        calls.append(True)
        return original(*args, **kwargs)

    monkeypatch.setattr(workflow, "_build_index", build)
    result = build_index(handoff, root, out)
    assert result["status"] == "blocked"
    assert result["stop_reason"] == "non_retryable_failure"
    assert result["diagnostics"][0]["code"] == "state_write_failed"
    assert result["attempt"] == 0
    assert calls == []
    assert out.read_text() == "occupied"
    corrected = build_index(handoff, root, tmp_path / "valid-index")
    assert corrected["attempt"] == 1
    assert calls == [True]


def test_reservation_failure_preserves_previous_failure_evidence(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff())
    failed = build_index(handoff, root, out)
    assert failed["diagnostics"][0]["code"] == "artifact_missing"
    before = (out / "state.json").read_bytes()

    def cannot_reserve(*args, **kwargs):
        raise OSError("reservation unavailable")

    monkeypatch.setattr(workflow, "_write_state", cannot_reserve)
    result = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=handoff, out_dir=out)
    assert result["status"] == "blocked"
    assert result["stop_reason"] == "non_retryable_failure"
    assert result["attempt"] == 1
    assert result["artifacts"] == failed["artifacts"]
    assert result["metrics"] == failed["metrics"]
    assert result["diagnostics"][:-1] == failed["diagnostics"]
    assert result["diagnostics"][-1]["code"] == "state_write_failed"
    assert (out / "state.json").read_bytes() == before


def test_standalone_validation_tracks_current_passing_entry_progress(tmp_path: Path) -> None:
    root, built, out = tmp_path / "step1", tmp_path / "built", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, built)
    saved = json.loads((built / "index.json").read_text())
    index = out / "index.json"
    for revision in range(1, 4):
        saved["entries"][0]["metrics"]["revision"] = revision
        write_json(index, saved)
        result = validate_saved_index(index, root)
        assert result["status"] == "pass"
        assert result["diagnostics"] == []
        assert result["attempt"] == revision
        history = json.loads((out / "state.json").read_text())["attempts"]
        assert history[-1]["improved"] is True
        assert history[-1]["no_progress_count"] == 0
    assert result["stop_reason"] == "attempt_limit_reached"
    saved["entries"][0]["metrics"]["revision"] = 4
    write_json(index, saved)
    stopped = validate_saved_index(index, root)
    assert stopped["status"] == "blocked"
    assert stopped["stop_reason"] == "attempt_limit_reached"
    assert stopped["attempt"] == 3
    assert stopped["diagnostics"] == []
    assert json.loads((out / "state.json").read_text())["entries"] == {}


def test_standalone_validation_failures_do_not_count_changed_error_wording_as_progress(tmp_path: Path) -> None:
    root, built, out = tmp_path / "step1", tmp_path / "built", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, built)
    saved = json.loads((built / "index.json").read_text())
    index = out / "index.json"
    write_json(index, saved)
    assert validate_saved_index(index, root)["status"] == "pass"
    failures = []
    for number in (1, 2):
        saved["entries"][0]["artifacts"] = [{"name": "inventory.json", "path": f"missing-{number}.json", "sha256": "missing"}]
        write_json(index, saved)
        result = validate_saved_index(index, root)
        failures.append(result)
        history = json.loads((out / "state.json").read_text())["attempts"]
        assert history[-1]["improved"] is False
        assert history[-1]["no_progress_count"] == number
        assert result["diagnostics"][0]["code"] == "artifact_missing"
    assert failures[1]["stop_reason"] == "no_progress"
    assert failures[0]["diagnostics"] != failures[1]["diagnostics"]
    assert json.loads((out / "state.json").read_text())["validated_entry_hashes"] == []


def test_resume_reuses_only_validated_unchanged_entries(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    entries = []
    for name in ("a", "b"):
        write_json(root / f"{name}/handoff.json", source_handoff(
            source={"path": f"{name}.xlsx", "sha256": name}, run_id=name, artifacts=[],
        ))
        entries.append({"relative_path": f"{name}.xlsx", "sha256": name, "run_id": name, "handoff_path": f"{name}/handoff.json"})
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": entries})
    build_index(batch, root, out)
    original = workflow._resolve_batch_entry
    calls = []

    def resolve(root, item, ordinal):
        calls.append(item["relative_path"])
        return original(root, item, ordinal)

    monkeypatch.setattr(workflow, "_resolve_batch_entry", resolve)
    unchanged = build_index(batch, root, out, resume=True)
    assert unchanged["metrics"]["entries_skipped"] == 2
    assert unchanged["metrics"]["entries_checked"] == 0
    assert calls == []
    assert len(json.loads((out / "state.json").read_text())["attempts"]) == 1

    write_json(root / "a/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[], metrics={"changed": True},
    ))
    changed = build_index(batch, root, out, resume=True)
    assert changed["metrics"]["entries_skipped"] == 1
    assert changed["metrics"]["entries_checked"] == 1
    assert calls == ["a.xlsx"]


@pytest.mark.parametrize("change", ["rename", "remove", "insert"])
def test_resume_reuses_entries_after_batch_positions_change(tmp_path: Path, monkeypatch, change: str) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    entries = {}
    for name in ("a", "b", "c", "z"):
        artifact = root / f"{name}/artifacts/custom.json"
        write_json(artifact, {"content": name})
        write_json(root / f"{name}/handoff.json", source_handoff(
            source={"path": f"{name}.xlsx", "sha256": name}, run_id=name,
            artifact_paths_relative_to=f"{name}/artifacts", artifacts=[{
                "name": "custom.json", "path": "custom.json",
                "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
            }],
        ))
        entries[name] = {"relative_path": f"{name}.xlsx", "handoff_path": f"{name}/handoff.json"}
    initial = ["b", "c"] if change == "insert" else ["a", "b", "c"]
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [entries[name] for name in initial]})
    assert build_index(batch, root, out)["attempt"] == 1
    original = workflow._resolve_batch_entry
    calls = []

    def resolve(root, item, ordinal):
        calls.append(item["relative_path"])
        return original(root, item, ordinal)

    monkeypatch.setattr(workflow, "_resolve_batch_entry", resolve)
    changed = {"rename": ["b", "c", "z"], "remove": ["b", "c"], "insert": ["a", "b", "c"]}[change]
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [entries[name] for name in changed]})
    resumed = build_index(batch, root, out, resume=True)
    processed = 0 if change == "remove" else 1
    assert calls == (["z.xlsx"] if change == "rename" else ["a.xlsx"] if change == "insert" else [])
    assert resumed["attempt"] == 2
    assert resumed["metrics"]["entries_skipped"] == 2
    assert resumed["metrics"]["entries_processed"] == processed
    assert resumed["metrics"]["entries_checked"] == processed
    assert resumed["metrics"]["artifacts_checked"] == processed
    assert len(json.loads((out / "state.json").read_text())["attempts"]) == 2


@pytest.mark.parametrize("item", [
    {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
    {"status": "blocked"},
])
def test_batch_duplicate_keys_survive_unrelated_insertion(tmp_path: Path, item: dict) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    write_json(root / "a/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[],
    ))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [item, item]})
    state = {"entries": {}, "attempts": []}
    _, entries, before, _ = workflow._load_entries(batch, "batch.json", root, out, state, False)
    assert len(entries) == len(before) == 2
    unrelated = {"relative_path": "b.xlsx", "handoff_path": "b/handoff.json"}
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [unrelated, item, item]})
    _, entries, after, _ = workflow._load_entries(batch, "batch.json", root, out, state, False)
    assert len(entries) == len(after) == 3
    assert set(before).issubset(after)
    if item.get("handoff_path"):
        assert [after[key]["input_sha256"] for key in before] == [saved["input_sha256"] for saved in before.values()]


def test_resume_reuses_unchanged_duplicate_batch_occurrences(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    write_json(root / "a/handoff.json", source_handoff(artifacts=[]))
    item = {"handoff_path": "a/handoff.json"}
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [item, item]})
    assert build_index(batch, root, out)["metrics"]["entries_checked"] == 2
    before = (out / "state.json").read_bytes()
    resumed = build_index(batch, root, out, resume=True)
    assert resumed["status"] == "partial"
    assert resumed["stop_reason"] is None
    assert resumed["attempt"] == 1
    assert resumed["metrics"]["entries_skipped"] == 2
    assert resumed["metrics"]["entries_processed"] == resumed["metrics"]["entries_checked"] == 0
    assert (out / "state.json").read_bytes() == before


@pytest.mark.parametrize("changed", ["first", "second"])
def test_resume_validates_changed_occurrence_with_duplicate_source_id(tmp_path: Path, monkeypatch, changed: str) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    items = []
    for name in ("first", "second"):
        artifact = root / f"{name}/custom.json"
        write_json(artifact, {"content": name})
        write_json(root / f"{name}/handoff.json", source_handoff(
            artifact_paths_relative_to=name, artifacts=[{
                "name": "custom.json", "path": "custom.json", "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
            }],
        ))
        items.append({"handoff_path": f"{name}/handoff.json"})
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": items})
    assert build_index(batch, root, out)["metrics"]["entries_checked"] == 2
    initial = json.loads((out / "index.json").read_text())["entries"]
    assert initial[0]["source_id"] == initial[1]["source_id"]
    (root / f"{changed}/custom.json").write_text("changed", encoding="utf-8")
    original, calls = workflow._resolve_batch_entry, []

    def resolve(root, item, ordinal):
        calls.append(item["handoff_path"])
        return original(root, item, ordinal)

    monkeypatch.setattr(workflow, "_resolve_batch_entry", resolve)
    resumed = build_index(batch, root, out, resume=True)
    assert calls == [f"{changed}/handoff.json"]
    assert resumed["status"] == "blocked"
    assert resumed["metrics"]["entries_skipped"] == 1
    assert resumed["metrics"]["entries_processed"] == resumed["metrics"]["entries_checked"] == 1
    assert resumed["metrics"]["artifacts_checked"] == 1
    assert [d["code"] for d in resumed["diagnostics"]] == ["artifact_checksum_mismatch"]
    assert resumed["diagnostics"][0]["path"] == f"{changed}/custom.json"
    states = {entry["output_sha256"]: entry["validation_status"] for entry in json.loads((out / "state.json").read_text())["entries"].values()}
    statuses = [states[workflow._hash_value(entry)] for entry in json.loads((out / "index.json").read_text())["entries"]]
    assert statuses == (["blocked", "pass"] if changed == "first" else ["pass", "blocked"])


def test_resume_validates_inserted_unidentified_occurrence(tmp_path: Path) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    write_json(root / "old/handoff.json", source_handoff(source={}, run_id=None, artifacts=[]))
    old = {"handoff_path": "old/handoff.json"}
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [old]})
    build_index(batch, root, out)
    write_json(root / "new/handoff.json", source_handoff(
        source={}, run_id=None, artifact_paths_relative_to="new", artifacts=[{
            "name": "custom.json", "path": "missing.json", "sha256": "missing",
        }],
    ))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [{"handoff_path": "new/handoff.json"}, old]})
    resumed = build_index(batch, root, out, resume=True)
    assert resumed["status"] == "blocked"
    assert resumed["metrics"]["entries_skipped"] == 1
    assert resumed["metrics"]["entries_processed"] == resumed["metrics"]["entries_checked"] == 1
    assert resumed["metrics"]["artifacts_checked"] == 1
    assert [d["code"] for d in resumed["diagnostics"]] == ["artifact_missing"]
    entries = json.loads((out / "index.json").read_text())["entries"]
    assert entries[0]["source_id"] == entries[1]["source_id"]
    states = {entry["output_sha256"]: entry["validation_status"] for entry in json.loads((out / "state.json").read_text())["entries"].values()}
    assert [states[workflow._hash_value(entry)] for entry in entries] == ["blocked", "pass"]


def test_resume_rechecks_changed_artifacts_and_saved_outputs(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    artifact = root / "runs/source-a/artifacts/custom.json"
    write_json(artifact, {"opaque": True})
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[{
        "name": "custom.json", "path": "custom.json", "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]))
    build_index(handoff, root, out)
    (out / "INDEX.md").write_text("changed", encoding="utf-8")
    rebuilt = build_index(handoff, root, out, resume=True)
    assert rebuilt["metrics"]["entries_skipped"] == 0
    write_json(artifact, {"opaque": "changed"})
    changed = build_index(handoff, root, out, resume=True)
    assert changed["metrics"]["entries_skipped"] == 0
    assert any(d["code"] == "artifact_checksum_mismatch" for d in changed["diagnostics"])


def test_resume_continues_interrupted_changed_entry_with_shared_budget(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    entries = []
    for name in ("a", "b"):
        write_json(root / f"{name}/handoff.json", source_handoff(
            source={"path": f"{name}.xlsx", "sha256": name}, run_id=name, artifacts=[],
        ))
        entries.append({"relative_path": f"{name}.xlsx", "handoff_path": f"{name}/handoff.json"})
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": entries})
    assert build_index(batch, root, out)["attempt"] == 1
    write_json(root / "a/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[], metrics={"changed": True},
    ))
    original_build, original_resolve = workflow._build_index, workflow._resolve_batch_entry

    def interrupted(*args, **kwargs):
        raise SystemExit("interrupted after reservation")

    monkeypatch.setattr(workflow, "_build_index", interrupted)
    with pytest.raises(SystemExit):
        build_index(batch, root, out, resume=True)
    reserved = json.loads((out / "state.json").read_text())
    assert len(reserved["attempts"]) == 2
    assert reserved["attempts"][1]["result"]["diagnostics"][0]["code"] == "attempt_interrupted"
    # The retry is explicit; an ordinary repeated call still stops.
    assert build_index(batch, root, out)["stop_reason"] == "repeated_tool_input"
    monkeypatch.setattr(workflow, "_build_index", original_build)
    calls = []

    def resolve(root, item, ordinal):
        calls.append(item["relative_path"])
        return original_resolve(root, item, ordinal)

    monkeypatch.setattr(workflow, "_resolve_batch_entry", resolve)
    resumed = build_index(batch, root, out, resume=True)
    assert resumed["status"] == "partial"
    assert resumed["attempt"] == 3
    assert resumed["metrics"]["entries_processed"] == 1
    assert resumed["metrics"]["entries_checked"] == 1
    assert resumed["metrics"]["entries_skipped"] == 1
    assert calls == ["a.xlsx"]
    completed = json.loads((out / "state.json").read_text())
    assert len(completed["attempts"]) == 3
    assert completed["attempts"][1] == reserved["attempts"][1]
    write_json(root / "a/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[], metrics={"changed": "again"},
    ))
    assert build_index(batch, root, out, resume=True)["stop_reason"] == "attempt_limit_reached"
    assert calls == ["a.xlsx"]


def test_writer_opaque_handoff_keeps_correction_and_specialist_actions(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    run = root / "runs/run-a"
    write_json(run / "source.json", source_metadata())
    quality = {
        "status": "fail", "ready_for_next_step": False, "metrics": {"opaque_parts": 2},
        "diagnostics": [{"code": "opaque_parts_preserved", "severity": "warning"}],
        "next_tool": {"name": "inventory.extract", "arguments": {"run": str(run)}},
    }
    _write_handoff(run, quality)
    emitted = json.loads((run / "handoff.json").read_text())
    assert [action["name"] for action in emitted["next_actions"]] == ["inventory.extract"]
    build_index(run / "handoff.json", root, out)
    entry = json.loads((out / "index.json").read_text())["entries"][0]
    assert [action["name"] for action in entry["next_actions"]] == ["inventory.extract", "specialist.review"]
    assert entry["next_actions"][0] == emitted["next_actions"][0]
    summary = (out / "INDEX.md").read_text(encoding="utf-8")
    assert "inventory.extract" in summary and "specialist.review" in summary

    # A handoff that already carries the human/specialist action keeps it once.
    emitted["next_actions"] = entry["next_actions"]
    write_json(run / "handoff.json", emitted)
    build_index(run / "handoff.json", root, out, resume=True)
    updated = json.loads((out / "index.json").read_text())["entries"][0]
    assert updated["next_actions"] == entry["next_actions"]


@pytest.mark.parametrize("corrupt_b", [False, True])
def test_resume_recovers_validated_sibling_after_interrupted_output_write(tmp_path: Path, monkeypatch, corrupt_b: bool) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    entries = []
    for name in ("a", "b"):
        write_json(root / f"{name}/handoff.json", source_handoff(
            source={"path": f"{name}.xlsx", "sha256": name}, run_id=name, artifacts=[],
        ))
        entries.append({"relative_path": f"{name}.xlsx", "handoff_path": f"{name}/handoff.json"})
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": entries})
    build_index(batch, root, out)
    write_json(root / "a/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[], metrics={"changed": True},
    ))
    original_validate, original_resolve = workflow._validate_saved_index, workflow._resolve_batch_entry

    def interrupted(*args, **kwargs):
        raise SystemExit("interrupted after index and summary writes")

    monkeypatch.setattr(workflow, "_validate_saved_index", interrupted)
    with pytest.raises(SystemExit):
        build_index(batch, root, out, resume=True)
    reserved = json.loads((out / "state.json").read_text())
    assert len(reserved["attempts"]) == 2
    assert not workflow._outputs_match(out, reserved)
    partial = json.loads((out / "index.json").read_text())
    assert partial["entries"][0]["metrics"] == {"changed": True}
    if corrupt_b:
        partial["entries"][1]["metrics"] = {"corrupt": True}
        workflow._write_index(out, workflow.Step2Index.model_validate(partial))
    calls, validated = [], []

    def resolve(root, item, ordinal):
        calls.append(item["relative_path"])
        return original_resolve(root, item, ordinal)

    def validate(index, root, skipped=None, **kwargs):
        saved = json.loads(index.read_text())
        validated.extend(entry["source_path"] for key, entry in zip(kwargs["entry_keys"], saved["entries"]) if key not in (skipped or set()))
        return original_validate(index, root, skipped, **kwargs)

    monkeypatch.setattr(workflow, "_resolve_batch_entry", resolve)
    monkeypatch.setattr(workflow, "_validate_saved_index", validate)
    resumed = build_index(batch, root, out, resume=True)
    expected = ["a.xlsx", "b.xlsx"] if corrupt_b else ["a.xlsx"]
    assert calls == validated == expected
    assert resumed["metrics"]["entries_skipped"] == (0 if corrupt_b else 1)
    assert resumed["attempt"] == 3
    assert resumed["stop_reason"] == "attempt_limit_reached"
    final = json.loads((out / "state.json").read_text())
    assert len(final["attempts"]) == 3
    assert final["attempts"][1] == reserved["attempts"][1]
    assert workflow._outputs_match(out, final)
    assert json.loads((out / "index.json").read_text())["entries"][1]["metrics"] == {"traceability_ratio": 1.0}


def test_cached_resume_returns_validated_output_evidence_after_failed_input(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[], diagnostics=[{"code": "preserved_note", "severity": "warning"}]))
    original = handoff.read_bytes()
    built = build_index(handoff, root, out)
    write_json(handoff, {"schema_version": "unsupported"})
    failed = build_index(handoff, root, out)
    assert failed["stop_reason"] == "non_retryable_failure"
    before = (out / "state.json").read_bytes()
    handoff.write_bytes(original)
    cached = build_index(handoff, root, out, resume=True)
    assert cached["status"] == "partial"
    assert cached["source"] == built["source"]
    assert cached["artifacts"] == built["artifacts"]
    assert cached["metrics"]["input_count"] == 1
    assert cached["metrics"]["artifact_count"] == built["metrics"]["artifact_count"]
    assert cached["metrics"]["entries_skipped"] == 1
    assert cached["metrics"]["entries_processed"] == cached["metrics"]["entries_checked"] == 0
    assert cached["diagnostics"] == built["diagnostics"]
    assert not any(d["code"] == "handoff_invalid" for d in cached["diagnostics"])
    assert cached["attempt"] == 2
    assert cached["stop_reason"] is None
    assert (out / "state.json").read_bytes() == before


@pytest.mark.parametrize("shared_source_id", [False, True])
def test_first_build_resume_reuses_entry_completed_before_artifact_interruption(tmp_path: Path, monkeypatch, shared_source_id: bool) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    entries, artifacts = [], {}
    for name in ("a", "b"):
        identity = "a" if shared_source_id else name
        artifact = root / f"{name}/custom.json"
        write_json(artifact, {"source": name})
        artifacts[name] = artifact
        write_json(root / f"{name}/handoff.json", source_handoff(
            source={"path": f"{identity}.xlsx", "sha256": identity}, run_id=identity,
            artifact_paths_relative_to=name, artifacts=[{
                "name": "custom.json", "path": "custom.json", "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
            }],
        ))
        entries.append({"relative_path": f"{identity}.xlsx", "handoff_path": f"{name}/handoff.json"})
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": entries})
    original_read, original_resolve = workflow._read_json, workflow._resolve_batch_entry
    reads = []

    def interrupted(path):
        if path in artifacts.values():
            reads.append(path)
        if path == artifacts["b"]:
            raise SystemExit("interrupted during B artifact validation")
        return original_read(path)

    monkeypatch.setattr(workflow, "_read_json", interrupted)
    with pytest.raises(SystemExit):
        build_index(batch, root, out)
    assert reads == [artifacts["a"], artifacts["b"]]
    reserved = json.loads((out / "state.json").read_text())
    assert len(reserved["attempts"]) == 1
    assert reserved["attempts"][0]["result"]["diagnostics"][0]["code"] == "attempt_interrupted"
    assert len(reserved["entries"]) == 1
    completed_a = next(iter(reserved["entries"].values()))
    assert completed_a["validation_status"] == "pass"
    partial = json.loads((out / "index.json").read_text())
    assert (partial["entries"][0]["source_id"] == partial["entries"][1]["source_id"]) is shared_source_id
    assert completed_a["source_id"] == partial["entries"][0]["source_id"]
    assert completed_a["output_sha256"] == workflow._hash_value(partial["entries"][0])
    calls, reads = [], []

    def read(path):
        if path in artifacts.values():
            reads.append(path)
        return original_read(path)

    def resolve(root, item, ordinal):
        calls.append(item["handoff_path"])
        return original_resolve(root, item, ordinal)

    monkeypatch.setattr(workflow, "_read_json", read)
    monkeypatch.setattr(workflow, "_resolve_batch_entry", resolve)
    resumed = build_index(batch, root, out, resume=True)
    assert calls == ["b/handoff.json"]
    assert reads == [artifacts["b"]]
    assert resumed["metrics"]["entries_skipped"] == 1
    assert resumed["metrics"]["entries_processed"] == resumed["metrics"]["entries_checked"] == 1
    assert resumed["metrics"]["artifacts_checked"] == 1
    assert resumed["attempt"] == 2
    final = json.loads((out / "state.json").read_text())
    assert len(final["entries"]) == len(final["attempts"]) == 2
    assert final["attempts"][0] == reserved["attempts"][0]
    assert workflow._outputs_match(out, final)
    assert build_index(batch, root, out)["stop_reason"] == "repeated_tool_input"


def test_shared_bounds_stop_before_dispatch_and_leave_step1_history_alone(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    history = root / "runs/source-a/attempt_history.json"
    write_json(handoff, source_handoff(artifacts=[]))
    write_json(history, {"attempt_limit": 3, "attempts": [{"tool": "inventory.extract"}]})
    before = history.read_bytes()
    build_index(handoff, root, out)
    repeated = execute_tool("step2.index.build", step1_root=root, handoff_path=handoff, out_dir=out)
    assert repeated["stop_reason"] == "repeated_tool_input"
    assert repeated["artifacts"]
    assert len(json.loads((out / "state.json").read_text())["attempts"]) == 1
    unchanged = validate_saved_index(out / "index.json", root)
    assert unchanged["metrics"]["entries_checked"] == 0
    assert unchanged["attempt"] == 1
    execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=handoff, out_dir=out)
    # A changed input with useful new evidence uses the same cumulative budget.
    write_json(handoff, source_handoff(artifacts=[], metrics={"changed": True}))
    third = build_index(handoff, root, out)
    assert third["attempt"] == 3
    monkeypatch.setattr(workflow, "_build_index", lambda *a, **k: pytest.fail("dispatch after limit"))
    stopped = execute_tool("step2.index.build", step1_root=root, handoff_path=handoff, out_dir=out)
    assert stopped["stop_reason"] == "attempt_limit_reached"
    assert stopped["diagnostics"] is not None
    cached = build_index(handoff, root, out, resume=True)
    assert cached["status"] == "partial"
    assert cached["metrics"]["entries_checked"] == 0
    assert cached["attempt"] == 3
    assert history.read_bytes() == before
    assert all(a["tool"].startswith("step2.") for a in json.loads((out / "state.json").read_text())["attempts"])


def test_no_progress_and_non_retryable_stops_keep_failure_evidence(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff())  # Missing inventory is a Step 1 repair recommendation.
    build_index(handoff, root, out)
    repeated_resume = build_index(handoff, root, out, resume=True)
    assert repeated_resume["stop_reason"] == "repeated_tool_input"
    assert repeated_resume["attempt"] == 1
    checked = validate_saved_index(out / "index.json", root)
    assert checked["stop_reason"] == "no_progress"
    stopped = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=handoff, out_dir=out)
    assert stopped["stop_reason"] == "no_progress"
    assert any(d["code"] == "artifact_missing" for d in stopped["diagnostics"])
    assert (out / "index.json").exists()

    bad_out = tmp_path / "bad-index"
    write_json(handoff, {"schema_version": "unsupported"})
    failed = build_index(handoff, root, bad_out)
    assert failed["stop_reason"] == "non_retryable_failure"
    stopped = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=handoff, out_dir=bad_out)
    assert stopped["stop_reason"] == "non_retryable_failure"
    assert stopped["diagnostics"][0]["code"] == "handoff_invalid"


def test_resolver_failure_wording_does_not_reset_no_progress(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    source, batch = root / "a/handoff.json", root / "batch.json"
    write_json(source, source_handoff(source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[]))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
    ]})
    resolved = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    assert resolved["diagnostics"] == []
    write_json(source, {"schema_version": "wrong"})
    first_failure = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    history = json.loads((out / "state.json").read_text())["attempts"]
    assert history[1]["improved"] is False
    assert history[1]["no_progress_count"] == 1
    assert first_failure["stop_reason"] is None
    write_json(source, {"schema_version": "step1.v1"})
    stopped = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    history = json.loads((out / "state.json").read_text())["attempts"]
    assert history[2]["improved"] is False
    assert history[2]["no_progress_count"] == 2
    assert stopped["attempt"] == 3
    assert stopped["stop_reason"] == "no_progress"
    assert stopped["diagnostics"][0]["code"] == "source_handoff_invalid"
    assert "source object" in stopped["diagnostics"][0]["message"]
    assert stopped["diagnostics"] != first_failure["diagnostics"]
    assert stopped["entries"][0]["diagnostics"] == stopped["diagnostics"]
    assert history[1]["result"] == first_failure
    assert history[2]["result"] == stopped


def test_resolver_progress_keeps_duplicate_source_occurrences(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    for name in ("first", "second"):
        write_json(root / f"{name}/handoff.json", source_handoff(artifacts=[]))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"handoff_path": "first/handoff.json"}, {"handoff_path": "second/handoff.json"},
    ]})
    execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    write_json(root / "first/handoff.json", source_handoff(artifacts=[], metrics={"changed": True}))
    execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    history = json.loads((out / "state.json").read_text())["attempts"]
    assert len(history[0]["resolved_entries"]) == len(history[1]["resolved_entries"]) == 2
    assert history[1]["improved"] is True
    assert history[1]["no_progress_count"] == 0


def test_resolver_useful_entry_changes_and_fewer_errors_remain_progress(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    write_json(root / "a/handoff.json", source_handoff(source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[]))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
        {"relative_path": "b.xlsx", "handoff_path": "b/handoff.json"},
    ]})
    execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    write_json(root / "a/handoff.json", source_handoff(source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[], metrics={"updated": True}))
    execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    changed = json.loads((out / "state.json").read_text())["attempts"][-1]
    assert changed["improved"] is True
    assert changed["no_progress_count"] == 0
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
    ]})
    execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=batch, out_dir=out)
    fewer_errors = json.loads((out / "state.json").read_text())["attempts"][-1]
    assert fewer_errors["result"]["diagnostics"] == []
    assert fewer_errors["improved"] is True
    assert fewer_errors["no_progress_count"] == 0


@pytest.mark.parametrize("initial_valid", [False, True])
def test_build_blocked_entry_changes_do_not_reset_no_progress(tmp_path: Path, monkeypatch, initial_valid: bool) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    source, batch = root / "a/handoff.json", root / "batch.json"
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
    ]})
    if initial_valid:
        write_json(source, source_handoff(source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[]))
        build_index(batch, root, out)
    write_json(source, {"schema_version": "wrong"})
    first = build_index(batch, root, out, resume=True)
    first_state = json.loads((out / "state.json").read_text())
    assert first_state["attempts"][-1]["improved"] is False
    assert first_state["attempts"][-1]["no_progress_count"] == 1
    assert all(entry["validation_status"] == "blocked" for entry in first_state["entries"].values())
    assert first["stop_reason"] is None
    write_json(source, {"schema_version": "step1.v1"})
    second = build_index(batch, root, out, resume=True)
    final = json.loads((out / "state.json").read_text())
    assert final["attempts"][-1]["improved"] is False
    assert final["attempts"][-1]["no_progress_count"] == 2
    assert all(entry["validation_status"] == "blocked" for entry in final["entries"].values())
    assert second["stop_reason"] == "no_progress"
    assert second["attempt"] == (3 if initial_valid else 2)
    assert "source object" in second["diagnostics"][0]["message"]
    assert second["diagnostics"] != first["diagnostics"]
    assert final["attempts"][-1]["result"] == second
    before = (out / "state.json").read_bytes()
    write_json(source, {"schema_version": "step1.v1", "source": {}})
    monkeypatch.setattr(workflow, "_build_index", lambda *a, **k: pytest.fail("dispatch after no-progress stop"))
    stopped = build_index(batch, root, out, resume=True)
    assert stopped["recovery_stopped"] is True
    assert stopped["diagnostics"] == second["diagnostics"]
    assert stopped["artifacts"] == second["artifacts"]
    assert stopped["attempt"] == second["attempt"]
    assert (out / "state.json").read_bytes() == before


def test_build_validated_entry_changes_and_fewer_errors_remain_progress(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    write_json(root / "a/handoff.json", source_handoff(source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[]))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
        {"relative_path": "b.xlsx", "handoff_path": "b/handoff.json"},
    ]})
    build_index(batch, root, out)
    write_json(root / "a/handoff.json", source_handoff(source={"path": "a.xlsx", "sha256": "a"}, run_id="a", artifacts=[], metrics={"updated": True}))
    build_index(batch, root, out, resume=True)
    changed = json.loads((out / "state.json").read_text())["attempts"][-1]
    assert changed["improved"] is True
    assert changed["no_progress_count"] == 0
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
    ]})
    build_index(batch, root, out, resume=True)
    fewer_errors = json.loads((out / "state.json").read_text())["attempts"][-1]
    assert fewer_errors["result"]["diagnostics"] == []
    assert fewer_errors["improved"] is True
    assert fewer_errors["no_progress_count"] == 0


def test_resolver_dispatches_once_and_opaque_action_is_visible(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[], metrics={"opaque_parts": 2}, diagnostics=[{"code": "opaque_parts_preserved", "message": "Preserved"}]))
    monkeypatch.setattr(workflow, "_build_index", lambda *a, **k: pytest.fail("resolver built an index"))
    resolved = execute_tool("step2.handoff.resolve", step1_root=root, handoff_path=handoff, out_dir=out)
    assert resolved["tool"] == "step2.handoff.resolve"
    assert not (out / "index.json").exists()
    assert resolved["entries"][0]["next_actions"][0]["name"] == "specialist.review"
    monkeypatch.undo()
    build_index(handoff, root, out)
    assert "specialist.review" in (out / "INDEX.md").read_text(encoding="utf-8")
    repeated = build_index(handoff, root, out)
    assert repeated["stop_reason"] == "repeated_tool_input"


def test_resume_preserves_partial_index_and_retries_corrected_entry(tmp_path: Path) -> None:
    root, out = tmp_path / "step1", tmp_path / "index"
    batch = root / "batch.json"
    write_json(root / "a/handoff.json", source_handoff(source={"path": "a.xlsx", "sha256": "a"}, artifacts=[]))
    write_json(batch, {"schema_version": "step1.batch.v1", "entries": [
        {"relative_path": "a.xlsx", "handoff_path": "a/handoff.json"},
        {"relative_path": "b.xlsx", "handoff_path": "b/handoff.json", "status": "fail", "metrics": {"opaque_parts": 1}},
    ]})
    first = build_index(batch, root, out)
    assert first["metrics"]["input_count"] == 2
    saved = json.loads((out / "index.json").read_text())["entries"]
    assert saved[1]["diagnostics"][-1]["code"] == "source_handoff_unavailable"
    assert saved[1]["next_actions"][0]["name"] == "specialist.review"
    write_json(root / "b/handoff.json", source_handoff(source={"path": "b.xlsx", "sha256": "b"}, artifacts=[]))
    resumed = build_index(batch, root, out, resume=True)
    assert resumed["metrics"]["entries_skipped"] == 1
    assert resumed["metrics"]["entries_checked"] == 1
    assert resumed["metrics"]["input_count"] == 2


def test_invalid_state_is_preserved_and_interrupted_dispatch_is_counted(tmp_path: Path, monkeypatch) -> None:
    from excel_to_act.steps.step2 import workflow

    root, out = tmp_path / "step1", tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    out.mkdir()
    state_path = out / "state.json"
    state_path.write_text("{", encoding="utf-8")
    before = state_path.read_bytes()
    failed = build_index(handoff, root, out)
    assert failed["stop_reason"] == "state_invalid"
    assert state_path.read_bytes() == before
    fresh_out = tmp_path / "fresh"

    def interrupted(*args, **kwargs):
        raise SystemExit("interrupted")

    monkeypatch.setattr(workflow, "_build_index", interrupted)
    with pytest.raises(SystemExit):
        build_index(handoff, root, fresh_out)
    saved = json.loads((fresh_out / "state.json").read_text())
    assert len(saved["attempts"]) == 1
    assert saved["attempts"][0]["result"]["diagnostics"][0]["code"] == "attempt_interrupted"
    stopped = build_index(handoff, root, fresh_out)
    assert stopped["stop_reason"] == "repeated_tool_input"


def test_builds_single_current_handoff_with_relative_artifact_refs_and_quality(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    input_path = root / "batches/batch-a/sources/source-a/handoff.json"
    inventory_path = root / "runs/source-a/artifacts/inventory.json"
    inventory = {
        "schema_version": "phase1.v1", "artifact_type": "workbook_inventory",
        "workbook_sha256": "source-hash", "sheets": [],
        "coverage": {"recognized_inventory_objects": 0, "unsupported_or_opaque_objects": 0, "discovered_workbook_objects": 0},
    }
    write_json(inventory_path, inventory)
    write_json(input_path, source_handoff(artifacts=[{
        "name": "inventory", "path": "inventory.json",
        "sha256": hashlib.sha256(inventory_path.read_bytes()).hexdigest(),
    }]))

    result = build_index(input_path, root, tmp_path / "index")
    payload = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))

    assert result["status"] == "partial"
    assert len(payload["entries"]) == 1
    entry = payload["entries"][0]
    assert entry["source_path"] == "models/a.xlsx"
    assert entry["source_sha256"] == "source-hash"
    assert entry["status"] == "pass"
    assert entry["handoff_path"] == "batches/batch-a/sources/source-a/handoff.json"
    assert entry["artifacts"][0]["path"] == "runs/source-a/artifacts/inventory.json"
    assert entry["metrics"] == {"traceability_ratio": 1.0}
    assert entry["thresholds"] == {"traceability": 1.0}
    assert entry["final_output"] == "final/batches/batch-a/source-a"
    assert "large content" not in (tmp_path / "index/index.json").read_text(encoding="utf-8")


def test_accepts_legitimate_failed_handoff_with_missing_hash_and_no_artifacts(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = root / "failed/handoff.json"
    write_json(handoff, {
        "schema_version": "step1.v1",
        "source": {"path": None, "sha256": None},
        "run_id": "failed-run",
        "status": "fail",
        "ready_for_next_step": False,
        "metrics_state": "unavailable",
        "metrics": {},
        "artifacts": [],
    })

    result = build_index(handoff, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["status"] == "fail"
    assert entry["source_sha256"] is None
    assert entry["artifacts"] == []
    assert entry["diagnostics"] == []


def test_rejects_source_metadata_shape_as_root_handoff(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = root / "run/source.json"
    write_json(handoff, source_metadata())

    result = build_index(handoff, root, tmp_path / "index")

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["code"] == "handoff_invalid"
    assert "source" in result["diagnostics"][0]["message"]


def test_rejects_source_metadata_shape_in_batch_and_preserves_batch_evidence(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    metadata_path = root / "run/source.json"
    batch_path = root / "batch_handoff.json"
    write_json(metadata_path, source_metadata())
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{
            "relative_path": "models/a.xlsx",
            "sha256": None,
            "run_id": "run-a",
            "status": "fail",
            "metrics_state": "unavailable",
            "metrics": {"preserved": True},
            "diagnostics": [{"code": "source_read_failed"}],
            "handoff_path": "run/source.json",
        }],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["source_path"] == "models/a.xlsx"
    assert entry["run_id"] == "run-a"
    assert entry["status"] == "fail"
    assert entry["metrics"] == {"preserved": True}
    assert entry["diagnostics"][0]["code"] == "source_read_failed"
    assert entry["diagnostics"][1]["code"] == "source_handoff_invalid"


def test_batch_keeps_failed_sources_and_missing_handoffs_with_evidence(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batches/batch-b/batch_handoff.json"
    write_json(root / "sources/pass/handoff.json", source_handoff(
        source={"path": "models/a.xlsx", "sha256": "sha-a"}, status="partial", run_id="run-pass", artifacts=[]
    ))
    write_json(root / "sources/fail/handoff.json", source_handoff(
        source={"path": "models/b.xlsx", "sha256": "sha-b"},
        status="partial",
        ready_for_next_step=False,
        run_id="run-fail",
        metrics={"opaque_rate": 0.2},
        artifacts=[],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "batch_id": "batch-b",
        "entries": [
            {"relative_path": "models/a.xlsx", "sha256": "sha-a", "run_id": "run-pass", "status": "partial", "handoff_path": "sources/pass/handoff.json", "metrics": {"fallback": 1}},
            {"relative_path": "models/b.xlsx", "sha256": "sha-b", "run_id": "run-fail", "status": "fail", "handoff_path": "sources/fail/handoff.json", "diagnostics": [{"code": "failed"}]},
            {"relative_path": "models/c.xlsx", "sha256": "sha-c", "run_id": "run-missing", "status": "fail", "metrics": {"retained": 3}, "diagnostics": [{"code": "source_error"}]},
        ],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entries = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"]

    assert result["status"] == "partial"
    assert [entry["status"] for entry in entries] == ["partial", "fail", "fail"]
    assert [entry["source_path"] for entry in entries] == ["models/a.xlsx", "models/b.xlsx", "models/c.xlsx"]
    assert [entry["source_sha256"] for entry in entries] == ["sha-a", "sha-b", "sha-c"]
    assert [entry["run_id"] for entry in entries] == ["run-pass", "run-fail", "run-missing"]
    assert all(entry["source_id"] for entry in entries)
    assert entries[1]["diagnostics"][0]["code"] == "failed"
    assert entries[2]["metrics"] == {"retained": 3}
    assert entries[2]["diagnostics"][0]["code"] == "source_error"
    assert any(item["code"] == "source_handoff_unavailable" for item in entries[2]["diagnostics"])
    assert entries[2]["handoff_path"] is None


@pytest.mark.parametrize(
    ("handoff_update", "mismatch"),
    [
        ({"source": {"path": "models/other.xlsx", "sha256": "sha-b"}}, "source path"),
        ({"source": {"path": "models/b.xlsx", "sha256": "other-sha"}}, "source SHA-256"),
        ({"run_id": "other-run"}, "run ID"),
    ],
)
def test_batch_rejects_handoff_with_mismatched_identity_and_keeps_batch_evidence(
    tmp_path: Path,
    handoff_update: dict,
    mismatch: str,
) -> None:
    root = tmp_path / "step1"
    updates = {"source": {"path": "models/b.xlsx", "sha256": "sha-b"}, "run_id": "run-b"}
    updates.update(handoff_update)
    handoff = source_handoff(**updates, metrics={"wrong": True})
    handoff_path = root / "sources/other/handoff.json"
    batch_path = root / "batch.json"
    write_json(handoff_path, handoff)
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "batch_id": "batch-a",
        "entries": [{
            "relative_path": "models/b.xlsx",
            "sha256": "sha-b",
            "run_id": "run-b",
            "status": "fail",
            "metrics": {"batch_evidence": True},
            "diagnostics": [{"code": "batch_diagnostic"}],
            "handoff_path": "sources/other/handoff.json",
        }],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["source_path"] == "models/b.xlsx"
    assert entry["source_sha256"] == "sha-b"
    assert entry["run_id"] == "run-b"
    assert entry["status"] == "fail"
    assert entry["metrics"] == {"batch_evidence": True}
    assert entry["artifacts"] == []
    assert entry["diagnostics"][0]["code"] == "batch_diagnostic"
    assert entry["diagnostics"][1]["code"] == "source_handoff_mismatch"
    assert mismatch in entry["diagnostics"][1]["message"]


def test_batch_identity_check_allows_missing_source_hash_and_run_id(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batch.json"
    write_json(root / "sources/a/handoff.json", source_handoff(
        source={"path": "models/a.xlsx", "sha256": None},
        run_id=None,
        status="fail",
        ready_for_next_step=False,
        artifacts=[],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{
            "relative_path": "models/a.xlsx",
            "sha256": "batch-sha",
            "run_id": "batch-run",
            "status": "fail",
            "handoff_path": "sources/a/handoff.json",
        }],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["source_sha256"] == "batch-sha"
    assert entry["run_id"] == "batch-run"
    assert entry["diagnostics"] == []


def test_batch_handoff_and_artifact_paths_use_layered_bases_not_final_output(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batch_handoff.json"
    write_json(root / "per-source/handoff.json", source_handoff(
        source={"path": "a.xlsx", "sha256": "a"},
        run_id="1",
        run_path="runs/alternate",
        artifact_paths_relative_to="runs/chosen",
        final_output="final/elsewhere",
        artifacts=[{"name": "facts", "path": "facts.json", "sha256": "facts-hash"}],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{"relative_path": "a.xlsx", "sha256": "a", "run_id": "1", "status": "pass", "handoff_path": "per-source/handoff.json"}],
    })

    build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert entry["handoff_path"] == "per-source/handoff.json"
    assert entry["artifacts"][0]["path"] == "runs/chosen/facts.json"
    assert entry["run_path"] == "runs/alternate"
    assert entry["final_output"] == "final/elsewhere"


def test_artifacts_fall_back_to_run_path_when_relative_base_is_absent(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = source_handoff(artifact_paths_relative_to=None, run_path="runs/fallback")
    input_path = root / "handoff.json"
    write_json(input_path, handoff)

    build_index(input_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert entry["artifacts"][0]["path"] == "runs/fallback/inventory.json"


def test_batch_path_outside_root_is_not_retained_as_a_handoff_reference(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    batch_path = root / "batch.json"
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{"relative_path": "a.xlsx", "sha256": "a", "status": "fail", "handoff_path": "../../outside.json"}],
    })

    result = build_index(batch_path, root, tmp_path / "index")
    entry = json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert result["status"] == "partial"
    assert entry["handoff_path"] is None
    assert any(item["code"] == "source_handoff_unavailable" for item in entry["diagnostics"])


def test_rejects_outside_root_and_malformed_root_handoffs(tmp_path: Path, monkeypatch) -> None:
    root = tmp_path / "step1"
    root.mkdir()
    monkeypatch.chdir(tmp_path)
    outside = tmp_path / "outside.json"
    write_json(outside, source_handoff())
    write_json(root / "outside.json", source_handoff())

    outside_result = build_index(Path("outside.json"), root, tmp_path / "outside-index")
    malformed_path = root / "bad.json"
    malformed_path.write_text("[]", encoding="utf-8")
    malformed_result = build_index(malformed_path, root, tmp_path / "bad-index")

    assert outside_result["status"] == "blocked"
    assert outside_result["diagnostics"][0]["code"] == "handoff_invalid"
    assert malformed_result["status"] == "blocked"
    assert malformed_result["diagnostics"][0]["code"] == "handoff_invalid"


def test_rejects_legacy_schema_and_keeps_escaping_artifact_out_of_refs(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    legacy_path = root / "legacy.json"
    write_json(legacy_path, {"schema_version": "phase1.v1", "artifacts": []})
    result = build_index(legacy_path, root, tmp_path / "legacy-index")

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["code"] == "handoff_invalid"

    handoff_path = root / "handoff.json"
    write_json(handoff_path, source_handoff(artifacts=[{"name": "bad", "path": "../../../../outside.json", "sha256": "x"}]))
    build_index(handoff_path, root, tmp_path / "artifact-index")
    entry = json.loads((tmp_path / "artifact-index/index.json").read_text(encoding="utf-8"))["entries"][0]

    assert entry["artifacts"] == []
    assert any(item["code"] == "artifact_path_invalid" for item in entry["diagnostics"])


def test_unchanged_inputs_produce_identical_index_files(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff())
    first = tmp_path / "first"
    second = tmp_path / "second"

    build_index(handoff, root, first)
    build_index(handoff, root, second)

    assert (first / "index.json").read_bytes() == (second / "index.json").read_bytes()
    assert (first / "INDEX.md").read_bytes() == (second / "INDEX.md").read_bytes()


def test_tools_catalogue_documents_resolver_and_builder_contracts() -> None:
    tools = tool_catalog()["tools"]

    assert [tool["name"] for tool in tools] == ["step2.handoff.resolve", "step2.index.build", "step2.index.validate"]
    for tool in tools:
        assert tool["command"]
        assert tool["inputs"]
        assert tool["outputs"]
        assert "next" in tool


def test_saved_index_validator_checks_paths_hashes_and_known_schemas(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    entry = index["entries"][0]
    graph = root / "runs/source-a/dependency_graph.json"
    write_json(graph, {"artifact_type": "wrong", "source_location": {"workbook_path": "a.xlsx", "object_type": "cell"}})
    outside = tmp_path / "outside.json"
    write_json(outside, {"opaque": True})
    opaque = root / "runs/source-a/unknown.json"
    write_json(opaque, {"opaque": True})
    changed_graph = root / "runs/source-a/changed_graph.json"
    write_json(changed_graph, {"nodes": [{"source_location": {"workbook_path": "a.xlsx", "sheet_name": "Sheet1", "address": "A1", "object_type": "cell"}}]})
    malformed = root / "runs/source-a/malformed.json"
    malformed.parent.mkdir(parents=True, exist_ok=True)
    malformed.write_text("{", encoding="utf-8")
    entry["artifacts"] = [
        {"name": "dependency_graph.json", "path": graph.relative_to(root).as_posix(), "sha256": hashlib.sha256(graph.read_bytes()).hexdigest()},
        {"name": "inventory.json", "path": "runs/source-a/missing.json", "sha256": "missing"},
        {"name": "outside.json", "path": "../outside.json", "sha256": hashlib.sha256(outside.read_bytes()).hexdigest()},
        {"name": "unknown.json", "path": opaque.relative_to(root).as_posix(), "sha256": hashlib.sha256(opaque.read_bytes()).hexdigest()},
        {"name": "inventory.json", "path": malformed.relative_to(root).as_posix(), "sha256": hashlib.sha256(malformed.read_bytes()).hexdigest()},
        {"name": "dependency_graph.json", "path": changed_graph.relative_to(root).as_posix(), "sha256": "old-checksum"},
        {"name": "inventory.json", "path": malformed.relative_to(root).as_posix(), "sha256": "old-checksum"},
    ]
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    assert {item["code"] for item in result["diagnostics"]} == {
        "artifact_path_invalid", "artifact_missing", "artifact_schema_invalid", "artifact_json_invalid", "artifact_checksum_mismatch"
    }
    assert all(item["source_id"] == entry["source_id"] for item in result["diagnostics"])
    assert all(item["artifact"] for item in result["diagnostics"])
    schema_error = next(item for item in result["diagnostics"] if item["code"] == "artifact_schema_invalid")
    assert schema_error["source_location"]["workbook_path"] == "a.xlsx"
    changed_error = next(item for item in result["diagnostics"] if item["code"] == "artifact_checksum_mismatch" and item["artifact"] == "dependency_graph.json")
    assert changed_error["source_location"]["address"] == "A1"
    malformed_error = next(item for item in result["diagnostics"] if item["code"] == "artifact_checksum_mismatch" and item["artifact"] == "inventory.json")
    assert "source_location" not in malformed_error
    assert result["next_tool"] is None


def test_saved_index_validator_checks_current_step1_identity_fields(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    source_record = source_metadata()
    source_record.update({"relative_path": "models/other.xlsx", "run_id": "run-b", "sha256": "another-workbook"})
    artifacts = {
        "workbook_manifest.json": {
            "schema_version": "phase1.v1", "artifact_type": "workbook_manifest",
            "workbook_path": "a.xlsx", "file_name": "a.xlsx", "file_size": 0,
            "sha256": "another-workbook", "sheets": [],
        },
        "source.json": source_record,
        "quality.json": {
            "schema_version": "step1.v1", "source_sha256": "another-workbook", "run_id": "run-a",
        },
        "custom.json": {"sha256": "opaque-hash"},
    }
    refs = []
    for name, payload in artifacts.items():
        path = root / "runs/source-a" / name
        write_json(path, payload)
        refs.append({"name": name, "path": path.relative_to(root).as_posix(), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    index["entries"][0]["artifacts"] = refs
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    mismatches = [item for item in result["diagnostics"] if item["code"] == "artifact_identity_mismatch"]
    assert {(item["artifact"], "source path" in item["message"]) for item in mismatches} == {
        ("workbook_manifest.json", False), ("source.json", True),
        ("source.json", False), ("source.json", False), ("quality.json", False),
    }
    assert not any(item["artifact"] == "custom.json" for item in mismatches)


@pytest.mark.parametrize(
    ("name", "payload"),
    [
        ("workbook_manifest.json", {
            "schema_version": "phase1.v1", "artifact_type": "workbook_manifest",
            "workbook_path": "a.xlsx", "file_name": "a.xlsx", "file_size": 0,
            "sha256": "wrong-source", "sheets": [],
        }),
        ("source.json", {"relative_path": "models/other.xlsx", "sha256": "wrong-source"}),
        ("quality.json", {"source_sha256": "wrong-source", "run_id": "run-a"}),
    ],
)
def test_build_index_blocks_mismatched_current_step1_artifacts(tmp_path: Path, name: str, payload: dict) -> None:
    root = tmp_path / "step1"
    handoff = root / "handoff.json"
    if name == "source.json":
        payload = source_metadata()
        payload.update({"relative_path": "models/other.xlsx", "sha256": "wrong-source"})
    artifact = root / "runs/source-a/artifacts" / name
    write_json(artifact, payload)
    write_json(handoff, source_handoff(artifacts=[{
        "name": name, "path": name,
        "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]))

    result = build_index(handoff, root, tmp_path / "index")

    assert result["status"] == "blocked"
    assert any(item["code"] == "artifact_identity_mismatch" and item["artifact"] == name for item in result["diagnostics"])
    assert json.loads((tmp_path / "index/index.json").read_text(encoding="utf-8"))["validation_status"] == "blocked"


def test_saved_index_validator_accepts_legacy_phase1_handoff(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    artifact = root / "runs/source-a/handoff.json"
    write_json(artifact, {
        "schema_version": "phase1.v1", "artifact_type": "handoff",
        "workbook_sha256": "source-hash", "run_id": "run-a",
    })
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    index["entries"][0]["artifacts"] = [{
        "name": "handoff.json", "path": artifact.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]
    write_json(index_path, index)

    assert validate_saved_index(index_path, root)["status"] == "pass"


def test_saved_index_validator_passes_known_and_opaque_artifacts(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    artifacts = [
        ("dependency_graph.json", {"schema_version": "phase1.v1", "artifact_type": "formula_graph", "nodes": [], "edges": []}),
        ("custom.json", {"custom": "opaque"}),
    ]
    refs = []
    for name, payload in artifacts:
        path = root / "runs/source-a" / name
        write_json(path, payload)
        refs.append({"name": name, "path": path.relative_to(root).as_posix(), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    index["entries"][0]["artifacts"] = refs
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "pass"
    assert result["diagnostics"] == []
    assert result["metrics"]["artifacts_checked"] == 2
    assert result["retryable"] is False


def test_saved_index_validator_accepts_and_repairs_writer_emitted_current_handoff(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    run_dir = root / "batches/batch-a/sources/source-a/runs/run-a"
    source = {
        "schema_version": "step1.v1", "batch_id": "batch-a", "run_id": "run-a",
        "source_path": "C:/models/a.xlsx", "relative_path": "models/a.xlsx",
        "sha256": "source-hash", "suffix": ".xlsx", "thresholds": {},
        "allow_opaque": True, "attempt_limit": 3,
    }
    write_json(run_dir / "source.json", source)
    write_json(run_dir / "quality.json", {
        "source_sha256": "source-hash", "run_id": "run-a", "status": "pass",
        "ready_for_next_step": True, "metrics_state": "current", "metrics": {},
    })
    quality = json.loads((run_dir / "quality.json").read_text(encoding="utf-8"))
    _write_handoff(run_dir, quality)
    handoff_path = run_dir / "handoff.json"
    build_index(handoff_path, root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    artifact = handoff_path
    index["entries"][0]["artifacts"] = [{
        "name": "handoff.json", "path": artifact.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(artifact.read_bytes()).hexdigest(),
    }]
    write_json(index_path, index)

    valid = validate_saved_index(index_path, root)
    assert valid["status"] == "pass"

    emitted = json.loads(artifact.read_text(encoding="utf-8"))
    emitted.pop("source")
    write_json(artifact, emitted)
    index["entries"][0]["artifacts"][0]["sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    index["input_handoff_sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    write_json(index_path, index)
    invalid = validate_saved_index(index_path, root)

    assert invalid["status"] == "blocked"
    assert any(item["code"] == "artifact_schema_invalid" and item["artifact"] == "handoff.json" for item in invalid["diagnostics"])
    artifact_error = next(item for item in invalid["diagnostics"] if item["code"] == "artifact_schema_invalid")
    assert artifact_error["next_tool"] == "report.handoff"
    assert invalid["next_tool"] is None

    _write_handoff(run_dir, quality)
    index["entries"][0]["artifacts"][0]["sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    index["input_handoff_sha256"] = hashlib.sha256(artifact.read_bytes()).hexdigest()
    write_json(index_path, index)
    fourth = validate_saved_index(index_path, root)
    assert fourth["stop_reason"] == "attempt_limit_reached"
    assert fourth["attempt"] == 3
    repaired_index = tmp_path / "repaired-index/index.json"
    write_json(repaired_index, index)
    assert validate_saved_index(repaired_index, root)["status"] == "pass"


def test_saved_index_validator_checks_input_handoff_reference(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, index_dir)
    handoff.write_text("{}", encoding="utf-8")

    result = validate_saved_index(index_dir / "index.json", root)

    assert result["status"] == "blocked"
    assert result["diagnostics"][0]["code"] == "input_handoff_checksum_mismatch"


def test_saved_index_validator_checks_input_handoff_schema_even_with_updated_digest(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    handoff = root / "handoff.json"
    write_json(handoff, source_handoff(artifacts=[]))
    build_index(handoff, root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    write_json(handoff, {})
    index["input_handoff_sha256"] = hashlib.sha256(handoff.read_bytes()).hexdigest()
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    assert any(item["code"] == "input_handoff_schema_invalid" and item["artifact"] == "input_handoff" for item in result["diagnostics"])


@pytest.mark.parametrize("invalid_handoff_path", ["../outside.json", "sources/pass/missing.json"])
def test_saved_index_validator_rechecks_batch_entry_handoff_path(tmp_path: Path, invalid_handoff_path: str) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    batch_path = root / "batch_handoff.json"
    per_source_handoff = root / "sources/pass/handoff.json"
    write_json(per_source_handoff, source_handoff(
        source={"path": "models/a.xlsx", "sha256": "sha-a"},
        run_id="run-a",
        artifacts=[],
    ))
    write_json(batch_path, {
        "schema_version": "step1.batch.v1",
        "entries": [{
            "relative_path": "models/a.xlsx",
            "sha256": "sha-a",
            "run_id": "run-a",
            "status": "pass",
            "handoff_path": "sources/pass/handoff.json",
        }],
    })
    build_index(batch_path, root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    assert validate_saved_index(index_path, root)["status"] == "pass"

    index["entries"][0]["handoff_path"] = invalid_handoff_path
    write_json(index_path, index)
    if invalid_handoff_path == "sources/pass/missing.json":
        per_source_handoff.unlink()

    result = validate_saved_index(index_path, root)

    expected_code = "source_handoff_path_invalid" if invalid_handoff_path.startswith("..") else "source_handoff_missing"
    diagnostic = next(item for item in result["diagnostics"] if item["code"] == expected_code)
    assert result["status"] == "blocked"
    assert diagnostic["artifact"] == "source_handoff"
    assert diagnostic["source_path"] == "models/a.xlsx"


def test_saved_index_validator_rejects_incomplete_source_record_artifact(tmp_path: Path) -> None:
    root = tmp_path / "step1"
    index_dir = tmp_path / "index"
    write_json(root / "handoff.json", source_handoff(artifacts=[]))
    build_index(root / "handoff.json", root, index_dir)
    index_path = index_dir / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8"))
    source_path = root / "runs/source-a/source.json"
    source_record = source_metadata()
    source_record.pop("attempt_limit")
    write_json(source_path, source_record)
    index["entries"][0]["artifacts"] = [{
        "name": "source.json",
        "path": source_path.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(source_path.read_bytes()).hexdigest(),
    }]
    write_json(index_path, index)

    result = validate_saved_index(index_path, root)

    assert result["status"] == "blocked"
    assert any(item["code"] == "artifact_schema_invalid" and item["artifact"] == "source.json" for item in result["diagnostics"])
