"""Small, hash-bound state for the six reviewed conversion checkpoints."""

from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path
import urllib.error
import urllib.request
import uuid
from typing import Any


MANIFEST_NAME = "workflow.json"
STAGES = (1, 2, 3, 4, 5, 6)
REJECTION_ROUTES = {
    1: {1},
    2: {1, 2},
    3: {1, 2, 3},
    4: {1, 2, 3, 4},
    5: {1, 2, 3, 4, 5},
    6: {1, 2, 3, 4, 5, 6},
}


def display_report_value(value: Any, missing: str = "Not recorded") -> str:
    """Render a nullable fact without changing false, zero, empty shapes, or literal text."""
    return missing if value is None else str(value)


def display_report_axes(target: Any, limit: int = 6, missing: str = "Not recorded") -> str:
    """Render recorded axis metadata and a bounded sample of typed keys for a report."""
    if not isinstance(target, dict):
        return missing
    axes = target.get("axes", target.get("axis"))
    if axes is None:
        return missing
    if not isinstance(axes, list):
        return _display_report_axis_value(axes, missing)
    if not axes:
        return "[]"
    summaries = []
    for axis in axes:
        if not isinstance(axis, dict):
            summaries.append(_display_report_axis_value(axis, missing))
            continue
        details = []
        for key in ("name", "axis_id", "role", "units", "source_range"):
            if key in axis:
                details.append(f"{key}: {display_report_value(axis[key], missing)}")
        if "keys" in axis:
            details.append(f"keys: {_display_report_axis_keys(axis['keys'], limit, missing)}")
        summaries.append("; ".join(details) or missing)
    return " / ".join(summaries).replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def _display_report_axis_value(value: Any, missing: str) -> str:
    if value is None:
        return missing
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, str):
        return value.replace("|", "\\|").replace("\r", " ").replace("\n", " ")
    if isinstance(value, (int, float)):
        return repr(value) if isinstance(value, float) else str(value)
    if isinstance(value, list):
        return "[" + ", ".join(_display_report_axis_key(item, missing) for item in value) + "]"
    return "structured value"


def _display_report_axis_key(value: Any, missing: str) -> str:
    if value is None:
        return missing
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, (int, float)):
        return repr(value) if isinstance(value, float) else str(value)
    if isinstance(value, list):
        return "[" + ", ".join(_display_report_axis_key(item, missing) for item in value) + "]"
    return "structured value"


def _display_report_axis_keys(value: Any, limit: int, missing: str) -> str:
    if value is None:
        return missing
    if not isinstance(value, list):
        return _display_report_axis_value(value, missing)
    shown = [_display_report_axis_key(item, missing) for item in value[:limit]]
    if len(value) > limit:
        shown.append(f"… {len(value) - limit} more")
    return "[" + ", ".join(shown) + "]"


def design_question_lines(questions: Any) -> list[str]:
    lines = [
        "",
        "## Step 3 design question history",
        "",
        "These bound Step 3 records retain their recorded answers and dispositions. An execution prerequisite is not itself a business interpretation. Later execution evidence is shown separately and does not fill an unrecorded answer.",
    ]
    if not isinstance(questions, list):
        return [*lines, "", "Design questions: Not recorded."]
    if not questions:
        return [*lines, "", "No Step 3 design questions are recorded."]

    def cell(value: Any) -> str:
        return display_report_value(value).replace("|", "\\|").replace("\r\n", "<br>").replace("\r", "<br>").replace("\n", "<br>")

    for question in questions:
        if isinstance(question, dict):
            disposition = cell(question.get("disposition"))
            punctuation = "" if disposition.endswith(".") else "."
            lines.extend([
                "",
                f"- **{cell(question.get('question_id'))}** — {cell(question.get('question'))}",
                f"  - Recorded answer: {cell(question.get('answer'))}; status: "
                f"{cell(question.get('status'))}; disposition: {disposition}{punctuation}",
            ])
        else:
            lines.extend(["", f"- {cell(question)}"])
    return lines


def review_tool_entries(stage: int) -> list[dict[str, Any]]:
    """Direct CLI entries shared by every numbered checkpoint catalogue."""
    if stage not in STAGES:
        raise ValueError(f"unsupported workflow stage: {stage}")
    reviewers = "agent|human|typesafe" if stage >= 3 else "agent|human"
    evidence_option = " [--evidence REVIEW.json]" if stage >= 3 else ""
    return_stages = "|".join(str(value) for value in range(1, stage + 1))
    entries = [
        {"name": "workflow.status", "command": "workflow status --workflow DIR",
         "inputs": {"workflow": "six-stage conversion workflow directory"},
         "outputs": ["checkpoint freshness", "recorded decisions", "next permitted stage"],
         "next": ["workflow.confirm", "workflow.reject"]},
        {"name": "workflow.confirm",
         "command": f"workflow confirm --workflow DIR --stage {stage} --reviewer {reviewers} --decision approve --message TEXT{evidence_option}",
         "inputs": {"workflow": "conversion workflow", "stage": str(stage),
                    "reviewer": "agent or human" if stage < 3 else "agent, human, or delegated TypeSafe",
                    "decision": "approve only; use workflow.reject to reject", "message": "review evidence or explicit user response"},
         "outputs": ["receipt bound to the current JSON/Markdown revision and upstream inputs"],
         "next": ["workflow.status"]},
        {"name": "workflow.reject",
         "command": f"workflow reject --workflow DIR --stage {stage} --reviewer {reviewers} --return-to {return_stages} --message TEXT{evidence_option}",
         "inputs": {"workflow": "conversion workflow", "stage": str(stage),
                    "return_to": f"current or earlier checkpoint {return_stages}"},
         "outputs": ["rejection receipt and invalidated downstream approvals"],
         "next": ["workflow.status"]},
    ]
    if stage == 3:
        entries.append({"name": "workflow.delegate",
                        "command": "workflow delegate --workflow DIR --authorization AUTH.json",
                        "inputs": {"authorization": "source/workflow-bound explicit delegation for TypeSafe reviews of stages 3..6"},
                        "outputs": ["registered delegation hash; default Agent+human policy otherwise remains unchanged"],
                        "next": ["workflow.status"]})
    if stage >= 3:
        entries.append({"name": "workflow.typesafe",
                        "command": f"workflow typesafe --workflow DIR --stage {stage} --request REQUEST.json",
                        "inputs": {"request": "current-artifact-bound TypeSafe Choice request; requires TYPESAFE_API_KEY"},
                        "outputs": ["live request/response/review files and a workflow invocation record"],
                        "next": ["workflow.confirm", "workflow.reject"]})
    return entries


def json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True, default=str) + "\n").encode("utf-8")


def hash_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _atomic_write(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_bytes(payload)
    os.replace(temporary, path)


def _write_manifest(root: Path, manifest: dict[str, Any]) -> None:
    _atomic_write(root / MANIFEST_NAME, json_bytes(manifest))


def load_workflow(root: Path) -> tuple[Path, dict[str, Any]]:
    root = root.expanduser().resolve()
    manifest = read_json(root / MANIFEST_NAME)
    if not isinstance(manifest, dict) or manifest.get("schema_version") != "conversion.workflow.v1":
        raise ValueError("workflow.json is missing or has an unsupported schema")
    return root, manifest


def create_workflow(root: Path, source: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    root = root.expanduser().resolve()
    manifest_path = root / MANIFEST_NAME
    if manifest_path.exists():
        existing = read_json(manifest_path)
        old_source = existing.get("source", {}) if isinstance(existing, dict) else {}
        identity_keys = ("workbook_sha256", "run_id", "step1_run_path")
        if not isinstance(existing, dict) or any(old_source.get(key) != source.get(key) for key in identity_keys):
            raise ValueError("workflow root already belongs to a different source workbook")
        return root, existing
    manifest = {"schema_version": "conversion.workflow.v1", "created_at": _now(), "source": source,
                "mode": "production", "stages": {str(stage): {"current_revision": 0, "revisions": []}
                                                    for stage in STAGES}}
    _write_manifest(root, manifest)
    return root, manifest


def bind_source_id(root: Path, source_id: str) -> None:
    root, manifest = load_workflow(root)
    existing = manifest["source"].get("source_id")
    if existing not in (None, source_id):
        raise ValueError("workflow is already bound to a different Step 2 source ID")
    manifest["source"]["source_id"] = source_id
    _write_manifest(root, manifest)


def _source_sha256(manifest: dict[str, Any]) -> str | None:
    source = manifest.get("source", {})
    value = source.get("workbook_sha256") or source.get("source_sha256")
    return value if isinstance(value, str) and value else None


def _validate_delegation(root: Path, manifest: dict[str, Any], authorization: Any) -> dict[str, Any]:
    if not isinstance(authorization, dict) or authorization.get("schema_version") != "workflow.review_delegation.authorization.v1":
        raise ValueError("authorization file has an unsupported schema")
    source = manifest.get("source", {})
    try:
        authorized_workflow = Path(authorization.get("workflow", "")).expanduser().resolve()
    except (OSError, TypeError, ValueError) as exc:
        raise ValueError("authorization has an invalid workflow path") from exc
    if os.path.normcase(str(authorized_workflow)) != os.path.normcase(str(root)):
        raise ValueError("authorization is for a different workflow directory")
    expected = {
        "authorized_by": "human",
        "source_sha256": _source_sha256(manifest),
        "source_id": source.get("source_id"),
        "run_id": source.get("run_id"),
        "delegate": "typesafe",
        "requested_model": "jev-latest",
        "stages": [3, 4, 5, 6],
        "default_policy_restored_for_other_workflows": True,
    }
    for key, value in expected.items():
        if authorization.get(key) != value or (key == "source_sha256" and value is None):
            raise ValueError(f"authorization does not match this workflow: {key}")
    if not isinstance(authorization.get("authorized_at"), str) or not authorization["authorized_at"].strip():
        raise ValueError("authorization is missing its authorization time")
    if not isinstance(authorization.get("user_instruction_translation"), str) or not authorization["user_instruction_translation"].strip():
        raise ValueError("authorization is missing the human instruction")
    if not isinstance(authorization.get("scope"), str) or not authorization["scope"].strip():
        raise ValueError("authorization is missing its scope")
    threshold = authorization.get("minimum_approval_probability")
    if (isinstance(threshold, bool) or not isinstance(threshold, (int, float))
            or not math.isfinite(threshold) or not 0.85 <= threshold <= 1):
        raise ValueError("authorization approval probability threshold must be between 0.85 and 1")
    return authorization


def register_review_delegation(root: Path, authorization_path: Path) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if manifest.get("mode") == "synthetic_fixture":
        raise ValueError("TypeSafe delegation cannot be registered for a synthetic workflow")
    path = authorization_path.expanduser().resolve()
    if not path.is_file():
        raise ValueError("authorization file does not exist")
    authorization = _validate_delegation(root, manifest, read_json(path))
    digest = hash_file(path)
    manifest["review_delegation"] = {
        "schema_version": "workflow.review_delegation.record.v1",
        "authorization_path": str(path),
        "authorization_sha256": digest,
        "source_sha256": authorization["source_sha256"],
        "source_id": authorization["source_id"],
        "run_id": authorization["run_id"],
        "stages": authorization["stages"],
        "requested_model": authorization["requested_model"],
        "minimum_approval_probability": authorization["minimum_approval_probability"],
        "registered_at": _now(),
    }
    _write_manifest(root, manifest)
    return {"status": "pass", "workflow": str(root), "authorization_sha256": digest,
            "stages": authorization["stages"], "requested_model": authorization["requested_model"],
            "minimum_approval_probability": authorization["minimum_approval_probability"]}


def _delegation_state(root: Path, manifest: dict[str, Any]) -> tuple[dict[str, Any] | None, str | None]:
    record = manifest.get("review_delegation")
    if record is None:
        return None, None
    if not isinstance(record, dict) or record.get("schema_version") != "workflow.review_delegation.record.v1":
        return None, "TypeSafe delegation record is invalid"
    path = Path(str(record.get("authorization_path", "")))
    if not path.is_file() or hash_file(path) != record.get("authorization_sha256"):
        return None, "TypeSafe authorization file changed or is missing"
    try:
        authorization = _validate_delegation(root, manifest, read_json(path))
    except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
        return None, f"TypeSafe authorization is no longer valid: {exc}"
    if any(record.get(key) != authorization.get(auth_key) for key, auth_key in (
        ("source_sha256", "source_sha256"), ("source_id", "source_id"), ("run_id", "run_id"),
        ("stages", "stages"), ("requested_model", "requested_model"),
        ("minimum_approval_probability", "minimum_approval_probability"),
    )):
        return None, "TypeSafe delegation record does not match its authorization"
    return record, None


def current_revision(manifest: dict[str, Any], stage: int) -> dict[str, Any] | None:
    stage_value = manifest.get("stages", {}).get(str(stage), {})
    revision = stage_value.get("current_revision", 0)
    revisions = stage_value.get("revisions", [])
    return next((item for item in revisions if item.get("revision") == revision), None)


def append_stage_artifact(
    root: Path,
    stage: int,
    json_name: str,
    payload: dict[str, Any],
    markdown: str,
    *,
    input_files: dict[str, Path] | None = None,
    input_stages: dict[int, str | dict[str, Any]] | None = None,
) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if stage not in STAGES:
        raise ValueError(f"unsupported workflow stage: {stage}")
    stage_entry = manifest["stages"][str(stage)]
    revision_number = len(stage_entry["revisions"]) + 1
    revision_dir = root / f"stage{stage}" / f"revision-{revision_number:04d}"
    json_path = revision_dir / json_name
    md_path = revision_dir / json_name.replace(".json", ".md")
    if json_path.exists() or md_path.exists():
        raise ValueError("stage revision output already exists; use a new workflow root")
    json_content = json_bytes(payload)
    md_content = markdown.encode("utf-8")
    _atomic_write(json_path, json_content)
    _atomic_write(md_path, md_content)
    files = {key: {"path": str(path.expanduser().resolve()), "sha256": hash_file(path)}
             for key, path in (input_files or {}).items()}
    stage_hashes: dict[str, dict[str, Any]] = {}
    for key, supplied in (input_stages or {}).items():
        prior = current_revision(manifest, int(key))
        if prior is None:
            raise ValueError(f"stage {key} has no artifact to bind")
        artifact = prior["artifact"]
        if isinstance(supplied, str) and supplied != artifact["json_sha256"]:
            raise ValueError(f"stage {key} input is not its current JSON artifact")
        if isinstance(supplied, dict) and any(supplied.get(field) != artifact.get(field)
                                              for field in ("revision", "json_sha256", "md_sha256")):
            raise ValueError(f"stage {key} input is not its current artifact pair")
        stage_hashes[str(key)] = {"revision": prior["revision"],
                                  "json_sha256": artifact["json_sha256"],
                                  "md_sha256": artifact["md_sha256"]}
    resolves_rejections = _pending_rejection_tokens(manifest, stage)
    revision = {"revision": revision_number,
                "artifact": {"json": str(json_path.relative_to(root)), "md": str(md_path.relative_to(root)),
                             "json_sha256": hash_bytes(json_content), "md_sha256": hash_bytes(md_content)},
                "input_files": files, "input_stages": stage_hashes, "decisions": [],
                "resolves_rejections": resolves_rejections}
    stage_entry["revisions"].append(revision)
    stage_entry["current_revision"] = revision_number
    if stage == 3 and payload.get("schema_version") == "step3.input_boundary.v1" and payload.get("phase") == "input_boundary":
        manifest["input_boundary"] = {
            "schema_version": "step3.input_boundary.pointer.v1",
            "revision": revision_number,
            "artifact": revision["artifact"],
            "boundary_sha256": payload.get("boundary_sha256"),
            "source_sha256": payload.get("source", {}).get("workbook_sha256"),
            "target_selection_sha256": hash_bytes(json_bytes(payload.get("target_selection"))),
            "submitted_at": _now(),
        }
    _write_manifest(root, manifest)
    return revision


def _revision_staleness(root: Path, manifest: dict[str, Any], stage: int, revision: dict[str, Any]) -> list[str]:
    reasons: list[str] = []
    for key, item in revision.get("input_files", {}).items():
        path = Path(item["path"])
        if not path.is_file() or hash_file(path) != item.get("sha256"):
            reasons.append(f"input file changed or missing: {key}")
    for key, expected in revision.get("input_stages", {}).items():
        prior = current_revision(manifest, int(key))
        actual = prior.get("artifact", {}) if prior else {}
        if isinstance(expected, str):  # Legacy manifests created before paired-artifact binding.
            stage_revisions = manifest.get("stages", {}).get(str(key), {}).get("revisions", [])
            matching = [item for item in stage_revisions
                        if item.get("artifact", {}).get("json_sha256") == expected]
            matches = (prior is not None and len(matching) == 1
                       and matching[0].get("revision") == prior.get("revision")
                       and actual.get("json_sha256") == expected)
        else:
            matches = (prior is not None and prior.get("revision") == expected.get("revision")
                       and actual.get("json_sha256") == expected.get("json_sha256")
                       and actual.get("md_sha256") == expected.get("md_sha256"))
        if not matches:
            reasons.append(f"stage {key} artifact changed")
    artifact = revision.get("artifact", {})
    for kind in ("json", "md"):
        path = root / artifact.get(kind, "")
        expected = artifact.get(f"{kind}_sha256")
        if not path.is_file() or hash_file(path) != expected:
            reasons.append(f"stage {stage} {kind} artifact changed")
    return reasons


def _validate_typesafe_evidence(
    root: Path,
    manifest: dict[str, Any],
    stage: int,
    revision: dict[str, Any],
    evidence_path: Path,
    delegation: dict[str, Any],
) -> dict[str, Any]:
    evidence_path = evidence_path.expanduser().resolve()
    if not evidence_path.is_file():
        raise ValueError("TypeSafe review evidence file does not exist")
    evidence = read_json(evidence_path)
    artifact = revision["artifact"]
    if not isinstance(evidence, dict) or evidence.get("schema_version") != "conversion.typesafe_review.v1":
        raise ValueError("TypeSafe review evidence has an unsupported schema")
    expected = {
        "stage": stage,
        "revision": revision["revision"],
        "source_sha256": _source_sha256(manifest),
        "artifact": {"json_sha256": artifact["json_sha256"], "md_sha256": artifact["md_sha256"]},
        "service": "TypeSafe",
        "requested_model": delegation["requested_model"],
        "question_id": "stage_decision",
    }
    for key, value in expected.items():
        if evidence.get(key) != value:
            raise ValueError(f"TypeSafe review evidence does not match the current stage artifact: {key}")
    invocation_id = evidence.get("invocation_id")
    if not isinstance(invocation_id, str) or not invocation_id:
        raise ValueError("TypeSafe evidence has no matching live CLI invocation")
    invocations = manifest.get("typesafe_invocations", [])
    invocation = next((item for item in invocations
                       if isinstance(item, dict) and item.get("invocation_id") == invocation_id), None)
    if invocation is None:
        raise ValueError("TypeSafe evidence has no matching live CLI invocation")
    resolved_model = evidence.get("resolved_model")
    if not isinstance(resolved_model, str) or not resolved_model.startswith("jev-"):
        raise ValueError("TypeSafe review evidence is missing its resolved Jev model")

    request_record = evidence.get("request")
    response_record = evidence.get("response")
    call_reference = evidence.get("call")
    if (not isinstance(request_record, dict) or not isinstance(response_record, dict)
            or not isinstance(call_reference, dict)):
        raise ValueError("TypeSafe review evidence must bind both request and response files")
    request_path = _evidence_file(evidence_path, request_record, "request")
    response_path = _evidence_file(evidence_path, response_record, "response")
    expected_review_path = str(evidence_path.resolve().relative_to(root.resolve()))
    request_invocation = invocation.get("request")
    response_invocation = invocation.get("response")
    review_invocation = invocation.get("review")
    call_invocation = invocation.get("call")
    if (invocation.get("stage") != stage or invocation.get("revision") != revision["revision"]
            or invocation.get("source_sha256") != _source_sha256(manifest)
            or invocation.get("artifact") != expected["artifact"]
            or invocation.get("requested_model") != delegation["requested_model"]
            or invocation.get("authorization_sha256") != delegation["authorization_sha256"]
            or not isinstance(request_invocation, dict) or request_invocation.get("sha256") != request_record["sha256"]
            or not isinstance(response_invocation, dict) or response_invocation.get("sha256") != response_record["sha256"]
            or not isinstance(review_invocation, dict) or review_invocation.get("path") != expected_review_path
            or review_invocation.get("sha256") != hash_file(evidence_path)
            or not isinstance(call_invocation, dict) or call_reference != call_invocation):
        raise ValueError("TypeSafe evidence does not match its logged live CLI invocation")
    call_record_path = _workflow_file(root, call_invocation, "call record")
    call_record = read_json(call_record_path)
    request_root_path = request_path.resolve().relative_to(root.resolve())
    response_root_path = response_path.resolve().relative_to(root.resolve())
    logged_request = call_record.get("request") if isinstance(call_record, dict) else None
    logged_response = call_record.get("response") if isinstance(call_record, dict) else None
    if (not isinstance(call_record, dict)
            or call_record.get("schema_version") != "workflow.typesafe_invocation.v1"
            or call_record.get("invocation_id") != invocation_id
            or call_record.get("client_correlation_id") != invocation.get("client_correlation_id")
            or call_record.get("correlation_id_origin") != "client-generated; not provider-issued"
            or call_record.get("stage") != stage
            or call_record.get("revision") != revision["revision"]
            or call_record.get("source_sha256") != _source_sha256(manifest)
            or call_record.get("artifact") != expected["artifact"]
            or call_record.get("authorization_sha256") != delegation["authorization_sha256"]
            or not isinstance(logged_request, dict)
            or logged_request.get("workflow_path") != str(request_root_path)
            or logged_request.get("sha256") != request_record["sha256"]
            or not isinstance(logged_response, dict)
            or logged_response.get("workflow_path") != str(response_root_path)
            or logged_response.get("sha256") != response_record["sha256"]
            or call_record.get("requested_model") != delegation["requested_model"]
            or call_record.get("resolved_model") != evidence.get("resolved_model")
            or call_record.get("usage") != evidence.get("usage")
            or invocation.get("usage") != evidence.get("usage")
            or call_record.get("transport") != "https" or call_record.get("status") != "pass"
            or call_record.get("http_status") != 200 or call_record.get("method") != "POST"
            or call_record.get("endpoint") != "https://api.typesafe.ai/v1/systemone"):
        raise ValueError("TypeSafe live invocation record is invalid")
    request = read_json(request_path)
    response = read_json(response_path)
    if not isinstance(request, dict) or request.get("model") != delegation["requested_model"]:
        raise ValueError("TypeSafe request does not use the delegated model")
    binding = request.get("state", {}).get("artifact_binding") if isinstance(request.get("state"), dict) else None
    expected_binding = {"stage": stage, "revision": revision["revision"],
                        "source_sha256": _source_sha256(manifest),
                        "artifact": expected["artifact"]}
    if binding != expected_binding:
        raise ValueError("TypeSafe request is not bound to the current source and artifact revision")
    state = request.get("state")
    if not isinstance(state, dict) or state.get("client_correlation_id") != invocation.get("client_correlation_id"):
        raise ValueError("TypeSafe request does not match its logged client correlation ID")
    questions = request.get("questions")
    question = questions.get("stage_decision") if isinstance(questions, dict) else None
    if not isinstance(question, dict) or question.get("type") != "choice":
        raise ValueError("TypeSafe request does not contain the required stage_decision choice")
    response_model, response_choice, response_usage = _validate_typesafe_response(
        response, set(question["criteria"]))
    if response_model != resolved_model or response_choice != evidence.get("decision"):
        raise ValueError("TypeSafe response does not match the review record")
    if (evidence.get("usage") != response_usage or call_record.get("usage") != response_usage
            or invocation.get("usage") != response_usage):
        raise ValueError("TypeSafe usage metadata does not match the live response")
    answer = response.get("answers", {}).get("stage_decision") if isinstance(response.get("answers"), dict) else None
    if not isinstance(answer, dict) or answer.get("type") != "choice":
        raise ValueError("TypeSafe response is missing the stage_decision choice")
    choice = answer.get("choice")
    decision = evidence.get("decision")
    if choice not in {"approve", "reject", "insufficient"} or decision != choice:
        raise ValueError("TypeSafe decision does not match its response")
    if choice == "insufficient":
        raise ValueError("TypeSafe marked the evidence insufficient; investigate instead of approving")
    probabilities = answer.get("probabilities")
    if not isinstance(probabilities, dict) or choice not in probabilities:
        raise ValueError("TypeSafe response is missing the selected decision probability")
    if any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value)
           or value < 0 or value > 1 for value in probabilities.values()):
        raise ValueError("TypeSafe response contains an invalid decision probability")
    if not math.isclose(sum(probabilities.values()), 1.0, rel_tol=0.0, abs_tol=1e-6):
        raise ValueError("TypeSafe response probabilities do not sum to one")
    probability = float(probabilities[choice])
    if choice == "approve" and probability < delegation["minimum_approval_probability"]:
        raise ValueError("TypeSafe approval is below the delegated confidence threshold")
    return {"path": str(evidence_path), "sha256": hash_file(evidence_path),
            "request": {"path": str(request_path), "sha256": request_record["sha256"]},
            "response": {"path": str(response_path), "sha256": response_record["sha256"]},
            "usage": response_usage,
            "resolved_model": resolved_model, "decision": choice,
            "decision_probability": probability,
            "authorization_sha256": delegation["authorization_sha256"],
            "invocation_id": invocation_id,
            "call": invocation["call"]}


def _workflow_file(root: Path, record: Any, label: str) -> Path:
    if not isinstance(record, dict):
        raise ValueError(f"TypeSafe invocation has an invalid {label} file reference")
    relative = record.get("path")
    digest = record.get("sha256")
    if not isinstance(relative, str) or not relative or not isinstance(digest, str) or len(digest) != 64:
        raise ValueError(f"TypeSafe invocation has an invalid {label} file reference")
    path = (root / relative).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError as exc:
        raise ValueError(f"TypeSafe invocation {label} is outside the workflow") from exc
    if not path.is_file() or hash_file(path) != digest:
        raise ValueError(f"TypeSafe invocation {label} changed or is missing")
    return path


def _validate_typesafe_request(
    request: Any,
    *,
    stage: int,
    revision: dict[str, Any],
    manifest: dict[str, Any],
    delegation: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(request, dict) or request.get("model") != delegation["requested_model"]:
        raise ValueError("TypeSafe request does not use the delegated model")
    state = request.get("state")
    binding = state.get("artifact_binding") if isinstance(state, dict) else None
    artifact = revision["artifact"]
    expected_binding = {"stage": stage, "revision": revision["revision"],
                        "source_sha256": _source_sha256(manifest),
                        "artifact": {"json_sha256": artifact["json_sha256"],
                                     "md_sha256": artifact["md_sha256"]}}
    if binding != expected_binding:
        raise ValueError("TypeSafe request is not bound to the current source and artifact revision")
    questions = request.get("questions")
    question = questions.get("stage_decision") if isinstance(questions, dict) else None
    criteria = question.get("criteria") if isinstance(question, dict) else None
    if (not isinstance(question, dict) or question.get("type") != "choice"
            or not isinstance(question.get("instructions"), (str, dict, list))
            or not isinstance(criteria, dict)
            or not {"approve", "reject", "insufficient"}.issubset(criteria)):
        raise ValueError("TypeSafe request must contain the stage_decision approve/reject/insufficient choice")
    return request


def _validate_typesafe_response(
    response: Any, expected_choices: set[str],
) -> tuple[str, str, dict[str, int]]:
    if not isinstance(response, dict):
        raise ValueError("TypeSafe API response is not a JSON object")
    model = response.get("model")
    answer = response.get("answers", {}).get("stage_decision") if isinstance(response.get("answers"), dict) else None
    if not isinstance(model, str) or not model.startswith("jev-"):
        raise ValueError("TypeSafe API response is missing its resolved Jev model")
    usage = response.get("usage")
    if not isinstance(usage, dict):
        raise ValueError("TypeSafe API response is missing the required usage object")
    token_usage: dict[str, int] = {}
    for key in ("input_tokens", "output_tokens"):
        value = usage.get(key)
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise ValueError(f"TypeSafe API response usage.{key} must be a non-negative integer")
        token_usage[key] = value
    if not isinstance(answer, dict) or answer.get("type") != "choice":
        raise ValueError("TypeSafe API response is missing the stage_decision choice")
    choice = answer.get("choice")
    probabilities = answer.get("probabilities")
    if choice not in {"approve", "reject", "insufficient"} or not isinstance(probabilities, dict):
        raise ValueError("TypeSafe API response has an unsupported stage_decision")
    if set(probabilities) != expected_choices:
        raise ValueError("TypeSafe response probability keys do not match the submitted choice criteria")
    if any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value)
           or value < 0 or value > 1 for value in probabilities.values()):
        raise ValueError("TypeSafe API response contains invalid probabilities")
    if not math.isclose(sum(probabilities.values()), 1.0, rel_tol=0.0, abs_tol=1e-6):
        raise ValueError("TypeSafe API response probabilities do not sum to one")
    if choice not in probabilities:
        raise ValueError("TypeSafe API response omits the selected choice probability")
    return model, choice, token_usage


def _post_typesafe(request: dict[str, Any], api_key: str) -> tuple[int, bytes]:
    payload = json_bytes(request)
    http_request = urllib.request.Request(
        "https://api.typesafe.ai/v1/systemone",
        data=payload,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(http_request, timeout=120) as response:
            return int(response.status), response.read()
    except urllib.error.HTTPError as exc:
        raise ValueError(f"TypeSafe API returned HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise ValueError(f"TypeSafe API request failed: {exc.reason}") from exc


def submit_typesafe_review(
    root: Path,
    stage: int,
    request_path: Path,
) -> dict[str, Any]:
    """Call TypeSafe for the current delegated artifact and log a valid live invocation."""
    root, manifest = load_workflow(root)
    if stage not in range(3, 7) or manifest.get("mode") == "synthetic_fixture":
        raise ValueError("TypeSafe calls are only available for delegated production stages 3..6")
    delegation, error = _delegation_state(root, manifest)
    if error or delegation is None or stage not in delegation.get("stages", []):
        raise ValueError(error or "this stage has no active TypeSafe delegation")
    revision = current_revision(manifest, stage)
    if revision is None:
        raise ValueError(f"stage {stage} has no current artifact")
    if _revision_staleness(root, manifest, stage, revision):
        raise ValueError("cannot submit a stale stage artifact")
    if stage > 1:
        require_stage_approved(root, stage - 1)
    status = stage_status(root, manifest, stage)
    if (status.get("agent") or {}).get("decision") != "approve":
        raise ValueError("the current artifact requires an Agent approval before TypeSafe review")
    if status.get("typesafe") is not None:
        raise ValueError("a TypeSafe decision is already recorded for the current artifact")
    if not request_path.is_file():
        raise ValueError("TypeSafe request file does not exist")
    request = _validate_typesafe_request(read_json(request_path), stage=stage, revision=revision,
                                         manifest=manifest, delegation=delegation)
    api_key = os.environ.get("TYPESAFE_API_KEY")
    if not api_key:
        raise ValueError("TYPESAFE_API_KEY is required for a live TypeSafe call")
    invocation_id = str(uuid.uuid4())
    client_correlation_id = str(uuid.uuid4())
    outbound = json.loads(json.dumps(request))
    outbound.setdefault("state", {})["client_correlation_id"] = client_correlation_id
    http_status, response_bytes = _post_typesafe(outbound, api_key)
    if http_status != 200:
        raise ValueError(f"TypeSafe API returned HTTP {http_status}")
    try:
        response = json.loads(response_bytes)
    except (TypeError, json.JSONDecodeError) as exc:
        raise ValueError("TypeSafe API response is not valid JSON") from exc
    question = request["questions"]["stage_decision"]
    resolved_model, decision, usage = _validate_typesafe_response(
        response, set(question["criteria"]))

    artifact = revision["artifact"]
    invocation_dir = root / "typesafe" / f"stage-{stage}" / f"revision-{revision['revision']:04d}" / invocation_id
    request_out = invocation_dir / "request.json"
    response_out = invocation_dir / "response.json"
    call_out = invocation_dir / "call.json"
    review_out = invocation_dir / "review.json"
    _atomic_write(request_out, json_bytes(outbound))
    _atomic_write(response_out, json_bytes(response))
    request_record = {"path": request_out.name, "sha256": hash_file(request_out)}
    response_record = {"path": response_out.name, "sha256": hash_file(response_out)}
    call_payload = {
        "schema_version": "workflow.typesafe_invocation.v1",
        "invocation_id": invocation_id,
        "client_correlation_id": client_correlation_id,
        "correlation_id_origin": "client-generated; not provider-issued",
        "transport": "https", "method": "POST",
        "endpoint": "https://api.typesafe.ai/v1/systemone", "http_status": http_status,
        "status": "pass", "source_sha256": _source_sha256(manifest),
        "stage": stage, "revision": revision["revision"],
        "artifact": {"json_sha256": artifact["json_sha256"], "md_sha256": artifact["md_sha256"]},
        "authorization_sha256": delegation["authorization_sha256"],
        "requested_model": delegation["requested_model"], "resolved_model": resolved_model,
        "decision": decision,
        # Token counts are retained as API metadata, not treated as proof of provenance.
        "usage": usage,
        "request": {**request_record, "workflow_path": str(request_out.resolve().relative_to(root.resolve()))},
        "response": {**response_record, "workflow_path": str(response_out.resolve().relative_to(root.resolve()))},
        "recorded_at": _now(),
    }
    _atomic_write(call_out, json_bytes(call_payload))
    call_record = {"path": str(call_out.resolve().relative_to(root.resolve())),
                   "sha256": hash_file(call_out)}
    review_payload = {
        "schema_version": "conversion.typesafe_review.v1", "stage": stage,
        "revision": revision["revision"], "source_sha256": _source_sha256(manifest),
        "artifact": {"json_sha256": artifact["json_sha256"], "md_sha256": artifact["md_sha256"]},
        "service": "TypeSafe", "requested_model": delegation["requested_model"],
        "resolved_model": resolved_model, "question_id": "stage_decision", "decision": decision,
        "usage": usage,
        "invocation_id": invocation_id, "call": call_record,
        "request": request_record, "response": response_record,
    }
    _atomic_write(review_out, json_bytes(review_payload))
    invocation = {
        "schema_version": "workflow.typesafe_invocation_log.v1",
        "invocation_id": invocation_id, "client_correlation_id": client_correlation_id,
        "source_sha256": _source_sha256(manifest), "stage": stage,
        "revision": revision["revision"],
        "artifact": review_payload["artifact"],
        "authorization_sha256": delegation["authorization_sha256"],
        "requested_model": delegation["requested_model"], "resolved_model": resolved_model,
        "usage": usage,
        "http_status": http_status, "method": "POST",
        "endpoint": "https://api.typesafe.ai/v1/systemone",
        "request": {"sha256": request_record["sha256"]},
        "response": {"sha256": response_record["sha256"]},
        "review": {"path": str(review_out.resolve().relative_to(root.resolve())),
                   "sha256": hash_file(review_out)},
        "call": call_record,
    }
    manifest.setdefault("typesafe_invocations", []).append(invocation)
    _write_manifest(root, manifest)
    return {"status": "pass", "invocation_id": invocation_id,
            "client_correlation_id": client_correlation_id,
            "correlation_id_origin": "client-generated; not provider-issued",
            "stage": stage, "revision": revision["revision"], "decision": decision,
            "resolved_model": resolved_model, "review": str(review_out),
            "review_sha256": hash_file(review_out), "call_sha256": call_record["sha256"]}


def _evidence_file(evidence_path: Path, record: dict[str, Any], label: str) -> Path:
    relative = record.get("path")
    digest = record.get("sha256")
    if not isinstance(relative, str) or not relative or not isinstance(digest, str) or len(digest) != 64:
        raise ValueError(f"TypeSafe evidence has an invalid {label} file reference")
    path = Path(relative)
    if not path.is_absolute():
        path = evidence_path.parent / path
    path = path.resolve()
    if not path.is_file() or hash_file(path) != digest:
        raise ValueError(f"TypeSafe {label} file changed or is missing")
    return path


def stage_status(
    root: Path, manifest: dict[str, Any], stage: int,
    *, revision_override: dict[str, Any] | None = None,
) -> dict[str, Any]:
    revision = revision_override if revision_override is not None else current_revision(manifest, stage)
    if revision is None:
        return {"stage": stage, "status": "pending", "revision": None, "reasons": ["artifact not produced"]}
    stale = _revision_staleness(root, manifest, stage, revision)
    delegation: dict[str, Any] | None = None
    delegated = False
    if stage >= 3 and manifest.get("review_delegation") is not None:
        delegation, delegation_error = _delegation_state(root, manifest)
        if delegation_error:
            stale.append(delegation_error)
        elif delegation is not None and stage in delegation.get("stages", []):
            delegated = True
    revision_tokens = set(revision.get("resolves_rejections", []))
    for rejected_stage, rejected_revision, receipt in _rejection_receipts(manifest):
        return_to = receipt.get("return_to")
        if isinstance(return_to, int) and return_to <= stage and rejected_stage > stage:
            token = _rejection_token(rejected_stage, rejected_revision, receipt)
            if token not in revision_tokens:
                stale.append(f"stage {rejected_stage} rejection routes to stage {return_to}; a revised checkpoint is required")
    for prior_stage in revision.get("input_stages", {}):
        prior_status = stage_status(root, manifest, int(prior_stage))["status"]
        allowed = {"approved"}
        if manifest.get("mode") == "synthetic_fixture":
            allowed.add("simulated_approved")
        if prior_status not in allowed:
            stale.append(f"stage {prior_stage} is {prior_status}")
    if stale:
        return {"stage": stage, "status": "stale", "revision": revision["revision"], "reasons": stale}
    phase_payload: dict[str, Any] = {}
    if stage == 3:
        try:
            phase_payload = read_json(root / revision["artifact"]["json"])
        except (OSError, ValueError, KeyError, TypeError):
            return {"stage": stage, "status": "stale", "revision": revision["revision"],
                    "reasons": ["stage 3 artifact payload is unreadable"]}
    is_input_boundary = (
        phase_payload.get("schema_version") == "step3.input_boundary.v1"
        and phase_payload.get("phase") == "input_boundary"
    )
    decisions: dict[str, dict[str, Any]] = {}
    proof_stale: list[str] = []
    # The input-boundary subgate always requires an actual human confirmation.
    # A workflow's later Stage 3–6 TypeSafe delegation cannot replace it.
    human_pair_override = delegated and delegation is not None and any(
        item.get("stage") == stage and item.get("revision") == revision["revision"]
        and item.get("reviewer") == "human" and item.get("decision") == "approve"
        and item.get("artifact_sha256") == revision["artifact"]["json_sha256"]
        and item.get("artifact_md_sha256") == revision["artifact"]["md_sha256"]
        and item.get("reviewer_pair_override") == {
            "schema_version": "workflow.review_pair_override.v1",
            "authorization_sha256": delegation["authorization_sha256"],
            "reviewer_pair": ["agent", "human"],
        }
        for item in revision.get("decisions", [])
    )
    reviewer_pair = ("agent", "human") if is_input_boundary or not delegated or human_pair_override else ("agent", "typesafe")
    advisory_typesafe_summary: dict[str, Any] | None = None
    for item in revision.get("decisions", []):
        if (item.get("artifact_sha256") == revision["artifact"]["json_sha256"]
                and item.get("artifact_md_sha256", revision["artifact"]["md_sha256"]) == revision["artifact"]["md_sha256"]):
            reviewer = item.get("reviewer", "")
            if reviewer == "typesafe":
                if is_input_boundary:
                    # Keep an optional TypeSafe recommendation visible, but
                    # never let it replace the human receipt required here.
                    advisory_typesafe_summary = _receipt_summary(item)
                    if not delegated or delegation is None:
                        advisory_typesafe_summary["evidence_status"] = "stale"
                        advisory_typesafe_summary["evidence_message"] = "TypeSafe recommendation has no active delegation"
                        continue
                    try:
                        current_proof = _validate_typesafe_evidence(
                            root, manifest, stage, revision, Path(item["typesafe_evidence"]["path"]), delegation)
                        if current_proof != item.get("typesafe_evidence") or current_proof["decision"] != item.get("decision"):
                            raise ValueError("TypeSafe evidence receipt changed")
                    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
                        advisory_typesafe_summary["evidence_status"] = "stale"
                        advisory_typesafe_summary["evidence_message"] = str(exc)
                    else:
                        advisory_typesafe_summary["evidence_status"] = "current"
                    continue
                if not delegated or delegation is None:
                    proof_stale.append("TypeSafe receipt has no active delegation")
                    continue
                try:
                    current_proof = _validate_typesafe_evidence(
                        root, manifest, stage, revision, Path(item["typesafe_evidence"]["path"]), delegation)
                    if current_proof != item.get("typesafe_evidence") or current_proof["decision"] != item.get("decision"):
                        raise ValueError("TypeSafe evidence receipt changed")
                except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
                    proof_stale.append(f"TypeSafe review evidence changed or is invalid: {exc}")
                    continue
            decisions[reviewer] = item
    if proof_stale:
        return {"stage": stage, "status": "stale", "revision": revision["revision"],
                "reasons": proof_stale}
    rejected = next((item for item in decisions.values() if item.get("decision") == "reject"), None)
    if rejected:
        return {"stage": stage, "status": "rejected", "revision": revision["revision"],
                "return_to": rejected.get("return_to"), "reasons": [rejected.get("message", "rejected")]}
    approved = all(decisions.get(role, {}).get("decision") == "approve" for role in reviewer_pair)
    status = "approved" if approved else "awaiting_confirmation"
    if manifest.get("mode") == "synthetic_fixture" and approved:
        status = "simulated_approved"
    if stage == 3:
        if is_input_boundary:
            if status == "approved":
                status = "input_boundary_confirmed"
            elif status == "simulated_approved":
                status = "simulated_input_boundary_confirmed"
        elif manifest.get("input_boundary_required") and phase_payload.get("schema_version") == "step3.analysis_design.v1":
            boundary_stale = _final_design_boundary_staleness(root, manifest, phase_payload)
            if boundary_stale:
                return {"stage": stage, "status": "stale", "revision": revision["revision"],
                        "reasons": boundary_stale}
    return {"stage": stage, "status": status, "revision": revision["revision"],
            "artifact_sha256": revision["artifact"]["json_sha256"],
            "reviewer_pair": list(reviewer_pair),
            "agent": _receipt_summary(decisions.get("agent")),
            "human": _receipt_summary(decisions.get("human")),
            "typesafe": advisory_typesafe_summary if is_input_boundary else _receipt_summary(decisions.get("typesafe")),
            "reasons": []}


def require_stage_approved(root: Path, stage: int) -> tuple[dict[str, Any], dict[str, Any]]:
    root, manifest = load_workflow(root)
    state = stage_status(root, manifest, stage)
    allowed = {"approved"}
    if manifest.get("mode") == "synthetic_fixture":
        allowed.add("simulated_approved")
    if state["status"] not in allowed:
        if stage == 3 and state["status"] in {"input_boundary_confirmed", "simulated_input_boundary_confirmed"}:
            raise ValueError("stage 3 input-boundary catalog is confirmed but does not release Step 4; a separately reviewed final design is required")
        reviewer_names = state.get("reviewer_pair")
        if reviewer_names:
            reviewer_pair = " and ".join(name.title() for name in reviewer_names)
        else:
            delegation, _error = _delegation_state(root, manifest)
            reviewer_pair = "Agent and TypeSafe" if stage >= 3 and delegation is not None else "Agent and human"
        raise ValueError(f"stage {stage} is {state['status']}; both {reviewer_pair} approval are required")
    revision = current_revision(manifest, stage)
    if stage == 3 and manifest.get("input_boundary_required"):
        payload = read_json(root / revision["artifact"]["json"]) if revision else {}
        if payload.get("schema_version") != "step3.analysis_design.v1":
            raise ValueError("stage 3 input-boundary confirmation does not release Step 4; a separately reviewed final design is required")
        stale = _final_design_boundary_staleness(root, manifest, payload)
        if stale:
            raise ValueError("stage 3 final design is stale: " + "; ".join(stale))
    return manifest, revision  # type: ignore[return-value]


def _input_boundary_revision(manifest: dict[str, Any]) -> dict[str, Any] | None:
    pointer = manifest.get("input_boundary")
    if not isinstance(pointer, dict):
        return None
    revisions = manifest.get("stages", {}).get("3", {}).get("revisions", [])
    revision = next((item for item in revisions if item.get("revision") == pointer.get("revision")), None)
    if revision is None or revision.get("artifact") != pointer.get("artifact"):
        return None
    return revision


def input_boundary_status(root: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    pointer = manifest.get("input_boundary")
    if not isinstance(pointer, dict):
        return {"status": "pending", "revision": None, "reasons": ["input boundary has not been submitted"]}
    revision = _input_boundary_revision(manifest)
    if revision is None:
        return {"status": "stale", "revision": pointer.get("revision"),
                "reasons": ["input boundary pointer does not match its saved revision"]}
    state = stage_status(root, manifest, 3, revision_override=revision)
    if state.get("status") == "approved":
        state["status"] = "input_boundary_confirmed"
    elif state.get("status") == "simulated_approved":
        state["status"] = "simulated_input_boundary_confirmed"
    state["boundary_sha256"] = pointer.get("boundary_sha256")
    return state


def _final_design_boundary_staleness(root: Path, manifest: dict[str, Any], payload: dict[str, Any]) -> list[str]:
    pointer = manifest.get("input_boundary")
    reference = payload.get("input_boundary_reference")
    if not isinstance(pointer, dict) or not isinstance(reference, dict):
        return ["final design does not bind a confirmed input-boundary catalog"]
    expected = {"revision": pointer.get("revision"), "boundary_sha256": pointer.get("boundary_sha256"),
                "artifact": pointer.get("artifact")}
    if any(reference.get(key) != value for key, value in expected.items()):
        return ["a newer input-boundary revision supersedes this final design"]
    state = input_boundary_status(root, manifest)
    if state.get("status") not in {"input_boundary_confirmed", "simulated_input_boundary_confirmed"}:
        return [f"bound input boundary is {state.get('status')}"]
    receipts = {key: state.get(key) for key in ("reviewer_pair", "agent", "human", "typesafe")}
    if reference.get("review_receipts") != receipts:
        return ["confirmed input-boundary review receipts changed after this design was prepared"]
    return []


def record_decision(
    root: Path,
    stage: int,
    reviewer: str,
    decision: str,
    message: str,
    *,
    return_to: int | None = None,
    evidence_path: Path | None = None,
) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    if stage not in STAGES or reviewer not in {"agent", "human", "typesafe"} or decision not in {"approve", "reject"}:
        raise ValueError("use stage 1..6, reviewer agent|human|typesafe, and decision approve|reject")
    if reviewer == "typesafe" and (stage < 3 or manifest.get("mode") == "synthetic_fixture"):
        raise ValueError("TypeSafe review is only available for delegated production stages 3..6")
    if reviewer != "typesafe" and evidence_path is not None:
        raise ValueError("--evidence is only valid for a TypeSafe reviewer")
    if not message.strip():
        raise ValueError("a review message or explicit user response is required")
    revision = current_revision(manifest, stage)
    if revision is None:
        raise ValueError(f"stage {stage} has no current artifact to review")
    stale = _revision_staleness(root, manifest, stage, revision)
    if stale:
        raise ValueError("cannot review stale stage artifact: " + "; ".join(stale))
    delegation: dict[str, Any] | None = None
    typesafe_evidence: dict[str, Any] | None = None
    if reviewer == "typesafe":
        delegation, delegation_error = _delegation_state(root, manifest)
        if delegation_error or delegation is None or stage not in delegation.get("stages", []):
            raise ValueError(delegation_error or "this stage has no active TypeSafe delegation")
        if evidence_path is None:
            raise ValueError("TypeSafe review requires --evidence REVIEW.json")
        typesafe_evidence = _validate_typesafe_evidence(root, manifest, stage, revision, evidence_path, delegation)
        if typesafe_evidence["decision"] != decision:
            raise ValueError("requested decision does not match the TypeSafe response")
    elif reviewer == "human" and stage >= 3:
        delegation, delegation_error = _delegation_state(root, manifest)
        if delegation_error or delegation is None or stage not in delegation.get("stages", []):
            delegation = None
    if decision == "reject" and return_to not in REJECTION_ROUTES[stage]:
        choices = ", ".join(str(value) for value in sorted(REJECTION_ROUTES[stage]))
        raise ValueError(f"stage {stage} rejection must route to one of: {choices}")
    if stage > 1:
        require_stage_approved(root, stage - 1)
    artifact_payload: dict[str, Any] | None = None
    if decision == "approve":
        artifact_payload = read_json(root / revision["artifact"]["json"])
        if isinstance(artifact_payload, dict) and artifact_payload.get("status") in {"blocked", "fail", "failed", "not_ready"}:
            raise ValueError("cannot approve a stage artifact whose status is blocked, failed, or not ready")
        if any(item.get("decision") == "reject" for item in revision.get("decisions", [])):
            raise ValueError("cannot approve an artifact revision that already has a rejection; create a replacement revision")
        if reviewer in {"human", "typesafe"} and not any(
            item.get("reviewer") == "agent"
            and item.get("decision") == "approve"
            and item.get("stage") == stage
            and item.get("revision") == revision["revision"]
            and item.get("artifact_sha256") == revision["artifact"]["json_sha256"]
            and item.get("artifact_md_sha256") == revision["artifact"]["md_sha256"]
            and item.get("input_files", {}) == revision.get("input_files", {})
            and item.get("input_stages", {}) == revision.get("input_stages", {})
            for item in revision.get("decisions", [])
        ):
            raise ValueError("Agent must approve this current JSON/Markdown revision before human or TypeSafe approval")
    receipt = {"stage": stage, "revision": revision["revision"],
               "artifact_sha256": revision["artifact"]["json_sha256"],
               "artifact_md_sha256": revision["artifact"]["md_sha256"],
               "stage_artifact_revision": revision["revision"],
               "input_files": revision.get("input_files", {}), "input_stages": revision.get("input_stages", {}),
               "reviewer": reviewer, "decision": decision, "message": message.strip(),
               "return_to": return_to if decision == "reject" else None, "recorded_at": _now(),
               "simulation": manifest.get("mode") == "synthetic_fixture"}
    if typesafe_evidence is not None:
        receipt["typesafe_evidence"] = typesafe_evidence
    if (reviewer == "human" and decision == "approve" and delegation is not None
            and not (stage == 3 and isinstance(artifact_payload, dict)
                    and artifact_payload.get("schema_version") == "step3.input_boundary.v1"
                    and artifact_payload.get("phase") == "input_boundary")):
        receipt["reviewer_pair_override"] = {
            "schema_version": "workflow.review_pair_override.v1",
            "authorization_sha256": delegation["authorization_sha256"],
            "reviewer_pair": ["agent", "human"],
        }
    revision["decisions"].append(receipt)
    _write_manifest(root, manifest)
    return _receipt_summary(receipt)


def _receipt_summary(receipt: dict[str, Any] | None) -> dict[str, Any] | None:
    if receipt is None:
        return None
    summary = {"receipt_id": hash_bytes(json_bytes(receipt)), "stage": receipt.get("stage"),
               "revision": receipt.get("revision"), "artifact_sha256": receipt.get("artifact_sha256"),
               "artifact_md_sha256": receipt.get("artifact_md_sha256"),
               "reviewer": receipt.get("reviewer"), "decision": receipt.get("decision"),
               "message": receipt.get("message"), "return_to": receipt.get("return_to"),
               "recorded_at": receipt.get("recorded_at"),
               "input_file_count": len(receipt.get("input_files", {})),
               "input_files_sha256": hash_bytes(json_bytes(receipt.get("input_files", {}))),
               "input_stages": receipt.get("input_stages", {})}
    if "reviewer_pair_override" in receipt:
        summary["reviewer_pair_override"] = receipt["reviewer_pair_override"]
    return summary


def _rejection_token(stage: int, revision: int, receipt: dict[str, Any]) -> str:
    return f"{stage}:{revision}:{receipt.get('recorded_at')}:{receipt.get('return_to')}"


def _rejection_receipts(manifest: dict[str, Any]):
    for stage in STAGES:
        revision = current_revision(manifest, stage)
        if revision is None:
            continue
        for receipt in revision.get("decisions", []):
            if receipt.get("decision") == "reject":
                yield stage, revision["revision"], receipt


def _pending_rejection_tokens(manifest: dict[str, Any], stage: int) -> list[str]:
    resolved: list[str] = []
    for rejected_stage, revision, receipt in _rejection_receipts(manifest):
        if receipt.get("return_to") == stage:
            resolved.append(_rejection_token(rejected_stage, revision, receipt))
    return sorted(resolved)


def workflow_status(root: Path) -> dict[str, Any]:
    root, manifest = load_workflow(root)
    stages = [stage_status(root, manifest, stage) for stage in STAGES]
    next_stage: int | None = None
    for stage, state in zip(STAGES, stages):
        if stage == 1 and state["status"] in {"approved", "simulated_approved"}:
            next_stage = 2
            continue
        if state["status"] in {"approved", "simulated_approved"}:
            next_stage = stage + 1 if stage < 6 else None
            continue
        if state["status"] == "rejected":
            next_stage = state.get("return_to")
        else:
            next_stage = stage
        break
    return {"schema_version": "conversion.workflow_status.v1", "workflow": str(root),
            "source": manifest.get("source"), "mode": manifest.get("mode", "production"),
            "stages": stages, "input_boundary": input_boundary_status(root, manifest),
            "next_permitted_stage": next_stage}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()
