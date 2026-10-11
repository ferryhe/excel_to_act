"""Deterministic Step 3 tool-interface export from approved design evidence."""

from __future__ import annotations

import math
import json
from pathlib import Path
from typing import Any

from excel_to_act.steps import conversion_workflow as conversion
from excel_to_act.steps.step3.input_boundary import confirmed_boundary_reference


_SCHEMA_ID = "step3.tool_interface.v1"
_EXPORTER_VERSION = "4"
_TYPES = {"boolean", "integer", "number", "string"}


def export_tool_interface(workflow_dir: Path) -> dict[str, Any]:
    try:
        root, _manifest = conversion.load_workflow(workflow_dir)
        manifest, design_revision = conversion.require_stage_approved(root, 3)
        boundary_reference, boundary_json, _boundary_md = confirmed_boundary_reference(root, manifest)
        boundary_state = conversion.input_boundary_status(root, manifest)
        design_state = conversion.stage_status(root, manifest, 3)
        if (boundary_state.get("status") != "input_boundary_confirmed"
                or boundary_state.get("reviewer_pair") != ["agent", "human"]
                or boundary_state.get("human", {}).get("decision") != "approve"):
            raise ValueError("the current input boundary needs Agent and actual human approval on its exact JSON/Markdown pair")
        boundary_revision = next((item for item in manifest["stages"]["3"]["revisions"]
                                  if item["revision"] == boundary_reference["revision"]), None)
        if boundary_revision is None:
            raise ValueError("the current input-boundary revision is missing")
        _require_current_receipts(boundary_state, boundary_revision)
        _require_current_receipts(design_state, design_revision)

        design_path = _artifact_path(root, design_revision["artifact"].get("json"))
        design = conversion.read_json(design_path)
        boundary = conversion.read_json(boundary_json)
        _check_bindings(manifest, design_revision, design, boundary_reference, boundary, boundary_state)
        support_domain = _support_domain(design.get("design", {}).get("tool_interface"))
        objects, parameters, source_map, blockers = _build_contract(design, boundary, support_domain)
        if support_domain is None:
            blockers.append("supported variable-input trace and producer policy are not approved")
        outputs, output_provenance, output_blockers = _build_outputs(design)
        source_map["outputs"] = output_provenance
        blockers = sorted(set([*blockers, *output_blockers]))
        unknown_facts = _descriptive_unknowns(objects, outputs)
        definition = {
            "schema_version": _SCHEMA_ID,
            "schema_id": _SCHEMA_ID,
            "exporter_version": _EXPORTER_VERSION,
            "status": "interface_generated",
            "operation": "calculate",
            "parameters": parameters,
            "support_domain": support_domain,
            "outputs": outputs,
            "objects": objects,
            "source_provenance": source_map,
            "bindings": {
                "source": {key: manifest["source"].get(key) for key in (
                    "source_id", "run_id", "workbook_sha256")},
                "design": {"revision": design_revision["revision"],
                           "artifact": design_revision["artifact"],
                           "review_receipts": {key: design_state.get(key) for key in
                                               ("reviewer_pair", "agent", "human", "typesafe")}},
                "input_boundary": {"revision": boundary_reference["revision"],
                                   "boundary_sha256": boundary_reference["boundary_sha256"],
                                   "artifact": boundary_reference["artifact"],
                                   "review_receipts": boundary_reference["review_receipts"]},
            },
            "contract_readiness": {"status": "blocked" if blockers else "defined",
                                    "blockers": blockers, "unknown_facts": unknown_facts},
            "runtime_readiness": {
                "status": "blocked",
                "blockers": [*blockers,
                             "Step 3 exports an interface definition only; runtime computation is not implemented."],
            },
        }
        json_content = conversion.json_bytes(definition)
        definition_sha = conversion.hash_bytes(json_content)
        markdown = _interface_markdown(definition, definition_sha)
        md_content = markdown.encode("utf-8")
        relative_dir = Path("tool_interfaces") / definition_sha
        json_relative = relative_dir / "interface.json"
        md_relative = relative_dir / "interface.md"
        json_path = _artifact_path(root, json_relative.as_posix())
        md_path = _artifact_path(root, md_relative.as_posix())
        entry = {
            "schema_version": "step3.tool_interface.registration.v1",
            "schema_id": _SCHEMA_ID,
            "exporter_version": _EXPORTER_VERSION,
            "definition_sha256": definition_sha,
            "artifact": {"json": json_relative.as_posix(),
                         "json_sha256": definition_sha,
                         "md": md_relative.as_posix(),
                         "md_sha256": conversion.hash_bytes(md_content)},
            "bindings": definition["bindings"],
        }
        registrations = manifest.get("tool_interfaces", [])
        if not isinstance(registrations, list) or any(not isinstance(item, dict) for item in registrations):
            raise ValueError("workflow tool_interfaces registry is invalid")
        existing = next((item for item in registrations
                         if item.get("definition_sha256") == definition_sha), None)
        if existing is not None:
            if (existing != entry or not json_path.is_file() or not md_path.is_file()
                    or conversion.hash_file(json_path) != definition_sha
                    or conversion.hash_file(md_path) != entry["artifact"]["md_sha256"]):
                raise ValueError("registered tool interface artifact is missing or changed")
            return _export_result(root, entry, definition)
        for path, content in ((json_path, json_content), (md_path, md_content)):
            if path.exists():
                if path.read_bytes() != content:
                    raise ValueError("unregistered tool interface output has different contents")
            else:
                conversion._atomic_write(path, content)
        manifest.setdefault("tool_interfaces", []).append(entry)
        conversion._write_manifest(root, manifest)
        return _export_result(root, entry, definition)
    except Exception as exc:
        return {"schema_version": _SCHEMA_ID, "tool": "step3.interface", "status": "blocked",
                "diagnostics": [{"code": "tool_interface_export_blocked", "severity": "error",
                                 "message": str(exc)}]}


def _artifact_path(root: Path, relative: Any) -> Path:
    if not isinstance(relative, str) or not relative:
        raise ValueError("workflow artifact path is missing")
    candidate = (root / relative).resolve()
    if not candidate.is_relative_to(root.resolve()):
        raise ValueError("workflow artifact path resolves outside the workflow directory")
    return candidate


def _require_current_receipts(state: dict[str, Any], revision: dict[str, Any]) -> None:
    artifact = revision["artifact"]
    expected_inputs = conversion.hash_bytes(conversion.json_bytes(revision.get("input_files", {})))
    for reviewer in state.get("reviewer_pair", []):
        receipt = state.get(reviewer)
        if not isinstance(receipt, dict) or any((
            receipt.get("stage") != 3,
            receipt.get("revision") != revision["revision"],
            receipt.get("reviewer") != reviewer,
            receipt.get("decision") != "approve",
            receipt.get("artifact_sha256") != artifact["json_sha256"],
            receipt.get("artifact_md_sha256") != artifact["md_sha256"],
            receipt.get("input_files_sha256") != expected_inputs,
            receipt.get("input_stages") != revision.get("input_stages", {}),
        )):
            raise ValueError(f"{reviewer} approval does not bind the exact current revision and inputs")


def _check_bindings(
    manifest: dict[str, Any], design_revision: dict[str, Any], design: Any,
    boundary_reference: dict[str, Any], boundary: Any, boundary_state: dict[str, Any],
) -> None:
    if not isinstance(design, dict) or design.get("schema_version") != "step3.analysis_design.v1":
        raise ValueError("current approved Stage 3 artifact is not a final analysis design")
    if not isinstance(boundary, dict) or boundary.get("schema_version") != "step3.input_boundary.v1":
        raise ValueError("current confirmed input boundary has an unsupported schema")
    source_keys = ("source_id", "run_id", "workbook_sha256")
    source = manifest.get("source", {})
    if any(design.get("source", {}).get(key) != source.get(key)
           or boundary.get("source", {}).get(key) != source.get(key) for key in source_keys):
        raise ValueError("approved design and input boundary do not match the current workflow source")
    if (boundary.get("source", {}).get("analysis_binding_sha256") != design.get("binding_sha256")
            or boundary.get("boundary_sha256") != boundary_reference.get("boundary_sha256")):
        raise ValueError("approved design and input boundary are cross-source or cross-revision")
    expected_boundary = design.get("input_boundary_reference")
    if (not isinstance(expected_boundary, dict)
            or any(expected_boundary.get(key) != boundary_reference.get(key)
                   for key in ("revision", "boundary_sha256", "artifact", "review_receipts"))):
        raise ValueError("approved design does not bind the exact current input-boundary review pair")
    if boundary_state.get("revision") != boundary_reference.get("revision"):
        raise ValueError("input-boundary approval is for a different revision")
    if not isinstance(design_revision.get("artifact"), dict):
        raise ValueError("current design artifact reference is invalid")
    validate_tool_interface_declarations(design.get("design", {}).get("tool_interface"))
    validate_output_interface_declarations(design.get("design", {}).get("targets"))


def _build_contract(design: dict[str, Any], boundary: dict[str, Any],
                    support_domain: dict[str, Any] | None = None) -> tuple[list[dict[str, Any]], dict[str, Any], dict[str, Any], list[str]]:
    variables = boundary.get("variables")
    if not isinstance(variables, list) or not variables:
        raise ValueError("confirmed input boundary has no logical variables")
    declarations = _design_object_declarations(design.get("design", {}).get("tool_interface"))
    logical_names: set[str] = set()
    boundary_ids: set[str] = set()
    axis_lengths = {axis["axis_id"]: axis["shape"][0]
                    for axis in boundary.get("axis_metadata", [])}
    shared_axis_keys: dict[str, tuple[int, bool, tuple[tuple[type, Any], ...]]] = {}
    objects: list[dict[str, Any]] = []
    properties: dict[str, Any] = {}
    required: list[str] = []
    source_map: dict[str, Any] = {}
    blockers: list[str] = []
    for variable in variables:
        if not isinstance(variable, dict):
            raise ValueError("confirmed input boundary contains an invalid logical variable")
        variable_id, logical_name = variable.get("variable_id"), variable.get("logical_name")
        if not isinstance(variable_id, str) or not variable_id or variable_id in boundary_ids:
            raise ValueError("confirmed input boundary variable_id values must be unique non-empty text")
        if not isinstance(logical_name, str) or not logical_name.strip() or logical_name in logical_names:
            raise ValueError("confirmed input boundary logical_name values must be unique non-empty text")
        boundary_ids.add(variable_id)
        logical_names.add(logical_name)
        declaration = declarations.pop(variable_id, {})
        mutability = declaration.get("mutability", "unknown")
        value_type = declaration.get("value_type", "unknown")
        requiredness = declaration.get("required", "unknown")
        default_policy = declaration.get("default_policy", {"status": "unknown"})
        allowed_range = declaration.get("allowed_range", {"status": "unknown"})
        units = declaration.get("units", "unknown")
        unit_basis = declaration.get("unit_basis", "unknown")
        kind, shape, axes = variable.get("kind"), variable.get("shape"), variable.get("axes")
        if kind not in {"scalar", "vector", "table"} or not isinstance(shape, list) or not isinstance(axes, list):
            raise ValueError(f"boundary variable {variable_id} has invalid kind, shape, or axes")
        if (len(shape) != {"scalar": 0, "vector": 1, "table": 2}[kind]
                or len(axes) != len(shape)
                or any(isinstance(item, bool) or not isinstance(item, int) or item <= 0 for item in shape)):
            raise ValueError(f"boundary variable {variable_id} has inconsistent kind, shape, or axes")
        if "axes" in declaration:
            declared_axes = declaration["axes"]
            if len(declared_axes) != len(axes):
                raise ValueError(f"{logical_name}: declared axis rank differs from the boundary")
            for dimension, (declared_axis, boundary_axis) in enumerate(zip(declared_axes, axes)):
                if ({key: value for key, value in declared_axis.items() if key != "keys"}
                        != {key: value for key, value in boundary_axis.items() if key != "keys"}
                        or ("keys" in boundary_axis and json.dumps(declared_axis.get("keys"))
                            != json.dumps(boundary_axis["keys"]))):
                    raise ValueError(f"{logical_name}: declared axis differs from boundary identity")
                if "keys" in declared_axis and len(declared_axis["keys"]) != shape[dimension]:
                    raise ValueError(f"{logical_name}: declared axis key count differs from shape")
            axes = declared_axes
            variable = {**variable, "axes": axes}
        for dimension, axis in enumerate(axes):
            axis_id = axis.get("axis_id")
            if axis_id is not None:
                length = shape[dimension]
                if axis_id in axis_lengths and axis_lengths[axis_id] != length:
                    raise ValueError(f"{logical_name}: axis metadata {axis_id} has conflicting length")
                signature = (length, "keys" in axis,
                             tuple((type(key), key) for key in axis.get("keys", [])))
                previous = shared_axis_keys.get(axis_id)
                if previous is not None and previous != signature:
                    conflict = "length" if previous[0] != length else "keys"
                    raise ValueError(f"{logical_name}: shared axis {axis_id} has conflicting {conflict}")
                shared_axis_keys[axis_id] = signature
        cells = declaration.get("mutable_cells")
        if cells is not None:
            if mutability != "mutable" or not shape or requiredness is not True or default_policy.get("status") != "none":
                raise ValueError(f"{logical_name}: per-cell inputs must be required mutable arrays without a default")
            indices = [tuple(cell["indices"]) for cell in cells]
            if (len(set(indices)) != len(indices) or any(len(index) != len(shape) or any(
                    coordinate >= shape[axis] for axis, coordinate in enumerate(index))
                    for index in indices)):
                raise ValueError(f"{logical_name}: mutable cell indices are duplicate or outside the shape")
            mapped = variable.get("coordinate_map")
            selected = [entry for entry in mapped or [] if isinstance(entry, dict)
                        and isinstance(entry.get("logical_index"), list)
                        and tuple(entry["logical_index"]) in indices]
            if (not isinstance(mapped, list) or len(selected) != len(cells)
                    or len({tuple(entry["logical_index"]) for entry in selected}) != len(cells)
                    or any(not isinstance(entry.get("source_cell"), str)
                           or "!" not in entry["source_cell"] for entry in selected)
                    or len({entry["source_cell"] for entry in selected}) != len(cells)):
                raise ValueError(f"{logical_name}: mutable cells need exact source coordinate mappings")
        item = {"variable_id": variable_id, "logical_name": logical_name, "kind": kind,
                "shape": shape, "axes": axes, "mutability": mutability,
                "value_type": value_type, "required": requiredness, "default_policy": default_policy,
                "allowed_range": allowed_range, "units": units, "unit_basis": unit_basis,
                "allow_blank": declaration.get("allow_blank", False)}
        if cells is not None:
            item["mutable_cells"] = cells
        objects.append(item)
        source_map.setdefault("inputs", {})[variable_id] = {
            key: variable[key] for key in ("logical_name", "group_id", "role", "logical_role",
                                           "consumer_targets", "source_extents", "value_source_policy")
            if key in variable}
        if mutability == "mutable":
            schema = _parameter_schema(variable, declaration)
            properties[logical_name] = schema
            if requiredness is True:
                required.append(logical_name)
            elif requiredness is False and default_policy.get("status") != "saved":
                blockers.append(f"{logical_name}: optional mutable input needs an approved saved default")
        if mutability == "unknown":
            blockers.append(f"{logical_name}: mutability is unknown")
        if value_type == "unknown" and mutability == "mutable" and cells is None:
            blockers.append(f"{logical_name}: value type is unknown")
        if requiredness == "unknown":
            blockers.append(f"{logical_name}: required policy is unknown")
        if default_policy.get("status") == "unknown":
            blockers.append(f"{logical_name}: default policy is unknown")
        if allowed_range.get("status") == "unknown" and mutability == "mutable" and cells is None:
            blockers.append(f"{logical_name}: allowed range is unknown")
        if (variable.get("value_source_policy", {}).get("available_now") is False
                and variable_id not in (support_domain or {}).get("external_policy", {})):
            blockers.append(f"{logical_name}: approved current external value is unavailable")
    if declarations:
        raise ValueError("tool_interface declares boundary variables that are not in the current input boundary: "
                         + ", ".join(sorted(declarations)))
    by_id = {item["variable_id"]: item for item in objects}
    for item in objects:
        policy = item["allowed_range"]
        if policy["status"] == "dependent_enum":
            dependency = by_id.get(policy["variable_id"])
            if (item["shape"] or dependency is None or dependency is item
                    or dependency["mutability"] != "mutable" or dependency["shape"]
                    or dependency["value_type"] != item["value_type"]
                    or dependency["allowed_range"]["status"] in {"unknown", "dependent_enum"}):
                raise ValueError(f"{item['logical_name']}: dependent enum needs a mutable scalar of the same type")
    objects.sort(key=lambda item: (item["logical_name"].casefold(), item["variable_id"]))
    source_map["inputs"] = {key: source_map.get("inputs", {})[key]
                             for key in sorted(source_map.get("inputs", {}))}
    parameters = {"type": "object", "properties": properties,
                  "required": sorted(required), "additionalProperties": False}
    return objects, parameters, source_map, sorted(set(blockers))


def _build_outputs(design: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], list[str]]:
    targets = design.get("design", {}).get("targets")
    if not isinstance(targets, list) or not targets:
        raise ValueError("approved design has no logical output targets")
    properties: dict[str, Any] = {}
    provenance: dict[str, Any] = {}
    blockers: list[str] = []
    for target in targets:
        if not isinstance(target, dict) or not isinstance(target.get("target_id"), str):
            raise ValueError("approved design contains an invalid output target")
        name = target["target_id"]
        shape = target.get("shape")
        if not isinstance(shape, list):
            raise ValueError(f"output target {name} has no declared shape")
        value_type = target.get("value_type", "unknown")
        allowed_range = target.get("allowed_range", {"status": "unknown"})
        axes = target.get("axes")
        if axes is None:
            axes = [] if not shape else "unknown"
        # Legacy units text may describe an open question; only the paired explicit basis confirms it.
        units = target.get("units") if isinstance(target.get("unit_basis"), str) and isinstance(
            target.get("units"), str) else "unknown"
        unit_basis = target.get("unit_basis") if isinstance(target.get("unit_basis"), str) else "unknown"
        properties[name] = _shape_schema(shape, value_type, allowed_range)
        properties[name].update({"x-shape": shape, "x-axes": axes, "x-units": units,
                                 "x-unit-basis": unit_basis, "x-allowed-range": allowed_range})
        if value_type == "unknown":
            blockers.append(f"output {name}: value type is unknown")
        if axes == "unknown":
            blockers.append(f"output {name}: axes are unknown")
        provenance[name] = {key: target[key] for key in ("selector", "location", "units",
                                                       "diagnostic_intermediates") if key in target}
    return ({"type": "object", "properties": properties,
             "required": sorted(properties), "additionalProperties": False}, provenance, blockers)


def _descriptive_unknowns(objects: list[dict[str, Any]], outputs: dict[str, Any]) -> list[str]:
    facts = []
    for item in objects:
        name = item["logical_name"]
        for field, label in (("units", "units"), ("unit_basis", "unit basis")):
            if item[field] == "unknown":
                facts.append(f"{name}: {label} are unknown" if field == "units"
                             else f"{name}: {label} is unknown")
        if item["mutability"] == "immutable" or "mutable_cells" in item:
            if item["value_type"] == "unknown":
                facts.append(f"{name}: value type is unknown")
            if item["allowed_range"]["status"] == "unknown":
                facts.append(f"{name}: allowed range is unknown")
    for name, schema in outputs["properties"].items():
        if schema["x-units"] == "unknown":
            facts.append(f"output {name}: units are unknown")
        if schema["x-unit-basis"] == "unknown":
            facts.append(f"output {name}: unit basis is unknown")
        if schema["x-allowed-range"]["status"] == "unknown":
            facts.append(f"output {name}: allowed range is unknown")
    return sorted(set(facts))


def _design_object_declarations(value: Any) -> dict[str, dict[str, Any]]:
    if value is None:
        return {}
    _reject_fields(value, {"objects", "support_domain"}, "design.tool_interface")
    objects = value.get("objects")
    if not isinstance(objects, list):
        raise ValueError("design.tool_interface.objects must be a list")
    result: dict[str, dict[str, Any]] = {}
    allowed = {"variable_id", "mutability", "value_type", "required", "default_policy",
               "allowed_range", "units", "unit_basis", "allow_blank", "mutable_cells", "axes"}
    for index, declaration in enumerate(objects):
        _reject_fields(declaration, allowed, f"design.tool_interface.objects[{index}]")
        variable_id = declaration.get("variable_id")
        if not isinstance(variable_id, str) or not variable_id or variable_id in result:
            raise ValueError("tool_interface variable_id values must be unique non-empty text")
        if declaration.get("mutability") not in {"mutable", "immutable", "unknown"}:
            raise ValueError(f"tool_interface {variable_id} must explicitly declare mutability")
        value_type = declaration.get("value_type", "unknown")
        if value_type not in _TYPES | {"unknown"}:
            raise ValueError(f"tool_interface {variable_id} value_type must be boolean, integer, number, string, or unknown")
        if "required" in declaration and not isinstance(declaration["required"], bool):
            raise ValueError(f"tool_interface {variable_id} required must be boolean")
        if "allow_blank" in declaration and not isinstance(declaration["allow_blank"], bool):
            raise ValueError(f"tool_interface {variable_id} allow_blank must be boolean")
        if "axes" in declaration:
            axes = declaration["axes"]
            if not isinstance(axes, list):
                raise ValueError(f"tool_interface {variable_id} axes must be a list")
            for axis in axes:
                _reject_fields(axis, {"name", "role", "axis_id", "keys"},
                               f"tool_interface {variable_id} axis")
                if "keys" in axis:
                    keys = axis["keys"]
                    if (not isinstance(keys, list) or any(
                            key is not None and type(key) not in {bool, int, float, str}
                            or type(key) is float and not math.isfinite(key) for key in keys)):
                        raise ValueError(f"tool_interface {variable_id} axis keys must be finite JSON scalars")
                    if len({(type(key), key) for key in keys}) != len(keys):
                        raise ValueError(f"tool_interface {variable_id} axis keys contain duplicates")
        for key in ("units", "unit_basis"):
            if key in declaration and (not isinstance(declaration[key], str) or not declaration[key].strip()):
                raise ValueError(f"tool_interface {variable_id} {key} must be non-empty text")
        default_policy = declaration.get("default_policy", {"status": "unknown"})
        _validate_default_policy(default_policy, variable_id)
        allowed_range = declaration.get("allowed_range", {"status": "unknown"})
        _validate_allowed_range(allowed_range, value_type, variable_id)
        if "mutable_cells" in declaration:
            cells = declaration["mutable_cells"]
            if not isinstance(cells, list) or not cells:
                raise ValueError(f"tool_interface {variable_id} mutable_cells must be nonempty")
            for cell in cells:
                _reject_fields(cell, {"indices", "value_type", "allowed_range", "allow_blank"},
                               f"tool_interface {variable_id} mutable cell")
                indices = cell.get("indices")
                if not isinstance(indices, list) or not indices or any(
                        type(index) is not int or index < 0 for index in indices):
                    raise ValueError(f"tool_interface {variable_id} mutable cell needs nonnegative indices")
                if cell.get("value_type") not in _TYPES or not isinstance(cell.get("allow_blank"), bool):
                    raise ValueError(f"tool_interface {variable_id} mutable cell needs type and blank policy")
                _validate_allowed_range(cell.get("allowed_range"), cell["value_type"], variable_id)
                if cell["allowed_range"]["status"] == "unknown":
                    raise ValueError(f"tool_interface {variable_id} mutable cell needs an approved range policy")
                if cell["allowed_range"]["status"] == "dependent_enum":
                    raise ValueError(f"tool_interface {variable_id} dependent enum is scalar-only")
        result[variable_id] = {**declaration, "value_type": value_type,
                               "default_policy": default_policy, "allowed_range": allowed_range}
    return result


def _validate_default_policy(value: Any, variable_id: str) -> None:
    if not isinstance(value, dict) or value.get("status") not in {"unknown", "none", "saved"}:
        raise ValueError(f"tool_interface {variable_id} default_policy must be unknown, none, or approved saved; "
                         "literal defaults need later approved value capture")
    _reject_fields(value, {"status"}, f"tool_interface {variable_id} default_policy")


def _validate_allowed_range(value: Any, value_type: str, variable_id: str) -> None:
    if not isinstance(value, dict) or value.get("status") not in {"unknown", "unbounded", "declared", "enum", "dependent_enum"}:
        raise ValueError(f"tool_interface {variable_id} allowed_range needs an explicit status")
    status = value["status"]
    allowed = ({"status", "minimum", "maximum"} if status == "declared" else
               {"status", "values", "variable_id"} if status == "dependent_enum" else
               {"status", "values"} if status == "enum" else {"status"})
    _reject_fields(value, allowed, f"tool_interface {variable_id} allowed_range")
    if status in {"enum", "dependent_enum"}:
        if status == "dependent_enum" and (not isinstance(value.get("variable_id"), str)
                                            or not value["variable_id"]):
            raise ValueError(f"tool_interface {variable_id} dependent enum needs a variable_id")
        values = value.get("values")
        if (not isinstance(values, list) or not values
                or any(item is not None and type(item) not in {bool, int, float, str} for item in values)
                or any(type(item) is float and not math.isfinite(item) for item in values)
                or any(item is not None and not (
                    value_type == "boolean" and type(item) is bool
                    or value_type == "integer" and type(item) is int
                    or value_type == "number" and type(item) in {int, float}
                    or value_type == "string" and type(item) is str) for item in values)
                or len({json.dumps(item, sort_keys=True) for item in values}) != len(values)):
            raise ValueError(f"tool_interface {variable_id} enum needs distinct finite scalar values")
        return
    if status != "declared":
        return
    if value_type not in {"integer", "number"}:
        raise ValueError(f"tool_interface {variable_id} numeric allowed_range needs a numeric value_type")
    bounds = [value[key] for key in ("minimum", "maximum") if key in value]
    if not bounds:
        raise ValueError(f"tool_interface {variable_id} declared allowed_range needs a bound")
    if any(isinstance(bound, bool) or not isinstance(bound, (int, float)) or not math.isfinite(bound)
           for bound in bounds):
        raise ValueError(f"tool_interface {variable_id} range bounds must be finite numbers")
    if "minimum" in value and "maximum" in value and value["minimum"] > value["maximum"]:
        raise ValueError(f"tool_interface {variable_id} range minimum exceeds maximum")


def _parameter_schema(variable: dict[str, Any], declaration: dict[str, Any]) -> dict[str, Any]:
    kind, shape = variable["kind"], variable["shape"]
    value_type = declaration.get("value_type", "unknown")
    allowed_range = declaration.get("allowed_range", {"status": "unknown"})
    cells = declaration.get("mutable_cells")
    schema = ({"type": "array", "minItems": len(cells), "maxItems": len(cells),
               "items": {"type": "object", "properties": {"indices": {"type": "array"},
                                                      "value": {}},
                         "required": ["indices", "value"], "additionalProperties": False}}
              if cells is not None else _shape_schema(shape, value_type, allowed_range))
    schema.update({"x-kind": kind, "x-shape": shape, "x-axes": variable["axes"],
                   "x-mutability": declaration["mutability"],
                   "x-required-policy": declaration.get("required", "unknown"),
                   "x-default-policy": declaration.get("default_policy", {"status": "unknown"}),
                   "x-allowed-range": allowed_range,
                   "x-units": declaration.get("units", "unknown"),
                   "x-unit-basis": declaration.get("unit_basis", "unknown"),
                   "x-allow-blank": declaration.get("allow_blank", False)})
    if cells is not None:
        schema["x-mutable-cells"] = cells
    return schema


def _shape_schema(shape: list[int], value_type: str, allowed_range: dict[str, Any]) -> dict[str, Any]:
    leaf: dict[str, Any] = {"type": value_type} if value_type in _TYPES else {}
    if allowed_range.get("status") == "declared":
        leaf.update({key: allowed_range[key] for key in ("minimum", "maximum") if key in allowed_range})
    schema = leaf
    for size in reversed(shape):
        schema = {"type": "array", "minItems": size, "maxItems": size, "items": schema}
    return schema


def validate_tool_interface_declarations(value: Any) -> None:
    """Validate the closed, optional per-object contract annotations in a design input."""
    _design_object_declarations(value)
    _support_domain(value)


def _support_domain(value: Any) -> dict[str, Any] | None:
    if value is None or value.get("support_domain") is None:
        return None
    domain = value["support_domain"]
    _reject_fields(domain, {"status", "active_trace_sha256", "trace_binding", "external_policy",
                            "tolerances", "cases"},
                   "design.tool_interface.support_domain")
    if domain.get("status") not in {"fixed_trace", "exact_cases"}:
        raise ValueError("variable support domain needs an approved active-trace SHA-256")
    if "trace_binding" in domain:
        if (domain["trace_binding"] != "post_discovery" or domain["status"] != "exact_cases"
                or "active_trace_sha256" in domain):
            raise ValueError("post-discovery exact cases cannot declare an active-trace SHA-256")
    elif (not isinstance(domain.get("active_trace_sha256"), str)
          or len(domain["active_trace_sha256"]) != 64 or not all(
              char in "0123456789abcdef" for char in domain["active_trace_sha256"])):
        raise ValueError("variable support domain needs an approved active-trace SHA-256")
    if domain["status"] == "exact_cases":
        cases = domain.get("cases")
        if not isinstance(cases, list) or not cases:
            raise ValueError("exact-case support needs bound trace and request artifacts")
        for case in cases:
            if not isinstance(case, dict) or set(case) not in ({
                    "trace", "raw_request", "public_request", "producer"}, {
                    "trace", "raw_request", "public_request", "producer", "plan", "profile"}):
                raise ValueError("exact-case artifact references are incomplete")
            for reference in case.values():
                if (not isinstance(reference, dict) or set(reference) != {"path", "sha256"}
                        or not isinstance(reference["path"], str) or not reference["path"]
                        or not isinstance(reference["sha256"], str)
                        or len(reference["sha256"]) != 64 or not all(
                            char in "0123456789abcdef" for char in reference["sha256"])):
                    raise ValueError("exact-case artifact path or SHA-256 is malformed")
    elif "cases" in domain:
        raise ValueError("fixed-trace support cannot declare exact-case artifacts")
    policies = domain.get("external_policy", {})
    tolerances = domain.get("tolerances")
    if (not isinstance(tolerances, dict) or set(tolerances) != {"absolute", "relative"}
            or any(type(number) not in {int, float} or not math.isfinite(number) or number < 0
                   for number in tolerances.values())):
        raise ValueError("variable support domain needs approved absolute and relative tolerances")
    if not isinstance(policies, dict):
        raise ValueError("external producer policy must name each same-request input")
    for key, policy in policies.items():
        _reject_fields(policy, {"mode", "source_sha256", "version"}, f"external producer {key}")
        if (not isinstance(key, str) or policy.get("mode") != "same_request"
                or not isinstance(policy.get("source_sha256"), str) or len(policy["source_sha256"]) != 64
                or not isinstance(policy.get("version"), str) or not policy["version"]):
            raise ValueError("external producer needs an approved same-request source and version")
    return domain


def validate_output_interface_declarations(targets: Any) -> None:
    if not isinstance(targets, list):
        raise ValueError("design targets must be a list")
    for target in targets:
        if not isinstance(target, dict):
            continue
        target_id = target.get("target_id", "unknown")
        value_type = target.get("value_type", "unknown")
        if value_type not in _TYPES | {"unknown"}:
            raise ValueError(f"target {target_id} value_type must be boolean, integer, number, string, or unknown")
        _validate_allowed_range(target.get("allowed_range", {"status": "unknown"}),
                                value_type, str(target_id))
        for key in ("units", "unit_basis"):
            if key in target and target[key] is not None and (
                    not isinstance(target[key], str) or not target[key].strip()):
                raise ValueError(f"target {target_id} {key} must be non-empty text or null")
        if "axes" in target and target["axes"] is not None:
            if not isinstance(target["axes"], list) or any(not isinstance(axis, dict) for axis in target["axes"]):
                raise ValueError(f"target {target_id} axes must be a list of axis objects or null")
            if isinstance(target.get("shape"), list) and len(target["axes"]) != len(target["shape"]):
                raise ValueError(f"target {target_id} axes rank does not match its shape")


def _reject_fields(value: Any, allowed: set[str], label: str) -> None:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be an object")
    extra = sorted(set(value) - allowed)
    if extra:
        raise ValueError(f"{label} contains unsupported fields: {', '.join(extra)}")


def _interface_markdown(definition: dict[str, Any], digest: str) -> str:
    parameters = definition["parameters"]["properties"]
    blockers = definition["contract_readiness"]["blockers"]
    unknown_facts = definition["contract_readiness"]["unknown_facts"]
    lines = ["# Step 3 Tool Interface", "", "Status: **Interface generated**", "",
             f"Schema ID: `{definition['schema_id']}`",
             f"Definition SHA-256: `{digest}`",
             f"Mutable logical parameters: {len(parameters)}",
             f"Logical objects: {len(definition['objects'])}",
             f"Contract readiness: **{definition['contract_readiness']['status']}**",
             "Runtime readiness: **blocked**", "",
             "Defined means the execution contract is complete; it does not settle descriptive business facts.",
             "This artifact defines logical-name parameters only. It does not perform parameterized computation, "
             "establish Excel equivalence, or establish production readiness.", "",
             "## Logical objects", "",
             "| Logical name | Kind | Shape | Mutability | Type | Required | Default | Allowed range |",
             "| --- | --- | --- | --- | --- | --- | --- | --- |"]
    for item in definition["objects"]:
        cells = [item["logical_name"], item["kind"], json.dumps(item["shape"], ensure_ascii=False),
                 item["mutability"], item["value_type"], str(item["required"]),
                 json.dumps(item["default_policy"], ensure_ascii=False, sort_keys=True),
                 json.dumps(item["allowed_range"], ensure_ascii=False, sort_keys=True)]
        lines.append("| " + " | ".join(cell.replace("|", "\\|") for cell in cells) + " |")
    if blockers:
        lines.extend(["", "## Execution blockers", "", *[f"- {item}" for item in blockers]])
    if unknown_facts:
        lines.extend(["", "## Unknown descriptive facts", "", *[f"- {item}" for item in unknown_facts]])
    return "\n".join(lines) + "\n"


def _export_result(root: Path, entry: dict[str, Any], definition: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": _SCHEMA_ID, "tool": "step3.interface", "status": "pass",
            "message": "interface generated", "workflow": str(root),
            "schema_id": entry["schema_id"], "artifact": entry["artifact"],
            "definition_sha256": entry["definition_sha256"],
            "contract_readiness": definition["contract_readiness"],
            "runtime_readiness": definition["runtime_readiness"]}
