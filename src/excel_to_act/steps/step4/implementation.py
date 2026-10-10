"""Validate and bind a Step 4 implementation map before bundle generation."""

from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Any

from excel_to_act.steps.conversion_workflow import hash_file, load_workflow, read_json, require_stage_approved
from excel_to_act.steps.step3.calculation import CalculationBlocked, _parse_reference
from excel_to_act.steps.step4.external_inputs import load_current_external_bundle
from excel_to_act.steps.step4.modular_bundle import (
    _raw_bindings,
    _source_index,
    add_adapter_metadata,
    _named_bindings,
)
from excel_to_act.steps.step4.modular_compiler import (
    build_source_cell_index,
    compile_mapped_families,
    render_grouped_driver,
    validate_family_coverage,
)


def _canonical_reference(value: str) -> tuple[str, int, int, int, int] | None:
    reference = _parse_reference(value, "Main")
    if reference is None or not reference.sheet:
        return None
    return (reference.sheet.casefold(), reference.min_row, reference.min_col,
            reference.max_row, reference.max_col)


def _apply_family_execution_refinements(
    plan: dict[str, Any], implementation: dict[str, Any], trace: dict[str, Any],
) -> list[dict[str, Any]]:
    """Apply explicit same-index refinements only when the current trace proves them."""
    execution = implementation.get("execution_plan", {})
    requested = execution.get("family_execution", []) if isinstance(execution, dict) else None
    if not isinstance(requested, list):
        raise CalculationBlocked("implementation family_execution must be a list")
    if not requested:
        return []

    mappings = {item.get("source_family_id"): item
                for item in plan.get("source_mappings", []) if isinstance(item, dict)}
    variables = plan.get("variables", {})
    if not isinstance(variables, dict):
        raise CalculationBlocked("semantic plan has no logical variables for schedule proof")
    source_index = build_source_cell_index(plan)
    formula_addresses = {
        _canonical_reference(item.get("address"))
        for item in trace.get("cells", []) if isinstance(item, dict)
        and item.get("role") in {"calculated_formula", "calculated_array_formula"}
    }
    formula_addresses.discard(None)
    edges_by_consumer: dict[tuple[str, int, int, int, int], list[dict[str, Any]]] = {}
    for edge in trace.get("edges", []):
        if not isinstance(edge, dict):
            continue
        consumer = _canonical_reference(edge.get("consumer", ""))
        if consumer is not None:
            edges_by_consumer.setdefault(consumer, []).append(edge)
    requested_ids: set[str] = set()
    proof: list[dict[str, Any]] = []
    for item in requested:
        if not isinstance(item, dict):
            raise CalculationBlocked("implementation family execution entry is malformed")
        family_id, new_execution = item.get("source_family_id"), item.get("execution")
        if (not isinstance(family_id, str) or family_id not in mappings
                or family_id in requested_ids or new_execution != "vectorized"):
            raise CalculationBlocked("only unique mapped families may request vectorized schedule refinement")
        requested_ids.add(family_id)
        mapping = mappings[family_id]
        variable_id = mapping.get("variable_id")
        variable = variables.get(variable_id)
        if not isinstance(variable, dict) or not isinstance(variable.get("axes"), list):
            raise CalculationBlocked(f"schedule refinement has no logical axis: {family_id}")
        segments = [segment for segment in variable.get("equation_segments", [])
                    if family_id in segment.get("source_family_ids", [])]
        if len(segments) != 1 or segments[0].get("execution") != "ascending_recurrence":
            raise CalculationBlocked(f"schedule refinement must name one declared recurrence segment: {family_id}")
        segment = segments[0]
        axis_name = segment.get("index_axis")
        dimension = next((index for index, axis in enumerate(variable["axes"])
                          if axis.get("name") == axis_name), None)
        if dimension is None:
            raise CalculationBlocked(f"schedule refinement axis is not declared: {family_id}")
        members = mapping.get("source_members", [])
        if not isinstance(members, list) or not members:
            raise CalculationBlocked(f"schedule refinement has no current active members: {family_id}")

        same_year_edges = 0
        direct_formula_edges = 0
        for member in members:
            owner = _canonical_reference(member)
            if owner is None:
                raise CalculationBlocked(f"schedule refinement member is not a qualified cell: {member}")
            owner_cells = source_index.get((owner[0], owner[1], owner[2]), [])
            owner_matches = [cell for cell in owner_cells if cell.variable_id == variable_id]
            if len({(cell.indices, cell.axes) for cell in owner_matches}) != 1:
                raise CalculationBlocked(f"schedule refinement member has no unique logical coordinate: {member}")
            owner_indices = next(iter(owner_matches)).indices
            if len(owner_indices) <= dimension:
                raise CalculationBlocked(f"schedule refinement output is missing its time coordinate: {member}")
            owner_index = owner_indices[dimension]
            for edge in edges_by_consumer.get(owner, []):
                relationship = edge.get("relationship")
                if relationship != "cell_reference":
                    raise CalculationBlocked(
                        f"vectorized schedule refinement requires direct cell edges only: {member}")
                prerequisite = _canonical_reference(edge.get("prerequisite", ""))
                if prerequisite not in formula_addresses:
                    continue
                direct_formula_edges += 1
                if prerequisite is None or prerequisite[1] != prerequisite[3] or prerequisite[2] != prerequisite[4]:
                    raise CalculationBlocked(f"schedule refinement prerequisite is not a formula cell: {member}")
                prereq_cells = source_index.get((prerequisite[0], prerequisite[1], prerequisite[2]), [])
                formula_matches = [cell for cell in prereq_cells
                                   if cell.role not in {"source_value", "adapter_metadata"}]
                if not formula_matches:
                    raise CalculationBlocked(f"formula prerequisite has no reviewed logical coordinate: {edge['prerequisite']}")
                for prereq_cell in formula_matches:
                    prereq_variable = variables.get(prereq_cell.variable_id, {})
                    if prereq_variable.get("role") == "external":
                        raise CalculationBlocked("vectorized schedule refinement cannot depend on an external formula boundary")
                    prereq_axes = prereq_variable.get("axes", [])
                    prereq_dimension = next((index for index, axis in enumerate(prereq_axes)
                                             if axis.get("name") == axis_name), None)
                    if prereq_dimension is None:
                        if prereq_variable.get("shape"):
                            raise CalculationBlocked(
                                f"formula prerequisite uses an unrelated logical axis: {edge['prerequisite']}")
                        continue
                    if len(prereq_cell.indices) <= prereq_dimension:
                        raise CalculationBlocked(f"formula prerequisite is missing its time coordinate: {edge['prerequisite']}")
                    prereq_index = prereq_cell.indices[prereq_dimension]
                    if prereq_index != owner_index:
                        relation = "prior-period" if prereq_index < owner_index else "future-period"
                        raise CalculationBlocked(
                            f"vectorized schedule refinement has a {relation} formula edge: {member} -> {edge['prerequisite']}")
                    same_year_edges += 1

        # This changes only an execution schedule, never an equation or logical variable.
        segment["execution"] = "vectorized"
        segment["recurrence_group"] = None
        segment["snapshot_before_update"] = False
        segment["update_order"] = "same_index_dependencies"
        proof.append({"source_family_id": family_id, "variable_id": variable_id,
                      "segment_id": segment.get("segment_id"), "axis": axis_name,
                      "old_execution": "ascending_recurrence", "new_execution": "vectorized",
                      "active_member_count": len(members),
                      "direct_formula_edge_count": direct_formula_edges,
                      "same_year_formula_edge_count": same_year_edges,
                      "cross_time_formula_edge_count": 0,
                      "trace_sha256": trace.get("active_trace_sha256")})
    return proof


def _current_trace(root: Path, stage3_revision: dict[str, Any], source_sha256: str) -> tuple[Path, dict[str, Any]]:
    discovery_root = root / "stage4" / "discovery"
    paths = sorted(discovery_root.glob("revision-*/discovery.json"), reverse=True) \
        if discovery_root.exists() else []
    for discovery_path in paths:
        try:
            discovery = read_json(discovery_path)
            trace_path = Path(discovery["active_trace_path"]).resolve()
            if (discovery.get("schema_version") == "step4.discovery.v1"
                    and discovery.get("status") == "pass"
                    and discovery.get("stage3_artifact_sha256") == stage3_revision["artifact"]["json_sha256"]
                    and discovery.get("source_sha256") == source_sha256
                    and trace_path.is_file()
                    and hash_file(trace_path) == discovery.get("active_trace_sha256")):
                return trace_path, read_json(trace_path)
        except (KeyError, OSError, ValueError, TypeError):
            continue
    raise CalculationBlocked("no current passing Step 4 discovery trace exists")


def _active_projection(
    semantic_plan: dict[str, Any], source_profile: dict[str, Any], trace: dict[str, Any],
    implementation: dict[str, Any], catalog: dict[str, Any], model_scope: dict[str, Any],
    source_path: Path, source_sha256: str,
) -> tuple[dict[str, Any], dict[str, Any], dict[str, int]]:
    cells = trace.get("cells")
    if not isinstance(cells, list):
        raise CalculationBlocked("active trace has no cell list")
    active_formulas = {
        _canonical_reference(item.get("address")): item.get("formula")
        for item in cells if isinstance(item, dict) and item.get("role") == "calculated_formula"
    }
    active_formulas.pop(None, None)
    active_followers = {
        _canonical_reference(item.get("address")): item
        for item in cells if isinstance(item, dict) and item.get("role") == "calculated_array_formula"
    }
    active_followers.pop(None, None)
    candidate_trace = read_json(Path(semantic_plan["evidence"]["source_candidate_trace"]["path"]))
    candidates = {
        _canonical_reference(item.get("address")): item.get("formula")
        for item in candidate_trace.get("cells", []) if isinstance(item, dict)
        and item.get("role") == "calculated_formula"
    }
    candidates.pop(None, None)
    mapped_candidates = {
        _canonical_reference(address)
        for mapping in semantic_plan.get("source_mappings", [])
        for address in mapping.get("source_members", [])
    }
    if None in mapped_candidates:
        raise CalculationBlocked("reviewed formula map contains an invalid source address")
    if not set(active_formulas).issubset(mapped_candidates):
        raise CalculationBlocked("fresh active formula lies outside the reviewed source-family candidates")
    if not set(active_formulas).issubset(set(candidates)):
        raise CalculationBlocked("fresh active formula is absent from the bound candidate trace")
    for address, formula in active_formulas.items():
        if not isinstance(formula, str) or candidates[address] != formula:
            raise CalculationBlocked("fresh formula text differs from the reviewed candidate trace")

    projected = copy.deepcopy(semantic_plan)
    active_family_ids: set[str] = set()
    projected_mappings = []
    for mapping in projected.get("source_mappings", []):
        members = [address for address in mapping.get("source_members", [])
                   if _canonical_reference(address) in active_formulas]
        if not members:
            continue
        mapping["source_members"] = members
        mapping["member_count"] = len(members)
        active_family_ids.add(mapping["source_family_id"])
        projected_mappings.append(mapping)
    projected["source_mappings"] = projected_mappings

    active_array_keys = set(active_followers)
    projected_arrays = []
    for family in projected.get("array_mappings", []):
        instances = []
        for instance in family.get("instances", []):
            anchor = _canonical_reference(instance.get("anchor", ""))
            members = [address for address in instance.get("source_members", [])
                       if _canonical_reference(address) in active_array_keys]
            if anchor not in active_formulas and not members:
                continue
            if anchor not in active_formulas:
                raise CalculationBlocked("active array followers are missing their calculated anchor")
            instance["source_members"] = members
            instance["member_count"] = len(members)
            instances.append(instance)
        if instances:
            family["instances"] = instances
            projected_arrays.append(family)
    projected["array_mappings"] = projected_arrays

    projected_profile = copy.deepcopy(source_profile)
    projected_families = []
    for family in projected_profile.get("families", []):
        members = [address for address in family.get("source_members", [])
                   if _canonical_reference(address) in active_formulas]
        if members:
            family["source_members"] = members
            family["member_count"] = len(members)
            projected_families.append(family)
    projected_profile["families"] = projected_families

    active_equation_ids = {item.get("equation_family_id")
                           for item in projected_mappings
                           if isinstance(item.get("equation_family_id"), str)}
    equations = projected.get("equation_families", [])
    if isinstance(equations, list):
        projected["equation_families"] = [item for item in equations
                                           if item.get("family_id") in active_equation_ids]
    elif isinstance(equations, dict):
        projected["equation_families"] = {key: value for key, value in equations.items()
                                           if key in active_equation_ids}

    base_mappings = {item["source_family_id"]: item for item in semantic_plan.get("source_mappings", [])}
    family_to_pass: dict[str, str] = {}
    passes = implementation.get("execution_plan", {}).get("passes")
    if not isinstance(passes, list) or not passes:
        raise CalculationBlocked("implementation has no ordered execution passes")
    module_records = semantic_plan.get("modules", [])
    module_order = [item.get("module_id") for item in module_records if isinstance(item, dict)]
    if [item.get("module_id") for item in sorted(passes, key=lambda record: record.get("order", -1))] != module_order:
        raise CalculationBlocked("implementation passes differ from the reviewed module order")
    variables = projected.get("variables", {})
    variable_ids = set(variables) if isinstance(variables, dict) else {
        item.get("variable_id") for item in variables if isinstance(item, dict)}
    for expected_order, item in enumerate(sorted(passes, key=lambda record: record.get("order", -1))):
        if item.get("order") != expected_order or not isinstance(item.get("pass_id"), str):
            raise CalculationBlocked("implementation pass order or identity is invalid")
        family_ids = item.get("source_family_ids")
        pass_variables = item.get("variable_ids")
        if (not isinstance(family_ids, list) or len(family_ids) != len(set(family_ids))
                or not isinstance(pass_variables, list) or len(pass_variables) != len(set(pass_variables))
                or not set(pass_variables).issubset(variable_ids)):
            raise CalculationBlocked(f"implementation pass has invalid family or variable references: {item.get('pass_id')}")
        for family_id in family_ids:
            mapping = base_mappings.get(family_id)
            if mapping is None or mapping.get("module_id") != item.get("module_id"):
                raise CalculationBlocked(f"implementation assigns an unknown family/module: {family_id}")
            if mapping.get("variable_id") not in pass_variables:
                raise CalculationBlocked(f"implementation pass omits the family output variable: {family_id}")
            if family_id in family_to_pass:
                raise CalculationBlocked(f"source family appears in multiple implementation passes: {family_id}")
            family_to_pass[family_id] = item["pass_id"]
    if set(family_to_pass) != set(base_mappings):
        raise CalculationBlocked("implementation passes do not assign each reviewed source family exactly once")

    active_passes = copy.deepcopy(implementation["execution_plan"])
    for item in active_passes["passes"]:
        item["source_family_ids"] = [family_id for family_id in item["source_family_ids"]
                                      if family_id in active_family_ids]
    projected["execution_plan"] = active_passes
    projected["execution_refinements"] = _apply_family_execution_refinements(
        projected, implementation, trace)
    follower_count = sum(len(instance.get("source_members", []))
                         for family in projected.get("array_mappings", [])
                         for instance in family.get("instances", []))
    instance_count = sum(len(family.get("instances", []))
                         for family in projected.get("array_mappings", []))
    projected["coverage"] = {
        **projected.get("coverage", {}),
        "ordinary_formula_members_mapped": len(active_formulas),
        "array_formula_followers_mapped": follower_count,
        "active_formula_members_mapped": len(active_formulas) + follower_count,
        "source_family_count": len(projected_mappings),
        "semantic_family_count": len(active_equation_ids),
        "array_instance_count": instance_count,
    }
    projected["reference_resolution"] = implementation.get("reference_resolution")
    if not isinstance(projected["reference_resolution"], dict):
        raise CalculationBlocked("implementation has no source reference-resolution map")
    _validate_reference_resolution(projected, projected["reference_resolution"], trace)
    metadata_summary = add_adapter_metadata(projected, catalog, model_scope, trace,
                                            source_path, source_sha256, candidate_trace)

    counts = validate_family_coverage(projected, projected_profile, trace)
    return projected, projected_profile, {**counts, **metadata_summary}


def _validate_reference_resolution(
    plan: dict[str, Any], resolution: dict[str, Any], trace: dict[str, Any],
) -> None:
    variables = plan.get("variables", {})
    if not isinstance(variables, dict):
        raise CalculationBlocked("approved plan has no logical variable map")
    declared_extents: dict[tuple[str, tuple[str, int, int, int, int], str], dict[str, Any]] = {}
    for variable_id, variable in variables.items():
        for extent in variable.get("source_extents", []):
            key = (variable_id, _canonical_reference(extent.get("ref", "")), extent.get("role"))
            declared_extents[key] = extent.get("coordinate_mapping")
    supplied = resolution.get("cell_ranges")
    if not isinstance(supplied, list):
        raise CalculationBlocked("implementation source-cell map is not a list")
    supplied_extents = {}
    for item in supplied:
        key = (item.get("variable_id"), _canonical_reference(item.get("source_ref", "")), item.get("role"))
        supplied_extents[key] = item.get("coordinate_mapping")
    if supplied_extents != declared_extents:
        raise CalculationBlocked("implementation source-cell extents differ from the approved 88-variable plan")

    trace_names = {}
    for item in trace.get("names", []):
        if isinstance(item, dict) and isinstance(item.get("name"), str):
            trace_names[(item.get("scope", "workbook").casefold(), item["name"].casefold())] = item
    names = resolution.get("named_references")
    if not isinstance(names, list):
        raise CalculationBlocked("implementation named-reference map is not a list")
    source_index = build_source_cell_index(plan)
    names_by_id = set()
    for item in names:
        name, scope, kind = item.get("name"), item.get("scope", "workbook"), item.get("kind")
        if not isinstance(name, str) or not isinstance(scope, str) or (scope.casefold(), name.casefold()) not in trace_names:
            raise CalculationBlocked(f"implementation name is absent from the current trace: {name!r}")
        key = (scope.casefold(), name.casefold())
        if key in names_by_id:
            raise CalculationBlocked(f"implementation contains a duplicate named reference: {name}")
        names_by_id.add(key)
        trace_name = trace_names[key]
        if _canonical_reference(item.get("source_ref", "")) != _canonical_reference(
                trace_name.get("destination", trace_name.get("definition", ""))):
            raise CalculationBlocked(f"implementation name destination differs from current source: {name}")
        if kind in {"scalar", "variable_coordinate"}:
            variable_id = item.get("variable_id")
            if not isinstance(variable_id, str) or variable_id not in variables:
                raise CalculationBlocked(f"implementation scalar name has no approved variable: {name}")
            indices = item.get("indices", []) if kind == "variable_coordinate" else []
            shape = variables[variable_id].get("shape", [])
            if (not isinstance(indices, list) or len(indices) != len(shape)
                    or any(not isinstance(index, int) or index < 0 or index >= shape[dimension]
                           for dimension, index in enumerate(indices))):
                if shape or indices:
                    raise CalculationBlocked(f"implementation name coordinate is outside its variable: {name}")
            source = _canonical_reference(item.get("source_ref", ""))
            if source is None or source[1] != source[3] or source[2] != source[4]:
                raise CalculationBlocked(f"implementation scalar name is not a single source cell: {name}")
            matches = [cell for cell in source_index.get((source[0], source[1], source[2]), [])
                       if cell.variable_id == variable_id and cell.indices == tuple(indices)]
            if len(matches) != 1:
                raise CalculationBlocked(f"implementation scalar name does not match its logical source coordinate: {name}")
        elif kind == "range":
            table_binding = item.get("table_binding")
            if not isinstance(table_binding, dict):
                raise CalculationBlocked(f"implementation table name has no binding: {name}")
        else:
            raise CalculationBlocked(f"implementation uses an unsupported name binding kind: {kind!r}")
    if names_by_id != set(trace_names):
        raise CalculationBlocked("implementation name map does not cover all current source names")


def _preflight_context(workflow_dir: Path, implementation_path: Path) -> dict[str, Any]:
    root, workflow = load_workflow(workflow_dir)
    _manifest, stage3_revision = require_stage_approved(root, 3)
    context, capture, capture_path, capture_file_sha = load_current_external_bundle(root)
    stage3_json = root / stage3_revision["artifact"]["json"]
    stage3 = read_json(stage3_json)
    source_sha256 = context["source_sha256"]
    from excel_to_act.steps.step4.generator import _latest_discovery_trace

    trace_path = _latest_discovery_trace(root, stage3_revision, source_sha256)
    trace = read_json(trace_path)
    trace["active_trace_sha256"] = hash_file(trace_path)
    semantic_plan = context["semantic_plan"]
    plan_path = context["semantic_plan_path"]
    check_path = context["semantic_check_path"]
    profile_path = context["evidence_paths"]["candidate_profile"][0]
    profile = context["candidate_profile"]
    scope_path = context["evidence_paths"]["model_scope"][0]
    model_scope = read_json(scope_path)
    input_path = implementation_path.expanduser().resolve()
    if not input_path.is_file():
        raise CalculationBlocked("Step 4 implementation input file is missing")
    implementation = read_json(input_path)
    expected_binding = {
        "active_trace_sha256": hash_file(trace_path),
        "candidate_trace_sha256": context["evidence_paths"]["candidate_trace"][1],
        "external_capture_sha256": capture_file_sha,
        "semantic_check_sha256": context["semantic_check_artifact"]["sha256"],
        "semantic_plan_sha256": context["semantic_plan_artifact"]["sha256"],
        "source_profile_sha256": context["evidence_paths"]["candidate_profile"][1],
        "source_sha256": source_sha256,
        "stage3_artifact_sha256": stage3_revision["artifact"]["json_sha256"],
        "stage3_revision": stage3_revision["revision"],
    }
    if (implementation.get("schema_version") != "step4.active_implementation.v1"
            or implementation.get("binding") != expected_binding):
        raise CalculationBlocked("implementation input is stale or not bound to the current source evidence")
    return {"root": root, "workflow": workflow, "stage3_revision": stage3_revision,
            "stage3": stage3, "stage3_json": stage3_json,
            "source_path": context["source_path"], "source_sha256": source_sha256,
            "capture": capture, "capture_path": capture_path, "capture_file_sha256": capture_file_sha,
            "trace_path": trace_path, "trace": trace, "semantic_plan": semantic_plan,
            "semantic_plan_path": plan_path, "semantic_check_path": check_path,
            "source_profile": profile, "source_profile_path": profile_path,
            "scope_path": scope_path, "model_scope": model_scope,
            "implementation": implementation, "implementation_path": input_path,
            "implementation_sha256": hash_file(input_path), "catalog": context["catalog"],
            "catalog_path": context["catalog_path"], "context": context}


def validate_implementation(workflow_dir: Path, implementation_path: Path) -> dict[str, Any]:
    """Run the full source-bound compiler/scheduler preflight in memory."""
    context = _preflight_context(workflow_dir, implementation_path)
    plan, profile, counts = _active_projection(
        context["semantic_plan"], context["source_profile"], context["trace"],
        context["implementation"], context["catalog"], context["model_scope"],
        context["source_path"], context["source_sha256"],
    )
    index, variables = _source_index(plan)
    raw_addresses = {_item["address"].casefold() for _item in _raw_bindings(plan, index, variables)}
    names = _named_bindings(plan, index, raw_addresses)

    def resolve_name(name: str, _sheet: str, _owner: str) -> str:
        if name.casefold() not in names:
            raise CalculationBlocked(f"current source name is not mapped: {name}")
        return f"runtime.lookup_name(values, {name.casefold()!r})"

    compiled = compile_mapped_families(plan, context["trace"], resolve_name)
    generated_source = render_grouped_driver(plan, compiled)
    compile(generated_source, "pricing.py", "exec")
    if compiled["compiled_member_count"] != counts["ordinary_formula_count"]:
        raise CalculationBlocked("preflight compiler member count differs from the active trace")
    if compiled.get("unsupported_references") or compiled.get("unsupported_names"):
        raise CalculationBlocked("preflight compiler left active source references or names unresolved")
    return {**context, "projected_plan": plan, "projected_profile": profile,
            "coverage": counts, "compiled": compiled, "driver_source": generated_source}


def create_implementation_plan(workflow_dir: Path, implementation_path: Path) -> dict[str, Any]:
    """Write an immutable Step 4 preflight only after all current bindings compile."""
    try:
        validated = validate_implementation(workflow_dir, implementation_path)
        root = validated["root"]
        plan_root = (root / "stage4" / "implementation_plans").resolve()
        revisions = [path for path in plan_root.glob("revision-*") if path.is_dir()] if plan_root.exists() else []
        next_revision = max((int(path.name.removeprefix("revision-")) for path in revisions
                             if path.name.removeprefix("revision-").isdigit()), default=0) + 1
        revision_dir = plan_root / f"revision-{next_revision:04d}"
        if not revision_dir.resolve().is_relative_to(plan_root) or revision_dir.exists():
            raise CalculationBlocked("implementation plan revision path is invalid or already exists")
        revision_dir.mkdir(parents=True)
        context = validated["context"]
        evidence = context["evidence_paths"]
        payload = {
            "schema_version": "step4.implementation_plan.v1", "tool": "step4.plan",
            "status": "pass", "source_sha256": validated["source_sha256"],
            "stage3": {"revision": validated["stage3_revision"]["revision"],
                       "artifact_json_sha256": validated["stage3_revision"]["artifact"]["json_sha256"],
                       "artifact_md_sha256": validated["stage3_revision"]["artifact"]["md_sha256"]},
            "scenario": context["scenario"],
            "catalog": {"revision": context["boundary_reference"]["revision"],
                        "boundary_sha256": context["boundary_reference"]["boundary_sha256"],
                        "artifact_json_sha256": context["catalog_artifact"]["json_sha256"],
                        "artifact_md_sha256": context["catalog_artifact"]["md_sha256"]},
            "implementation_input": {"path": str(validated["implementation_path"]),
                                     "sha256": validated["implementation_sha256"]},
            "semantic_plan": {"path": str(validated["semantic_plan_path"]),
                              "sha256": context["semantic_plan_artifact"]["sha256"]},
            "semantic_check": {"path": str(validated["semantic_check_path"]),
                               "sha256": context["semantic_check_artifact"]["sha256"]},
            "candidate_trace": {"path": str(evidence["candidate_trace"][0]),
                                "sha256": evidence["candidate_trace"][1]},
            "candidate_profile": {"path": str(evidence["candidate_profile"][0]),
                                  "sha256": evidence["candidate_profile"][1]},
            "model_scope": {"path": str(evidence["model_scope"][0]),
                            "sha256": evidence["model_scope"][1]},
            "external_capture": {"path": str(validated["capture_path"]),
                                 "artifact_sha256": validated["capture_file_sha256"],
                                 "capture_sha256": validated["capture"].get("external_capture_sha256"),
                                 "values_sha256": validated["capture"].get("external_values_sha256")},
            "active_trace": {"path": str(validated["trace_path"]),
                             "sha256": hash_file(validated["trace_path"])},
            "reference_resolution": validated["implementation"]["reference_resolution"],
            "execution_plan": validated["implementation"]["execution_plan"],
            "execution_refinements": validated["projected_plan"].get("execution_refinements", []),
            "adapter_metadata_groups": validated["projected_plan"].get("adapter_metadata_groups", {}),
            "adapter_metadata_cells": validated["projected_plan"].get("adapter_metadata_cells", []),
            "metadata_seed_bindings": validated["projected_plan"].get("metadata_seed_bindings", []),
            "coverage": validated["coverage"],
            "compiled_variant_count": validated["compiled"]["function_count"],
            "source_family_count": validated["compiled"]["source_family_count"],
            "business_input_count": sum(item.get("role") in {"raw", "external"}
                                         for item in validated["projected_plan"]["variables"].values()),
            "raw_business_input_count": sum(item.get("role") == "raw"
                                             for item in validated["projected_plan"]["variables"].values()),
            "external_business_input_count": sum(item.get("role") == "external"
                                                  for item in validated["projected_plan"]["variables"].values()),
            "model_variable_count": len(validated["projected_plan"]["variables"]),
            "derived_variable_count": sum(item.get("role") == "derived"
                                           for item in validated["projected_plan"]["variables"].values()),
            "formula_cache_inputs": False, "native_excel_called": False,
        }
        json_path = revision_dir / "implementation_plan.json"
        json_path.write_text(json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
                             encoding="utf-8")
        lines = ["# Step 4 Implementation Preflight", "", "Status: **Pass**", "",
                 f"Source SHA-256: `{payload['source_sha256']}`",
                 f"Approved Stage 3 revision: {payload['stage3']['revision']}",
                 f"Fresh active trace SHA-256: `{payload['active_trace']['sha256']}`", "",
                 f"Active formulas: {payload['coverage']['ordinary_formula_count']}; array followers: {payload['coverage']['array_follower_count']}; compiled variants: {payload['compiled_variant_count']}.",
                 f"Business inputs: {payload['business_input_count']} ({payload['raw_business_input_count']} raw and {payload['external_business_input_count']} formula-derived external); model variables: {payload['model_variable_count']} ({payload['derived_variable_count']} derived).",
                 f"Raw source coordinates: {payload['coverage']['raw_source_coordinate_count']}; catalog metadata coordinates: {payload['coverage']['source_metadata_coordinate_count']} across {len(payload['adapter_metadata_groups'])} groups; metadata-initialized derived variables: {payload['coverage']['metadata_seed_binding_count']}.",
                 f"Trace-proved same-index schedule refinements: {len(payload['execution_refinements'])}.",
                 "The reviewed 88-variable plan is projected to the fresh active trace. Source names, table bindings, and execution passes compiled in memory.",
                 "This preflight calls no native Excel and reads no formula caches. It does not append or approve a Stage 4 generation report.", ""]
        md_path = revision_dir / "implementation_plan.md"
        md_path.write_text("\n".join(lines), encoding="utf-8")
        coverage_summary = {
            key: payload["coverage"].get(key)
            for key in ("ordinary_formula_count", "array_follower_count",
                        "active_formula_member_count", "source_family_count",
                        "semantic_family_count", "array_instance_count",
                        "source_coordinate_count", "raw_source_coordinate_count",
                        "source_metadata_coordinate_count", "metadata_seed_binding_count")
        }
        return {"schema_version": payload["schema_version"], "tool": payload["tool"], "status": "pass",
                "workflow": str(root), "revision": f"revision-{next_revision:04d}",
                "artifact": str(json_path), "artifact_sha256": hash_file(json_path),
                "markdown": str(md_path), "compiled_variant_count": payload["compiled_variant_count"],
                "source_family_count": payload["source_family_count"],
                "business_input_count": payload["business_input_count"],
                "raw_business_input_count": payload["raw_business_input_count"],
                "external_business_input_count": payload["external_business_input_count"],
                "model_variable_count": payload["model_variable_count"],
                "derived_variable_count": payload["derived_variable_count"],
                "execution_refinement_count": len(payload["execution_refinements"]),
                "raw_source_coordinate_count": payload["coverage"]["raw_source_coordinate_count"],
                "source_metadata_coordinate_count": payload["coverage"]["source_metadata_coordinate_count"],
                "metadata_seed_binding_count": payload["coverage"]["metadata_seed_binding_count"],
                "coverage": coverage_summary, "stage_advanced": False}
    except Exception as exc:
        return {"schema_version": "step4.implementation_plan.v1", "tool": "step4.plan",
                "status": "blocked", "reason": str(exc), "stage_advanced": False}


def load_current_implementation_plan(
    workflow_dir: Path,
) -> tuple[dict[str, Any], Path, str, dict[str, Any]]:
    """Reload and revalidate the latest preflight against all current upstream bytes."""
    root = Path(workflow_dir).expanduser().resolve()
    plan_root = root / "stage4" / "implementation_plans"
    paths = sorted(plan_root.glob("revision-*/implementation_plan.json"), reverse=True) \
        if plan_root.exists() else []
    if not paths:
        raise CalculationBlocked("no current Step 4 implementation preflight exists; run `step4 plan` first")
    path = paths[0]
    record = read_json(path)
    input_file = record.get("implementation_input", {})
    input_path = Path(input_file.get("path", "")).resolve()
    if (record.get("schema_version") != "step4.implementation_plan.v1" or record.get("status") != "pass"
            or not input_path.is_file() or hash_file(input_path) != input_file.get("sha256")):
        raise CalculationBlocked("latest Step 4 implementation preflight or its input is stale")
    validated = validate_implementation(root, input_path)
    context = validated["context"]
    evidence = context["evidence_paths"]
    current = {
        "source_sha256": validated["source_sha256"],
        "scenario": context["scenario"],
        "stage3": {"revision": validated["stage3_revision"]["revision"],
                   "artifact_json_sha256": validated["stage3_revision"]["artifact"]["json_sha256"],
                   "artifact_md_sha256": validated["stage3_revision"]["artifact"]["md_sha256"]},
        "catalog": {"revision": context["boundary_reference"]["revision"],
                    "boundary_sha256": context["boundary_reference"]["boundary_sha256"],
                    "artifact_json_sha256": context["catalog_artifact"]["json_sha256"],
                    "artifact_md_sha256": context["catalog_artifact"]["md_sha256"]},
        "semantic_plan": {"path": str(validated["semantic_plan_path"]),
                           "sha256": context["semantic_plan_artifact"]["sha256"]},
        "semantic_check": {"path": str(validated["semantic_check_path"]),
                            "sha256": context["semantic_check_artifact"]["sha256"]},
        "candidate_trace": {"path": str(evidence["candidate_trace"][0]),
                             "sha256": evidence["candidate_trace"][1]},
        "candidate_profile": {"path": str(evidence["candidate_profile"][0]),
                               "sha256": evidence["candidate_profile"][1]},
        "model_scope": {"path": str(evidence["model_scope"][0]),
                        "sha256": evidence["model_scope"][1]},
        "external_capture": {"path": str(validated["capture_path"]),
                             "artifact_sha256": validated["capture_file_sha256"],
                             "capture_sha256": validated["capture"].get("external_capture_sha256"),
                             "values_sha256": validated["capture"].get("external_values_sha256")},
        "active_trace": {"path": str(validated["trace_path"]),
                         "sha256": hash_file(validated["trace_path"])},
        "execution_plan": validated["implementation"]["execution_plan"],
        "execution_refinements": validated["projected_plan"].get("execution_refinements", []),
        "reference_resolution": validated["implementation"]["reference_resolution"],
        "adapter_metadata_groups": validated["projected_plan"].get("adapter_metadata_groups", {}),
        "adapter_metadata_cells": validated["projected_plan"].get("adapter_metadata_cells", []),
        "metadata_seed_bindings": validated["projected_plan"].get("metadata_seed_bindings", []),
        "coverage": validated["coverage"],
        "compiled_variant_count": validated["compiled"]["function_count"],
        "source_family_count": validated["compiled"]["source_family_count"],
    }
    if _implementation_plan_binding_mismatches(record, current):
        raise CalculationBlocked("latest implementation plan no longer matches current Stage 3/4 inputs")
    return record, path, hash_file(path), validated


def _implementation_plan_binding_mismatches(record: dict[str, Any], current: dict[str, Any]) -> list[str]:
    """Compare the serialized plan inputs with their freshly validated bindings."""
    return [key for key, value in current.items() if record.get(key) != value]
