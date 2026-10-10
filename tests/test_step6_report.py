from __future__ import annotations

import json
import re
from pathlib import Path

from excel_to_act.steps.conversion_workflow import (
    _write_manifest,
    append_stage_artifact,
    create_workflow,
    hash_file,
    load_workflow,
    read_json,
    record_decision,
)
from excel_to_act.steps.step6.workflow import (
    _load_input_catalog,
    _report_markdown,
    _result_row,
    create_conversion_report,
)


def _approve(root: Path, stage: int) -> None:
    for reviewer in ("agent", "human"):
        record_decision(root, stage, reviewer, "approve", f"Synthetic fixture review for stage {stage}.")


def _append(root: Path, stage: int, payload: dict, prior: dict | None = None) -> dict:
    return append_stage_artifact(root, stage, f"stage{stage}.json", payload,
                                 f"Synthetic Stage {stage} evidence.",
                                 input_stages={stage - 1: prior["artifact"]} if prior else None)


def _catalog(root: Path, source_sha: str) -> tuple[dict, Path, Path]:
    catalog = {
        "schema_version": "step3.input_boundary.v1",
        "tool": "step3.input_boundary",
        "phase": "input_boundary",
        "status": "ready_for_review",
        "boundary_sha256": "fixture-boundary",
        "source": {"workbook_sha256": source_sha},
        "counts": {
            "logical_input_object_count": 5,
            "source_raw_variable_count": 4,
            "formula_derived_external_variable_count": 1,
            "scalar_field_count": 2,
            "vector_object_count": 2,
            "table_count": 1,
            "source_record_group_count": 1,
        },
        "groups": [
            {"group_id": "saved_case_inputs", "name": "Saved case inputs", "classification": "business_input"},
            {"group_id": "projection_rates", "name": "Projection rate series", "classification": "assumption_curves"},
            {"group_id": "category_table", "name": "Category matrix", "classification": "configuration_table"},
            {"group_id": "external_rates", "name": "External rate boundary", "classification": "formula_derived_external"},
            {"group_id": "empty_metadata", "name": "Empty metadata group", "classification": "metadata"},
        ],
        "variables": [
            {"logical_name": "StartAge", "group_id": "saved_case_inputs", "kind": "scalar",
             "role": "source_raw", "shape": [], "axes": [], "variable_id": "inputs.b2",
             "source_extents": [{"sheet": "Inputs", "range": "B2"}]},
            {"logical_name": "RunMode", "group_id": "saved_case_inputs", "kind": "scalar",
             "role": "source_raw", "shape": [], "axes": [], "variable_id": "inputs.c4",
             "source_extents": [{"sheet": "Inputs", "range": "C4"}]},
            {"logical_name": "ProjectionRate", "group_id": "projection_rates", "kind": "vector",
             "role": "source_raw", "shape": [4], "axes": [{"name": "policy_year", "role": "key"}],
             "variable_id": "inputs.d2_d5", "source_extents": [{"sheet": "Inputs", "range": "D2:D5"}]},
            {"logical_name": "CategoryWeights", "group_id": "category_table", "kind": "table",
             "role": "source_raw", "shape": [2, 3],
             "axes": [{"name": "record", "role": "record"}, {"name": "category", "role": "category"}],
             "variable_id": "inputs.d2_f3", "source_extents": [{"sheet": "Inputs", "range": "D2:F3"}]},
            {"logical_name": "ExternalCurve", "group_id": "external_rates", "kind": "vector",
             "role": "formula_derived_external", "shape": [4],
             "axes": [{"name": "policy_year", "role": "key"}], "variable_id": "external.curve",
             "source_extents": [{"sheet": "Rates", "range": "A2:A5"}]},
        ],
        "source_records": [{
            "record_id": "fixture.table.boundary",
            "classification": "source_assumption",
            "group_id": "category_table",
            "source_extents": [{"sheet": "Inputs", "range": "D2:F3"}],
            "provenance": {"description": "Fixture category table records retained from the source."},
        }],
    }
    folder = root / "catalog"
    folder.mkdir()
    json_path = folder / "input_boundary.json"
    markdown_path = folder / "input_boundary.md"
    json_path.write_text(json.dumps(catalog, sort_keys=True), encoding="utf-8")
    markdown_path.write_text("Synthetic input catalog.", encoding="utf-8")
    reference = {
        "schema_version": "step3.input_boundary.reference.v1",
        "revision": 1,
        "boundary_sha256": catalog["boundary_sha256"],
        "artifact": {
            "json": str(json_path.relative_to(root)),
            "json_sha256": hash_file(json_path),
            "md": str(markdown_path.relative_to(root)),
            "md_sha256": hash_file(markdown_path),
        },
    }
    return reference, json_path, markdown_path


def _build_workflow(root: Path) -> tuple[Path, Path, Path]:
    root.mkdir()
    source_sha = "fixture-workbook"
    source_path = root / "source.xlsx"
    source_path.write_bytes(b"synthetic source workbook")
    create_workflow(root, {
        "workbook_sha256": source_sha,
        "workbook_path": str(source_path),
        "run_id": "fixture-run",
        "source_id": "fixture",
    })
    _, manifest = load_workflow(root)
    manifest["mode"] = "synthetic_fixture"
    _write_manifest(root, manifest)

    step1 = _append(root, 1, {"status": "ready_for_review", "quality": {
        "status": "complete", "ready_for_next_step": True,
        "metrics": {"logical_objects_accounted": 5, "logical_objects_total": 5,
                    "package_parts_preserved": 8, "package_parts_total": 8,
                    "parsed_package_parts": 5, "parsed_package_parts_total": 8,
                    "opaque_parts": 3, "opaque_rate": 0.375}}})
    _approve(root, 1)
    step2 = _append(root, 2, {"status": "ready_for_review"}, step1)
    _approve(root, 2)

    catalog_reference, catalog_json, catalog_markdown = _catalog(root, source_sha)
    design = {
        "targets": [
            {"target_id": "T_RETURN", "selector": "NetReturn", "result_order": 0, "shape": [],
             "units": "ratio", "description": "Aggregated saved-case net return.",
             "diagnostic_intermediates": ["ReserveRatio"]},
        ],
        "scenarios": [{
            "scenario_id": "base-case",
            "primary_inputs": {
                "RunMode": {"source_cell": "Inputs!C4", "source_literal": "base"},
                "StartAge": {"source_cell": "Inputs!B2", "source_literal": 60},
            },
        }],
        "scope": {
            "boundary": "One saved input case and two requested result values.",
            "included": ["The selected base-case calculation."],
            "excluded": ["Other scenarios", "An iterative feedback solver"],
            "source_behavior_policy": "Retain the accepted source behavior for the selected case.",
        },
        "execution_design": {
            "backend": "Python",
            "ordering": "Inputs -> exposure projection -> return aggregation.",
            "emission": "Reusable functions and named vectors; no function per source cell.",
            "input_policy": "Source literals and separately supplied external values are kept distinct.",
            "preserve_formulas": ["The recurrence uses the prior-period state.",
                                  "Same-year decrement order follows explicit seeds and snapshots.",
                                  "Final result equation: Result = A / (B - C); intermediate outputs are reductions."],
        },
        "shared_modules": [
            {"order": 0, "module_id": "InputAdapter", "purpose": "Load and validate accepted input values."},
            {"order": 1, "module_id": "ExposureProjection", "purpose": "Build ordered policy-year state vectors."},
            {"order": 2, "module_id": "ReturnAggregation", "purpose": "Reduce projected state to requested results."},
        ],
        "field_groups": [
            {"id": "source_inputs", "shape": "Two scalars and a policy-year vector",
             "source_extents": ["Inputs!B2:B5"]},
            {"id": "projection", "shape": "Same-year state recurrence with prior-period input",
             "source_extents": ["Projection!A2:A5"]},
        ],
        "coverage": {
            "known_path_ids": ["R_BASE", "R_ADJUST", "R_FEEDBACK"],
            "documented_path_ids": ["R_BASE", "R_ADJUST", "R_FEEDBACK"],
            "resolved_path_ids": ["R_BASE", "R_ADJUST", "R_FEEDBACK"],
            "planned_path_ids": ["R_BASE", "R_ADJUST", "R_FEEDBACK"],
            "implemented_path_ids": [],
            "runtime_verified_path_ids": [],
            "all_configuration_coverage": None,
        },
        "options": [
            {"option_id": "O_BASE", "name": "Base case complete calculation", "selected": True,
             "scope": "One base case", "status": "Selected", "numerical_paths": ["R_BASE", "R_ADJUST"],
             "condition_only_paths": ["R_FEEDBACK"], "exclusions": ["An iterative feedback solver"]},
            {"option_id": "O_EXTENDED", "name": "Expanded scenario set", "selected": False,
             "scope": "Additional cases", "status": "Not selected", "all_configuration_coverage": None,
             "exclusions": ["Additional input selections"]},
        ],
        "validation_plan": [],
        "open_questions": [{
            "question_id": "Q_RUN_MODE",
            "question": "Is the base run mode the only approved business interpretation?",
            "disposition": "Keep the current saved value; request business review for any alternate meaning.",
        }],
    }
    stage3 = _append(root, 3, {
        "status": "ready_for_review",
        "source": {"workbook_sha256": source_sha},
        "design": design,
        "input_boundary_reference": catalog_reference,
    }, step2)
    _approve(root, 3)

    bundle_dir = root / "bundle"
    bundle_dir.mkdir()
    bundle_files = {}
    for filename, content in {
        "model.py": "def run(): return 0.125\n",
        "model_manifest.json": '{"schema_version": "fixture"}',
        "source_map.json": '{"sources": []}',
    }.items():
        path = bundle_dir / filename
        path.write_text(content, encoding="utf-8")
        bundle_files[filename] = hash_file(path)
    implementation_evidence = [
        {"path_id": "R_BASE", "status": "generated", "address": "Results!B2", "source_formula": "=SUM(A1:A4)"},
        {"path_id": "R_ADJUST", "status": "generated", "address": "Results!B3", "source_formula": "=B2*C2"},
        {"path_id": "R_UNKNOWN", "status": "generated", "address": "Results!B4", "source_formula": "=1"},
    ]
    stage4 = _append(root, 4, {
        "status": "ready_for_review",
        "bundle": {"path": str(bundle_dir), "files": bundle_files,
                   "manifest_path": str(bundle_dir / "model_manifest.json")},
        "formula_count": 6,
        "compiled_function_count": 6,
        "source_family_count": 4,
        "semantic_equation_group_count": 3,
        "array_member_count": 9,
        "array_instance_count": 3,
        "business_input_count": 5,
        "raw_business_input_count": 4,
        "external_business_input_count": 1,
        "model_variable_count": 6,
        "derived_variable_count": 2,
        "raw_input_count": 12,
        "raw_source_coordinate_count": 12,
        "source_metadata_coordinate_count": 2,
        "external_input_coordinate_count": 4,
        "external_capture": {"ranges": ["Inputs!H2:H5", "Inputs!J2:J5"],
                             "formula_cut_member_count": 4,
                             "age_axis": {"source_range": "Inputs!A2:A5"}},
        "target_names": ["NetReturn", "ReserveRatio"],
        "runtime_dependencies": ["Python standard library"],
        "path_coverage": {
            "known_path_ids": ["R_BASE", "R_ADJUST", "R_FEEDBACK"],
            "numerical_path_ids": ["R_BASE", "R_ADJUST"],
            "condition_only_path_ids": ["R_FEEDBACK"],
            "implemented_path_ids": ["R_BASE", "R_ADJUST", "R_UNKNOWN"],
            "implementation_evidence": implementation_evidence,
        },
    }, stage3)
    _approve(root, 4)
    (root / "excel_oracle.json").write_text("{}", encoding="utf-8")
    (root / "standalone_validation.json").write_text("{}", encoding="utf-8")
    _append(root, 5, {
        "status": "ready_for_review",
        "compared": {
            "named_targets": 3,
            "required_targets": 3,
            "primary_inputs": 2,
            "active_formula_members": 6,
            "required_active_formula_members": 6,
            "active_formula_identity_matches": 6,
            "required_active_formula_identities": 6,
            "active_formula_max_abs_difference": 2.22e-16,
            "external_input_coordinates": 4,
            "required_external_input_coordinates": 4,
            "external_formula_cut_members": 4,
            "required_external_formula_cut_members": 4,
            "external_age_axis_keys": 4,
            "required_external_age_axis_keys": 4,
        },
        "target_results": {
            "NetReturn": {"model": 0.125, "excel": 0.125, "matched": True},
            "ReserveRatio": {"model": 1.25, "excel": 1.25, "matched": True},
            "SupplementalMeasure": {"model": 2.0, "excel": 2.0, "matched": True},
        },
        "tolerances": {"absolute": 1e-12, "relative": 1e-12},
        "mismatch_count": 0,
        "oracle": {
            "engine": "Microsoft Excel", "macros_executed": False, "input_overrides": [],
            "engine_settings": {"calculation": -4105, "version": "16.0", "iteration": True,
                                "max_change": 0.001, "max_iterations": 5},
            "path": str(root / "excel_oracle.json"),
        },
        "standalone_validation": {"path": str(root / "standalone_validation.json")},
        "standalone_project_imports": False,
        "formula_cache_inputs": False,
        "path_coverage": {
            "implemented_path_ids": ["R_BASE", "R_ADJUST", "R_UNKNOWN"],
            "runtime_verified_path_ids": ["R_BASE", "R_UNKNOWN"],
            "runtime_evidence": [
                {"path_id": "R_BASE", "status": "matched", "address": "Results!B2",
                 "source_formula": "=SUM(A1:A4)", "comparison": {"kind": "native_named_target"}},
                {"path_id": "R_UNKNOWN", "status": "matched", "address": "Results!B4",
                 "source_formula": "=1", "comparison": {"kind": "native_formula_range"}},
            ],
            "condition_evidence": [{
                "path_id": "R_FEEDBACK", "status": "inactive_condition_checked_no_solver",
                "evidence": "The condition is inactive in this case; no solver was generated.",
            }],
        },
    }, stage4)
    _approve(root, 5)
    return root, catalog_json, catalog_markdown


def test_step6_human_report_preserves_evidence_and_binds_catalog_files(tmp_path: Path) -> None:
    root, catalog_json, catalog_markdown = _build_workflow(tmp_path / "workflow")
    result = create_conversion_report(root)
    assert result["status"] == "pass", result
    report = read_json(root / result["artifact"]["json"])
    summary = report["design_summary"]
    assert summary["execution_design"]["ordering"] == "Inputs -> exposure projection -> return aggregation."
    assert [module["module_id"] for module in summary["shared_modules"]] == [
        "InputAdapter", "ExposureProjection", "ReturnAggregation",
    ]
    assert summary["scenarios"][0]["primary_inputs"]["StartAge"]["source_literal"] == 60
    assert summary["targets"][0]["result_order"] == 0
    assert summary["targets"][0]["diagnostic_intermediates"] == ["ReserveRatio"]
    assert summary["input_catalog"]["available"] is True
    assert len(summary["input_catalog"]["content"]["source_records"]) == 1
    assert summary["input_catalog"]["markdown_artifact_path"] == str(catalog_markdown.resolve())
    assert summary["input_catalog"]["markdown_artifact_sha256"] == hash_file(catalog_markdown)
    assert summary["open_questions"] == [{
        "question_id": "Q_RUN_MODE",
        "question": "Is the base run mode the only approved business interpretation?",
        "disposition": "Keep the current saved value; request business review for any alternate meaning.",
    }]
    coverage = summary["actual_coverage"]
    assert coverage["implemented_path_ids"] == ["R_ADJUST", "R_BASE"]
    assert coverage["runtime_verified_path_ids"] == ["R_BASE"]
    assert coverage["condition_assessed_inactive_ids"] == ["R_FEEDBACK"]
    assert report["target_results"]["NetReturn"]["excel"] == 0.125
    assert report["validation_context"]["oracle"]["engine"] == "Microsoft Excel"
    assert report["validation_context"]["standalone_project_imports"] is False
    _, manifest = load_workflow(root)
    revision = manifest["stages"]["6"]["revisions"][-1]
    assert revision["input_files"]["step3_input_boundary_json"]["path"] == str(catalog_json.resolve())
    assert revision["input_files"]["step3_input_boundary_json"]["sha256"] == hash_file(catalog_json)
    assert revision["input_files"]["step3_input_boundary_markdown"]["path"] == str(catalog_markdown.resolve())
    assert revision["input_files"]["step3_input_boundary_markdown"]["sha256"] == hash_file(catalog_markdown)

    markdown = (root / result["artifact"]["md"]).read_text(encoding="utf-8")
    required_sections = [
        "## Executive summary",
        "## Purpose/intended use/scope",
        "## Source model and inputs/assumptions",
        "## Calculation design and time/module order",
        "## Conversion approach",
        "## Validation and reconciliation results",
        "## Differences/limitations/open decisions",
        "## Handover/run instructions",
        "## Acceptance",
    ]
    assert all(section in markdown for section in required_sections)
    assert markdown.index("## Executive summary") < markdown.index("## Source model")
    assert "| Classification | Measure |" in markdown
    assert markdown.index("| Requested result | NetReturn |") < markdown.index("| Diagnostic intermediate | ReserveRatio |")
    assert markdown.index("| Diagnostic intermediate | ReserveRatio |") < markdown.index("| Additional Step 5 check | SupplementalMeasure |")
    assert "| Order | Requested result |" not in markdown
    assert "flowchart LR" in markdown
    assert "InputAdapter" in markdown and "ReturnAggregation" in markdown
    assert "NetReturn" in markdown and "ReserveRatio" in markdown
    assert "Inputs!B2" in markdown and "60" in markdown
    assert "policy_year" in markdown and "2 × 3" in markdown
    assert "2 scalars, 2 vectors, 1 table" in markdown
    assert "Distinct catalog variables by source tab" in markdown
    catalog_link = re.search(r"Input-boundary catalog: \[catalog Markdown\]\(([^)]+)\)", markdown)
    assert catalog_link and catalog_link.group(1).endswith("input_boundary.md")
    assert "| Inputs | 4 | 0 | 4 |" in markdown
    assert "| Rates | 0 | 1 | 1 |" in markdown
    assert "Empty metadata group" not in markdown
    assert "Same-year state recurrence" in markdown
    assert "prior-period state" in markdown
    assert "Result equation and reductions" in markdown
    assert "Result = A / (B - C)" in markdown
    assert "Step 4 generated 2 numerical routes" in markdown
    assert "Step 5 verified 1 of them" in markdown
    assert "R_BASE" in markdown and "R_ADJUST" in markdown
    assert "R_UNKNOWN" not in markdown
    assert "Condition-only route checks assessed inactive: 1 of 1 (R_FEEDBACK). No feedback solver is included." in markdown
    assert "Comparison tolerances: absolute 1e-12; relative 1e-12. Mismatches: 0." in markdown
    assert "Validation method" in markdown and "Microsoft Excel" in markdown
    assert "Standalone Python run" in markdown and "Native workbook comparison" in markdown
    assert "Macros executed for the oracle run: No." in markdown
    assert "Oracle input overrides: none recorded." in markdown
    assert "Formula-cache inputs: No." in markdown
    assert "Recorded Excel run conditions: Excel version 16.0; iteration enabled; maximum 5 iterations; maximum change 0.001." in markdown
    main_body = markdown.split("## Appendices", maxsplit=1)[0]
    assert "-4105" not in main_body
    assert "| Primary source inputs | 2 | Not recorded |" in markdown
    assert "| External formula-cut members | 4 | 4 |" in markdown
    assert "Five external CI vectors remain absent" not in markdown
    assert "missing external-input capture prerequisite was completed" in markdown
    assert "Source extraction was partial: 5 of 8 package parts were parsed and 3 are recorded as opaque" in markdown
    assert "Unresolved business interpretation" in markdown
    assert "Is the base run mode the only approved business interpretation?" in markdown
    assert "Question ID" in markdown and "Original disposition" in markdown
    assert "Keep the current saved value" in markdown
    assert "fixture.table.boundary" in markdown
    assert "[Runnable model]" in markdown or "[model.py]" in markdown
    assert "Current workflow acceptance ledger" in markdown and "workflow.json" in markdown
    assert not re.search(r"\b[a-fA-F0-9]{64}\b", main_body)
    assert '"target_id":' not in main_body
    assert "[Society of Actuaries" in markdown
    assert "SHA-256" in markdown.split("## Appendices", maxsplit=1)[1]
    assert "Step 5 method artifacts" in markdown and "Standalone Python validation" in markdown
    appendix = markdown.split("## Appendices", maxsplit=1)[1]
    assert "| calculation | -4105 |" in appendix
    accepted_design_link = re.search(r"\[Accepted design\]\(([^)]+)\)", markdown)
    assert accepted_design_link and accepted_design_link.group(1).endswith("stage3.md")
    assert "\n\n- The selected base-case calculation." in markdown
    assert "| External key-axis entries | 4 | 4 |\n\nMaximum absolute difference" in markdown



def test_step6_reader_preserves_non_scalar_positions_metadata_and_scalar_style(tmp_path: Path) -> None:
    import copy

    root, _catalog_json, _catalog_markdown = _build_workflow(tmp_path / "workflow")
    result = create_conversion_report(root)
    assert result["status"] == "pass", result
    report = read_json(root / result["artifact"]["json"])
    target = report["design_summary"]["targets"][0]
    target.update({"selector": "Quarter Matrix", "shape": [2, 2], "result_kind": "table",
                   "units": "USD", "axes": [{"name": "policy year", "role": "time", "keys": [2025, 2026]}],
                   "diagnostic_intermediates": ["VectorCheck"]})
    bool_target = {"target_id": "T_ACTIVE", "selector": "active_flag", "result_order": 1,
                   "shape": [], "result_kind": "scalar", "units": "boolean"}
    report["design_summary"]["targets"].append(bool_target)
    report["design_summary"]["scenarios"] = [{
        "scenario_id": "literal-case",
        "primary_inputs": {
            "BooleanFalse": {"source_cell": "Inputs!A1", "source_literal": False},
            "TextNo": {"source_cell": "Inputs!A2", "source_literal": "No"},
        },
    }]
    report["target_results"] = {
        "VectorCheck": {"model": [0, False], "excel": [0, False], "matched": True},
        "Quarter Matrix": {"model": [[11, 12], [21, 22]],
                           "excel": [[11, 12], [21, 22]], "matched": True},
        "active_flag": {"model": False, "excel": "No", "matched": True},
        "SupplementalMeasure": {"model": 2.0, "excel": 2.0, "matched": True},
    }
    original = copy.deepcopy(report)
    markdown = _report_markdown(report)
    primary = markdown.split("## Purpose/intended use/scope", 1)[0]

    requested_line = next(line for line in primary.splitlines() if "| Requested result | Quarter Matrix |" in line)
    diagnostic_line = next(line for line in primary.splitlines() if "| Diagnostic intermediate | VectorCheck |" in line)
    assert "shape [2, 2]; kind table; units USD; axes name: policy year; role: time; keys: [2025, 2026]" in requested_line
    assert "position [0, 0] = 11" in requested_line and "position [0, 1] = 12" in requested_line
    assert "position [1, 0] = 21" in requested_line and "position [1, 1] = 22" in requested_line
    assert primary.index("| Requested result | Quarter Matrix |") < primary.index("| Diagnostic intermediate | VectorCheck |")
    assert "shape Not recorded" in diagnostic_line and "position [0] = 0" in diagnostic_line
    assert "position [1] = False" in diagnostic_line
    bool_line = next(line for line in primary.splitlines() if "| Requested result | active_flag |" in line)
    assert bool_line == "| Requested result | active_flag | False | No | Not recorded | Matched |"
    assert "| literal-case | BooleanFalse | Inputs!A1 | False |" in markdown
    assert "| literal-case | TextNo | Inputs!A2 | No |" in markdown
    assert _result_row("Requested result", "GP", {"model": 0.125, "excel": 0.125, "matched": True}) == [
        "Requested result", "GP", "0.125", "0.125", "0.0", "Matched",
    ]
    assert report == original

def test_step6_legacy_report_marks_catalog_facts_unavailable_without_guessing() -> None:
    report = {
        "source": {}, "stages": {
            "3": {"markdown_path": "design.md", "json_path": "design.json"},
            "4": {"json_path": "generation.json"},
            "5": {"json_path": "validation.json"},
        }, "review_policy": {}, "design_summary": {},
        "generation_summary": {}, "validation_summary": {},
    }
    markdown = _report_markdown(report)
    assert "Input catalog facts are unavailable" in markdown
    assert "Counts, shapes, axes, and source-record detail are therefore not inferred." in markdown
    assert "logical input objects" not in markdown
    assert "Step 4 generated not recorded numerical routes" in markdown
    assert "do not establish a complete reconciliation" in markdown
    assert "[Accepted design](design.md)" in markdown
    assert "[Generation report](generation.json)" in markdown
    assert "[Validation report](validation.json)" in markdown


def test_step6_rejects_catalog_facts_when_a_recorded_digest_is_stale(tmp_path: Path) -> None:
    root = tmp_path / "workflow"
    root.mkdir()
    reference, catalog_json, _ = _catalog(root, "source-sha")
    stage3 = {
        "source": {"workbook_sha256": "source-sha"},
        "input_boundary_reference": reference,
    }
    catalog_json.write_text(catalog_json.read_text(encoding="utf-8") + " ", encoding="utf-8")
    summary, input_paths = _load_input_catalog(root, stage3)
    assert summary["available"] is False
    assert "digest does not match" in summary["reason"]
    assert input_paths == {}


def test_step6_requires_nonempty_accepted_source_digest_for_catalog_facts(tmp_path: Path) -> None:
    root = tmp_path / "workflow"
    root.mkdir()
    reference, _catalog_json, _catalog_markdown = _catalog(root, "source-sha")
    for source in ({}, {"workbook_sha256": ""}, {"workbook_sha256": "   "}):
        summary, input_paths = _load_input_catalog(root, {
            "source": source,
            "input_boundary_reference": reference,
        })
        assert summary["available"] is False
        assert "no nonempty source workbook digest" in summary["reason"]
        assert input_paths == {}


def test_step6_catalog_link_falls_back_to_json_when_markdown_is_not_referenced(tmp_path: Path) -> None:
    root = tmp_path / "workflow"
    root.mkdir()
    reference, _catalog_json, _catalog_markdown = _catalog(root, "source-sha")
    reference["artifact"].pop("md")
    reference["artifact"].pop("md_sha256")
    summary, input_paths = _load_input_catalog(root, {
        "source": {"workbook_sha256": "source-sha"},
        "input_boundary_reference": reference,
    })
    assert summary["available"] is True
    assert "markdown_artifact_path" not in summary
    assert set(input_paths) == {"json"}
    markdown = _report_markdown({
        "source": {}, "stages": {}, "review_policy": {},
        "design_summary": {"input_catalog": summary},
        "generation_summary": {}, "validation_summary": {},
    }, workflow_root=root, report_directory=root / "stage6" / "revision-0001")
    catalog_link = re.search(r"Input-boundary catalog: \[catalog JSON\]\(([^)]+)\)", markdown)
    assert catalog_link and catalog_link.group(1).endswith("input_boundary.json")


def test_step6_marks_delegated_typesafe_as_a_separate_reviewer() -> None:
    report = {
        "source": {}, "stages": {},
        "review_policy": {
            "stage6_required_reviewer_pair": ["agent", "typesafe"],
            "delegation": {"stages": [3, 4, 5, 6],
                           "requested_model": "jev-latest",
                           "authorization_sha256": "auth-sha"},
        },
        "design_summary": {}, "generation_summary": {}, "validation_summary": {},
    }
    markdown = _report_markdown(report)
    assert "Step 6 requires Agent and delegated TypeSafe approval" in markdown
    assert "authorization SHA-256 auth-sha" in markdown


def test_step6_legacy_generation_summary_remains_readable() -> None:
    report = {
        "source": {}, "stages": {}, "review_policy": {}, "design_summary": {},
        "generation_summary": {"formula_count": 30, "unique_equation_family_count": 87,
                               "array_instance_count": 8, "array_member_count": 14,
                               "runtime_dependencies": []},
        "validation_summary": {},
    }
    markdown = _report_markdown(report)
    assert "87 unique equation-family functions" in markdown
    assert "Not recorded source families" in markdown
    assert "Maximum absolute difference" not in markdown
