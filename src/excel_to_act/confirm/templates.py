"""Human confirmation template generation."""

from __future__ import annotations

from excel_to_act.schemas import ConfirmationQuestion, ConfirmationTemplate, ModuleClassification


class ConfirmationTemplateBuilder:
    name = "confirmation_template_builder"

    def build(self, classification: ModuleClassification) -> ConfirmationTemplate:
        questions: list[ConfirmationQuestion] = []
        for item in classification.items:
            if item.confidence < 0.7:
                questions.append(
                    ConfirmationQuestion(
                        id=f"confirm:{item.id}",
                        prompt=f"Confirm category for {item.id}: {item.category}",
                        question_type="category_override",
                        source_location=item.source_location,
                        options=["input", "data_table", "formula_block", "lookup_block", "output", "presentation", "external_dependency", "unsupported_opaque", "other"],
                        default=str(item.category),
                        rationale="Low-confidence rule-based classification requires human review.",
                    )
                )
            if item.category == "unsupported_opaque":
                questions.append(
                    ConfirmationQuestion(
                        id=f"risk:{item.id}",
                        prompt=f"Review unsupported or opaque workbook feature: {'; '.join(item.reasons)}",
                        question_type="unsupported_feature_review",
                        source_location=item.source_location,
                        options=["accept_record_only", "manual_follow_up", "ignore_for_phase1"],
                        default="manual_follow_up",
                        rationale="Unsupported Excel features are recorded and should be reviewed before migration.",
                    )
                )
        return ConfirmationTemplate(questions=questions)
