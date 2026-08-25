from __future__ import annotations

from pathlib import Path

from mfmt.thinking_rules import (
    DecisionStatus,
    RuleExecutionContext,
    ThinkingRuleCatalog,
    ThinkingRuleEngine,
)


ROOT = Path(__file__).parents[1]
CATALOG_PATH = ROOT / "resources" / "thinking" / "thinking_rule_catalog.json"


def test_all_reports_and_candidates_are_compiled_with_source_trace() -> None:
    catalog = ThinkingRuleCatalog.from_json(CATALOG_PATH)
    completeness = catalog.validate_completeness()

    assert completeness["complete"] is True
    assert completeness["processed_reports"] == 45
    assert completeness["compiled_rules"] == 185
    assert all(rule.source_file and rule.source_text for rule in catalog.rules)
    assert all(rule.rule_id.startswith("TH-") for rule in catalog.rules)


def test_missing_goal_blocks_every_rule() -> None:
    engine = ThinkingRuleEngine(ThinkingRuleCatalog.from_json(CATALOG_PATH))
    evaluation = engine.evaluate({"evidence_status": "verified"})

    assert evaluation["overall_status"] == DecisionStatus.BLOCKED.value
    assert evaluation["counts"][DecisionStatus.BLOCKED.value] == 185
    assert all("一體" in item["reasons"][0] for item in evaluation["outcomes"])


def test_complete_low_risk_context_preserves_human_and_prompt_boundaries() -> None:
    engine = ThinkingRuleEngine(ThinkingRuleCatalog.from_json(CATALOG_PATH))
    context = RuleExecutionContext(
        goal="建立可追溯的 FMEDA 文字化流程",
        stakeholders=("FMEDA 工程師", "審查者", "主管", "跨部門", "Editor 使用者"),
        evidence_status="verified",
        evidence_items=("source hash", "formula catalog", "validation report"),
        reversible=True,
        failure_cost="low",
        success_criteria=("公式可回查", "原始檔不變"),
        output_destination="GitHub",
        human_owner="FMEDA Engineer",
    )
    evaluation = engine.evaluate(context)

    assert evaluation["overall_status"] in {"review_required", "prompt_only", "accepted"}
    assert evaluation["counts"][DecisionStatus.REVIEW_REQUIRED.value] > 0
    assert evaluation["counts"][DecisionStatus.PROMPT_ONLY.value] > 0
    assert evaluation["counts"][DecisionStatus.BLOCKED.value] == 0


def test_high_cost_irreversible_context_is_blocked() -> None:
    engine = ThinkingRuleEngine(ThinkingRuleCatalog.from_json(CATALOG_PATH))
    evaluation = engine.evaluate(
        {
            "goal": "直接全面切換流程",
            "evidence_status": "verified",
            "reversible": False,
            "failure_cost": "high",
            "success_criteria": ["完成"],
        }
    )

    assert evaluation["overall_status"] == DecisionStatus.BLOCKED.value
    assert evaluation["counts"][DecisionStatus.BLOCKED.value] > 0
    assert any("不可逆" in reason for item in evaluation["outcomes"] for reason in item["reasons"])


def test_complete_habit_catalog_compiles_one_rule_per_habit_definition() -> None:
    habit_catalog = ROOT / "resources" / "thinking" / "habit_program_catalog.json"
    catalog = ThinkingRuleCatalog.from_json(habit_catalog)
    completeness = catalog.validate_completeness()

    assert completeness["complete"] is True
    assert completeness["processed_reports"] == 45
    assert completeness["compiled_rules"] == 235
    assert all(rule.rule_id.startswith("TH-") for rule in catalog.rules)
    assert all(rule.human_confirmation_required is True for rule in catalog.rules)
