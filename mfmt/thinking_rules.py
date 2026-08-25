from __future__ import annotations

from dataclasses import asdict, dataclass
from enum import StrEnum
import json
from pathlib import Path
import re
from typing import Any, Iterable


class ExecutionLevel(StrEnum):
    """How far a natural-language habit may be automated safely."""

    DETERMINISTIC_CHECK = "deterministic_check"
    HUMAN_REVIEW = "human_review"
    ASSISTIVE_PROMPT = "assistive_prompt"


class DecisionStatus(StrEnum):
    ACCEPTED = "accepted"
    REVIEW_REQUIRED = "review_required"
    BLOCKED = "blocked"
    PROMPT_ONLY = "prompt_only"


@dataclass(frozen=True)
class ThinkingRule:
    rule_id: str
    group_id: str
    report_no: int
    source_file: str
    source_text: str
    core_question: str
    trigger_conditions: str
    evidence_and_uncertainty: str
    stop_or_pivot_conditions: str
    required_outputs: str
    execution_level: str
    evidence_required: bool
    human_confirmation_required: bool
    source_ambiguity: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class RuleExecutionContext:
    goal: str = ""
    stakeholders: tuple[str, ...] = ()
    evidence_status: str = "unknown"
    evidence_items: tuple[str, ...] = ()
    reversible: bool | None = None
    failure_cost: str = "unknown"
    success_criteria: tuple[str, ...] = ()
    output_destination: str = ""
    human_owner: str = ""

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "RuleExecutionContext":
        return cls(
            goal=str(value.get("goal", "")),
            stakeholders=tuple(str(item) for item in value.get("stakeholders", [])),
            evidence_status=str(value.get("evidence_status", "unknown")),
            evidence_items=tuple(str(item) for item in value.get("evidence_items", [])),
            reversible=value.get("reversible"),
            failure_cost=str(value.get("failure_cost", "unknown")),
            success_criteria=tuple(str(item) for item in value.get("success_criteria", [])),
            output_destination=str(value.get("output_destination", "")),
            human_owner=str(value.get("human_owner", "")),
        )


def _classify_execution_level(text: str, ambiguity: str) -> ExecutionLevel:
    normalized = f"{text} {ambiguity}".lower()
    human_terms = (
        "價值", "信任", "公平", "關係", "說服", "接受", "合作", "領導", "權力",
        "責任與權限", "身心安全", "誠實", "動機", "雙贏", "談判", "利害關係人",
        "真人", "專業覆核", "人工", "主管", "人負責", "人審核", "人決定",
    )
    deterministic_terms = (
        "hash", "sha-256", "公式", "錯誤", "版本", "欄位", "格式", "來源", "證據",
        "檢查表", "紀錄", "比對", "輸入", "輸出", "門檻", "計算", "數據", "資料",
        "時間", "比例", "機率", "條件", "範圍", "清單", "流程",
    )
    if any(term in normalized for term in human_terms):
        return ExecutionLevel.HUMAN_REVIEW
    if any(term in normalized for term in deterministic_terms):
        return ExecutionLevel.DETERMINISTIC_CHECK
    return ExecutionLevel.ASSISTIVE_PROMPT


def _requires_evidence(text: str, evidence: str) -> bool:
    normalized = f"{text} {evidence}".lower()
    return any(term in normalized for term in (
        "證據", "來源", "驗證", "查證", "數據", "資料", "事實", "可信", "核對", "紀錄",
    ))


def _requires_human_confirmation(level: ExecutionLevel, text: str, ambiguity: str) -> bool:
    normalized = f"{text} {ambiguity}".lower()
    return level != ExecutionLevel.DETERMINISTIC_CHECK or any(term in normalized for term in (
        "待驗證", "推論", "未明示", "人工", "專業", "真人", "主管", "例外",
    ))


class ThinkingRuleCatalog:
    def __init__(self, rules: list[ThinkingRule], summary: dict[str, Any]):
        self.rules = rules
        self.summary = summary

    @classmethod
    def from_json(cls, path: str | Path) -> "ThinkingRuleCatalog":
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        rules: list[ThinkingRule] = []
        habit_catalog = payload.get("catalog_kind") == "one_program_rule_candidate_per_habit_definition"
        for group in payload.get("rules", []):
            ambiguity = str(group.get("ambiguities_and_conflicts", ""))
            if habit_catalog and group.get("habit_program_rules"):
                candidates = [
                    str(item.get("rule_candidate", "未明示"))
                    for item in group["habit_program_rules"]
                ]
            else:
                candidates = list(group.get("program_rule_candidates", []))
            if not candidates:
                candidates = [
                    f"未萃取到明確候選；以核心問題作為 assistive prompt：{group.get('core_question', '未明示')}"
                ]
            for index, candidate in enumerate(candidates, start=1):
                level = _classify_execution_level(candidate, ambiguity)
                rules.append(
                    ThinkingRule(
                        rule_id=f"{group['rule_group_id']}-R{index:02d}",
                        group_id=str(group["rule_group_id"]),
                        report_no=int(group["report_no"]),
                        source_file=str(group["source_file"]),
                        source_text=str(candidate),
                        core_question=str(group.get("core_question", "未明示")),
                        trigger_conditions=str(group.get("trigger_conditions", "未明示")),
                        evidence_and_uncertainty=str(group.get("evidence_and_uncertainty", "未明示")),
                        stop_or_pivot_conditions=str(group.get("stop_or_pivot_conditions", "未明示")),
                        required_outputs=str(group.get("required_outputs", "未明示")),
                        execution_level=level.value,
                        evidence_required=_requires_evidence(candidate, str(group.get("evidence_and_uncertainty", ""))),
                        human_confirmation_required=_requires_human_confirmation(level, candidate, ambiguity),
                        source_ambiguity=any(
                            marker in ambiguity
                            for marker in (
                                "待驗證", "推論", "未明示", "未定義", "衝突", "矛盾",
                                "模糊", "缺乏", "不一致", "尚待",
                            )
                        ),
                    )
                )
        return cls(rules, payload.get("summary", {}))

    @property
    def report_count(self) -> int:
        return len({rule.report_no for rule in self.rules})

    def validate_completeness(self, expected_reports: int = 45) -> dict[str, Any]:
        groups = {rule.group_id for rule in self.rules}
        report_numbers = {rule.report_no for rule in self.rules}
        return {
            "expected_reports": expected_reports,
            "processed_reports": len(report_numbers),
            "expected_rule_groups": expected_reports,
            "processed_rule_groups": len(groups),
            "compiled_rules": len(self.rules),
            "complete": len(report_numbers) == expected_reports and len(groups) == expected_reports,
            "missing_report_numbers": sorted(set(range(1, expected_reports + 1)) - report_numbers),
        }

    def write_compiled_catalog(self, path: str | Path) -> None:
        output = {
            "schema_version": "joker-thinking-program-rules-v1",
            "source_summary": self.summary,
            "completeness": self.validate_completeness(),
            "execution_levels": {
                level.value: sum(rule.execution_level == level.value for rule in self.rules)
                for level in ExecutionLevel
            },
            "rules": [rule.to_dict() for rule in self.rules],
        }
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


class ThinkingRuleEngine:
    """Run safe readiness checks and emit human-reviewable rule outcomes.

    The engine never turns a subjective habit into an irreversible action. It can
    mark deterministic evidence checks as accepted, but subjective or ambiguous
    rules remain review_required or prompt_only until a person decides.
    """

    def __init__(self, catalog: ThinkingRuleCatalog):
        self.catalog = catalog

    def _status_for(self, rule: ThinkingRule, context: RuleExecutionContext) -> tuple[DecisionStatus, list[str]]:
        reasons: list[str] = []
        if not context.goal.strip():
            return DecisionStatus.BLOCKED, ["缺少一體：尚未定義這一輪真正要達成的目標"]
        if rule.source_ambiguity:
            reasons.append("來源報告含有未定義、推論或待驗證內容")
        if rule.evidence_required and context.evidence_status in {"unknown", "unverified", "conflict"}:
            reasons.append(f"證據狀態為 {context.evidence_status}，不足以自動通過")
        if "stakeholder" in rule.trigger_conditions.lower() or "利害關係人" in rule.source_text:
            if not context.stakeholders:
                return DecisionStatus.REVIEW_REQUIRED, reasons + ["缺少兩面：尚未列出受影響的其他人"]
        if any(term in rule.source_text for term in ("停損", "停止", "放棄", "轉向", "失敗")):
            if not context.success_criteria:
                return DecisionStatus.REVIEW_REQUIRED, reasons + ["尚未定義成功／失敗或轉向條件"]
        if context.failure_cost == "high" and context.reversible is False:
            return DecisionStatus.BLOCKED, reasons + ["高代價且不可逆，必須先建立小型可逆驗證"]
        if rule.execution_level == ExecutionLevel.HUMAN_REVIEW.value:
            return DecisionStatus.REVIEW_REQUIRED, reasons + ["此規則涉及價值、關係、權責或專業判斷"]
        if rule.execution_level == ExecutionLevel.ASSISTIVE_PROMPT.value:
            return DecisionStatus.PROMPT_ONLY, reasons + ["此規則目前只能提供提示，不能自動判定"]
        if rule.human_confirmation_required:
            return DecisionStatus.REVIEW_REQUIRED, reasons + ["需要人工確認來源與適用情境"]
        return DecisionStatus.ACCEPTED, reasons + ["可執行的機械檢查條件已滿足"]

    def evaluate(self, context: RuleExecutionContext | dict[str, Any]) -> dict[str, Any]:
        if isinstance(context, dict):
            context = RuleExecutionContext.from_dict(context)
        outcomes = []
        for rule in self.catalog.rules:
            status, reasons = self._status_for(rule, context)
            outcomes.append({
                "rule_id": rule.rule_id,
                "group_id": rule.group_id,
                "report_no": rule.report_no,
                "source_file": rule.source_file,
                "status": status.value,
                "execution_level": rule.execution_level,
                "source_text": rule.source_text,
                "reasons": reasons,
                "human_confirmation_required": rule.human_confirmation_required,
            })
        counts = {status.value: sum(item["status"] == status.value for item in outcomes) for status in DecisionStatus}
        overall = "blocked" if counts[DecisionStatus.BLOCKED.value] else (
            "review_required" if counts[DecisionStatus.REVIEW_REQUIRED.value] else (
                "prompt_only" if counts[DecisionStatus.PROMPT_ONLY.value] else "accepted"
            )
        )
        return {
            "schema_version": "joker-thinking-rule-evaluation-v1",
            "overall_status": overall,
            "context": asdict(context),
            "catalog_completeness": self.catalog.validate_completeness(),
            "counts": counts,
            "outcomes": outcomes,
        }

    def write_report(self, evaluation: dict[str, Any], path: str | Path) -> None:
        lines = [
            "# Thinking Rule Evaluation",
            "",
            f"**Overall status**: `{evaluation['overall_status']}`  ",
            f"**Compiled rules**: {evaluation['catalog_completeness']['compiled_rules']}  ",
            f"**Reports covered**: {evaluation['catalog_completeness']['processed_reports']} / {evaluation['catalog_completeness']['expected_reports']}  ",
            "",
            "## Summary",
            "",
            "| Status | Count |",
            "|---|---:|",
        ]
        for key, value in evaluation["counts"].items():
            lines.append(f"| `{key}` | {value} |")
        lines.extend(["", "## Outcomes requiring attention", ""])
        lines.append("| Rule | Report | Status | Execution level | Reason |")
        lines.append("|---|---:|---|---|---|")
        for item in evaluation["outcomes"]:
            if item["status"] == DecisionStatus.ACCEPTED.value:
                continue
            reason = "；".join(item["reasons"]).replace("|", "\\|")
            lines.append(
                f"| `{item['rule_id']}` | {item['report_no']} | `{item['status']}` | `{item['execution_level']}` | {reason} |"
            )
        lines.extend([
            "",
            "> This evaluation is a decision aid. It does not replace the user's judgment or domain review.",
        ])
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("\n".join(lines) + "\n", encoding="utf-8")
