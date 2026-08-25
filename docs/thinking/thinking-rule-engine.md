# Joker Thinking Rule Engine

## Purpose

This directory records the complete, report-derived thinking-rule catalog used by the project. The source archive contains 45 numbered decision reports. Each report was read in full and converted into a structured group with a core question, habit definitions, decision sequence, trigger conditions, evidence and uncertainty, stop or pivot conditions, required outputs, ambiguities, and program-rule candidates.

The catalog is evidence-preserving. A candidate written in natural language is not silently promoted to an irreversible automated action. The engine therefore separates rule records from rule execution level.

## Coverage

| Measure | Result |
|---|---:|
| Numbered reports read | 45 |
| Processed report groups | 45 |
| Habit definition lines | 235 |
| Program-rule candidate lines | 185 |
| Source errors | 0 |
| Catalog schema | `joker-thinking-rule-catalog-v1` |
| Compiled schema | `joker-thinking-program-rules-v1` |

The full machine-readable source-derived catalog is `resources/thinking/thinking_rule_catalog.json`. The compiled rule records are written to `resources/thinking/compiled_program_rules.json`. The reading checklist is `docs/thinking/thinking_habits_checklist.md`.

## Execution levels

The engine classifies each candidate into one of three levels. This is an implementation boundary, not a claim that a natural-language judgment has become objective.

| Execution level | Meaning | Default behavior |
|---|---|---|
| `deterministic_check` | The candidate can be checked mechanically from explicit data such as a file, value, formula, hash, status, count, or declared criterion. | May return `accepted` only when all required evidence is present. |
| `human_review` | The candidate involves value, trust, fairness, relationship, responsibility, professional judgment, or an ambiguous source definition. | Returns `review_required`; the engine does not decide for the user. |
| `assistive_prompt` | The candidate is useful as a prompt or question, but its operational meaning is not yet sufficiently defined for a reliable automatic check. | Returns `prompt_only`. |

## Evaluation contract

A decision context may contain `goal`, `stakeholders`, `evidence_status`, `evidence_items`, `reversible`, `failure_cost`, `success_criteria`, `output_destination`, and `human_owner`. These fields implement the user's 4! structure without pretending that every field can be inferred automatically.

The engine applies the following guards in order. A missing goal blocks the evaluation because the one thing being decided has not been defined. A high-cost and irreversible action blocks until a small reversible experiment is specified. An unverified, unknown, or conflicting evidence state requires review when the rule requires evidence. A missing stakeholder list requires review for rules that explicitly concern affected people. Rules involving subjective or ambiguous material remain review-required even when the input data appears complete.

## Usage

Compile the complete catalog:

```powershell
thinking-rules compile `
  --catalog resources/thinking/thinking_rule_catalog.json `
  --output resources/thinking/compiled_program_rules.json
```

Evaluate a project decision context:

```powershell
thinking-rules evaluate `
  --catalog resources/thinking/thinking_rule_catalog.json `
  --context demo-output/thinking-rules/fmeda-context.json `
  --json-output demo-output/thinking-rules/evaluation.json `
  --markdown-output demo-output/thinking-rules/evaluation.md
```

The JSON output is intended for Editor or pipeline integration. The Markdown output is intended for a person to review. Both are generated from the same evaluation object.

## What this engine does not do

The engine does not claim that every report has a complete numerical threshold. The normalized catalog records ambiguities and conflicts explicitly, including missing definitions of confidence, utility, fairness, cost, sample representativeness, responsibility-authority symmetry, and other context-dependent terms. Those items remain human-review or prompt-only until a domain owner defines and verifies an operational standard.

The engine also does not make a final domain decision for FMEDA, safety, compliance, personnel, commercial, or relationship questions. Its job is to make the relevant questions, evidence gaps, affected parties, and stop conditions visible before a person commits to an action.
