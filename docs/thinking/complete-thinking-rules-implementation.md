# Complete Thinking Habits to Program Rules

## Decision

The project now uses the complete supplied thinking decision database as a source-derived rule catalog. The implementation does not treat the database as a collection of slogans. It preserves each report's source, core question, habit definition, decision sequence, trigger, evidence and uncertainty, stop or pivot condition, required output, ambiguity, and rule candidate.

All 45 numbered reports were read in full. The complete source-derived data contains 235 habit-definition lines and 185 report-level program-rule candidate lines. These are retained as two related views because they answer different completeness questions: the report-level view preserves what each report explicitly proposed as a candidate, while the habit-level view guarantees that every individual habit definition receives a corresponding rule candidate.

## Two catalog views

| View | Count | Purpose | File |
|---|---:|---|---|
| Report-level candidates | 185 | Preserve explicit candidate rules extracted from each report | `resources/thinking/thinking_rule_catalog.json` |
| Habit-level candidates | 235 | Create one rule candidate for every habit definition | `resources/thinking/habit_program_catalog.json` |
| Compiled report rules | 185 | Runtime rule records for report-level evaluation | `resources/thinking/compiled_program_rules.json` |
| Compiled habit rules | 235 | Runtime rule records for complete habit-level evaluation | `resources/thinking/compiled_habit_program_rules.json` |

The full reading checklist is `docs/thinking/thinking_habits_checklist.md`. The raw extraction result is retained under `resources/thinking/extract_all_thinking_rules.json` so each normalized record can be traced back to the individual report processed by the extraction step.

## Program-rule contract

Every habit-level rule is represented as an `IF / THEN / GUARD / STOP-PIVOT / OUTPUT` candidate. The candidate is deliberately not treated as production automation merely because it has been formatted like a rule.

```text
IF       the report's trigger condition is present
THEN     perform the habit's stated thinking action
GUARD    check the report's evidence and uncertainty requirements
STOP     stop, pivot, reduce scope, or ask for a person when its boundary is reached
OUTPUT   record the decision, evidence, result, and next reusable asset
```

Each rule also preserves `execution_level`, `formalization_status`, `activation_status`, `human_confirmation_required`, and `source_ambiguity`.

## Execution levels

The engine uses three execution levels. `deterministic_check` is limited to explicit facts that can be checked mechanically. `human_review` covers values, relationships, trust, fairness, persuasion, responsibility, safety, and professional judgment. `assistive_prompt` covers useful questions whose operational meaning is not sufficiently defined for reliable automatic evaluation.

All 235 habit-level rules currently have `formalization_status=candidate`, `activation_status=disabled_until_owner_confirmation`, and `human_confirmation_required=true`. This is intentional. The database includes multiple items marked as inference, pending verification, ambiguous, or dependent on context-specific thresholds. Turning these into automatic actions without owner confirmation would contradict the source database's own evidence and stop-condition principles.

## Preflight context

The runtime engine accepts a structured decision context:

| Field | Meaning |
|---|---|
| `goal` | The one thing being decided or achieved |
| `stakeholders` | People, teams, customers, reviewers, or future users affected |
| `evidence_status` | `unknown`, `unverified`, `verified`, or `conflict` |
| `evidence_items` | Evidence already collected |
| `reversible` | Whether the next action can be undone or safely rolled back |
| `failure_cost` | Approximate consequence class such as `low`, `medium`, or `high` |
| `success_criteria` | Conditions for success, failure, stop, or pivot |
| `output_destination` | Where the result will become a reusable asset |
| `human_owner` | Person responsible for final judgment |

A missing goal blocks evaluation. A high-cost irreversible action blocks until a small reversible experiment is defined. Unknown, unverified, or conflicting evidence requires review when the rule depends on evidence. Subjective or ambiguous rules remain `review_required`; assistive rules remain `prompt_only`.

## FMEDA integration

The rule engine is intentionally separate from the FMEDA numerical validator, but the two can share the same project context. The FMEDA pipeline supplies concrete evidence such as source hashes, formula catalogs, external workbook hashes, validation reports, reviewer manifests, and derived revision IDs. The thinking-rule engine then checks whether the development decision has defined its goal, affected people, evidence, reversibility, failure cost, success conditions, human owner, and reusable output.

This means the tool can answer questions such as whether a proposed external-link materialization has enough evidence for a production decision, while still leaving the actual FMEDA acceptance decision to the designated engineer. It can also require a reversible trial before a high-cost migration, and it can ensure that an output is documented instead of disappearing after one run.

## Verification result

The complete habit catalog compiles successfully with 45 reports, 45 rule groups, and 235 compiled habit rules. The project test suite passes with 31 tests. A representative FMEDA context produces `review_required` for all 235 habit-level rules, not because the rules failed to load, but because every habit-level rule is intentionally awaiting owner confirmation before it can become an enabled decision rule.

## Next activation step

The next step is not to activate all 235 rules at once. The owner should confirm each rule's intended operational meaning, especially where the source records missing numerical thresholds, inferred habits, conflicting numbers, or context-dependent definitions. A confirmed rule can then move from `candidate` to an enabled rule with a specific test fixture and a defined action. Unconfirmed rules remain available as review prompts and evidence checklists.
