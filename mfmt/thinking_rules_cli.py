from __future__ import annotations

import argparse
import json
from pathlib import Path

from .thinking_rules import RuleExecutionContext, ThinkingRuleCatalog, ThinkingRuleEngine


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compile and evaluate the user's complete thinking-rule catalog."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    compile_parser = subparsers.add_parser("compile", help="Compile report candidates into program-rule records")
    compile_parser.add_argument("--catalog", type=Path, required=True)
    compile_parser.add_argument("--output", type=Path, required=True)

    evaluate_parser = subparsers.add_parser("evaluate", help="Evaluate a decision context against all compiled rules")
    evaluate_parser.add_argument("--catalog", type=Path, required=True)
    evaluate_parser.add_argument("--context", type=Path, required=True)
    evaluate_parser.add_argument("--json-output", type=Path, required=True)
    evaluate_parser.add_argument("--markdown-output", type=Path, required=True)

    args = parser.parse_args()
    catalog = ThinkingRuleCatalog.from_json(args.catalog)
    if args.command == "compile":
        catalog.write_compiled_catalog(args.output)
        print(json.dumps(catalog.validate_completeness(), ensure_ascii=False))
        return 0

    context = json.loads(args.context.read_text(encoding="utf-8"))
    evaluation = ThinkingRuleEngine(catalog).evaluate(context)
    args.json_output.parent.mkdir(parents=True, exist_ok=True)
    args.json_output.write_text(json.dumps(evaluation, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    ThinkingRuleEngine(catalog).write_report(evaluation, args.markdown_output)
    print(json.dumps({"overall_status": evaluation["overall_status"], "counts": evaluation["counts"]}, ensure_ascii=False))
    return 2 if evaluation["overall_status"] == "blocked" else 0


if __name__ == "__main__":
    raise SystemExit(main())
