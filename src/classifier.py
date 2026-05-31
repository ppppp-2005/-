from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass
class Rule:
    rule_type: str
    patterns: list[re.Pattern[str]]


def compile_rules(raw_rules: list[dict]) -> list[Rule]:
    compiled: list[Rule] = []
    for item in raw_rules:
        compiled.append(
            Rule(
                rule_type=item["type"],
                patterns=[re.compile(p) for p in item["patterns"]],
            )
        )
    return compiled


def classify_paragraph(text: str, rules: list[Rule]) -> str:
    normalized = text.strip()
    if not normalized:
        return "empty"

    for rule in rules:
        for pattern in rule.patterns:
            if pattern.search(normalized):
                return rule.rule_type

    return "body"
