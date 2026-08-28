"""Build the two focused evidence images requested during report review."""

from __future__ import annotations

import ast
import csv
import json
from pathlib import Path

from build_evidence import ROOT, SCREENSHOTS, render_terminal


def function_source(path: Path, function_name: str) -> list[str]:
    source = path.read_text(encoding="utf-8")
    lines = source.splitlines()
    module = ast.parse(source)
    function = next(
        node
        for node in module.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name == function_name
    )
    decorator_lines = [decorator.lineno for decorator in function.decorator_list]
    start_line = min(decorator_lines, default=function.lineno)
    return lines[start_line - 1 : function.end_lineno]


def main() -> None:
    SCREENSHOTS.mkdir(parents=True, exist_ok=True)

    events_path = ROOT / "evidence" / "reports" / "bugzot_events.csv"
    with events_path.open(newline="", encoding="utf-8") as stream:
        events = list(csv.DictReader(stream))

    selected_events = []
    for outcome in ("confirmed", "rejected_capacity", "rejected_duplicate"):
        selected_events.append(next(event for event in events if event["outcome"] == outcome))

    render_terminal(
        "Bugzot raw event records",
        "python -m educore.demo --report-directory evidence/reports",
        json.dumps(selected_events, indent=2).splitlines(),
        SCREENSHOTS / "bugzot_raw_events.png",
    )

    test_path = ROOT / "tests" / "test_domain.py"
    render_terminal(
        "Representative parameterised test",
        "pytest tests/test_domain.py::test_learner_rejects_invalid_email",
        function_source(test_path, "test_learner_rejects_invalid_email"),
        SCREENSHOTS / "representative_test_case.png",
    )


if __name__ == "__main__":
    main()
