"""Run reproducible checks and render their real outputs as evidence images."""

from __future__ import annotations

import io
import json
import os
from pathlib import Path
import pstats
import subprocess
import sys
import textwrap

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "evidence" / "reports"
PROFILING = ROOT / "evidence" / "profiling"
SCREENSHOTS = ROOT / "evidence" / "screenshots"
PYTHON = ROOT / ".venv" / "bin" / "python"
FONT_REGULAR = "/System/Library/Fonts/Supplemental/Arial.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
FONT_MONO = "/System/Library/Fonts/SFNSMono.ttf"


def run(command: list[str], *, env: dict[str, str] | None = None) -> str:
    result = subprocess.run(
        command,
        cwd=ROOT,
        env=env,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    if result.returncode:
        raise RuntimeError(f"Command failed ({result.returncode}): {' '.join(command)}\n{result.stdout}")
    return result.stdout


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        return ImageFont.load_default()


def render_terminal(title: str, command: str, lines: list[str], destination: Path) -> None:
    width = 1600
    wrapped: list[str] = []
    for line in lines:
        wrapped.extend(textwrap.wrap(line, width=118, subsequent_indent="  ") or [""])
    height = max(500, 165 + len(wrapped) * 29 + 45)
    image = Image.new("RGB", (width, height), "#0b1220")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((18, 18, width - 18, height - 18), radius=22, fill="#111827", outline="#334155", width=3)
    draw.rectangle((20, 20, width - 20, 92), fill="#1e293b")
    for index, colour in enumerate(("#fb7185", "#fbbf24", "#34d399")):
        draw.ellipse((45 + index * 36, 45, 65 + index * 36, 65), fill=colour)
    draw.text((175, 39), title, fill="#e2e8f0", font=font(FONT_BOLD, 28))
    draw.text((45, 112), f"$ {command}", fill="#5eead4", font=font(FONT_MONO, 22))
    y = 158
    mono = font(FONT_MONO, 21)
    for line in wrapped:
        colour = "#86efac" if any(token in line for token in ("passed", "confirmed", "True", "93%")) else "#dbeafe"
        draw.text((45, y), line, fill=colour, font=mono)
        y += 29
    destination.parent.mkdir(parents=True, exist_ok=True)
    image.save(destination)


def main() -> None:
    REPORTS.mkdir(parents=True, exist_ok=True)
    PROFILING.mkdir(parents=True, exist_ok=True)
    SCREENSHOTS.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["PYTHONPATH"] = str(ROOT / "src")

    tests = run([str(PYTHON), "-m", "pytest", "--cov=educore", "--cov-report=term-missing"], env=env)
    (REPORTS / "test_output.txt").write_text(tests, encoding="utf-8")
    relevant_tests = [
        line for line in tests.splitlines()
        if line.startswith(("platform ", "collected ", "tests/", "src/", "TOTAL", "="))
    ][-25:]
    render_terminal(
        "Automated test and coverage evidence",
        "python -m pytest --cov=educore --cov-report=term-missing",
        relevant_tests,
        SCREENSHOTS / "automated_tests.png",
    )

    demo = run([str(PYTHON), "-m", "educore.demo", "--report-directory", str(REPORTS)], env=env)
    (REPORTS / "demo_output.txt").write_text(demo, encoding="utf-8")
    demo_lines = demo.splitlines()
    render_terminal(
        "Domain model and design patterns",
        "python -m educore.demo",
        demo_lines[:12],
        SCREENSHOTS / "domain_and_patterns.png",
    )
    registration_lines = [
        line for line in demo_lines
        if line.startswith(("Registration:", "Registration summary:", "Integrity:"))
    ]
    render_terminal(
        "Concurrent registration processing",
        "python -m educore.demo",
        registration_lines,
        SCREENSHOTS / "concurrent_registration.png",
    )

    report = json.loads((REPORTS / "bugzot_performance_report.json").read_text(encoding="utf-8"))
    render_terminal(
        "Bugzot performance report",
        "python -m educore.demo --report-directory evidence/reports",
        json.dumps(report, indent=2, sort_keys=True).splitlines(),
        SCREENSHOTS / "bugzot_report.png",
    )

    sys.path.insert(0, str(ROOT / "src"))
    from educore.benchmark import benchmark_as_dict

    benchmark = benchmark_as_dict()
    (REPORTS / "optimisation_benchmark.json").write_text(json.dumps(benchmark, indent=2), encoding="utf-8")
    render_terminal(
        "Before-and-after optimisation benchmark",
        "python -m educore.benchmark",
        json.dumps(benchmark, indent=2).splitlines(),
        SCREENSHOTS / "optimisation_benchmark.png",
    )

    profile_path = PROFILING / "registration_profile.prof"
    run(
        [str(PYTHON), "-m", "cProfile", "-o", str(profile_path), "-m", "educore.demo", "--report-directory", str(REPORTS)],
        env=env,
    )
    stream = io.StringIO()
    pstats.Stats(str(profile_path), stream=stream).strip_dirs().sort_stats("cumulative").print_stats(18)
    profile_summary = stream.getvalue()
    (PROFILING / "profile_summary.txt").write_text(profile_summary, encoding="utf-8")
    profile_lines = [line.rstrip() for line in profile_summary.splitlines() if line.strip()][2:23]
    render_terminal(
        "cProfile cumulative-time analysis",
        "python -m cProfile -o evidence/profiling/registration_profile.prof -m educore.demo",
        profile_lines,
        SCREENSHOTS / "profiling_results.png",
    )


if __name__ == "__main__":
    main()

