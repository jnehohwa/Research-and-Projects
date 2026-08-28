"""Reproducible CLI demonstration used for assessment evidence."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from .benchmark import benchmark_as_dict
from .bugzot import Bugzot
from .domain import Assessment, Course, Learner, TicketCategory
from .patterns import (
    ApplicationConfig,
    BestAttemptStrategy,
    PassFailStrategy,
    SupportTicketFactory,
    WeightedAverageStrategy,
)
from .registration import RegistrationRequest, RegistrationService


def build_demo_service(capacity: int = 10) -> RegistrationService:
    learners = [
        Learner(f"L{index:03d}", f"Learner {index}", f"learner{index}@example.org")
        for index in range(1, 21)
    ]
    return RegistrationService(learners, [Course("PY701", "Enterprise Python", capacity)], Bugzot())


def run_demo(report_directory: Path) -> None:
    print("EDUCORE ENTERPRISE APPLICATION DEMONSTRATION")
    print("=" * 56)

    learner = Learner("L001", "Thabo Mokoena", "thabo@example.org")
    course = Course("PY701", "Enterprise Python Development", 10)
    assessment = Assessment("A001", course.course_id, "Concurrency Practical")
    course.add_assessment(assessment.assessment_id)
    print(f"Domain: {learner.name} | {course.title} | {assessment.title}")

    config_one = ApplicationConfig()
    config_two = ApplicationConfig()
    print(f"Singleton: same instance = {config_one is config_two}; id = {id(config_one)}")

    for category in TicketCategory:
        ticket = SupportTicketFactory.create(category, learner.learner_id, f"{category.title()} help", "Please assist")
        print(f"Factory: {ticket.category.value:<9} -> priority {ticket.priority.value}")

    scores = [68, 76, 81]
    print(f"Strategy weighted average: {WeightedAverageStrategy((0.2, 0.3, 0.5)).calculate(scores)}")
    print(f"Strategy best attempt: {BestAttemptStrategy().calculate(scores)}")
    print(f"Strategy pass/fail: {PassFailStrategy(50).calculate(scores)}")

    service = build_demo_service(capacity=10)
    requests = [RegistrationRequest(f"L{index:03d}", "PY701") for index in range(1, 16)]
    requests.append(RegistrationRequest("L001", "PY701"))
    results = service.process_concurrently(requests, max_workers=8)
    for result in results:
        print(f"Registration: {result.learner_id} -> {result.status.value}")
    print("Registration summary:", json.dumps(service.summarise(results), sort_keys=True))
    print(f"Integrity: confirmed={service.course_registration_count('PY701')} capacity=10")

    bugzot = service.monitor
    report_directory.mkdir(parents=True, exist_ok=True)
    bugzot.export_json(report_directory / "bugzot_performance_report.json")
    bugzot.export_csv(report_directory / "bugzot_events.csv")
    print("Bugzot report:", json.dumps(bugzot.generate_report(), sort_keys=True))
    print("Optimisation benchmark:", json.dumps(benchmark_as_dict(), sort_keys=True))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report-directory", type=Path, default=Path("evidence/reports"))
    args = parser.parse_args()
    run_demo(args.report_directory)


if __name__ == "__main__":
    main()
