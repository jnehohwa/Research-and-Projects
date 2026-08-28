import json

from educore.bugzot import Bugzot, BugzotEvent


def test_bugzot_records_useful_context() -> None:
    monitor = Bugzot()
    monitor.record_event(
        BugzotEvent(
            event_type="duplicate_registration",
            severity="WARNING",
            component="registration_service",
            outcome="rejected_duplicate",
            message="duplicate prevented",
            correlation_id="TX-001",
            learner_id="L001",
            course_id="PY701",
        )
    )
    event = monitor.events[0]
    assert event.correlation_id == "TX-001"
    assert event.learner_id == "L001"


def test_bugzot_generates_performance_metrics() -> None:
    monitor = Bugzot()
    for outcome, duration in [("confirmed", 1.0), ("confirmed", 2.0), ("rejected_capacity", 3.0)]:
        monitor.record_event(
            BugzotEvent(
                event_type="registration_transaction",
                severity="INFO",
                component="registration_service",
                outcome=outcome,
                message="processed",
                correlation_id=outcome,
                duration_ms=duration,
            )
        )
    report = monitor.generate_report()
    assert report["total_events"] == 3
    assert report["successful_transactions"] == 2
    assert report["average_duration_ms"] == 2.0
    assert report["p95_duration_ms"] == 3.0
    assert report["throughput_transactions_per_second"] > 0


def test_bugzot_exports_json(tmp_path) -> None:
    monitor = Bugzot()
    monitor.record_event(
        BugzotEvent("test", "INFO", "tests", "success", "ok", "TX-001", duration_ms=1.0)
    )
    destination = tmp_path / "report.json"
    monitor.export_json(destination)
    assert json.loads(destination.read_text())["total_events"] == 1
