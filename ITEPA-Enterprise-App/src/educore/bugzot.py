"""Thread-safe structured monitoring and performance reporting."""

from __future__ import annotations

from collections import Counter
from contextlib import contextmanager
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
import csv
import json
from pathlib import Path
from statistics import mean
from threading import Lock
from time import perf_counter
from typing import Iterator


@dataclass(frozen=True, slots=True)
class BugzotEvent:
    event_type: str
    severity: str
    component: str
    outcome: str
    message: str
    correlation_id: str
    learner_id: str | None = None
    course_id: str | None = None
    duration_ms: float | None = None
    timestamp: datetime = field(default_factory=lambda: datetime.now(UTC))

    def to_record(self) -> dict[str, object]:
        record = asdict(self)
        record["timestamp"] = self.timestamp.isoformat()
        return record


class Bugzot:
    """Collects diagnostic events and transaction metrics in memory."""

    def __init__(self) -> None:
        self._events: list[BugzotEvent] = []
        self._lock = Lock()

    @property
    def events(self) -> tuple[BugzotEvent, ...]:
        with self._lock:
            return tuple(self._events)

    def record_event(self, event: BugzotEvent) -> None:
        with self._lock:
            self._events.append(event)

    @contextmanager
    def measure_transaction(
        self,
        *,
        event_type: str,
        correlation_id: str,
        learner_id: str | None = None,
        course_id: str | None = None,
    ) -> Iterator[dict[str, str]]:
        state = {
            "outcome": "success",
            "message": "transaction completed",
            "event_type": event_type,
        }
        started_at = perf_counter()
        try:
            yield state
        except Exception as error:
            state.update(outcome="error", message=str(error))
            raise
        finally:
            self.record_event(
                BugzotEvent(
                    event_type=state["event_type"],
                    severity=(
                        "ERROR"
                        if state["outcome"] == "error"
                        else "WARNING"
                        if state["outcome"].startswith("rejected_")
                        else "INFO"
                    ),
                    component="registration_service",
                    outcome=state["outcome"],
                    message=state["message"],
                    correlation_id=correlation_id,
                    learner_id=learner_id,
                    course_id=course_id,
                    duration_ms=round((perf_counter() - started_at) * 1000, 4),
                )
            )

    def generate_report(self) -> dict[str, object]:
        events = self.events
        durations = [event.duration_ms for event in events if event.duration_ms is not None]
        outcomes = Counter(event.outcome for event in events)
        types = Counter(event.event_type for event in events)
        ordered = sorted(durations)
        p95_index = max(0, int(len(ordered) * 0.95 + 0.9999) - 1) if ordered else 0
        successful = outcomes.get("success", 0) + outcomes.get("confirmed", 0)
        if events:
            elapsed_seconds = (
                max(event.timestamp for event in events) - min(event.timestamp for event in events)
            ).total_seconds()
            elapsed_seconds += (max(durations) / 1000) if durations else 0.0
        else:
            elapsed_seconds = 0.0
        return {
            "total_events": len(events),
            "successful_transactions": successful,
            "failed_transactions": len(events) - successful,
            "success_rate_percent": round(successful / len(events) * 100, 2) if events else 0.0,
            "average_duration_ms": round(mean(durations), 4) if durations else 0.0,
            "minimum_duration_ms": round(min(durations), 4) if durations else 0.0,
            "maximum_duration_ms": round(max(durations), 4) if durations else 0.0,
            "p95_duration_ms": round(ordered[p95_index], 4) if ordered else 0.0,
            "throughput_transactions_per_second": (
                round(len(events) / elapsed_seconds, 2) if elapsed_seconds > 0 else 0.0
            ),
            "events_by_type": dict(sorted(types.items())),
            "outcomes": dict(sorted(outcomes.items())),
        }

    def export_json(self, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(json.dumps(self.generate_report(), indent=2), encoding="utf-8")

    def export_csv(self, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        records = [event.to_record() for event in self.events]
        if not records:
            destination.write_text("", encoding="utf-8")
            return
        with destination.open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=records[0].keys())
            writer.writeheader()
            writer.writerows(records)
