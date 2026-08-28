"""Sequential and concurrent learner registration processing."""

from __future__ import annotations

from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from threading import Event, Lock
from uuid import uuid4

from .bugzot import Bugzot
from .domain import Course, Learner, Registration, RegistrationStatus


@dataclass(frozen=True, slots=True)
class RegistrationRequest:
    learner_id: str
    course_id: str
    correlation_id: str = ""

    def resolved_correlation_id(self) -> str:
        return self.correlation_id or f"TX-{uuid4().hex[:10].upper()}"


@dataclass(frozen=True, slots=True)
class RegistrationResult:
    learner_id: str
    course_id: str
    status: RegistrationStatus
    message: str
    correlation_id: str
    registration: Registration | None = None


class RegistrationService:
    """Enforces registration invariants under sequential or concurrent load."""

    def __init__(self, learners: list[Learner], courses: list[Course], bugzot: Bugzot) -> None:
        self._learners = {learner.learner_id: learner for learner in learners}
        self._courses = {course.course_id: course for course in courses}
        self._registrations: dict[tuple[str, str], Registration] = {}
        self._course_counts: Counter[str] = Counter()
        self._lock = Lock()
        self._bugzot = bugzot

    @property
    def registrations(self) -> tuple[Registration, ...]:
        with self._lock:
            return tuple(self._registrations.values())

    @property
    def monitor(self) -> Bugzot:
        return self._bugzot

    def course_registration_count(self, course_id: str) -> int:
        with self._lock:
            return self._course_counts[course_id]

    def process(self, request: RegistrationRequest) -> RegistrationResult:
        correlation_id = request.resolved_correlation_id()
        with self._bugzot.measure_transaction(
            event_type="registration_transaction",
            correlation_id=correlation_id,
            learner_id=request.learner_id,
            course_id=request.course_id,
        ) as transaction:
            result = self._process_atomically(request, correlation_id)
            transaction["outcome"] = result.status.value
            transaction["message"] = result.message
            transaction["event_type"] = {
                RegistrationStatus.CONFIRMED: "registration_success",
                RegistrationStatus.REJECTED_DUPLICATE: "duplicate_registration",
                RegistrationStatus.REJECTED_CAPACITY: "course_capacity_violation",
                RegistrationStatus.REJECTED_INVALID: "validation_failure",
            }[result.status]
            return result

    def _process_atomically(
        self, request: RegistrationRequest, correlation_id: str
    ) -> RegistrationResult:
        with self._lock:
            learner = self._learners.get(request.learner_id)
            course = self._courses.get(request.course_id)
            if learner is None or course is None or not learner.active:
                return RegistrationResult(
                    request.learner_id,
                    request.course_id,
                    RegistrationStatus.REJECTED_INVALID,
                    "learner or course is invalid",
                    correlation_id,
                )

            key = (request.learner_id, request.course_id)
            if key in self._registrations:
                return RegistrationResult(
                    request.learner_id,
                    request.course_id,
                    RegistrationStatus.REJECTED_DUPLICATE,
                    "duplicate registration prevented",
                    correlation_id,
                )

            if self._course_counts[request.course_id] >= course.capacity:
                return RegistrationResult(
                    request.learner_id,
                    request.course_id,
                    RegistrationStatus.REJECTED_CAPACITY,
                    "course capacity reached",
                    correlation_id,
                )

            registration = Registration(request.learner_id, request.course_id)
            self._registrations[key] = registration
            self._course_counts[request.course_id] += 1
            return RegistrationResult(
                request.learner_id,
                request.course_id,
                RegistrationStatus.CONFIRMED,
                "registration confirmed",
                correlation_id,
                registration,
            )

    def process_concurrently(
        self, requests: list[RegistrationRequest], max_workers: int = 8
    ) -> list[RegistrationResult]:
        if max_workers <= 0:
            raise ValueError("max_workers must be positive")
        start_signal = Event()

        def process_after_signal(request: RegistrationRequest) -> RegistrationResult:
            start_signal.wait()
            return self.process(request)

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(process_after_signal, request) for request in requests]
            start_signal.set()
            return [future.result() for future in futures]

    @staticmethod
    def summarise(results: list[RegistrationResult]) -> dict[str, int]:
        counts = Counter(result.status.value for result in results)
        return {status.value: counts.get(status.value, 0) for status in RegistrationStatus}
