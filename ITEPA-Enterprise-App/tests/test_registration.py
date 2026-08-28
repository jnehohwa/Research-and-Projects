from collections import Counter

from educore.bugzot import Bugzot
from educore.domain import Course, Learner, RegistrationStatus
from educore.registration import RegistrationRequest, RegistrationService


def build_service(capacity: int = 5, learner_count: int = 20) -> RegistrationService:
    learners = [
        Learner(f"L{index:03d}", f"Learner {index}", f"learner{index}@example.org")
        for index in range(1, learner_count + 1)
    ]
    return RegistrationService(learners, [Course("PY701", "Enterprise Python", capacity)], Bugzot())


def test_successful_registration() -> None:
    result = build_service().process(RegistrationRequest("L001", "PY701"))
    assert result.status is RegistrationStatus.CONFIRMED
    assert result.registration is not None


def test_invalid_learner_is_rejected() -> None:
    result = build_service().process(RegistrationRequest("UNKNOWN", "PY701"))
    assert result.status is RegistrationStatus.REJECTED_INVALID


def test_duplicate_registration_is_prevented() -> None:
    service = build_service()
    service.process(RegistrationRequest("L001", "PY701"))
    duplicate = service.process(RegistrationRequest("L001", "PY701"))
    assert duplicate.status is RegistrationStatus.REJECTED_DUPLICATE
    assert len(service.registrations) == 1


def test_capacity_is_enforced() -> None:
    service = build_service(capacity=1)
    first = service.process(RegistrationRequest("L001", "PY701"))
    second = service.process(RegistrationRequest("L002", "PY701"))
    assert first.status is RegistrationStatus.CONFIRMED
    assert second.status is RegistrationStatus.REJECTED_CAPACITY


def test_concurrent_processing_protects_capacity_and_uniqueness() -> None:
    capacity = 5
    service = build_service(capacity=capacity)
    requests = [RegistrationRequest(f"L{index:03d}", "PY701") for index in range(1, 21)]
    requests.extend(RegistrationRequest("L001", "PY701") for _ in range(5))
    results = service.process_concurrently(requests, max_workers=8)
    statuses = Counter(result.status for result in results)

    assert statuses[RegistrationStatus.CONFIRMED] == capacity
    assert service.course_registration_count("PY701") == capacity
    assert len({(item.learner_id, item.course_id) for item in service.registrations}) == capacity


def test_summary_includes_zero_count_statuses() -> None:
    service = build_service()
    summary = service.summarise([service.process(RegistrationRequest("L001", "PY701"))])
    assert summary == {
        "confirmed": 1,
        "rejected_duplicate": 0,
        "rejected_capacity": 0,
        "rejected_invalid": 0,
    }

