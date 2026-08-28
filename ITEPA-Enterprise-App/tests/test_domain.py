import pytest

from educore.domain import (
    Assessment,
    Course,
    Learner,
    SupportTicket,
    TicketCategory,
    TicketPriority,
    ValidationError,
)


def test_learner_normalises_email() -> None:
    learner = Learner("L001", "Thabo Mokoena", " THABO@EXAMPLE.ORG ")
    assert learner.email == "thabo@example.org"


@pytest.mark.parametrize("email", ["", "invalid", "a@", "@example.org"])
def test_learner_rejects_invalid_email(email: str) -> None:
    with pytest.raises(ValidationError):
        Learner("L001", "Thabo", email)


@pytest.mark.parametrize("capacity", [0, -1, False])
def test_course_rejects_invalid_capacity(capacity: int) -> None:
    with pytest.raises(ValidationError):
        Course("PY701", "Enterprise Python", capacity)


def test_course_tracks_assessment_relationship() -> None:
    course = Course("PY701", "Enterprise Python", 10)
    assessment = Assessment("A001", course.course_id, "Practical")
    course.add_assessment(assessment.assessment_id)
    assert assessment.assessment_id in course.assessment_ids


@pytest.mark.parametrize("score", [-0.1, 100.1])
def test_assessment_rejects_out_of_range_score(score: float) -> None:
    assessment = Assessment("A001", "PY701", "Practical")
    with pytest.raises(ValidationError):
        assessment.validate_score(score)


def test_support_ticket_status_workflow() -> None:
    ticket = SupportTicket(
        "L001",
        "Cannot sign in",
        "The portal rejects my password",
        TicketCategory.TECHNICAL,
        TicketPriority.HIGH,
    )
    ticket.start_work()
    ticket.resolve()
    assert ticket.status.value == "resolved"
    with pytest.raises(ValidationError):
        ticket.start_work()

