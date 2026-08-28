"""Validated domain entities for the EduCore training platform."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from enum import StrEnum
import re
from uuid import uuid4


class ValidationError(ValueError):
    """Raised when domain data violates a business invariant."""


class RegistrationStatus(StrEnum):
    CONFIRMED = "confirmed"
    REJECTED_DUPLICATE = "rejected_duplicate"
    REJECTED_CAPACITY = "rejected_capacity"
    REJECTED_INVALID = "rejected_invalid"


class TicketCategory(StrEnum):
    ACADEMIC = "academic"
    TECHNICAL = "technical"
    BILLING = "billing"
    GENERAL = "general"


class TicketPriority(StrEnum):
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


class TicketStatus(StrEnum):
    OPEN = "open"
    IN_PROGRESS = "in_progress"
    RESOLVED = "resolved"


_EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _require_text(value: str, field_name: str) -> str:
    cleaned = value.strip()
    if not cleaned:
        raise ValidationError(f"{field_name} is required")
    return cleaned


@dataclass(slots=True)
class Learner:
    learner_id: str
    name: str
    email: str
    active: bool = True

    def __post_init__(self) -> None:
        self.learner_id = _require_text(self.learner_id, "learner_id")
        self.name = _require_text(self.name, "name")
        self.email = self.email.strip().lower()
        if not _EMAIL_PATTERN.fullmatch(self.email):
            raise ValidationError("email must be a valid email address")


@dataclass(slots=True)
class Course:
    course_id: str
    title: str
    capacity: int
    assessment_ids: set[str] = field(default_factory=set)

    def __post_init__(self) -> None:
        self.course_id = _require_text(self.course_id, "course_id")
        self.title = _require_text(self.title, "title")
        if isinstance(self.capacity, bool) or not isinstance(self.capacity, int) or self.capacity <= 0:
            raise ValidationError("capacity must be a positive integer")

    def add_assessment(self, assessment_id: str) -> None:
        self.assessment_ids.add(_require_text(assessment_id, "assessment_id"))


@dataclass(frozen=True, slots=True)
class Registration:
    learner_id: str
    course_id: str
    status: RegistrationStatus = RegistrationStatus.CONFIRMED
    registration_id: str = field(default_factory=lambda: f"REG-{uuid4().hex[:10].upper()}")
    created_at: datetime = field(default_factory=lambda: datetime.now(UTC))


@dataclass(slots=True)
class Assessment:
    assessment_id: str
    course_id: str
    title: str
    maximum_mark: float = 100.0

    def __post_init__(self) -> None:
        self.assessment_id = _require_text(self.assessment_id, "assessment_id")
        self.course_id = _require_text(self.course_id, "course_id")
        self.title = _require_text(self.title, "title")
        if self.maximum_mark <= 0:
            raise ValidationError("maximum_mark must be greater than zero")

    def validate_score(self, score: float) -> float:
        if not 0 <= score <= self.maximum_mark:
            raise ValidationError(f"score must be between 0 and {self.maximum_mark:g}")
        return float(score)


@dataclass(slots=True)
class SupportTicket:
    learner_id: str
    subject: str
    description: str
    category: TicketCategory
    priority: TicketPriority
    ticket_id: str = field(default_factory=lambda: f"TKT-{uuid4().hex[:10].upper()}")
    status: TicketStatus = TicketStatus.OPEN
    created_at: datetime = field(default_factory=lambda: datetime.now(UTC))

    def __post_init__(self) -> None:
        self.learner_id = _require_text(self.learner_id, "learner_id")
        self.subject = _require_text(self.subject, "subject")
        self.description = _require_text(self.description, "description")

    def start_work(self) -> None:
        if self.status is TicketStatus.RESOLVED:
            raise ValidationError("a resolved ticket cannot be reopened by this operation")
        self.status = TicketStatus.IN_PROGRESS

    def resolve(self) -> None:
        self.status = TicketStatus.RESOLVED
