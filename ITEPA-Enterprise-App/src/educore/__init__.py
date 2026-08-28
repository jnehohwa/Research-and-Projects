"""EduCore enterprise training management prototype."""

from .bugzot import Bugzot, BugzotEvent
from .domain import Assessment, Course, Learner, Registration, SupportTicket
from .registration import RegistrationRequest, RegistrationResult, RegistrationService

__all__ = [
    "Assessment",
    "Bugzot",
    "BugzotEvent",
    "Course",
    "Learner",
    "Registration",
    "RegistrationRequest",
    "RegistrationResult",
    "RegistrationService",
    "SupportTicket",
]
