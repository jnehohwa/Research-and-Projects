"""Required enterprise design-pattern implementations."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from threading import Lock
from typing import ClassVar

from .domain import SupportTicket, TicketCategory, TicketPriority, ValidationError


class ApplicationConfig:
    """Thread-safe Singleton containing process-wide configuration."""

    _instance: ClassVar[ApplicationConfig | None] = None
    _instance_lock: ClassVar[Lock] = Lock()

    def __new__(cls) -> ApplicationConfig:
        if cls._instance is None:
            with cls._instance_lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance.worker_count = 8
                    cls._instance.log_level = "INFO"
                    cls._instance.report_directory = Path("evidence/reports")
        return cls._instance


class SupportTicketFactory:
    """Factory that assigns enterprise defaults to each ticket category."""

    _PRIORITIES = {
        TicketCategory.ACADEMIC: TicketPriority.MEDIUM,
        TicketCategory.TECHNICAL: TicketPriority.HIGH,
        TicketCategory.BILLING: TicketPriority.HIGH,
        TicketCategory.GENERAL: TicketPriority.LOW,
    }

    @classmethod
    def create(
        cls,
        category: TicketCategory | str,
        learner_id: str,
        subject: str,
        description: str,
    ) -> SupportTicket:
        try:
            resolved_category = TicketCategory(category)
        except ValueError as error:
            raise ValidationError(f"unsupported ticket category: {category}") from error
        return SupportTicket(
            learner_id=learner_id,
            subject=subject,
            description=description,
            category=resolved_category,
            priority=cls._PRIORITIES[resolved_category],
        )


class AssessmentStrategy(ABC):
    """Strategy interface for interchangeable assessment calculations."""

    @abstractmethod
    def calculate(self, scores: list[float]) -> float | bool:
        """Calculate a result from validated percentage scores."""

    @staticmethod
    def _validate_scores(scores: list[float]) -> None:
        if not scores:
            raise ValidationError("at least one score is required")
        if any(score < 0 or score > 100 for score in scores):
            raise ValidationError("scores must be between 0 and 100")


@dataclass(frozen=True, slots=True)
class WeightedAverageStrategy(AssessmentStrategy):
    weights: tuple[float, ...]

    def calculate(self, scores: list[float]) -> float:
        self._validate_scores(scores)
        if len(scores) != len(self.weights):
            raise ValidationError("the number of scores must match the number of weights")
        if abs(sum(self.weights) - 1.0) > 1e-9 or any(weight < 0 for weight in self.weights):
            raise ValidationError("weights must be non-negative and total 1.0")
        return round(sum(score * weight for score, weight in zip(scores, self.weights)), 2)


class BestAttemptStrategy(AssessmentStrategy):
    def calculate(self, scores: list[float]) -> float:
        self._validate_scores(scores)
        return max(scores)


@dataclass(frozen=True, slots=True)
class PassFailStrategy(AssessmentStrategy):
    pass_mark: float = 50.0

    def calculate(self, scores: list[float]) -> bool:
        self._validate_scores(scores)
        if not 0 <= self.pass_mark <= 100:
            raise ValidationError("pass_mark must be between 0 and 100")
        return sum(scores) / len(scores) >= self.pass_mark

