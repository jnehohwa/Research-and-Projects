import pytest

from educore.domain import TicketCategory, ValidationError
from educore.patterns import (
    ApplicationConfig,
    BestAttemptStrategy,
    PassFailStrategy,
    SupportTicketFactory,
    WeightedAverageStrategy,
)


def test_application_config_is_singleton() -> None:
    assert ApplicationConfig() is ApplicationConfig()


@pytest.mark.parametrize(
    ("category", "priority"),
    [
        (TicketCategory.ACADEMIC, "medium"),
        (TicketCategory.TECHNICAL, "high"),
        (TicketCategory.BILLING, "high"),
        (TicketCategory.GENERAL, "low"),
    ],
)
def test_ticket_factory_applies_priority(category: TicketCategory, priority: str) -> None:
    ticket = SupportTicketFactory.create(category, "L001", "Help", "Details")
    assert ticket.category is category
    assert ticket.priority.value == priority


def test_ticket_factory_rejects_unknown_category() -> None:
    with pytest.raises(ValidationError):
        SupportTicketFactory.create("unknown", "L001", "Help", "Details")


def test_weighted_average_strategy() -> None:
    assert WeightedAverageStrategy((0.2, 0.3, 0.5)).calculate([68, 76, 81]) == 76.9


def test_best_attempt_strategy() -> None:
    assert BestAttemptStrategy().calculate([68, 76, 81]) == 81


def test_pass_fail_strategy_boundary() -> None:
    assert PassFailStrategy(50).calculate([40, 60]) is True


def test_strategy_rejects_invalid_scores() -> None:
    with pytest.raises(ValidationError):
        BestAttemptStrategy().calculate([101])
