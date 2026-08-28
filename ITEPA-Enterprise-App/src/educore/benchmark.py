"""Reproducible list-versus-set duplicate lookup benchmark."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from time import perf_counter


@dataclass(frozen=True, slots=True)
class BenchmarkResult:
    dataset_size: int
    repetitions: int
    list_seconds: float
    set_seconds: float
    improvement_percent: float


def run_lookup_benchmark(dataset_size: int = 10_000, repetitions: int = 2_000) -> BenchmarkResult:
    registrations = [(f"L{index:05d}", "PY701") for index in range(dataset_size)]
    registration_index = set(registrations)
    target = registrations[-1]

    started_at = perf_counter()
    for _ in range(repetitions):
        assert target in registrations
    list_seconds = perf_counter() - started_at

    started_at = perf_counter()
    for _ in range(repetitions):
        assert target in registration_index
    set_seconds = perf_counter() - started_at

    improvement = (list_seconds - set_seconds) / list_seconds * 100 if list_seconds else 0.0
    return BenchmarkResult(
        dataset_size=dataset_size,
        repetitions=repetitions,
        list_seconds=round(list_seconds, 6),
        set_seconds=round(set_seconds, 6),
        improvement_percent=round(improvement, 2),
    )


def benchmark_as_dict(**kwargs: int) -> dict[str, int | float]:
    return asdict(run_lookup_benchmark(**kwargs))

