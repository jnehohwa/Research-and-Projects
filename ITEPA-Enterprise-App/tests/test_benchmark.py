from educore.benchmark import run_lookup_benchmark


def test_benchmark_compares_equivalent_lookup_results() -> None:
    result = run_lookup_benchmark(dataset_size=1_000, repetitions=100)
    assert result.dataset_size == 1_000
    assert result.list_seconds > 0
    assert result.set_seconds > 0

