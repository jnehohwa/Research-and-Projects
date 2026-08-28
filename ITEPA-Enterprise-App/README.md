# EduCore Enterprise Application

EduCore is a Python prototype for managing learners, courses, registrations,
assessments, support tickets, and operational monitoring in a training provider.

## Requirements

- Python 3.11 or newer
- macOS, Linux, or Windows

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install '.[dev]'
python -m pytest --cov=educore --cov-report=term-missing
```

## Run the application demonstration

```bash
python -m educore.demo --report-directory evidence/reports
```

The demonstration exercises the domain relationships, Singleton, Factory and
Strategy patterns, 16 concurrent registrations, capacity and duplicate rules,
Bugzot reporting, and the optimisation benchmark.

## Regenerate evidence and report

```bash
python scripts/build_evidence.py
python scripts/build_visuals.py
python scripts/build_report.py
```

## Architecture

- `domain.py`: validated entities and state types
- `patterns.py`: Singleton, Factory and Strategy implementations
- `registration.py`: sequential and thread-safe concurrent registration service
- `bugzot.py`: structured events and performance reporting
- `benchmark.py`: reproducible list-versus-set comparison
- `demo.py`: complete assessment demonstration
- `tests/`: 32 automated tests covering core behaviour and concurrency

## Submission approach

The implementation is developed incrementally against the ITEPA3-33 marking
criteria. Evidence, profiling output, diagrams, and report material will be kept
in dedicated folders so each claim can be reproduced.

The `submission/` folder and final ZIP are generated only after the verification
commands pass. Review `SUBMISSION_CHECKLIST.md` before uploading.
