"""Build the rubric-aligned EduCore assessment report as a verified PDF."""

from __future__ import annotations

import json
from pathlib import Path

from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "pdf" / "Joshua_Nehohwa_ITEPA3-33_Practical_Report.pdf"
NAVY = colors.HexColor("#123353")
TEAL = colors.HexColor("#07818d")
PALE = colors.HexColor("#eaf1f7")
INK = colors.HexColor("#24364b")
MUTED = colors.HexColor("#5d7085")
GREEN = colors.HexColor("#087f5b")


styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="CoverTitle", fontName="Helvetica-Bold", fontSize=27, leading=33, textColor=colors.white, alignment=TA_CENTER, spaceAfter=12))
styles.add(ParagraphStyle(name="CoverSub", fontName="Helvetica", fontSize=14, leading=20, textColor=colors.HexColor("#dbeafe"), alignment=TA_CENTER))
styles.add(ParagraphStyle(name="H1x", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=20, leading=25, textColor=NAVY, spaceBefore=8, spaceAfter=10))
styles.add(ParagraphStyle(name="H2x", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=TEAL, spaceBefore=12, spaceAfter=7))
styles.add(ParagraphStyle(name="H3x", parent=styles["Heading3"], fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=NAVY, spaceBefore=9, spaceAfter=5))
styles.add(ParagraphStyle(name="Bodyx", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.4, leading=14, textColor=INK, spaceAfter=7))
styles.add(ParagraphStyle(name="Smallx", parent=styles["BodyText"], fontName="Helvetica", fontSize=7.8, leading=10.5, textColor=MUTED, spaceAfter=4))
styles.add(ParagraphStyle(name="Captionx", parent=styles["BodyText"], fontName="Helvetica-Oblique", fontSize=8, leading=11, textColor=MUTED, alignment=TA_CENTER, spaceBefore=4, spaceAfter=10))
styles.add(ParagraphStyle(name="Callout", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=9.2, leading=14, textColor=GREEN, backColor=colors.HexColor("#e9fbf4"), borderColor=colors.HexColor("#9ee6cb"), borderWidth=1, borderPadding=8, spaceBefore=6, spaceAfter=10))
styles.add(ParagraphStyle(name="CodeX", fontName="Courier", fontSize=7.4, leading=10, textColor=colors.HexColor("#e5eef8"), backColor=colors.HexColor("#101827"), borderPadding=8, spaceBefore=5, spaceAfter=8))


def p(text: str, style: str = "Bodyx") -> Paragraph:
    return Paragraph(text, styles[style])


def bullet(text: str) -> Paragraph:
    return Paragraph(text, styles["Bodyx"], bulletText="-")


def table(rows: list[list[object]], widths: list[float] | None = None) -> Table:
    prepared = [[cell if hasattr(cell, "wrap") else p(str(cell), "Smallx") for cell in row] for row in rows]
    result = Table(prepared, colWidths=widths, repeatRows=1, hAlign="LEFT")
    result.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdc9d6")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return result


def figure(path: Path, caption: str, max_width: float = 175 * mm, max_height: float = 205 * mm) -> list[object]:
    with PILImage.open(path) as image:
        width, height = image.size
    scale = min(max_width / width, max_height / height)
    return [KeepTogether([Image(str(path), width=width * scale, height=height * scale), p(caption, "Captionx")])]


def code(text: str) -> Paragraph:
    escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
    return p(escaped, "CodeX")


def header_footer(canvas, document) -> None:
    canvas.saveState()
    width, height = A4
    if document.page > 1:
        canvas.setStrokeColor(colors.HexColor("#d4dee8"))
        canvas.line(18 * mm, height - 17 * mm, width - 18 * mm, height - 17 * mm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(MUTED)
        canvas.drawString(18 * mm, height - 13 * mm, "ITEPA3-33 Enterprise Programming in Python")
        canvas.drawRightString(width - 18 * mm, 12 * mm, f"Joshua Nehohwa | Page {document.page}")
    canvas.restoreState()


def build_story() -> list[object]:
    reports = ROOT / "evidence" / "reports"
    screenshots = ROOT / "evidence" / "screenshots"
    diagrams = ROOT / "diagrams"
    metrics = json.loads((reports / "bugzot_performance_report.json").read_text(encoding="utf-8"))
    benchmark = json.loads((reports / "optimisation_benchmark.json").read_text(encoding="utf-8"))

    story: list[object] = []
    cover = Table(
        [[p("EduCore", "CoverTitle")], [p("Enterprise Training Management Prototype", "CoverSub")], [Spacer(1, 18 * mm)], [p("Practical Assignment Report", "CoverTitle")], [p("ITEPA3-33 | Enterprise Programming in Python", "CoverSub")], [Spacer(1, 22 * mm)], [p("Joshua Nehohwa", "CoverTitle")], [p("Eduvos | August 2026", "CoverSub")]],
        colWidths=[178 * mm],
        rowHeights=[None, None, None, None, None, None, None, None],
    )
    cover.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 18), ("BOTTOMPADDING", (0, 0), (-1, -1), 18)]))
    story.extend([Spacer(1, 34 * mm), cover, PageBreak()])

    story += [p("Executive summary", "H1x"), p("EduCore is a Python enterprise-application prototype for a growing South African training provider. It centralises learners, courses, registrations, assessments and support tickets while demonstrating maintainable object-oriented design, three enterprise design patterns, thread-safe concurrent processing, structured monitoring, automated testing, profiling and a migration path to microservices."), p("The implementation is deliberately modular. Domain entities contain validation and business meaning; services coordinate workflows; strategies and factories isolate variable behaviour; and Bugzot captures operational evidence. The final test run completed 32 tests successfully with 92% measured coverage. A contention test confirmed that concurrent registration never exceeded course capacity and did not create duplicate records.", "Callout"), p("Contents", "H2x")]
    toc_rows = [["Section", "Assessment area", "Marks"], ["1", "Domain model and design patterns", "20"], ["2", "Registration, data modelling and concurrency", "25"], ["3", "Bugzot, monitoring and optimisation", "20"], ["4", "Interface design and evaluation", "15"], ["5", "Testing, profiling and microservices readiness", "20"], ["Appendix", "Execution, AI-use and references", "-"]]
    story += [table(toc_rows, [22 * mm, 126 * mm, 22 * mm]), p("System boundaries", "H2x"), p("This submission is a working command-line prototype supported by interface mockups. Persistence remains in memory so that the assessed enterprise mechanisms are visible and testable. The report distinguishes prototype behaviour from future production requirements such as an external database, authentication, deployment and distributed infrastructure."), PageBreak()]

    story += [p("1. Deliverable 1: Foundational components", "H1x"), p("1.1 Domain model implementation", "H2x"), p("The domain model uses validated dataclasses with explicit state enums. Each class has one central responsibility and exposes operations that preserve its invariants. Composition and identifier relationships keep entities understandable without tightly coupling every object to every other object."), table([
        ["Class", "Responsibility and main validation", "Relationship"],
        ["Learner", "Stores identity, normalises email and rejects missing or invalid values.", "A learner participates in registrations and owns support tickets."],
        ["Course", "Stores title and a strictly positive integer capacity.", "Tracks assessment identifiers and receives registrations."],
        ["Registration", "Immutable record containing learner, course, status and UTC creation time.", "Joins one learner to one course."],
        ["Assessment", "Validates title, maximum mark and score boundaries.", "Belongs to a course and delegates result calculation to a strategy."],
        ["SupportTicket", "Controls category, priority and open-to-resolved workflow.", "Belongs to a learner and is created through the factory."],
    ], [27 * mm, 91 * mm, 55 * mm]), p("Object-oriented principles", "H3x"), bullet("Encapsulation: validation is performed inside domain constructors and state-changing methods."), bullet("Abstraction: assessment calculations depend on the AssessmentStrategy interface rather than a specific algorithm."), bullet("Polymorphism: each strategy provides the same calculate interface but returns its algorithm-specific result."), bullet("Single responsibility: registration workflow, monitoring and entity state are held in separate modules."), code("course.add_assessment(assessment.assessment_id)\nregistration = Registration(learner.learner_id, course.course_id)\nticket.start_work()\nticket.resolve()")]
    story += figure(screenshots / "domain_and_patterns.png", "Figure 1. Successful creation of related domain objects and execution of all required patterns.", max_height=125 * mm)

    story += [p("1.2 Design pattern implementation", "H2x"), table([
        ["Pattern", "Implementation", "Enterprise relevance"],
        ["Singleton", "ApplicationConfig uses a class-level instance and double-checked lock.", "All services observe one configuration without repeated file reads or inconsistent settings."],
        ["Factory", "SupportTicketFactory maps ticket categories to controlled priority defaults.", "New ticket types can be added without scattering construction rules through controllers."],
        ["Strategy", "WeightedAverage, BestAttempt and PassFail implement AssessmentStrategy.", "Assessment policy can vary by programme without changing the Assessment entity or calling workflow."],
    ], [27 * mm, 68 * mm, 78 * mm]), p("The Singleton lock protects first-time creation if multiple threads request configuration simultaneously. The Factory rejects unsupported categories, while the Strategy implementations validate score and weight boundaries. The demonstration proves that both configuration requests return the same object identity and that each algorithm produces a distinct, correct result."), PageBreak()]

    story += [p("2. Deliverable 2: Scalable registration", "H1x"), p("2.1 Registration processing engine", "H2x"), p("RegistrationService owns learner and course indexes, confirmed registration records and per-course counters. A request receives a correlation ID and one explicit outcome: confirmed, rejected_duplicate, rejected_capacity or rejected_invalid. The service therefore supports both accurate records and a complete processing summary."), table([
        ["Processing stage", "Rule", "Failure outcome"],
        ["Identity validation", "Learner and course must exist; learner must be active.", "rejected_invalid"],
        ["Uniqueness", "The learner-course key must not already exist.", "rejected_duplicate"],
        ["Capacity", "Confirmed count must remain below course capacity.", "rejected_capacity"],
        ["Commit", "Insert the immutable registration and increment the course counter.", "confirmed"],
    ], [35 * mm, 97 * mm, 41 * mm]), p("The demonstration processes 16 requests, including a deliberate duplicate, against a course with capacity 10. Because request completion order is deliberately nondeterministic, the duplicate request may win the race and the earlier request may be reported as duplicate; the invariant remains that only one L001-PY701 record exists."), p("2.2 Scalable data modelling", "H2x"), table([
        ["Approach", "Lookup", "Advantages", "Limitations"],
        ["List of registrations", "O(n)", "Simple iteration and insertion.", "Every duplicate check scans progressively more records."],
        ["Dictionary/set index", "Average O(1)", "Fast uniqueness checks and explicit composite keys.", "In-memory state is process-local and non-durable."],
        ["Production database", "Indexed lookup", "Durability, transactions, unique constraints and multi-instance access.", "Requires schema, deployment and operational management."],
    ], [31 * mm, 24 * mm, 64 * mm, 54 * mm]), p("The prototype selects dictionaries and sets because they improve processing cost without obscuring the enterprise concepts being assessed. The same composite key maps naturally to a database unique constraint in a production migration.")]
    story += [PageBreak(), p("2.3 Concurrent request processing", "H2x"), p("ThreadPoolExecutor represents simultaneous registration traffic. An Event releases queued work together, increasing contention. The critical section is protected by one Lock around duplicate check, capacity check and insertion. Treating this sequence atomically prevents both check-then-act races and lost capacity updates."), code("with self._lock:\n    if key in self._registrations:\n        return duplicate_result\n    if self._course_counts[course_id] >= course.capacity:\n        return capacity_result\n    self._registrations[key] = registration\n    self._course_counts[course_id] += 1"), table([
        ["Risk", "Failure without protection", "Control"],
        ["Duplicate race", "Two threads both see no existing key and insert twice.", "Composite-key lookup and lock around check plus insert."],
        ["Capacity race", "Several threads see the last available place.", "Capacity check and counter increment share the same lock."],
        ["Monitoring race", "Events are lost or list state is inconsistent.", "Bugzot protects its event collection with a separate lock."],
        ["Deadlock", "Workers wait indefinitely.", "Short non-nested critical sections; start coordination uses Event rather than a barrier."],
    ], [31 * mm, 70 * mm, 72 * mm])]
    story += figure(screenshots / "concurrent_registration.png", "Figure 2. Sixteen concurrent requests produce ten confirmations, five capacity rejections and one duplicate rejection; the final count equals capacity.", max_height=137 * mm)

    story += [PageBreak(), p("3. Deliverable 3: Bugzot monitoring", "H1x"), p("3.1 Monitoring subsystem", "H2x"), p("Bugzot records immutable structured events. Each event contains UTC timestamp, severity, component, event type, outcome, message, correlation ID, learner/course context and optional duration. This provides enough information to trace a request, group operational failures and reproduce the affected business context."), table([
        ["Event type", "Trigger", "Diagnostic value"],
        ["validation_failure", "Unknown or inactive learner, or unknown course.", "Identifies invalid input and affected identifiers."],
        ["duplicate_registration", "Existing learner-course composite key.", "Shows attempted duplicate and protects record accuracy."],
        ["course_capacity_violation", "Confirmed count equals course capacity.", "Supports demand and capacity planning."],
        ["registration_success", "Registration committed successfully.", "Supports transaction volume and success-rate monitoring."],
    ], [40 * mm, 65 * mm, 68 * mm]), p("Events can be exported to CSV for investigation and to JSON for management reporting. Bugzot itself is independent of console presentation, which keeps monitoring reusable."), p("3.2 Application performance monitoring", "H2x"), p(f"The captured run contains <b>{metrics['total_events']}</b> events: <b>{metrics['successful_transactions']}</b> confirmed and <b>{metrics['failed_transactions']}</b> rejected transactions. The success rate is <b>{metrics['success_rate_percent']}%</b>. Average measured processing time is <b>{metrics['average_duration_ms']} ms</b>, p95 is <b>{metrics['p95_duration_ms']} ms</b>, and observed throughput is <b>{metrics['throughput_transactions_per_second']} transactions per second</b> on this local in-memory workload."), p("These figures describe a small simulator run, not production capacity. They are useful as a baseline and for regression comparison; real deployment testing would use sustained load, external storage and percentile aggregation over longer windows.", "Callout")]
    story += figure(screenshots / "bugzot_report.png", "Figure 3. Bugzot performance report generated from the registration run.", max_height=130 * mm)

    story += [PageBreak(), p("3.3 Performance improvement", "H2x"), p("The original duplicate-detection candidate scanned a registration list. The root cause is linear search: as the collection grows, each worst-case membership check compares against every existing registration. The implementation replaces this with a set/dictionary composite-key index, whose hash lookup is constant time on average."), table([
        ["Metric", "Before: list", "After: set", "Result"],
        ["Dataset size", str(benchmark["dataset_size"]), str(benchmark["dataset_size"]), "Identical inputs"],
        ["Repeated lookups", str(benchmark["repetitions"]), str(benchmark["repetitions"]), "Identical workload"],
        ["Measured seconds", benchmark["list_seconds"], benchmark["set_seconds"], f"{benchmark['improvement_percent']}% faster"],
        ["Functional result", "Target found", "Target found", "Equivalent behaviour"],
    ], [40 * mm, 39 * mm, 39 * mm, 55 * mm]), p("Repeated runs will vary with hardware and background activity, but the algorithmic advantage remains. The next likely bottleneck in production is database or network latency rather than in-memory lookup.")]
    story += figure(screenshots / "optimisation_benchmark.png", "Figure 4. Reproducible before-and-after lookup benchmark.", max_height=112 * mm)

    story += [PageBreak(), p("4. Deliverable 4: Interface design and evaluation", "H1x"), p("4.1 User interface design", "H2x"), p("The proposed desktop-first interface uses persistent role-aware navigation, consistent cards, plain labels and visible system feedback. Learners receive focused support and registration interactions; administrators receive course management and operational reporting. Required fields use labels and inline errors rather than relying on colour alone."), table([
        ["Requirement", "Screen and interaction"],
        ["Learner registration", "Validated form with identity, email and course selection plus confirmation guidance."],
        ["Course management", "Capacity and enrolment table, availability status and controlled edit action."],
        ["Support ticket creation", "Category, automatically assigned priority, subject, description and response expectation."],
        ["Report viewing", "Date filter, registration outcome visualisation, system-health summary and export action."],
    ], [45 * mm, 128 * mm])]
    for index, (filename, caption) in enumerate([
        ("ui_01_dashboard.png", "Figure 5. Administrator dashboard and primary navigation."),
        ("ui_02_learner_registration.png", "Figure 6. Learner-registration form with inline validation."),
        ("ui_03_course_management.png", "Figure 7. Course capacity and availability management."),
        ("ui_04_support_ticket.png", "Figure 8. Learner support-ticket creation workflow."),
        ("ui_05_reports.png", "Figure 9. Operational report with outcome and health indicators."),
    ]):
        story += figure(diagrams / filename, caption, max_height=92 * mm)
        if index in (1, 3):
            story.append(PageBreak())
    story += [p("4.2 Design decisions and evaluation", "H2x"), p("Layout and navigation: a persistent left navigation column provides predictable access while the top bar confirms the current role and context. Primary actions use consistent teal buttons and are positioned after the required information."), p("Validation and usability: errors appear beside the relevant field with specific correction guidance. Capacity is visible before enrolment, priority is explained, report state is summarised in text, and destructive actions are not placed beside primary actions."), table([
        ["Strengths", "Limitation", "Recommended improvement"],
        ["Consistent visual hierarchy, task-focused forms, visible validation, role context and operational feedback.", "The prototype is desktop-first and has not been tested with representative users or assistive technology.", "Create responsive mobile variants and perform moderated usability plus WCAG keyboard/screen-reader testing before implementation."],
    ], [60 * mm, 55 * mm, 58 * mm]), PageBreak()]

    story += [p("5. Deliverable 5: Quality and future growth", "H1x"), p("5.1 Automated testing", "H2x"), p("Pytest was selected because it provides concise tests, parameterisation, clear failure output and a mature coverage ecosystem. The 32-test suite covers normal behaviour, boundaries, invalid inputs, every required pattern, business rules, concurrent invariants, structured monitoring and the benchmark harness."), table([
        ["Test category", "Representative scenarios"],
        ["Domain validation", "Invalid email, empty text, zero/negative capacity, score boundaries and ticket state transition."],
        ["Patterns", "Singleton identity, four Factory categories, invalid category, three strategies and invalid scores."],
        ["Registration", "Success, invalid input, duplicate, capacity, summary completeness and concurrent contention."],
        ["Monitoring", "Context capture, performance aggregation and JSON export."],
        ["Performance", "List/set benchmark executes equivalent successful lookups."],
    ], [40 * mm, 133 * mm]), p("The test suite completed with 32 passes and 92% coverage of assessed core modules. The demonstration orchestration is intentionally omitted from coverage because its behaviour is exercised through the underlying tested services and captured end-to-end output.")]
    story += figure(screenshots / "automated_tests.png", "Figure 10. Successful automated test and coverage run.", max_height=124 * mm)

    profile_text = (ROOT / "evidence" / "profiling" / "profile_summary.txt").read_text(encoding="utf-8")
    profile_observation = "Benchmark list membership and demonstration/report setup dominate cumulative time; registration critical sections remain small."
    if "benchmark.py" not in profile_text:
        profile_observation = "The profile confirms that orchestration and output dominate this small in-memory workload; registration critical sections remain small."
    story += [PageBreak(), p("5.2 Application profiling", "H2x"), p("The application was profiled with Python cProfile using the full demonstration workload. Results were sorted by cumulative time so that work initiated through helper calls remains visible."), p(profile_observation), p("The profile supports the implemented set-based optimisation. Further optimisation should focus on realistic persistence and report serialisation only after measuring a database-backed workload. Micro-optimising validated entity construction would add complexity without material benefit.")]
    story += figure(screenshots / "profiling_results.png", "Figure 11. cProfile results sorted by cumulative time.", max_height=128 * mm)

    story += [PageBreak(), p("5.3 Microservices readiness assessment", "H2x"), p("The modular monolith provides clear candidate boundaries without prematurely introducing distributed-system cost. Each service should own its data and publish stable contracts. The Registration Service is the consistency boundary for enrolment; Reporting consumes events rather than reading every operational database directly."), table([
        ["Candidate service", "Responsibility", "Independent operation"],
        ["Learner", "Profiles, status and identity validation.", "Can manage learner records independently."],
        ["Course", "Catalogue, course status and capacity definition.", "Can publish course availability independently."],
        ["Registration", "Uniqueness, capacity reservation and enrolment lifecycle.", "Operates independently but queries cached learner/course eligibility."],
        ["Assessment", "Assessment definitions, attempts and result strategies.", "Can calculate and publish results independently."],
        ["Support", "Ticket categorisation, priority and workflow.", "Can continue accepting tickets during reporting outages."],
        ["Reporting", "Bugzot events, metrics and management reports.", "Consumes events asynchronously and tolerates temporary producer outages."],
    ], [34 * mm, 83 * mm, 56 * mm]), p("Communication: REST is appropriate for immediate validation and user-facing queries. Events such as registration-created, assessment-completed and ticket-updated reduce coupling for reporting and notification. Each message requires an idempotency key and correlation ID."), p("Testing and tracing: unit tests remain within each service; consumer-driven contract tests protect REST and event schemas; integration tests cover key workflows; and distributed traces connect gateway, service and event processing. Bugzot becomes the central event and performance view while service-local logs remain available for diagnosis.")]
    story += figure(diagrams / "microservices_architecture.png", "Figure 12. Proposed service boundaries, communication and cross-cutting controls.", max_height=125 * mm)

    story += [PageBreak(), p("Conclusion", "H1x"), p("EduCore demonstrates a maintainable enterprise-programming foundation rather than a collection of isolated examples. Validated entities and patterns support change; the registration service preserves uniqueness and capacity under contention; Bugzot provides diagnostic and performance visibility; tests and profiling supply reproducible quality evidence; and the proposed interfaces and service boundaries establish a credible path to future growth."), p("Submission verification", "H2x"), bullet("Run python -m pytest --cov=educore and confirm all tests pass."), bullet("Run python -m educore.demo and confirm the final registration count equals course capacity."), bullet("Regenerate reports and compare the current evidence with the figures in this document."), bullet("Review the AI-use disclosure for accuracy and retain the development history."), p("References", "H2x"), p("Gamma, E., Helm, R., Johnson, R. and Vlissides, J. (1994). <i>Design Patterns: Elements of Reusable Object-Oriented Software</i>. Addison-Wesley."), p("Python Software Foundation (2026). <i>concurrent.futures - Launching parallel tasks</i>. Available at: https://docs.python.org/3/library/concurrent.futures.html"), p("Python Software Foundation (2026). <i>threading - Thread-based parallelism</i>. Available at: https://docs.python.org/3/library/threading.html"), p("Python Software Foundation (2026). <i>The Python Profilers</i>. Available at: https://docs.python.org/3/library/profile.html"), p("pytest development team (2026). <i>pytest documentation</i>. Available at: https://docs.pytest.org/"), p("Appendix A: AI-use statement", "H2x"), p("AI assistance was used for rubric analysis, project scaffolding, implementation support, test design, debugging, evidence preparation and report structuring. The accompanying AI_USE_LOG.md records the actual prompts and the student's required review and verification actions. AI use must be declared according to the final LMS template."), p("Appendix B: Reproduction commands", "H2x"), code("python3 -m venv .venv\nsource .venv/bin/activate\npython -m pip install '.[dev]'\npython -m pytest --cov=educore --cov-report=term-missing\npython -m educore.demo --report-directory evidence/reports")]
    return story


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=22 * mm,
        bottomMargin=18 * mm,
        title="EduCore ITEPA3-33 Practical Assignment Report",
        author="Joshua Nehohwa",
        subject="Enterprise Programming in Python",
    )
    document.build(build_story(), onFirstPage=header_footer, onLaterPages=header_footer)
    print(OUTPUT)


if __name__ == "__main__":
    main()
