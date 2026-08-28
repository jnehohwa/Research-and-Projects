from __future__ import annotations

import io
import math
from pathlib import Path

from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4, LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    CondPageBreak,
    BaseDocTemplate,
    Frame,
    Flowable,
    HRFlowable,
    KeepTogether,
    ListFlowable,
    ListItem,
    PageBreak,
    PageTemplate,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents


ROOT = Path("/Users/joshuanehohwa/Documents/Research and Projects")
WORK = ROOT / "tmp/itopa_final"
COVER_SOURCE = Path(
    "/Users/joshuanehohwa/Library/CloudStorage/OneDrive-Eduvos/Documents/"
    "Block 4 2025 Projects/Individual Assignment Coversheet V1.2 - Copy (2).pdf"
)
OUTPUT = ROOT / "ITOPA Practical/ITOPA3-33 - Assignment - Mowbray - EDUV4948467.pdf"
ANSWER_PDF = WORK / "answer-pages.pdf"
FILLED_COVER = WORK / "filled-cover.pdf"

NAVY = colors.HexColor("#111827")
BLUE = colors.HexColor("#1F4D78")
TABLE_BLUE = colors.HexColor("#2E74B5")
PALE_BLUE = colors.HexColor("#E6EEF7")
PALE_GREY = colors.HexColor("#F3F5F7")
MID_GREY = colors.HexColor("#667085")
GREEN = colors.HexColor("#287D4F")
RED = colors.HexColor("#B42318")
INK = colors.HexColor("#182230")


def register_fonts() -> tuple[str, str, str]:
    carlito_dir = Path(
        "/Users/joshuanehohwa/.cache/codex-runtimes/codex-primary-runtime/dependencies/"
        "native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype"
    )
    carlito_regular = carlito_dir / "Carlito-Regular.ttf"
    carlito_bold = carlito_dir / "Carlito-Bold.ttf"
    regular_candidates = [
        Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/Library/Fonts/Arial.ttf"),
    ]
    bold_candidates = [
        Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
        Path("/Library/Fonts/Arial Bold.ttf"),
    ]
    regular = next((p for p in regular_candidates if p.exists()), None)
    bold = next((p for p in bold_candidates if p.exists()), None)
    if regular and bold:
        pdfmetrics.registerFont(TTFont("ArialLocal", str(regular)))
        pdfmetrics.registerFont(TTFont("ArialLocal-Bold", str(bold)))
    if carlito_regular.exists() and carlito_bold.exists():
        pdfmetrics.registerFont(TTFont("CarlitoLocal", str(carlito_regular)))
        pdfmetrics.registerFont(TTFont("CarlitoLocal-Bold", str(carlito_bold)))
        return "CarlitoLocal", "CarlitoLocal-Bold", "ArialLocal" if regular else "Helvetica"
    return "ArialLocal" if regular else "Helvetica", "ArialLocal-Bold" if bold else "Helvetica-Bold", "ArialLocal" if regular else "Helvetica"


FONT, FONT_BOLD, COVER_FONT = register_fonts()


styles = getSampleStyleSheet()
styles.add(
    ParagraphStyle(
        name="AssignmentTitle",
        parent=styles["Title"],
        fontName=FONT_BOLD,
        fontSize=19.9,
        leading=25,
        textColor=NAVY,
        alignment=TA_CENTER,
        spaceAfter=4 * mm,
    )
)
styles.add(
    ParagraphStyle(
        name="Question",
        parent=styles["Heading1"],
        fontName=FONT_BOLD,
        fontSize=18,
        leading=22,
        textColor=BLUE,
        alignment=TA_CENTER,
        spaceBefore=5 * mm,
        spaceAfter=3.5 * mm,
        keepWithNext=True,
    )
)
styles.add(
    ParagraphStyle(
        name="SubQuestion",
        parent=styles["Heading2"],
        fontName=FONT_BOLD,
        fontSize=15,
        leading=19,
        textColor=BLUE,
        alignment=TA_CENTER,
        spaceBefore=4.5 * mm,
        spaceAfter=2.5 * mm,
        keepWithNext=True,
    )
)
styles.add(
    ParagraphStyle(
        name="BodyA",
        parent=styles["BodyText"],
        fontName=FONT,
        fontSize=12,
        leading=18,
        textColor=INK,
        alignment=TA_JUSTIFY,
        firstLineIndent=18,
        spaceAfter=3.2 * mm,
        allowWidows=0,
        allowOrphans=0,
    )
)
styles.add(
    ParagraphStyle(
        name="BodySmall",
        parent=styles["BodyText"],
        fontName=FONT,
        fontSize=10,
        leading=13.5,
        textColor=INK,
        spaceAfter=1.7 * mm,
    )
)
styles.add(
    ParagraphStyle(
        name="BodySmallWhite",
        parent=styles["BodySmall"],
        fontName=FONT_BOLD,
        textColor=INK,
    )
)
styles.add(
    ParagraphStyle(
        name="BulletA",
        parent=styles["BodyText"],
        fontName=FONT,
        fontSize=12,
        leading=18,
        leftIndent=7 * mm,
        firstLineIndent=0,
        bulletIndent=0,
        spaceAfter=1.6 * mm,
        textColor=INK,
    )
)
styles.add(
    ParagraphStyle(
        name="Callout",
        parent=styles["BodyText"],
        fontName=FONT,
        fontSize=11.5,
        leading=16,
        textColor=INK,
        borderColor=BLUE,
        borderWidth=0.8,
        borderPadding=7,
        backColor=PALE_BLUE,
        spaceBefore=2 * mm,
        spaceAfter=4 * mm,
    )
)
styles.add(
    ParagraphStyle(
        name="CodeA",
        parent=styles["Code"],
        fontName="Courier",
        fontSize=11,
        leading=15,
        textColor=INK,
        backColor=PALE_GREY,
        borderPadding=6,
        spaceBefore=1.5 * mm,
        spaceAfter=3 * mm,
    )
)
styles.add(
    ParagraphStyle(
        name="Reference",
        parent=styles["BodyText"],
        fontName=FONT,
        fontSize=10.5,
        leading=14.5,
        leftIndent=5 * mm,
        firstLineIndent=-5 * mm,
        spaceAfter=2.2 * mm,
        textColor=INK,
    )
)
styles.add(
    ParagraphStyle(
        name="TocHeading",
        parent=styles["Heading1"],
        fontName=FONT_BOLD,
        fontSize=16.1,
        leading=20,
        textColor=BLUE,
        spaceAfter=5 * mm,
    )
)


class AssignmentDocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        frame = Frame(
            self.leftMargin,
            self.bottomMargin,
            self.width,
            self.height,
            id="assignment-body",
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )
        self.addPageTemplates(
            PageTemplate(id="assignment-pages", frames=[frame], onPage=add_header_footer)
        )

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and flowable.style.name == "Question":
            text = flowable.getPlainText()
            self.notify("TOCEntry", (0, text, self.page + 1))


def P(text: str, style: str = "BodyA") -> Paragraph:
    return Paragraph(text, styles[style])


def bullet_items(items: list[str]) -> ListFlowable:
    return ListFlowable(
        [ListItem(P(item, "BulletA"), leftIndent=4 * mm) for item in items],
        bulletType="bullet",
        start="circle",
        leftIndent=7 * mm,
        bulletFontName=FONT,
        bulletFontSize=7,
        bulletColor=BLUE,
        spaceAfter=2 * mm,
    )


class ResourceAllocationGraph(Flowable):
    def __init__(self, width: float = 528, height: float = 270):
        super().__init__()
        self.width = width
        self.height = height

    @staticmethod
    def _arrow(c, x1, y1, x2, y2, color, dashed=False):
        c.saveState()
        c.setStrokeColor(color)
        c.setFillColor(color)
        c.setLineWidth(1.4)
        if dashed:
            c.setDash(4, 3)
        c.line(x1, y1, x2, y2)
        angle = math.atan2(y2 - y1, x2 - x1)
        size = 6
        left = angle + math.pi * 0.82
        right = angle - math.pi * 0.82
        path = c.beginPath()
        path.moveTo(x2, y2)
        path.lineTo(x2 + size * math.cos(left), y2 + size * math.sin(left))
        path.lineTo(x2 + size * math.cos(right), y2 + size * math.sin(right))
        path.close()
        c.drawPath(path, fill=1, stroke=0)
        c.restoreState()

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#A9B8CB"))
        c.roundRect(0, 0, self.width, self.height, 7, fill=1, stroke=1)
        c.setFont(FONT_BOLD, 10)
        c.setFillColor(NAVY)
        c.drawString(12, self.height - 18, "Resource Allocation Graph at the blocking point")

        def process(x, y, label):
            c.setFillColor(PALE_BLUE)
            c.setStrokeColor(BLUE)
            c.circle(x, y, 19, fill=1, stroke=1)
            c.setFillColor(INK)
            c.setFont(FONT_BOLD, 9)
            c.drawCentredString(x, y - 3, label)

        def resource(x, y, label):
            c.setFillColor(PALE_GREY)
            c.setStrokeColor(MID_GREY)
            c.roundRect(x - 18, y - 15, 36, 30, 3, fill=1, stroke=1)
            c.setFillColor(INK)
            c.setFont(FONT_BOLD, 9)
            c.drawCentredString(x, y - 3, label)

        # Cycle 1 plus B: B -> P0 -> F -> P2 -> A -> P0.
        resource(28, 150, "B")
        process(100, 150, "P0")
        resource(185, 195, "F")
        process(270, 150, "P2")
        resource(185, 100, "A")
        self._arrow(c, 47, 150, 81, 150, BLUE)
        self._arrow(c, 117, 159, 166, 188, RED, dashed=True)
        self._arrow(c, 204, 188, 253, 159, BLUE)
        self._arrow(c, 253, 141, 204, 107, RED, dashed=True)
        self._arrow(c, 166, 107, 117, 141, BLUE)

        # Cycle 2: P1 -> D -> P3 -> C -> P1.
        process(335, 150, "P1")
        resource(420, 195, "D")
        process(505, 150, "P3")
        resource(420, 100, "C")
        self._arrow(c, 352, 159, 401, 188, RED, dashed=True)
        self._arrow(c, 439, 188, 488, 159, BLUE)
        self._arrow(c, 488, 141, 439, 107, RED, dashed=True)
        self._arrow(c, 401, 107, 352, 141, BLUE)

        # Resource E has two instances: one allocated to P1 and one free.
        c.setFillColor(PALE_GREY)
        c.setStrokeColor(MID_GREY)
        c.roundRect(312, 42, 46, 31, 3, fill=1, stroke=1)
        c.setFont(FONT_BOLD, 8.5)
        c.setFillColor(INK)
        c.drawCentredString(335, 61, "E (2)")
        c.setFillColor(BLUE)
        c.circle(328, 50, 3, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setStrokeColor(MID_GREY)
        c.circle(342, 50, 3, fill=1, stroke=1)
        self._arrow(c, 335, 74, 335, 130, BLUE)

        c.setFont(FONT, 7.7)
        c.setFillColor(RED)
        c.drawString(12, 25, "Dashed red: request edge (process -> resource)")
        c.setFillColor(BLUE)
        c.drawString(278, 25, "Blue: assignment edge (resource -> process)")
        c.setFillColor(MID_GREY)
        c.drawString(12, 10, "All resources A-F are shown. E has two instances: one allocated to P1 and one free.")
        c.restoreState()


def add_header_footer(c: canvas.Canvas, doc: SimpleDocTemplate):
    page_number = doc.page + 1  # The official coversheet becomes page 1 after merging.
    width, height = LETTER
    c.saveState()
    c.setFont(COVER_FONT, 9.1)
    c.setFillColor(colors.black)
    c.drawRightString(
        width - 36,
        27,
        f"ITOPA3-33_Assignment_Mowbray_eduv4948467 | Page {page_number}",
    )
    c.restoreState()


def metadata_table() -> Table:
    data = [
        [P("<b>Student</b>", "BodySmall"), P("Joshua Nehohwa", "BodySmall"), P("<b>Student number</b>", "BodySmall"), P("EDUV4948467", "BodySmall")],
        [P("<b>Campus</b>", "BodySmall"), P("Mowbray", "BodySmall"), P("<b>Lecturer</b>", "BodySmall"), P("Bongani Mahlangu", "BodySmall")],
        [P("<b>Module</b>", "BodySmall"), P("ITOPA3-33", "BodySmall"), P("<b>Assessment</b>", "BodySmall"), P("Individual Assignment", "BodySmall")],
    ]
    table = Table(data, colWidths=[27 * mm, 55 * mm, 34 * mm, 52 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), PALE_BLUE),
                ("BACKGROUND", (2, 0), (2, -1), PALE_BLUE),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4D4")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


def build_story():
    s = []
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle(
            name="TOCLevel0",
            fontName=FONT,
            fontSize=12,
            leading=26,
            leftIndent=0,
            firstLineIndent=0,
            textColor=INK,
        )
    ]
    s.append(HRFlowable(width="100%", thickness=1, color=TABLE_BLUE, spaceAfter=8 * mm))
    s.append(P("Table of Contents", "TocHeading"))
    s.append(toc)
    s.append(Spacer(1, 6 * mm))
    s.append(HRFlowable(width="100%", thickness=1, color=TABLE_BLUE))
    s.append(PageBreak())

    s.append(P("Operating Systems: System Calls, Multithreading, Deadlock and File Management", "AssignmentTitle"))
    s.append(HRFlowable(width="100%", thickness=1, color=TABLE_BLUE, spaceAfter=5 * mm))

    # Question 1
    s.append(P("Question 1 - Operating-system system calls (20 marks)", "Question"))
    s.append(P("1.1 Basic file-management system calls and their importance (8 marks)", "SubQuestion"))
    s.append(P(
        "When a customer requests a bank statement, the banking application should not communicate with the disk by itself. It should make controlled requests to the operating system through system calls. The main file-management calls and their relevance are summarised below. Eduvos explains that system calls let user programs request protected operating-system services, while file-system management covers file creation, deletion and access (Eduvos, 2026a, slides 41 and 43).",
    ))
    q1_calls = Table(
        [
            [P("System call", "BodySmallWhite"), P("Purpose", "BodySmallWhite"), P("Importance for the bank statement", "BodySmallWhite")],
            [P("create / delete", "BodySmall"), P("Create a new file or remove an existing one.", "BodySmall"), P("Used when generating a new statement or applying an authorised retention rule; the OS prevents arbitrary file changes.", "BodySmall")],
            [P("open / close", "BodySmall"), P("Open returns a protected descriptor or handle; close releases it and related kernel resources.", "BodySmall"), P("Gives the application controlled access to the correct statement and prevents leaked handles after viewing or printing.", "BodySmall")],
            [P("read / write", "BodySmall"), P("Transfer bytes from a file to memory or from memory to a file/device.", "BodySmall"), P("Read obtains the statement data; write supports an authorised export or sends output through the printer spooler.", "BodySmall")],
            [P("seek / reposition", "BodySmall"), P("Move the current file offset to a required position.", "BodySmall"), P("Lets the application retrieve a selected period or section without reading the complete file sequentially.", "BodySmall")],
            [P("get / set attributes", "BodySmall"), P("Read or change metadata such as size, dates, ownership and permissions.", "BodySmall"), P("Supports validation and audit information; setting sensitive attributes remains permission-controlled.", "BodySmall")],
        ],
        colWidths=[31 * mm, 57 * mm, 80 * mm],
        repeatRows=1,
    )
    q1_calls.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PALE_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), INK),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4D4")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    s.append(q1_calls)
    s.append(Spacer(1, 2.5 * mm))
    s.append(P(
        "Together, these calls enforce permissions, isolate one customer's process from another customer's data, coordinate simultaneous access, provide consistent error handling, support caching and buffering, and create an auditable access path. They also hide device-specific details, so the application can keep the same interface if the bank changes its disks or printers. Therefore, system calls turn a risky hardware operation into a controlled operating-system service (The Open Group, 2024).",
    ))

    s.append(P("1.2 Consequences of direct hard-drive and printer access (6 marks)", "SubQuestion"))
    s.append(P(
        "The developer's performance argument is understandable because avoiding a mode switch might appear to remove overhead. However, the small saving would introduce much larger risks. First, <b>security</b> would be weakened. Direct access could bypass normal permission checks, process isolation and protected file handles. A defect or malicious input could expose statements, transaction records or audit logs belonging to other customers. Modern operating systems deliberately separate user mode from kernel mode so that ordinary applications cannot freely control hardware (Microsoft, 2025a).",
    ))
    s.append(P(
        "Second, <b>data integrity and reliability</b> would be threatened. Raw disk access can ignore file-system metadata, journaling, locks and cache-coherency rules. One incorrect block address could corrupt a statement, a transaction database or an entire volume. Direct printer control can also bypass the spooler, causing mixed jobs, lost pages or a blocked device. Third, <b>concurrency</b> becomes unsafe because thousands of banking sessions may issue conflicting operations without the OS scheduler, queues and locking mechanisms. Fourth, <b>maintenance and portability</b> become worse because the application would need device-specific driver logic for every disk controller and printer model. Finally, real performance may decline because the design gives up optimised OS features such as caching, buffering, direct memory access, I/O scheduling and printer spooling (Eduvos, 2026a; Microsoft, 2025b).",
    ))

    s.append(P("1.3.1 CIO decision (2 marks)", "SubQuestion"))
    s.append(P(
        "I would <b>reject the proposal</b>. A banking system should not trade its security boundary, auditability and data integrity for a speculative performance improvement.",
    ))
    s.append(P("1.3.2 Justification (4 marks)", "SubQuestion"))
    s.append(P(
        "System calls are the controlled gateway between an application running in user mode and privileged operating-system services. The kernel validates a request, checks permissions, coordinates shared resources, invokes trusted drivers and returns a defined result or error. This matches the course distinction between user mode and kernel mode (Eduvos, 2026a, slide 41). The design protects the application, hardware and other processes at the same time. It is especially important in a regulated bank, where confidentiality, correctness, recovery and audit trails are more valuable than a minor reduction in call overhead. If performance is genuinely poor, the bank should profile the bottleneck and use supported options such as asynchronous I/O, batching, caching, faster storage or an approved driver update instead of bypassing the OS.",
    ))

    # Question 2
    s.append(P("Question 2 - Multithreaded ShopFast server (25 marks)", "Question"))
    s.append(P("2.1a Why the single-threaded server performed poorly (3 marks)", "SubQuestion"))
    s.append(P(
        "The server handled requests serially, so one request had to finish before the next request could make progress. During Black Friday, requests arrived faster than that one thread could complete them. The waiting queue therefore grew, response time increased, requests timed out and customers abandoned their carts. The design also failed to use the quad-core machine properly because only one thread could run application work at a time.",
    ))
    s.append(P("2.1b Blocking and its effect on clients (3 marks)", "SubQuestion"))
    s.append(P(
        "Blocking happens when a thread cannot continue until an event completes. For example, a ShopFast request may wait for a product database query, payment gateway response, disk read or slow network client. In the single-threaded design, that blocked request stops the only server thread. A customer browsing products must then wait even though their request is unrelated to the blocked payment request. This reduces both responsiveness and throughput (Eduvos, 2026b).",
    ))

    s.append(KeepTogether([
        P("2.2a Three benefits of multithreading (3 marks)", "SubQuestion"),
        bullet_items([
            "<b>Responsiveness:</b> one blocked thread does not freeze the whole server process.",
            "<b>Economy and resource sharing:</b> threads share code, process data and resources, and are cheaper to manage than separate processes.",
            "<b>Scalability:</b> runnable threads can execute in parallel across the four CPU cores (Eduvos, 2026b, slide 46).",
        ]),
    ]))
    s.append(KeepTogether([
        P("2.2b Direct application to ShopFast (3 marks)", "SubQuestion"),
        bullet_items([
            "<b>Responsiveness:</b> while one customer waits for the payment gateway, other workers can continue serving product searches and cart updates, reducing timeouts and abandoned carts.",
            "<b>Economy and resource sharing:</b> request threads can reuse ShopFast's configuration, connection pools and product cache, leaving more memory and processing capacity for customer work.",
            "<b>Scalability:</b> product browsing, cart updates and payment requests can run on different cores at the same time, increasing peak throughput provided the database does not become the next bottleneck.",
        ]),
    ]))

    s.append(P("2.3a Task parallelism compared with data parallelism (3 marks)", "SubQuestion"))
    s.append(P(
        "<b>Task parallelism</b> assigns different tasks or functions to different processing units. The tasks may execute different code and work on different data. <b>Data parallelism</b> divides a dataset into parts and applies the same operation to each part at the same time. A useful example is calculating the same discount rule across four sections of a very large product list (Eduvos, 2026b, slide 47).",
    ))
    s.append(P("2.3b Classification of the ShopFast server (3 marks)", "SubQuestion"))
    s.append(P(
        "The redesigned server mainly exhibits <b>task parallelism</b>. First, each HTTP request is an independent task belonging to a particular customer. Second, requests can follow different code paths: one thread may browse products, another may update a cart and another may process a payment. These tasks can run concurrently across the four cores even though their operations and data are different. Some repeated handlers may resemble data parallelism, but the overall server design is best classified as task parallelism.",
    ))

    s.append(P("2.4a Thread pool (2 marks)", "SubQuestion"))
    s.append(P(
        "A thread pool is a managed set of reusable worker threads. The server creates or maintains workers independently of individual requests. When a worker finishes one request, it returns to the pool and can process another request.",
    ))
    s.append(P("2.4b Why pooling is preferred (2 marks)", "SubQuestion"))
    s.append(P(
        "Creating a new thread for every request adds creation and destruction overhead and can allow a traffic spike to create thousands of threads. That consumes memory, increases context switching and can crash the server. A bounded pool reuses threads and limits the amount of work running at once, giving ShopFast predictable resource use. This follows the course description of fixed reusable workers as a way to control overhead and resources (Eduvos, 2026b, slide 54). The best worker count must still be measured: CPU-heavy work may stay near the core count, while I/O-heavy work may benefit from more than four workers (Oracle, 2025).",
    ))
    s.append(P("2.4c Role of the task queue (3 marks)", "SubQuestion"))
    s.append(P(
        "The task queue sits between request acceptance and worker execution. Each accepted HTTP request becomes a task and waits in the queue until a worker is available. The next worker removes a task, processes it and returns to the pool. The queue absorbs short bursts and can enforce first-in-first-out order or priorities. If it is bounded, it also provides backpressure: when both workers and queue are full, the server can reject or shed load in a controlled way instead of exhausting memory. ShopFast should monitor queue length and waiting time because a permanently growing queue means the system is overloaded, not healthy.",
    ))

    # Question 3
    s.append(P("Question 3 - Deadlock in the manufacturing system (20 marks)", "Question"))
    s.append(P("3.1 Resource Allocation Graph and deadlock result (10 marks)", "SubQuestion"))
    s.append(P(
        "Executing the resource requests line-by-line from P0 to P3 produces the following state. P0 obtains A, P1 obtains C, P2 obtains F and P3 obtains D. On the next line, P0 obtains B and P1 obtains one instance of E, while P2 waits for A and P3 waits for C. On the third request, P0 waits for F and P1 waits for D. The second instance of E is still free, but none of the blocked processes needs it next.",
    ))
    trace_table = Table(
        [
            [P("Request round", "BodySmallWhite"), P("P0", "BodySmallWhite"), P("P1", "BodySmallWhite"), P("P2", "BodySmallWhite"), P("P3", "BodySmallWhite")],
            [P("1", "BodySmall"), P("gets A", "BodySmall"), P("gets C", "BodySmall"), P("gets F", "BodySmall"), P("gets D", "BodySmall")],
            [P("2", "BodySmall"), P("gets B", "BodySmall"), P("gets E1", "BodySmall"), P("waits for A", "BodySmall"), P("waits for C", "BodySmall")],
            [P("3", "BodySmall"), P("waits for F", "BodySmall"), P("waits for D", "BodySmall"), P("already blocked", "BodySmall"), P("already blocked", "BodySmall")],
        ],
        colWidths=[34 * mm, 33.5 * mm, 33.5 * mm, 33.5 * mm, 33.5 * mm],
        repeatRows=1,
    )
    trace_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PALE_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), INK),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4D4")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    s.append(trace_table)
    s.append(Spacer(1, 2.5 * mm))
    s.append(ResourceAllocationGraph())
    s.append(Spacer(1, 3 * mm))
    s.append(P(
        "The graph contains two cycles: <b>P0 -> F -> P2 -> A -> P0</b> and <b>P1 -> D -> P3 -> C -> P1</b>. In the first cycle, P0 holds A and waits for F, while P2 holds F and waits for A. In the second cycle, P1 holds C and waits for D, while P3 holds D and waits for C. Since A, C, D and F each have one instance, the cycles confirm deadlock (Eduvos, 2026c, slide 66). The unused second instance of E cannot break either cycle because no blocked process is waiting for E; P1 already holds E1 and is waiting for D.",
    ))

    s.append(P("3.2 Safe request order (5 marks)", "SubQuestion"))
    s.append(P(
        "I would impose one global resource order: <b>A &lt; B &lt; C &lt; D &lt; E &lt; F</b>. Every process must request only in ascending order and should release in reverse order. The modified requests are:",
    ))
    order_table = Table(
        [
            [P("Process", "BodySmallWhite"), P("Original order", "BodySmallWhite"), P("Safe order", "BodySmallWhite"), P("Reason for change", "BodySmallWhite")],
            [P("P0", "BodySmall"), P("A -> B -> F", "BodySmall"), P("A -> B -> F", "BodySmall"), P("Already ascending", "BodySmall")],
            [P("P1", "BodySmall"), P("C -> E -> D", "BodySmall"), P("C -> D -> E", "BodySmall"), P("D before E", "BodySmall")],
            [P("P2", "BodySmall"), P("F -> A", "BodySmall"), P("A -> F", "BodySmall"), P("A before F", "BodySmall")],
            [P("P3", "BodySmall"), P("D -> C", "BodySmall"), P("C -> D", "BodySmall"), P("C before D", "BodySmall")],
        ],
        colWidths=[22 * mm, 43 * mm, 43 * mm, 60 * mm],
        repeatRows=1,
    )
    order_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PALE_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), INK),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4D4")),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    s.append(order_table)
    s.append(Spacer(1, 2.5 * mm))
    s.append(P(
        "This rule removes the <b>circular-wait condition</b>, one of the four necessary conditions for deadlock. A process holding a higher-ranked resource can never request a lower-ranked one, so a closed cycle of waiting edges cannot form. It also preserves the two instances of E; both are treated as the same resource class in the ordering (Eduvos, 2026c, slides 65 and 70).",
    ))

    s.append(P("3.3 Prevention or avoidance (5 marks)", "SubQuestion"))
    s.append(P(
        "<b>Deadlock prevention</b> is more suitable for this manufacturing system. The four processes repeat known operations and use a small, fixed set of resources, so the company can enforce a global request order in the control software. The rule is simple, deterministic and easy to test, which matters in a physical production line where a deadlock can stop machinery and delay output. Deadlock avoidance, such as the Banker's algorithm, would check whether each allocation keeps the system in a safe state. It can allow more flexible allocations, but it needs accurate maximum-demand information and adds a decision step to every request. That overhead and complexity are unnecessary here because the resource pattern is predictable. The prevention approach may reduce some concurrency, but the safety and operational simplicity justify that trade-off (Eduvos, 2026c).",
    ))

    # Question 4
    s.append(P("Question 4 - File management and caching (35 marks)", "Question"))
    s.append(P("4.1 Three advantages of directories (3 marks)", "SubQuestion"))
    s.append(bullet_items([
        "<b>Organisation:</b> related files can be grouped by department or purpose, making employee records easier to locate and maintain.",
        "<b>Name separation:</b> different directories can contain files with the same name without a conflict because the full paths are different.",
        "<b>Security and administration:</b> permissions, backup rules and quotas can be applied to a directory and inherited by the files inside it. POSIX also describes files as a hierarchy in which directories form the non-terminal nodes (The Open Group, 2024).",
    ]))

    s.append(P("4.2 Full path and Command Prompt line (4 marks)", "SubQuestion"))
    s.append(P("The file is under Root (C:), then IT, then Software. Its full path is:", "BodyA"))
    s.append(P("C:\\IT\\Software\\Antivirus.exe", "CodeA"))
    s.append(P("A Command Prompt line that runs it using the full path is:", "BodyA"))
    s.append(P('"C:\\IT\\Software\\Antivirus.exe"', "CodeA"))
    s.append(P(
        "The quotation marks are safe practice because they also work when a future directory name contains spaces. A fully qualified Windows path starts with the drive and follows each parent directory to the target file (Microsoft, 2025c).",
    ))

    s.append(P("4.3 HR and Public contents after the actions (4 marks)", "SubQuestion"))
    action_table = Table(
        [
            [P("HR directory", "BodySmallWhite"), P("Public directory", "BodySmallWhite")],
            [P("Employees.xlsx<br/>Salaries.xlsx<br/>Contracts.docx<br/><b>Policies.pdf</b> (copied in)", "BodySmall"), P("Policies.pdf<br/>Forms.docx", "BodySmall")],
        ],
        colWidths=[84 * mm, 84 * mm],
    )
    action_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PALE_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), INK),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8C4D4")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    s.append(action_table)
    s.append(Spacer(1, 2.5 * mm))
    s.append(P(
        "Public still contains Policies.pdf because the administrator <b>copied</b> the file rather than moving it. Deleting JanBackup.zip changes only IT\\Backup. Renaming Expenses.xlsx changes only Finance, where the new name becomes Expenses_2026.xlsx. Therefore those two actions do not alter HR or Public.",
    ))

    s.append(P("4.4 Delete-on-logout compared with keep-until-deleted (8 marks)", "SubQuestion"))
    s.append(P(
        "A system that automatically deletes user files at logout or job termination has strong <b>security and storage-management</b> benefits. Temporary data does not remain for the next user on a shared computer, abandoned jobs do not fill the disk, and administrators spend less time cleaning old files. This approach suits kiosks, examination labs, temporary cloud jobs and systems that process sensitive short-lived data. Its main weakness is the risk of permanent data loss. A crash, forgotten save request or misunderstood policy can erase legitimate work. It also places a continuous burden on users to decide what must be preserved.",
    ))
    s.append(P(
        "A system that keeps files until the user deletes them gives better <b>continuity and recovery</b>. Users can log out and continue later, records remain available for audits, and an accidental logout does not destroy work. This suits personal computers, office systems and business records. The disadvantages are that storage fills with stale files, private information remains exposed for longer, and users often fail to clean up or apply retention rules. Keeping everything can also increase backup cost and make searches less effective.",
    ))
    s.append(P(
        "Neither extreme is ideal for every system. I would use a hybrid policy: automatically delete known temporary files and session caches; keep deliberately saved user documents; use quotas, retention periods and warnings; and provide a recycle or backup window for recovery. The policy should match the value, sensitivity and legal retention needs of the data.",
    ))

    s.append(P("4.5a Why systems treat file types differently (4 marks)", "SubQuestion"))
    s.append(P(
        "Some systems track file type because the operating system can then choose a suitable application, icon and permitted operation. For example, Windows file associations determine which application launches when a user opens a particular extension and which actions appear in the interface (Microsoft, 2021). File types can also help the OS distinguish regular files, directories, executables, devices and other special objects, improving validation and protection.",
    ))
    s.append(P(
        "Other systems leave type information to the user or application, often through a filename extension or the file's internal format. This is simpler and flexible because the operating system can treat the content as a sequence of bytes. Systems that do not implement many file types reduce kernel complexity and avoid locking applications into a fixed type system. The trade-off is that applications must interpret the data correctly, and a misleading extension may cause the wrong program to be selected.",
    ))
    s.append(P("4.5b Which system is better? (4 marks)", "SubQuestion"))
    s.append(P(
        "There is no universally better design. Strong OS-managed typing is useful where safety, device control and a consistent user experience matter. Application-managed typing is better where flexibility, portability and simple file-system design matter. My preferred approach is a <b>hybrid</b>: the OS should keep a small set of essential structural types, such as directory, regular file, symbolic link and device, while applications use extensions, registered associations, MIME information or content signatures for detailed formats. This gives the OS enough information to protect the system without forcing every future document format into the kernel.",
    ))

    s.append(P("4.6a How caches improve performance (4 marks)", "SubQuestion"))
    s.append(P(
        "A cache keeps recently or frequently used data in a faster storage layer. When the requested item is already present, the system serves a <b>cache hit</b> instead of repeating a slower disk, network or main-memory access. This reduces average access time and traffic to the slower resource. Caches work because programs often show <b>temporal locality</b>, meaning recently used data is likely to be used again, and <b>spatial locality</b>, meaning nearby data is likely to be needed soon. The course slides describe cache management as temporarily storing frequently accessed data in faster storage (Eduvos, 2026a, slide 43). File-system caches can therefore avoid repeated disk reads, while CPU caches reduce delays caused by main memory.",
    ))
    s.append(P(
        "The improvement can be expressed as <b>average access time = hit time + (miss rate x miss penalty)</b>. For example, if a cache hit takes 1 ms, the miss rate is 10%, and a miss adds another 9 ms, the average is 1 + (0.10 x 9) = <b>1.9 ms</b>, compared with approximately 10 ms when every request uses the slower layer. Coherency and invalidation rules are still required so that faster access does not return stale data (Linux Kernel, 2026).",
    ))
    s.append(P("4.6b Why systems do not use unlimited or larger caches (4 marks)", "SubQuestion"))
    s.append(P(
        "Fast cache memory is limited, expensive and consumes chip area or system RAM. A larger cache can improve the hit rate, but the benefit eventually becomes smaller because not all workloads reuse the extra data. Larger structures may also take longer to search, use more power, require more metadata and make coherency between cores or devices more complex. At file-system level, a very large cache can take memory away from applications and must still manage dirty data, invalidation and recovery. It can also retain useless data and cause cache pollution. The best cache size is therefore workload-specific: Intel notes that cache effectiveness depends on whether the program's critical code and data fit the cache, and a poorly matched cache may provide little or even negative benefit (Intel, 2023).",
    ))

    # References and disclosure
    s.append(P("Reference list", "Question"))
    references = [
        "Eduvos. 2026a. <i>ITOPA3-T11 Week 0 and 1 Lessons 1-4</i>. Internal course slides, especially slides 41 and 43.",
        "Eduvos. 2026b. <i>ITOPA3-T11 Week 2 Lessons 5-7</i>. Internal course slides, especially slides 46, 47 and 54.",
        "Eduvos. 2026c. <i>ITOPA3-T11 Week 3 Lessons 8-10</i>. Internal course slides, especially slides 65, 66 and 70.",
        "Intel. 2023. <i>Effective Use of Cache Memory</i>. Nios II Processor Reference Guide. Available at: https://www.intel.com/content/www/us/en/docs/programmable/683836/current/effective-use-of-cache-memory.html (Accessed 27 August 2026).",
        "Linux Kernel. 2026. <i>Network Filesystem Caching API</i>. Available at: https://docs.kernel.org/filesystems/caching/netfs-api.html (Accessed 27 August 2026).",
        "Microsoft. 2021. <i>How File Associations Work</i>. Available at: https://learn.microsoft.com/en-us/windows/win32/shell/fa-how-work (Accessed 27 August 2026).",
        "Microsoft. 2025a. <i>User Mode and Kernel Mode</i>. Available at: https://learn.microsoft.com/en-us/windows-hardware/drivers/gettingstarted/user-mode-and-kernel-mode (Accessed 27 August 2026).",
        "Microsoft. 2025b. <i>Windows Security Model for Driver Developers</i>. Available at: https://learn.microsoft.com/en-us/windows-hardware/drivers/driversecurity/windows-security-model (Accessed 27 August 2026).",
        "Microsoft. 2025c. <i>About Path Syntax</i>. Available at: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_path_syntax (Accessed 27 August 2026).",
        "Oracle. 2025. <i>ThreadPoolExecutor, Java SE 24 API</i>. Available at: https://docs.oracle.com/en/java/javase/24/docs/api/java.base/java/util/concurrent/ThreadPoolExecutor.html (Accessed 27 August 2026).",
        "The Open Group. 2024. <i>POSIX.1-2024 General Concepts and System Interfaces</i>. Available at: https://pubs.opengroup.org/onlinepubs/9799919799/ (Accessed 27 August 2026).",
    ]
    for ref in references:
        s.append(P(ref, "Reference"))

    s.append(PageBreak())
    s.append(P("AI assistance disclosure", "Question"))
    s.append(P(
        "<b>Tool:</b> OpenAI Codex (GPT-5.6).<br/><b>Use:</b> The tool extracted and interpreted the assignment brief, reviewed the supplied Week 0-3 lecture slides, researched missing file-management details using official technical sources, drafted and refined answers to Questions 1-4 against the mark allocations, produced the resource-allocation graph and execution trace, filled the non-signature fields of the supplied coversheet, and formatted this PDF.<br/><b>Student responsibility:</b> The student must verify the content against the course material, revise the wording where required, complete any originality process, and make an honest decision about signing the coversheet. No signature or declaration response was inserted by the AI.",
        "Callout",
    ))
    s.append(P("Prompt record", "SubQuestion"))
    s.append(bullet_items([
        '"If you go to my downloads on my macbook, there is a file title ITOPA practical, its recent, extract it here please and go through it and let me know what it needs me to do."',
        '"Its due Friday night, so lets plan and get it done."',
        '"mowbray, eduv4948467, Bongani Mahlangu, use these files for now" with three Eduvos ITOPA slide decks.',
        '"continue till the end and answer all question in my tone. also that is the coversheet we should use" with the Individual Assessment Coversheet V1.2 PDF.',
        '"So let\'s say this was AI generated and I had sent this to you to review. How many marks do you think I\'d be getting if I submitted it?"',
        '"so how can we bump it up" followed by "okay lets do that then".',
    ]))
    return s


def create_answer_pdf():
    doc = AssignmentDocTemplate(
        str(ANSWER_PDF),
        pagesize=LETTER,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=45,
        title="ITOPA3-33 Operating Systems Assignment",
        author="Joshua Nehohwa",
        subject="Operating Systems individual assignment - AI-assisted study draft",
    )
    doc.multiBuild(build_story())


def create_filled_cover():
    source = PdfReader(str(COVER_SOURCE))
    page = source.pages[0]
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)

    overlay_buffer = io.BytesIO()
    c = canvas.Canvas(overlay_buffer, pagesize=(width, height))
    c.setFillColor(INK)
    c.setFont(COVER_FONT, 10)
    fields = [
        ("Mowbray", 139, 621.3),
        ("Information Technology", 139, 604.1),
        ("ITOPA3-33", 139, 586.9),
        ("Individual", 139, 569.7),
        ("Bongani Mahlangu", 139, 552.5),
        ("Joshua Nehohwa", 139, 535.3),
        ("EDUV4948467", 139, 518.1),
    ]
    for value, x, y in fields:
        c.drawString(x, y, value)
    c.save()
    overlay_buffer.seek(0)
    overlay = PdfReader(overlay_buffer).pages[0]
    page.merge_page(overlay)
    writer = PdfWriter()
    letter_page = writer.add_blank_page(width=LETTER[0], height=LETTER[1])
    scale = 0.88
    x_offset = (LETTER[0] - width * scale) / 2
    y_offset = (LETTER[1] - height * scale) / 2 + 5
    letter_page.merge_transformed_page(
        page,
        Transformation().scale(scale).translate(x_offset, y_offset),
    )

    footer_buffer = io.BytesIO()
    footer = canvas.Canvas(footer_buffer, pagesize=LETTER)
    footer.setStrokeColor(TABLE_BLUE)
    footer.setLineWidth(0.8)
    footer.line(36, LETTER[1] - 36, LETTER[0] - 36, LETTER[1] - 36)
    footer.setFont(COVER_FONT, 9.1)
    footer.setFillColor(colors.black)
    footer.drawRightString(
        LETTER[0] - 36,
        27,
        "ITOPA3-33_Assignment_Mowbray_eduv4948467 | Page 1",
    )
    footer.save()
    footer_buffer.seek(0)
    letter_page.merge_page(PdfReader(footer_buffer).pages[0])
    writer.add_metadata({
        "/Title": "ITOPA3-33 Assignment Coversheet",
        "/Author": "Joshua Nehohwa",
    })
    with FILLED_COVER.open("wb") as handle:
        writer.write(handle)


def merge_final():
    writer = PdfWriter()
    for source_path in (FILLED_COVER, ANSWER_PDF):
        reader = PdfReader(str(source_path))
        for page in reader.pages:
            writer.add_page(page)
    writer.add_metadata({
        "/Title": "ITOPA3-33 - Assignment - Mowbray - EDUV4948467",
        "/Author": "Joshua Nehohwa",
        "/Subject": "Operating Systems individual assignment - AI-assisted study draft",
        "/Keywords": "ITOPA3-33, Operating Systems, Eduvos, Mowbray",
    })
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("wb") as handle:
        writer.write(handle)


def validate():
    reader = PdfReader(str(OUTPUT))
    if len(reader.pages) < 8:
        raise RuntimeError(f"Unexpectedly short final PDF: {len(reader.pages)} pages")
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    required = [
        "Joshua Nehohwa",
        "EDUV4948467",
        "Question 1",
        "Question 2",
        "Question 3",
        "Question 4",
        "Resource Allocation Graph",
        "AI assistance disclosure",
        "C:\\IT\\Software\\Antivirus.exe",
    ]
    missing = [item for item in required if item not in text]
    if missing:
        raise RuntimeError(f"Required content missing: {missing}")
    print(f"Created {OUTPUT}")
    print(f"Pages: {len(reader.pages)}")
    print(f"Bytes: {OUTPUT.stat().st_size}")


if __name__ == "__main__":
    WORK.mkdir(parents=True, exist_ok=True)
    create_answer_pdf()
    create_filled_cover()
    merge_final()
    validate()
