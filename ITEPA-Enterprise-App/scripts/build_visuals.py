"""Create polished UI mockups and the microservices architecture diagram."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
DIAGRAMS = ROOT / "diagrams"
FONT_REGULAR = "/System/Library/Fonts/Supplemental/Arial.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    path = FONT_BOLD if bold else FONT_REGULAR
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        return ImageFont.load_default()


def rounded(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str, outline: str = "#dbe3ef", radius: int = 18, width: int = 2) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def shell(title: str, role: str) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (1400, 900), "#f5f7fb")
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1400, 86), fill="#0f2c4c")
    draw.text((42, 24), "EduCore", fill="white", font=font(34, True))
    draw.text((238, 31), title, fill="#c8d8eb", font=font(24))
    rounded(draw, (1170, 20, 1355, 66), "#163d66", "#3f6385", 20)
    draw.text((1200, 32), role, fill="white", font=font(20, True))
    draw.rectangle((0, 86, 245, 900), fill="#ffffff")
    items = ["Dashboard", "Learners", "Courses", "Support", "Reports"]
    for index, item in enumerate(items):
        y = 130 + index * 64
        if item.lower() in title.lower() or (title == "Operations dashboard" and item == "Dashboard"):
            rounded(draw, (24, y - 12, 220, y + 38), "#e6f2ff", "#b5d5f7", 12)
        draw.text((48, y), item, fill="#17324d", font=font(22, item.lower() in title.lower()))
    return image, draw


def label(draw: ImageDraw.ImageDraw, x: int, y: int, text: str, required: bool = False) -> None:
    draw.text((x, y), text + (" *" if required else ""), fill="#27364b", font=font(19, True))


def field(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], placeholder: str, error: str | None = None) -> None:
    rounded(draw, box, "#ffffff", "#dc3545" if error else "#aebed0", 10)
    draw.text((box[0] + 15, box[1] + 14), placeholder, fill="#61748a", font=font(19))
    if error:
        draw.text((box[0], box[3] + 7), error, fill="#b42318", font=font(16))


def button(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], text: str, secondary: bool = False) -> None:
    rounded(draw, box, "#ffffff" if secondary else "#087e8b", "#087e8b", 11)
    draw.text((box[0] + 24, box[1] + 13), text, fill="#087e8b" if secondary else "white", font=font(19, True))


def build_dashboard() -> None:
    image, draw = shell("Operations dashboard", "Administrator")
    draw.text((290, 125), "Good afternoon, Administrator", fill="#132c45", font=font(34, True))
    draw.text((290, 173), "Live operational overview", fill="#64758a", font=font(21))
    cards = [
        ("2,481", "Active learners", "#0e7490"),
        ("18", "Open courses", "#2563eb"),
        ("94.2%", "Registration success", "#059669"),
        ("27", "Open tickets", "#d97706"),
    ]
    for index, (value, title, colour) in enumerate(cards):
        x = 290 + index * 265
        rounded(draw, (x, 225, x + 235, 365), "#ffffff")
        draw.rectangle((x, 225, x + 8, 365), fill=colour)
        draw.text((x + 28, 250), value, fill="#132c45", font=font(34, True))
        draw.text((x + 28, 307), title, fill="#64758a", font=font(18))
    rounded(draw, (290, 405, 870, 820), "#ffffff")
    draw.text((320, 435), "Registration activity", fill="#132c45", font=font(25, True))
    points = [(335, 735), (415, 680), (495, 705), (575, 585), (655, 620), (735, 505), (815, 535)]
    draw.line(points, fill="#087e8b", width=7)
    for point in points:
        draw.ellipse((point[0] - 8, point[1] - 8, point[0] + 8, point[1] + 8), fill="#087e8b")
    rounded(draw, (900, 405, 1350, 820), "#ffffff")
    draw.text((930, 435), "Needs attention", fill="#132c45", font=font(25, True))
    notices = [("PY701", "2 spaces remaining"), ("5 tickets", "Older than 24 hours"), ("3 failures", "Review validation log")]
    for index, (title, detail) in enumerate(notices):
        y = 510 + index * 92
        draw.ellipse((930, y, 952, y + 22), fill="#f59e0b")
        draw.text((970, y - 4), title, fill="#27364b", font=font(20, True))
        draw.text((970, y + 27), detail, fill="#64758a", font=font(17))
    image.save(DIAGRAMS / "ui_01_dashboard.png")


def build_learner() -> None:
    image, draw = shell("Learner registration", "Administrator")
    draw.text((290, 125), "Register a learner", fill="#132c45", font=font(34, True))
    draw.text((290, 173), "Required fields are marked with an asterisk.", fill="#64758a", font=font(20))
    rounded(draw, (290, 220, 1320, 780), "#ffffff")
    label(draw, 335, 265, "Full name", True); field(draw, (335, 300, 790, 360), "e.g. Thabo Mokoena")
    label(draw, 835, 265, "Email address", True); field(draw, (835, 300, 1275, 360), "thabo@example", "Enter a valid email address")
    label(draw, 335, 420, "Learner ID", True); field(draw, (335, 455, 790, 515), "L002481")
    label(draw, 835, 420, "Course", True); field(draw, (835, 455, 1275, 515), "Select an available course")
    draw.text((335, 570), "The learner will receive a confirmation email after registration.", fill="#64758a", font=font(18))
    button(draw, (335, 650, 535, 710), "Register learner"); button(draw, (560, 650, 690, 710), "Cancel", True)
    image.save(DIAGRAMS / "ui_02_learner_registration.png")


def build_courses() -> None:
    image, draw = shell("Course management", "Administrator")
    draw.text((290, 125), "Course management", fill="#132c45", font=font(34, True))
    button(draw, (1110, 120, 1315, 180), "+ Add course")
    rounded(draw, (290, 220, 1320, 755), "#ffffff")
    headers = ["Course", "Capacity", "Enrolled", "Availability", "Actions"]
    xs = [330, 700, 850, 980, 1190]
    draw.rectangle((291, 221, 1319, 290), fill="#eaf0f7")
    for x, text in zip(xs, headers): draw.text((x, 243), text, fill="#27364b", font=font(18, True))
    rows = [("PY701 - Enterprise Python", "30", "28", "2 spaces", "Edit"), ("DS610 - Data Systems", "25", "25", "Full", "Edit"), ("SE520 - Software Design", "40", "17", "23 spaces", "Edit")]
    for index, row in enumerate(rows):
        y = 320 + index * 115
        draw.line((310, y + 77, 1295, y + 77), fill="#e1e7ef", width=2)
        for x, text in zip(xs, row):
            colour = "#b42318" if text == "Full" else "#27364b"
            draw.text((x, y + 20), text, fill=colour, font=font(18, text in ("Full", "Edit")))
    draw.text((325, 690), "Capacity changes are validated against current enrolment.", fill="#64758a", font=font(18))
    image.save(DIAGRAMS / "ui_03_course_management.png")


def build_support() -> None:
    image, draw = shell("Support ticket", "Learner")
    draw.text((290, 125), "How can we help?", fill="#132c45", font=font(34, True))
    rounded(draw, (290, 210, 1320, 790), "#ffffff")
    label(draw, 335, 255, "Category", True); field(draw, (335, 290, 790, 350), "Technical support")
    label(draw, 835, 255, "Priority"); field(draw, (835, 290, 1275, 350), "High - assigned automatically")
    label(draw, 335, 395, "Subject", True); field(draw, (335, 430, 1275, 490), "Cannot access assessment")
    label(draw, 335, 535, "Description", True); field(draw, (335, 570, 1275, 665), "Describe what happened and any error shown")
    button(draw, (335, 705, 520, 765), "Create ticket"); button(draw, (545, 705, 675, 765), "Cancel", True)
    draw.text((850, 720), "Typical response: within 4 hours", fill="#64758a", font=font(18))
    image.save(DIAGRAMS / "ui_04_support_ticket.png")


def build_reports() -> None:
    image, draw = shell("Operational reports", "Administrator")
    draw.text((290, 125), "Operational reports", fill="#132c45", font=font(34, True))
    field(draw, (850, 120, 1070, 180), "Last 30 days"); button(draw, (1095, 120, 1318, 180), "Export report")
    rounded(draw, (290, 220, 1320, 795), "#ffffff")
    draw.text((330, 255), "Registration outcomes", fill="#132c45", font=font(24, True))
    outcomes = [("Confirmed", 842, "#059669"), ("Duplicate", 63, "#d97706"), ("Capacity", 41, "#dc2626"), ("Invalid", 18, "#7c3aed")]
    max_value = max(value for _, value, _ in outcomes)
    for index, (name, value, colour) in enumerate(outcomes):
        y = 330 + index * 80
        draw.text((330, y), name, fill="#27364b", font=font(19))
        draw.rounded_rectangle((490, y, 1110, y + 33), radius=12, fill="#eaf0f7")
        draw.rounded_rectangle((490, y, 490 + int(620 * value / max_value), y + 33), radius=12, fill=colour)
        draw.text((1140, y + 2), str(value), fill="#27364b", font=font(19, True))
    rounded(draw, (330, 675, 1270, 750), "#ecfdf5", "#a7f3d0", 12)
    draw.text((360, 698), "System health: 94.2% successful registrations; p95 latency 18.4 ms", fill="#065f46", font=font(20, True))
    image.save(DIAGRAMS / "ui_05_reports.png")


def arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int], colour: str = "#4f6b88") -> None:
    draw.line((*start, *end), fill=colour, width=5)
    x, y = end
    draw.polygon([(x, y), (x - 16, y - 9), (x - 16, y + 9)], fill=colour)


def build_architecture() -> None:
    image = Image.new("RGB", (1600, 1000), "#f5f7fb")
    draw = ImageDraw.Draw(image)
    draw.text((60, 35), "EduCore microservices readiness architecture", fill="#102a43", font=font(36, True))
    rounded(draw, (60, 120, 270, 220), "#0f2c4c", "#0f2c4c")
    draw.text((110, 151), "Web / mobile", fill="white", font=font(24, True))
    rounded(draw, (390, 120, 650, 220), "#087e8b", "#087e8b")
    draw.text((445, 151), "API Gateway", fill="white", font=font(25, True))
    arrow(draw, (270, 170), (390, 170))
    services = [
        ("Learner Service", "Profile and identity", "#2563eb"),
        ("Course Service", "Catalogue and capacity", "#2563eb"),
        ("Registration Service", "Rules and enrolment", "#7c3aed"),
        ("Assessment Service", "Results and strategies", "#7c3aed"),
        ("Support Service", "Ticket workflow", "#d97706"),
        ("Reporting Service", "Metrics and Bugzot", "#059669"),
    ]
    positions = [(90, 340), (550, 340), (1010, 340), (90, 620), (550, 620), (1010, 620)]
    for (title, subtitle, colour), (x, y) in zip(services, positions):
        rounded(draw, (x, y, x + 380, y + 150), "#ffffff", colour, 18, 4)
        draw.rectangle((x, y, x + 380, y + 55), fill=colour)
        draw.text((x + 25, y + 15), title, fill="white", font=font(21, True))
        draw.text((x + 25, y + 77), subtitle, fill="#31465c", font=font(18))
        draw.text((x + 25, y + 111), "Service-owned database", fill="#64758a", font=font(16))
    for x in (280, 740, 1200):
        arrow(draw, (520, 220), (x, 335))
    rounded(draw, (390, 850, 1210, 940), "#e7eef8", "#9bb0c6", 16)
    draw.text((430, 873), "Event bus: registration-created | assessment-completed | ticket-updated", fill="#27364b", font=font(21, True))
    for x in (280, 740, 1200):
        draw.line((x, 770, x, 850), fill="#4f6b88", width=4)
    draw.text((1130, 160), "Cross-cutting controls", fill="#102a43", font=font(23, True))
    draw.text((1130, 205), "- Correlation IDs\n- Distributed tracing\n- Contract tests\n- Central Bugzot monitoring", fill="#4f6275", font=font(19), spacing=12)
    image.save(DIAGRAMS / "microservices_architecture.png")


def main() -> None:
    DIAGRAMS.mkdir(parents=True, exist_ok=True)
    build_dashboard(); build_learner(); build_courses(); build_support(); build_reports(); build_architecture()


if __name__ == "__main__":
    main()

