"""Build an ITECA-style ITEPA draft and attach the official coversheet."""

from __future__ import annotations

import io
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import build_report as base


ROOT = Path(__file__).resolve().parents[1]
COVER_SOURCE = Path(
    "/Users/joshuanehohwa/Library/CloudStorage/OneDrive-Eduvos/Documents/"
    "Block 4 2025 Projects/Individual Assignment Coversheet V1.2 - Copy (2).pdf"
)
WORK = ROOT / "tmp" / "pdfs" / "style_matched_draft"
BODY_PDF = WORK / "itepa_body_letter.pdf"
FILLED_COVER = WORK / "itepa_coversheet_draft.pdf"
OUTPUT = ROOT / "output" / "pdf" / "Joshua_Nehohwa_ITEPA3-33_Style_Matched_Draft.pdf"
FOOTER = "ITEPA_PRACTICAL_MOWBRAY_EDUV4948467"

ARIAL = "/System/Library/Fonts/Supplemental/Arial.ttf"
ARIAL_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
ARIAL_ITALIC = "/System/Library/Fonts/Supplemental/Arial Italic.ttf"
COURIER = "/System/Library/Fonts/Courier New.ttf"


def register_fonts() -> None:
    pdfmetrics.registerFont(TTFont("UserArial", ARIAL))
    pdfmetrics.registerFont(TTFont("UserArialBold", ARIAL_BOLD))
    pdfmetrics.registerFont(TTFont("UserArialItalic", ARIAL_ITALIC))
    if Path(COURIER).exists():
        pdfmetrics.registerFont(TTFont("UserCourier", COURIER))


def apply_reference_styles() -> None:
    base.styles["H1x"].fontName = "UserArialBold"
    base.styles["H1x"].fontSize = 14
    base.styles["H1x"].leading = 17
    base.styles["H1x"].textColor = colors.black
    base.styles["H1x"].spaceBefore = 7
    base.styles["H1x"].spaceAfter = 7

    base.styles["H2x"].fontName = "UserArialBold"
    base.styles["H2x"].fontSize = 11.5
    base.styles["H2x"].leading = 14
    base.styles["H2x"].textColor = colors.black
    base.styles["H2x"].spaceBefore = 9
    base.styles["H2x"].spaceAfter = 5

    base.styles["H3x"].fontName = "UserArialBold"
    base.styles["H3x"].fontSize = 10
    base.styles["H3x"].leading = 13
    base.styles["H3x"].textColor = colors.black

    base.styles["Bodyx"].fontName = "UserArial"
    base.styles["Bodyx"].fontSize = 9.5
    base.styles["Bodyx"].leading = 13.5
    base.styles["Bodyx"].textColor = colors.black
    base.styles["Bodyx"].spaceAfter = 6

    base.styles["Smallx"].fontName = "UserArial"
    base.styles["Smallx"].fontSize = 7.7
    base.styles["Smallx"].leading = 10
    base.styles["Smallx"].textColor = colors.black

    base.styles["Captionx"].fontName = "UserArialItalic"
    base.styles["Captionx"].fontSize = 7.5
    base.styles["Captionx"].leading = 10
    base.styles["Captionx"].textColor = colors.black

    base.styles["Callout"].fontName = "UserArial"
    base.styles["Callout"].fontSize = 9.5
    base.styles["Callout"].leading = 13.5
    base.styles["Callout"].textColor = colors.black
    base.styles["Callout"].backColor = colors.HexColor("#f2f2f2")
    base.styles["Callout"].borderColor = colors.HexColor("#b7b7b7")
    base.styles["Callout"].borderWidth = 0.5

    base.styles["CodeX"].fontName = "UserCourier" if "UserCourier" in pdfmetrics.getRegisteredFontNames() else "Courier"
    base.styles["CodeX"].fontSize = 7.2
    base.styles["CodeX"].leading = 9.5
    base.styles["CodeX"].textColor = colors.black
    base.styles["CodeX"].backColor = colors.HexColor("#f4f4f4")


def restyle_tables(flowables: list[object]) -> None:
    for item in flowables:
        if isinstance(item, Table):
            item.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                        ("FONTNAME", (0, 0), (-1, 0), "UserArialBold"),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9a9a9a")),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7f7f7")]),
                    ]
                )
            )
        elif isinstance(item, KeepTogether):
            restyle_tables(list(item._content))


def body_footer(pdf_canvas: canvas.Canvas, document: SimpleDocTemplate) -> None:
    pdf_canvas.saveState()
    width, _ = letter
    pdf_canvas.setFont("UserArial", 7)
    pdf_canvas.setFillColor(colors.black)
    pdf_canvas.drawString(17 * mm, 10 * mm, FOOTER)
    pdf_canvas.setFillColor(colors.HexColor("#666666"))
    pdf_canvas.drawRightString(width - 17 * mm, 10 * mm, "DRAFT - AI-ASSISTED")
    pdf_canvas.restoreState()


def create_body() -> None:
    story = base.build_story()
    # Remove the decorative generated cover; the official Eduvos coversheet is used instead.
    if len(story) >= 3 and isinstance(story[2], PageBreak):
        story = story[3:]

    title_style = ParagraphStyle(
        "ReferenceTitle",
        fontName="UserArialBold",
        fontSize=17,
        leading=21,
        textColor=colors.black,
        spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "ReferenceSubtitle",
        fontName="UserArial",
        fontSize=10,
        leading=14,
        textColor=colors.black,
        spaceAfter=9,
    )
    draft_style = ParagraphStyle(
        "DraftNotice",
        fontName="UserArialBold",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#555555"),
        backColor=colors.HexColor("#eeeeee"),
        borderPadding=5,
        spaceAfter=9,
    )
    opening = [
        Paragraph("ITEPA3-33 Practical Assignment: EduCore Enterprise Application", title_style),
        Paragraph("Python enterprise application prototype, evidence and design evaluation", subtitle_style),
        Paragraph("WORKING DRAFT - AI ASSISTANCE IS DISCLOSED IN THE REPORT AND AI USE LOG", draft_style),
    ]
    story = opening + story
    restyle_tables(story)

    document = SimpleDocTemplate(
        str(BODY_PDF),
        pagesize=letter,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=17 * mm,
        title="ITEPA3-33 EduCore Style-Matched Draft",
        author="Joshua Nehohwa",
        subject="Working draft with AI assistance disclosed",
    )
    document.build(story, onFirstPage=body_footer, onLaterPages=body_footer)


def create_filled_cover() -> None:
    reader = PdfReader(str(COVER_SOURCE))
    source_page = reader.pages[0]
    width = float(source_page.mediabox.width)
    height = float(source_page.mediabox.height)

    overlay_buffer = io.BytesIO()
    overlay_canvas = canvas.Canvas(overlay_buffer, pagesize=(width, height))
    overlay_canvas.setFillColor(colors.black)
    overlay_canvas.setFont("UserArial", 10)
    fields = [
        ("Mowbray", 139, 638.5),
        ("Information Technology", 139, 621.3),
        ("ITEPA3-33", 139, 604.1),
        ("Individual", 139, 586.9),
        ("", 139, 569.7),
        ("Joshua Nehohwa", 139, 552.5),
        ("EDUV4948467", 139, 535.3),
    ]
    for value, x, y in fields:
        overlay_canvas.drawString(x, y, value)

    overlay_canvas.setFont("UserArialBold", 8)
    overlay_canvas.setFillColor(colors.HexColor("#666666"))
    overlay_canvas.drawRightString(width - 40, height - 35, "DRAFT - AI-ASSISTED")
    overlay_canvas.save()
    overlay_buffer.seek(0)
    source_page.merge_page(PdfReader(overlay_buffer).pages[0])

    writer = PdfWriter()
    writer.add_page(source_page)
    writer.add_metadata({
        "/Title": "ITEPA3-33 Draft Coversheet",
        "/Author": "Joshua Nehohwa",
        "/Subject": "Unsigned working draft",
    })
    with FILLED_COVER.open("wb") as stream:
        writer.write(stream)


def merge_final() -> None:
    writer = PdfWriter()
    for path in (FILLED_COVER, BODY_PDF):
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
    writer.add_metadata({
        "/Title": "ITEPA3-33 Practical - Mowbray - EDUV4948467 - Style-Matched Draft",
        "/Author": "Joshua Nehohwa",
        "/Subject": "Working draft with AI assistance disclosed",
        "/Keywords": "ITEPA3-33, EduCore, Mowbray, draft, AI-assisted",
    })
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("wb") as stream:
        writer.write(stream)


def validate() -> None:
    reader = PdfReader(str(OUTPUT))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    required = [
        "Individual Assessment Coversheet",
        "Mowbray",
        "ITEPA3-33",
        "Joshua Nehohwa",
        "EDUV4948467",
        "DRAFT - AI-ASSISTED",
        "Deliverable 1",
        "Deliverable 5",
        "AI-use statement",
    ]
    missing = [value for value in required if value not in text]
    if missing:
        raise RuntimeError(f"Required content missing from draft: {missing}")
    print(OUTPUT)
    print(f"Pages: {len(reader.pages)}")
    print(f"Bytes: {OUTPUT.stat().st_size}")


def main() -> None:
    WORK.mkdir(parents=True, exist_ok=True)
    register_fonts()
    apply_reference_styles()
    create_body()
    create_filled_cover()
    merge_final()
    validate()


if __name__ == "__main__":
    main()

