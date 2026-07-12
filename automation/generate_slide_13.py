from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs/decks/slide-13.deck.pptx"

BG = RGBColor(246, 248, 250)
TITLE = RGBColor(20, 20, 20)
SUB = RGBColor(70, 70, 70)
WHITE = RGBColor(255, 255, 255)
CHARCOAL = RGBColor(54, 58, 66)
TEAL = RGBColor(30, 95, 102)
PANEL = RGBColor(241, 244, 247)
BORDER = RGBColor(220, 225, 230)


def add_text(slide, left, top, width, height, text, size=14, bold=False, color=TITLE, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return box


def add_rect(slide, left, top, width, height, fill, line=BORDER):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    return shape


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = BG

    add_text(slide, 0.45, 0.16, 12.2, 0.5, "Worked Example: Moderate Income Household", size=30, bold=True)
    add_text(slide, 0.45, 0.74, 12.2, 0.42, "Illustrative progression from occupancy to ownership using income-linked payments.", size=12, color=SUB)

    # Left assumptions and mini table
    add_rect(slide, 0.55, 1.35, 4.2, 4.95, PANEL, BORDER)
    add_text(slide, 0.75, 1.5, 3.8, 0.2, "Example Assumptions", size=10, bold=True, color=SUB)
    assumptions = [
        "Starting income: $95k",
        "Payment rate: 10%",
        "Income growth: 3% p.a.",
        "Annual floor: $8k",
        "Target cap: indexed",
    ]
    y = 1.76
    for line in assumptions:
        add_text(slide, 0.78, y, 3.7, 0.2, "• " + line, size=9, color=TITLE)
        y += 0.26

    add_text(slide, 0.75, 3.1, 3.8, 0.2, "Mini Progress Table", size=10, bold=True, color=SUB)
    add_rect(slide, 0.75, 3.34, 3.8, 2.65, RGBColor(255, 255, 255), RGBColor(210, 216, 223))

    rows = [
        ("Year 1", "$9.5k", "$9.5k"),
        ("Year 5", "$10.7k", "$50.7k"),
        ("Year 10", "$12.4k", "$108.8k"),
        ("Year 15", "$14.4k", "$175.4k"),
    ]
    add_text(slide, 0.9, 3.45, 1.0, 0.16, "Year", size=8, bold=True, color=SUB)
    add_text(slide, 2.15, 3.45, 1.2, 0.16, "Annual", size=8, bold=True, color=SUB)
    add_text(slide, 3.3, 3.45, 1.2, 0.16, "Cumulative", size=8, bold=True, color=SUB)
    y = 3.72
    for year, annual, cum in rows:
        add_text(slide, 0.9, y, 1.1, 0.16, year, size=8, color=TITLE)
        add_text(slide, 2.15, y, 1.2, 0.16, annual, size=8, color=TITLE)
        add_text(slide, 3.3, y, 1.2, 0.16, cum, size=8, color=TITLE)
        y += 0.43

    # Right chart panel
    add_rect(slide, 4.95, 1.35, 7.9, 4.95, PANEL, BORDER)
    add_text(slide, 5.15, 1.5, 3.8, 0.2, "Cumulative Progress Trajectory", size=10, bold=True, color=SUB)

    chart_left = 5.35
    chart_top = 1.95
    chart_w = 7.2
    chart_h = 3.9
    add_rect(slide, chart_left, chart_top, chart_w, chart_h, RGBColor(255, 255, 255), RGBColor(210, 216, 223))

    # Axes
    add_rect(slide, chart_left + 0.45, chart_top + 0.2, 0.01, 3.3, RGBColor(170, 176, 184), RGBColor(170, 176, 184))
    add_rect(slide, chart_left + 0.45, chart_top + 3.5, 6.45, 0.01, RGBColor(170, 176, 184), RGBColor(170, 176, 184))

    # Y labels
    add_text(slide, chart_left + 0.05, chart_top + 0.2, 0.35, 0.16, "$300k", size=7, color=SUB, align=PP_ALIGN.RIGHT)
    add_text(slide, chart_left + 0.05, chart_top + 1.3, 0.35, 0.16, "$200k", size=7, color=SUB, align=PP_ALIGN.RIGHT)
    add_text(slide, chart_left + 0.05, chart_top + 2.4, 0.35, 0.16, "$100k", size=7, color=SUB, align=PP_ALIGN.RIGHT)

    # Draw line as connected small segments
    points = [
        (chart_left + 0.6, chart_top + 3.38),
        (chart_left + 1.7, chart_top + 2.95),
        (chart_left + 3.0, chart_top + 2.35),
        (chart_left + 4.4, chart_top + 1.55),
        (chart_left + 5.8, chart_top + 0.72),
        (chart_left + 6.8, chart_top + 0.38),
    ]
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i + 1]
        # approximate segment with narrow rectangle
        seg_w = max(0.08, x2 - x1)
        seg_y = min(y1, y2)
        seg_h = max(0.06, abs(y2 - y1) + 0.06)
        add_rect(slide, x1, seg_y, seg_w, seg_h, TEAL, TEAL)

    for x, y in points:
        add_rect(slide, x - 0.03, y - 0.03, 0.06, 0.06, TEAL, TEAL)

    add_text(slide, chart_left + 5.65, chart_top + 0.1, 1.4, 0.2, "Completion zone", size=8, bold=True, color=TEAL)
    add_rect(slide, chart_left + 6.45, chart_top + 0.22, 0.5, 0.24, RGBColor(230, 244, 247), RGBColor(180, 210, 216))

    add_rect(slide, 0.45, 6.35, 12.4, 0.75, CHARCOAL, CHARCOAL)
    add_text(
        slide,
        0.72,
        6.58,
        12.0,
        0.28,
        "Under moderate-income assumptions, ownership progression is visible, measurable, and policy-auditable.",
        size=12,
        bold=True,
        color=WHITE,
    )
    add_text(slide, 0.45, 7.16, 12.2, 0.2, "Illustrative scenario from slide-map Section 13 assumptions.", size=8, color=SUB)

    notes = (
        "This worked example is illustrative and intended to show plausibility under clear assumptions.\n\n"
        "The mini table and line trajectory communicate that moderate earners can make visible progress toward completion over time.\n\n"
        "The next slide compares how this timeline changes for higher-income pathways under the same rules."
    )
    notes_frame = slide.notes_slide.notes_text_frame
    notes_frame.clear()
    notes_frame.text = notes

    prs.save(str(OUT))


if __name__ == "__main__":
    main()