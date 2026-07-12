from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs/decks/slide-10.deck.pptx"

BG = RGBColor(246, 248, 250)
TITLE = RGBColor(20, 20, 20)
SUB = RGBColor(70, 70, 70)
WHITE = RGBColor(255, 255, 255)
CHARCOAL = RGBColor(54, 58, 66)
TEAL = RGBColor(30, 95, 102)
ORANGE = RGBColor(198, 94, 31)
SLATE = RGBColor(71, 94, 146)
PANEL = RGBColor(241, 244, 247)
BORDER = RGBColor(220, 225, 230)


def add_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    size=14,
    bold=False,
    color=TITLE,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
):
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


def add_lane(slide, left, top, width, height, title, color, bullets):
    add_rect(slide, left, top, width, height, PANEL, BORDER)
    add_rect(slide, left, top, width, 0.4, color, color)
    add_text(slide, left + 0.14, top + 0.11, width - 0.28, 0.22, title, size=11, bold=True, color=WHITE)

    y = top + 0.5
    for bullet in bullets:
        add_text(slide, left + 0.16, y, width - 0.3, 0.27, "• " + bullet, size=9, color=TITLE)
        y += 0.33


def add_arrow(slide, left, top, width=0.42, height=0.22):
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(left), Inches(top), Inches(width), Inches(height))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor(160, 167, 176)
    arrow.line.color.rgb = RGBColor(160, 167, 176)
    return arrow


def add_chip(slide, left, top, width, text):
    add_rect(slide, left, top, width, 0.34, RGBColor(255, 255, 255), RGBColor(206, 213, 221))
    add_text(slide, left + 0.08, top + 0.1, width - 0.16, 0.16, text, size=8, bold=True, color=SUB, align=PP_ALIGN.CENTER)


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG

    add_text(slide, 0.45, 0.16, 12.2, 0.5, "Procurement Architecture: Public Rules, Competitive Delivery", size=28, bold=True)
    add_text(
        slide,
        0.45,
        0.74,
        12.2,
        0.44,
        "Set the specification once, open the market to compliant bidders, and enforce transparent cost and quality controls.",
        size=12,
        color=SUB,
    )

    # Three-lane system map
    lane_top = 1.35
    lane_h = 3.95
    lane_w = 3.75

    add_lane(
        slide,
        0.55,
        lane_top,
        lane_w,
        lane_h,
        "1. Government Framework",
        TEAL,
        [
            "Define fixed housing product and standards",
            "Set indicative cost cap and review settings",
            "Set compliance, audit, and reporting rules",
            "Control serviced land release into pipeline",
        ],
    )

    add_lane(
        slide,
        4.78,
        lane_top,
        lane_w,
        lane_h,
        "2. Procurement Engine",
        ORANGE,
        [
            "Open participation tender model",
            "Eligibility gate: capability + compliance",
            "Evaluate price, speed, quality, scalability",
            "Award across multiple contractors",
        ],
    )

    add_lane(
        slide,
        9.01,
        lane_top,
        lane_w,
        lane_h,
        "3. Delivery Pathways",
        SLATE,
        [
            "Prefab and modular pathway",
            "Onsite systemised pathway",
            "3D-assisted automated pathway",
            "All pathways meet one output standard",
        ],
    )

    add_arrow(slide, 4.35, 3.05)
    add_arrow(slide, 8.58, 3.05)

    # Right-side outcome chips within bottom row
    chip_y = 5.45
    add_chip(slide, 0.75, chip_y, 2.95, "Open market participation")
    add_chip(slide, 3.9, chip_y, 2.95, "Standardised output quality")
    add_chip(slide, 7.05, chip_y, 2.95, "Cost transparency + discipline")
    add_chip(slide, 10.2, chip_y, 2.35, "Scalable delivery")

    add_rect(slide, 0.45, 6.25, 12.4, 0.72, CHARCOAL, CHARCOAL)
    add_text(
        slide,
        0.72,
        6.46,
        12.0,
        0.27,
        "Government sets the rules and guardrails; private industry competes inside them to deliver the same audited housing product.",
        size=11,
        bold=True,
        color=WHITE,
    )
    add_text(slide, 0.45, 7.04, 12.2, 0.2, "Architecture from docs/slides/slide-10.plan.md.", size=8, color=SUB)

    notes = (
        "This slide presents procurement as a system, not a one-off build contract.\n\n"
        "Government sets the non-negotiables: product definition, compliance, cost-discipline settings, and land-release coordination.\n\n"
        "Procurement remains open to multiple compliant bidders, and contracts are awarded on competitive delivery performance.\n\n"
        "Different construction methods can compete, but each must deliver the same audited output standard.\n\n"
        "This is how policy control and market competition can coexist at scale."
    )
    notes_frame = slide.notes_slide.notes_text_frame
    notes_frame.clear()
    notes_frame.text = notes

    prs.save(str(OUT))


if __name__ == "__main__":
    main()