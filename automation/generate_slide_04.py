from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs/decks/slide-04.deck.pptx"


BG = RGBColor(246, 248, 250)
TITLE = RGBColor(20, 20, 20)
SUB = RGBColor(70, 70, 70)
CHARCOAL = RGBColor(54, 58, 66)
SLATE_BLUE = RGBColor(71, 94, 146)
BURNT_ORANGE = RGBColor(198, 94, 31)
DEEP_RED = RGBColor(170, 45, 45)
WHITE = RGBColor(255, 255, 255)


def add_text(slide, left, top, width, height, text, size=14, bold=False, color=TITLE):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    return box


def add_shape(slide, shape_type, left, top, width, height, fill, line=None):
    s = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = fill if line is None else line
    return s


def add_notes(slide, txt):
    notes = slide.notes_slide
    tf = notes.notes_text_frame
    tf.clear()
    tf.text = txt


def draw_waterfall_chart(slide, left, top, width, height):
    add_shape(slide, 1, left, top, width, height, RGBColor(241, 244, 247), RGBColor(220, 225, 230))
    add_text(slide, left + 0.15, top + 0.05, 4.8, 0.25, "Waterfall: cost formation per dwelling", size=11, bold=True, color=SUB)
    add_text(slide, left + 4.5, top + 0.05, 3.2, 0.25, "Range: $600k to $1.0m", size=10, bold=True, color=SUB)

    labels = [
        "Raw\nland",
        "Rezoning\n/ scarcity",
        "Infra\n+ civil",
        "Construction",
        "Taxes\n+ compliance",
        "Overheads\n+ margin",
        "Final\nprice",
    ]
    ranges = ["20k-30k", "80k-130k", "80k-130k", "300k-500k", "60k-100k", "80k-130k", "600k-1,000k"]
    increments = [20000, 100000, 100000, 400000, 80000, 100000]
    total = 800000
    pct_labels = ["2.5%", "12.5%", "12.5%", "50.0%", "10.0%", "12.5%", "100%"]

    colors = [
        BURNT_ORANGE,
        BURNT_ORANGE,
        BURNT_ORANGE,
        SLATE_BLUE,
        SLATE_BLUE,
        SLATE_BLUE,
        CHARCOAL,
    ]

    plot_left = left + 0.25
    plot_top = top + 0.45
    plot_w = width - 0.5
    plot_h = height - 1.95

    baseline_y = plot_top + plot_h
    add_shape(slide, 1, plot_left, baseline_y, plot_w, 0.01, RGBColor(180, 186, 194), RGBColor(180, 186, 194))

    bar_w = 0.78
    gap = (plot_w - (bar_w * 7)) / 6
    y_scale = plot_h / total

    cum = 0
    for i in range(7):
        x = plot_left + i * (bar_w + gap)
        if i < 6:
            start = cum
            end = cum + increments[i]
            cum = end
            y_start = baseline_y - (start * y_scale)
            y_end = baseline_y - (end * y_scale)
            y = y_end
            h = max(0.06, y_start - y_end)
        else:
            y = baseline_y - (total * y_scale)
            h = total * y_scale

        add_shape(slide, 1, x, y, bar_w, h, colors[i])
        add_text(slide, x - 0.05, y - 0.18, bar_w + 0.1, 0.18, pct_labels[i], size=9, bold=True)

        # Category label and dollar-range below each bar
        add_text(slide, x - 0.07, baseline_y + 0.05, bar_w + 0.14, 0.38, labels[i], size=8, color=SUB)
        add_text(slide, x - 0.06, baseline_y + 0.40, bar_w + 0.12, 0.18, ranges[i], size=8, bold=True, color=TITLE)

    add_text(slide, left + 0.15, top + height - 0.28, width - 0.3, 0.2, "Range labels under bars are dollar equivalents across the $600k to $1.0m band.", size=8, color=SUB)


def draw_mortgage_panel(slide, left, top, width, height):
    add_shape(slide, 1, left, top, width, height, RGBColor(241, 244, 247), RGBColor(220, 225, 230))
    add_text(slide, left + 0.15, top + 0.05, width - 0.3, 0.25, "Mortgage-capacity scenario", size=11, bold=True, color=SUB)

    in_top = top + 0.35
    add_text(slide, left + 0.15, in_top, width - 0.3, 0.2, "Inputs", size=10, bold=True, color=SUB)
    inputs = [
        "Income: $120,000 p.a.",
        "Deposit: $40,000",
        "Lending multiple: 5x",
        "Rate: 6.0%",
        "Term: 30 years",
    ]
    y = in_top + 0.18
    for line in inputs:
        add_text(slide, left + 0.2, y, width - 0.4, 0.2, line, size=9, color=TITLE)
        y += 0.2

    out_top = y + 0.08
    add_text(slide, left + 0.15, out_top, width - 0.3, 0.2, "Outputs", size=10, bold=True, color=SUB)

    outputs = [
        "Max loan @ 5x: $600,000",
        "Assumed purchase: $600,000",
        "Loan required: $560,000",
        "Monthly repayment: $3,358",
        "Annual repayment: $40,296",
        "Repayment burden: 33.6% of income",
    ]
    y = out_top + 0.18
    for i, line in enumerate(outputs):
        add_text(slide, left + 0.2, y, width - 0.4, 0.2, line, size=9, bold=(i in (0, 5)), color=DEEP_RED if i in (3, 5) else TITLE)
        y += 0.2

    add_text(
        slide,
        left + 0.15,
        top + height - 0.25,
        width - 0.3,
        0.2,
        "Repayment excludes rates, insurance, maintenance, and transaction costs.",
        size=8,
        color=SUB,
    )


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG

    # Top band
    add_text(slide, 0.45, 0.16, 12.2, 0.55, "Diagnosis: What Is Actually Expensive?", size=30, bold=True)
    add_text(
        slide,
        0.45,
        0.76,
        12.2,
        0.45,
        "Show cost build-up as a waterfall from raw land to final price, then show how lending-capacity settings push households to max borrow.",
        size=12,
        color=SUB,
    )

    draw_waterfall_chart(slide, left=0.45, top=1.35, width=8.15, height=4.75)
    draw_mortgage_panel(slide, left=8.75, top=1.35, width=4.1, height=3.35)

    # Right-side insights
    insight = add_shape(slide, 1, 8.75, 4.85, 4.1, 1.25, RGBColor(248, 250, 252), RGBColor(224, 228, 233))
    _ = insight
    add_text(slide, 8.9, 4.92, 3.8, 0.2, "Insights", size=10, bold=True, color=SUB)
    bullets = [
        "Construction is largest, but not the only price driver.",
        "Raw land is small; uplift happens later in the pipeline.",
        "Land transformation, infra, and compliance create major uplift.",
        "Max-lend norms can support inflated pricing toward debt limits.",
    ]
    y = 5.12
    for line in bullets:
        add_text(slide, 9.0, y, 3.6, 0.2, line, size=9, color=TITLE)
        y += 0.23

    # Bottom takeaway
    add_shape(slide, 1, 0.45, 6.28, 12.4, 0.72, CHARCOAL)
    add_text(
        slide,
        0.72,
        6.50,
        12.0,
        0.28,
        "Prices are often pulled toward max deposit plus max borrowing capacity, allowing margins and embedded costs to expand beyond core build economics.",
        size=12,
        bold=True,
        color=WHITE,
    )
    add_text(
        slide,
        0.45,
        7.05,
        12.2,
        0.2,
        "Illustrative scenario calibrated from project research ranges in new_home_development.txt.",
        size=8,
        color=SUB,
    )

    notes = (
        "A common framing is that housing is expensive because it is expensive to build. This chart shows a broader reality.\n\n"
        "The left chart uses a waterfall to show cumulative price formation from raw land through delivery layers to final price. In outer major-city areas, the final price can sit in a 600,000 to 1,000,000 range.\n\n"
        "Under each bar, ranges show how the same category scales across that 600,000 to 1,000,000 market band.\n\n"
        "Land-system and delivery layers create large uplift before the keys are handed over: land transformation, infrastructure, compliance, and risk-loaded margins.\n\n"
        "On the right, the mortgage-capacity panel shows the financing anchor. For a household on 120,000 with a 40,000 deposit and a 5x lending multiple, maximum borrowing capacity is 600,000, but this scenario holds purchase price at 600,000.\n\n"
        "At 6% over 30 years, the required 560,000 loan implies repayments of about 3,358 per month, or about 33.6% of gross income.\n\n"
        "So margins and embedded costs do not just reflect technical delivery input costs. They can expand toward available finance.\n\n"
        "That is why policy reform has to focus on system economics, not only construction productivity."
    )
    add_notes(slide, notes)

    prs.save(str(OUT))


if __name__ == "__main__":
    main()
