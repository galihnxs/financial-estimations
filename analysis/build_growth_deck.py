"""
Teman Journey - growth recommendation deck (19 slides + 6 appendix = 25), plain-English version for readers whose
first language is not English and who are not business specialists. Chart-led: every number-heavy message has a
native PowerPoint chart.

Run from financial-estimations/:   python analysis/build_growth_deck.py
Every scenario number is computed by growth_scenarios.py at build time; prices and costs come from model v13.
Colours follow temanjourney.id: slate blue dominant on white, terracotta and beige only as accents.
Writes TemanJourney_Growth_Recommendation.pptx.
"""
import json
import sys
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).parent))
import growth_scenarios as G

OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "TemanJourney_Growth_Recommendation.pptx")
FONT = "Arial"
rgb = lambda h: RGBColor.from_string(h)
# temanjourney.id: text and buttons #475571 / #576C94 / #6379A1 on white; CSS brand variables add terracotta and beige
NAVY, BLUE, MBLUE, LBLUE = rgb("475571"), rgb("576C94"), rgb("6379A1"), rgb("A9B6D1")
TERRA, RED = rgb("9B3809"), rgb("CD5D44")
BEIGE, BEIGE_L = rgb("B3A37C"), rgb("F1DEBB")
INK, MUTED, WHITE, GREY = rgb("374151"), rgb("6B7280"), rgb("FFFFFF"), rgb("B8BEC9")
PANEL, BORDER, BLUE_BG, RED_BG = rgb("F3F5F9"), rgb("E5E7EB"), rgb("DDE3EE"), rgb("F5DCD0")

# ------------------------------------------------------------------ numbers (all from the scenario engine)
L = {name: G.scen(st) for name, st in G.LADDER}
S1, S2, S3, S3B, S4, S5, S6 = (L[n] for n, _ in G.LADDER)
S4_ADM2 = G.scen(["fair", "eap", "cof", "bank", "adm2"])       # full team, CTO paid by Theta
CURVE = G.eap_curve()
UPLIFT, _ = G.price_only_uplift()
RUNS = json.loads((Path(__file__).resolve().parent.parent / "output" / "v14_runs.json").read_text())   # excel_runner.py
REC = RUNS["plan_founders_may"]                                  # our plan: real timing + growth plan, founders paid from May
F27 = REC["fy27"]
NOBANK, BICUT = RUNS["plan_founders_may_no_bank"]["fy27"], RUNS["plan_founders_may_bi_cut"]["fy27"]
FSTART = [("Jan 2027", "plan_founders_jan"), ("Mar 2027", "plan_founders_mar"), ("May 2027", "plan_founders_may"),
          ("Jul 2027", "plan_founders_jul")]
LOWM = G.MONTHS[F27["low_idx"]].replace("-", " 20")
FAIR_MO = G.FAIR_DELTA + G.ADS
FIX_NOW = G.ROWS["F"][3]
FIX_FAIR = FIX_NOW + FAIR_MO
PSY_DELTA = G.FAIR["psych"] - G.CUR["psych"]
EXIST_DELTA = FAIR_MO - PSY_DELTA
STEP1_M = (S1["N"] - 12 * EXIST_DELTA) / S1["R"]
EAP_GP = 12 * (G.EAP_R - G.EAP_C)
GROUP_GP = G.BANK_R - G.BANK_C - G.PMFEE
THETA_MO = 2.5                                                   # Rp5k x 500 staff, per EAP contract
B2B = G.ROWS["b2b"]
FAC = G.ROWS["fac"]
over_pm = [m for m, e in zip(G.MONTHS, B2B) if e > 2]
over_fac = [m for m, f in zip(G.MONTHS, FAC) if f > 2]
BI = sum(G.ROWS["BI"][G.Y27])
EAP_REV1, B2C_REV = 180.0, 123.0
OTHER1 = S1["R"] - BI - EAP_REV1 - B2C_REV
BI_S1, BI_S4 = BI / S1["R"], BI / S4["R"]
EAP_SHARE_S1 = EAP_REV1 / S1["R"]
EAP_SHARE_S3 = 3 * EAP_REV1 / S3["R"]
EAP_SHARE_S4 = 3 * EAP_REV1 / S4["R"]
MKT_MO = G.FAIR["leadmkt"] + G.FAIR["socmed"] + G.ADS
NET27 = REC["net"][3:]
LOSS_MONTHS = sum(1 for v in NET27 if v < 0)
BUFFER = 6 * REC["fixed"][-1]                                    # 6 months of fixed cost, Dec-27 team
INVEST_28 = 0.5 * S4["N"]
M27 = [m[:3] for m in G.MONTHS[3:]]
M15 = [m.replace("-", " ") for m in G.MONTHS]
pc = lambda v, d=0: f"{v * 100:.{d}f}%"
rp = lambda v, d=0: f"Rp{v:,.{d}f}M"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
PAGE = [0]


# ------------------------------------------------------------------ helpers
def text(slide, x, y, w, h, paras, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, margin=0.08, space=4, name=None):
    """paras: str or list of str / (text, opts) / list of runs [(text, {opts})]. Returns the shape."""
    if fill is None and line is None and shape == MSO_SHAPE.RECTANGLE:
        sh = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid(); sh.fill.fore_color.rgb = fill
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line; sh.line.width = Pt(1)
        sh.shadow.inherit = False
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = 0.08
    if name:
        sh.name = name
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_after = Pt(space)
        runs = p if isinstance(p, list) else [p] if isinstance(p, tuple) else [(p, {})]
        for t, o in runs:
            r = para.add_run()
            r.text = t
            f = r.font
            f.name = FONT
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", False)
            f.color.rgb = o.get("color", color)
    return sh


def card(slide, x, y, w, h, paras, size=14, fill=PANEL, **kw):
    return text(slide, x, y, w, h, paras, size=size, fill=fill, line=BORDER if fill == PANEL else None,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, margin=kw.pop("margin", 0.18), **kw)


TAGS = {"goal": ("Our goal", TERRA), "answer": ("Our answer", NAVY), "problem": ("The problem", TERRA),
        "1": ("1  Price", NAVY), "2": ("2  People", BLUE), "3": ("3  Growth", MBLUE), "next": ("Next steps", NAVY),
        "app": ("Appendix", MUTED)}


def new_slide(title, source=None, dark=False, tag=None):
    s = prs.slides.add_slide(BLANK)
    PAGE[0] += 1
    bg = s.background.fill
    bg.solid(); bg.fore_color.rgb = NAVY if dark else WHITE
    if title:
        text(s, 0.5, 0.32, 10.9, 1.0, title, size=24, bold=True, color=NAVY, name="Title")
    if tag:
        lab, f = TAGS[tag]
        text(s, 11.55, 0.4, 1.3, 0.4, lab, size=12, bold=True, color=WHITE, fill=f, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, margin=0.02, name="Tag")
    if source:
        text(s, 0.5, 7.0, 11.6, 0.35, "Source: " + source, size=10, color=MUTED, name="Source")
    text(s, 12.3, 7.0, 0.55, 0.35, str(PAGE[0]), size=10, color=WHITE if dark else MUTED, align=PP_ALIGN.RIGHT, name="Page")
    return s


def circle(slide, x, y, d, label, fill=NAVY, size=16):
    return text(slide, x, y, d, d, label, size=size, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE, fill=fill, shape=MSO_SHAPE.OVAL, margin=0)


def table(slide, x, y, w, col_w, rows, size=12, header_fill=NAVY, row_h=0.36, first_col_bold=True, fills=None):
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr)).table
    for j, cw in enumerate(col_w):
        gt.columns[j].width = Inches(cw)
    for i, r in enumerate(rows):
        gt.rows[i].height = Inches(row_h)
        for j, v in enumerate(r):
            c = gt.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.03)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.word_wrap = True
            tf.paragraphs[0].text = ""
            run = tf.paragraphs[0].add_run()
            run.text = str(v)
            f = run.font
            f.name, f.size = FONT, Pt(size)
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = header_fill
                f.bold, f.color.rgb = True, WHITE
            else:
                c.fill.fore_color.rgb = (fills or {}).get((i, j), PANEL if i % 2 else WHITE)
                f.color.rgb = INK
                f.bold = first_col_bold and j == 0
    return gt


def _style_title(ch, title):
    ch.has_title = True
    ch.chart_title.text_frame.text = title
    tp = ch.chart_title.text_frame.paragraphs[0].runs[0].font
    tp.size, tp.bold, tp.name, tp.color.rgb = Pt(13), True, FONT, INK


def bar(slide, x, y, w, h, cats, series, title, colors=None, point_colors=None, fmt='0%', stacked=False,
        horizontal=False, size=12, legend=False, gap=60, texts=None):
    """series: list of (name, values). colors: one per series. point_colors: per category (single series)."""
    cd = CategoryChartData()
    cd.categories = cats
    for n, v in series:
        cd.add_series(n, v)
    kind = {(False, False): XL_CHART_TYPE.COLUMN_CLUSTERED, (True, False): XL_CHART_TYPE.COLUMN_STACKED,
            (False, True): XL_CHART_TYPE.BAR_CLUSTERED, (True, True): XL_CHART_TYPE.BAR_STACKED}[(stacked, horizontal)]
    ch = slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    _style_title(ch, title)
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size, ch.legend.font.name = Pt(size), FONT
        ch.legend.font.color.rgb = INK
    pl = ch.plots[0]
    pl.gap_width = gap
    if stacked:
        pl.overlap = 100
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format, dl.number_format_is_linked = fmt, False
    dl.position = XL_LABEL_POSITION.CENTER if stacked else XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.name, dl.font.bold = Pt(size + (0 if stacked else 1)), FONT, True
    dl.font.color.rgb = WHITE if stacked else INK
    for i, (n, v) in enumerate(series):
        s = pl.series[i]
        if colors:
            s.format.fill.solid(); s.format.fill.fore_color.rgb = colors[i]
        s.invert_if_negative = False
    if point_colors:
        for i, c in enumerate(point_colors):
            pt = pl.series[0].points[i]
            pt.format.fill.solid(); pt.format.fill.fore_color.rgb = c
            dpt = pl.series[0]._element.get_or_add_dPt_for_point(i)   # PowerPoint paints negative bars white unless each point says no
            if dpt.find(qn("c:invertIfNegative")) is None:
                el = OxmlElement("c:invertIfNegative"); el.set("val", "0")
                dpt.find(qn("c:idx")).addnext(el)
    if texts:                                      # fixed label text: dot decimals on any machine, blanks for small parts
        for i, tl in enumerate(texts):
            for j, t in enumerate(tl):
                tf = pl.series[i].points[j].data_label.text_frame
                tf.text = t
                if tf.paragraphs[0].runs:
                    f = tf.paragraphs[0].runs[0].font
                    f.size, f.bold, f.name = Pt(size + (0 if stacked else 1)), True, FONT
                    f.color.rgb = WHITE if stacked else INK
    allv = [sum(t) for t in zip(*[v for _, v in series])] if stacked else [a for _, v in series for a in v]
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.maximum_scale = max(allv) * 1.25
    va.minimum_scale = min(0, min(allv) * 1.35)
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(size), FONT
    ca.tick_labels.font.color.rgb = MUTED
    ca.format.line.color.rgb = BORDER
    if horizontal:
        ca.reverse_order = True
    if va.minimum_scale < 0:
        from pptx.enum.chart import XL_TICK_LABEL_POSITION
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch


def line(slide, x, y, w, h, cats, series, title, colors, fmt='#,##0', size=12, labels_last=True):
    cd = CategoryChartData()
    cd.categories = cats
    for n, v in series:
        cd.add_series(n, v)
    ch = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    _style_title(ch, title)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(size), FONT
    for i, s in enumerate(ch.plots[0].series):
        s.format.line.color.rgb = colors[i]
        s.format.line.width = Pt(2.5 if i == 0 else 1.75)
        s.smooth = False
        s.marker.format.fill.solid(); s.marker.format.fill.fore_color.rgb = colors[i]
        s.marker.format.line.color.rgb = colors[i]
        if i > 0:
            from pptx.enum.chart import XL_MARKER_STYLE
            s.marker.style = XL_MARKER_STYLE.NONE
        if i == 0 and labels_last:
            for j in (0, len(cats) - 1):
                dl = s.points[j].data_label
                dl.has_text_frame = False
                dl.show_value = True
                dl.number_format, dl.number_format_is_linked = fmt, False
                dl.position = XL_LABEL_POSITION.ABOVE
                dl.font.size, dl.font.bold, dl.font.name = Pt(size), True, FONT
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = BORDER
    va.tick_labels.font.size, va.tick_labels.font.name = Pt(size - 1), FONT
    va.tick_labels.font.color.rgb = MUTED
    va.tick_labels.number_format, va.tick_labels.number_format_is_linked = fmt, False
    va.format.line.fill.background()
    va.minimum_scale = 0
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(size - 1), FONT
    ca.tick_labels.font.color.rgb = MUTED
    ca.format.line.color.rgb = BORDER
    return ch


def combo(slide, x, y, w, h, cats, bars, lines, title, size=10):
    """Clustered columns plus lines on the same axis. bars/lines: list of (name, values, colour[, dashed])."""
    cd = CategoryChartData()
    cd.categories = cats
    for n, v, *_ in bars + lines:
        cd.add_series(n, v)
    ch = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    _style_title(ch, title)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(size + 1), FONT
    pl = ch.plots[0]
    pl.gap_width, pl.overlap = 60, -10
    for i, (n, v, c, *_) in enumerate(bars):
        s = pl.series[i]
        s.format.fill.solid(); s.format.fill.fore_color.rgb = c
        s.invert_if_negative = False
    bc = ch._chartSpace.chart.plotArea.find(qn("c:barChart"))
    sers = bc.findall(qn("c:ser"))[len(bars):]
    lc = parse_xml(f'<c:lineChart {nsdecls("c")}><c:grouping val="standard"/><c:varyColors val="0"/></c:lineChart>')
    for ser, (n, v, c, *dash) in zip(sers, lines):
        bc.remove(ser)
        for tag in ("c:spPr", "c:invertIfNegative"):
            for el in ser.findall(qn(tag)):
                ser.remove(el)
        d = '<a:prstDash val="dash"/>' if dash and dash[0] else ""
        sp = parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:ln w="34925" cap="rnd"><a:solidFill><a:srgbClr val="{c}"/>'
                       f'</a:solidFill>{d}<a:round/></a:ln></c:spPr>')
        mk = parse_xml(f'<c:marker {nsdecls("c")}><c:symbol val="none"/></c:marker>')
        tx = ser.find(qn("c:tx"))
        tx.addnext(sp); sp.addnext(mk)
        ser.append(parse_xml(f'<c:smooth {nsdecls("c")} val="0"/>'))
        lc.append(ser)
    lc.append(parse_xml(f'<c:marker {nsdecls("c")} val="1"/>'))
    for ax in bc.findall(qn("c:axId")):
        lc.append(deepcopy(ax))
    bc.addnext(lc)
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = BORDER
    va.tick_labels.font.size, va.tick_labels.font.name = Pt(size), FONT
    va.tick_labels.font.color.rgb = MUTED
    va.tick_labels.number_format, va.tick_labels.number_format_is_linked = '#,##0', False
    va.format.line.fill.background()
    va.minimum_scale = 0
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(size), FONT
    ca.tick_labels.font.color.rgb = MUTED
    ca.format.line.color.rgb = BORDER
    return ch

def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


MODEL = "Our 2027 financial model (version 13) and analysis/growth_scenarios.py"
ok = lambda m: BLUE if m >= 0.15 else TERRA
B = lambda t: (t, {"bold": True})
N = lambda t: (t, {})
H = lambda t, c=NAVY, sz=16: [(t, {"bold": True, "color": c, "size": sz})]

# ================================================================== 1. Cover: goal + answer
s = new_slide(None, dark=True)
text(s, 0.8, 0.9, 11.5, 0.5, "TEMAN JOURNEY  |  GROWTH PLAN  |  5 OCTOBER 2026", size=14, bold=True, color=BEIGE_L)
text(s, 0.8, 1.5, 11.5, 1.9, "Keep our profit, pay everyone fairly, and get ready to grow", size=40, bold=True, color=WHITE)
text(s, 0.8, 3.55, 11.6, 1.5,
     [[("Our goal: ", {"bold": True, "color": BEIGE_L}),
       N("keep a net margin of at least 15% and pay everyone fairly, at the same time. Then we are ready for investors, "
         "and we can grow into new business units.")],
      [("Our answer: ", {"bold": True, "color": BEIGE_L}),
       N("keep the new prices, sell more EAP (monthly counseling contracts for companies), and hire only when signed work pays for the person.")]],
     size=17, color=WHITE, space=10)
for i, (lab, f) in enumerate([("1  Price", BLUE), ("2  People", MBLUE), ("3  Growth", TERRA)]):
    text(s, 0.8 + i * 2.3, 5.55, 2.1, 0.5, lab, size=14, bold=True, color=WHITE, fill=f, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.8, 6.35, 11.5, 0.4, "Based on our 2027 financial model (version 13). All money is in Rupiah million (Rp M) unless stated.",
     size=12, color=BEIGE_L)
notes(s, "This deck is written to be read without a presenter. Slides 2 and 3 explain the goal. Slide 4 gives the "
         "answer. Slides 5 to 19 show the problem and the plan in three parts: price, people and growth. Slides 20 to 25 "
         "hold the detailed numbers, a word list, the risks, the model checks and our estimates.")

# ================================================================== 2. The goal in simple words
s = new_slide("Our goal has two parts, and we must reach both at the same time",
              source="Industry net margin: listed education and training companies, about 5% (FullRatio). Jakarta minimum wage 2026: Jakarta Government, Dec 2025.",
              tag="goal")
cols = [
    ("Part 1: a net margin of at least 15%", NAVY,
     [[B("What it means: "), N("from every Rp100 a client pays us, Rp15 stays with us after we pay all costs, all salaries and tax.")],
      [B("Why we need it:")],
      N("Safety. Some months have no big events. Profit from good months pays salaries in quiet months (slide 3)."),
      N("Money to grow. Profit pays for new staff, Instagram and new units, without loans."),
      N("Trust. Large training firms keep about 5%. We are small: one lost client hurts us more, so we need a bigger cushion.")]),
    ("Part 2: fair pay for everyone", BLUE,
     [[B("What it means: "), N(f"full-time staff earn at least the Jakarta minimum wage (Rp5.73M a month), plus BPJS "
                               "(health and work insurance) and THR (holiday bonus). Part-time pay is already fair.")],
      [B("Why we need it:")],
      N("The law. The minimum wage is the legal floor for employees."),
      N("Good people stay. A new person needs time to learn, and clients notice."),
      N("Better care. Tired, underpaid staff give weaker support to clients."),
      N("Investors check it. Profit that depends on low pay is not real profit.")]),
    ("What both parts give us", TERRA,
     [[B("Ready for investors: "), N("we can show a profit that is real and fair.")],
      [B("Ready to grow: "), N("the same team, system and cash can start new business units, for example the Theta app and Coreitera assessments (slide 18).")],
      [B("Our test: "), N("fair pay is a must, because it is the law. Every other idea must keep the year's margin at 15% or more. If it does not, we wait.")]]),
]
for i, (h, f, b) in enumerate(cols):
    x = 0.5 + i * 4.15
    text(s, x, 1.45, 3.95, 0.7, h, size=15, bold=True, color=WHITE, fill=f, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
    card(s, x, 2.25, 3.95, 4.15, b, size=13, space=6)
notes(s, "Read this slide first if business words are new to you. Net margin measures how much of our sales we keep. "
         "Fair pay measures how we treat our people. Fair pay is not optional; the margin target guides every other choice.")

# ================================================================== 3. Why 15%: quiet months
s = new_slide(f"Why 15%: we lose money in {LOSS_MONTHS} of 12 months, so the good months must carry the whole year",
              source="Model version 14, recalculated in Excel: the plan on slide 19 (new work from March, fair pay, founders paid from May).", tag="goal")
vals = [round(v, 1) for v in NET27]
bar(s, 0.5, 1.45, 8.0, 5.45, M27, [("Net profit", vals)], "Net profit each month in 2027 (Rp M)",
    point_colors=[BLUE if v >= 0 else TERRA for v in vals], fmt='#,##0', gap=40)
worst = min(vals)
card(s, 8.8, 1.45, 4.05, 5.45,
     [H("What this chart shows"),
      [B("Red = a month with a loss. "), N("Salaries and office costs come every month. Big programs do not.")],
      [B(f"The worst months lose about {rp(-worst)}. "), N("January and February have almost no events.")],
      [B("Good months must pay for bad ones. "), N(f"A 15% margin on {rp(F27['R'])} of sales gives "
                                                  f"{rp(0.15 * F27['R'])} a year to cover quiet months and to grow.")],
      [B("Without the cushion, "), N("one lost program in a good month can turn the whole year into a loss.")]],
     size=13, space=8)
notes(s, "This is the simplest reason for a 15% target: our income is uneven, our costs are not. The cushion keeps "
         "salaries safe in quiet months and leaves money to grow.")

# ================================================================== 4. Executive summary
s = new_slide("Our answer: keep the new prices, pay our team fairly, and sell more EAP before we hire more people",
              source=MODEL + ". Full-year figures use one year of the 2027 sales plan.", tag="answer")
card(s, 0.5, 1.45, 12.3, 1.6,
     [N(f"In a normal full year, the new prices give {pc(S1['m'])}, but only because our full-time staff earn less than the minimum "
        f"wage. Fair pay drops it to {pc(S2['m'])}; two more EAP contracts bring it back to {pc(S3['m'])}. In 2027 itself, new work "
        f"starts in March and founder pay in May: margin {pc(F27['m'], 1)}, lowest cash {rp(F27['low'])} in {LOWM}.")],
     size=16, fill=NAVY, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
cards = [
    ("1", "Price: what we charge", NAVY,
     ["Big programs for a regulator's head office: about Rp100M.",
      "The same program for each branch: Rp28.5M to Rp61M a day.",
      "EAP: a monthly fee for each employee.",
      "Every price is at least 3 times our cost."]),
    ("2", "People: who we pay", BLUE,
     [f"Pay 4 full-time staff fairly: {rp(FAIR_MO)} more a month, with Instagram ads.",
      "Hire only when new signed work pays for the person.",
      "Better contracts for psychologists, trainers and the project manager."]),
    ("3", "Growth: new sales", TERRA,
     [f"EAP brings income every month: from {pc(EAP_SHARE_S1)} to {pc(EAP_SHARE_S3)} of sales.",
      "Find a co-founder for sales first.",
      "Instagram builds trust; Theta vouchers bring new clients."]),
]
for i, (n, h, f, b) in enumerate(cards):
    x = 0.5 + i * 4.15
    card(s, x, 3.25, 3.95, 3.1, "")
    circle(s, x + 0.2, 3.4, 0.5, n, fill=f)
    text(s, x + 0.8, 3.4, 3.05, 0.5, h, size=15, bold=True, color=f, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.2, 4.0, 3.6, 2.3, b, size=14, space=6)
text(s, 0.5, 6.45, 12.3, 0.45,
     [[B("EAP (Employee Assistance Program): "), N("a company pays us every month, and we give its staff counseling and wellbeing support.")]],
     size=12, color=MUTED)
notes(s, "The answer comes first, then three groups that do not overlap: what we charge (price), who we pay and how "
         "(people), and where new sales come from (growth). Slides 5 to 19 follow this order.")

# ================================================================== 5. Problem: margin bars + why not prices
s = new_slide("The new prices fix our profit only because our team is underpaid, and a second price rise cannot fix it",
              source=MODEL + ". A normal full year at the 2027 sales plan. Fair pay = minimum wage + 18% for BPJS and THR (estimate).", tag="problem")
cats = ["Old prices", "New prices", "+ fair pay", "+ 2 EAP", "+ co-founder and bank deal"]
vals = [G.OLD["m"], S1["m"], S2["m"], S3["m"], S4["m"]]
bar(s, 0.5, 1.45, 7.3, 5.45, cats, [("Net margin", vals)], "Net margin in a normal full year (blue = 15% or more)",
    point_colors=[ok(v) for v in vals])
card(s, 8.1, 1.45, 4.75, 3.05,
     [H("How to read the bars", NAVY, 15),
      [B("New prices: "), N(f"{pc(S1['m'])}, but staff still earn below the minimum wage.")],
      [B("+ fair pay: "), N(f"{pc(S2['m'])}. Fair pay and Instagram ads cost {rp(12 * FAIR_MO)} a year.")],
      [B("+ 2 EAP: "), N(f"{pc(S3['m'])}. Each EAP contract adds {rp(EAP_GP)} of gross profit a year.")],
      [B("+ co-founder and bank deal: "), N(f"{pc(S4['m'])}.")]],
     size=12, space=5)
card(s, 8.1, 4.6, 4.75, 2.3,
     [H("Why not raise prices again?", TERRA, 15),
      N(f"Without new EAP, every event price would need +{pc(UPLIFT)}: the 3-day program at about Rp116M, the 1-day "
        "regional program at about Rp33M. Regulators have budget limits, and clients may leave.")],
     size=12, space=5)
notes(s, "This answers both investor concerns. Full-time staff are underpaid: this plan fixes it. Part-time staff are "
         "paid enough: we keep their rates and only change how their contracts work (slides 13 and 14). Each EAP "
         "contract adds about 5 to 7 points of margin: " + ", ".join(f"{n} new = {pc(m)}" for n, m in CURVE) + ".")

# ================================================================== 6. Fixed cost structure
s = new_slide(f"Fair pay adds {rp(FAIR_MO)} a month to our fixed costs, and we pay it even in quiet months",
              source="Our 2027 budget (model version 13) and the fair-pay table on slide 22.", tag="problem")
cat = ["Today (2027 budget)", "With fair pay"]
parts = [("Founders", [25.0, 25.0], NAVY), ("Staff", [11.3, sum(G.FAIR.values())], TERRA),
         ("Project manager", [5.0, 5.0], MBLUE), ("Office", [5.8, 5.8], BLUE),
         ("Software and ads", [2.08, 2.08 + G.ADS], BEIGE), ("Accountant and sales travel", [3.5, 3.5], GREY)]
bar(s, 0.5, 1.45, 7.2, 5.45, cat, [(n, v) for n, v, _ in parts], "Fixed cost per month (Rp M)",
    colors=[c for *_, c in parts], fmt='#,##0.0', stacked=True, legend=True, gap=80,
    texts=[[f"{a:.1f}" if a >= 4 else "" for a in v] for _, v, _ in parts])
card(s, 8.0, 1.45, 4.85, 5.45,
     [H("What changes"),
      [B("Fixed cost "), N("= what we pay every month, even with no events: salaries, office, software.")],
      [B(f"Total: {rp(FIX_NOW, 1)} to {rp(FIX_FAIR, 1)} a month. "), N("Staff pay rises the most, shown in red.")],
      [B("Founders stay at Rp25M. "), N("We do not raise founder pay in this plan.")],
      [B("Why it matters: "), N(f"each month we must earn at least {rp(FIX_FAIR, 0)} of gross profit before we keep any profit. "
                                "Gross profit = money left after we pay the people who deliver the work.")]],
     size=13, space=8)
notes(s, "Fixed cost is the base we must cover every month. Fair pay makes this base higher, which is why we need "
         "income that also comes every month: EAP.")

# ================================================================== 7. Real timing: sales, cost and cash
s = new_slide("Real timing: we pitch and sign from October to February, and new work starts in March",
              source="Model version 14, Outlook sheet, recalculated in Excel: no new sales before March 2027, growth plan on, founders paid from May 2027.",
              tag="problem")
combo(s, 0.5, 1.45, 8.4, 5.45, M15,
      [("Sales", [round(v) for v in REC["sales"]], BLUE), ("All costs", [round(v) for v in REC["cost"]], GREY)],
      [("Cash in the bank", [round(v) for v in REC["cash"]], "475571"),
       ("Safety floor (1 month of costs)", [round(v) for v in REC["floor"]], "9B3809", True)],
      "Each month, Rp M: sales and costs (bars), cash (lines)")
card(s, 9.15, 1.45, 3.7, 5.45,
     [H("How to read it", NAVY, 15),
      [B("Grey bar taller than blue: "), N("we spend more than we sell that month.")],
      [B("Oct to Feb: "), N(f"only EAP and counseling bring money. Cash falls from Rp210M to {rp(REC['cash'][4])}.")],
      [B("Mar to Jun: "), N("new programs start, but clients pay 0 to 60 days later, so cash stays low.")],
      [B("Jul to Dec: "), N(f"one tight month in July (costs of a busy month come first), then cash grows to {rp(F27['end'])}.")],
      [B("Use it every month: "), N("put real sales and costs next to these bars. If the dark line moves toward the red line, act early.")]],
     size=12, space=6)
notes(s, "This is the most important change from model version 13. Version 13 assumed 13 business events in October "
         "to December 2026. We now assume that this time is used to pitch and sign, and that new work starts in March 2027.")

# ================================================================== 8. Founder pay start
s = new_slide("Start founder salaries in May 2027: if they start in January, cash almost runs out in April",
              source="Model version 14, recalculated in Excel for each start month. Same plan, only the founder salary start changes.", tag="problem")
labs = [a for a, _ in FSTART]
lows = [RUNS[k]["fy27"]["low"] for _, k in FSTART]
belows = [RUNS[k]["fy27"]["below_floor"] for _, k in FSTART]
bar(s, 0.5, 1.45, 6.1, 4.4, labs, [("Lowest cash", lows)], "Lowest cash in the bank (Rp M)",
    point_colors=[TERRA if v < 100 else BLUE for v in lows], fmt='#,##0', texts=[[f"Rp{v:.0f}M" for v in lows]])
bar(s, 6.75, 1.45, 6.1, 4.4, labs, [("Months", belows)], "Months below the safety floor (Oct 2026 to Dec 2027)",
    point_colors=[TERRA if v > 1 else BLUE for v in belows], fmt='0', texts=[[str(int(v)) for v in belows]])
card(s, 0.5, 6.0, 12.3, 0.9,
     [[B("Founder salaries start in: "), N("May 2027. The founders still get 8 months of pay in 2027 (Rp200M). July is safer (no month "
        "below the floor) but pays 2 months less. If March sales are late, move founder pay to July. Months not paid can be paid later as a bonus.")]],
     size=13, fill=BEIGE_L, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Fair pay for staff starts in January because it is the law. Founder pay is the one cost we can choose to move, "
         "so it is the safety valve for cash in early 2027.")

# ================================================================== 8. Price journey
s = new_slide("Set prices along the client journey: start at the head office, spread to branches, keep the client with EAP",
              source="Price list and inputs in our 2027 model (version 13). The EAP price per employee is an estimate; test it with 2 bank HR heads.", tag="1")
cols = [
    ("1. START", "Regulator head office (BI, OJK, LPS)", NAVY,
     ["3-day trainer program: Rp100M (Rp94M training + Rp6M assessment)",
      "Price close to the client's budget limit (the most they are allowed to pay)",
      "Do not raise the price of an old item. Add a new item, such as an assessment. A new item gets its own budget.",
      "Contract: one yearly agreement with the head office"]),
    ("2. SPREAD", "Their branches (BI regional offices, bank branches)", BLUE,
     ["1-day regional program: Rp28.5M",
      "State bank Rp51M, private bank Rp61M, per group per day",
      "Discount: 5% for 5 or more groups, 10% for 10 or more. Never more.",
      "Contract: yearly agreement + a work order (SPK) for each office"]),
    ("3. KEEP", "EAP: support for the client's staff", TERRA,
     ["Today: Rp15M per contract per month",
      "New: Rp30k per employee per month, at least 500 staff (= Rp15M)",
      "Plus Theta app: Rp5k per employee (paid to Theta)",
      "Contract: 12 months, renews by itself, paid monthly in advance"]),
]
for i, (tg, who, f, b) in enumerate(cols):
    x = 0.5 + i * 4.15
    text(s, x, 1.45, 3.95, 0.95, [[(tg, {"bold": True, "size": 18})], [(who, {"size": 13})]], color=WHITE, fill=f,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE, margin=0.18, space=2)
    card(s, x, 2.5, 3.95, 4.4, b, size=14, space=8)
    if i < 2:
        text(s, x + 3.97, 1.7, 0.16, 0.45, "", fill=BEIGE, shape=MSO_SHAPE.CHEVRON)
notes(s, "Every rupiah of business sales fits one of three steps: start (first program at the head office), spread "
         "(the same program in the branches) or keep (a monthly EAP contract). These are our three business goals.")

# ================================================================== 9. Price benchmarks + 3x rule
s = new_slide("Our prices sit between the government fee guide and big training firms, and most pass our 3 times rule",
              source="Government fee guide: Finance Ministry SBM 2026 (PMK 32/2025), fee budget for a 1-day ToT. Big-firm rate: founders' estimate (2.5 times our old bank price).", tag="1")
pd_cats = ["Government fee guide, 1 day", "Our regional program", "Our 3-day program, per day", "Our state bank price",
           "Our private bank price", "Big training firms (banks)"]
pd_vals = [12.6, 28.5, round(100 / 3, 1), 51.0, 61.0, 87.5]
bar(s, 0.5, 1.45, 6.3, 4.6, pd_cats, [("Rp M per day", pd_vals)], "Price for one training day (Rp M)",
    point_colors=[GREY, BLUE, BLUE, BLUE, BLUE, GREY], fmt='#,##0.0', horizontal=True, gap=45,
    texts=[[f"{v:.1f}" for v in pd_vals]])
mx_cats = ["3-day program", "Regional program", "State bank", "Private bank", "EAP", "Public event", "Counseling"]
mx_vals = [4.0, round(28.5 / 8, 2), 4.25, round(61 / 12, 2), 3.0, 2.5, 2.5]
bar(s, 7.0, 1.45, 5.85, 4.6, mx_cats, [("Times our cost", mx_vals)], "Price divided by our cost (rule: 3 or more)",
    point_colors=[BLUE if v >= 3 else TERRA for v in mx_vals], fmt='0.0"x"', horizontal=True, gap=45,
    texts=[[f"{v:.1f}x" for v in mx_vals]])
card(s, 0.5, 6.15, 12.3, 0.75,
     [[B("Grey = outside prices. Blue = ours. Red = below our rule. "),
       N("Raise public events to Rp24M. Keep the counseling price: it brings new clients and is only 1.5% of sales.")]],
     size=13, fill=BEIGE_L, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "We want to be clearly cheaper than big training firms but never the cheapest. The 3 times rule makes sure "
         "each event also pays its share of salaries and office.")

# ================================================================== 10. Margin per hire step
steps_c = ["New prices, today's pay", "1. Fair pay for staff", "2. + psychologist and 2 EAP", "3. + co-founder and bank deal",
           "4 to 6. Full team, Theta pays CTO", "If we paid the CTO"]
steps_v = [S1["m"], STEP1_M, S3["m"], S4["m"], S4_ADM2["m"], S6["m"]]
s = new_slide("Each hire comes with the work that pays for it, so the margin is back above 15% after step 1",
              source=MODEL + ". Net margin for one full year after each step.", tag="2")
bar(s, 0.5, 1.45, 8.2, 5.45, steps_c, [("Net margin", steps_v)], "Net margin after each hiring step",
    point_colors=[ok(v) for v in steps_v], gap=45)
card(s, 9.0, 1.45, 3.85, 5.45,
     [H("Two red bars, two lessons"),
      [B(f"Step 1 ({pc(STEP1_M)}): "), N("fair pay starts in January, before new EAP. It cannot wait, because it is the law. "
                                         "EAP contract 2 in April brings us back.")],
      [B(f"Last bar ({pc(S6['m'])}): "), N("if Teman Journey paid the CTO, the full team would fall below 15%. So Theta's own income pays the CTO.")],
      [B("The rule: "), N("no signed work, no hire.")]],
     size=13, space=9)
notes(s, "This chart is the hiring plan in one picture. Each blue bar is a step where the new person and the new work "
         "arrive together.")

# ================================================================== 12. Capacity
s = new_slide(f"Our real limit is people: the project manager is over capacity in {len(over_pm)} of 15 months, trainers in {len(over_fac)}",
              source="Model version 13, Volume and Capacity sheets. Assumes one event = one project, and each trainer does at most 2 events a month.", tag="2")
bar(s, 0.5, 1.45, 6.1, 4.7, M15, [("Business events", B2B)], "Business events per month (project manager can run 2)",
    point_colors=[TERRA if e > 2 else BLUE for e in B2B], fmt='0', gap=35, size=10)
bar(s, 6.75, 1.45, 6.1, 4.7, M15, [("Trainers needed", FAC)], "Trainers needed per month (we have 2 today)",
    point_colors=[TERRA if f > 2 else BLUE for f in FAC], fmt='0', gap=35, size=10)
card(s, 0.5, 6.25, 12.3, 0.65,
     [[B("Red = more work than our people can do. "),
       N("Fix: a second project manager on standby (paid per project) and train trainers up to 8. Plan for 11 is too many.")]],
     size=13, fill=BEIGE_L, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "October to December is the busy season in both years. These two people limits, not money, decide how many "
         "programs we can sell.")

# ================================================================== 13. Kraljic
s = new_slide("Treat each type of worker differently, and protect the people who are hard to replace",
              source="Kraljic, P. (1983), Purchasing must become supply management, Harvard Business Review. Team numbers from our 2027 model.", tag="2")
gx, gy, gw, gh = 1.2, 1.5, 7.4, 5.0
qw, qh = gw / 2, gh / 2
quads = [
    (0, 0, "Hard to replace, smaller effect", BLUE,
     [f"Project manager: only 1 person, 1 to 2 projects a month (slide 12)",
      "Trainers: 2 certified today, up to 7 needed in a busy month", "What to do: secure them"]),
    (1, 0, "Hard to replace, big effect", TERRA,
     ["Full-time psychologist", "Lead trainers that regulators trust", "What to do: long-term partnership, fair pay, a career path"]),
    (0, 1, "Easy to replace, smaller effect", MUTED,
     ["Event helpers, design, interns", "Venue and travel (the client pays)", "What to do: keep it simple, pay per day"]),
    (1, 1, "Easy to replace, big effect", NAVY,
     ["Part-time psychologists for EAP and counseling", "Many licensed psychologists in Jakarta",
      "What to do: build a group and pay per session"]),
]
for c, r, h, f, b in quads:
    x, y = gx + c * qw, gy + r * qh
    card(s, x + 0.05, y + 0.05, qw - 0.1, qh - 0.1, [[(h, {"bold": True, "color": f, "size": 14})]] + b,
         size=13, margin=0.15, space=6)
text(s, 0.85 - gh / 2, gy + gh / 2 - 0.2, gh, 0.4, "How hard to replace them", size=12, bold=True, color=MUTED,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, name="RiskAxis").rotation = -90
text(s, gx, gy + gh + 0.02, gw, 0.35, "How much they affect our results and our clients' trust", size=12, bold=True,
     color=MUTED, align=PP_ALIGN.CENTER)
text(s, 8.95, 1.45, 3.9, 5.45,
     [H("What this means"),
      N("This is the Kraljic matrix, a tool buyers use to sort suppliers. Here, the suppliers are our people."),
      [B("Do not cut pay. "), N("The right price is not the lowest price.")],
      [B("Saving is not the goal. "), N("Delivery is 32% of sales, so saving 10% gives only about Rp48M a year.")],
      [B("Keeping people is the goal. "), N("One project manager and two trainers limit how many events we can sell.")],
      [B("Pay part-timers per piece of work, "), N("so quiet months do not burn cash.")]],
     size=13, space=8)
notes(s, "The Kraljic matrix sorts suppliers by two questions: how hard is it to replace them, and how much do they "
         "affect our results? Here the suppliers are our own people.")

# ================================================================== 14. Contracts
s = new_slide("Improve every contract in four areas: pay, time promised, quality checks and risk",
              source="Session fees and day rates are estimates; compare them with the market before you sign.", tag="2")
rows = [["Area", "Psychologists (full-time and part-time)", "Trainers", "Project manager"],
        ["1. How we pay", "Full-time: market salary + bonus when an EAP client renews. Part-time: fee per session, higher for senior licence",
         "Day rate by level, extra for the 3-day program; client pays travel; trainer repays training cost if they leave before 4 events",
         "Small monthly fee to stay available + fee per project (bigger project, bigger fee)"],
        ["2. What they promise", "Part-time: at least 8 hours a month on Theta; reply within 24 hours; 3 months notice before leaving",
         "Yearly agreement; at most 2 events a month; confirm 30 days before; we get the first call",
         "Up to 2 projects a month; we get the first call; a handover file for each project"],
        ["3. How we check quality", "Client rating 4.5 of 5 or higher; few missed sessions; reports on time; review every 3 months",
         "Participant scores; test results before and after the training",
         "On time; invoice sent before BI's payment date; client rating; on budget"],
        ["4. How we reduce risk", "Keep client data private (data protection law); senior supervision; insurance; no direct work with our clients for 12 months",
         "Teman Journey owns the training material; trainers use it only for us; a backup trainer for each event",
         "A written step-by-step guide; a second project manager joins one project every 3 months to learn"]]
table(s, 0.5, 1.45, 12.3, [2.0, 3.5, 3.5, 3.3], rows, size=12, row_h=0.88)
card(s, 0.5, 6.1, 12.3, 0.75,
     [[("Why psychologists will join us: ", {"bold": True, "color": NAVY}),
       N("a minimum number of paid hours, supervision and SKP points (professional credits), work with BI and OJK, and flexible hours on Theta.")]],
     size=13, fill=BEIGE_L, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "The four areas cover any service contract without overlap: how we pay, how much time they promise, how we "
         "check their work, and how we protect ourselves. We keep part-time pay levels; we change the structure.")

# ================================================================== 15. Revenue mix
s = new_slide(f"Growth also makes us safer: BI falls from {pc(BI_S1)} to {pc(BI_S4)} of our sales",
              source=MODEL + ". 'With growth' = one full year with 2 more EAP contracts and 10 bank training days.", tag="3")
mix = [("BI (Bank Indonesia)", [BI, BI], NAVY), ("Other regulators and banks", [OTHER1, OTHER1 + 510.0], MBLUE),
       ("EAP (monthly)", [EAP_REV1, 3 * EAP_REV1], TERRA), ("Public events and counseling", [B2C_REV, B2C_REV], GREY)]
bar(s, 0.5, 1.45, 7.2, 5.45, ["2027 plan", "With growth"], [(n, [round(a) for a in v]) for n, v, _ in mix],
    "Sales for one year by type of client (Rp M)", colors=[c for *_, c in mix], fmt='#,##0', stacked=True, legend=True, gap=80)
card(s, 8.0, 1.45, 4.85, 5.45,
     [H("Why this matters"),
      [B("One big client is a risk. "), N(f"Today BI gives {pc(BI_S1)} of our sales. If BI cuts its budget, we feel it at once.")],
      [B(f"With growth, BI is {pc(BI_S4)}. "), N("Our target is below 40%.")],
      [B(f"Monthly income rises from {pc(EAP_SHARE_S1)} to {pc(EAP_SHARE_S4)}. "), N("EAP money comes every month, also in quiet months.")],
      [B("Investors like both: "), N("many clients, and income that repeats.")]],
     size=13, space=9)
notes(s, "BI stays the same size in rupiah; the other parts grow around it. That is the safest way to reduce dependence "
         "on one client: grow others, do not shrink BI.")

# ================================================================== 16. Growth loop
s = new_slide("Grow in one loop: win the head office, spread to its branches, and keep the client with EAP",
              source=MODEL + ".", tag="3")
loop = [("Instagram", "Builds trust: 3 posts a week, client stories (with permission)", NAVY),
        ("Start", "A big program at the regulator head office", BLUE),
        ("Spread", "The same program in each branch", MBLUE),
        ("Keep", "EAP and Theta vouchers for staff", TERRA),
        ("New clients", "Vouchers bring individual counseling and referrals", RED)]
for i, (h, b, f) in enumerate(loop):
    x = 0.5 + i * 2.5
    text(s, x, 1.45, 2.2, 1.45, [[(h, {"bold": True, "size": 16})], [(b, {"size": 12})]], color=WHITE, fill=f,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE, margin=0.12, space=3, align=PP_ALIGN.CENTER)
    if i < 4:
        text(s, x + 2.22, 2.0, 0.26, 0.35, "", fill=BEIGE, shape=MSO_SHAPE.RIGHT_ARROW)
text(s, 0.5, 2.95, 12.2, 0.35, "New clients and their stories go back to Instagram, and the loop starts again.",
     size=12, color=MUTED, align=PP_ALIGN.CENTER)
boxes = [
    ("Instagram targets", [
        "Reach that grows every month",
        "Messages that turn into counseling bookings",
        "1 client story every 3 months",
        f"Cost: about {rp(MKT_MO)} a month for marketing staff and ads",
        "Set exact numbers after 1 month of data"]),
    ("How we deliver EAP", [
        "Our full-time psychologist manages each client",
        "Part-time psychologists run sessions, paid per hour",
        "A report for the client every 3 months",
        "First check: is our current EAP contract still active?"]),
    ("Theta vouchers: free new clients", [
        "Each EAP client's staff get Theta vouchers",
        "Staff who like the service book more sessions on their own",
        "We pay nothing to find these clients",
        "Track: how many vouchers turn into paid sessions"]),
]
for i, (h, b) in enumerate(boxes):
    x = 0.5 + i * 4.15
    card(s, x, 3.4, 3.95, 3.5, [[(h, {"bold": True, "color": NAVY, "size": 15})]] + b, size=13, space=7)
notes(s, "The loop joins our three business goals: enter the regulator, spread across its branches, and keep it with "
         "EAP. Theta vouchers bring new clients at the 'keep' step; Instagram makes us look professional to all of them.")

# ================================================================== 17. Co-founder
s = new_slide("Choose a sales co-founder first: 3 bank training days pay their cost, 10 days pay it more than 3 times",
              source="Model version 13. Profit per bank day = Rp51M price minus delivery and project manager cost. Check share terms with a lawyer.", tag="3")
rows = [["Question", "Sales co-founder", "Marketing co-founder"],
        ["Does it fix our biggest limit?", "Yes: the founders do all sales, and one founder studies in Singapore", "No: marketing does not close regulator deals"],
        ["How do our clients buy?", "Regulators and banks buy through relationships and trust", "Ads help the public side (counseling) most"],
        ["Can we hire it as staff instead?", "Hard: good sellers to regulators are rare", f"Yes: marketing staff + ads cost about {rp(MKT_MO)} a month"],
        ["Our choice", "First", "Later, or hire as staff"]]
table(s, 0.5, 1.45, 7.0, [2.2, 2.6, 2.2], rows, size=12, row_h=0.8,
      fills={(4, 1): BLUE_BG, (4, 2): RED_BG})
bar(s, 7.8, 1.45, 5.05, 3.6, ["Co-founder cost", "Profit, 3 bank days", "Profit, 10 bank days"],
    [("Rp M a year", [12 * G.COF, 3 * GROUP_GP, 10 * GROUP_GP])], "One year (Rp M)",
    point_colors=[TERRA, BLUE, BLUE], fmt='#,##0', gap=50)
card(s, 7.8, 5.15, 5.05, 1.75,
     [H("The deal", NAVY, 14),
      N(f"Salary about {rp(G.COF)} a month + shares earned over 4 years (nothing if they leave in year 1)."),
      N("Year-1 goals: 1 bank agreement, first OJK or LPS program, 2 EAP renewals.")],
     size=12, space=4, margin=0.15)
text(s, 0.5, 5.6, 7.0, 1.3,
     [[B("Co-founder "), N("= a partner who owns part of the company. Shares are earned over time, so a partner who leaves early does not keep them.")]],
     size=13, color=MUTED)
notes(s, "A sales co-founder fixes the limit that matters most today: who sells to regulators and banks. Marketing can "
         "be hired as staff for less.")

# ================================================================== 18. Ready to scale
s = new_slide("We are ready to invest in new business units once our cash passes 6 months of fixed costs",
              source="Model version 14, Outlook sheet, recalculated in Excel (real timing, founders paid from May 2027).", tag="3")
line(s, 0.5, 1.45, 7.6, 5.45, M15, [("Cash in the bank (plan)", [round(v) for v in REC["cash"]]),
                                     ("Safety cushion: 6 months of fixed cost", [round(BUFFER)] * 15)],
     "Cash at the end of each month (Rp M)", [NAVY, TERRA])
when = "at the end of 2027" if F27["end"] >= BUFFER else "in early 2028"
card(s, 8.4, 1.45, 4.45, 5.45,
     [H("Our rule for growth money"),
      [B("1. Keep a cushion. "), N(f"6 months of fixed cost, about {rp(BUFFER)}, stays in the bank for bad times.")],
      [B("2. Invest above it. "), N("Up to half of each year's net profit can go to new business units, such as the Theta app and Coreitera.")],
      [B("When: "), N(f"{when}. Cash is {rp(F27['end'])} at the end of 2027.")],
      [B("How much: "), N(f"a normal full year with this plan makes about {rp(S4['N'])} net profit, so about {rp(INVEST_28)} a year for new units.")]],
     size=13, space=9)
notes(s, "This is what 'ready to grow' means in numbers. We do not invest the cushion. We invest only what we earn above it.")

# ================================================================== 19. Roadmap
s = new_slide("Pitch and sign now, deliver from March, and take three decisions this month",
              source="Model version 14, recalculated in Excel. Each hire starts in its trigger month; cash includes client payment delays.", tag="next")
phases = [("Oct 2026 to Feb 2027", "Pitch and sign", NAVY, ["Pitch BI 2027 programs before budgets close, at the new prices", "Pitch EAP to 2 banks, LPS and OJK; sign for delivery from March", "Fair pay for 3 full-time staff from January"]),
          ("Mar to Jun 2027", "Deliver", BLUE, ["First new programs in March", "EAP 2 starts in April; hire the full-time psychologist", "Founder salaries start in May"]),
          ("Jul to Sep 2027", "Grow", MBLUE, ["EAP 3 starts; sales co-founder joins", "Build a group of 4 to 6 part-time psychologists", "Train trainers up to 8, not 11"]),
          ("Oct to Dec 2027", "Scale", TERRA, ["First bank agreement: 1 training day a month", "Busy season: second project manager on standby",
                                               f"Theta pays {pc(3 * THETA_MO / G.CTO)} of the CTO salary"])]
for i, (q, h, f, b) in enumerate(phases):
    x = 0.5 + i * 3.1
    text(s, x, 1.45, 2.95, 0.75, [[(q, {"bold": True, "size": 14})], [(h, {"size": 12})]], color=WHITE, fill=f,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE, margin=0.15, space=0)
    card(s, x, 2.3, 2.95, 1.95, b, size=12, margin=0.15, space=6)
text(s, 0.5, 4.45, 6.6, 2.45,
     [H("Three decisions for the founders"),
      [B("1. "), N("Pitch now. Goal by February: BI 2027 programs, 2 EAP contracts and 1 bank deal signed.")],
      [B("2. "), N(f"Fair pay from January ({rp(EXIST_DELTA, 1)} a month). Founder salaries from May; July if March sales are late.")],
      [B("3. "), N("Start the co-founder search. Theta, not Teman Journey, pays the CTO.")]],
     size=14, space=7)
text(s, 7.4, 4.45, 5.4, 2.45,
     [[("Year 2027 with this plan", {"bold": True, "color": WHITE, "size": 16})],
      [(f"Net margin {pc(F27['m'], 1)}, with fair pay", {"bold": True, "size": 20, "color": BEIGE_L})],
      N(f"Sales {rp(F27['R'])}, net profit {rp(F27['N'])}. Lowest cash {rp(F27['low'])} ({LOWM}), {rp(F27['end'])} at the end of 2027."),
      N(f"Risks: no bank deal gives {pc(NOBANK['m'], 1)}. BI buying 30% less gives {pc(BICUT['m'], 1)} and cash as low as {rp(BICUT['low'])}.")],
     size=14, color=WHITE, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, margin=0.2, space=5)
notes(s, "The margin in 2027 is high partly because founders are paid for 8 months, not 12. A normal full year with "
         "founders paid all year is the number on slide 5.")

# ================================================================== A1. Scenario tables
s = new_slide("The numbers behind the charts", source=MODEL + "; model version 14 recalculated in Excel. Tax is 0.5% of sales.", tag="app")
text(s, 0.5, 1.0, 10.9, 0.35, "Top: a normal full year, each row adds one change. Bottom: 2027 month by month; there the project manager cost starts later (slide 25).",
     size=12, color=MUTED)
rows = [["A normal full year (2027 sales plan)", "Sales", "Gross profit", "Fixed cost", "Net profit", "Net margin"],
        ["Old prices, today's pay", f"{G.OLD['R']:,.0f}", "", "", f"{G.OLD['N']:,.0f}", pc(G.OLD['m'], 1)]]
plain = ["New prices, today's pay", "+ fair pay, psychologist, Instagram ads", "+ 2 more EAP contracts",
         "+ co-founder, but no new deal", "+ co-founder wins 10 bank training days", "+ Teman Journey pays the CTO",
         "+ second admin"]
for lab, (name, _) in zip(plain, G.LADDER):
    v = L[name]
    rows.append([lab, f"{v['R']:,.0f}", f"{v['GP']:,.0f}", f"{v['F']:,.0f}", f"{v['N']:,.0f}", pc(v['m'], 1)])
rows.append(["Full team, Theta pays the CTO", f"{S4_ADM2['R']:,.0f}", f"{S4_ADM2['GP']:,.0f}", f"{S4_ADM2['F']:,.0f}",
             f"{S4_ADM2['N']:,.0f}", pc(S4_ADM2['m'], 1)])
fills = {(i, 5): (BLUE_BG if float(rows[i][5].rstrip('%')) >= 15 else RED_BG) for i in range(1, len(rows))}
table(s, 0.5, 1.4, 12.3, [5.3, 1.4, 1.4, 1.4, 1.4, 1.4], rows, size=12, row_h=0.3, fills=fills)
rows2 = [["Year 2027 (model version 14)", "Sales", "Net profit", "Net margin", "Lowest cash", "Cash end 2027", "Months below floor"]]
for lab, k in [("Old timing (sales from Oct 2026), today's pay", "v13_timing_today_pay"),
               ("Real timing (new work from Mar 2027), today's pay", "real_timing_today_pay"),
               ("Real timing + growth plan, founders paid from Jan", "plan_founders_jan"),
               ("Same, founders paid from May (our plan)", "plan_founders_may"),
               ("Our plan, but no bank deal", "plan_founders_may_no_bank"),
               ("Our plan, but BI buys 30% less", "plan_founders_may_bi_cut")]:
    f = RUNS[k]["fy27"]
    rows2.append([lab, f"{f['R']:,.0f}", f"{f['N']:,.0f}", pc(f['m'], 1),
                  f"{f['low']:,.0f} ({G.MONTHS[f['low_idx']].replace('-', ' ')})", f"{f['end']:,.0f}", str(f['below_floor'])])
table(s, 0.5, 4.55, 12.3, [4.6, 1.05, 1.15, 1.15, 1.6, 1.35, 1.4], rows2, size=12, row_h=0.33, header_fill=BLUE)

# ================================================================== 11. Hire ladder table
s = new_slide("Hire in a fixed order, only when signed work pays for the person",
              source=MODEL + ". Salaries for the co-founder, CTO and second admin are estimates. Net margin is for one full year.", tag="app")
text(s, 0.5, 1.0, 10.9, 0.4, "Trigger = the event that tells us it is time to hire. No trigger, no hire.", size=13, color=MUTED)
rows = [["Step", "Role", "Extra cost a month", "Trigger: hire when...", "What pays for it", "Net margin after"],
        ["1", "Fair pay for admin, social media and marketing lead, plus Instagram ads", f"+{rp(EXIST_DELTA, 1)}",
         "January 2027, when the new prices start", "The new prices", f"{pc(STEP1_M)} (before more EAP)"],
        ["2", "Full-time psychologist (replaces the part-time PIC)", f"+{rp(PSY_DELTA, 1)}", "The second EAP contract is signed",
         "Less than 1 EAP contract", f"{pc(S3['m'])} with 2 more EAP"],
        ["3", "Sales co-founder (salary + shares)", f"+{rp(G.COF)}", "2 EAP contracts running and a bank, OJK or LPS offer is open",
         "3 bank training days a year", f"{pc(S4['m'])} with 10 bank days"],
        ["4", "CTO for the Theta app", f"+{rp(G.CTO)}", "Theta's own income covers the salary", "4 EAP contracts with Theta",
         "Theta pays, not Teman Journey"],
        ["5", "Second project manager (paid per project)", "Per project", "More than 2 projects a month, 2 months in a row",
         "The fee of each project", "No change"],
        ["6", "Second admin", f"+{rp(G.ADM2, 1)}", "Over 12 invoices or 6 business events a month, for 3 months",
         "Until then, a daily freelancer", f"{pc(S4_ADM2['m'])} with the full team"]]
table(s, 0.5, 1.45, 12.3, [0.6, 3.2, 1.4, 3.0, 2.2, 1.9], rows, size=12, row_h=0.66)
card(s, 0.5, 6.2, 12.3, 0.7,
     [[B("PIC = person in charge. "), N("Today a part-time psychologist leads our clinical work; step 2 makes this a full-time role.")]],
     size=13, fill=BEIGE_L, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Sales roles start on a date when deals are open. Delivery roles start when the work stays high. Steps 5 and 6 "
         "stay flexible until the work proves them.")

# ================================================================== A2. Fair pay
s = new_slide("Fair pay: what each full-time role costs at market pay",
              source="Our 2027 budget; Jakarta minimum wage 2026 Rp5,729,876; psychologist pay Rp7M to Rp13M (JobStreet); BPJS + THR 18% (estimate).", tag="app")
rows = [["Role", "Pay today (Rp M a month)", "Fair pay", "Fair pay + BPJS and THR", "Increase", "Based on"]]
basis = {"admin": "Jakarta minimum wage", "socmed": "Jakarta minimum wage", "leadmkt": "Minimum wage + 22% for a lead role",
         "psych": "Psychologist market pay (low end)", "interns": "Intern allowance, no change"}
label = {"admin": "Admin", "socmed": "Social media staff", "leadmkt": "Marketing lead",
         "psych": "Full-time psychologist (was part-time)", "interns": "2 interns (marketing, operations)"}
gross = {"admin": G.UMP, "socmed": G.UMP, "leadmkt": 7.0, "psych": 9.0, "interns": 2.0}
for k in G.CUR:
    rows.append([label[k], f"{G.CUR[k]:.1f}", f"{gross[k]:.1f}", f"{G.FAIR[k]:.1f}", f"+{G.FAIR[k] - G.CUR[k]:.1f}", basis[k]])
rows.append(["More Instagram ads and content", "", "", f"{G.ADS:.1f}", f"+{G.ADS:.1f}", "Founders' budget choice"])
rows.append(["Total", f"{sum(G.CUR.values()):.1f}", "", f"{sum(G.FAIR.values()) + G.ADS:.1f}", f"+{FAIR_MO:.1f}", ""])
table(s, 0.5, 1.45, 12.3, [3.6, 1.8, 1.3, 1.9, 1.2, 2.5], rows, size=12, row_h=0.5)
card(s, 0.5, 5.6, 12.3, 1.1,
     [[B("In short: "),
       N(f"fixed costs rise from {rp(FIX_NOW, 1)} to {rp(FIX_FAIR, 1)} a month (founders included). To keep 15% profit, we need "
         f"about {rp(FIX_FAIR / 0.525 * 12)} of sales a year. Our plan has {rp(S1['R'])}. EAP and the bank deal must close this gap.")]],
     size=14, fill=BEIGE_L, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================== A3. Word list
s = new_slide("Word list: the business words in this deck, in simple terms", tag="app",
              source="Example numbers use one EAP contract: the client pays Rp15M a month and the work costs us Rp5M.")
words_l = [["Word", "What it means"],
           ["Sales (revenue)", "All the money clients pay us"],
           ["Delivery cost", "What one job costs us: psychologist and trainer fees, materials"],
           ["Gross profit", "Sales minus delivery cost. One EAP contract: 15 - 5 = Rp10M a month"],
           ["Fixed cost", "What we pay every month, even with no work: salaries, office, software"],
           ["Net profit", "Gross profit minus fixed cost and tax: what we really keep"],
           ["Net margin", "Net profit as a share of sales. 15% = we keep Rp15 of every Rp100"],
           ["Break-even", "The point where income covers a cost exactly; after it, we make profit"],
           ["Cash", "Money in the bank today. Profit is not cash until the client pays"],
           ["Cash cushion", "Money we keep for bad times: 6 months of fixed cost"]]
words_r = [["Word", "What it means"],
           ["Fair pay", "At least the minimum wage, plus BPJS and THR, for full-time staff"],
           ["UMP, BPJS, THR", "Minimum wage; state health and work insurance; holiday bonus"],
           ["EAP", "A company pays us monthly; we support its staff with counseling"],
           ["Recurring income", "Money that comes every month without a new sale, like EAP"],
           ["Trigger", "The event that tells us it is time to hire"],
           ["Co-founder, shares", "A partner who owns part of the company; shares are earned over years"],
           ["Business unit", "A separate part of the business, for example Theta or Coreitera"],
           ["Framework, SPK", "A yearly agreement with a client; a work order for one job under it"],
           ["Client share", "How much of our sales comes from one client, such as BI"]]
table(s, 0.5, 1.45, 6.0, [1.8, 4.2], words_l, size=12, row_h=0.52)
table(s, 6.8, 1.45, 6.0, [1.8, 4.2], words_r, size=12, row_h=0.52)

# ================================================================== A4. Risks
s = new_slide("Risks in our people supply, and how we respond",
              source="Chance and impact are our own judgement. Review this list every 3 months.", tag="app")
text(s, 0.5, 1.0, 10.9, 0.35, "Ways to respond: reduce the risk, share it (for example with insurance), accept it, or stop the activity.",
     size=12, color=MUTED)
rows = [["Risk", "Chance", "Impact", "Response", "Action", "Who watches, and when to act"],
        ["One project manager has too much work", "High", "High", "Reduce", f"Second project manager on standby; written guide (slide 12)", "Operations lead; 3 or more projects in a month"],
        ["A certified trainer leaves", "Medium", "High", "Reduce", "Yearly agreement; trainer repays training cost; train up to 8", "Galih; fewer trainers than the busiest month needs + 1"],
        ["A client hires our trainer directly", "Medium", "Medium", "Reduce", "No direct work with our clients for 12 months; we own the material", "Co-founder; any direct approach"],
        ["A clinical or ethics problem", "Low", "High", "Share + reduce", "Insurance; senior supervision; clear steps to escalate", "Full-time psychologist; any case"],
        ["Staff data leaks (EAP, Theta)", "Low", "High", "Reduce", "Consent, access control, data agreement (data protection law)", "CTO; check every 3 months"],
        ["BI cuts its budget", "Medium", "High", "Reduce", f"Grow EAP and banks: BI share of sales from {pc(BI_S1)} to {pc(BI_S4)}", "Founders; BI below plan for 2 months"]]
table(s, 0.5, 1.45, 12.3, [2.6, 1.0, 0.9, 1.4, 3.9, 2.5], rows, size=12, row_h=0.68)
text(s, 0.5, 6.3, 12.3, 0.6,
     [[B("Bad-year check: "), N("in the model's worst case (no new business January to June, BI buys 30% less), margin is 8% with "
                                "today's pay. Do not start hiring step 3 or later while this risk is real.")]], size=12)

# ================================================================== A6. Model checks + estimates
s = new_slide("What v14 fixed, what is still open, and the estimates to check", tag="app",
              source="Review of model versions 13 and 14 and the projection guide, 5 October 2026.")
left = [H("What model version 14 fixed", NAVY, 15),
        "Real timing: no new sales before March 2027 (switch in Inputs)",
        "Fair pay, new EAP, co-founder and bank days are inside the model",
        "New Outlook sheet: sales, costs and cash in one chart",
        H("Still open", TERRA, 15),
        "Project manager is a fixed salary in the model; she is paid per project",
        "Theta and CTO costs are in no model",
        "Is our current EAP contract active, and for how many staff?",
        "Will BI accept the new 2027 prices?",
        "Tax: does the 0.5% small-business tax still apply as sales grow?"]
card(s, 0.5, 1.45, 4.6, 5.45, left, size=12, space=5)
rows = [["Estimate", "Value we used", "How to check"],
        ["BPJS and THR on top of salary", "18% of salary", "Payroll provider quote"],
        ["Full-time psychologist salary", "Rp9M a month", "3 job ads or offers"],
        ["EAP price per employee", "Rp30k a month, 500 staff", "Ask 2 bank HR heads"],
        ["Theta access per employee", "Rp5k a month", "Same 2 bank HR heads"],
        ["Sales co-founder salary", "Rp8M a month + shares", "Talks with candidates"],
        ["CTO salary", "Rp10M a month", "CTO's expectation"],
        ["Bank deal", "1 training day a month from Oct 2027", "First bank proposal"],
        ["Bank payment timing", "20% / 50% / 30% over 60 days", "First bank contract"],
        ["New work starts", "March 2027", "Signed contracts by February"],
        ["Income tax", "0.5% of sales", "Tax consultant"]]
table(s, 5.3, 1.45, 7.5, [2.7, 2.6, 2.2], rows, size=12, row_h=0.49)

assert len(prs.slides) == 25, len(prs.slides)
prs.save(OUT)
print("wrote", OUT, len(prs.slides), "slides")
