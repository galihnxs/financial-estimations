"""
Teman Journey - growth recommendation deck (40 slides incl. 7 dark, question-led section openers): pyramid answer first, then A. B2B, B. B2C, C. people, D. money, plain-English version for readers whose
first language is not English and who are not business specialists. Chart-led, one short headline per slide.
Look: the Coreitera/Teman Journey design pattern (warm-white pages, soft cards, big numbers, one accent phrase).
Estimates are named conservative / current / aggressive (no model version numbers on slides).

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
HEAD = "Georgia"                                   # headlines and big numbers (Roca One / Playfair are not installed)
# Design pattern (Coreitera/Teman Journey design skill): warm-white pages, sky-tint cards, cream callouts, navy for dark
# slides, one accent phrase per headline. Colours stay with temanjourney.id: slate blue for data, terracotta as the accent.
BG, INK_D, BODY, MUTED = rgb("FAF7F2"), rgb("1E2A3A"), rgb("3D4F6A"), rgb("6B7A99")
SKY, CREAM, SAND, ZEBRA, LINE = rgb("EAF2FF"), rgb("F5EDD8"), rgb("EEE9DF"), rgb("F6F2EA"), rgb("E3DED3")
AMBER, DARK2, WHITE = rgb("E8904A"), rgb("2A3A52"), rgb("FFFFFF")
NAVY, BLUE, LBLUE = INK_D, rgb("576C94"), rgb("B4C2DC")
TERRA, RED, BEIGE, GREY = rgb("9B3809"), rgb("CD5D44"), rgb("C9B78D"), rgb("C3C9D4")
BLUE_BG, RED_BG = rgb("DDE6F5"), rgb("F7DDD2")

# ------------------------------------------------------------------ numbers (all from the scenario engine)
L = {name: G.scen(st) for name, st in G.LADDER}
S1, S2, S3, S3B, S4, S5, S6 = (L[n] for n, _ in G.LADDER)
S4_ADM2 = G.scen(["fair", "eap", "cof", "bank", "adm2"])       # full team, CTO paid by Theta
CURVE = G.eap_curve()
UPLIFT, _ = G.price_only_uplift()
RUNS = json.loads((Path(__file__).resolve().parent.parent / "output" / "model_runs.json").read_text())   # excel_runner.py
REC = RUNS["plan"]                                               # our plan: model v15 with the founders' 5 Oct decisions
F27 = REC["fy27"]
NOBANK, BICUT, EAPLATE = RUNS["plan_no_bank"]["fy27"], RUNS["plan_bi_cut"]["fy27"], RUNS["plan_eap_late"]["fy27"]
GUARD = RUNS["plan_both_guards"]["fy27"]
LOWM = G.MONTHS[F27["low_idx"]].replace("-", " 20")
LEVERS = RUNS["plan_all_levers"]["fy27"]                         # guard rule + Flagship down payment + 3 events moved
import openpyxl                                                  # noqa: E402  (values cached by Excel, excel_runner.py --save)
WB = openpyxl.load_workbook(Path(__file__).resolve().parent.parent / "TemanJourney_Model_2027_v15.xlsx", data_only=True)
_cap = WB["Capacity_OPEX"]
FAC15 = [int(_cap.cell(8, c).value or 0) for c in range(5, 20)]      # trainers needed per month
CERT15 = [int(_cap.cell(47, c).value or 0) for c in range(5, 20)]    # trainers certified that month (ahead of need)
PM_CAP = WB["Inputs"]["C115"].value
PM_FEE = WB["Inputs"]["C114"].value
EAP2 = "EAP contract 2: the warm lead (chase now)"
_pp = WB["Pay_Plan"]
ADS_X = _pp["H12"].value                                         # extra IG ads a month from Jan-27
ADM = [_pp.cell(6, c).value for c in range(5, 20)]               # admin pay (per-client scheme, minimum-wage floor from Jul-27)
LP = WB["Inputs"]["C124"].value                                  # Lead Partnership pay (estimate)
PERF = [WB["Inputs"][c].value for c in ("C123", "C126", "C127")]   # sales targets that unlock pay steps 1-3 (Rp M a month)
_step = [int(_pp.cell(30, c).value or 0) for c in range(5, 20)]
STEP_M = [next((G.MONTHS[i].replace("-", " 20") for i, v in enumerate(_step) if v >= k), "after 2027") for k in (1, 2, 3)]
_b = WB["B2C"]
B2C = {k: [_b.cell(r, c).value or 0 for c in range(5, 20)] for k, r in
       dict(fol=20, br=21, spc=22, ig=24, vouch=25, sess=26, rev=31, b2b=32, gp=35, psy=36, ads=37, cov=38, demand=41,
            cap=42, load=43, ft=46, ftneed=47, refsess=52, ltv=56, cac=57, ltvcac_ig=58, ltvcac_ref=59).items()}
SN_ = {"est": 21}                                                # slide numbers used in cross-references
B2C_Q = [[_b.cell(r, c).value for c in range(3, 8)] for r in (6, 7, 8)]          # quarterly inputs
QS = [slice(0, 3), slice(3, 6), slice(6, 9), slice(9, 12), slice(12, 15)]
QN = ["Q4 2026", "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027"]
RM = {WB["Roadmap"].cell(r, 1).value: [WB["Roadmap"].cell(r, c).value for c in range(1, 12)] for r in range(15, 36)}
BELOW = [G.MONTHS[i].replace("-", " 20") for i, (c, f) in enumerate(zip(REC["cash"], REC["floor"])) if c < f]
PHASES = [(0, "Now (Oct 2026)"), (3, "Jan 2027"), (5, "Mar 2027"), (9, "Jul 2027"), (12, "Oct 2027")]
FAIR_MO = G.FAIR_DELTA + G.ADS
FIX_NOW = G.ROWS["F"][3]
FIX_FAIR = FIX_NOW + FAIR_MO
PSY_DELTA = G.FAIR["lp"] - G.CUR["lp"]                           # the new hire (Lead Partnership) is not "existing staff"
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
def text(slide, x, y, w, h, paras, size=14, color=None, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, margin=0.08, space=4, name=None, font=FONT, radius=0.08):
    """paras: str or list of str / (text, opts) / list of runs [(text, {opts})]. opts: size, bold, italic, color, font."""
    color = BODY if color is None else color
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
            sh.adjustments[0] = radius
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
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", False)
            f.color.rgb = o.get("color", color)
    return sh


def card(slide, x, y, w, h, paras, size=13, fill=None, **kw):
    return text(slide, x, y, w, h, paras, size=size, fill=SKY if fill is None else fill,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, margin=kw.pop("margin", 0.2), radius=kw.pop("radius", 0.06), **kw)


def stat(slide, x, y, w, h, value, label, vcolor=None, fill=None, lcolor=None, vsize=30, lsize=12, align=PP_ALIGN.LEFT):
    """Big number on a soft card: the number first, a short plain label under it."""
    return card(slide, x, y, w, h, [[(value, {"font": HEAD, "size": vsize, "bold": True, "color": vcolor or INK_D})],
                                    [(label, {"size": lsize, "color": lcolor or BODY})]],
                fill=fill, anchor=MSO_ANCHOR.MIDDLE, space=2, align=align, margin=0.18)


TAGS = {"b2b": ("A · B2B", INK_D, WHITE), "b2c": ("B · B2C", RED_BG, TERRA), "ppl": ("C · PEOPLE", SKY, BLUE),
        "money": ("D · MONEY", CREAM, TERRA), "goal": ("OUR GOAL", RED_BG, TERRA), "answer": ("OUR ANSWER", INK_D, WHITE), "problem": ("THE PROBLEM", RED_BG, TERRA),
        "1": ("1 · PRICE", SKY, BLUE), "2": ("2 · PEOPLE", SKY, BLUE), "3": ("3 · GROWTH", SKY, BLUE),
        "next": ("NEXT STEPS", INK_D, WHITE), "app": ("APPENDIX", SAND, MUTED)}


def new_slide(title, sub=None, source=None, dark=False, tag=None, accent=None):
    """Light slide: warm-white background, one-line serif headline with one accent phrase, a plain subtitle, a chip."""
    s = prs.slides.add_slide(BLANK)
    PAGE[0] += 1
    bg = s.background.fill
    bg.solid(); bg.fore_color.rgb = INK_D if dark else BG
    if title:
        tc, ac = (WHITE, AMBER) if dark else (INK_D, TERRA)
        runs = [(title, {})]
        if accent and accent in title:
            a, b = title.split(accent, 1)
            runs = [(a, {}), (accent, {"color": ac}), (b, {})]
        text(s, 0.5, 0.36, 10.8, 0.62, [runs], size=28, bold=True, color=tc, font=HEAD, name="Title",
             anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    if sub:
        text(s, 0.5, 0.96, 10.8, 0.4, sub, size=14, color=rgb("C9D2E3") if dark else BODY, name="Subtitle", margin=0.02)
    if tag:
        lab, f, c = TAGS[tag]
        text(s, 11.45, 0.46, 1.4, 0.34, lab, size=10, bold=True, color=c, fill=f, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, margin=0.02, name="Tag", radius=0.5)
    if source:
        text(s, 0.5, 7.02, 11.6, 0.32, "Source: " + source, size=9, color=MUTED, name="Source")
    text(s, 12.3, 7.02, 0.55, 0.32, str(PAGE[0]), size=10, bold=True, color=WHITE if dark else MUTED,
         align=PP_ALIGN.RIGHT, name="Page")
    return s


def circle(slide, x, y, d, label, fill=INK_D, size=16):
    return text(slide, x, y, d, d, label, size=size, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font=HEAD,
                anchor=MSO_ANCHOR.MIDDLE, fill=fill, shape=MSO_SHAPE.OVAL, margin=0)


def callout(slide, x, y, w, h, paras, size=13, **kw):
    return card(slide, x, y, w, h, paras, size=size, fill=CREAM, anchor=kw.pop("anchor", MSO_ANCHOR.MIDDLE), **kw)


def table(slide, x, y, w, col_w, rows, size=12, header_fill=None, row_h=0.36, first_col_bold=True, fills=None):
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr)).table
    for j, cw in enumerate(col_w):
        gt.columns[j].width = Inches(cw)
    for i, r in enumerate(rows):
        gt.rows[i].height = Inches(row_h)
        for j, v in enumerate(r):
            c = gt.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.08)
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
                c.fill.fore_color.rgb = header_fill or INK_D
                f.bold, f.color.rgb = True, WHITE
            else:
                c.fill.fore_color.rgb = (fills or {}).get((i, j), WHITE if i % 2 else ZEBRA)
                f.color.rgb = INK_D if (first_col_bold and j == 0) else BODY
                f.bold = first_col_bold and j == 0
    return gt


def _style_title(ch, title):
    ch.has_title = True
    ch.chart_title.text_frame.text = title
    tp = ch.chart_title.text_frame.paragraphs[0].runs[0].font
    tp.size, tp.bold, tp.name, tp.color.rgb = Pt(12), True, FONT, INK_D


def _chart_bg(ch):                                   # transparent chart area so the warm-white slide shows through
    cs = ch._chartSpace
    sp = cs.find(qn("c:spPr"))
    if sp is None:
        sp = parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')
        cs.find(qn("c:chart")).addnext(sp)


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
    _chart_bg(ch)
    _style_title(ch, title)
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size, ch.legend.font.name = Pt(size), FONT
        ch.legend.font.color.rgb = BODY
    pl = ch.plots[0]
    pl.gap_width = gap
    if stacked:
        pl.overlap = 100
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format, dl.number_format_is_linked = fmt, False
    dl.position = XL_LABEL_POSITION.CENTER if stacked else XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.name, dl.font.bold = Pt(size + (0 if stacked else 1)), FONT, True
    dl.font.color.rgb = WHITE if stacked else INK_D
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
                    f.color.rgb = WHITE if stacked else INK_D
    allv = [sum(t) for t in zip(*[v for _, v in series])] if stacked else [a for _, v in series for a in v]
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.maximum_scale = max(allv) * 1.25
    va.minimum_scale = min(0, min(allv) * 1.35)
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(size), FONT
    ca.tick_labels.font.color.rgb = BODY
    ca.format.line.color.rgb = LINE
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
    _chart_bg(ch)
    _style_title(ch, title)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(size), FONT
    ch.legend.font.color.rgb = BODY
    for i, s in enumerate(ch.plots[0].series):
        s.format.line.color.rgb = colors[i]
        s.format.line.width = Pt(3 if i == 0 else 1.75)
        s.smooth = False
        s.marker.format.fill.solid(); s.marker.format.fill.fore_color.rgb = colors[i]
        s.marker.format.line.color.rgb = colors[i]
        if i > 0:
            from pptx.enum.chart import XL_MARKER_STYLE
            s.marker.style = XL_MARKER_STYLE.NONE
            s.format.line.dash_style = 4
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
    va.major_gridlines.format.line.color.rgb = LINE
    va.tick_labels.font.size, va.tick_labels.font.name = Pt(size - 1), FONT
    va.tick_labels.font.color.rgb = MUTED
    va.tick_labels.number_format, va.tick_labels.number_format_is_linked = fmt, False
    va.format.line.fill.background()
    va.minimum_scale = 0
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(size - 1), FONT
    ca.tick_labels.font.color.rgb = MUTED
    ca.format.line.color.rgb = LINE
    return ch


def combo(slide, x, y, w, h, cats, bars, lines, title, size=10):
    """Clustered columns plus lines on the same axis. bars/lines: list of (name, values, colour[, dashed])."""
    cd = CategoryChartData()
    cd.categories = cats
    for n, v, *_ in bars + lines:
        cd.add_series(n, v)
    ch = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    _chart_bg(ch)
    _style_title(ch, title)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(size + 1), FONT
    ch.legend.font.color.rgb = BODY
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
        sp = parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:ln w="38100" cap="rnd"><a:solidFill><a:srgbClr val="{c}"/>'
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
    va.major_gridlines.format.line.color.rgb = LINE
    va.tick_labels.font.size, va.tick_labels.font.name = Pt(size), FONT
    va.tick_labels.font.color.rgb = MUTED
    va.tick_labels.number_format, va.tick_labels.number_format_is_linked = '#,##0', False
    va.format.line.fill.background()
    va.minimum_scale = 0
    ca = ch.category_axis
    ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(size), FONT
    ca.tick_labels.font.color.rgb = MUTED
    ca.format.line.color.rgb = LINE
    return ch


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t



SLIDE_NO = {}                                                        # slide number of each block, for "see slide X" references


def mark(key):
    SLIDE_NO[key] = PAGE[0]


PARTS_NAV = [("A", "B2B"), ("B", "B2C"), ("C", "People"), ("D", "Money")]


PILLAR_NAMES = {1: "Clients", 2: "People", 3: "Company", 4: "Community"}


def divider(letter, chip, question, answer, stats, asks, part, pillars=()):
    """Dark section opener: the question this part answers, the short answer, two numbers, and the questions inside."""
    s = new_slide(None, dark=True)
    text(s, 0.8, 0.75, 2.6, 0.36, chip, size=11, bold=True, color=INK_D, fill=AMBER, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         radius=0.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 0.8, 1.2, 1.5, 1.75, letter, size=88, bold=True, color=AMBER, font=HEAD, margin=0, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 2.45, 1.25, 10.4, 1.1, question, size=33, bold=True, color=WHITE, font=HEAD, margin=0, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 2.45, 2.4, 10.2, 0.95, [[("Our answer: ", {"bold": True, "color": AMBER}), (answer, {})]], size=17,
         color=rgb("C9D2E3"), margin=0)
    for i, (v, lab) in enumerate(stats):
        stat(s, 2.45 + i * 3.35, 3.6, 3.15, 1.3, v, lab, vcolor=AMBER, fill=DARK2, lcolor=rgb("C9D2E3"), vsize=28, lsize=11)
    x0 = 2.45 + len(stats) * 3.35
    text(s, x0, 3.55, 12.85 - x0, 2.4, [[("Questions this part answers", {"bold": True, "color": AMBER, "size": 12})]] +
         [[("•  ", {"color": AMBER, "bold": True}), (t, {})] for t in asks], size=12, color=rgb("C9D2E3"), space=3, margin=0)
    if pillars:
        text(s, 2.45, 5.1, 6.5, 0.4, [[("Serves pillars:  ", {"bold": True, "color": AMBER}), ("  ·  ".join(f"{k} {PILLAR_NAMES[k]}" for k in pillars), {})]],
             size=13, color=rgb("C9D2E3"), margin=0, anchor=MSO_ANCHOR.MIDDLE)
    for i, (l, n) in enumerate(PARTS_NAV):
        on = part == i
        text(s, 2.45 + i * 2.6, 6.25, 2.45, 0.48, [[(f"{l}  ", {"bold": True}), (n, {})]], size=13, bold=on,
             color=INK_D if on else rgb("9FAAC0"), fill=AMBER if on else DARK2, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             radius=0.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return s


def fill_refs():                                                 # replace §key§ with the real slide number everywhere
    import re as _re
    sub = lambda t: _re.sub(r"§(\w+)§", lambda m: str(SLIDE_NO[m.group(1)]), t)
    for sl in prs.slides:
        frames = [sh.text_frame for sh in sl.shapes if sh.has_text_frame]
        frames += [c.text_frame for sh in sl.shapes if sh.has_table for row in sh.table.rows for c in row.cells]
        for tf in frames:
            for p in tf.paragraphs:
                for r in p.runs:
                    if "§" in r.text:
                        r.text = sub(r.text)
        if sl.has_notes_slide:
            nt = sl.notes_slide.notes_text_frame
            if "§" in nt.text:
                nt.text = sub(nt.text)

FULL = "Our model, a normal full year at the 2027 sales plan"          # the step-by-step 'what if' view
CUR = "Our model, current estimate for 2027, month by month"             # the calendar view (Excel)
ok = lambda m: BLUE if m >= 0.15 else TERRA
B = lambda t: (t, {"bold": True, "color": INK_D})
N = lambda t: (t, {})
H = lambda t, c=None, sz=16: [(t, {"bold": True, "color": c or INK_D, "size": sz, "font": HEAD})]

# ================================================================== 1. Cover
s = new_slide(None, dark=True)
text(s, 0.8, 0.85, 4.6, 0.38, "TEMAN JOURNEY  ·  GROWTH PLAN  ·  5 OCT 2026", size=11, bold=True, color=INK_D, fill=AMBER,
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.8, 1.5, 11.6, 1.9, [[("Create ", {}), ("real value for our clients", {"color": AMBER}), (", in a business that pays for itself", {})]],
     size=46, bold=True, color=WHITE, font=HEAD, margin=0)
text(s, 0.8, 3.45, 10.8, 0.9, "Two goals, tracked through four impact pillars: clients, people, company and community. "
                              "Fair pay and a 15%+ margin are how we show it works.", size=18, color=rgb("C9D2E3"), margin=0)
for i, (v, lab) in enumerate([(pc(F27["m"], 1), "net margin in 2027 (current estimate)"),
                              (rp(F27["low"]), f"lowest cash, {LOWM}: our weak point"),
                              ("March 2027", "new work starts; we pitch and sign now")]):
    stat(s, 0.8 + i * 3.95, 4.7, 3.7, 1.35, v, lab, vcolor=AMBER, fill=DARK2, lcolor=rgb("C9D2E3"), vsize=30, lsize=12)
text(s, 0.8, 6.4, 11.5, 0.4, "Three estimates in this deck: conservative, current and aggressive. All money is in Rupiah million (Rp M).",
     size=12, color=rgb("9FAAC0"), margin=0)
notes(s, "This deck is written to be read without a presenter. Slide §summary§ gives the whole answer. Each part opens "
         "with a dark slide and the question it answers: the goal and the problem (slide §d_goal§), A. B2B (§d_a§), B. B2C (§d_b§), "
         "C. people (§d_c§), D. money (§d_d§), next steps (§d_next§) and the appendix (§d_app§).")

mark("cover")

# ================================================================== S. Summary (the top of the pyramid)
FY = lambda k: sum(B2C[k][3:])
B2B_SH = FY("b2b") / (FY("b2b") + FY("rev"))
first_cov = next((G.MONTHS[i].replace("-", " 20") for i in range(3, 15) if B2C["sess"][i] * 0.6 >= B2C["psy"][i]), "after 2027")
s = new_slide(f"Value for clients pays for fair pay and {pc(F27['m'], 1)} margin", accent=pc(F27["m"], 1), tag="answer",
              sub="Goals and pillars: slide §goal§. Money in (A. B2B, B. B2C), money out (C. people), timing (D. cash).",
              source=CUR + ". Current estimate; conservative and aggressive estimates on slide §estimates§.")
PARTS = [("A", "B2B is our engine", pc(B2B_SH), "of 2027 sales", INK_D,
          ["Win the head office, spread to its branches, keep the client with EAP",
           "Send BI one full-year 2027 proposal by 17 October",
           f"EAP and banks cut BI's share from {pc(BI_S1)} to {pc(BI_S4)}"]),
         ("B", "B2C pays for our clinic", pc(1 - B2B_SH), "of 2027 sales", TERRA,
          [f"Counseling pays our 2 part-time psychologists from {first_cov}",
           "Referral wins a client at a quarter of the ad cost",
           "Full-time psychologist only when B2C can pay for it (2028)"]),
         ("C", "People follow the work", rp(REC['running_total'][12]), f"running cost a month by Oct 2027, up in steps from {rp(REC['running_total'][0])}", BLUE,
          ["No signed work, no hire", "Pay rises unlock with sales targets, not dates",
           "The limit moves: founders now, project manager from July"]),
         ("D", "Cash is protected", rp(F27["low"]), f"lowest cash ({LOWM})", rgb("6379A1"),
          ["Sell 4 to 6 months before the work", "Guard rule + 30% down payment from BI",
           f"Bad year (conservative): still {pc(RUNS['conservative']['fy27']['m'], 1)} margin"])]
for i, (l, h, v, lab, f, b) in enumerate(PARTS):
    x = 0.5 + i * 3.1
    card(s, x, 1.55, 2.95, 4.05, "")
    circle(s, x + 0.2, 1.72, 0.55, l, fill=f, size=18)
    text(s, x + 0.85, 1.65, 2.05, 0.7, h, size=14, bold=True, color=INK_D, font=HEAD, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    text(s, x + 0.2, 2.45, 2.6, 1.05, [[(v, {"font": HEAD, "size": 30, "bold": True, "color": f if f != INK_D else TERRA})],
                                      [(lab, {"size": 11, "color": MUTED})]], margin=0.02, space=0)
    text(s, x + 0.2, 3.6, 2.6, 1.95, [[("•  ", {"color": TERRA, "bold": True}), (t, {})] for t in b], size=11, space=5, margin=0.02)
callout(s, 0.5, 5.8, 12.25, 1.1,
        [[B("The two problems, and where we solve them: "), N("profit that depends on low pay (slide §problem§) is solved in A, B and C: "
                                                            "more B2B and B2C sales pay for fair pay, released in steps. The cash gap in the pitch months "
                                                            "(slide §timing§) is solved in A and D: sell early, get paid earlier, guard the cash.")]], size=13)
notes(s, "This is the whole deck on one page. A and B are where money comes from, C is what we spend on people, D is "
         "when the money moves. The rest of the deck proves each part, in this order.")

mark("summary")

# ================================================================== Divider: What do we want, and what stops us?
divider('?', 'WHY WE ACT', 'What do we want, and what stops us?', 'Real value for clients, in a business that pays for itself. Today profit relies on low pay, and cash runs thin before March.', [("2 + 4", "goals + impact pillars"), ("15%+", "net margin we need")], ['What are our goals, and how do we measure them?', 'Why 15%?', 'Why do we lose money in some months?', 'Why are the new prices not enough?', 'Why is cash tight?'], None)
mark("d_goal")

# ================================================================== 2. Goals and impact pillars
s = new_slide("Two goals, four impact pillars", sub="The goals say why we exist. The pillars show, with numbers, that it is working.",
              accent="four impact pillars", tag="goal",
              source=CUR + ". Fair-pay floor: Jakarta minimum wage 2026. Before/after scores and EAP renewals: new measures, tracked from Q4 2026.")
GOALS = [("1", "Create real value for our clients", "The people we train and counsel change what they know and do, and clients can see it."),
         ("2", "Build a business that pays for itself", "A 15%+ margin and safe cash, without outside money, so we can keep doing goal 1.")]
for i, (n, h, t) in enumerate(GOALS):
    x = 0.5 + i * 6.25
    card(s, x, 1.5, 6.05, 1.2, "", fill=INK_D)
    circle(s, x + 0.2, 1.8, 0.6, n, fill=AMBER, size=20)
    text(s, x + 0.95, 1.58, 4.95, 1.05, [[(h, {"font": HEAD, "size": 18, "bold": True, "color": WHITE})], [(t, {"size": 11, "color": rgb("C9D2E3")})]],
         margin=0.02, space=2, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.5, 2.78, 12.3, 0.3, "WE KNOW IT IS WORKING WHEN THESE NUMBERS MOVE", size=11, bold=True, color=MUTED, align=PP_ALIGN.CENTER, margin=0)
_sess27 = sum(B2C["sess"][3:])
_vol, _inp = WB["Volume"], WB["Inputs"]
REACH_PER_TRAINER = 2                                            # estimate: people one training participant passes it on to (founder: total 1-2k)
TRAINERS27 = _vol["D23"].value * _inp["C63"].value + _vol["D24"].value * _inp["C64"].value   # Flagship + KPW/LPS ToT
EVENT_PAX27 = _vol["D25"].value * _inp["C65"].value + _vol["D27"].value * _inp["C66"].value  # bank days + public events
REACH27 = TRAINERS27 * REACH_PER_TRAINER
PILLARS = [("Clients", "BI, banks, EAP clients", INK_D,
            [(f"{WB['Volume']['E26'].value} → {WB['Volume']['N26'].value}", f"EAP contracts active, by {G.MONTHS[9].replace('-', ' 20')}; then we track renewals (slide §b2b_deals§)"), ("Before / after", "scores on every training, from Q4 2026"),
             (f"{B2C['spc'][14]:.1f}", "sessions per B2C client from Q3 2027: our plan (slide §b2c§)")]),
           ("People", "Staff and psychologists", BLUE,
            [("Rp5.73M", "fair-pay floor + BPJS and THR (slide §payscheme§)"), (f"Rp{PERF[0]:.0f}/{PERF[1]:.0f}/{PERF[2]:.0f}M", "monthly sales that unlock the 3 pay steps"),
             (pc(max(B2C["load"][3:])), "busiest load on paid part-timers (healthy 60-80%) (slide §psycap§)")]),
           ("Company", "Founders and future investors", TERRA,
            [(pc(F27["m"], 1), "net margin in 2027; we need 15%+ (slide §estimates§)"), (rp(F27["low"]), f"lowest cash ({LOWM}) (slide §cashplan§)"),
             (f"{pc(BI_S1)} → {pc(BI_S4)}", "BI's share of sales, new prices only → full plan (slide §revmix§)")]),
           ("Community", "People beyond our clients", rgb("6379A1"),
            [(f"{TRAINERS27:,.0f} + ~{round(REACH27, -2):,.0f}", f"people we train in 2027 + people they reach (estimate: {REACH_PER_TRAINER} each)"),
             (f"{EVENT_PAX27:,.0f} + {_sess27:,.0f}", "people at our bank and public events + counseling sessions (slide §b2c§)"),
             ("20%", "new clients from referral by Q4 2027 (slide §referral§)")])]
for i, (h, who, f, mets) in enumerate(PILLARS):
    x = 0.5 + i * 3.1
    card(s, x, 3.15, 2.95, 3.55, "")
    circle(s, x + 0.18, 3.3, 0.48, str(i + 1), fill=f, size=15)
    text(s, x + 0.75, 3.27, 2.15, 0.55, [[(h, {"font": HEAD, "size": 16, "bold": True, "color": INK_D})], [(who, {"size": 10, "color": MUTED})]],
         margin=0.02, space=0, anchor=MSO_ANCHOR.MIDDLE)
    for j, (v, lab) in enumerate(mets):
        text(s, x + 0.18, 3.95 + j * 0.9, 2.65, 0.88, [[(v, {"font": HEAD, "size": 17, "bold": True, "color": f if f != INK_D else TERRA})],
                                                       [(lab, {"size": 10, "color": BODY})]], margin=0.02, space=0)
notes(s, "Goal 1 is why we exist: value for clients. Goal 2 is what lets us keep doing it: a business that pays for itself. "
         "Fair pay is not a goal on its own; it is one result we track, in the People pillar. Each pillar has three numbers. "
         "Most already sit in parts A to D; the slide numbers show where. Two are new and start in Q4 2026: before/after scores "
         "on each training, and EAP renewals. They are our proof of value for the next client. Community reach: "
         f"{TRAINERS27:,.0f} training participants in 2027 ({WB['Volume']['D23'].value} Flagship programs x {WB['Inputs']['C63'].value} people, "
         f"{WB['Volume']['D24'].value} regional ToT x {WB['Inputs']['C64'].value}), each passing it on to about {REACH_PER_TRAINER} people "
         "(estimate; replace with BI's ambassador reports). Plus people at bank and public events, and B2C counseling sessions.")

mark("goal")

# ================================================================== 3. Why 15%
vals = [round(v, 1) for v in REC["net"]]                          # Oct 2026 to Dec 2027
worst = min(vals)
LOSS15 = sum(1 for v in vals if v < 0)
s = new_slide(f"We lose money in {LOSS15} of the next 15 months", accent=f"{LOSS15} of the next 15 months", tag="goal",
              sub="Salaries come every month, big programs do not. Until March we only pitch, so the good months must carry us.",
              source=CUR + ". Net profit = sales minus all costs and tax.")
bar(s, 0.5, 1.5, 8.3, 5.4, M15, [("Net profit", vals)], "Net profit each month, Oct 2026 to Dec 2027 (Rp M). Red = a loss.",
    point_colors=[BLUE if v >= 0 else TERRA for v in vals], fmt='#,##0', gap=35, size=10)
stat(s, 9.1, 1.6, 3.75, 1.5, rp(-worst), f"lost in the worst month ({M15[vals.index(worst)]})", vcolor=TERRA)
stat(s, 9.1, 3.25, 3.75, 1.5, rp(-sum(v for v in vals[:5] if v < 0)), "lost from October to February, while we pitch and sign")
callout(s, 9.1, 4.9, 3.75, 2.0, [[B("Without the cushion, ")], N("one lost program in a good month can turn the whole year into a loss.")], size=13)
notes(s, "Our income is uneven, our costs are not. The 15% cushion keeps salaries safe in quiet months and leaves money to grow.")

mark("why15")

# ================================================================== 5. Problem
s = new_slide("The new prices work only because pay is low", accent="pay is low", tag="problem",
              sub="Fair pay alone drops the margin to " + pc(S2["m"]) + ". More EAP, not a second price rise, brings it back.",
              source=FULL + ". Fair pay = minimum wage + 18% for BPJS and THR (estimate).")
cats = ["Old prices", "New prices", "+ fair pay", "+ 2 EAP", "+ co-founder and bank deal"]
vals5 = [G.OLD["m"], S1["m"], S2["m"], S3["m"], S4["m"]]
bar(s, 0.5, 1.5, 7.4, 5.4, cats, [("Net margin", vals5)], "Net margin in a normal full year (blue = 15% or more)",
    point_colors=[ok(v) for v in vals5])
card(s, 8.15, 1.6, 4.7, 2.85,
     [H("How to read the bars", sz=15),
      [B("+ fair pay: "), N(f"costs {rp(12 * FAIR_MO)} a year.")],
      [B("+ 2 EAP: "), N(f"each contract adds {rp(EAP_GP)} of gross profit a year.")],
      [B("+ co-founder and bank deal: "), N(f"{pc(S4['m'])}.")]], size=13, space=7)
callout(s, 8.15, 4.6, 4.7, 2.3,
        [H("Why not raise prices again?", TERRA, 15),
         N(f"Every event would need +{pc(UPLIFT)}: the 3-day program at about Rp116M. Regulators have budget limits, and clients may leave.")],
        size=13, anchor=MSO_ANCHOR.TOP, space=6)
notes(s, "Full-time staff are underpaid: this plan fixes it. Part-time staff are paid enough: we keep their rates and only "
         "change how their contracts work (slides §kraljic§ and §contracts§). Each EAP contract adds about 5 to 7 points of margin: "
         + ", ".join(f"{n} new = {pc(m)}" for n, m in CURVE) + ".")

mark("problem")

# ================================================================== 6. Real timing
s = new_slide(f"Cash is our weak point: lowest {rp(F27['low'])} in {LOWM}", accent=f"{rp(F27['low'])} in {LOWM}", tag="problem",
              sub="We pitch and sign from October to February. New work starts in March, and clients pay 0 to 60 days later.",
              source=CUR + " (Outlook sheet). No new events before March 2027.")
combo(s, 0.5, 1.5, 8.5, 5.4, M15,
      [("Sales", [round(v) for v in REC["sales"]], BLUE), ("All costs", [round(v) for v in REC["cost"]], GREY)],
      [("Cash in the bank", [round(v) for v in REC["cash"]], "1E2A3A"),
       ("Safety floor (1 month of costs)", [round(v) for v in REC["floor"]], "9B3809", True)],
      "Each month, Rp M: sales and costs (bars), cash (lines)")
stat(s, 9.25, 1.6, 1.75, 1.25, str(len(BELOW)), "months under the red line", vcolor=TERRA, vsize=28, lsize=11)
stat(s, 11.1, 1.6, 1.75, 1.25, rp(F27["end"]), "cash at end of 2027", vsize=22, lsize=11)
callout(s, 9.25, 3.0, 3.6, 2.15,
        [H("Guard rule", TERRA, 14),
         N(f"If February ends under Rp80M (now: {rp(REC['cash'][4])}), the second founder's pay waits until May and extra ads start in March.")],
        size=12, anchor=MSO_ANCHOR.TOP, space=4, margin=0.16)
card(s, 9.25, 5.3, 3.6, 1.6,
     [H("Two more cash levers", sz=14), N(f"A 30% down payment on head-office programs, and 3 events moved to Jul-Sep. All three: lowest {rp(LEVERS['low'])}. More: slide §cashplan§.")],
     size=12, space=4, margin=0.16)
notes(s, "Grey bar taller than blue: we spend more than we sell that month. The red dashed line is one month of costs. "
         "Put real monthly sales and costs next to these bars every month.")

mark("timing")

# ================================================================== Divider: How do we grow B2B, our main engine?
divider('A', 'PART A OF D', 'How do we grow B2B, our main engine?', 'Win the head office, spread to its branches, keep the client with EAP, and sell 4 to 6 months before the work.', [(pc(B2B_SH), "of 2027 sales"), ("17 Oct", "BI proposal deadline")], ['Who do we sell to, and in what order?', 'Which deals must we sign, and by when?', 'What do we charge, and is it fair?', 'When must we start selling?', 'How do we depend less on BI?', 'Who sells with us?'], 0, (1, 3))
mark("d_a")

# ================================================================== A1. B2B strategy
s = new_slide("Win the head office, spread to branches, keep", accent="spread to branches", tag="b2b",
              sub="Who we serve: financial institutions that must change how people handle money.",
              source="Rate card and Roadmap sheet in our model. Bank buying order: our positioning review, October 2026.")
STEPS = [("1  Start", "Head-office trainer program", "about Rp100M", "3 days, yearly agreement", INK_D),
         ("2  Spread", "The same program in each branch", "Rp28.5-61M a day", "a work order per office", BLUE),
         ("3  Keep", "EAP for the client's staff", "Rp15M a month", "12 months, renews", TERRA)]
for i, (h, what, price, how, f) in enumerate(STEPS):
    x = 0.5 + i * 4.15
    text(s, x, 1.55, 3.95, 0.55, h, size=16, bold=True, color=WHITE, fill=f, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.2, font=HEAD)
    card(s, x, 2.15, 3.95, 1.55, [[(what, {"bold": True, "color": INK_D, "size": 14})],
                                  [(price, {"font": HEAD, "bold": True, "size": 22, "color": f if f != INK_D else TERRA})],
                                  N(how)], size=12, space=2)
    if i < 2:
        text(s, x + 3.97, 1.7, 0.16, 0.3, "", fill=BEIGE, shape=MSO_SHAPE.RIGHT_ARROW)
text(s, 0.5, 3.95, 12.3, 0.35, "Who we sell to, in this order", size=14, bold=True, color=INK_D, font=HEAD, margin=0.02)
ORDER = [("BI", "46 regional offices", "Now: 2027 budget"), ("OJK and LPS", "head office, then branches", "Pitch Jun-Sep 2027 for 2028"),
         ("State banks", "BRI, Mandiri and others", "EAP 3 and training days, 2027"), ("Private banks", "learning partnerships", "After the first state bank")]
for i, (who, what, when) in enumerate(ORDER):
    x = 0.5 + i * 3.1
    card(s, x, 4.35, 2.95, 1.25, [[(f"{i + 1}  {who}", {"font": HEAD, "bold": True, "size": 15, "color": INK_D})], N(what),
                                  [(when, {"bold": True, "color": TERRA, "size": 11})]], size=12, space=1, fill=SKY if i else CREAM)
callout(s, 0.5, 5.8, 12.3, 1.1,
        [[B("Why they choose us: "), N("psychology-led programs with tests before and after (Coreitera), so the client gets proof of "
                                       "behaviour change, not an attendance list.")],
         [B("Guardrail: "), N("we sell proven delivery and results, never 'connections' or access to regulators.")]], size=13, space=4)
notes(s, "This is our B2B strategy on one page. Start with one big program at a head office, spread the same program to its "
         "branches, then keep the client with a monthly EAP contract. BI is first because we already deliver its flagship program.")

mark("b2b_strategy")

# ================================================================== A2. B2B deals to sign
_pl = WB["PnL"]
_fy = lambda r: sum((_pl.cell(r, c).value or 0) for c in range(8, 20))
_eapn = [WB["Volume"].cell(26, c).value or 0 for c in range(8, 20)]
DEALS = [("BI 2027: head-office and regional programs", "BI 2027: Flagship and regional programs (first in March)", _fy(6) + _fy(7)),
         ("EAP 2: the warm lead", EAP2, 15 * sum(1 for n in _eapn if n >= 2)),
         ("EAP 3: a bank", "EAP contract 3 with a bank", 15 * sum(1 for n in _eapn if n >= 3)),
         ("Bank agreement: training days", "Bank agreement: training days", _fy(8)),
         ("OJK and LPS programs for 2028", "OJK and LPS programs for 2028", None)]
fmt_d = lambda d: d.strftime("%d %b %Y").lstrip("0") if d else "-"
rows = [["Deal", "2027 value (Rp M)", "Start pitching", "Signed by", "Owner", "Status today"]]
st_fill = {}
for i, (lab, key, val) in enumerate(DEALS, 1):
    r_ = RM[key]
    rows.append([lab, f"{val:,.0f}" if val else "2028", fmt_d(r_[8]), fmt_d(r_[5]), r_[9], r_[10]])
    st_fill[(i, 5)] = {"Late": RED_BG, "At risk": RED_BG, "Start now": CREAM}.get(r_[10], BLUE_BG)
s = new_slide("The 2027 deals we must sign, and by when", accent="by when", tag="b2b",
              sub="Regulators and banks plan budgets a year ahead, so most of 2027 is decided in the next three months.",
              source="Roadmap sheet in our model (lead times are estimates). Values: current estimate for 2027.")
table(s, 0.5, 1.55, 8.6, [2.85, 1.2, 1.2, 1.2, 1.15, 1.0], rows, size=11, row_h=0.55, fills=st_fill)
TO_SIGN = sum(v for *_, v in DEALS if v)
stat(s, 9.4, 1.6, 3.45, 1.45, rp(TO_SIGN), "of 2027 sales still to sign", vcolor=TERRA)
stat(s, 9.4, 3.2, 3.45, 1.45, "17 Oct", "BI proposal deadline: inside BI's 2027 budget window")
callout(s, 0.5, 5.3, 12.35, 1.6,
        [H("How we sell", sz=14),
         [("•  ", {"color": TERRA, "bold": True}), ("One full-year proposal per client, not one event at a time", {})],
         [("•  ", {"color": TERRA, "bold": True}), ("Ask for a 30% down payment on head-office programs, and EAP billed a quarter ahead", {})],
         [("•  ", {"color": TERRA, "bold": True}), ("Log every deal's real dates in the Roadmap sheet, so next year's plan uses our own speed", {})]],
        size=12, anchor=MSO_ANCHOR.TOP, space=3)
notes(s, "Start now means the latest start date is within a month. Kia leads regulators and banks; Galih leads EAP 2. "
         "The BD co-founder takes over OJK, LPS and banks after joining.")

mark("b2b_deals")

# ================================================================== 7. Rate card
_in = WB["Inputs"]
I = lambda c: _in[c].value
AS = I("C92")                                                      # assessment + impact report, per participant
RC = [  # product, today's price, new price, delivery cost, note
    ("Regional training (BI offices, LPS), a day", I("C10"), I("C88") + AS * I("C64"), I("C17")),
    ("Head-office trainer program, 3 days", I("C9"), I("C87") + AS * I("C63"), I("C16")),
    ("State bank program, a group a day", I("C11"), I("C89") + AS * I("C65"), I("C18")),
    ("Private bank program, a group a day", I("C11"), I("C90") + AS * I("C65"), I("C18")),
    ("EAP for client staff, a month", I("C12"), I("C12"), I("C19")),
    ("Public event (new price is a proposal)", I("C13"), 24, I("C20")),
    ("Counseling, a session", I("C14"), I("C14"), I("C14") * I("C21"))]
f1 = lambda v: f"{v:g}" if v == int(v) else f"{v:.1f}"
K_NOW, K_NEW, K_COST = RC[0][1], RC[0][2], RC[0][3]
K_TRAIN = I("C88")
K_N27 = WB["Volume"]["D24"].value
s = new_slide(f"Regional training: Rp{f1(K_NOW)}M now, Rp{f1(K_NEW)}M in 2027", accent=f"Rp{f1(K_NEW)}M in 2027", tag="b2b",
              sub=f"It costs us Rp{f1(K_COST)}M to deliver. So we keep Rp{f1(K_NOW - K_COST)}M today, and Rp{f1(K_NEW - K_COST)}M at the new price.",
              source="Rate card and delivery costs in our model (Inputs). Margin = price minus delivery cost, before salaries and office. "
                     "The Coreitera assessment fee is already inside the delivery cost.")
bar(s, 0.5, 1.5, 4.5, 3.55, ["Today", "From Jan 2027"],
    [("Delivery cost", [K_COST, K_COST]), ("Margin: training", [K_NOW - K_COST, K_TRAIN - K_COST]), ("Margin: assessment", [0, K_NEW - K_TRAIN])],
    "Regional training, Rp M a day", colors=[GREY, BLUE, TERRA], stacked=True, legend=True, size=11, gap=55,
    texts=[[f"Cost {f1(K_COST)}", f"Cost {f1(K_COST)}"], [f"Margin {f1(K_NOW - K_COST)}", f"Training {f1(K_TRAIN - K_COST)}"],
           ["", f"Report {f1(K_NEW - K_TRAIN)}"]])
stat(s, 0.5, 5.15, 2.2, 1.0, f"{f1(K_NOW - K_COST)} → {f1(K_NEW - K_COST)}", "Rp M we keep a day", vcolor=TERRA, vsize=22, lsize=10)
stat(s, 2.8, 5.15, 2.2, 1.0, f"+Rp{int((K_NEW - K_NOW) * K_N27 + 0.5)}M", f"more margin in 2027 ({K_N27} regional days)", vcolor=TERRA, vsize=22, lsize=10)
rows = [["Product", "Today", "New", "Our cost", "We keep"]]
for prod, now, new, cost, *_ in RC:
    rows.append([prod, f1(now), f1(new), f1(cost), f"{f1(now - cost)} → {f1(new - cost)}"])
table(s, 5.3, 1.5, 7.5, [3.35, 0.8, 0.8, 0.9, 1.65], rows, size=12, row_h=0.5, fills={(1, j): CREAM for j in range(5)})
text(s, 5.3, 5.55, 7.5, 0.6, [[B("New price = training + assessment and impact report "), N(f"(Rp{AS * 1000:.0f}k a participant). "
                                "Add-ons: Bank Elite +30% (trainer licence + 3 months follow-up); Theta app Rp5k a staff a month.")]],
     size=11, margin=0.02)
callout(s, 0.5, 6.25, 12.3, 0.62,
        [[B("Rules: "), N(f"price at least 3 times cost (training floor: 3 × Rp{f1(K_COST)}M = Rp{f1(3 * K_COST)}M)  ·  discounts at most 10%  ·  "
                          "first pilot at most 20% off  ·  client pays travel  ·  a regulator price rise always comes with a new item")]], size=12)
notes(s, f"Regional training (BI regional offices and LPS) is our most frequent product: {K_N27} days in 2027. Today we charge "
         f"Rp{f1(K_NOW)}M and it costs Rp{f1(K_COST)}M, so we keep Rp{f1(K_NOW - K_COST)}M. Our 3x rule sets the floor for the training "
         f"part at Rp{f1(3 * K_COST)}M. Adding the assessment and impact report ({I('C64')} participants x Rp{AS * 1000:.0f}k) brings "
         f"the price to Rp{f1(K_NEW)}M, so we keep Rp{f1(K_NEW - K_COST)}M. Counseling is the one product below 3x (2.5x) on purpose: "
         "it keeps sessions affordable. The public event price of Rp24M is a proposal; the model still uses Rp20M.")

mark("ratecard")

# ================================================================== 8. Price benchmarks
s = new_slide("Our prices sit between the state guide and big firms", accent="between", tag="b2b",
              sub="Blue = our prices. Grey = outside prices. Red = below our 3 times rule.",
              source="State fee guide: Finance Ministry SBM 2026 (PMK 32/2025), 1-day ToT. Big-firm rate: founders' estimate (2.5x our old bank price).")
pd_cats = ["State fee guide, 1 day", "Our regional program", "Our 3-day program, per day", "Our state bank price",
           "Our private bank price", "Big training firms (banks)"]
pd_vals = [12.6, 28.5, round(100 / 3, 1), 51.0, 61.0, 87.5]
bar(s, 0.5, 1.5, 6.3, 4.55, pd_cats, [("Rp M per day", pd_vals)], "Price for one training day (Rp M)",
    point_colors=[GREY, BLUE, BLUE, BLUE, BLUE, GREY], fmt='#,##0.0', horizontal=True, gap=45,
    texts=[[f"{v:.1f}" for v in pd_vals]])
mx_cats = ["3-day program", "Regional program", "State bank", "Private bank", "EAP", "Public event", "Counseling"]
mx_vals = [4.0, round(28.5 / 8, 2), 4.25, round(61 / 12, 2), 3.0, 2.5, 2.5]
bar(s, 7.0, 1.5, 5.85, 4.55, mx_cats, [("Times our cost", mx_vals)], "Price divided by our cost (rule: 3 or more)",
    point_colors=[BLUE if v >= 3 else TERRA for v in mx_vals], fmt='0.0"x"', horizontal=True, gap=45,
    texts=[[f"{v:.1f}x" for v in mx_vals]])
callout(s, 0.5, 6.2, 12.3, 0.7, [[B("Two changes: "), N("raise public events to Rp24M. Keep counseling at Rp1M: it brings new clients and is only 1.5% of sales.")]], size=13)
notes(s, "We want to be clearly cheaper than big training firms but never the cheapest. The 3 times rule makes sure each "
         "event also pays its share of salaries and office.")

mark("bench")

# ================================================================== 19. Roadmap (worked backward)
from datetime import date, datetime                                # noqa: E402
s = new_slide("Start selling 4 to 6 months before the work", accent="4 to 6 months", tag="b2b",
              sub="Regulators and banks plan next year's budget now, so this quarter decides most of 2027.",
              source="Roadmap sheet in our model. Lead times are estimates until we log real deals.")
GX, GW, GY, RH = 3.55, 9.3, 1.85, 0.3
MW = GW / 15


def gx(d):                                                         # date -> x position (Oct 2026 = month 0)
    d = d.date() if isinstance(d, datetime) else d
    m = (d.year - 2026) * 12 + d.month - 10 + (d.day - 1) / 31
    return GX + max(0.0, min(15.0, m)) * MW


for k, m in enumerate(G.MONTHS):
    text(s, GX + k * MW, 1.52, MW, 0.28, m[:3], size=10, bold=True, color=MUTED, align=PP_ALIGN.CENTER, margin=0)
text(s, GX + 3 * MW, 1.3, 3 * MW, 0.26, "2027 →", size=8, color=MUTED, margin=0)
text(s, 0.85, 1.42, 2.6, 0.3, [[("■ ", {"color": LBLUE}), ("sell or prepare   ", {}), ("■ ", {"color": INK_D}), ("work runs", {})]],
     size=9, color=MUTED, margin=0)
rm = lambda key, col: RM[key][col]
D = lambda y, mo, d=1: date(y, mo, d)
ROWS19 = [
    ("Sales", INK_D, "BI 2027 proposal (Kia)", [(rm("BI 2027: Flagship and regional programs (first in March)", 8), rm("BI 2027: Flagship and regional programs (first in March)", 5), "sell"), (D(2027, 3), D(2027, 12, 31), "work")]),
    ("Sales", INK_D, "EAP 2: warm lead (Galih)", [(D(2026, 10, 5), rm(EAP2, 5), "sell"), (D(2027, 2), D(2027, 12, 31), "work")]),
    ("Sales", INK_D, "EAP 3 + bank days (Kia, BD)", [(rm("EAP contract 3 with a bank", 8), rm("EAP contract 3 with a bank", 5), "sell"), (D(2027, 7), D(2027, 12, 31), "work")]),
    ("Sales", INK_D, "B2C events: Instagram first", [(rm("B2C events (Instagram campaign, then tickets)", 8), rm("B2C events (Instagram campaign, then tickets)", 5), "sell"), (D(2027, 10), D(2027, 12, 31), "work")]),
    ("Sales", INK_D, "OJK, LPS 2028 (BD co-founder)", [(rm("OJK and LPS programs for 2028", 8), rm("OJK and LPS programs for 2028", 4), "sell")]),
    ("People", BLUE, "Trainers: 3 for Mar, 8 for Oct-Dec", [(rm("3 certified trainers", 8), D(2027, 3), "sell"), (rm("8 certified trainers", 8), D(2027, 10), "sell")]),
    ("People", BLUE, "Standby project manager", [(rm("Standby project manager (paid per project)", 8), D(2027, 7), "sell"), (D(2027, 7), D(2027, 12, 31), "work")]),
    ("People", BLUE, "Psychologist pool; Lead Partnership", [(rm("Part-time psychologist pool for EAP 2", 8), D(2027, 2), "sell"), (rm("Lead Partnership (non-psychologist)", 8), D(2027, 10), "sell"), (D(2027, 10), D(2027, 12, 31), "work")]),
    ("People", BLUE, "Sales co-founder", [(rm("Sales co-founder", 8), D(2027, 7), "sell"), (D(2027, 7), D(2027, 12, 31), "work")]),
]
lanes = {}
for i, (lane, col, lab, segs) in enumerate(ROWS19):
    y = GY + i * RH
    lanes.setdefault(lane, [y, y, col])[1] = y + RH
    if i % 2 == 0:
        text(s, 0.85, y, GX + GW - 0.85, RH, "", fill=rgb("F3EEE4"), shape=MSO_SHAPE.RECTANGLE)
    text(s, 0.9, y, 2.6, RH, lab, size=11, color=INK_D, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    for a, b, kind in segs:
        x0, x1 = gx(a), gx(b)
        text(s, x0, y + 0.07, max(0.06, x1 - x0), RH - 0.14, "", fill=LBLUE if kind == "sell" else INK_D,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        if kind == "sell":
            text(s, x1 - 0.09, y + 0.06, 0.18, RH - 0.12, "", fill=TERRA, shape=MSO_SHAPE.DIAMOND)
for lane, (y0, y1, col) in lanes.items():
    text(s, 0.5, y0, 0.3, y1 - y0, "", fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
    lt = text(s, 0.5 + 0.15 - (y1 - y0) / 2, (y0 + y1) / 2 - 0.15, y1 - y0, 0.3, lane.upper(), size=10, bold=True,
              color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    lt.rotation = 270
yc = GY + len(ROWS19) * RH + 0.06
text(s, 0.5, yc, 0.3, 0.5, "", fill=TERRA, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
lt = text(s, 0.4, yc + 0.1, 0.5, 0.3, "CASH", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
          anchor=MSO_ANCHOR.MIDDLE, margin=0); lt.rotation = 270
text(s, 0.9, yc, 2.6, 0.5, "Cash checks", size=11, color=INK_D, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
for d, lab, pos in [(D(2027, 1, 15), "Ask a 30% down payment", "left"), (D(2027, 2, 28), f"Guard check (slide §timing§): {rp(REC['cash'][4])}", "below"),
                    (D(2027, 4, 15), f"Lowest cash: {rp(F27['low'])}", "right")]:
    x = gx(d)
    text(s, x - 0.08, yc + 0.06, 0.16, 0.16, "", fill=TERRA, shape=MSO_SHAPE.DIAMOND)
    box = {"left": (x - 2.12, yc, 2.0, PP_ALIGN.RIGHT), "below": (x - 1.1, yc + 0.22, 2.2, PP_ALIGN.CENTER),
           "right": (x + 0.12, yc, 2.0, PP_ALIGN.LEFT)}[pos]
    text(s, box[0], box[1], box[2], 0.28, lab, size=9, color=TERRA, bold=True, align=box[3], margin=0)
xt = gx(date(2026, 10, 5))
ln_ = s.shapes.add_connector(1, Inches(xt), Inches(1.78), Inches(xt), Inches(yc + 0.5))
ln_.line.color.rgb = TERRA; ln_.line.width = Pt(1.5); ln_.line.dash_style = 4
text(s, xt + 0.04, 1.3, 0.6, 0.26, "Today", size=8, bold=True, color=TERRA, margin=0)
by = yc + 0.66
card(s, 0.5, by, 7.5, 6.92 - by,
     [H("This month", sz=14),
      [B("1. BI (Kia): "), N("one full-year 2027 proposal by 17 October; ask for a 30% down payment.")],
      [B("2. EAP (Galih): "), N("chase the warm lead; sign by mid-January.")],
      [B("3. People: "), N("start the BD co-founder search; agree the target-based pay scheme; register BPJS.")]],
     size=11, space=2, margin=0.16)
stat(s, 8.2, by, 2.25, 6.92 - by, pc(F27["m"], 1), "net margin 2027, current estimate", vcolor=AMBER, fill=INK_D, lcolor=WHITE, vsize=26, lsize=11)
stat(s, 10.6, by, 2.25, 6.92 - by, rp(LEVERS["low"]), "lowest cash with the guard rule and both cash levers", vsize=24, lsize=10)
notes(s, "Read from the right: each dark bar is work we must deliver. The light bar before it is the time to find, pitch and "
         "sign the client, or to recruit and train people. Regulators and banks plan next year's budget in the second half "
         "of this year, so BI, EAP 3 and bank days all start now. Dates come from the Roadmap sheet in the model. "
         f"Without the levers, lowest cash is {rp(F27['low'])} ({LOWM}) and cash is under the floor in {len(BELOW)} months.")

mark("roadmap")

# ================================================================== 15. Revenue mix
s = new_slide(f"EAP and banks cut our BI share from {pc(BI_S1)} to {pc(BI_S4)}", accent=f"{pc(BI_S1)} to {pc(BI_S4)}", tag="b2b",
              sub="BI stays the same size in rupiah. Other clients grow around it.",
              source=FULL + ". 'With growth' = 2 more EAP contracts and 10 bank training days.")
mix = [("BI (Bank Indonesia)", [BI, BI], INK_D), ("Other regulators and banks", [OTHER1, OTHER1 + 510.0], BLUE),
       ("EAP (monthly)", [EAP_REV1, 3 * EAP_REV1], TERRA), ("Public events and counseling", [B2C_REV, B2C_REV], GREY)]
bar(s, 0.5, 1.5, 7.4, 5.4, ["2027 plan", "With growth"], [(n, [round(a) for a in v]) for n, v, _ in mix],
    "Sales for one year by type of client (Rp M)", colors=[c for *_, c in mix], fmt='#,##0', stacked=True, legend=True, gap=80)
stat(s, 8.2, 1.6, 4.65, 1.6, f"{pc(BI_S1)} → {pc(BI_S4)}", "BI's share of our sales. Target: below 40%, so one budget cut cannot sink us.", vcolor=INK_D)
stat(s, 8.2, 3.35, 4.65, 1.6, f"{pc(EAP_SHARE_S1)} → {pc(EAP_SHARE_S4)}", "monthly EAP income: it comes in quiet months too.", vcolor=TERRA)
callout(s, 8.2, 5.1, 4.65, 1.8, [[B("Investors like both: ")], N("many clients, and income that repeats every month.")], size=14)
notes(s, "Grow other clients, do not shrink BI. That is the safest way to reduce dependence on one client.")

mark("revmix")

# ================================================================== 17. Co-founder
s = new_slide("Hire a sales co-founder first", accent="sales", tag="b2b",
              sub="3 bank training days pay their cost; 10 days pay it more than 3 times.",
              source=FULL + ". Profit per bank day = Rp51M price minus delivery and project manager cost. Check share terms with a lawyer.")
rows = [["Question", "Sales co-founder", "Marketing co-founder"],
        ["Fixes our biggest limit?", "Yes: founders do all sales; one is in Singapore", "No: marketing does not close regulator deals"],
        ["How do clients buy?", "Regulators and banks buy on trust", "Ads help the public side most"],
        ["Could we hire staff instead?", "Hard: good sellers to regulators are rare", f"Yes: staff + ads cost about {rp(MKT_MO)} a month"],
        ["Our choice", "First", "Later, or as staff"]]
table(s, 0.5, 1.55, 7.0, [2.1, 2.6, 2.3], rows, size=12, row_h=0.78, fills={(4, 1): BLUE_BG, (4, 2): RED_BG})
bar(s, 7.8, 1.5, 5.05, 3.55, ["Co-founder cost", "Profit, 3 bank days", "Profit, 10 bank days"],
    [("Rp M a year", [12 * G.COF, 3 * GROUP_GP, 10 * GROUP_GP])], "One year (Rp M)",
    point_colors=[TERRA, BLUE, BLUE], fmt='#,##0', gap=50)
callout(s, 7.8, 5.2, 5.05, 1.7,
        [H("The deal", sz=14), N(f"About {rp(G.COF)} a month + shares earned over 4 years (none if they leave in year 1)."),
         N("Year-1 goals: 1 bank agreement, a first OJK or LPS program, 2 EAP renewals.")], size=12, space=3, anchor=MSO_ANCHOR.TOP)
text(s, 0.5, 5.65, 7.0, 1.2, [[B("Co-founder "), N("= a partner who owns part of the company. Shares are earned over time, so a partner who leaves early does not keep them.")]],
     size=12, color=MUTED)
notes(s, "A sales co-founder fixes the limit that matters most today: who sells to regulators and banks. Marketing can be hired as staff for less.")

mark("cofounder")

# ================================================================== Divider: Can B2C pay for our clinic?
divider('B', 'PART B OF D', 'Can B2C pay for our clinic?', 'Yes, from about April 2027: Instagram, referral and Theta vouchers grow B2C until it pays our part-time psychologists.', [(pc(1 - B2B_SH), "of 2027 sales"), (f"{B2C['ltvcac_ref'][14]:.1f}x", "referral: profit per client / cost")], ['How do followers turn into paid sessions?', 'What does it cost to win one client?'], 1, (4, 2))
mark("d_b")

# ================================================================== 17. B2C: Instagram to paid sessions
CNS = lambda i: B2C["sess"][i]
first_cov = next((G.MONTHS[i].replace("-", " 20") for i in range(3, 15) if B2C["sess"][i] * 0.6 >= B2C["psy"][i]), "after 2027")
fy27 = lambda k: sum(B2C[k][3:])
s = new_slide(f"B2C grows to {pc(fy27('rev') / (fy27('rev') + fy27('b2b')))} of sales and pays our psychologists", accent="pays our psychologists", tag="b2c",
              sub=f"From {first_cov}, counseling profit pays our 2 part-time psychologists. Bookings matter more than followers.",
              source=CUR + " (B2C sheet). Followers: @temanjourney_id on 5 Oct 2026. Today about 2 paid sessions a month.")
qb2b = [round(sum(B2C["b2b"][q])) for q in QS]
qb2c = [round(sum(B2C["rev"][q])) for q in QS]
bar(s, 0.5, 1.5, 6.6, 3.75, QN, [("B2B: regulators, banks, EAP", qb2b), ("B2C: counseling and public events", qb2c)],
    "Revenue per quarter (Rp M)", colors=[INK_D, TERRA], fmt='#,##0', stacked=True, legend=True, gap=60,
    texts=[[f"{v}" for v in qb2b], [f"{v}" if v >= 15 else "" for v in qb2c]])
rows = [["Per quarter (our levers)"] + QN,
        ["Follower growth"] + [pc(v) for v in B2C_Q[0]],
        ["Bookings per 1,000 followers"] + [f"{v:.2f}" for v in B2C_Q[1]],
        ["Sessions per client"] + [f"{v:.2f}" for v in B2C_Q[2]]]
table(s, 0.5, 5.45, 6.6, [2.35, 0.85, 0.85, 0.85, 0.85, 0.85], rows, size=11, row_h=0.36)
stat(s, 7.4, 1.6, 2.65, 1.3, f"{B2C['fol'][0] / 1000:.1f}k → {B2C['fol'][-1] / 1000:.0f}k", "Instagram followers, Oct 2026 to Dec 2027", vsize=22, lsize=11)
stat(s, 10.2, 1.6, 2.65, 1.3, f"2 → {CNS(14):.0f}", "paid counseling sessions a month", vcolor=TERRA, vsize=22, lsize=11)
stat(s, 7.4, 3.05, 2.65, 1.3, rp(fy27("rev")), "B2C revenue in 2027, incl. 4 public events", vsize=22, lsize=11)
stat(s, 10.2, 3.05, 2.65, 1.3, f"{B2C['cov'][14]:.1f}x", "psychologists' pay covered by B2C profit, Dec 2027", vsize=22, lsize=11)
callout(s, 7.4, 4.5, 5.45, 2.4,
        [H("When can we hire a full-time psychologist?", TERRA, 14),
         N(f"When counseling profit pays the lead part-timer plus a full-timer: about {rp(B2C['ftneed'][14])} a month for 3 months. "
           f"Dec 2027: {rp(B2C['ft'][14], 1)}. So early 2028, not 2027. Until then: one more paid part-timer, plus our pool of 20."),
         [B("Watch: "), N("8+ paid sessions a month by March, 13+ by June. If not, cut the extra ads back to Rp1M.")]],
        size=12, anchor=MSO_ANCHOR.TOP, space=5, margin=0.18)
notes(s, "The funnel: followers x booking rate (new paying clients a month per 1,000 followers) x sessions per client, plus "
         "sessions from Theta vouchers given to EAP staff. Today about 2 paid sessions a month from 3,439 followers. "
         "Growth is set per quarter on the B2C sheet. 10,000 followers at today's booking rate would give only about 6 sessions a month: "
         "bookings and repeat sessions (packages of sessions) do most of the work.")

mark("b2c")

# ================================================================== B2. Referral and unit economics
_ltv = B2C["ltv"][14]
_cac_jun, _cac_dec = B2C["cac"][8], B2C["cac"][14]
REF_REV = sum(B2C["refsess"][3:]) * 1.0
s = new_slide("Referral wins a client at a quarter of the ad cost", accent="a quarter of the ad cost", tag="b2c",
              sub="We ask every new client where they came from, and give a small session credit to clients who bring a friend.",
              source=CUR + " (B2C sheet, unit economics). Ads: Rp4.5M a month in total. Reward: Rp100k credit to each side.")
cats = ["Profit per client (LTV)", "Instagram client, Jun 2027", "Instagram client, Dec 2027", "Referred client"]
vals = [round(_ltv * 1000), round(_cac_jun * 1000), round(_cac_dec * 1000), 200]
bar(s, 0.5, 1.5, 6.3, 3.3, cats, [("Rp thousand", vals)], "Profit per client vs cost to win one client (Rp thousand)",
    point_colors=[INK_D, GREY, GREY, TERRA], fmt='#,##0', horizontal=True, gap=40, texts=[[f"{v:,}" for v in vals]])
stat(s, 0.5, 4.95, 3.05, 1.95, f"{B2C['ltvcac_ref'][14]:.1f}x", "profit per client ÷ cost, referral", vcolor=TERRA, vsize=28)
stat(s, 3.75, 4.95, 3.05, 1.95, f"{B2C['ltvcac_ig'][8]:.1f}x → {B2C['ltvcac_ig'][14]:.1f}x", "same for Instagram ads, Jun to Dec 2027", vsize=24)
PROG = [("1", "Ask the source", "One required question at every booking: Instagram, friend (code), Theta voucher, company, event, other"),
        ("2", "Give a code", "After the first session: a personal code (e.g. TJ-ANDI24) by WhatsApp template"),
        ("3", "Track in one sheet", "Date, client ID (no names), source, code, sessions, revenue, reward. Admin fills it"),
        ("4", "Review monthly", "New clients by source, referral share (target 20% by Q4 2027), cost per client by channel")]
for i, (n, h, t) in enumerate(PROG):
    y = 1.55 + i * 0.95
    card(s, 7.1, y, 5.75, 0.85, "")
    circle(s, 7.25, y + 0.17, 0.5, n, fill=[INK_D, BLUE, TERRA, rgb("6379A1")][i])
    text(s, 7.9, y + 0.05, 4.85, 0.78, [[(h, {"font": HEAD, "bold": True, "size": 13, "color": INK_D})], N(t)], size=11, space=1, margin=0.02)
callout(s, 7.1, 5.45, 5.75, 1.45,
        [[B("Fair and ethical: "), N("the reward is a session credit, never cash. No testimonials. Referring is the client's choice, "
                                     "and we never tell anyone who referred whom (HIMPSI code). Lead marketing owns the referral KPI.")]],
        size=11, anchor=MSO_ANCHOR.TOP)
notes(s, f"LTV = Rp1M price x 60% after the psychologist fee x 1.4 sessions = Rp{_ltv * 1000:,.0f}k. Instagram cost per client = "
         "all ad spend divided by new Instagram clients, so it is cautious. Ads earn back less than they cost until about July 2027. "
         f"Referral at 20% of new clients adds about {rp(REF_REV)} of B2C revenue in 2027. Upgrade the sheet to codes in the Theta app at 40+ new clients a month.")

mark("referral")

# ================================================================== Divider: Who do we hire, and when do we pay more?
divider('C', 'PART C OF D', 'Who do we hire, and when do we pay more?', 'No signed work, no hire. Pay rises when sales targets are met, and we fix each limit before it stops us.', [(rp(REC["running_total"][12]), "running cost a month, Oct 2027"), ("3", "pay steps, unlocked by sales")], ['When do we hire?', 'When does pay rise?', 'How does our running cost grow?', 'What limits us, and when?', 'How do we treat each type of worker?'], 2, (2,))
mark("d_c")

# ================================================================== 9. Hiring steps
steps_c = ["New prices, today's pay", "1. Fair pay for staff", "2. + Lead Partnership and 2 EAP", "3. + co-founder and bank deal",
           "4 to 6. Full team, Theta pays CTO", "If we paid the CTO"]
steps_v = [S1["m"], STEP1_M, S3["m"], S4["m"], S4_ADM2["m"], S6["m"]]
s = new_slide("Every hire comes with the work that pays for it", accent="the work that pays for it", tag="ppl",
              sub="No signed work, no hire. Red bars fall below our 15% target.", source=FULL + ". Net margin after each step.")
bar(s, 0.5, 1.5, 8.3, 5.4, steps_c, [("Net margin", steps_v)], "Net margin after each hiring step",
    point_colors=[ok(v) for v in steps_v], gap=45, texts=[[pc(v, 1) for v in steps_v]])
stat(s, 9.1, 1.6, 3.75, 1.55, pc(STEP1_M, 1), "if all staff get fair pay at once. So in 2027 pay rises in steps (slide §payscheme§).", vcolor=TERRA, lsize=12)
stat(s, 9.1, 3.3, 3.75, 1.55, pc(S6["m"], 1), "if Teman Journey paid the CTO. So Theta's own income pays the CTO.", vcolor=TERRA, lsize=12)
callout(s, 9.1, 5.0, 3.75, 1.9, [H("The rule", TERRA, 15), N("A new person and the new work that pays for them arrive together.")],
        size=13, anchor=MSO_ANCHOR.TOP, space=5)
notes(s, "This chart is the hiring plan in one picture. Each blue bar is a step where the new person and the new work arrive together.")

mark("hiring")

# ================================================================== A2. Pay scheme: rises unlock with sales targets
s = new_slide("Pay rises when sales targets are met", accent="sales targets", tag="ppl",
              sub="A step unlocks when average sales of the last 3 months reach the target, and never goes back down.",
              source=CUR + " (Pay_Plan sheet, rows 27-30). Rp M a month. The Instagram booking bonus is new and not yet in the model.")
rows = [["Role", "Now", f"Step 1: sales Rp{PERF[0]:g}M+", f"Step 2: sales Rp{PERF[1]:g}M+", f"Step 3: sales Rp{PERF[2]:g}M+", "Monthly KPI and bonus"],
        ["Admin", "1.8 + Rp17k a client + 0.62 bonus", "same", "at least 5.73", "same", "Invoice within 2 working days of client sign-off; reply to bookings the same day"],
        ["Social media", "1.5", "3.0", "4.5", "5.73", "3 posts a week; followers on track to 10k; Rp25k for each paid session booked through Instagram"],
        ["Lead marketing (has shares)", "2.5", "4.0 (cap)", "4.0", "4.0", "10k followers by December; ads must earn 3 times their cost (ROAS 3x)"],
        ["Psychologist (part-time)", "2.3", "2.3", "2.3", "2.3", "No full-time hire for now"],
        ["Lead Partnership", "-", "-", f"{LP:g} (Oct 2027)", f"{LP:g}", "Starts when 3 EAP contracts run; brings EAP clients (estimate)"],
        ["Founders", "10 (one paid)", "25 from March (room for both)", "25", "25", "Rp10M for one founder is the minimum; the second start follows the guard rule"]]
table(s, 0.5, 1.5, 12.3, [2.2, 1.75, 1.6, 1.55, 1.45, 3.75], rows, size=11, row_h=0.6)
for i, (m, k) in enumerate(zip(STEP_M, ("Step 1", "Step 2", "Step 3"))):
    stat(s, 0.5 + i * 2.75, 5.85, 2.6, 1.05, m, f"{k} in the current estimate", vsize=20, lsize=11)
callout(s, 8.85, 5.85, 3.95, 1.05, [[B("Always: "), N("BPJS and THR (18%) from January. Below-minimum pay is agreed in writing (PP 36/2021).")]], size=11)
notes(s, "Founders' decision, 5 October 2026: targets first, then pay. In a bad year the steps do not unlock, so the cost "
         "follows the sales. Sales = all money clients pay us in a month. Ask a labour consultant to confirm the written agreements.")

mark("payscheme")

# ================================================================== 11. Organisation chart evolution
s = new_slide("Our team and running cost grow in steps", accent="in steps", tag="ppl",
              sub=f"From {rp(REC['running_total'][0])} to {rp(REC['running_total'][12])} a month. Each column shows only what is new; each step follows signed work.",
              source=CUR + " (Pay_Plan sheet). Per-project and per-session people are paid only when there is work.")
ORG = [
    ["Founders x2 (one paid, Rp10M)", "Lead marketing (semi full-time, has shares)", "Social media", "Admin (base + fee per client)",
     "Lead psychologist, part-time", "2 interns", "Project manager (per project)", "2 trainers (per event)"],
    ["BPJS + THR for all staff", f"Instagram ads Rp{1 + ADS_X:g}M in total", "Pay scheme: rises unlock with sales targets",
     "Magang Nasional intern (bonus, if approved)"],
    ["Room for both founders (Rp25M)", "EAP part-time psychologists (per session)"],
    [f"Pay step 1 (since {STEP_M[0][:3]}): social media Rp3M, lead marketing Rp4M", "Sales co-founder", "EAP contract 3",
     "Standby project manager (per project)"],
    ["Pay step 2: admin at minimum wage, social media Rp4.5M", "Lead Partnership (non-psychologist)", "Bank training days",
     "Intern for invoices (from Sep)"],
]
TRIG = ["Today", "New prices start", "New work; EAP 2 since Feb", f"EAP 3; sales over Rp{PERF[0]:g}M", f"3 EAP; sales over Rp{PERF[1]:g}M"]
cw = 2.38
for j, ((i, lab), items) in enumerate(zip(PHASES, ORG)):
    x = 0.5 + j * (cw + 0.105)
    text(s, x, 1.5, cw, 0.85, [[(lab, {"bold": True, "size": 13})], [(rp(REC['running_total'][i]) + " a month", {"size": 17, "font": HEAD, "bold": True, "color": AMBER if j == 0 else WHITE})]],
         color=WHITE, fill=INK_D if j == 0 else BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE,
         margin=0.1, space=0, align=PP_ALIGN.CENTER)
    text(s, x, 2.4, cw, 0.3, [[("Trigger: ", {"bold": True}), (TRIG[j], {})]], size=10, color=MUTED, align=PP_ALIGN.CENTER, margin=0.02)
    for k, it in enumerate(items):
        new = j > 0
        text(s, x, 2.75 + k * 0.52, cw, 0.46, it, size=10, color=INK_D, fill=CREAM if new else SKY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE, margin=0.08, align=PP_ALIGN.CENTER, radius=0.2)
    if j < 4:
        text(s, x + cw + 0.005, 1.78, 0.1, 0.3, "", fill=BEIGE, shape=MSO_SHAPE.RIGHT_ARROW)
notes(s, "Read left to right. Blue boxes = today's team; cream boxes = new at that step. Everything from earlier columns stays.")

mark("org")

# ================================================================== 12. Capacity (COO review)
PROJ = [int(v) for v in REC["projects"]]
s = new_slide("Our limit moves during the year", accent="limit moves", tag="ppl",
              sub="Founders' selling time now, the project manager from July, trainers from October.",
              source=CUR + " (capacity rows). One event = one project. Each trainer runs at most 2 events a month.")
bar(s, 0.5, 1.5, 6.1, 3.25, M15, [("Projects", PROJ)], f"Projects per month (one project manager runs {PM_CAP})",
    point_colors=[TERRA if e > PM_CAP else BLUE for e in PROJ], fmt='0', gap=35, size=10,
    texts=[[str(v) if v else "" for v in PROJ]])
bar(s, 6.75, 1.5, 6.1, 3.25, M15, [("Trainers needed", FAC15), ("Trainers certified that month", CERT15)],
    "Trainers needed (blue) and new trainers certified 2 months earlier (sand)", colors=[BLUE, BEIGE], fmt='0', gap=35,
    size=10, legend=False, texts=[[str(v) if v else "" for v in FAC15], [str(v) if v else "" for v in CERT15]])
LIM = [("Founders' selling time", "Oct 2026 to Feb 2027", INK_D, "4 to 5 new clients, about 40 hours each; one founder is in Singapore.",
        "Kia (CEO) leads BI and banks; Galih leads EAP 2 and people."),
       ("Project manager", "Jul to Dec 2027", TERRA, f"Up to {max(PROJ)} projects a month: {max(PROJ) / PM_CAP:.0f} times what she can run.",
        f"A standby project manager at Rp{PM_FEE:g}M a project, booked by May."),
       ("Trainers", "Oct to Dec 2027", BLUE, f"{max(FAC15)} needed in the busiest month; we have 2 today.",
        "Recruit in July, certify by September. 8 is enough, not 11.")]
for j, (who, when, f, why, fix) in enumerate(LIM):
    x = 0.5 + j * 4.15
    text(s, x, 4.9, 3.95, 0.6, [[(who, {"bold": True, "size": 14, "font": HEAD})], [(when, {"size": 11})]], color=WHITE, fill=f,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, anchor=MSO_ANCHOR.MIDDLE, margin=0.15, space=0)
    card(s, x, 5.55, 3.95, 1.35, [[B("Why: "), N(why)], [B("Fix: "), N(fix)]], size=11, margin=0.14, space=3)
notes(s, "A business moves only as fast as its slowest step, and that step changes during the year. Red bars = more projects "
         "than one project manager can run. Trainers must be certified about 2 months before the work.")

mark("limits")

# ================================================================== 13. Kraljic
s = new_slide("Treat each type of worker differently", accent="differently", tag="ppl",
              sub="Sort our people by how hard they are to replace and how much they affect our results.",
              source="Kraljic, P. (1983), Purchasing must become supply management, Harvard Business Review. Team numbers from our model.")
gx_, gy_, gw_, gh_ = 1.2, 1.55, 7.4, 4.95
qw, qh = gw_ / 2, gh_ / 2
quads = [(0, 0, "Hard to replace, smaller effect", BLUE, SKY,
          ["Project manager: only 1 person", f"Trainers: 2 today, up to {max(FAC15)} needed"], "Secure them"),
         (1, 0, "Hard to replace, big effect", TERRA, RED_BG,
          ["Lead psychologist (part-time)", "Lead trainers that regulators trust"], "Long-term partnership, fair pay, a career path"),
         (0, 1, "Easy to replace, smaller effect", MUTED, SAND,
          ["Event helpers, design, interns", "Venue and travel (client pays)"], "Keep it simple, pay per day"),
         (1, 1, "Easy to replace, big effect", INK_D, CREAM,
          ["Part-time psychologists for EAP", "Many licensed psychologists in Jakarta"], "Build a group, pay per session")]
for c, r, h, f, fill, b, do in quads:
    x, y = gx_ + c * qw, gy_ + r * qh
    card(s, x + 0.05, y + 0.05, qw - 0.1, qh - 0.1,
         [[(h, {"bold": True, "color": f, "size": 14, "font": HEAD})]] + b + [[("Do: ", {"bold": True, "color": f}), (do, {"bold": True, "color": INK_D})]],
         fill=fill, size=12, margin=0.16, space=5)
text(s, 0.85 - gh_ / 2, gy_ + gh_ / 2 - 0.2, gh_, 0.4, "How hard to replace them", size=12, bold=True, color=MUTED,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, name="RiskAxis").rotation = -90
text(s, gx_, gy_ + gh_ + 0.02, gw_, 0.35, "How much they affect our results and our clients' trust", size=12, bold=True,
     color=MUTED, align=PP_ALIGN.CENTER)
stat(s, 8.95, 1.6, 3.9, 1.55, "Rp48M", "a year: all we save if we cut delivery costs by 10%. Keeping people matters more.", vcolor=TERRA)
card(s, 8.95, 3.3, 3.9, 3.6,
     [H("What this means", sz=15),
      [("•  ", {"color": TERRA, "bold": True}), ("Do not cut pay: the right price is not the lowest price.", {})],
      [("•  ", {"color": TERRA, "bold": True}), ("One project manager and two trainers limit how much we can sell.", {})],
      [("•  ", {"color": TERRA, "bold": True}), ("Pay part-timers per piece of work, so quiet months do not burn cash.", {})]],
     size=13, space=8)
notes(s, "This is the Kraljic matrix, a tool buyers use to sort suppliers. Here the suppliers are our own people.")

mark("kraljic")

# ================================================================== Divider: Is our cash safe, and when can we invest?
divider('D', 'PART D OF D', 'Is our cash safe, and when can we invest?', 'Safe in all three estimates if we sell early and get paid earlier. We invest only above a 6-month cushion.', [(pc(F27["m"], 1), "net margin, current estimate"), (rp(F27["low"]), "lowest cash, Feb 2027")], ['What if the year is bad, or good?', 'How do we protect cash until March?', 'When can we invest in new units?'], 3, (3,))
mark("d_d")

# ================================================================== D1. Three estimates
EST3 = [("Conservative", "BI buys 30% less, no bank deal", "conservative", GREY), ("Current", "our plan", "plan", INK_D),
        ("Aggressive", "bank deal from July, 2 days a month", "aggressive", BLUE)]
s = new_slide("Three estimates: a bad year, our plan, a good year", accent="a bad year", tag="money",
              sub="Even in the bad year the margin stays above 15%, because pay steps and hires wait for sales.",
              source=CUR + "; every estimate recalculated in Excel.")
bar(s, 0.5, 1.5, 5.6, 5.4, [e[0] for e in EST3], [("Net margin", [RUNS[e[2]]["fy27"]["m"] for e in EST3])],
    "Net margin 2027 (target: 15% or more)", point_colors=[e[3] for e in EST3], fmt='0.0%', gap=60,
    texts=[[pc(RUNS[e[2]]["fy27"]["m"], 1) for e in EST3]])
rows = [["2027", "Conservative", "Current", "Aggressive"],
        ["What it assumes"] + [e[1] for e in EST3],
        ["Sales (Rp M)"] + [f"{RUNS[e[2]]['fy27']['R']:,.0f}" for e in EST3],
        ["Net profit (Rp M)"] + [f"{RUNS[e[2]]['fy27']['N']:,.0f}" for e in EST3],
        ["Lowest cash (Rp M)"] + [f"{RUNS[e[2]]['fy27']['low']:,.0f}" for e in EST3],
        ["Cash at end of 2027 (Rp M)"] + [f"{RUNS[e[2]]['fy27']['end']:,.0f}" for e in EST3],
        ["Months below the safety floor"] + [str(RUNS[e[2]]['fy27']['below_floor']) for e in EST3]]
table(s, 6.4, 1.55, 6.45, [2.25, 1.4, 1.4, 1.4], rows, size=11, row_h=0.5, fills={(i, 2): BLUE_BG for i in range(1, 7)})
callout(s, 6.4, 5.3, 6.45, 1.6, [[B("What makes the bad year safe: "), N("the pay steps do not unlock, no one is hired without signed work, and B2C "
                                                                        "pays our psychologists. What makes it unsafe: cash in February to April, so the guard rule stays.")]],
        size=12, anchor=MSO_ANCHOR.TOP)
notes(s, "All three use the same costs and B2C plan. Only the B2B sales change. The current estimate is the plan we manage to.")

mark("estimates")

# ================================================================== A5. Cash savings, Oct 2026 to Mar 2027
SN = RUNS["save_now"]
s = new_slide("Protect cash in the pitch months", accent="Protect cash", tag="money",
              sub=f"Oct to Feb: we earn about Rp11M a month after delivery costs, but running cost is {rp(min(REC['running_total'][:5]))} to {rp(max(REC['running_total'][:5]))} a month.",
              source=CUR + ". Steps 2 and 4 recalculated in Excel; steps 3, 5, 6 and 7 are estimates.")
rows = [["#", "Step", "Cash gained", "What it takes", "When"],
        ["1", "Keep the extra Instagram ads, and check the return every month", "Rp0: ads pay for themselves",
         "Break-even: 6 paid sessions a month from ads (ROAS 1.7x). Target 3x", "Do now"],
        ["2", "Pay rises unlock with sales targets, not dates (slide §payscheme§)", "about +Rp7M by Feb, more in 2027", "Agree the scheme with staff now", "In the plan"],
        ["3", "Bill EAP 2 one quarter in advance", "about +Rp30M by Feb", "Maybe a 2-3% discount", "Do now"],
        ["4", "30% down payment on the March head-office program, paid in February", "+Rp30M by Feb", "Ask BI in the 17 October proposal", "Do now"],
        ["5", "One public event between December and February", "about +Rp12M", "Team time in pitch season", "Do now"],
        ["6", "Second founder's pay waits longer. One founder's Rp10M is the minimum and stays", "+Rp15M a month of delay", "Founders' decision", "Reserve"],
        ["7", "Smaller office or coworking space in the pitch months", "about +Rp15M", "Only if the lease allows", "Reserve"]]
fl = {(i, 4): BLUE_BG for i in (1, 2, 3, 4, 5)}
fl.update({(i, 4): CREAM for i in (6, 7)})
table(s, 0.5, 1.5, 8.3, [0.35, 3.45, 1.55, 2.15, 0.8], rows, size=11, row_h=0.56, fills=fl)
stat(s, 9.1, 1.6, 3.75, 1.3, f"{rp(REC['cash'][4])} → {rp(SN['cash'][4])}", "cash at the end of February, with step 4", vsize=24)
stat(s, 9.1, 3.05, 3.75, 1.3, f"{rp(F27['low'])} → {rp(SN['fy27']['low'])}", "lowest cash in 2027, with step 4", vcolor=TERRA, vsize=24)
stat(s, 9.1, 4.5, 3.75, 1.0, "+Rp42M", "more by February if steps 3 and 5 also work (estimate)", vsize=20, lsize=11)
callout(s, 9.1, 5.65, 3.75, 1.25, [[B("Do not cut: "), N("sales travel, trainer certification in January, BPJS and THR, and founder pay below Rp10M.")]], size=12)
text(s, 0.5, 6.15, 8.3, 0.75,
     [[B("Check first: "), N("was September's EAP invoice sent (about +Rp15M in October)? What do our 4 small corporate clients pay, and when? Can the office lease change?")]],
     size=11, color=BODY)
notes(s, "The cash drop from October to February is a sales gap, not overspending. Ads stay because this is when we need "
         "sales; they must earn back their cost (ROAS = sales from ads divided by ad cost). Getting paid earlier (EAP billed "
         "in advance, a down payment from BI) helps most. Steps 6 and 7 are a reserve if BI says no to a down payment.")

mark("cashplan")

# ================================================================== 18. Ready to scale
_mo = -(-(BUFFER - F27["end"]) // (S4["N"] / 12))
when = "end of 2027" if F27["end"] >= BUFFER else ["January", "February", "March", "April", "May", "June", "July", "August",
                                                    "September", "October", "November", "December"][int(_mo) - 1] + " 2028"
s = new_slide("Invest in new units only above a 6-month cushion", accent="6-month cushion", tag="money",
              sub=f"The cushion is about {rp(BUFFER)}. Cash is {rp(F27['end'])} at the end of 2027, so the earliest start is {when}.",
              source=CUR + " (Outlook sheet).")
line(s, 0.5, 1.5, 7.7, 5.4, M15, [("Cash in the bank (current estimate)", [round(v) for v in REC["cash"]]),
                                   ("Safety cushion: 6 months of fixed cost", [round(BUFFER)] * 15)],
     "Cash at the end of each month (Rp M)", [INK_D, TERRA])
rules = [("1", "Keep the cushion", f"About {rp(BUFFER)} stays in the bank for bad times."),
         ("2", "Invest above it", "Up to half of each year's net profit goes to new units, such as Theta and Coreitera."),
         ("3", "How much", f"A normal full year makes about {rp(S4['N'])}, so about {rp(INVEST_28)} a year for new units.")]
for i, (n, h, t) in enumerate(rules):
    y = 1.6 + i * 1.78
    card(s, 8.5, y, 4.35, 1.6, "")
    circle(s, 8.7, y + 0.2, 0.5, n, fill=[INK_D, BLUE, TERRA][i])
    text(s, 9.32, y + 0.15, 3.4, 1.35, [[(h, {"font": HEAD, "bold": True, "size": 15, "color": INK_D})], N(t)], size=12, space=3)
notes(s, "This is what 'ready to grow' means in numbers. We do not invest the cushion. We invest only what we earn above it.")

mark("scale")

# ================================================================== Divider: What do we do in the next 90 days?
divider('→', 'WHAT WE DO NOW', 'What do we do in the next 90 days?', 'Eight actions with an owner and a date, and six decisions only the founders can make.', [("17 Oct", "first deadline: BI proposal"), ("Kia + Galih", "owners until the BD co-founder joins")], ['What must happen, by whom and by when?', 'What must the founders decide?'], None)
mark("d_next")

# ================================================================== E1. Next 90 days
s = new_slide("The next 90 days: who does what, by when", accent="who does what", tag="next",
              sub="Owners for now: Kia (CEO) and Galih, until the BD co-founder joins. Review every Monday for 30 minutes.",
              source="Roadmap sheet in our model and the stop dates on the risks slide.")
rows = [["#", "Action", "Owner", "By when", "If it slips"],
        ["1", "BI: one full-year 2027 proposal; 3 events moved to Jul-Sep; 30% down payment", "Kia", "17 Oct 2026", "Escalate to the regional offices"],
        ["2", "Start the BD co-founder search", "Kia and Galih", "Oct 2026", "Founders keep selling alone"],
        ["3", "Hire the second paid part-time psychologist (starts January)", "Galih", "30 Nov 2026", "Pool of 20 covers evenings"],
        ["4", "Banks: put EAP 3 and training days into their 2027 budgets", "Kia", "31 Dec 2026", "Bank deals move to 2028"],
        ["5", "EAP 2: sign the warm lead, billed a quarter ahead", "Galih", "15 Jan 2027", "No EAP 2 by 15 Dec: delay the April pay step"],
        ["6", "Pay scheme in writing with staff; BPJS registered", "Kia", "31 Dec 2026", "No pay change without it"],
        ["7", "Booking link, session packages and referral tracking live", "Galih + admin", "31 Jan 2027", "Cut extra ads to Rp1M"],
        ["8", "Certify 1 more trainer for March; one public event", "Galih", "Feb 2027", "No BI work order by 31 Jan: March slips"]]
table(s, 0.5, 1.55, 12.3, [0.4, 5.6, 1.55, 1.45, 3.3], rows, size=12, row_h=0.58,
      fills={(i, 3): CREAM for i in (1, 3, 5)})
notes(s, "These eight actions start the plan. Items 1, 4 and 5 bring the money; 2, 3 and 8 make sure we can deliver; 6 and 7 "
         "set up pay and B2C tracking. Check them every Monday.")

mark("next90")

# ================================================================== E2. Decisions needed
s = new_slide("What we need to decide now", accent="decide now", tag="next",
              sub="Each decision changes a number in this plan. Our proposal is in each box.",
              source="Our model; founders' discussion, 5 October 2026.")
DEC = [("Second founder's pay", "Room of Rp15M a month from March; the guard rule says wait until May if February cash is under Rp80M.", "End of Feb 2027"),
       ("Sales targets for pay steps", "Rp100M, Rp150M and Rp200M a month (3-month average). Agree them with staff in writing.", "Dec 2026"),
       ("Lead Partnership hire", f"A non-psychologist who manages EAP clients from Oct 2027, about Rp{LP:g}M a month (estimate).", "Jul 2027"),
       ("Referral reward", "Rp100k session credit to each side. Our lead psychologist checks it against the HIMPSI code.", "Nov 2026"),
       ("How the pool of 20 is paid", "Per session (40% of the price), with 8 open hours a month each. Trainees cannot counsel alone.", "Oct 2026"),
       ("Down payment from BI", "Ask for 30% before the March head-office program, or invoices in stages.", "17 Oct 2026")]
for i, (h, t, when) in enumerate(DEC):
    x, y = 0.5 + (i % 3) * 4.15, 1.6 + (i // 3) * 2.7
    card(s, x, y, 3.95, 2.5, [[(h, {"font": HEAD, "bold": True, "size": 15, "color": INK_D})], N(t), [("", {"size": 4})],
                              [("Decide by: ", {"bold": True, "color": TERRA}), (when, {"bold": True, "color": TERRA})]],
         size=12, space=5, fill=SKY if i % 2 == 0 else CREAM)
notes(s, "These are the open founder decisions. Each one is a blue input in the model, so the numbers update when you decide.")

mark("decisions")

# ================================================================== Divider: Where do the numbers come from?
divider('+', 'APPENDIX', 'Where do the numbers come from?', 'The tables, rules and checks behind each slide, so anyone can test the plan.', [], ['What are the numbers behind each chart?', 'How are running cost, contracts and capacity set?', 'What are the risks?', 'What do the business words mean?', 'What is still an estimate?'], None)
mark("d_app")

# ================================================================== A1. Numbers
s = new_slide("The numbers behind the charts", sub="Three estimates for 2027, then a normal full year, one change per row.",
              tag="app", source="Our model; Excel recalculated every case. Tax is 0.5% of sales.")
EST = [("Conservative: BI buys 30% less, no bank deal", "conservative", RED_BG), ("Current: our plan", "plan", BLUE_BG),
       ("Aggressive: bank deal from July, 2 days a month", "aggressive", None),
       ("Current + guard rule", "plan_both_guards", None), ("Current + guard rule + both cash levers", "plan_all_levers", None),
       ("Current, but EAP 2 signed with a bank (April)", "plan_eap2_bank", None)]
rows2 = [["2027 estimate", "Sales", "Net profit", "Net margin", "Lowest cash", "Cash end 2027", "Months below floor"]]
fills2 = {}
for i, (lab, k, fl) in enumerate(EST, 1):
    f = RUNS[k]["fy27"]
    rows2.append([lab, f"{f['R']:,.0f}", f"{f['N']:,.0f}", pc(f['m'], 1),
                  f"{f['low']:,.0f} ({G.MONTHS[f['low_idx']].replace('-', ' ')})", f"{f['end']:,.0f}", str(f['below_floor'])])
    if fl:
        for j in range(7):
            fills2[(i, j)] = fl
table(s, 0.5, 1.5, 12.3, [4.6, 1.05, 1.15, 1.15, 1.6, 1.35, 1.4], rows2, size=11, row_h=0.29, fills=fills2)
rows = [["A normal full year (2027 sales plan)", "Sales", "Gross profit", "Fixed cost", "Net profit", "Net margin"],
        ["Old prices, today's pay", f"{G.OLD['R']:,.0f}", "", "", f"{G.OLD['N']:,.0f}", pc(G.OLD['m'], 1)]]
plain = ["New prices, today's pay", "+ fair pay, Lead Partnership, Instagram ads", "+ 2 more EAP contracts",
         "+ co-founder, but no new deal", "+ co-founder wins 10 bank training days", "+ Teman Journey pays the CTO",
         "+ second admin"]
for lab, (name, _) in zip(plain, G.LADDER):
    v = L[name]
    rows.append([lab, f"{v['R']:,.0f}", f"{v['GP']:,.0f}", f"{v['F']:,.0f}", f"{v['N']:,.0f}", pc(v['m'], 1)])
rows.append(["Full team, Theta pays the CTO", f"{S4_ADM2['R']:,.0f}", f"{S4_ADM2['GP']:,.0f}", f"{S4_ADM2['F']:,.0f}",
             f"{S4_ADM2['N']:,.0f}", pc(S4_ADM2['m'], 1)])
fills = {(i, 5): (BLUE_BG if float(rows[i][5].rstrip('%')) >= 15 else RED_BG) for i in range(1, len(rows))}
table(s, 0.5, 3.75, 12.3, [5.3, 1.4, 1.4, 1.4, 1.4, 1.4], rows, size=11, row_h=0.28, fills=fills, header_fill=BLUE)

mark("numbers")

# ================================================================== 10. Running cost by phase
s = new_slide(f"Running cost grows in steps: {rp(REC['running_total'][0])} to {rp(REC['running_total'][12])}",
              accent=f"{rp(REC['running_total'][0])} to {rp(REC['running_total'][12])}", tag="app",
              sub="Running cost = all fixed costs in one month. Each step starts after the work that pays for it is signed.",
              source=CUR + " (Pay_Plan sheet).")
grp = [("Founders", "founders", INK_D), ("Team, with BPJS and THR", "team", BLUE), ("Marketing and ads", "marketing", TERRA),
       ("Office and accountant", "office", GREY), ("Sales and delivery support", "sales_support", BEIGE)]
bar(s, 0.5, 1.5, 7.7, 5.4, [lab for _, lab in PHASES], [(n, [REC[k][i] for i, _ in PHASES]) for n, k, _ in grp],
    "Running cost per month at each step (Rp M)", colors=[c for *_, c in grp], fmt='#,##0', stacked=True, legend=True, gap=55,
    texts=[[f"{REC[k][i]:.0f}" if REC[k][i] >= 4 else "" for i, _ in PHASES] for _, k, _ in grp])
steps10 = [("Jan 2027", f"BPJS and THR; Instagram ads Rp{1 + ADS_X:g}M in total"),
           ("Mar 2027", "room for both founders (Rp25M); first new programs"),
           ("Jul 2027", f"pay step 1 (since {STEP_M[0][:3]}); sales co-founder; standby project manager"),
           ("Oct 2027", "pay step 2: admin at minimum wage; Lead Partnership; first bank days")]
for i, (d, t) in enumerate(steps10):
    card(s, 8.5, 1.6 + i * 1.33, 4.35, 1.2, [[(d, {"font": HEAD, "bold": True, "size": 15, "color": TERRA})], N(t)],
         size=12, space=2, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
notes(s, "Running cost is what we pay every month even with no events. Delivery costs (trainer and session fees) come on top, only when there is work.")

mark("runcost")

# ================================================================== 14. Contracts
s = new_slide("Better contracts in four areas", accent="four areas", tag="app",
              sub="We keep part-time pay levels. We change how the contracts work.",
              source="Session fees and day rates are estimates; compare them with the market before you sign.")
rows = [["Area", "Psychologists", "Trainers", "Project manager"],
        ["1. How we pay", "Full-time: market salary + bonus when an EAP client renews. Part-time: fee per session",
         "Day rate by level; client pays travel; repays training cost if leaving before 4 events",
         "Small monthly fee to stay available + Rp2M per project"],
        ["2. What they promise", "Part-time: 8+ hours a month on Theta; reply within 24 hours; 3 months notice",
         "Yearly agreement; at most 2 events a month; confirm 30 days before",
         "Up to 2 projects a month; we get the first call; a handover file"],
        ["3. How we check quality", "Client rating 4.5 of 5 or more; reports on time; review every 3 months",
         "Participant scores; tests before and after the training", "On time; invoice before BI's payment date; client rating"],
        ["4. How we reduce risk", "Client data kept private; senior supervision; insurance; no direct work with our clients for 12 months",
         "We own the training material; a backup trainer for each event", "A written guide; a second project manager learns one project every 3 months"]]
table(s, 0.5, 1.5, 12.3, [2.2, 3.6, 3.3, 3.2], rows, size=12, row_h=0.86)
callout(s, 0.5, 6.15, 12.3, 0.75, [[B("Why psychologists will join us: "), N("paid hours we promise, supervision and SKP points (professional credits), "
                                                                            "work with BI and OJK, and flexible hours on Theta.")]], size=13)
notes(s, "The four areas cover any service contract without overlap: how we pay, what time they promise, how we check their work, and how we protect ourselves.")

mark("contracts")

# ================================================================== A5b. Psychologist capacity and schedule (COO)
s = new_slide("Psychologist capacity: two part-timers and a pool", accent="two part-timers and a pool", tag="app",
              sub=f"Two paid part-timers cover up to {B2C['cap'][14]:.0f} sessions a month. Our pool of 20 covers peaks, Sundays and special cases.",
              source=CUR + " (B2C sheet, capacity rows). Shifts: 3 a week per part-timer, 3 sessions each. EAP: 10 sessions a month per contract.")
eap_s = [round(d - b) for d, b in zip(B2C["demand"], B2C["sess"])]
combo(s, 0.5, 1.5, 6.9, 5.4, M15, [("B2C sessions", [round(v) for v in B2C["sess"]], TERRA), ("EAP sessions", eap_s, INK_D)],
      [("Paid part-time capacity", [round(v) for v in B2C["cap"]], "576C94", True)], "Sessions to deliver each month vs paid capacity")
GRID = [["", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
        ["Day (EAP)", "A", "A", "A", "B", "B", "-", "-"],
        ["Evening 18-21 (B2C)", "B", "A", "B", "A", "Pool", "B (day)", "Pool"]]
gfill = {(i, j): (CREAM if GRID[i][j] == "Pool" else SKY if GRID[i][j].startswith("A") else BLUE_BG if GRID[i][j].startswith("B") else WHITE)
         for i in (1, 2) for j in range(1, 8)}
table(s, 7.7, 1.55, 5.15, [1.45] + [0.53] * 7, GRID, size=11, row_h=0.42, fills=gfill)
text(s, 7.7, 2.85, 5.15, 0.36, "A = lead part-timer. B = new part-timer (Jan). Pool = 20, per session.", size=10, color=MUTED)
stat(s, 7.7, 3.3, 2.5, 1.25, pc(max(B2C["load"][3:])), "busiest load on the 2 part-timers (healthy: 60-80%)", vsize=24, lsize=10)
stat(s, 10.35, 3.3, 2.5, 1.25, f"{max(B2C['demand']):.0f}", "sessions a month at peak (B2C + EAP)", vcolor=TERRA, vsize=24, lsize=10)
callout(s, 7.7, 4.7, 5.15, 2.2,
        [H("Rules for good availability", sz=14),
         [("•  ", {"color": TERRA, "bold": True}), ("Every new client gets a slot within 48 hours", {})],
         [("•  ", {"color": TERRA, "bold": True}), ("WhatsApp reply within 1 hour; booking link open 24/7 on Theta", {})],
         [("•  ", {"color": TERRA, "bold": True}), ("Each pool member opens at least 8 hours a month", {})]],
        size=12, anchor=MSO_ANCHOR.TOP, space=4)
notes(s, "B2C clients mostly book evenings and weekends; EAP sessions are mostly in office hours. So the two paid "
         "part-timers split days and evenings, and the pool takes Friday evening and Sunday. Confirm how the pool of 20 is paid (per session).")

mark("psycap")

# ================================================================== A4. Risks
s = new_slide("Risks and how we respond", accent="how we respond", tag="app",
              sub="Ways to respond: reduce the risk, share it (for example insurance), accept it, or stop the activity.",
              source="Chance and impact are our own judgement. Review this list every 3 months.")
rows = [["Risk", "Chance", "Impact", "Response", "Action", "Who watches, and when to act"],
        ["We start selling too late and miss budget windows", "High", "High", "Reduce", "Stop dates: BI proposal by 17 Oct; no EAP 2 by 15 Dec = delay the April pay step; no BI work order by 31 Jan = March slips", "Kia and Galih; Roadmap sheet monthly"],
        ["One project manager has too much work", "High", "High", "Reduce", "Standby project manager, Rp2M a project, booked by May (slide §limits§)", "Kia; 3 or more projects in a month"],
        ["A certified trainer leaves", "Medium", "High", "Reduce", "Yearly agreement; trainer repays training cost; train up to 8", "Galih; fewer trainers than the busiest month needs + 1"],
        ["A client hires our trainer directly", "Medium", "Medium", "Reduce", "No direct work with our clients for 12 months; we own the material", "Co-founder; any direct approach"],
        ["A clinical or ethics problem", "Low", "High", "Share + reduce", "Insurance; senior supervision; clear steps to escalate", "Lead psychologist; any case"],
        ["Staff data leaks (EAP, Theta)", "Low", "High", "Reduce", "Consent, access control, data agreement (data protection law)", "CTO; check every 3 months"],
        ["BI cuts its budget", "Medium", "High", "Reduce", f"Grow EAP and banks: BI share of sales from {pc(BI_S1)} to {pc(BI_S4)}", "Kia; BI below plan for 2 months"]]
risk_fill = {(i, j): RED_BG for i in (1, 2) for j in (1, 2)}
table(s, 0.5, 1.5, 12.3, [2.6, 0.95, 0.9, 1.35, 4.0, 2.5], rows, size=11, row_h=0.58, fills=risk_fill)
callout(s, 0.5, 6.3, 12.3, 0.6, [[B("Conservative estimate: "), N(f"BI buys 30% less and no bank deal. Margin {pc(RUNS['conservative']['fy27']['m'], 1)}, lowest cash "
                                                                f"{rp(RUNS['conservative']['fy27']['low'])}. If BI falls below plan, do not start the July steps.")]], size=12)

mark("risks")

# ================================================================== A3. Word list
s = new_slide("Word list", accent="Word list", tag="app", sub="The business words in this deck, in simple terms.",
              source="Example numbers use one EAP contract: the client pays Rp15M a month and the work costs us Rp5M.")
words_l = [["Word", "What it means"],
           ["Sales (revenue)", "All the money clients pay us"],
           ["Delivery cost", "What one job costs us: psychologist and trainer fees, materials"],
           ["Gross profit", "Sales minus delivery cost. One EAP contract: 15 - 5 = Rp10M a month"],
           ["Fixed cost", "What we pay every month, even with no work: salaries, office, software"],
           ["Net profit", "Gross profit minus fixed cost and tax: what we really keep"],
           ["Net margin", "Net profit as a share of sales. 15% = we keep Rp15 of every Rp100"],
           ["Running cost", "All fixed costs in one month: pay, office, software, ads"],
           ["Cash", "Money in the bank today. Profit is not cash until the client pays"],
           ["Cash cushion", "Money we keep for bad times: 6 months of fixed cost"]]
words_r = [["Word", "What it means"],
           ["Conservative, current, aggressive", "Three estimates: a bad year, our plan, a good year"],
           ["Fair pay", "At least the minimum wage, plus BPJS and THR, for full-time staff"],
           ["UMP, BPJS, THR", "Minimum wage; state health and work insurance; holiday bonus"],
           ["EAP", "A company pays us monthly; we support its staff with counseling"],
           ["Recurring income", "Money that comes every month without a new sale, like EAP"],
           ["Trigger", "The event that tells us it is time to hire"],
           ["Co-founder, shares", "A partner who owns part of the company; shares are earned over years"],
           ["Business unit", "A separate part of the business, for example Theta or Coreitera"],
           ["Framework, SPK", "A yearly agreement with a client; a work order for one job under it"]]
table(s, 0.5, 1.5, 6.0, [1.9, 4.1], words_l, size=12, row_h=0.52)
table(s, 6.8, 1.5, 6.0, [1.9, 4.1], words_r, size=12, row_h=0.52)

mark("words")

# ================================================================== A6. Model checks + estimates
s = new_slide("What the model includes, and what to confirm", accent="what to confirm", tag="app",
              sub="Every estimate below changes a number in this deck. Replace it with the real number when we have it.",
              source="Review of our financial model and the projection guide, 5 October 2026.")
left = [H("The model now includes", sz=15),
        "Real timing: no new events before March 2027",
        "Pay set month by month, with BPJS and THR",
        "New EAP, co-founder and bank days",
        "Sales, costs and cash in one chart",
        "A roadmap: the latest start for each deal and hire",
        [("", {"size": 6})],
        H("Still open", TERRA, 15),
        "Owners: Kia and Galih for now (proposal)",
        "Theta and CTO costs are in no model",
        "Is our current EAP contract still active?",
        "Will BI accept the new 2027 prices?",
        "Does the 0.5% small-business tax still apply?"]
card(s, 0.5, 1.5, 4.6, 5.4, left, size=12, space=4)
rows = [["Estimate", "Value we used", "How to check"],
        ["Second founder's pay", "Rp15M a month from March (room)", "Founders decide"],
        ["BPJS and THR on top of pay", "18% of pay", "Payroll provider quote"],
        ["Lead Partnership pay", f"Rp{LP:g}M a month from Oct 2027", "2 or 3 candidate talks"],
        ["EAP price per employee", "Rp30k a month, 500 staff", "Ask 2 bank HR heads"],
        ["EAP contract 2 start", "February 2027", "Signed by mid-January"],
        ["Sales co-founder salary", "Rp8M a month + shares", "Talks with candidates"],
        ["CTO salary", "Rp10M a month", "CTO's expectation"],
        ["Bank deal", "1 training day a month from Oct 2027", "First bank proposal"],
        ["Lead times to sign a client", "2 weeks to 6 weeks of paperwork", "Log real deals"],
        ["Income tax", "0.5% of sales", "Tax consultant"],
        ["People each trainee reaches", f"{REACH_PER_TRAINER} each", "BI ambassador reports"]]
table(s, 5.3, 1.5, 7.5, [2.7, 2.6, 2.2], rows, size=12, row_h=0.45)

mark("checks")

# ================================================================== A7. Bonus: Magang Nasional
s = new_slide("Bonus: Magang Nasional gives us one paid intern", accent="one paid intern", tag="app",
              sub="The government pays the intern. Not in our numbers: if approved, it is a bonus.",
              source="Kemnaker MagangHub (maganghub.kemnaker.go.id); Antara and Kompas, May 2026. Programme for 2027 not yet confirmed.")
cards3 = [("What it is", INK_D, ["Kemnaker's national internship (MagangHub)", "For people who graduated in the last 12 months",
                                 "In 2026 the state pays about the minimum wage (about Rp5.7M)"]),
          ("What we need", BLUE, ["A PT with NIB or WLKP, and a company NPWP", "A SIAPkerja account and an agreement with Kemnaker",
                                  "A mentor and a real learning plan", "Max 20% of our employees: 1 person"]),
          ("How we use it", TERRA, ["1 new psychology graduate for content and assessment admin",
                                    "Support to a paid person, never a replacement", "Not for counseling: that needs a licensed psychologist"])]
for i, (h, f, b) in enumerate(cards3):
    x = 0.5 + i * 4.15
    text(s, x, 1.55, 3.95, 0.6, h, size=16, bold=True, color=WHITE, fill=f, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         anchor=MSO_ANCHOR.MIDDLE, margin=0.2, font=HEAD)
    card(s, x, 2.25, 3.95, 3.8, [[("•  ", {"color": TERRA, "bold": True}), (t, {})] for t in b], size=13, space=10)
callout(s, 0.5, 6.2, 12.3, 0.7, [[B("Next step: "), N("register on MagangHub this month, then apply for the next batch. Other interns stay true learning "
                                                       "roles with a written agreement, an allowance and BPJS work-accident cover.")]], size=12)

mark("magang")

fill_refs()
assert len(prs.slides) == 40, len(prs.slides)
prs.save(OUT)
print("wrote", OUT, len(prs.slides), "slides")
