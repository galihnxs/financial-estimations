"""
Build TemanJourney_Model_2027_v16.xlsx from v15 (v15 is never changed).

Run from financial-estimations/CompanyWideAnalysis:   python analysis/build_model_v16.py
Then recalculate with analysis/lo_runner.py --save (LibreOffice) or open the file in Excel.

What v16 adds to v15 (founders' feedback, 16 Oct 2026), all as live formulas behind one switch (Inputs v16_on):
  1. Latest B2C prices (B2C Pricing/TJ_Pricing_Tier_Psikolog_Profit_Sharing.xlsx): psychologist tiers, profit
     share, platform cost and a tier mix replace v15's Rp1m session at 40% (sheet B2C_Pricing).
  2. EAP as the main channel (sheet Channels + Volume rows 38-50): 3-month EAP packages for BI regional offices
     (BI KPw, priced from the BI Sibolga costing v4), renewals, the Commit-2 scheme, paid add-on counseling,
     and trainings sold to EAP clients.
  3. Pricing scheme (sheet Pricing_Scheme): a small intro offer, list price, and Commit-2 (two terms bought upfront
     for a lower overall price), with a break-even rule for the discount.
  4. Targets by priority: BI kept at today's volume (BAU premium), LPS / OJK secondary, banks tertiary, and a new
     school channel (group interest and aptitude tests, high school to college).
  5. B2C public events priced from the Lyfe at 30s RAB (small paid format) instead of v10's Rp20m mass event.
With Inputs v16_on = 0 the workbook gives exactly the v15 results.
"""
import json
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as CL

ROOT = Path(__file__).resolve().parent.parent
SRC, OUT = ROOT / "TemanJourney_Model_2027_v15.xlsx", ROOT / "TemanJourney_Model_2027_v16.xlsx"
wb = openpyxl.load_workbook(SRC)
I, V, C, P, CF, S, T = (wb[n] for n in ["Inputs", "Volume", "Capacity_OPEX", "PnL", "CashFlow", "Summary", "To_Fill"])
PP, BC, U = wb["Pay_Plan"], wb["B2C"], wb["Unit_Economics"]
COLS = range(5, 20)                     # E..S = Oct-26 .. Dec-27
MON = ["Oct-26", "Nov-26", "Dec-26"] + [m + "-27" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()]
NORM, BOLD = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
BLUE, ITAL = Font(name="Arial", size=10, color="0000FF"), Font(name="Arial", size=9, italic=True)
TITLE = Font(name="Arial", bold=True, size=13)
HEAD, YEL = PatternFill("solid", fgColor="DDE3EE"), PatternFill("solid", fgColor="FFF2CC")


def style_like(dst, src):
    dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
    dst.alignment, dst.number_format = copy(src.alignment), src.number_format


def put(ws, r, c, v, font=NORM, fmt=None, fill=None, wrap=False):
    cell = ws.cell(r, c, v); cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if wrap:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    return cell


def header(ws, r, labels, c0=1):
    for c, t in enumerate(labels, c0):
        put(ws, r, c, t, BOLD, fill=HEAD, wrap=True)


def month_cols(ws, r):                 # month labels across E..S, plus Q4 2026 / FY 2027 in C / D
    for c in range(1, 20):
        style_like(ws.cell(r, c), V.cell(4, c))
        ws.cell(r, c).value = V.cell(4, c).value if c < 5 else f"=Volume!{CL(c)}4"
    ws.cell(r, 1, "Line"); ws.cell(r, 2, "Unit"); ws.cell(r, 3, "Q4 2026"); ws.cell(r, 4, "FY 2027")


def lag(L_col, k, row, sheet=""):      # reference to the same row k months earlier, or 0 before Oct-26
    c = L_col - k
    return f"{sheet}{CL(c)}{row}" if c >= 5 else "0"


# ------------------------------------------------------------------ 0. B2C_Pricing sheet (latest B2C price list)
BP = wb.create_sheet("B2C_Pricing", wb.sheetnames.index("B2C") + 1)
put(BP, 1, 1, "B2C prices: psychologist tiers and profit share (latest price list, 13 Oct 2026)", TITLE)
put(BP, 2, 1, "Source: B2C Pricing/TJ_Pricing_Tier_Psikolog_Profit_Sharing.xlsx (Pricing_Counseling, Inputs). Blue = copy "
              "of that file; change both together. Rp million per session. Feeds Inputs b2c_p_new and b2c_f_new.", ITAL)
header(BP, 4, ["Tier (psychologist experience)", "Online 60 min price", "Offline 60 min price", "Share to psychologist",
               "Share of sessions (mix)", "Fee online", "Fee offline", "TJ margin online", "TJ margin offline"])
TIERS = [("Tier 1 (0-2 years)", 0.500, 0.580, 0.60, 0.30), ("Tier 2 (2-5 years)", 0.500, 0.580, 0.65, 0.35),
         ("Tier 3 (5-10 years)", 0.725, 0.800, 0.65, 0.25), ("Tier 4 (over 10 years)", 0.850, 0.925, 0.70, 0.10)]
for k, (t, on, off, sh, mix) in enumerate(TIERS):
    r = 5 + k
    put(BP, r, 1, t)
    for c, v, f in ((2, on, "0.000"), (3, off, "0.000"), (4, sh, "0%"), (5, mix, "0%")):
        put(BP, r, c, v, BLUE, f, YEL if c == 5 else None)
    put(BP, r, 6, f"=B{r}*D{r}", fmt="0.000"); put(BP, r, 7, f"=C{r}*D{r}", fmt="0.000")
    put(BP, r, 8, f"=B{r}-F{r}", fmt="0.000"); put(BP, r, 9, f"=C{r}-G{r}", fmt="0.000")
put(BP, 9, 1, "Check: mix adds to 100%", ITAL); put(BP, 9, 5, '=IF(ABS(SUM(E5:E8)-1)<0.0001,"OK","CHECK")', BOLD)
BPI = [(11, "Share of sessions online (rest offline)", 0.70, "0%", "Estimate: fill from real bookings"),
       (12, "Platform and payment cost per session (paid by TJ)", 0.05, "0.000", "Estimate: pricing file default (Rp50k)"),
       (13, "4-session package discount", 0.10, "0%", "Known: ACC sheet, ~10%"),
       (14, "Share of clients on the 4-session package", "=Inputs!$C$PACK", "0%", "Inputs b2c_pack (0 = price list as is)")]
for r, t, v, f, note in BPI:
    put(BP, r, 1, t); put(BP, r, 2, v, BLUE if not str(v).startswith("=") else NORM, f, YEL if r in (11, 12) else None)
    put(BP, r, 3, note, ITAL)
OUTR = [(16, "Average list price per session", "=SUMPRODUCT(E5:E8,B5:B8)*B11+SUMPRODUCT(E5:E8,C5:C8)*(1-B11)"),
        (17, "Average psychologist fee per session", "=SUMPRODUCT(E5:E8,F5:F8)*B11+SUMPRODUCT(E5:E8,G5:G8)*(1-B11)"),
        (18, "Average price actually paid (after package discount)", "=B16*(1-B14*B13)"),
        (19, "Psychologist fee on the price paid (profit share is on the price paid)", "=B17*(1-B14*B13)"),
        (20, "TJ margin per session after platform cost", "=B18-B19-B12"),
        (21, "Cost share used by the model: (fee + platform) / price", "=(B19+B12)/B18"),
        (22, "v15 for comparison: price, and TJ margin per session", "=Inputs!$C$P_OLD"),
        (23, "v15 TJ margin per session (Rp1m x 60%)", "=Inputs!$C$P_OLD*(1-Inputs!$C$F_OLD)"),
        (24, "New margin vs v15 margin", "=B20/B23-1")]
put(BP, 15, 1, "Blended result (feeds the model)", BOLD, fill=HEAD)
for r, t, f in OUTR:
    put(BP, r, 1, t, BOLD if r in (18, 20, 21) else NORM); put(BP, r, 2, f, fmt="0%" if r in (21, 24) else "0.000")
put(BP, 26, 1, "Reference: assessments (one price for all assessors, 50:50 share after the test-tool cost)", BOLD, fill=HEAD)
for k, (t, age, pr) in enumerate([("Tes Minat Bakat (interest and aptitude)", "14+", 1.625), ("Tes IQ", "child to adult", 0.725),
                                  ("Tes Profiling", "6-13", 1.225), ("MHCU (adult mental health check)", "adult", 0.375)]):
    put(BP, 27 + k, 1, t); put(BP, 27 + k, 2, pr, BLUE, "0.000"); put(BP, 27 + k, 3, age, ITAL)
put(BP, 32, 1, "Read: the latest price list earns about Rp170k per session after the psychologist and platform, against "
               "Rp600k in v15. Counseling alone cannot pay for Instagram ads (see B2C rows 56-59); it pays when clients come "
               "through EAP, schools and referral.", ITAL)
BP.column_dimensions["A"].width = 58
for c in range(2, 10):
    BP.column_dimensions[CL(c)].width = 13

# ------------------------------------------------------------------ 1. Channels sheet: EAP packages for BI KPw + schools
CH = wb.create_sheet("Channels", wb.sheetnames.index("B2C_Pricing") + 1)
put(CH, 1, 1, "Channels (v16): EAP packages for BI regional offices, and schools", TITLE)
put(CH, 2, 1, "Package prices and costs: SPKnProposalAnalysis/BI Sibolga/Internal_-_BI_Sibolga_Costing_format_TJ_v4.xlsx "
              "(costs with every session paid at the Rp350k honor; before the 2% PPh 23 withholding). Rp million.", ITAL)
header(CH, 4, ["BI KPw package (3 months)", "Price", "Cost (all sessions paid)", "Counseling sessions", "Mix (blue)",
               "Margin", "Price / cost", "Price for a 30% margin", "Gap to that price"])
PKG = [("Sapa (intro: 15 sessions)", 7.5, 5.925, 15, 0.25), ("Tumbuh (24 sessions + MHCU + report)", 14.0, 10.70, 24, 0.35),
       ("Seperjalanan (36 sessions + 2 MHCU + webinar + 2 penyuluhan)", 23.0, 16.64, 36, 0.25),
       ("Berbagi (Seperjalanan, webinar free)", 20.0, 18.74, 36, 0.05), ("Penyuluhan (Seperjalanan, no webinar)", 20.0, 16.64, 36, 0.10)]
for k, (t, pr, co, se, mx) in enumerate(PKG):
    r = 5 + k
    put(CH, r, 1, t)
    for c, v, f in ((2, pr, "0.00"), (3, co, "0.00"), (4, se, "0"), (5, mx, "0%")):
        put(CH, r, c, v, BLUE, f, YEL if c == 5 else None)
    put(CH, r, 6, f"=1-C{r}/B{r}", fmt="0%"); put(CH, r, 7, f"=B{r}/C{r}", fmt="0.00")
    put(CH, r, 8, f"=C{r}/(1-0.3)", fmt="0.00"); put(CH, r, 9, f"=MAX(0,H{r}-B{r})", fmt="0.00")
put(CH, 10, 1, "Average package (mix-weighted)", BOLD)
for c, f in ((2, "=SUMPRODUCT($E$5:$E$9,B5:B9)"), (3, "=SUMPRODUCT($E$5:$E$9,C5:C9)"), (4, "=SUMPRODUCT($E$5:$E$9,D5:D9)"),
             (5, "=SUM(E5:E9)"), (6, "=1-C10/B10"), (7, "=B10/C10")):
    put(CH, 10, c, f, BOLD, "0%" if c in (5, 6) else "0.00")
put(CH, 11, 1, "Read: with every session paid at Rp350k the packages earn 7-27%, far below the 3x rule for training. They make "
               "money only when the paid part-timers deliver the sessions inside their fixed pay. Next offices: price "
               "Tumbuh and Seperjalanan at or above column H, and keep Sapa as the small intro.", ITAL)
put(CH, 13, 1, "Monthly plan (blue = change freely)", BOLD, fill=HEAD)
month_cols(CH, 14)
PLAN = [(15, "New BI KPw packages signed (new offices; Nov-26 = Sibolga)", "packages",
         [0, 1, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, 2, 1, 0]),
        (16, "Schools tested (group interest and aptitude test; Feb pilot, Jul grade 10, Aug-Nov grade 12)", "schools",
         [0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0])]
for r, t, u, vals in PLAN:
    put(CH, r, 1, t); put(CH, r, 2, u)
    put(CH, r, 3, f"=SUM(E{r}:G{r})", fmt="0"); put(CH, r, 4, f"=SUM(H{r}:S{r})", fmt="0")
    for j, v in enumerate(vals):
        put(CH, r, 5 + j, v, BLUE, "0", YEL)
put(CH, 18, 1, "Why these dates: BI KPw offices spend their own yearly budget, most of it in the second half; schools test "
               "grade 10 before the major choice (Jun-Jul) and grade 12 before college applications (Aug-Nov).", ITAL)
CH.column_dimensions["A"].width = 60
for c in range(2, 20):
    CH.column_dimensions[CL(c)].width = 10

# ------------------------------------------------------------------ 2. Inputs section 14
rows = [
    ("14. v16 (16 Oct 2026): EAP as the main channel, pricing scheme, schools, latest B2C prices", None),
    ("v16_on", "v16 changes on (1 = yes, 0 = exactly v15)", 1, "switch", "Your plan - confirm", "Turns on every v16 line below."),
    ("b2c_p_new", "B2C counseling: average price paid per session (latest price list)", "=B2C_Pricing!$B$18", "Rp m / session",
     "Known", "B2C_Pricing sheet: tier prices x tier mix x online share."),
    ("b2c_f_new", "B2C counseling: psychologist fee + platform cost, % of price", "=B2C_Pricing!$B$21", "%", "Known",
     "Profit share 60/65/65/70% by tier + Rp50k platform cost per session."),
    ("b2c_p_old", "v15 B2C counseling price (kept for the v15 check)", 1.0, "Rp m / session", "Known", "v10 file."),
    ("b2c_f_old", "v15 B2C counseling cost share (kept for the v15 check)", 0.4, "%", "Known", "v10 file."),
    ("b2c_pack", "Share of B2C clients on the 4-session package (10% off)", 0.0, "%", "Your plan - confirm",
     "0 = price list as is. Test with more sessions per client (B2C row 8)."),
    ("ev_p_new", "B2C public event revenue (Lyfe at 30s format: Rp259k ticket x 16 people)", 4.144, "Rp m / event", "Estimate",
     "Lyfe at 30s final proposal: 12-20 people at Rp259k."),
    ("ev_c_new", "B2C public event cost (Lyfe RAB at 16 people)", 2.73, "Rp m / event", "Estimate", "Psychologist, kit, F&B, ads cap."),
    ("ev_p_old", "v15 B2C mass event revenue (kept for the v15 check)", 20.0, "Rp m / event", "Known", "v10 file: 50 people."),
    ("ev_c_old", "v15 B2C mass event cost (kept for the v15 check)", 8.0, "Rp m / event", "Known", "v10 file."),
    ("lyfe_m", "Lyfe at 30s collab event (signed)", "Oct-26", "month", "Known", "Sat 31 Oct 2026, Kanteen n Co."),
    ("kpw_price", "BI KPw EAP package: average price (3 months)", "=Channels!$B$10", "Rp m / package", "Estimate",
     "Channels sheet: BI Sibolga v4 packages x mix."),
    ("kpw_cost", "BI KPw EAP package: average cost, every session paid", "=Channels!$C$10", "Rp m / package", "Estimate",
     "Rp350k honor + Rp35k admin per session, MHCU, report, contract."),
    ("kpw_sess", "BI KPw EAP package: counseling sessions", "=Channels!$D$10", "sessions", "Estimate", "Delivered over 3 months."),
    ("commit_s", "Commit-2: share of buyers who buy two terms upfront", 0.4, "%", "Estimate",
     "Pricing_Scheme sheet. Applies to KPw packages in the model."),
    ("commit_d", "Commit-2: cash discount on the overall price of the two terms", 0.05, "%", "Your plan - confirm",
     "Plus 2 free penyuluhan (bonus below). 10% only pays when the package earns 35%+ (Pricing_Scheme)."),
    ("bonus_c", "Commit-2 bonus: TJ cost of 2 free 30-min penyuluhan", 0.2, "Rp m / deal", "Estimate",
     "BI Sibolga v4: Rp100k materials each, hosted in-house."),
    ("bonus_v", "Commit-2 bonus: value to the client of those 2 penyuluhan", 2.0, "Rp m / deal", "Estimate",
     "BI Sibolga v4 values one penyuluhan at Rp1m."),
    ("renew_r", "Renewal without Commit-2: share who buy a second term later at full price", 0.3, "%", "Estimate",
     "Replace with real renewals after the first 3 offices."),
    ("addon", "Paid add-on counseling a month per active KPw package (B2C price)", 1.0, "sessions", "Estimate",
     "Staff and family who pay for more sessions after the quota."),
    ("upsell", "Share of new EAP clients who buy one training within 3 months", 0.25, "%", "Estimate",
     "Webinar or penyuluhan inside the package opens the training sale."),
    ("eap_bi", "Existing EAP contract is with BI? (1 = yes)", 1, "switch", "Your plan - confirm", "For the BI share."),
    ("lps_start", "LPS / OJK programs: first month", "Jul-27", "month", "Your plan - confirm",
     "Their 2027 budget is set; mid-year revisions are the realistic window."),
    ("lps_rate", "LPS / OJK programs a month from then", 0.34, "events", "Estimate", "About one a quarter, priced as a regional ToT."),
    ("sch_price", "School test (Tes Minat Bakat, group format): price per student", 0.3, "Rp m / student", "Estimate",
     "Individual test is Rp1.625m; group format with short individual report."),
    ("sch_tool", "School test: test-tool cost per student", 0.07, "Rp m / student", "Estimate", "MHCU tool cost used as a proxy."),
    ("sch_share", "School test: assessor share of (price - tool cost)", 0.5, "%", "Known", "Pricing file: tests 50:50."),
    ("sch_pax", "School test: students per school", 80, "students", "Estimate", "About 2-3 classes of one grade."),
    ("sch_up", "School test: students who book a 1:1 follow-up session (B2C price)", 0.05, "%", "Estimate", "Feeds B2C sessions."),
    ("sch_s0", "School test: share paid in the month of the test (rest next month)", 0.5, "%", "Estimate", "Parents or school."),
]
r0 = 130
first_sect = I["A84"]
ADDR = {}
for k, row in enumerate(rows):
    r = r0 + k
    if row[1] is None:
        I.cell(r, 1, row[0]); style_like(I.cell(r, 1), first_sect)
        continue
    key, item, val, unit, status, src = row
    for c, v in enumerate([key, item, val, unit, status, src], 1):
        cell = I.cell(r, c, v)
        style_like(cell, I.cell(86 if isinstance(val, str) and not val.startswith("=") else 97, c))
    if isinstance(val, str) and val.startswith("="):
        I.cell(r, 3).font = Font(name="Arial", size=10, color="000000")
    if unit == "%":
        I.cell(r, 3).number_format = "0%"
    elif unit in ("switch", "students"):
        I.cell(r, 3).number_format = "0"
    elif unit != "month":
        I.cell(r, 3).number_format = "0.000"
    ADDR[key] = f"C{r}"
A = lambda k: "Inputs!$C$" + ADDR[k][1:]
for r in range(11, 25):                 # fill the placeholders on B2C_Pricing
    v = BP.cell(r, 2).value
    if isinstance(v, str):
        BP.cell(r, 2).value = (v.replace("$C$PACK", "$C$" + ADDR["b2c_pack"][1:]).replace("$C$P_OLD", "$C$" + ADDR["b2c_p_old"][1:])
                               .replace("$C$F_OLD", "$C$" + ADDR["b2c_f_old"][1:]))
ON = A("v16_on")
# v15 price cells become switches between the v15 and v16 values (rows 13, 14, 20, 21 keep their keys)
I["C13"] = f"=IF({ON}=1,{A('ev_p_new')},{A('ev_p_old')})"
I["C14"] = f"=IF({ON}=1,{A('b2c_p_new')},{A('b2c_p_old')})"
I["C20"] = f"=IF({ON}=1,{A('ev_c_new')},{A('ev_c_old')})"
I["C21"] = f"=IF({ON}=1,{A('b2c_f_new')},{A('b2c_f_old')})"
for r in (13, 14, 20, 21):
    I.cell(r, 3).font = Font(name="Arial", size=10, color="000000")
I["F13"] = "v16: Lyfe at 30s format (Inputs section 14). v15: your v10 file."
I["F14"] = "v16: latest price list (B2C_Pricing sheet). v15: Rp1,000,000 per session."
I["F20"] = "v16: Lyfe RAB at 16 people. v15: venue, ads, materials, fees."
I["F21"] = "v16: psychologist share by tier + platform cost (B2C_Pricing). v15: Rp400k of Rp1m."

# ------------------------------------------------------------------ 3. Volume rows 38-50
hdr = V["A33"]
V["A38"] = "v16: EAP channel, LPS / OJK, schools"; style_like(V["A38"], hdr)
VR = [(39, "New BI KPw EAP packages (new offices)", "packages", "SUM"),
      (40, "Commit-2 second terms (bought upfront with the first)", "packages", "SUM"),
      (41, "Renewals at full price (no Commit-2)", "packages", "SUM"),
      (42, "Active BI KPw packages (3 months each)", "packages", "MAX"),
      (43, "BI KPw counseling sessions delivered", "sessions", "SUM"),
      (44, "Trainings sold to EAP clients: BI KPw", "events", "SUM"),
      (45, "Trainings sold to EAP clients: new enterprise contracts", "events", "SUM"),
      (46, "LPS / OJK programs (secondary target)", "events", "SUM"),
      (47, "Schools tested", "schools", "SUM"),
      (48, "Students tested", "students", "SUM"),
      (49, "Paid counseling through channels (KPw add-ons + school follow-ups)", "sessions", "SUM"),
      (50, "BI KPw package value signed this month (Rp m)", "Rp m", "SUM")]
for r, t, u, agg in VR:
    V.cell(r, 1, t); V.cell(r, 2, u)
    for c in (1, 2):
        style_like(V.cell(r, c), V.cell(29, c))
    V.cell(r, 3, f"={agg}(E{r}:G{r})"); V.cell(r, 4, f"={agg}(H{r}:S{r})")
    style_like(V.cell(r, 3), V["C29"]); style_like(V.cell(r, 4), V["D29"])
    V.cell(r, 3).number_format = V.cell(r, 4).number_format = "0.0"
s, d, rr = A("commit_s"), A("commit_d"), A("renew_r")
for c in COLS:
    L = CL(c)
    new_ent = (f"MAX(0,{lag(c, 3, 35)}-{lag(c, 4, 35)})" if c - 3 >= 5 else "0")
    F = {39: f"={ON}*Channels!{L}15",
         40: f"={lag(c, 3, 39)}*{s}",
         41: f"={lag(c, 3, 39)}*(1-{s})*{rr}",
         42: "=" + "+".join(f"{lag(c, k, 39)}+{lag(c, k, 40)}+{lag(c, k, 41)}" for k in range(3)).replace("+0+0+0", ""),
         43: f"={L}42*{A('kpw_sess')}/3",
         44: f"={ON}*{A('upsell')}*{lag(c, 3, 39)}*(1-{L}34)",
         45: f"={ON}*{A('upsell')}*{new_ent}*(1-{L}34)",
         46: f"={ON}*IF({L}6>=IFERROR(MATCH({A('lps_start')},$E$4:$S$4,0),99),{A('lps_rate')},0)*(1-{L}34)",
         47: f"={ON}*Channels!{L}16",
         48: f"={L}47*{A('sch_pax')}",
         49: f"={ON}*({L}42*{A('addon')}+{L}48*{A('sch_up')})",
         50: f"={A('kpw_price')}*({L}39*(1-{s}*{d})+{L}40*(1-{d})+{L}41)"}
    for r, f in F.items():
        V[f"{L}{r}"] = f; style_like(V[f"{L}{r}"], V[f"{L}29"]); V[f"{L}{r}"].number_format = "0.0"
    # Lyfe at 30s collab event (Oct-26) on top of the planned public events
    V[f"{L}27"] = V[f"{L}27"].value + f"+{ON}*IF({L}6=IFERROR(MATCH({A('lyfe_m')},$E$4:$S$4,0),99),1,0)"
V["A31"] = ("Notes: v15 adds the timing switch and the growth plan (Inputs sections 11 and 12). v16 adds the EAP channel, "
            "LPS / OJK and schools (rows 38-50, Inputs section 14, Channels sheet).")

# ------------------------------------------------------------------ 4. PnL rows 46-58
P["A46"] = "v16 new lines: EAP channel, LPS / OJK, schools"; style_like(P["A46"], P["A26"])
PR = [(47, "Revenue: BI KPw EAP packages (spread over 3 months)"), (48, "Revenue: trainings sold to EAP clients"),
      (49, "Revenue: LPS / OJK programs"), (50, "Revenue: school tests"), (51, "Total v16 revenue (in row 12)"),
      (53, "Delivery cost: BI KPw EAP packages"), (54, "Delivery cost: trainings sold to EAP clients"),
      (55, "Delivery cost: LPS / OJK programs"), (56, "Delivery cost: school tests (tool + assessor share)"),
      (57, "Total v16 delivery cost (in row 21)")]
for r, t in PR:
    P.cell(r, 1, t); P.cell(r, 2, "Rp m")
    src = 12 if r in (51, 57) else 6
    for c in range(1, 20):
        style_like(P.cell(r, c), P.cell(src, c))
    P.cell(r, 1).value = t; P.cell(r, 2).value = "Rp m"
    P.cell(r, 3, f"=SUM(E{r}:G{r})"); P.cell(r, 4, f"=SUM(H{r}:S{r})")
for c in COLS:
    L = CL(c)
    kpw_p = f"IF(Volume!{L}11=1,Unit_Economics!$B$5,Unit_Economics!$N$5)"
    kpw_c = f"IF(Volume!{L}11=1,Unit_Economics!$F$5,Unit_Economics!$O$5)"
    val = "+".join(f"Volume!{lag(c, k, 50)}" if c - k >= 5 else "0" for k in range(3))
    F = {47: f"=({val})/3",
         48: f"=(Volume!{L}44+Volume!{L}45)*{kpw_p}",
         49: f"=Volume!{L}46*{kpw_p}",
         50: f"=Volume!{L}48*{A('sch_price')}",
         51: f"=SUM({L}47:{L}50)",
         53: f"=Volume!{L}42*{A('kpw_cost')}/3+Volume!{L}39*{s}*{A('bonus_c')}",
         54: f"=(Volume!{L}44+Volume!{L}45)*{kpw_c}",
         55: f"=Volume!{L}46*{kpw_c}",
         56: f"=Volume!{L}48*({A('sch_tool')}+{A('sch_share')}*({A('sch_price')}-{A('sch_tool')}))",
         57: f"=SUM({L}53:{L}56)"}
    for r, f in F.items():
        P[f"{L}{r}"] = f
    P[f"{L}12"] = f"=SUM({L}6:{L}11)+{L}51"
    P[f"{L}21"] = f"=SUM({L}15:{L}20)+{L}57"
    eapbi = f"IF(Volume!{L}26=0,0,{L}9*Volume!{L}18/Volume!{L}26*{A('eap_bi')})"
    P[f"{L}40"] = (f"={L}6+{L}7*Inputs!$C$83+{ON}*({eapbi}+{L}47"
                   f"+IF(Volume!{L}44+Volume!{L}45=0,0,{L}48*Volume!{L}44/(Volume!{L}44+Volume!{L}45)))")
    P[f"{L}42"] = f"=SUM({L}6:{L}9)+{L}51"
    P[f"{L}44"] = f"=(Volume!{L}28+Volume!{L}26*Inputs!$C$68+Volume!{L}43)*Inputs!$C$67"
P["A12"] = "Total revenue (incl. v16 lines, row 51)"
P["A21"] = "Total delivery cost (incl. v16 lines, row 57)"
P["A40"] = "Revenue from BI (Flagship + BI share of KPW; v16 adds BI EAP, KPw packages and their trainings)"
P["A42"] = "B2B revenue (incl. EAP channel, LPS / OJK and schools)"

# ------------------------------------------------------------------ 5. CashFlow rows 31-35
CF["A31"] = "v16 cash in"; style_like(CF["A31"], CF["A5"])
CR = [(32, "BI KPw packages: 50% at signing, 50% two months later"), (33, "Trainings to EAP clients + LPS / OJK (payday rule)"),
      (34, "Schools"), (35, "Total v16 cash in (in row 11)")]
for r, t in CR:
    for c in range(1, 20):
        style_like(CF.cell(r, c), CF.cell(11 if r == 35 else 8, c))
    CF.cell(r, 1, t); CF.cell(r, 2, "Rp m")
    CF.cell(r, 3, f"=SUM(E{r}:G{r})"); CF.cell(r, 4, f"=SUM(H{r}:S{r})")
for c in COLS:
    L = CL(c)
    CF[f"{L}32"] = f"=0.5*Volume!{L}50+0.5*{lag(c, 2, 50, 'Volume!') if c - 2 >= 5 else '0'}"
    prev = f"+(PnL!{CL(c - 1)}48+PnL!{CL(c - 1)}49)*(1-Inputs!$C$76)" if c > 5 else ""
    CF[f"{L}33"] = f"=(PnL!{L}48+PnL!{L}49)*Inputs!$C$76{prev}"
    prev = f"+PnL!{CL(c - 1)}50*(1-{A('sch_s0')})" if c > 5 else ""
    CF[f"{L}34"] = f"=PnL!{L}50*{A('sch_s0')}{prev}"
    CF[f"{L}35"] = f"=SUM({L}32:{L}34)"
    CF[f"{L}11"] = f"=SUM({L}6:{L}10)+{L}35"

# ------------------------------------------------------------------ 6. Load: trainers, projects, invoices, admin, psychologists
for c in COLS:
    L = CL(c)
    C[f"{L}6"] = C[f"{L}6"].value + f"+(Volume!{L}44+Volume!{L}45+Volume!{L}46)*Inputs!$C$27"
    C[f"{L}7"] = C[f"{L}7"].value + f"+(Volume!{L}44+Volume!{L}45+Volume!{L}46)*Inputs!$C$27*Inputs!$C$31"
    C[f"{L}18"] = C[f"{L}18"].value + f"+Volume!{L}39+Volume!{L}40+Volume!{L}41+Volume!{L}44+Volume!{L}45+Volume!{L}46+Volume!{L}47"
    C[f"{L}43"] = f"=Volume!{L}29+Volume!{L}44+Volume!{L}45+Volume!{L}46+Volume!{L}47"
    old = f"(Volume!{L}28+Volume!{L}26*Inputs!$C$68)"
    PP[f"{L}6"] = PP[f"{L}6"].value.replace(old, f"(Volume!{L}28+Volume!{L}26*Inputs!$C$68+Volume!{L}43)")
    BC[f"{L}26"] = f"={L}24+{L}25+{L}52+Volume!{L}49"
    BC[f"{L}41"] = f"={L}26+Volume!{L}26*Inputs!$C$68+Volume!{L}43"
C["A43"] = "B2B projects this month (one event, training or school = one project)"
BC["A26"] = "Paid B2C counseling sessions (Instagram + vouchers + referral + channels)"
BC["A41"] = "COO: sessions to deliver (B2C + EAP + BI KPw packages)"

# ------------------------------------------------------------------ 7. Pricing_Scheme sheet
PS = wb.create_sheet("Pricing_Scheme", wb.sheetnames.index("Rate_Card") + 1)
put(PS, 1, 1, "Pricing scheme: a small intro, the list price, and Commit-2 (two terms bought upfront, lower overall price)", TITLE)
put(PS, 2, 1, "First principle: a discount is worth giving only if it buys something worth more than it costs: a second sale "
              "we would otherwise lose, cash earlier, and one less sales cycle. Rp million.", ITAL)
put(PS, 4, 1, "Scheme inputs (from Inputs)", BOLD, fill=HEAD)
put(PS, 5, 1, "Commit-2 discount on the overall price"); put(PS, 5, 3, f"={d}", fmt="0%")
put(PS, 6, 1, "Chance a buyer comes back for a second term without Commit-2 (renewal)"); put(PS, 6, 3, f"={rr}", fmt="0%")
put(PS, 7, 1, "Commit-2 bonus: 2 free penyuluhan, TJ cost / value to client"); put(PS, 7, 3, f"={A('bonus_c')}", fmt="0.0")
put(PS, 7, 4, f"={A('bonus_v')}", fmt="0.0")
header(PS, 8, ["Product", "Intro offer (small scope, full unit price)", "Intro price", "List price (one term or event)",
               "Delivery cost", "Commit-2 price (two terms)", "Client saves (cash + bonus value)", "Profit: one term at list",
               "Profit: Commit-2 (two terms, after bonus cost)", "Break-even: Commit-2 wins if renewal is below",
               "Largest cash discount that still wins", "Verdict"])
PROD = [("BI KPw EAP package (3 months)", "Sapa: 15 sessions, no screening", "=Channels!$B$5", f"={A('kpw_price')}", f"={A('kpw_cost')}"),
        ("Regional ToT (BI KPw, LPS, OJK)", "2-hour penyuluhan or webinar for staff", 5.0, "=Unit_Economics!$B$5", "=Unit_Economics!$F$5"),
        ("State bank training day", "Half-day workshop for one branch", "=Unit_Economics!$B$6/2", "=Unit_Economics!$B$6", "=Unit_Economics!$F$6"),
        ("Enterprise EAP (6 months)", "3-month pilot, same monthly price", "=Unit_Economics!$B$7*3", "=Unit_Economics!$B$7*6",
         "=Unit_Economics!$C$7*6"),
        ("School test (one school, one grade)", "One class of 30 students", f"=30*{A('sch_price')}", f"={A('sch_pax')}*{A('sch_price')}",
         f"={A('sch_pax')}*({A('sch_tool')}+{A('sch_share')}*({A('sch_price')}-{A('sch_tool')}))"),
        ("Flagship ToT (BI head office)", "None: sell on the past contract", None, "=Unit_Economics!$B$4", "=Unit_Economics!$F$4")]
for k, (name, intro, ip, lp, co) in enumerate(PROD):
    r = 9 + k
    put(PS, r, 1, name); put(PS, r, 2, intro, wrap=True)
    put(PS, r, 3, ip, BLUE if isinstance(ip, float) else NORM, "0.0"); put(PS, r, 4, lp, fmt="0.0"); put(PS, r, 5, co, fmt="0.0")
    put(PS, r, 6, f"=2*D{r}*(1-$C$5)", fmt="0.0")
    put(PS, r, 7, f'=TEXT((2*D{r}-F{r}+$D$7)/(2*D{r}),"0%")&" (Rp"&TEXT(2*D{r}-F{r}+$D$7,"0.0")&"m)"')
    put(PS, r, 8, f"=D{r}-E{r}", fmt="0.0"); put(PS, r, 9, f"=F{r}-2*E{r}-$C$7", fmt="0.0")
    put(PS, r, 10, f"=IF(H{r}<=0,0,I{r}/H{r}-1)", fmt="0%")
    put(PS, r, 11, f"=1-(E{r}+(1+$C$6)*(D{r}-E{r})/2+$C$7/2)/D{r}", fmt="0%")
    verdict = ('"No Commit discount: BI never lets an old price rise again. Ask a 30% down payment instead."' if "Flagship" in name
               else f'IF(K{r}<0.03,"Price too thin for any discount: raise the list price first",IF($C$5<=K{r},"Offer Commit-2: "&TEXT($C$5,"0%")&" off + 2 free penyuluhan, 50% at signing","Cash discount too deep: max "&TEXT(K{r},"0%")))')
    put(PS, r, 12, "=" + verdict, wrap=True)
RULES = ["Rules for every offer:",
         "1. Intro = smaller scope at the full unit price, never a lower unit price. BI and other regulators will not accept a "
         "price rise on an item they bought before (Rate_Card rule 7), so a cheap unit price becomes the permanent price.",
         "2. The intro fee is credited to Commit-2 if the client signs within 60 days. The intro is a step to the bigger sale.",
         "3. Commit-2 = two terms (or two events) in one SPK, both dates fixed, 50% paid at signing. Give 'cheaper overall' "
         "as 5% off + 2 free penyuluhan (the client sees about 11% of value; TJ gives about 6%). Never above column K.",
         "4. Floor: after the discount each term must still earn at least 30% (EAP) or 3x the external cost (training).",
         "5. BI head office: no Commit discount; ask a 30% down payment. Regional offices: Commit-2 on packages is fine "
         "(one office, one budget year).",
         "6. B2C: the entry price is the Tier 1-2 session (Rp500k online). The 4-session package (10% off) is the B2C "
         "Commit: offer it after the first session (Inputs b2c_pack).",
         "How to read column J: if fewer buyers than this would come back on their own, Commit-2 earns more than "
         "selling one term at a time. Column K is the deepest discount that still wins at today's renewal guess."]
for k, t in enumerate(RULES):
    put(PS, 17 + k, 1, t, BOLD if k == 0 else NORM)
PS.column_dimensions["A"].width = 36; PS.column_dimensions["B"].width = 30; PS.column_dimensions["L"].width = 46
for c in range(3, 12):
    PS.column_dimensions[CL(c)].width = 12
PS.row_dimensions[8].height = 54
for r in range(9, 15):
    PS.row_dimensions[r].height = 30

# ------------------------------------------------------------------ 8. Segments sheet: revenue by target priority + 2027 targets
SG = wb.create_sheet("Segments", 3)
put(SG, 1, 1, "Targets by priority: where 2027 revenue comes from (Rp million)", TITLE)
put(SG, 2, 1, "Priority (founders, 16 Oct 2026): 1. EAP as the channel to counseling and training. BI stays at today's volume "
              "as a premium client (BAU). 2. LPS and OJK. 3. Banks (private banks last). New: schools. Rows add up to PnL row 12.", ITAL)
month_cols(SG, 4)
p_cc = "Inputs!$C$14"
SEG = [(5, "1. EAP channel (primary): KPw packages, new EAP contracts, their trainings, channel counseling",
        f"=PnL!{{L}}47+PnL!{{L}}48+IF(Volume!{{L}}26=0,0,PnL!{{L}}9*(Volume!{{L}}35+Volume!{{L}}18*(1-{A('eap_bi')}*{ON}))/Volume!{{L}}26)"
        f"+Volume!{{L}}49*{p_cc}"),
       (6, "BI BAU premium: Flagship, BI regional ToT, existing EAP (kept at today's volume)",
        f"=PnL!{{L}}6+PnL!{{L}}7*Inputs!$C$83+IF(Volume!{{L}}26=0,0,PnL!{{L}}9*Volume!{{L}}18*{A('eap_bi')}*{ON}/Volume!{{L}}26)"),
       (7, "2. LPS / OJK (secondary)", "=PnL!{L}7*(1-Inputs!$C$83)+PnL!{L}49"),
       (8, "3. Banks: state and private (tertiary)", "=PnL!{L}8"),
       (9, "New: schools", "=PnL!{L}50"),
       (10, "B2C direct: Instagram, referral, vouchers, public events", f"=PnL!{{L}}10+PnL!{{L}}11-Volume!{{L}}49*{p_cc}"),
       (11, "Total", "=SUM({L}5:{L}10)"),
       (12, "Check: difference to PnL total revenue (must be 0)", "=ROUND({L}11-PnL!{L}12,6)"),
       (14, "Share: EAP channel", "=IF({L}11=0,0,{L}5/{L}11)"), (15, "Share: BI BAU", "=IF({L}11=0,0,{L}6/{L}11)"),
       (16, "Share: LPS / OJK", "=IF({L}11=0,0,{L}7/{L}11)"), (17, "Share: banks", "=IF({L}11=0,0,{L}8/{L}11)"),
       (18, "Share: schools", "=IF({L}11=0,0,{L}9/{L}11)"), (19, "Share: B2C direct", "=IF({L}11=0,0,{L}10/{L}11)"),
       (20, "All money from BI (BI BAU + KPw packages + their trainings): concentration risk", "=PnL!{L}40"),
       (21, "Share of revenue from BI", "=IF({L}11=0,0,{L}20/{L}11)")]
for r, t, f in SEG:
    pct = r in (14, 15, 16, 17, 18, 19, 21)
    put(SG, r, 1, t, BOLD if r in (5, 11, 21) else NORM); put(SG, r, 2, "%" if pct else "Rp m")
    for c in COLS:
        put(SG, r, c, f.format(L=CL(c)), fmt="0%" if pct else "0.0")
    if pct:
        num = {14: 5, 15: 6, 16: 7, 17: 8, 18: 9, 19: 10, 21: 20}[r]
        put(SG, r, 3, f"=IF(C11=0,0,C{num}/C11)", fmt="0%"); put(SG, r, 4, f"=IF(D11=0,0,D{num}/D11)", fmt="0%")
    else:
        put(SG, r, 3, f"=SUM(E{r}:G{r})", fmt="0.0"); put(SG, r, 4, f"=SUM(H{r}:S{r})", fmt="0.0")
put(SG, 13, 1, "Mix", BOLD, fill=HEAD)
put(SG, 23, 1, "2027 targets: the numbers to track every month", BOLD, fill=HEAD)
header(SG, 24, ["Priority", "Target for 2027", "Unit", "Plan (FY 2027)", "Owner"])
TGT = [("1. EAP channel", "New BI KPw packages signed", "packages", "=Volume!D39", "Kia"),
       ("1. EAP channel", "Commit-2 second terms + renewals", "packages", "=Volume!D40+Volume!D41", "Kia"),
       ("1. EAP channel", "Enterprise EAP contracts running in Dec 2027 (incl. existing)", "contracts", "=Volume!S26", "Galih"),
       ("1. EAP channel", "Trainings sold to EAP clients", "events", "=Volume!D44+Volume!D45", "Kia"),
       ("1. EAP channel", "Paid counseling sessions through channels", "sessions", "=Volume!D49", "Galih"),
       ("BI BAU", "Flagship + BI regional ToT (same as v15)", "events", "=Volume!D23+Volume!D24*Inputs!$C$83", "Kia"),
       ("2. LPS / OJK", "LPS / OJK programs (incl. non-BI regional ToT)", "events", "=Volume!D46+Volume!D24*(1-Inputs!$C$83)", "Kia, then BD co-founder"),
       ("3. Banks", "Bank training days and events", "events", "=Volume!D25", "Kia, then BD co-founder"),
       ("New: schools", "Schools tested / students", "schools", "=Volume!D47", "Galih's wife (schools lead)"),
       ("New: schools", "Students tested", "students", "=Volume!D48", "Galih's wife (schools lead)"),
       ("B2C direct", "Paid B2C sessions (all sources)", "sessions", "=B2C!D26", "Galih (lead marketing)"),
       ("Result", "BI share of revenue (all BI money)", "%", "=D21", "Kia"),
       ("Result", "Net margin FY 2027", "%", "=PnL!D36", "Both founders")]
for k, (pr, t, u, f, o) in enumerate(TGT):
    r = 25 + k
    put(SG, r, 1, pr); put(SG, r, 2, t); put(SG, r, 3, u)
    put(SG, r, 4, f, BOLD, "0%" if u == "%" else "0.0"); put(SG, r, 5, o, BLUE)
SG.column_dimensions["A"].width = 70; SG.column_dimensions["B"].width = 44
for c in range(3, 20):
    SG.column_dimensions[CL(c)].width = 9
SG.column_dimensions["E"].width = 14
SG.freeze_panes = "C5"

# ------------------------------------------------------------------ 9. To_Fill: v16 estimates
last = T.max_row
new_fill = [("v16 B2C", "Tier mix and online share of real bookings", "B2C_Pricing!E5", "Sets the average B2C price."),
            ("v16 B2C", "Platform and payment cost per session", "B2C_Pricing!B12", "Taken from every session."),
            ("v16 B2C", "Real public-event headcount and cost (after 31 Oct)", "Inputs!" + ADDR["ev_p_new"], "Lyfe at 30s is the only data point."),
            ("v16 EAP", "Package mix BI KPw offices choose", "Channels!E5", "Sets the average package price."),
            ("v16 EAP", "New BI KPw packages per month", "Channels!E15", "Main v16 growth line."),
            ("v16 EAP", "Commit-2 take-up and renewal without it", "Inputs!" + ADDR["commit_s"], "Decides if the scheme pays."),
            ("v16 EAP", "Paid add-on sessions per package and training upsell", "Inputs!" + ADDR["upsell"], "The 'channel' effect."),
            ("v16 EAP", "Is the existing EAP contract with BI?", "Inputs!" + ADDR["eap_bi"], "BI share."),
            ("v16 LPS/OJK", "LPS / OJK start month and programs a month", "Inputs!" + ADDR["lps_start"], "Secondary target."),
            ("v16 Schools", "School price per student, tool cost, students per school", "Inputs!" + ADDR["sch_price"], "New channel."),
            ("v16 Schools", "Schools per month", "Channels!E16", "Calendar-driven."),
            ("v16 Tax", "Surat Keterangan PP 55 so BI withholds 0.5% final tax, not 2% PPh 23", "Inputs!C80", "1.5% of every BI invoice in cash.")]
for k, (pil, item, addr, why) in enumerate(new_fill):
    r = last + 1 + k
    vals = [int(T.cell(last, 1).value) + 1 + k, pil, item, f"={addr}", "", why, addr, None, "Open"]
    for c, v in enumerate(vals, 1):
        T.cell(r, c, v); style_like(T.cell(r, c), T.cell(last, c))

# ------------------------------------------------------------------ 10. Summary: v16 focus block + scenario table from the runs
S["A1"] = "Teman Journey - 2027 plan (v16): EAP as the channel, BI kept as a premium client, new school channel"
S["A44"] = "E. Where the money comes from (v16 targets, Segments sheet)"; style_like(S["A44"], S["A26"])
for k, (t, f, u, m) in enumerate([("EAP channel share of 2027 revenue", "=Segments!D14", "%", "Primary target."),
                                  ("BI BAU share of 2027 revenue", "=Segments!D15", "%", "Kept at today's volume."),
                                  ("LPS / OJK share", "=Segments!D16", "%", "Secondary."),
                                  ("Banks share", "=Segments!D17", "%", "Tertiary."),
                                  ("Schools share", "=Segments!D18", "%", "New."),
                                  ("B2C direct share", "=Segments!D19", "%", "Instagram, referral, events."),
                                  ("All money from BI (incl. KPw packages)", "=Segments!D21", "%", "Concentration: lower is safer.")]):
    r = 45 + k
    for c, v in enumerate([t, f, u, m], 1):
        S.cell(r, c, v); style_like(S.cell(r, c), S.cell(27, c))
    S.cell(r, 2).number_format = "0%"
RUNS_JSON = ROOT / "output" / "v16_runs.json"            # written by lo_runner.py; rebuild after a new run
if RUNS_JSON.exists():
    runs = json.loads(RUNS_JSON.read_text())
    cols = [("v15 plan (check)", "v15_check"), ("v16 plan", "plan"), ("v16 recommended (new prices + cash guard)", "recommended"),
            ("v16 conservative", "conservative"), ("v16 aggressive", "aggressive")]
    S["A34"] = "Scenario comparison (v16 values; refresh with analysis/lo_runner.py, then rebuild)"
    for j, (lab_, k) in enumerate(cols):
        S.cell(35, 2 + j, lab_)
        if k not in runs:
            continue
        f = runs[k]["fy27"]
        for r, v in [(36, f["R"]), (37, f["N"]), (38, f["m"]), (39, f["low"]), (40, MON[f["low_idx"]]),
                     (41, f["below_floor"]), (42, f["end"])]:
            S.cell(r, 2 + j, v)

wb.save(OUT)
print("wrote", OUT.name, "| v16 inputs:", ", ".join(f"{k}={v}" for k, v in ADDR.items()))
