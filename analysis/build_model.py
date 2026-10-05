"""
Build TemanJourney_Model_2027_v15.xlsx from v13 (v13 is never changed).

Run from financial-estimations/:   python analysis/build_model.py
What v15 adds to v13, all as live formulas:
  1. Timing (Inputs section 11): "pipeline first" - no new B2B or B2C events before the first delivery month
     (Mar-27). Existing EAP and counseling clients continue.
  2. Growth plan (Inputs section 12): new EAP contracts, a sales co-founder and bank training days.
  3. Pay_Plan sheet: gross pay for every role, month by month (blue inputs), BPJS + THR, founder pay and extra IG
     ads, plus the running cost by group. This is where the step-by-step fair-pay path lives.
  4. Outlook sheet: sales, total cost, net profit and cash each month, with one combined chart.
v14 (the first version of this script) put fair pay in one step in January; v15 follows the founders' 5 October 2026
decisions. openpyxl drops cached values, so recalculate in Excel afterwards (analysis/excel_runner.py does this).
"""
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as CL

ROOT = Path(__file__).resolve().parent.parent
SRC, OUT = ROOT / "TemanJourney_Model_2027_v13.xlsx", ROOT / "TemanJourney_Model_2027_v15.xlsx"
wb = openpyxl.load_workbook(SRC)
I, V, C, P, CF, S, T = (wb[n] for n in ["Inputs", "Volume", "Capacity_OPEX", "PnL", "CashFlow", "Summary", "To_Fill"])
COLS = range(5, 20)                     # E..S = Oct-26 .. Dec-27


def style_like(dst, src):
    dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
    dst.alignment, dst.number_format = copy(src.alignment), src.number_format


def month(cell):                       # month index of an Inputs month label, 99 if the label is not found
    return f"IFERROR(MATCH(Inputs!${cell[0]}${cell[1:]},Volume!$E$4:$S$4,0),99)"


# ------------------------------------------------------------------ 1. Inputs sections 11 and 12
rows = [
    ("11. Timing - when new sales really start (v15)", None),
    ("pipe_on", "Pipeline-first timing? (1 = no new sales before the month below)", 1, "switch", "Your plan - confirm",
     "Oct 2026 to Feb 2027 is used to pitch and sign. Existing EAP and counseling clients continue."),
    ("sales_from", "First month new events are delivered", "Mar-27", "month", "Your plan - confirm",
     "Must match a month label (e.g. Mar-27). Earlier months get no new B2B events."),
    ("pipe_b2c", "Also stop B2C mass events before that month? (1 = yes)", 1, "switch", "Your plan - confirm",
     "1 = the team focuses on pitching. 0 = keep public events running."),
    ("12. Growth plan (v15) - pay is set month by month on the Pay_Plan sheet", None),
    ("gp_on", "Growth plan switch (1 = on, 0 = v13 cost base)", 1, "switch", "Your plan - confirm",
     "Turns on this section and the Pay_Plan sheet."),
    ("oncost", "BPJS + THR on top of staff pay", 0.18, "%", "Estimate",
     "Required for employees even when the micro-business wage rule applies; confirm with a payroll provider."),
    ("eap2_start", "New EAP contract 2 starts", "Feb-27", "month", "Your plan - confirm",
     "Each EAP contract: Rp15m a month, cost Rp5m. Signed during the pitching months."),
    ("eap3_start", "New EAP contract 3 starts", "Jul-27", "month", "Your plan - confirm", "Must match a month label."),
    ("cof_cost", "Sales co-founder cash pay", 8.0, "Rp m / month", "Estimate", "Plus shares that vest over 4 years."),
    ("cof_start", "Sales co-founder starts", "Jul-27", "month", "Your plan - confirm", "Start only when a bank, OJK or LPS offer is open."),
    ("bank_start", "Bank agreement: first training day", "Oct-27", "month", "Your plan - confirm",
     "Bank days are priced and costed as bank events (Rate_Card)."),
    ("bank_days", "Bank training days per month from that month", 1, "days", "Your plan - confirm", "Set 0 to test 'no bank deal'."),
    ("13. Operations (COO review, 5 Oct 2026)", None),
    ("pm_fee", "Project manager fee per B2B project", 2.0, "Rp m / project", "Known",
     "She is paid per project, not a salary: usually Rp2m per project (founders, 5 Oct 2026)."),
    ("pm_cap", "Projects one project manager can run in a month", 2, "projects", "Known",
     "Founders: she handles 1 or 2 projects a month. More projects = a second (standby) project manager."),
    ("cert_lead", "Certify new trainers this many months before they are needed", 2, "months", "Estimate",
     "Certification takes 1-2 months (observe, co-facilitate, lead). 0 = certify in the month of need (v13 logic)."),
    ("dp_flag", "Down payment on Flagship programs, paid the month before", 0.0, "%", "Your plan - confirm",
     "Taken from the same-month share. Test 30% in the scenario runs; ask BI whether a down payment or staged payment is possible."),
    ("adm_base", "Admin: base pay a month (paid for days she comes in)", 1.8, "Rp m / month", "Known", "Kia, 5 Oct 2026: Rp1.8m if she comes in."),
    ("adm_fee", "Admin: fee per client handled", 0.017, "Rp m / client", "Known", "Rp17,000 per client (counseling clients + EAP sessions)."),
    ("adm_bonus", "Admin: monthly bonus when she handles 5 clients a week", 0.62, "Rp m / month", "Known", "Rp155,000 a week x 4 weeks."),
    ("adm_thr", "Admin: clients a month that earn the bonus", 20, "clients", "Estimate", "5 clients a week x 4 weeks."),
    ("ump", "Jakarta minimum wage (fair-pay floor)", 5.73, "Rp m / month", "Known", "UMP DKI 2026."),
    ("perf1", "Pay step 1 unlocks when average sales of the last 3 months reach", 100, "Rp m / month", "Your plan - confirm",
     "Founders, 5 Oct 2026: targets first, then pay rises. Social media Rp3.0m, lead marketing Rp4.0m."),
    ("lp_cost", "Lead Partnership (non-psychologist) pay", 6.0, "Rp m / month", "Estimate",
     "Kia, 5 Oct 2026: no full-time psychologist for now; a non-psychologist Lead Partnership manages EAP clients and brings clients (no conflict of interest)."),
    ("lp_start", "Lead Partnership starts", "Oct-27", "month", "Your plan - confirm", "When 3 EAP contracts run."),
    ("perf2", "Pay step 2 unlocks at average sales of", 150, "Rp m / month", "Your plan - confirm",
     "Social media Rp4.5m; admin gets at least the minimum wage."),
    ("perf3", "Pay step 3 unlocks at average sales of", 200, "Rp m / month", "Your plan - confirm", "Social media at the minimum wage."),
    ("b2c_extra", "Extra public event (cash-savings step 5)", "Feb-27", "month", "Your plan - confirm",
     "One public event in the pitch months; the B2C funnel itself is on the B2C sheet."),
]
r0 = 100
first_input, sect_hdr = I["C98"], I["A84"]
ADDR = {}
for k, row in enumerate(rows):
    r = r0 + k
    if row[1] is None:
        I.cell(r, 1, row[0]); style_like(I.cell(r, 1), sect_hdr)
        continue
    key, item, val, unit, status, src = row
    for c, v in enumerate([key, item, val, unit, status, src], 1):
        cell = I.cell(r, c, v)
        style_like(cell, I.cell(86 if isinstance(val, str) else 97, c))
    if unit == "%":
        I.cell(r, 3).number_format = "0%"
    elif unit in ("switch", "days", "projects", "months", "clients"):
        I.cell(r, 3).number_format = "0"
    elif unit != "month":
        I.cell(r, 3).number_format = "0.000" if val < 0.1 else "0.00"
    ADDR[key] = f"C{r}"
A = lambda k: "Inputs!$C$" + ADDR[k][1:]

# ------------------------------------------------------------------ 2. Volume: timing flag, new EAP, bank days
hdr = V["A22"]
V["A33"] = "Timing and growth plan (v15)"; style_like(V["A33"], hdr)
lab = [(34, "Pipeline gap: no new sales yet (1 = yes)", "flag", "MAX"),
       (35, "New EAP contracts active (growth plan)", "contracts", "MAX"),
       (36, "Bank training days (growth plan)", "days", "SUM")]
for r, t, u, agg in lab:
    V.cell(r, 1, t); V.cell(r, 2, u)
    for c in (1, 2):
        style_like(V.cell(r, c), V.cell(29, c))
    V.cell(r, 3, f"={agg}(E{r}:G{r})"); V.cell(r, 4, f"={agg}(H{r}:S{r})")
    style_like(V.cell(r, 3), V["C29"]); style_like(V.cell(r, 4), V["D29"])
for c in COLS:
    L = CL(c)
    V[f"{L}34"] = f"=IF(AND({A('pipe_on')}=1,{L}6<IFERROR(MATCH({A('sales_from')},$E$4:$S$4,0),1)),1,0)"
    V[f"{L}35"] = (f"={A('gp_on')}*(IF({L}6>=IFERROR(MATCH({A('eap2_start')},$E$4:$S$4,0),99),1,0)"
                   f"+IF({L}6>=IFERROR(MATCH({A('eap3_start')},$E$4:$S$4,0),99),1,0))")
    V[f"{L}36"] = f"={A('gp_on')}*IF({L}6>=IFERROR(MATCH({A('bank_start')},$E$4:$S$4,0),99),{A('bank_days')},0)*(1-{L}34)"
    for r in (34, 35, 36):
        style_like(V[f"{L}{r}"], V[f"{L}29"])
    V[f"{L}23"] = V[f"{L}23"].value + f"*(1-{L}34)"
    V[f"{L}24"] = V[f"{L}24"].value + f"*(1-{L}34)"
    V[f"{L}25"] = V[f"{L}25"].value + f"*(1-{L}34)+{L}36"
    V[f"{L}26"] = f"={L}18+{L}35"
    V[f"{L}27"] = f"={L}19*(1-{L}34*{A('pipe_b2c')})"
V["A25"] = "Bank KCP events (incl. growth-plan bank days)"
V["A26"] = "Bank EAP contracts active (existing + growth plan)"
V["A31"] = ("Notes: v15 adds the timing switch (no new sales before the first delivery month) and the growth plan "
            "(new EAP contracts and bank days). Change them in Inputs sections 11 and 12.")

# ------------------------------------------------------------------ 3. Pay_Plan: pay for every role, month by month
PP = wb.create_sheet("Pay_Plan", 2)
PP["A1"] = "Pay plan: gross pay a month for every role (Rp million). Blue cells are your decisions; change them freely."
PP["A1"].font = Font(name="Arial", bold=True, size=13)
PP["A2"] = ("Decided 5 Oct 2026: BPJS + THR from Jan-27; social media up in steps; admin to the minimum wage from Jul-27; "
            "lead marketing capped at Rp4m in 2027 (KPI: 10k IG followers) with Rp3.5m of her budget moved to ads; "
            "admin paid per client (Rp1.8m base + Rp17k a client + Rp620k bonus); pay rises for admin, social media and lead "
            "marketing unlock when sales targets are met (rows 27-30), not by date; "
            "psychologist stays part-time (no full-time hire for now); a non-psychologist Lead Partnership from Oct-27; "
            "IG ads today Rp1m (software budget) + Rp3.5m moved from lead marketing; founders: only one is paid "
            "today (Rp10m); Rp25m from Mar-27 keeps room for the second founder, whose start is still to decide.")
PP["A2"].font = Font(name="Arial", italic=True, size=10)
for c in range(1, 20):
    style_like(PP.cell(4, c), V.cell(4, c))
    PP.cell(4, c).value = V.cell(4, c).value if c < 5 else f"=Volume!{CL(c)}4"
PP["A4"], PP["B4"], PP["C4"], PP["D4"] = "Line", "Unit", "Q4 2026", "FY 2027"
MON = ["Oct-26", "Nov-26", "Dec-26"] + [m + "-27" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()]


def path(steps):                       # steps: {first month: value}; value holds until the next step
    out, cur = [], None
    for m in MON:
        cur = steps.get(m, cur); out.append(cur)
    return out


PAY = [(6, "Admin (formula: base + fee per client + weekly bonus; minimum wage once step 2 is unlocked)", path({"Oct-26": 0.0})),
       (7, "Social media specialist (formula: 1.5, then 3.0 / 4.5 / minimum wage at pay steps 1 / 2 / 3)", path({"Oct-26": 0.0})),
       (8, "Lead marketing (formula: 2.5, then 4.0 cap at pay step 1; KPI 10k IG followers; has shares)", path({"Oct-26": 0.0})),
       (9, "Psychologists, paid part-time: PIC now + 1 more from Jan-27 (no full-time hire in 2027)", path({"Oct-26": 2.3, "Jan-27": 4.6})),
       (10, "Interns, not Magang (2 people)", path({"Oct-26": 2.0})),
       (11, "Founders: Rp10m now (one founder paid); Rp25m from Mar-27 = room for both (second founder's start: to decide)",
        path({"Oct-26": 10.0, "Mar-27": 25.0})),
       (12, "Extra IG ads (Rp3.5m moved from lead marketing; Rp1m is already in the software budget)", path({"Oct-26": 0.0, "Jan-27": 3.5})),
       (13, "Lead Partnership, non-psychologist (formula: from Inputs)", path({"Oct-26": 0.0})),
       (14, "BPJS + THR paid on staff pay? (1 = yes)", path({"Oct-26": 0, "Jan-27": 1}))]
PP["A5"] = "Gross pay a month (blue = your input)"; style_like(PP["A5"], V["A14"])
for r, lab_, vals in PAY:
    PP.cell(r, 1, lab_); style_like(PP.cell(r, 1), V.cell(15, 1))
    PP.cell(r, 2, "flag" if r == 14 else "Rp m"); style_like(PP.cell(r, 2), V.cell(15, 2))
    PP.cell(r, 3, f"=MAX(E{r}:G{r})" if r == 14 else f"=SUM(E{r}:G{r})")
    PP.cell(r, 4, f"=MAX(H{r}:S{r})" if r == 14 else f"=SUM(H{r}:S{r})")
    style_like(PP.cell(r, 3), V["C15"]); style_like(PP.cell(r, 4), V["D15"])
    for j, v in enumerate(vals):
        cell = PP.cell(r, 5 + j, v); style_like(cell, V.cell(15, 5))
        cell.number_format = "0" if r == 14 else "0.00"
PP["A27"] = "Pay steps: unlocked by sales, not by date (founders, 5 Oct 2026)"; style_like(PP["A27"], V["A14"])
PP["A28"] = "Average sales of the previous 3 months (Rp m)"
PP["A29"] = "Highest so far (a pay step never goes back down)"
PP["A30"] = "Pay step unlocked this month (0 to 3)"
for k, c in enumerate(COLS):
    L = CL(c)
    prev = [CL(c - d) for d in (1, 2, 3) if c - d >= 5]
    PP[f"{L}28"] = f"=AVERAGE({','.join('PnL!' + x + '12' for x in prev)})" if prev else "=0"
    PP[f"{L}29"] = f"=MAX({CL(c - 1)}29,{L}28)" if k else f"={L}28"
    PP[f"{L}30"] = f"=({L}29>={A('perf1')})+({L}29>={A('perf2')})+({L}29>={A('perf3')})"
    for r in (28, 29, 30):
        PP[f"{L}{r}"].font = Font(name="Arial", size=10); PP[f"{L}{r}"].number_format = "0" if r == 30 else "0.0"
    for r in (28, 29, 30):
        PP.cell(r, 1).font = Font(name="Arial", size=10)
for c in COLS:                          # admin, social media, lead marketing and Lead Partnership are formulas
    L = CL(c)
    cl = f"(Volume!{L}28+Volume!{L}26*Inputs!$C$68)"                # clients a month: counseling clients + EAP sessions
    dyn = f"{A('adm_base')}+{A('adm_fee')}*{cl}+IF({cl}>={A('adm_thr')},{A('adm_bonus')},0)"
    PP[f"{L}6"] = f"=IF({L}30>=2,MAX({A('ump')},{dyn}),{dyn})"
    PP[f"{L}7"] = f"=CHOOSE({L}30+1,1.5,3.0,4.5,{A('ump')})"
    PP[f"{L}8"] = f"=IF({L}30>=1,4.0,2.5)"
    PP[f"{L}13"] = f"=IF(Volume!{L}6>={month(ADDR['lp_start'])},{A('lp_cost')},0)"
    for r in (6, 7, 8, 13):
        PP[f"{L}{r}"].font = Font(name="Arial", size=10)
CALC = [(15, "Team cost incl. BPJS and THR (staff + psychologist + interns)",
         "=({L}6+{L}7+{L}8+{L}9+{L}13)*(1+{oc}*{L}14)+{L}10"),
        (16, "Total pay plan (team + founders + extra ads)", "={L}15+{L}11+{L}12")]
for r, lab_, f in CALC:
    PP.cell(r, 1, lab_); style_like(PP.cell(r, 1), V.cell(29, 1)); PP.cell(r, 2, "Rp m")
    PP.cell(r, 3, f"=SUM(E{r}:G{r})"); PP.cell(r, 4, f"=SUM(H{r}:S{r})")
    for c in COLS:
        PP.cell(r, c, f.format(L=CL(c), oc=A("oncost"))); style_like(PP.cell(r, c), V.cell(29, 5))
        PP.cell(r, c).number_format = "0.0"
PP["A18"] = "Running cost by group: all fixed costs a month (Rp million)"; style_like(PP["A18"], V["A14"])
GROUPS = [(19, "Founders", "=Capacity_OPEX!{L}26"),
          (20, "Team: staff, psychologist, interns, BPJS, THR", "=Capacity_OPEX!{L}22+Capacity_OPEX!{L}37"),
          (21, "Marketing: software and IG ads", "=Capacity_OPEX!{L}23+Capacity_OPEX!{L}38"),
          (22, "Office and accountant", "=Capacity_OPEX!{L}24+Capacity_OPEX!{L}25"),
          (23, "Sales and delivery support: co-founder, travel, project manager, 2nd admin, trainer training",
           "=Capacity_OPEX!{L}27+Capacity_OPEX!{L}28+Capacity_OPEX!{L}29+Capacity_OPEX!{L}30+Capacity_OPEX!{L}31+Capacity_OPEX!{L}39"),
          (24, "Total running cost (equals Capacity_OPEX total)", "=SUM({L}19:{L}23)"),
          (25, "Check: difference to Capacity_OPEX total (must be 0)", "=ROUND({L}24-Capacity_OPEX!{L}32,6)")]
for r, lab_, f in GROUPS:
    PP.cell(r, 1, lab_); style_like(PP.cell(r, 1), V.cell(29 if r == 24 else 23, 1)); PP.cell(r, 2, "Rp m")
    PP.cell(r, 3, f"=SUM(E{r}:G{r})"); PP.cell(r, 4, f"=SUM(H{r}:S{r})")
    for c in COLS:
        PP.cell(r, c, f.format(L=CL(c))); style_like(PP.cell(r, c), V.cell(29 if r == 24 else 23, 5))
        PP.cell(r, c).number_format = "0.0"
PP.column_dimensions["A"].width = 58
for c in range(2, 20):
    PP.column_dimensions[CL(c)].width = 8

# ------------------------------------------------------------------ 4. Capacity_OPEX: route the pay plan into fixed costs
C["A34"] = "Growth plan fixed costs (v15)"; style_like(C["A34"], C["A21"])
crow = [(35, "Growth plan on? (1 = yes)", "flag", "MAX"),
        (36, "BPJS + THR paid? (Pay_Plan)", "flag", "MAX"),
        (37, "Team pay change vs the 2027 budget (Pay_Plan, incl. BPJS and THR)", "Rp m", "SUM"),
        (38, "Extra IG ads (Pay_Plan) + referral rewards (B2C sheet)", "Rp m", "SUM"),
        (39, "Sales co-founder", "Rp m", "SUM"),
        (40, "Growth plan fixed cost", "Rp m", "SUM")]
for r, t, u, agg in crow:
    C.cell(r, 1, t); C.cell(r, 2, u)
    src = 32 if r == 40 else (17 if u == "flag" else 26)
    for c in range(1, 20):
        style_like(C.cell(r, c), C.cell(src, c))
    C.cell(r, 3, f"={agg}(E{r}:G{r})"); C.cell(r, 4, f"={agg}(H{r}:S{r})")
for c in COLS:
    L = CL(c)
    C[f"{L}35"] = f"={A('gp_on')}"
    C[f"{L}36"] = f"={L}35*Pay_Plan!{L}14"
    C[f"{L}37"] = f"={L}35*(Pay_Plan!{L}15-{L}22)"
    C[f"{L}38"] = f"={L}35*(Pay_Plan!{L}12+B2C!{L}53)"
    C[f"{L}39"] = f"={A('gp_on')}*IF(Volume!{L}6>={month(ADDR['cof_start'])},{A('cof_cost')},0)"
    C[f"{L}40"] = f"=SUM({L}37:{L}39)"
    C[f"{L}26"] = f"=IF({A('gp_on')}=1,Pay_Plan!{L}11,{L}17*(Inputs!$C$49+Inputs!$C$50))"
    C[f"{L}32"] = f"=SUM({L}22:{L}31)+{L}40"
C["A26"] = "Founder salaries (Pay_Plan when the growth plan is on)"

# ------------------------------------------------------------------ 4b. Operations (COO): PM per project, trainers certified early
C["A42"] = "Operations load (COO review)"; style_like(C["A42"], C["A21"])
orow = [(43, "B2B projects this month (one event = one project)", "projects", "SUM", "=Volume!{L}29"),
        (44, "Project managers needed", "people", "MAX", "=ROUNDUP({L}43/{pm_cap},0)"),
        (45, "Load on one project manager (100% = full)", "%", "MAX", "={L}43/{pm_cap}"),
        (46, "Trainer load (placements / what the trainer pool can do)", "%", "MAX",
         "=IF({L}9=0,0,{L}6/({L}9*Inputs!$C$34))"),
        (47, "Trainers to certify this month (ahead of need)", "people", "SUM",
         "=IF({gp}=1,IFERROR(INDEX($E$10:$S$10,Volume!{L}6+{lead}),0),{L}10)")]
for r, t, u, agg, f in orow:
    C.cell(r, 1, t); C.cell(r, 2, u)
    for c in range(1, 20):
        style_like(C.cell(r, c), C.cell(10, c))
    C.cell(r, 3, f"={agg}(E{r}:G{r})"); C.cell(r, 4, f"={agg}(H{r}:S{r})")
    for c in COLS:
        L = CL(c)
        C[f"{L}{r}"] = f.format(L=L, pm_cap=A("pm_cap"), gp=A("gp_on"), lead=A("cert_lead"))
        if u == "%":
            C[f"{L}{r}"].number_format = "0%"
    if u == "%":
        C.cell(r, 3).number_format = C.cell(r, 4).number_format = "0%"
for c in COLS:
    L = CL(c)
    C[f"{L}29"] = f"=IF({A('gp_on')}=1,{A('pm_fee')}*{L}43,{L}15*Inputs!$C$56)"
    C[f"{L}31"] = f"={L}47*Inputs!$C$37"
    C[f"{L}28"] = f"=IF({A('gp_on')}=1,Inputs!$C$55,{C[f'{L}28'].value[1:]})"   # we pitch from Oct-26, so travel starts now
C["A29"] = "L2 Project Manager (per project when the growth plan is on)"
C["A31"] = "L3 Facilitator certification (one-off, ahead of need)"
C["A28"] = "L1 Sales travel (from Oct-26 when the growth plan is on: we pitch now)"

# CashFlow: optional down payment on Flagship programs (paid the month before the event)
for j, c in enumerate(COLS):
    L, dp = CL(c), A("dp_flag")
    f = f"=PnL!{L}6*(Inputs!$C$73-{dp})+PnL!{CL(c + 1)}6*{dp}"
    if j >= 1:
        f += f"+PnL!{CL(c - 1)}6*Inputs!$C$74"
    if j >= 2:
        f += f"+PnL!{CL(c - 2)}6*Inputs!$C$75"
    CF[f"{L}7"] = f
CF["A7"] = "BI central - Flagship (incl. optional down payment)"

# ------------------------------------------------------------------ 5. PnL: route the new costs into the fixed lines
for c in COLS:
    L = CL(c)
    P[f"{L}27"] = P[f"{L}27"].value + f"+Capacity_OPEX!{L}37+Capacity_OPEX!{L}38"
    P[f"{L}29"] = P[f"{L}29"].value + f"+Capacity_OPEX!{L}39"
P["A27"] = "Base: team, software, office, accountant (+ pay plan and extra ads)"
P["A29"] = "B2B sales (co-founder, Account Manager, travel)"

# ------------------------------------------------------------------ 6. Outlook sheet + combined chart
O = wb.create_sheet("Outlook", 1)   # before Pay_Plan
O["A1"] = "Outlook: sales, costs and cash each month (Rp million)"
O["A1"].font = Font(name="Arial", bold=True, size=14)
O["A2"] = ("Read the bars against the line: when the grey cost bar is taller than the blue sales bar, cash goes down. "
           "Cash is paid later than sales, because clients pay 0 to 60 days after the event.")
O["A2"].font = Font(name="Arial", italic=True, size=10)
lines = [("Month", None), ("Sales (revenue)", "=PnL!{L}12"), ("Delivery cost", "=PnL!{L}21"), ("Fixed cost", "=PnL!{L}31"),
         ("Tax", "=PnL!{L}34"), ("Total cost", "={O}6+{O}7+{O}8"), ("Net profit", "=PnL!{L}35"),
         ("Cash in", "=CashFlow!{L}11"), ("Cash out", "=CashFlow!{L}17"), ("Cash at end of month", "=CashFlow!{L}21"),
         ("Cash floor (1 month of costs)", "=CashFlow!{L}24"), ("Above the floor? (1 = yes)", "=CashFlow!{L}25")]
for k, (lab_, f) in enumerate(lines):
    r = 4 + k
    O.cell(r, 1, lab_).font = Font(name="Arial", bold=(k in (0, 1, 5, 9)), size=10)
    for j, c in enumerate(COLS):
        L = CL(c)
        cell = O.cell(r, 2 + j, f"=Volume!{L}4" if f is None else f.format(L=L, O=CL(2 + j)))
        cell.font = Font(name="Arial", size=10, bold=(k == 0), color="008000")
        cell.number_format = "0" if k == 11 else "#,##0"
O.column_dimensions["A"].width = 30
for j in range(15):
    O.column_dimensions[CL(2 + j)].width = 8
bar = BarChart(); bar.type = "col"; bar.grouping = "clustered"
bar.add_data(Reference(O, min_col=1, max_col=16, min_row=5, max_row=5), from_rows=True, titles_from_data=True)
bar.add_data(Reference(O, min_col=1, max_col=16, min_row=9, max_row=9), from_rows=True, titles_from_data=True)
bar.set_categories(Reference(O, min_col=2, max_col=16, min_row=4, max_row=4))
bar.series[0].graphicalProperties.solidFill = "576C94"
bar.series[1].graphicalProperties.solidFill = "B8BEC9"
ln = LineChart()
ln.add_data(Reference(O, min_col=1, max_col=16, min_row=13, max_row=13), from_rows=True, titles_from_data=True)
ln.add_data(Reference(O, min_col=1, max_col=16, min_row=14, max_row=14), from_rows=True, titles_from_data=True)
ln.series[0].graphicalProperties.line.solidFill = "475571"; ln.series[0].graphicalProperties.line.width = 38100
ln.series[1].graphicalProperties.line.solidFill = "9B3809"; ln.series[1].graphicalProperties.line.dashStyle = "dash"
ln.y_axis.axId = bar.y_axis.axId
bar.title = "Sales and total cost (bars) vs cash in the bank (line), Rp million"
bar.y_axis.title = "Rp million"
bar.y_axis.majorGridlines = None
for sr in ln.series:
    sr.smooth = False
bar.x_axis.delete = False
bar.y_axis.delete = False
bar.x_axis.title = None
bar += ln
bar.height, bar.width = 11, 30
bar.legend.position = "t"
O.add_chart(bar, "A18")
O.page_setup.orientation = "landscape"
O.page_setup.fitToWidth, O.page_setup.fitToHeight = 1, 1
O.sheet_properties.pageSetUpPr.fitToPage = True

# ------------------------------------------------------------------ 6b. Roadmap: backward plan from each start date (COO)
from datetime import date
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, PatternFill

R = wb.create_sheet("Roadmap", 2)       # after Outlook
BLUE_IN, BOLD = Font(name="Arial", size=10, color="0000FF"), Font(name="Arial", size=10, bold=True)
NORM, HEAD = Font(name="Arial", size=10), PatternFill("solid", fgColor="DDE3EE")
YEL = PatternFill("solid", fgColor="FFF2CC")
DF = "d mmm yyyy"


def put(r, c, v, font=NORM, fmt=None, fill=None):
    cell = R.cell(r, c, v); cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def header(r, labels):
    for c, t in enumerate(labels, 1):
        cell = put(r, c, t, BOLD, fill=HEAD); cell.alignment = Alignment(wrap_text=True, vertical="top")


put(1, 1, "Roadmap: when we must start selling and preparing, worked backward from each start date",
    Font(name="Arial", bold=True, size=13))
put(2, 1, "Rule: start from the date the work begins and go back through three clocks: sales (lead to signed), delivery "
          "(people ready) and cash (work done to cash in). Regulators and banks also need us inside their budget, "
          "which they plan in the second half of the year before.", Font(name="Arial", italic=True, size=10))
put(4, 1, "Today (blue: change it when you review)", BOLD); put(4, 3, date(2026, 10, 5), BLUE_IN, DF)
TODAY = "$C$4"

put(6, 1, "A. Lead times by client type, in days (Estimate until we log 3 real deals of each type in section E)", BOLD)
header(7, ["Client type", "Contract and paperwork (legal, vendor registration, work order)", "Pitch, scoping, proposal",
           "Finding qualified leads"])
TYPES = [("Private company", 14, 45, 30), ("Regulator (BI, OJK, LPS)", 42, 75, 30),
         ("State-owned bank", 42, 75, 30), ("B2C public", 0, 30, 30)]
for k, row in enumerate(TYPES):
    for c, v in enumerate(row, 1):
        put(8 + k, c, v, BLUE_IN if c > 1 else NORM, fill=YEL if c > 1 else None)
LT = "$A$8:$D$11"

put(13, 1, "B. Sales clock: the latest date to start each deal", BOLD)
header(14, ["Milestone", "Client type", "New client? (1 = yes)", "Work starts", "Client budget window closes",
            "Sign by", "Start pitching", "Start finding leads", "Latest start", "Owner", "Status today"])
SALES = [("BI 2027: Flagship and regional programs (first in March)", "Regulator (BI, OJK, LPS)", 0, date(2027, 3, 1), date(2026, 12, 31)),
         ("BI 2027: regional events Jul-Sep (moved from the Q4 peak)", "Regulator (BI, OJK, LPS)", 0, date(2027, 7, 1), date(2026, 12, 31)),
         ("BI 2027: regional peak Oct-Dec", "Regulator (BI, OJK, LPS)", 0, date(2027, 10, 1), date(2026, 12, 31)),
         ("EAP contract 2: the warm lead (chase now)", "Private company", 0, date(2027, 2, 1), None),
         ("EAP contract 2 if the client is a bank", "State-owned bank", 1, date(2027, 2, 1), date(2026, 12, 31)),
         ("EAP contract 3 with a bank", "State-owned bank", 1, date(2027, 7, 1), date(2026, 12, 31)),
         ("Bank agreement: training days", "State-owned bank", 1, date(2027, 10, 1), date(2026, 12, 31)),
         ("B2C events (Instagram campaign, then tickets)", "B2C public", 1, date(2027, 10, 1), None),
         ("OJK and LPS programs for 2028", "Regulator (BI, OJK, LPS)", 1, date(2028, 3, 1), date(2027, 9, 30)),
         ("BI 2028 programs", "Regulator (BI, OJK, LPS)", 0, date(2028, 3, 1), date(2027, 12, 31))]
# Owners (proposal, 5 Oct 2026): Kia (CEO) leads regulators and banks; Galih leads EAP 2, B2C and people.
# The BD co-founder takes over OJK/LPS and banks after joining.
OWNER = {"Lead Partnership (non-psychologist)": "Kia", "EAP contract 2: the warm lead (chase now)": "Galih", "EAP contract 2 if the client is a bank": "Galih",
         "B2C events (Instagram campaign, then tickets)": "Galih (lead marketing)",
         "OJK and LPS programs for 2028": "Kia, then BD co-founder", "Bank agreement: training days": "Kia, then BD co-founder",
         "Standby project manager (paid per project)": "Kia", "Sales co-founder": "Kia and Galih"}
r = 15
for name, typ, new, start, close in SALES:
    put(r, 1, name); put(r, 2, typ, BLUE_IN); put(r, 3, new, BLUE_IN); put(r, 4, start, BLUE_IN, DF)
    put(r, 5, close, BLUE_IN, DF)
    put(r, 6, f"=D{r}-VLOOKUP(B{r},{LT},2,0)", fmt=DF)
    put(r, 7, f"=F{r}-VLOOKUP(B{r},{LT},3,0)", fmt=DF)
    put(r, 8, f"=G{r}-C{r}*VLOOKUP(B{r},{LT},4,0)", fmt=DF)
    put(r, 9, f'=IF(E{r}="",H{r},MIN(H{r},E{r}-VLOOKUP(B{r},{LT},3,0)))', fmt=DF)
    put(r, 10, OWNER.get(name, "Kia"), BLUE_IN)
    put(r, 11, f'=IF(G{r}<{TODAY},"Late",IF(I{r}<{TODAY},"At risk",IF(I{r}-{TODAY}<=31,"Start now","On track")))')
    r += 1
SALES_END = r - 1
put(r, 1, "Latest start = the earlier of: start finding leads, or the budget window close minus the pitch time.",
    Font(name="Arial", italic=True, size=9))

r += 2
put(r, 1, "C. Delivery clock: people ready before the work", BOLD); r += 1
header(r, ["Item", "Why", "", "Needed by", "Lead time (months)", "", "", "", "Latest start", "Owner", "Status today"]); r += 1
DELIV = [("3 certified trainers", "First programs in March need 3 trainers", date(2027, 3, 1), 3),
         ("Part-time psychologist pool for EAP 2", "EAP 2 starts in February", date(2027, 2, 1), 1),
         ("Second paid part-time psychologist", "B2C evenings and EAP days from January", date(2027, 1, 1), 2),
         ("Standby project manager (paid per project)", "From July, more than 2 projects in a month", date(2027, 7, 1), 2),
         ("Sales co-founder", "Joins in July; searching takes longer than a normal hire", date(2027, 7, 1), 6),
         ("8 certified trainers", "Busiest months: Oct to Dec 2027", date(2027, 10, 1), 3),
         ("Lead Partnership (non-psychologist)", "Manages EAP clients from October; brings clients", date(2027, 10, 1), 3),
         ("Intern for invoices", "Invoices grow to about 11 a month in Q4", date(2027, 9, 1), 1)]
DEL_START = r
for name, why, need, lead in DELIV:
    put(r, 1, name); put(r, 2, why); put(r, 4, need, BLUE_IN, DF); put(r, 5, lead, BLUE_IN)
    put(r, 9, f"=EDATE(D{r},-E{r})", fmt=DF); put(r, 10, OWNER.get(name, "Galih"), BLUE_IN)
    put(r, 11, f'=IF(I{r}<{TODAY},"Late",IF(I{r}-{TODAY}<=31,"Start now","On track"))')
    r += 1
DEL_END = r - 1

r += 1
put(r, 1, "D. Cash clock: work done, cash not yet in", BOLD); r += 1
cash = [("Lowest cash in the plan (Rp m)", "=MIN(Outlook!B13:P13)", "#,##0"),
        ("Month of the lowest cash", "=INDEX(Outlook!B4:P4,MATCH(MIN(Outlook!B13:P13),Outlook!B13:P13,0))", None),
        ("Months below the safety floor", "=COUNTIF(Outlook!B15:P15,0)", "0"),
        ("Down payment on Flagship programs (Inputs)", f"={A('dp_flag')}", "0%")]
for t, f, fmt in cash:
    put(r, 1, t); put(r, 4, f, fmt=fmt); r += 1
put(r, 1, "Costs leave in the event month; BI central pays 60/30/10% over 60 days, regional offices per payday, banks "
          "20/50/30%, EAP the next month. Fix a gap before signing: a down payment, staged invoices, or a later start.",
    Font(name="Arial", italic=True, size=9)); r += 2

put(r, 1, "E. Sales speed log: fill in after every won or lost deal, then replace the estimates in section A", BOLD); r += 1
header(r, ["Client", "Client type", "First contact", "Proposal sent", "Signed", "Work started", "Cash in",
           "Days to sign", "Days to cash", "Won or lost", "Note"]); r += 1
for k in range(8):
    for c in (3, 4, 5, 6, 7):
        put(r, c, None, BLUE_IN, DF)
    put(r, 8, f'=IF(OR(C{r}="",E{r}=""),"",E{r}-C{r})'); put(r, 9, f'=IF(OR(F{r}="",G{r}=""),"",G{r}-F{r})')
    r += 1

for rng in (f"K15:K{SALES_END}", f"K{DEL_START}:K{DEL_END}"):
    first = rng.split(":")[0]
    for word, col in (("Late", "F5DCD0"), ("At risk", "F5DCD0"), ("Start now", "F1DEBB"), ("On track", "DDE3EE")):
        R.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{word}"'], fill=PatternFill("solid", fgColor=col)))
widths = [52, 24, 10, 12, 14, 12, 12, 12, 12, 18, 12]
for c, w in enumerate(widths, 1):
    R.column_dimensions[CL(c)].width = w
R.row_dimensions[7].height = R.row_dimensions[14].height = 42
R.freeze_panes = "A5"

# ------------------------------------------------------------------ 6c. B2C funnel: Instagram followers to paid sessions
BC = wb.create_sheet("B2C", 3)
BLUE_F = Font(name="Arial", size=10, color="0000FF")
BC["A1"] = "B2C funnel: Instagram followers to paid counseling sessions, and B2C vs B2B revenue (Rp million)"
BC["A1"].font = Font(name="Arial", bold=True, size=13)
BC["A2"] = ("Growth is set per quarter (blue). Followers on 5 Oct 2026: @temanjourney_id. Today about 2 paid sessions a month, "
            "so the booking rate starts at about 0.6 new paying clients a month per 1,000 followers. Feeds Volume row 28.")
BC["A2"].font = Font(name="Arial", italic=True, size=10)
QH = ["Q4 2026", "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027"]
QIN = [(6, "Follower growth in the quarter", [0.15, 0.25, 0.28, 0.28, 0.24], "0%"),
       (7, "Booking rate: new paying clients a month per 1,000 followers", [0.58, 0.70, 0.85, 1.00, 1.00], "0.00"),
       (8, "Sessions per client", [1.00, 1.10, 1.25, 1.40, 1.40], "0.00"),
       (9, "Referral share of new clients (target 20% by Q4 2027)", [0.0, 0.05, 0.10, 0.15, 0.20], "0%")]
BC["A4"] = "Quarterly inputs (blue = change freely)"; style_like(BC["A4"], V["A14"])
for j, q in enumerate(QH):
    BC.cell(5, 3 + j, q).font = Font(name="Arial", size=10, bold=True)
for r, lab_, vals, fmt in QIN:
    BC.cell(r, 1, lab_).font = Font(name="Arial", size=10)
    for j, v in enumerate(vals):
        c = BC.cell(r, 3 + j, v); c.font = BLUE_F; c.number_format = fmt
SIN = [(10, "Instagram followers on 5 Oct 2026", 3439, "#,##0"), (11, "Theta voucher sessions a month per EAP contract", 2, "0"),
       (12, "Theta vouchers start", "Apr-27", "@"), (13, "Sessions one paid part-time psychologist covers a month (3 shifts a week x 3)", 36, "0"),
       (14, "Paid part-time psychologists from Jan-27", 2, "0"), (15, "Pool capacity: 20 part-timers x 8 hours a month", 160, "0"),
       (16, "Full-time psychologist pay, if hired (gross)", 9.0, "0.0"),
       (17, "Referral reward per referred client (Rp100k credit to each side)", 0.2, "0.00")]
for r, lab_, v, fmt in SIN:
    BC.cell(r, 1, lab_).font = Font(name="Arial", size=10)
    c = BC.cell(r, 3, v); c.font = BLUE_F; c.number_format = fmt
BC["E17"] = "Base Instagram ads already in the software budget (Rp m a month)"; BC["E17"].font = Font(name="Arial", size=10)
BC["K17"] = 1.0; BC["K17"].font = BLUE_F
for c in range(1, 20):
    style_like(BC.cell(18, c), V.cell(4, c))
    BC.cell(18, c).value = V.cell(4, c).value if c < 5 else f"=Volume!{CL(c)}4"
BC["A18"], BC["B18"], BC["C18"], BC["D18"] = "Line", "Unit", "Q4 2026", "FY 2027"
p_cc, f_cc = "Inputs!$C$14", "Inputs!$C$21"
MROWS = [(19, "Quarter (1 = Q4 2026 ... 5 = Q4 2027)", "#", None, "=INT((Volume!{L}6-1)/3)+1", "0"),
         (20, "Instagram followers (end of month)", "people", "MAX", None, "#,##0"),
         (21, "Booking rate", "per 1,000", "MAX", "=INDEX($C$7:$G$7,{L}19)", "0.00"),
         (22, "Sessions per client", "x", "MAX", "=INDEX($C$8:$G$8,{L}19)", "0.00"),
         (23, "New paying clients from Instagram", "clients", "SUM", "={L}20/1000*{L}21", "0.0"),
         (24, "Sessions from Instagram", "sessions", "SUM", "={L}23*{L}22", "0.0"),
         (25, "Sessions from Theta vouchers (EAP staff)", "sessions", "SUM",
          "=IF(Volume!{L}6>=IFERROR(MATCH($C$12,Volume!$E$4:$S$4,0),99),$C$11*Volume!{L}26,0)", "0.0"),
         (26, "Paid B2C counseling sessions (Instagram + vouchers + referral)", "sessions", "SUM", "={L}24+{L}25+{L}52", "0.0"),
         (28, "Counseling revenue", "Rp m", "SUM", "=PnL!{L}11", "0.0"),
         (29, "Public events", "events", "SUM", "=Volume!{L}27", "0"),
         (30, "Public event revenue", "Rp m", "SUM", "=PnL!{L}10", "0.0"),
         (31, "B2C revenue", "Rp m", "SUM", "={L}28+{L}30", "0.0"),
         (32, "B2B revenue (regulators, banks, EAP)", "Rp m", "SUM", "=PnL!{L}12-{L}31", "0.0"),
         (33, "B2C share of all revenue", "%", None, "=IF(PnL!{L}12=0,0,{L}31/PnL!{L}12)", "0%"),
         (35, "B2C gross profit (after psychologist session fees and event costs)", "Rp m", "SUM",
          "=PnL!{L}10+PnL!{L}11-PnL!{L}19-PnL!{L}20", "0.0"),
         (36, "Paid psychologists' pay incl. BPJS and THR", "Rp m", "SUM", "=Pay_Plan!{L}9*(1+" + A("oncost") + "*Pay_Plan!{L}14)", "0.0"),
         (37, "Extra Instagram ads", "Rp m", "SUM", "=Pay_Plan!{L}12", "0.0"),
         (38, "Psychologists' pay covered by B2C gross profit", "x", None, "=IF({L}36=0,0,{L}35/{L}36)", "0.0"),
         (39, "B2C profit after psychologists, ads and referral rewards", "Rp m", "SUM", "={L}35-{L}36-{L}37-{L}53", "0.0"),
         (41, "COO: sessions to deliver (B2C + EAP)", "sessions", "SUM", "={L}26+Volume!{L}26*Inputs!$C$68", "0"),
         (42, "Paid part-time capacity", "sessions", "SUM", "=IF(Volume!{L}6>=4,$C$14,1)*$C$13", "0"),
         (43, "Load on paid part-timers", "%", None, "={L}41/{L}42", "0%"),
         (44, "Sessions for the pool of 20", "sessions", "SUM", "=MAX(0,{L}41-{L}42)", "0"),
         (46, "Full-time test: counseling profit, 3-month average", "Rp m", None, None, "0.0"),
         (47, "Pay of the PIC part-timer + one full-timer, incl. BPJS and THR", "Rp m", None,
          "=(2.3+$C$16)*(1+" + A("oncost") + ")", "0.0"),
         (48, "Full-time psychologist affordable? (1 = yes)", "flag", "MAX", "=IF({L}46>={L}47,1,0)", "0"),
         (51, "New clients from referral", "clients", "SUM", "={L}23*INDEX($C$9:$G$9,{L}19)/(1-INDEX($C$9:$G$9,{L}19))", "0.0"),
         (52, "Sessions from referral", "sessions", "SUM", "={L}51*{L}22", "0.0"),
         (53, "Referral rewards (cost)", "Rp m", "SUM", "={L}51*$C$17", "0.00"),
         (54, "Referral share of new clients", "%", None, "=IF({L}23+{L}51=0,0,{L}51/({L}23+{L}51))", "0%"),
         (56, "Profit per client (LTV) = price x (1 - psychologist fee) x sessions per client", "Rp m", None, "=" + p_cc + "*(1-" + f_cc + ")*{L}22", "0.00"),
         (57, "Cost to win a client through Instagram (all ads / Instagram clients)", "Rp m", None,
          "=IF({L}23=0,0,($K$17+Pay_Plan!{L}12)/{L}23)", "0.00"),
         (58, "LTV / CAC, Instagram", "x", None, "=IF({L}57=0,0,{L}56/{L}57)", "0.0"),
         (59, "LTV / CAC, referral", "x", None, "={L}56/$C$17", "0.0")]
for r, lab_, u, agg, f, fmt in MROWS:
    BC.cell(r, 1, lab_); BC.cell(r, 2, u)
    for c in (1, 2):
        style_like(BC.cell(r, c), V.cell(23, c))
    if agg:
        BC.cell(r, 3, f"={agg}(E{r}:G{r})"); BC.cell(r, 4, f"={agg}(H{r}:S{r})")
    for c in (3, 4):
        style_like(BC.cell(r, c), V.cell(23, c)); BC.cell(r, c).number_format = fmt
    for k, c in enumerate(COLS):
        L = CL(c)
        if r == 20:
            g = "INDEX($C$6:$G$6,{L}19)".format(L=L)
            v = f"=$C$10*(1+{g})^(1/3)" if k == 0 else f"={CL(c - 1)}20*(1+{g})^(1/3)"
        elif r == 46:
            prev = [CL(c - d) for d in (0, 1, 2) if c - d >= 5]
            v = f"=AVERAGE({','.join(x + '26' for x in prev)})*{p_cc}*(1-{f_cc})"
        else:
            v = f.format(L=L)
        cell = BC.cell(r, c, v); style_like(cell, V.cell(23, c)); cell.number_format = fmt
BC["C33"] = "=IF(SUM(PnL!E12:G12)=0,0,C31/SUM(PnL!E12:G12))"; BC["D33"] = "=IF(SUM(PnL!H12:S12)=0,0,D31/SUM(PnL!H12:S12))"
BC["C38"] = "=IF(C36=0,0,C35/C36)"; BC["D38"] = "=IF(D36=0,0,D35/D36)"
for r in (33, 38):
    for c in (3, 4):
        BC.cell(r, c).number_format = "0%" if r == 33 else "0.0"
for r, t in ((27, "Revenue"), (34, "Can B2C pay for our psychologists?"), (40, "Capacity (COO)"), (45, "When can we afford a full-time psychologist?"),
             (50, "Referral (track the source of every booking)"), (55, "Unit economics per client")):
    BC.cell(r, 1, t); style_like(BC.cell(r, 1), V["A14"])
BC.column_dimensions["A"].width = 62
for c in range(2, 20):
    BC.column_dimensions[CL(c)].width = 9
BC.freeze_panes = "E19"
for c in COLS:                          # Volume: counseling sessions come from the funnel; one extra public event (step 5)
    L = CL(c)
    V[f"{L}28"] = f"=IF({A('gp_on')}=1,B2C!{L}26,{L}20)"
    V[f"{L}27"] = (V[f"{L}27"].value + f"+{A('gp_on')}*IF(Volume!{L}6=IFERROR(MATCH({A('b2c_extra')},$E$4:$S$4,0),99),1,0)")
V["A28"] = "B2C counseling sessions (B2C funnel sheet when the growth plan is on)"

# ------------------------------------------------------------------ 7. To_Fill: list the new estimates
last = T.max_row
new_fill = [("Timing", "First month new events are delivered", "Inputs!" + ADDR["sales_from"], "Decides how long we live on cash only."),
            ("Pay plan", "BPJS + THR on top of staff pay", "Inputs!" + ADDR["oncost"], "Cost of every pay step."),
            ("Pay plan", "Lead Partnership pay (Oct-27)", "Inputs!" + ADDR["lp_cost"], "Hiring step in Oct-27."),
            ("Growth plan", "EAP contract 2 and 3 start months", "Inputs!" + ADDR["eap2_start"], "Main source of new monthly income."),
            ("Growth plan", "Sales co-founder cash pay and start", "Inputs!" + ADDR["cof_cost"], "Hiring step in Jul-27."),
            ("Growth plan", "Bank training days per month and start", "Inputs!" + ADDR["bank_days"], "The swing factor for 2027."),
            ("Operations", "Project manager fee per project", "Inputs!" + ADDR["pm_fee"], "Cost of every B2B project."),
            ("Operations", "Lead times by client type (days)", "Roadmap!B8", "Sets the latest date to start selling."),
            ("Operations", "Down payment on Flagship programs", "Inputs!" + ADDR["dp_flag"], "Lifts the April cash low."),
            ("B2C", "Follower growth, booking rate and sessions per client by quarter", "B2C!C6", "Sets B2C revenue."),
            ("B2C", "Instagram followers and paid sessions, last 3-6 months", "B2C!C10", "Replaces the target path with a real trend.")]
for k, (pil, item, addr, why) in enumerate(new_fill):
    r = last + 1 + k
    vals = [int(T.cell(last, 1).value) + 1 + k, pil, item, f"={addr}", "", why, addr, None, "Open"]
    for c, v in enumerate(vals, 1):
        T.cell(r, c, v); style_like(T.cell(r, c), T.cell(last, c))

# ------------------------------------------------------------------ 8. Summary: scenario snapshot from the Excel runs
RUNS_JSON = ROOT / "output" / "model_runs.json"         # written by excel_runner.py; rebuild after a new run
if RUNS_JSON.exists():
    import json
    runs = json.loads(RUNS_JSON.read_text())
    cols = [("Old timing, v13 costs", "v13_timing_today_pay"), ("Real timing, v13 costs", "real_timing_today_pay"),
            ("Our plan (v15)", "plan"), ("Our plan, no bank deal", "plan_no_bank"), ("Our plan, BI buys 30% less", "plan_bi_cut")]
    S["A34"] = "Scenario comparison (v15 values from Excel; refresh with analysis/excel_runner.py, then rebuild)"
    for j, (lab_, k) in enumerate(cols):
        if k not in runs:
            continue
        f = runs[k]["fy27"]
        c = 2 + j
        S.cell(35, c, lab_)
        for r, v in [(36, f["R"]), (37, f["N"]), (38, f["m"]), (39, f["low"]), (40, MON[f["low_idx"]]),
                     (41, f["below_floor"]), (42, f["end"])]:
            S.cell(r, c, v)
    for c in range(1, 7):
        S.cell(43, c).value = None

wb.save(OUT)
print("wrote", OUT.name, "inputs:", ", ".join(f"{k}={v}" for k, v in ADDR.items()))
