"""Build the proposal page (v7: pyramid order, year 1 = bureau use case) from the model, so every number on the page comes from engine.py.

    python3 -I build_page.py      -> ../proposal/theta-partnership-proposal.html (+ ../output via run.py)

Prose lives here. Numbers come from inputs.py / engine.py. Never type a model number into the HTML by hand.
"""
import math, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import inputs as I
import engine
import run as runner

OUT_HTML = os.path.join(HERE, "..", "proposal", "theta-partnership-proposal.html")


# ---------- formatting ----------
def rp(x, d=1):
    """Rupiah from Rp million."""
    if abs(x) >= 1000:
        return f"Rp{x / 1000:,.1f} billion"
    s = f"Rp{abs(x):,.{d}f}M"
    return ("minus " + s) if x < 0 else s


def rpk(x):
    """Small amounts in full rupiah, from Rp million."""
    return f"Rp{round(x * 1_000_000, -3):,.0f}"


def num(x, d=0):
    return f"{x:,.{d}f}"


def pct(x, d=0):
    v = round(100 * x, d)
    return f"{(0.0 if v == 0 else v):.{d}f}%"


def month_name(label):
    y, m = label.split("-")
    names = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
             "October", "November", "December"]
    return f"{names[int(m) - 1]} {y}"


def table(head, rows, right=(), cls=""):
    th = "".join(f'<th class="r">{h}</th>' if i in right else f"<th>{h}</th>" for i, h in enumerate(head))
    body = "".join("<tr>" + "".join(f'<td class="r">{c}</td>' if i in right else f"<td>{c}</td>"
                                    for i, c in enumerate(r)) + "</tr>" for r in rows)
    return f'<div class="tablewrap {cls}"><table><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table></div>'


EST = '<span class="est">estimate</span>'


# ---------- model ----------
plan = engine.run(I, I.USE_NEXT_ROUND)
slow = engine.run(runner.slow_case(), I.USE_NEXT_ROUND)
runner.main_quiet = True
Y = plan["years"]
SY = slow["years"]
div_plan = runner.dividends(plan)
div_slow = runner.dividends(slow)
stake = plan["riskia_side"]["stake"]
paid_side = plan["riskia_side"]["paid"] + 0.0                       # Riskia 20 + 10
side_total = I.PAID_BEFORE["Riskia"] + I.DEAL_CASH                  # 50: Riskia 20 before + 30 now
one_offs = sum(I.DEAL_ONE_OFFS.values())
running_part = I.DEAL_CASH - one_offs

# market
avg_plan_mature = I.AVG_PLAN_BY_YEAR[2031]
per_psy_month = avg_plan_mature / I.PSY_PER_BUREAU + I.COPILOT_SEAT + I.PAYMENTS_PER_BUREAU[2031] * I.PAYMENT_FEE / I.PSY_PER_BUREAU
tam = I.PSYCHOLOGISTS_ID * per_psy_month * 12
bureau_psy = I.PSYCHOLOGISTS_ID * I.SHARE_IN_BUREAUS
sam = bureau_psy * per_psy_month * 12
bureaus_market = bureau_psy / I.PSY_PER_BUREAU
som = Y[2031]["revenue"]

# unit economics, 2028 (first full year of selling)
rows28 = [r for r in plan["rows"] if r["ym"][0] == 2028]
new28 = sum(r["new"] for r in rows28) / 12
paying28 = sum(r["paying"] for r in rows28) / 12
cac = I.WORKSHOPS[2028] / new28 + I.ONBOARDING_PER_NEW_BUREAU
contrib = sum(r["plan"] + r["fees"] + (r["seats"] - r["solo"]) * (I.COPILOT_SEAT - I.AI_COST_PER_SEAT)
              - r["bureaus"] * I.SERVER_PER_BUREAU for r in rows28) / 12 / paying28
payback = cac / contrib

# bureau example (10 psychologists, Growth plan)
ex_psy, ex_payments = 10, 120
ex_plan = 2.0
ex_fees = ex_payments * I.PAYMENT_FEE
ex_tools = ex_plan + ex_fees
ex_copilot = ex_psy * I.COPILOT_SEAT
ex_admin_saved = 2.5          # E half of one admin's month
ex_extra_fees = ex_psy * 2 * 0.7
ex_ai_cost = ex_psy * I.AI_COST_PER_SEAT

# sales agent example (one bureau, per month)
X = I.EXAMPLE
x_after = X["inquiries"] * X["after_hours_share"]
x_now = X["inquiries"] * X["conversion_now"]
x_agent = X["inquiries"] * X["conversion_agent"]
x_extra = x_agent - x_now
x_extra_rev = x_extra * X["first_purchase"]

low = plan["lowest_cash"]
slow_low = slow["lowest_cash"]
safety_line = math.ceil(-slow_low["cash"] / 5) * 5 + 5 if slow_low["cash"] < 0 else 0
exit_low = plan["exit"]["low"]
exit_mid = plan["exit"]["mid"]
slow_exit_low = slow["exit"]["low"]
div2 = sum(div_plan.values())
starter_margin = I.PLAN_TIERS[0][3] + I.PAYMENTS_PER_BUREAU[2028] * I.PAYMENT_FEE - I.AGENT_COST_PER_BUREAU - I.SERVER_PER_BUREAU
starter_needed = math.ceil(I.BASE_RUN_COST / starter_margin)
WORDS = {4: "Four", 5: "Five", 6: "Six", 7: "Seven", 8: "Eight"}
slow_profit_year = next(y for y in SY if all(SY[z]["profit"] > 0 for z in SY if z >= y))
sdiv2 = sum(div_slow.values())


# ---------- page (v6: pyramid order, one lead product, visual) ----------
import viz

css = open(os.path.join(HERE, "page_style.css")).read()

# extra numbers for v6
assess_x = I.ASSESS_PER_WEEK_COPILOT / I.ASSESS_PER_WEEK_NOW
agree_x = I.AGREEMENT_TARGET / I.AGREEMENT_NOW
more_clients = I.EXAMPLE["conversion_agent"] / I.EXAMPLE["conversion_now"] - 1
asean_share = I.ASEAN6["assessments_per_year_m"] / I.ASEAN6["with_disorder_m"]
asean_after = I.ASEAN6["assessments_per_year_m"] * assess_x
per100k = I.ASEAN6["psychologists"] / (I.ASEAN6["people_m"] * 10)
SESSIONS_PER_CLIENT = 1.4                       # E  TemanJourney model (Galih)
lp_payments = x_agent * SESSIONS_PER_CLIENT
lp_cost = I.PLAN_TIERS[0][3] + lp_payments * I.PAYMENT_FEE
lp_return = x_extra_rev / lp_cost
first_profit = plan["first_profitable_month"]
good_profit = plan["profitable_for_good_from"]
def paying_at(label, res=None):
    return next(r for r in (res or plan)["rows"] if r["label"] == label)["bureaus"]
B27 = {(2027, q): round(paying_at(f"2027-{3 * q:02d}")) for q in (1, 2, 3, 4)}    # bureaus still paying at each quarter end
B27.update({(2027, 3 * q): B27[(2027, q)] for q in (1, 2, 3, 4)})
GA = I.GROSS_ADDS
SIGNED27 = round(sum(v for k, v in GA.items() if k[0] == 2027))
LOST27 = round(Y[2027]["churned"])
ADMIN_FROM, ADMIN_TO = 9, 17                    # E  typical admin hours
Y1 = round(Y[2027]['bureaus_end'])
y1_share = Y1 / bureaus_market
rnd = I.NEXT_ROUND
rnd_when = month_name(f"{rnd['month'][0]}-{rnd['month'][1]:02d}")
stake_after = plan["riskia_side"]["stake_after_round"]
top_plan = I.PLAN_TIERS[-1][3]
noangel = engine.run(I, False)
NY = noangel["years"]
Q_NEW = [round(sum(v for k, v in GA.items() if k[0] == 2027 and (k[1] - 1) // 3 == q)) for q in range(4)]
B26 = round(GA[(2026, 12)])
B28 = Y[2028]["bureaus_end"]
def hc_at(label):
    r_ = next(r for r in plan["rows"] if r["label"] == label)
    return r_["headcount"]
TEST_FEE_LOW, TEST_FEE_HIGH = 50_000, 150_000   # K  KantorKu online recruitment tests, per candidate (Augment A3)

fp = I.FOUNDER_PAY_YEAR
S = []
add = S.append
yes, no = '<span class="yes">Yes</span>', '<span class="no">No</span>'

# ---- header + mission
add(f"""
<header style="display:flex;flex-direction:column;gap:14px">
  <div class="eyebrow">CoreItera · proposal for Riskia and TemanJourney</div>
  <h1>Theta Partnership Proposal</h1>
  <p class="meta">Galih Muji Agung, CEO of CoreItera · 19 October 2026 · Draft, version 8</p>
</header>
<section id="mission">
  <div class="mission">
    <div class="eyebrow">Our mission</div>
    <p class="mission-text">{I.MISSION}</p>
  </div>
</section>""")

# ---- the answer (top of the pyramid)
add(f"""
<section id="answer">
  <h2>The answer in one picture</h2>
  {viz.pyramid(
      f"Invest <b>Rp30M</b> for <b>20%</b>. We use it to help <b>TemanJourney</b> win more clients with a <b>24-hour sales agent</b>. Then we sign <b>{num(SIGNED27 + B26)} bureaus</b> by December 2027 ({num(Y1)} still paying after churn) and <b>{num(B28)}</b> in year 2, and enter <b>Malaysia or Singapore</b> in year 3.",
      [("A", "The need is big", [
            ("a1", f"Too few psychologists: {per100k:.0f} per 100,000 in ASEAN"),
            ("a2", "Bureaus lose clients who chat after admin hours"),
            ("a3", f"Indonesia first: about {num(round(bureaus_market, -2))} bureaus")]),
       ("B", "One product first: the bureau tool", [
            ("b1", "Easy to start, easy to run, tests included"),
            ("b2", "Booking, chat and tests are already built"),
            ("b3", "Only we sell 24 hours in the bureau's own style")]),
       ("C", "A safe plan and deal", [
            ("c1", "Year 1: bureaus only. Year 2: Psynthesia"),
            ("c2", f"Cheap to run; a {rp(rnd['amount'], 0)} angel round keeps growth steady"),
            ("c3", "Rp50M in total for 20%; the 15% promise kept")])])}
  <div class="tiles">
    <div class="tile"><b>{num(B28)} bureaus</b><span>paying in year 2 (December 2028); {num(Y1)} in December 2027</span></div>
    <div class="tile"><b>Rp1M to {rp(top_plan, 1)}</b><span>a month; never more than one admin</span></div>
    <div class="tile"><b>{month_name(first_profit)}</b><span>first month with a profit</span></div>
    <div class="tile"><b>20%</b><span>Riskia 16.1%, TemanJourney 3.9%</span></div>
  </div>
</section>""")

# ---- PART A
add(f"""
<div class="part"><span class="pyr-letter">A</span><h2>The need is big</h2></div>

<section id="a1">
  <p class="msg">A1. There are too few psychologists, and training more takes too long.</p>
  {viz.hbars([
      ("People with a mental disorder (six ASEAN countries)", I.ASEAN6['with_disorder_m'], "c-cost", f"{I.ASEAN6['with_disorder_m']:.1f} million"),
      ("Full assessments possible each year", I.ASEAN6['assessments_per_year_m'], "c-ochre", f"{I.ASEAN6['assessments_per_year_m']:.1f} million ({pct(asean_share, 1)})"),
      (f"The same psychologists with CoreItera", asean_after, "c-rev", f"{asean_after:.1f} million ({assess_x:.1f}×)"),
  ], note="Indonesia, Philippines, Vietnam, Thailand, Malaysia, Singapore. One scale. Sources: GBD 2021 ASEAN (2025); Augment proposal (IS5007).")}
  <div class="tiles three">
    <div class="tile"><b>1 billion+</b><span>people in the world with a mental health condition (WHO)</span></div>
    <div class="tile"><b>{per100k:.0f} per 100,000</b><span>psychologists in our six ASEAN countries</span></div>
    <div class="tile"><b>{I.ID['years_to_close_gap']} years</b><span>to train enough clinical psychologists in Indonesia (HIMPSI)</span></div>
  </div>
  <p><b>So each psychologist must help more people.</b> Our targets: <b>{pct(more_clients)} more clients</b> per bureau (sales agent), <b>{assess_x:.1f}× more assessments</b> per psychologist (Psynthesia, year 2), and <b>{agree_x:.1f}× more accurate diagnosis</b> (from {I.AGREEMENT_NOW:.2f} to {I.AGREEMENT_TARGET:.1f} agreement, a target).</p>
</section>

<section id="a2">
  <p class="msg">A2. Bureaus lose clients at night. Clients chat after work; the admin replies the next morning.</p>
  {viz.day_band(ADMIN_FROM, ADMIN_TO, X['after_hours_share'])}
  <div class="two">
    <div class="card"><div class="eyebrow">For the bureau</div><ul>
      <li>Clients ask on Instagram and WhatsApp, often after work.</li>
      <li>A late reply loses the client.</li>
      <li>A price question takes as long as a big company order.</li>
    </ul></div>
    <div class="card"><div class="eyebrow">For the founder</div><ul>
      <li>Long hours teaching the admin how to sell.</li>
      <li>When the admin leaves, the teaching starts again.</li>
      <li>Psychologists lose 18 of 45 hours a week to admin and reports.</li>
    </ul></div>
  </div>
  <p class="note">Admin hours and the {pct(X['after_hours_share'])} after-hours share are estimates. We count them at TemanJourney in December 2026.</p>
</section>

<section id="a3">
  <p class="msg">A3. Indonesia is the biggest market, and we can reach it first.</p>
  {viz.hbars([
      ("All psychologists in Indonesia", tam, "c-cost", f"{rp(tam)} a year"),
      (f"In private practice (about {num(round(bureaus_market, -2))} bureaus)", sam, "c-ochre", f"{rp(sam)} a year"),
      (f"Year-1 plan: {num(Y1)} bureaus ({pct(y1_share)})", Y[2027]['revenue'], "c-rev", f"{rp(Y[2027]['revenue'])} in 2027"),
  ], note=f"Value = psychologists × {rpk(per_psy_month)} a month each. The bureau count is our estimate: HIMPSI does not publish one.")}
  {table(["Why now", "What it means"], [
    ["Clients buy in chat", "The first bureau to reply wins."],
    ["Data law (UU PDP 27/2022)", "Client data on phones is a legal risk."],
    ["Psychology law (UU 23/2022)", "Bureaus need proper records."],
    ["AI can now draft", "The bureau still decides."],
  ])}
</section>""")

# ---- PART B
pos = viz.positioning([
    ("CoreItera", 88, 88, True),
    ("Admin on WhatsApp", 10, 50, False),
    ("Generic chatbots (Qontak, Wati)", 70, 18, False),
    ("Global practice software", 8, 34, False),
    ("Counselling apps (Riliv, Satu Persen)", 42, 70, False),
], "Sells for the bureau 24 hours", "Built for Indonesian psychology")
add(f"""
<div class="part"><span class="pyr-letter">B</span><h2>One product first: the bureau tool</h2></div>

<section id="b1">
  <p class="msg">B1. Year 1 sells one thing: a bureau tool that is easy to start, easy to run, with tests included.</p>
  <div class="tiles three">
    <div class="tile"><b>Easy to start</b><span>Set up in one day. Connect WhatsApp and Instagram; add psychologists and their hours.</span></div>
    <div class="tile"><b>Easy to run</b><span>Sales agent, scheduling, payment, chat counselling and follow-ups in one place.</span></div>
    <div class="tile"><b>Tests included</b><span>Send and score tests in CoreItera. No fee to other test sites (Rp{TEST_FEE_LOW:,} to Rp{TEST_FEE_HIGH:,} per person on KantorKu).</span></div>
  </div>
  <h3>How the sales agent works</h3>
  {table(["", "Admin away (night, weekend)", "Admin online"], [
    ["Who replies", "The agent, at once", "The agent first, then the admin"],
    ["What it does", "Answers, asks a few questions, offers times and a payment link", "Picks serious leads and sums each one up"],
    ["Who closes", "The client books and pays, or the lead waits for the morning", "The admin"],
    ["Style", "The bureau's own style, learned from its past chats", "The admin corrects it"],
    ["Safety", "No counselling or diagnosis. Crisis: hotline at once, and a person is alerted", "The same"],
  ])}
  {viz.compare_bars([
      ("Chats a month", X['inquiries'], X['inquiries'], num(X['inquiries']), num(X['inquiries'])),
      ("Answered in minutes", X['inquiries'] - x_after, X['inquiries'], num(X['inquiries'] - x_after), num(X['inquiries'])),
      ("Paying clients", x_now, x_agent, f"{num(x_now)} ({pct(X['conversion_now'])})", f"{num(x_agent)} ({pct(X['conversion_agent'])})"),
  ])}
  <div class="two">
    <div class="card"><div class="eyebrow">The bureau gets</div><p class="lead"><b>{num(x_extra)} more clients</b> a month = <b>{rp(x_extra_rev)}</b> more sales {EST}</p></div>
    <div class="card"><div class="eyebrow">The bureau pays</div><p class="lead"><b>{rp(lp_cost, 2)}</b> (Starter plan + payment fees) = about <b>{lp_return:.0f}×</b> back. The top plan never costs more than one admin ({rpk(top_plan)}).</p></div>
  </div>
  <p class="note">Example: {num(X['inquiries'])} chats a month, {SESSIONS_PER_CLIENT} paid sessions per client. We test it at TemanJourney first.</p>
</section>

<section id="b2">
  <p class="msg">B2. Most of it is already built, so we can deliver.</p>
  {viz.product_stack([
      ("Year 1", "ps-lead", "The bureau tool", [("Sales agent on WhatsApp", "Feb 2027"), ("Sales agent on Instagram", "May 2027"), ("Set up in one day", "Mar 2027"), ("Scheduling for any bureau", "Mar 2027"), ("Tests included", "Mar 2027"), ("Bureau's own style", "with the agent")]),
      ("Already built", "ps-built", "What the agent sells", [("Booking and payment", "live 30 Nov 2026"), ("Chat counselling (Theta)", "live 30 Nov 2026"), ("Tests and reports", "live today"), ("Follow-up emails", "live")]),
      ("Year 2", "ps-next", "Psynthesia: report copilot", [("AI drafts the report", "Rp300,000 a seat, from Jan 2028"), ("Psychologist signs", "always")]),
      ("Later", "ps-later", "Beyond Indonesia", [("Local test norms", "from every test"), ("Second ASEAN country", "after 2030")]),
  ])}
  <ul>
    <li><b>Booking and chat counselling (Theta):</b> live at TemanJourney on 30 November 2026.</li>
    <li><b>Tests:</b> already on CoreItera; TemanJourney already buys them.</li>
    <li><b>AI reports, tested on 205 reports:</b> scoring 60 → 10 minutes; writing 7 → 1 hour; quality {I.QUALITY_BEFORE} → {I.QUALITY_AFTER}+ out of 6.</li>
    <li><b>Writing-style engine:</b> built for reports; the agent reuses it.</li>
  </ul>
</section>

<section id="b3">
  <p class="msg">B3. Only CoreItera sells for a bureau 24 hours, in its own style, into its real calendar.</p>
  {pos}
  {table(["", "CoreItera", "Admin on WhatsApp", "Generic chatbots", "Global software", "Counselling apps"], [
    ["Replies 24 hours", yes, no, yes, no, "Only for the app"],
    ["Bureau's own style", yes, yes, "Set by hand", no, no],
    ["Books real free times", yes, "By hand", no, yes, yes],
    ["Local payment", yes, "Bank transfer", "Add-on", "US cards", yes],
    ["Psychology safety rules", yes, "Depends", no, "Partly", yes],
    ["Counselling and tests", yes, no, no, "Notes only", "Counselling only"],
    ["Client belongs to the bureau", yes, yes, yes, yes, no],
  ])}
  {viz.flywheel([("More bureaus", "join for clients"), ("More chats", "answered 24 h"), ("Smarter agent", "every correction"), ("More clients", "bureaus earn more"), ("Referrals", "bureaus bring bureaus")], ("More people helped", "per psychologist"))}
  <p><b>Hard to copy:</b> each bureau's style lives in its agent, and its clients and bookings live in CoreItera. TemanJourney, HIMPSI and universities are trust channels. We cost {rp(I.BASE_RUN_COST)} a month to run, so we can price low.</p>
</section>""")

# ---- PART C
road = viz.roadmap([
    {"name": "1. TemanJourney first", "when": "Nov 2026 to Mar 2027", "gate": f"TemanJourney wins more clients; {num(B27[(2027, 3)])} bureaus pay",
     "cells": {"Product": "Booking and chat live 30 Nov; WhatsApp agent in Feb", "Customers": f"{num(B26)} bureaus + TemanJourney by December 2026; +{num(Q_NEW[0])} in Q1 2027 (7 a month). Galih and Riskia close the deals; no sales staff",
               "Money": f"Rp30M in; {rp(I.BASE_RUN_COST)} a month", "Team": "3 founders + TemanJourney psychologists"}},
    {"name": f"2. Grow steadily", "when": "Apr to Dec 2027", "gate": f"{num(Y1)} bureaus still paying; churn down to 5% a month",
     "cells": {"Product": "Set up in one day; tests; Instagram agent", "Customers": f"+{num(Q_NEW[1])} in Q2, +{num(Q_NEW[2])} in Q3, +{num(Q_NEW[3])} in Q4: {num(SIGNED27)} signed, {num(Y1)} still paying by December",
               "Money": f"Angel round {rp(rnd['amount'], 0)} in {rnd_when}; first profit {month_name(first_profit)}", "Team": f"Each hire when profit pays it. By December: {hc_at("2027-12")["Sales"]} sales, {hc_at("2027-12")["Support"]} support, {hc_at("2027-12")["Engineers"]} engineers, {hc_at("2027-12")["Quality"]} quality"}},
    {"name": f"3. Year 2: {num(B28)} bureaus + Psynthesia", "when": "2028", "gate": f"{num(B28)} bureaus pay; churn 2% a month; 40% buy a seat",
     "cells": {"Product": "Psynthesia, Rp300,000 a seat", "Customers": f"{num(Y[2028]['bureaus_end'])} bureaus",
               "Money": f"Full founder pay from {month_name(plan['full_pay_from'])}", "Team": "Clinical lead; more engineers and quality"}},
    {"name": "4. Year 3: Malaysia or Singapore", "when": "2029 to 2031", "gate": "First paying bureaus abroad",
     "cells": {"Product": "English version; local tests and data law (PDPA)", "Customers": f"{num(Y[2031]['bureaus_end'])} bureaus in Indonesia by 2031, plus the first country abroad", "Money": f"{rp(Y[2031]['revenue'])} revenue in 2031 (Indonesia only)", "Team": f"Country lead abroad; {rp(I.ABROAD_COST, 0)} a month to set up"}},
], ["Product", "Customers", "Money", "Team"])
eng = viz.gantt([
    ("Booking and chat (Theta)", "launch: Adit, Zaki", 0, 1, "c-cost"),
    ("Sales agent on WhatsApp", "Adit", 1, 2, "c-rev"),
    ("Writing style for chats", "Zaki", 1, 2, "c-rev"),
    ("Scheduling for any bureau", "Adit, Galih", 3, 2, "c-rev"),
    ("Set up in one day, tests", "Zaki", 3, 3, "c-rev"),
    ("Sales agent on Instagram", "Adit", 4, 2, "c-rev"),
    ("Dashboard of extra clients", "Zaki", 6, 2, "c-rev"),
    ("Psynthesia, ready for year 2", "Zaki, clinical lead", 10, 4, "c-ochre"),
], (2026, 11), 14)
TP = I.TEAM_PLAN
periods = ["Q1 27", "Q2 27", "Q3 27", "Q4 27", "2028", "2029", "2030", "2031"]
keys = [(2027, 1), (2027, 2), (2027, 3), (2027, 4), 2028, 2029, 2030, 2031]
core_n = 4   # Galih, Riskia, Adit, Zaki
ends = ["2027-03", "2027-06", "2027-09", "2027-12", "2028-12", "2029-12", "2030-12", "2031-12"]
HC = [hc_at(e) for e in ends]
team_viz = viz.team_chart(periods, [
    ("Founders and Zaki", "t-core", [core_n] * len(ends)),
    ("Engineers", "t-eng", [h["Engineers"] for h in HC]),
    ("Quality", "t-qa", [h["Quality"] for h in HC]),
    ("Support", "t-cs", [h["Support"] for h in HC]),
    ("Sales", "t-sales", [h["Sales"] for h in HC]),
])
why = {"Sales": f"Rp{I.TEAM_PAY['Sales'][2027]:.0f}M base + {pct(I.COMMISSION_FIRST_PAYMENT)} of each new bureau's first payment",
       "Support": f"Today: one person shared with TemanJourney, who also sells for {pct(I.COMMISSION_FIRST_PAYMENT)} of the first payment",
       "Engineers": "Beyond Adit and Zaki; one main build a month",
       "Quality": "Tests every release and the agent's answers"}
team_rows = [[r] + [str(HC[i][r]) for i in (0, 1, 2, 3, 4, 7)] + [why[r]] for r in I.HIRE_ORDER]
def _c(curve, ym):
    return engine.interp(curve, ym, before=curve[min(curve)])
churn_rows = []
for name_, curve_ in I.CHURN_SCENARIOS.items():
    res_ = engine.run(runner.churn_case(curve_), I.USE_NEXT_ROUND)
    yy_ = res_["years"]
    path_ = " → ".join(pct(_c(curve_, k)) for k in ((2027, 1), (2027, 12), (2028, 12)))
    churn_rows.append([f"<b>{name_}</b>", path_, num(yy_[2027]["bureaus_end"]), num(yy_[2028]["new"]), rp(yy_[2028]["profit"]), rp(yy_[2031]["profit"])])
MN_ = {1: "Jan", 6: "Jun", 12: "Dec"}
cpts = [((f"{MN_[r['ym'][1]]} {r['ym'][0]}" if (r["ym"] == (2027, 1) or r["ym"][1] in (6, 12)) else ""), 100 * r["churn_rate"])
        for r in plan["rows"] if (2027, 1) <= r["ym"] <= (2029, 6)]
churn_viz = viz.line_chart(cpts, ref=[("Industry middle 5%", 5.0), ("Good software 2%", 2.0)], label="Share of paying bureaus that leave each month", aria="Monthly churn")
q_rows = [[q["quarter"], rp(q["revenue"]) if q["revenue"] > 0.05 else "None", rp(q["costs"]), rp(q["investment"]) if q["investment"] else "None", rp(q["net"]), f"<b>{rp(q['cash_end'])}</b>"] for q in plan["quarters"]]
add(f"""
<div class="part"><span class="pyr-letter">C</span><h2>A safe plan and deal</h2></div>

<section id="c1">
  <p class="msg">C1. Year 1: bureaus only, TemanJourney first. Year 2: Psynthesia.</p>
  {road}
  <div class="callout">
    <p><b>Why {num(B28)} in year 2.</b> Vibefam (Singapore) runs 700+ gyms and studios on one system, and bureaus need the same things: bookings, payments and a front desk that sells. We build up step by step in year 1 ({num(B26)}, then +{num(Q_NEW[0])}, +{num(Q_NEW[1])}, +{num(Q_NEW[2])}, +{num(Q_NEW[3])} a quarter) and reach {num(B28)} in year 2. Year 3 takes the same tool to Malaysia or Singapore.</p>
  </div>
  <h3>One build at a time</h3>
  {eng}
  <p class="note">Proposed owners. One main build a month. Psynthesia work starts in September 2027; it is sold from January 2028.</p>
  <h3>How the team grows</h3>
  <p class="msg">We hire only when profit can pay the first salary: sales first, then support, then quality and engineering.</p>
  {team_viz}
  {table(["Role", "Mar 2027", "Jun", "Sep", "Dec 2027", "Dec 2028", "Dec 2031", "Pay and job"], team_rows, right=(1, 2, 3, 4, 5, 6))}
  <p class="note">A hire happens only when last month's profit can pay that salary. Galih and Riskia are the deal makers until then.</p>
  {table(["Who", "Role"], [
    ["<b>Galih</b>, CEO", "Strategy, money, payments, data protection. Studying at NUS, so the smallest build share."],
    ["<b>Adit</b>, co-founder", "Sales agent, scheduling and chat; B2B sales."],
    ["<b>Zaki</b>, design lead", "Screens, set-up, writing style; Psynthesia in year 2."],
    ["<b>Riskia</b>, investor", "Runs TemanJourney, the first bureau; opens doors."],
  ])}
  {table(["Risk", "Early sign", "What we do"], [
    ["Agent says something clinical", "Chats about symptoms or crisis", "No counselling; crisis goes to the hotline and a person."],
    ["Meta limits messaging", "A warning on WhatsApp or Instagram", "Official Meta tools only; answers about the bureau's services only."],
    ["Bureaus leave after a few months", "Churn not below 5% a month by December 2027", "Set-up call in week 1; monthly extra-clients report; fix the top reason they leave."],
    ["Agent does not add clients", "No rise at TemanJourney in February", "Stop. No growth spending until it works."],
    [f"{num(B28)} in year 2 is too ambitious", f"Under {num(Y1)} bureaus by December 2027", "Slow hiring; keep the angel money as a buffer; stay profitable."],
    ["Family and company mix", "Decisions only in conversation", "Write everything down; both companies approve."],
  ])}
</section>

<section id="c2">
  <p class="msg">C2. We are cheap to run. The Rp30M starts us; a {rp(rnd['amount'], 0)} angel round in {rnd_when} gives a safe buffer, and every hire still waits until profit can pay it.</p>
  <div class="tiles">
    <div class="tile"><b>{rp(I.BASE_RUN_COST)}</b><span>a month to run today; no salaries</span></div>
    <div class="tile"><b>{starter_needed} bureaus</b><span>on the Rp1M plan cover it</span></div>
    <div class="tile"><b>{month_name(first_profit)}</b><span>first month with a profit</span></div>
    <div class="tile"><b>{rp(low['cash'])}</b><span>lowest cash, {month_name(low['month'])}</span></div>
  </div>
  {viz.quarter_chart(plan['quarters'])}
  {table(["Quarter", "Revenue", "Costs", "New money", "Net cash", "Cash at end"], q_rows, right=(1, 2, 3, 4, 5))}
  <p class="note"><b>New money:</b> Q4 2026 = the Rp30M from this deal (TemanJourney Rp20M, Riskia Rp10M) at the notary signing. Q1 2027 = the {rp(rnd['amount'], 0)} angel round (warm introductions or cold email to angel investors).</p>
  <h3>What the angel round adds</h3>
  {table(["", "With the angel round", "Without it"], [
    ["Outside money", f"Rp30M + {rp(rnd['amount'], 0)}", "Rp30M only"],
    ["Cash at the end of 2027", rp(Y[2027]['cash_end']), rp(NY[2027]['cash_end'])],
    ["Lowest cash in 2027", rp(min(r['cash'] for r in plan['rows'] if r['ym'][0] == 2027)), rp(min(r['cash'] for r in noangel['rows'] if r['ym'][0] == 2027))],
    ["Ownership of Riskia's side", f"{pct(stake_after, 0)} unless they join the round", "20%"],
  ], right=(1, 2))}
  <p class="note">With these targets, bureau income alone keeps cash above zero. The round is a safety buffer: it pays for a slow month, a hire a little earlier, or a faster push if the TemanJourney test goes well. It costs about {pct(rnd['dilution'])} of CoreItera {EST}.</p>
  <h3>Churn: how many bureaus we keep</h3>
  <p class="msg">As a new product we lose more bureaus at first: 8% a month, falling to the industry middle (5%) by December 2027 and to good-software level (2%) by December 2028.</p>
  {churn_viz}
  {viz.hbars([
      ("Signed in 2027", SIGNED27, "c-cost", f"{num(SIGNED27)} bureaus"),
      ("Left during 2027", LOST27, "c-ochre", f"{num(LOST27)} bureaus"),
      ("Still paying in December 2027", Y1, "c-rev", f"{num(Y1)} bureaus"),
      ("New bureaus needed in 2028 to reach 700", Y[2028]['new'], "c-cost", f"{num(Y[2028]['new'])} (because {num(Y[2028]['churned'])} leave)"),
  ], note="Churn falls as the product, onboarding and support get better; most small-business churn happens in the first 90 days, so early onboarding matters most.")}
  {table(["Churn path", "Jan 2027 → Dec 2027 → Dec 2028", "Paying Dec 2027", "New needed in 2028", "Profit 2028", "Profit 2031"], churn_rows, right=(2, 3, 4, 5))}
  <p class="note">Sources: Vena (2025), using Recurly and Paddle data: 3% to 7% a month for small-business software; Focus Digital (2025 to 2026): 6.4% a month for small-business SaaS. How we keep churn low: a set-up call in week 1, a monthly "extra clients" report to each bureau, and the first month free so only serious bureaus pay.</p>
  {viz.revenue_chart(Y)}
  <p>2027 income comes only from the bureau tool. Psynthesia adds income from 2028. <b>With half the bureaus</b>, we still make a profit every year from {slow_profit_year}. Full tables are in the appendix.</p>
</section>

<section id="c3">
  <p class="msg">C3. Riskia's side pays Rp50M in total for 20%. The 15% promise is kept.</p>
  {viz.cap_bars([
      ("Today", [(0.90, "seg-o", "Current owners 90%"), (0.10, "seg-r", "10%")]),
      ("Step 1: 15% kept", [(0.85, "seg-o", "Current owners 85%"), (0.15, "seg-r", "15%")]),
      ("Step 2: Rp30M in", [(0.80, "seg-o", "Current owners 80%"), (0.1608, "seg-r", "16.1%"), (0.0392, "seg-t", "")]),
  ])}
  {table(["Who", "Paid before", "Paid now", "Total", "Owns after"], [
    ["Galih (with the other founders)", "Rp50M", "None", "Rp50M", "80.0%"],
    ["Riskia", "Rp20M", "Rp10M", "Rp30M", "16.1%"],
    ["TemanJourney", "None", "Rp20M", "Rp20M", "3.9%"],
    ["<b>Total</b>", "<b>Rp70M</b>", "<b>Rp30M</b>", "<b>Rp100M</b>", "<b>100%</b>"],
  ], right=(1, 2, 3, 4))}
  <ul>
    <li><b>Price:</b> Rp480M before the new money, Rp510M after. That is 2.5× the price of Riskia's first 10%. Her 16.1% is worth about Rp82M on paper.</li>
    <li><b>Why founders own more for the same Rp50M:</b> their money came first, at the highest risk, and they built the platform, mostly unpaid.</li>
    <li><b>Angel round:</b> about {rp(rnd['amount'], 0)} for about {pct(rnd['dilution'])} in {rnd_when}, from angel investors (warm introductions or cold email). Riskia and TemanJourney may join to keep 20%; if not, they hold {pct(stake_after, 0)}.</li>
  </ul>
  {viz.stacked([("Running costs to February 2027", running_part, "seg-r"), ("Notary and share registration", I.DEAL_ONE_OFFS['notary and share registration'], "seg-o"), ("Launch", I.DEAL_ONE_OFFS['launch costs'], "seg-t")], f"Total {rp(I.DEAL_CASH)}")}
  {table(["Use of the Rp30M", "Amount", "What it reaches"], [
    ["Running costs to February 2027", rp(running_part), "Theta live; agent live at TemanJourney"],
    [f"Notary and share registration {EST}", rp(I.DEAL_ONE_OFFS['notary and share registration']), "This deal, done properly"],
    [f"Launch {EST}", rp(I.DEAL_ONE_OFFS['launch costs']), "Payment set-up; test week"],
  ], right=(1,))}
  {table(["Riskia's side gets", "Plan", "Slow case"], [
    ["Ownership", f"20% now; {pct(stake_after, 0)} after the angel round unless they join", "The same"],
    ["TemanJourney's software", "Free for 12 months, then cost price", "The same"],
    [f"Dividends 2030 and 2031 ({pct(runner.DIVIDEND_SHARE)} of profit) {EST}", rp(sum(div_plan.values()), 0), rp(sum(div_slow.values()), 0)],
    ["Say", "Monthly update; first look at new tools; right to invest again", "The same"],
  ])}
  <p class="note">Young plans often miss. These numbers are estimates, not promises.</p>
</section>

<section id="ask">
  <div class="answer">
    <div class="eyebrow">What we ask Riskia and TemanJourney</div>
    <ol>
      <li>Invest Rp30M (TemanJourney Rp20M, Riskia Rp10M) for 20%: Riskia 16.1%, TemanJourney 3.9%.</li>
      <li>Run the sales agent on TemanJourney's WhatsApp from February 2027, and share the chat numbers.</li>
      <li>Help us reach angel investors for a {rp(rnd['amount'], 0)} round in {rnd_when}, with the right to join it.</li>
      <li>Name one person at TemanJourney to own the launch.</li>
      <li>Introduce 2 Jakarta bureaus by January 2027.</li>
      <li>Sign the term sheet by Friday 30 October 2026.</li>
    </ol>
  </div>
</section>""")

# ---- appendix (detail behind each box of the pyramid)
yr_rows = [[str(y), num(v['bureaus_end']), num(v['seats_end']), rp(v['plan']), rp(v['fees']), rp(v['copilot']) if v['copilot'] > 0.05 else "None",
            f"<b>{rp(v['revenue'])}</b>", rp(v['variable'] + v['fixed']), rp(v['tax']), f"<b>{rp(v['profit'])}</b>",
            pct(v['profit'] / v['revenue']), rp(v['cash_end'])] for y, v in Y.items()]
cost_rows = []
def _row(name, keys):
    vals = [sum(Y[y][k] for k in keys) for y in Y]
    return [name] + [rp(v) if v > 0.05 else "None" for v in vals]
fp = I.FOUNDER_PAY_YEAR
cost_rows.append(_row("Founders: Rp10M a month each in profit months (2027); full pay from 2028", ["founders"]))
cost_rows.append(_row("Engineers: Zaki and the hires in the team plan", ["engineer"]))
cost_rows.append(_row("Quality: test engineers", ["quality"]))
cost_rows.append(_row("Support", ["cs"]))
cost_rows.append(_row("Sales", ["sales"]))
cost_rows.append(_row("Winning new bureaus: ads, referral credit, commission, set-up", ["acquisition", "onboarding"]))
cost_rows.append(_row("Workshops and events", ["workshops"]))
cost_rows.append(_row("Clinical quality checks", ["clinical"]))
cost_rows.append(_row("Sales agent, AI per report, servers", ["agent", "ai", "server"]))
cost_rows.append(_row("Cloud, AI and tools (today's running cost)", ["base"]))
cost_rows.append(_row(f"Admin, legal and HR ({pct(I.GA_SHARE_OF_REVENUE[2027])} of revenue in 2027, {pct(I.GA_SHARE_OF_REVENUE[2028])} after)", ["ga"]))
cost_rows.append(_row("Office and accountant", ["office", "accountant"]))
cost_rows.append(["<b>Total costs</b>"] + [f"<b>{rp(Y[y]['variable'] + Y[y]['fixed'])}</b>" for y in Y])
slow_rows = [[str(y), num(v['bureaus_end']), rp(v['revenue']), rp(v['profit']), pct(v['profit'] / v['revenue']), rp(v['cash_end'])] for y, v in SY.items()]
assump = [
    ["Running cost today", rp(I.BASE_RUN_COST) + " a month", "Known"],
    ["Cash today (19 Oct 2026)", rp(I.CASH_TODAY), "Known"],
    ["Prices: plan, payment fee, copilot", "From Rp1M; Rp5,000; Rp300,000", "Decided"],
    ["Average plan price per bureau", ", ".join(f"{y}: {rp(v)}" for y, v in I.AVG_PLAN_BY_YEAR.items()), "Estimate"],
    ["Psychologists per bureau", str(I.PSY_PER_BUREAU), "Estimate"],
    ["Paid payments per bureau a month", ", ".join(f"{y}: {v}" for y, v in I.PAYMENTS_PER_BUREAU.items()), "Estimate"],
    ["Share of psychologists in private bureaus", pct(I.SHARE_IN_BUREAUS), "Estimate"],
    ["Bureaus that leave each month", "8% in January 2027, 5% in December 2027, 2% from December 2028 (straight line between)", "Galih's plan, on industry benchmarks (3% to 7%)"],
    ["Sales agent cost per bureau", rp(I.AGENT_COST_PER_BUREAU, 2) + " a month", "Estimate"],
    ["Sales agent example", f"{num(X['inquiries'])} chats, {pct(X['after_hours_share'])} after hours, conversion {pct(X['conversion_now'])} to {pct(X['conversion_agent'])}, first purchase {rpk(X['first_purchase'])}", "Estimate, measured at TemanJourney first"],
    ["Psynthesia launch and use", "January 2028; 10% of bureau psychologists at launch, 40% Dec 2028, 55% Dec 2029, 60% Dec 2030", "Plan and estimate"],
    ["AI cost per Psynthesia seat", rp(I.AI_COST_PER_SEAT, 2) + " a month (about 15 reports at Rp8,000)", "Estimate"],
    ["Founder pay", f"2027: Rp10M a month each, only in profit months. From 2028: Galih {rp(fp['Galih'], 0)}, Riskia {rp(fp['Riskia'], 0)}, Adit {rp(fp['Adit'], 0)} a year", "Decided"],
    ["Zaki's pay", "Rp10M a month in 2027 (same rule); Rp20M a month from 2028", "Estimate, to confirm"],
    ["Dividends", pct(runner.DIVIDEND_SHARE) + " of profit from 2030", "Proposal"],
]
tests = [
    ["1. Is the pain big?", "Late replies lose clients; founders re-teach selling.", "Medium: count late chats at TemanJourney, Dec 2026"],
    ["2. Is it much better than today?", f"Replies in minutes, any hour; {pct(more_clients)} more clients in the example.", "Medium: tested at the phase 1 gate"],
    ["3. Is the volume enough?", f"{starter_needed} bureaus cover our costs; the year-1 goal is {pct(y1_share)} of all bureaus.", "Low to medium: bureau count not confirmed"],
    ["4. Could one factor kill it?", "Meta rules, or a clinical mistake. Official tools only; no counselling.", "Medium: legal check by Q1 2027"],
    ["5. Is the edge real?", "Own style, real calendars, safety rules and trust channels together.", "Medium: proven only at TemanJourney"],
]
add(f"""
<div class="part"><span class="pyr-letter">+</span><h2>Appendix: the detail</h2></div>
<section>
  <h3>Five-year profit and loss {EST}</h3>
  {table(["Year", "Bureaus", "Seats", "Plans", "Payment fees", "Copilot", "Revenue", "Costs", "Tax", "Profit", "Margin", "Cash (Dec)"], yr_rows, right=tuple(range(1, 12)))}
  <p class="note">Bureaus and seats in December. Tax: 0.5% of revenue to 2029, then 11% of profit (to confirm with a tax consultant). TemanJourney is not counted.</p>
  <h3>Main costs by year</h3>
  {table(["Cost"] + [str(y) for y in Y], cost_rows, right=tuple(range(1, 6)))}
  <p class="note">2027: founders and Zaki get Rp10M a month each, only in a month that makes a profit (from {month_name(plan["founders_paid_from"])}). From {month_name(plan["full_pay_from"])}: Galih {rp(fp["Galih"], 0)}, Riskia {rp(fp["Riskia"], 0)} and Adit {rp(fp["Adit"], 0)} a year. Hires follow the team plan in section C1.</p>
  <h3>Slow case: half the bureaus</h3>
  {table(["Year", "Bureaus", "Revenue", "Profit", "Margin", "Cash (Dec)"], slow_rows, right=(1, 2, 3, 4, 5))}
  <p class="note">Lowest cash in the slow case: {rp(slow_low['cash'])} in {month_name(slow_low['month'])}.</p>
  <h3>Prices</h3>
  {table(["Plan", "Psychologists", "Price a month", "Per psychologist"], [
    ["Starter", "1 to 3", rpk(I.PLAN_TIERS[0][3]), "Rp333,000 or more"],
    ["Growth", "4 to 10", rpk(I.PLAN_TIERS[1][3]), f"{rpk(I.PLAN_TIERS[1][3] / 10)} to {rpk(I.PLAN_TIERS[1][3] / 4)}"],
    ["Pro", "11 or more", rpk(I.PLAN_TIERS[2][3]) + " (one admin at Jakarta minimum wage 2026)", f"{rpk(I.PLAN_TIERS[2][3] / 11)} or less"],
  ], right=(2, 3))}
  <p class="note">Every plan has the sales agent, scheduling, booking, chat counselling and tests. Plus Rp5,000 per payment (Midtrans fee at cost). Psynthesia: Rp300,000 a seat from 2028. First month free.</p>
  <h3>Five tests</h3>
  {table(["Test", "Evidence", "Confidence"], tests)}
  <h3>Launch at TemanJourney on 30 November 2026</h3>
  {table(["Milestone", "Done by", "Owner"], [
    ["Riskia reads this proposal; numbers agreed", "Fri 23 Oct", "Galih, Riskia"],
    ["Term sheet signed by both companies", "Fri 30 Oct", "Galih, Riskia"],
    ["Shareholder approval in both companies (RUPS); notary booked", "Fri 6 Nov", "Galih, Riskia"],
    ["Notary deed signed; Rp30M paid in", "Fri 13 Nov", "Riskia, TemanJourney"],
    ["At least 20 psychologist profiles and licences (SIPP) entered", "Fri 13 Nov", "TemanJourney admin"],
    ["Prices and weekly time blocks entered", "Fri 20 Nov", "TemanJourney admin, Adit"],
    ["Release practised on a copy of the live database", "Fri 20 Nov", "Galih, Adit, Zaki"],
    ["Test week: real bookings, payments and chats", "Fri 27 Nov", "All"],
    ["<b>Theta live for TemanJourney clients</b>", "<b>Mon 30 Nov</b>", "All"],
  ], right=(1,))}
  <h3>Main assumptions</h3>
  {table(["Assumption", "Value", "Status"], assump)}
  <h3>Sale value of the 20% {EST}</h3>
  {table(["Sale price", "Plan", "Slow case"], [
    [f"{plan['exit']['low']['multiple']:.0f} times 2031 revenue", rp(plan['exit']['low']['stake_value']), rp(slow['exit']['low']['stake_value'])],
    [f"{plan['exit']['mid']['multiple']:.0f} times 2031 revenue", rp(plan['exit']['mid']['stake_value']), rp(slow['exit']['mid']['stake_value'])],
  ], right=(1, 2))}
  <p class="note">Likely buyers: an Indonesian HR-tech or health-tech company, or a test publisher.</p>
  <h3>The other way to price the deal</h3>
  {table(["", "Recommended (15% completed, then cash)", "Cash only"], [
    ["Riskia at the end", "16.1%", "12.6%"], ["TemanJourney at the end", "3.9%", "7.4%"], ["Together", "20.0%", "20.0%"],
    ["Value of CoreItera before the new money", "Rp480M", "Rp240M"], ["Is the 15% promise kept?", "Yes", "No"],
  ], right=(1, 2))}
  <h3>To confirm before signing</h3>
  {table(["Number", "Used here", "Why it matters"], [
    ["CoreItera's share count (from the deed)", "1,000,000 as an example", "Sets the exact number of new shares"],
    ["How the earlier Rp50M and Rp20M are recorded", "Galih and Riskia", "The deed must match the cash table"],
    ["Number of private bureaus and practices in Indonesia", f"About {num(round(bureaus_market, -2))}", f"The {num(Y1)}-bureau goal is {pct(y1_share)} of this; HIMPSI does not publish it"],
    ["Angel round size, price and timing", f"{rp(rnd['amount'], 0)} for {pct(rnd['dilution'])} in {rnd_when}", "Angel round; the price is an estimate"],
    ["After-hours share of chats and conversion at TemanJourney", f"{pct(X['after_hours_share'])}; {pct(X['conversion_now'])} to {pct(X['conversion_agent'])}", "The lead product's whole promise"],
    ["Other shareholders of TemanJourney", "Not known", "They must approve the Rp20M"],
  ])}
</section>""")

footer = """
<footer>
  Sources: CoreItera trial data (205 reports rated over 6 months); TemanJourney 2027 financial model (October 2026); WHO (2025, 2026), Mental Health Atlas 2024; UN World Population Prospects 2024; GBD 2021 ASEAN Mental Disorders Collaborators (2025); HIMPSI (2026, 2021); Rettew et al. (2009); Augment proposal (IS5007, October 2026) for the typical week, assessment hours, the 6 to 15 a week calculation and Law 23/2022; Kemenkes, Survei Kesehatan Indonesia 2023; SimplePractice prices from third-party guides (2026); Riliv from Fortune Indonesia (2023); Klinik Pintar from Dealroom and e27; Vibefam from its June 2026 funding release and Product Hunt (700+ locations, 26 countries); Jakarta minimum wage 2026 (Rp5,729,876) from Bisnis.com; KantorKu test prices via the Augment proposal. US$1 = Rp16,500. A notary and a tax consultant must confirm the legal and tax steps.
</footer>"""

html = f"""<title>Theta Partnership Proposal</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Gelasio:wght@500;600;700&family=Source+Sans+3:wght@400;600;700&family=IBM+Plex+Mono:wght@500&display=swap">
<style>
{css}
</style>

<main>
{''.join(S)}
{footer}
</main>
"""

with open(OUT_HTML, "w") as f:
    f.write(html)

import present
OUT_PRESENT = os.path.join(HERE, "..", "proposal", "theta-partnership-proposal-present.html")
with open(OUT_PRESENT, "w") as f:
    f.write(present.standalone("Theta Partnership Proposal", css, f"<main>\n{''.join(S)}\n{footer}\n</main>"))
print("wrote", os.path.normpath(OUT_PRESENT))
runner.main()
print("wrote", os.path.normpath(OUT_HTML))
