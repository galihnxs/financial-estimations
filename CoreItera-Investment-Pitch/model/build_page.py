"""Build the proposal page (v9: one message per slide, MECE pyramid, deck-ready) from the model, so every number on the page comes from engine.py.

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
stake_after = plan["riskia_side"]["stake"] * (1 - I.NEXT_ROUND["dilution"])   # if the optional round happens and they do not join
top_plan = I.PLAN_TIERS[-1][3]
withangel = engine.run(I, True)
WY = withangel["years"]
_s = __import__("types").SimpleNamespace(**{k: v for k, v in vars(I).items() if k.isupper()})
_s.GROSS_ADDS = {k: (0 if k <= (2027, 3) else v) for k, v in I.GROSS_ADDS.items()}
late = engine.run(_s, False)["lowest_cash"]          # stress: no bureau pays before April 2027
low27 = min(r["cash"] for r in plan["rows"] if r["ym"][0] == 2027)
low27_m = min((r for r in plan["rows"] if r["ym"][0] == 2027), key=lambda r: r["cash"])["label"]
Q_NEW = [round(sum(v for k, v in GA.items() if k[0] == 2027 and (k[1] - 1) // 3 == q)) for q in range(4)]
B26 = round(GA[(2026, 12)])
B28 = Y[2028]["bureaus_end"]
def hc_at(label):
    r_ = next(r for r in plan["rows"] if r["label"] == label)
    return r_["headcount"]
TEST_FEE_LOW, TEST_FEE_HIGH = 50_000, 150_000   # K  KantorKu online recruitment tests, per candidate (Augment A3)

fp = I.FOUNDER_PAY_YEAR
yes, no = '<span class="yes">Yes</span>', '<span class="no">No</span>'

# ---- header + mission

# ---- the answer (top of the pyramid)

pain_loop = viz.flywheel([
    ("Hundreds of leads", "Instagram, WhatsApp"),
    ("Slow replies", "under 5% close"),
    ("Founder re-teaches", "2 to 4 h a week"),
    ("Spreadsheet matching", "clients give up"),
    ("No feedback asked", "good reviews lost"),
    ("More ads", "to replace lost clients"),
], ("The endless loop", "for founders and admins"), mid="pl-arrow", extra="pl")
# ---- PART A

# ---- PART B
live = lambda t: f'<span class="pill live">{t}</span>'
soon = lambda t: f'<span class="pill soon">{t}</span>'
PF_ROWS = ([
    (1, "Hundreds of leads", "good and poor leads look the same", "Lead sorting", "the agent asks need, budget and time, and marks serious leads", "The admin sees serious leads first", soon("Feb 2027")),
    (2, "Slow replies", f"night chats wait; under {pct(X['conversion_now'])} close", "24-hour sales agent", "replies in minutes on WhatsApp (Instagram in May), offers open times and a payment link", f"Reply in under 5 minutes; closing {pct(X['conversion_now'])} → {pct(X['conversion_agent'], 1)}", soon("Feb 2027")),
    (3, "Founder re-teaches", "2 to 4 hours a week of SOPs", "SOP inside the agent", "written once, in the bureau's own style; the admin gets a summary of every lead", "Founder SOP time under 1 hour a week", soon("Feb 2027")),
    (4, "Spreadsheet matching", "clients wait, then give up", "Self-booking", "clients pick a psychologist's real free time; the admin only solves clashes", "Most bookings need no admin work", live("Live 30 Nov 2026")),
    (5, "No feedback asked", "good reviews lost", "Feedback and review requests", "asked the same day; happy clients get a Google Maps link", "Feedback from most sessions; new Google reviews every month", soon("Jan 2027")),
    (6, "More ads", "to replace lost clients", "Follow-ups and rebooking", "reminders bring clients back (email now, WhatsApp in Q1 2027); reviews and referrals bring new ones", "More clients book again; lower cost per new client", live("Email live")),
])
pain_fix_win = viz.pain_fix(PF_ROWS[:3])
pain_fix_keep = viz.pain_fix(PF_ROWS[3:])
pos = viz.positioning([
    ("CoreItera", 88, 88, True),
    ("Admin on WhatsApp", 10, 50, False),
    ("Generic chatbots (Qontak, Wati)", 70, 18, False),
    ("Global practice software", 8, 34, False),
    ("Counselling apps (Riliv, Satu Persen)", 42, 70, False),
], "Sells for the bureau 24 hours", "Built for Indonesian psychology")

# ---- PART C
road = viz.roadmap([
    {"name": "1. TemanJourney first", "when": "Nov 2026 to Mar 2027", "gate": f"TemanJourney wins more clients; {num(B27[(2027, 3)])} bureaus pay",
     "cells": {"Product": "Booking and chat live 30 Nov; WhatsApp agent in Feb", "Customers": f"{num(B26)} bureaus + TemanJourney by December 2026; +{num(Q_NEW[0])} in Q1 2027 (7 a month). Galih and Riskia close the deals; no sales staff",
               "Money": f"Rp30M in; {rp(I.BASE_RUN_COST)} a month", "Team": "3 founders + TemanJourney psychologists"}},
    {"name": f"2. Grow steadily", "when": "Apr to Dec 2027", "gate": f"{num(Y1)} bureaus still paying; churn down to 5% a month",
     "cells": {"Product": "Set up in one day; tests; Instagram agent", "Customers": f"+{num(Q_NEW[1])} in Q2, +{num(Q_NEW[2])} in Q3, +{num(Q_NEW[3])} in Q4: {num(SIGNED27)} signed, {num(Y1)} still paying by December",
               "Money": f"Own income; first profit {month_name(first_profit)}. Optional {rp(rnd['amount'], 0)} round from {rnd_when} if traction is good", "Team": f"Each hire when profit pays it. By December: {hc_at("2027-12")["Sales"]} sales, {hc_at("2027-12")["Support"]} support, {hc_at("2027-12")["Engineers"]} engineers, {hc_at("2027-12")["Quality"]} quality"}},
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
    ("Feedback and review requests", "Zaki", 2, 1, "c-rev"),
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
    ["Company tax", f"0.5% of turnover to 2027 (registered {I.TAX_REGISTERED}); from 2028, 22% of profit with the Pasal 31E cut (11% on the first Rp4.8 billion of turnover)", "Law; a tax consultant to confirm"],
]
tests = [
    ["1. Is the pain big?", "Late replies lose clients; founders re-teach selling.", "Medium: count late chats at TemanJourney, Dec 2026"],
    ["2. Is it much better than today?", f"Replies in minutes, any hour; {pct(more_clients)} more clients in the example.", "Medium: tested at the phase 1 gate"],
    ["3. Is the volume enough?", f"{starter_needed} bureaus cover our costs; the year-1 goal is {pct(y1_share)} of all bureaus.", "Low to medium: bureau count not confirmed"],
    ["4. Could one factor kill it?", "Meta rules, or a clinical mistake. Official tools only; no counselling.", "Medium: legal check by Q1 2027"],
    ["5. Is the edge real?", "Own style, real calendars, safety rules and trust channels together.", "Medium: proven only at TemanJourney"],
]


# ===================================================================== SLIDES (v9)
# One slide = one message (its title) + one main visual + at most a short note.
# Order follows the pyramid: answer first, then A (why), B (what), C (how), then the ask; detail in the appendix.
SL = []


def slide(sid, sec, title, body, note="", kind="content", lead=""):
    SL.append({"id": sid, "sec": sec, "title": title, "lead": lead, "body": body, "note": note, "kind": kind})


def divider(sid, sec, letter, title, points):
    SL.append({"id": sid, "sec": sec, "title": title, "letter": letter, "points": points, "kind": "divider"})


# ------------------------------------------------------------------ opening
slide("cover", "Start", "Theta Partnership Proposal", f"""
  <p class="mission-text">Helping TemanJourney, and then every psychology bureau, serve more clients with less work.</p>
  <p class="meta">CoreItera · proposal for Riskia and TemanJourney · Galih Muji Agung, CEO · 20 October 2026 · Draft, version 10</p>""", kind="cover")

slide("look-back", "Start", "Before the pitch: we built many things, but not the one TemanJourney needed most", viz.journey([
    (1, "The first idea", "Riskia working long hours to write one psychologist report",
     "An online testing platform, not a simple report helper she could use", "It brought some business, but her report time did not fall", ""),
    (2, "Synthesia", "Psychologists need help with the whole workflow",
     "A big psychologist workspace as our first product", "Too big: nobody really tested it", ""),
    (3, "B2B testing", "A company client asked for online tests",
     "Online testing for companies", "Revenue came, but daily operations stayed hard", ""),
    (4, "Theta, now", "TemanJourney's B2B grew; Riskia was exhausted running B2C",
     "A platform to help TemanJourney run as a bureau", "We found the real problems: admin training, money on ads, deals not closing, feedback lost", "jy-now"),
  ]), note="Each step taught us something. The pattern: we built what was interesting to build, and too much at once, before checking it solved Riskia's daily problem.")

slide("lessons", "Start", "This time: small, right, and tested at TemanJourney first", table(["What went wrong", "What we do now", "Where you see it"], [
    ["We built what was interesting, not what Riskia needed every day", "Start from TemanJourney's daily problem: the endless loop of leads, slow replies, re-teaching, matching and lost feedback", "A2"],
    ["The MVP was too big, and nobody tested it", "A small MVP, used by TemanJourney first; we sell it to others only after the proof numbers are met", "B1, C1"],
    ["Features made money but did not fix operations", "Every feature fixes one step of the loop and has one proof number", "B1"],
    ["Too many products at once", "One product in year 1 (the bureau tool); Psynthesia only in year 2; abroad only in year 3", "C1"],
    ["Growth before the base was strong", "Make TemanJourney succeed first; it becomes the proof for every other bureau", "C1, the ask"],
  ]), note="Just enough, and right: no feature is built unless it solves a real problem for TemanJourney and helps it grow.")

slide("vision", "Start", "Our vision and mission, set right this time", f"""
  <div class="vm">
    <div class="card wide"><b class="k">Vision</b><p>{I.VISION}</p></div>
    <div class="card"><b class="k">First step (2027)</b><p>{I.FIRST_STEP}</p></div>
    <div class="card"><b class="k">Mission (2027 to 2031)</b><p>{I.MISSION.split(': ', 1)[1][0].upper() + I.MISSION.split(': ', 1)[1][1:]}</p></div>
  </div>""", note="Responsible business: firms whose work changes people's lives, where quality and trust matter more than speed. Mental health firms are the first.")

slide("answer", "Start", "The answer in one picture", viz.pyramid(
      f"Invest <b>Rp30M</b> for <b>20%</b>. We use it to help <b>TemanJourney</b> win more clients with a <b>24-hour sales agent</b>. Then we sign <b>{num(SIGNED27 + B26)} bureaus</b> by December 2027, reach <b>{num(B28)}</b> in year 2, and enter <b>Malaysia or Singapore</b> in year 3.",
      [("A", "The need is big", [
            ("a1", f"Too few psychologists: {per100k:.0f} per 100,000 in ASEAN"),
            ("a2", "Founders and admins are stuck in an endless loop"),
            ("a3", f"Indonesia first: about {num(round(bureaus_market, -2))} bureaus")]),
       ("B", "One product first: the bureau tool", [
            ("b1", "The MVP fixes each step of the loop"),
            ("b2", "More clients for less than one admin"),
            ("b3", "Only we sell 24 hours in the bureau's own style")]),
       ("C", "A safe plan and deal", [
            ("c1", "Year 1: bureaus only. Year 2: Psynthesia"),
            ("c2", "Cheap to run; we grow on our own income"),
            ("c3", "Rp50M in total for 20%; the 15% promise kept")])]))

slide("summary", "Start", "Executive summary: a small, proven tool first, then 700 bureaus", f"""
  <div class="exs">
    <ol>
      <li><b>Situation</b><span>Too few psychologists ({per100k:.0f} per 100,000 in ASEAN). Indonesian bureaus win clients in Instagram and WhatsApp chats.</span></li>
      <li><b>Problem</b><span>Founders and admins are stuck in a loop: hundreds of leads, under {pct(X['conversion_now'])} close, founders re-teach SOPs 2 to 4 hours a week, schedules are matched in spreadsheets, feedback is never asked.</span></li>
      <li><b>Solution</b><span>A small MVP that fixes each step: a 24-hour sales agent in the bureau's own style, self-booking, and feedback requests. Tested at TemanJourney first.</span></li>
      <li><b>Plan</b><span>{num(SIGNED27 + B26)} bureaus signed by December 2027, {num(B28)} paying in year 2, Malaysia or Singapore in year 3. Psynthesia seats from year 2.</span></li>
      <li><b>Ask</b><span>Rp30M now (TemanJourney Rp20M, Riskia Rp10M) for 20%: Riskia 16.1%, TemanJourney 3.9%.</span></li>
    </ol>
    <div class="kn">
      <div><strong>{pct(more_clients)} more</strong><span>clients for a bureau from the sales agent (estimate)</span></div>
      <div><strong>Rp1M to {rp(top_plan, 1)}</strong><span>a month for a bureau; never more than one admin</span></div>
      <div><strong>{month_name(first_profit)}</strong><span>first month with a profit, on our own income</span></div>
      <div><strong>{num(B28)} bureaus</strong><span>paying in year 2 (December 2028)</span></div>
    </div>
  </div>""", note="All numbers are estimates from the CoreItera model (model/inputs.py); details in the appendix.")

# ------------------------------------------------------------------ A. why
divider("part-a", "A", "A", "The need is big", ["Too few psychologists", "Founders and admins stuck in a loop", "Indonesia is the market to start"])

slide("a1", "A", "A1. There are too few psychologists, and training more takes too long", viz.hbars([
      ("People with a mental disorder (six ASEAN countries)", I.ASEAN6['with_disorder_m'], "c-cost", f"{I.ASEAN6['with_disorder_m']:.1f} million"),
      ("Full assessments possible each year", I.ASEAN6['assessments_per_year_m'], "c-ochre", f"{I.ASEAN6['assessments_per_year_m']:.1f} million ({pct(asean_share, 1)})"),
      ("The same psychologists with CoreItera", asean_after, "c-rev", f"{asean_after:.1f} million ({assess_x:.1f}×)"),
  ]) + f"""
  <div class="tiles three">
    <div class="tile"><b>1 billion+</b><span>people in the world with a mental health condition (WHO)</span></div>
    <div class="tile"><b>{per100k:.0f} per 100,000</b><span>psychologists in our six ASEAN countries</span></div>
    <div class="tile"><b>{I.ID['years_to_close_gap']} years</b><span>to train enough clinical psychologists in Indonesia (HIMPSI)</span></div>
  </div>""",
  note=f"So each psychologist must help more people: {pct(more_clients)} more clients per bureau (sales agent), {assess_x:.1f}× more assessments (Psynthesia, year 2), {agree_x:.1f}× more accurate diagnosis (a target). Six countries: Indonesia, Philippines, Vietnam, Thailand, Malaysia, Singapore. Sources: WHO; GBD 2021 ASEAN (2025); HIMPSI; Augment proposal (IS5007).")

slide("a2", "A", "A2. Founders and admins are stuck in an endless loop", pain_loop,
      lead=f"Hundreds of leads come in, under {pct(X['conversion_now'])} close, and the same problems come back every week.",
      note="Founder hours and the closing rate come from bureau founders. We measure every step at TemanJourney in December 2026.")

slide("a2-cost", "A", "A2. Every turn of the loop costs the bureau clients and the founder time", table(["Step", "What happens today", "What it costs"], [
    ["<b>Leads</b>", "Hundreds a month from Instagram and WhatsApp; good and poor leads look the same", f"Under {pct(X['conversion_now'])} become clients"],
    ["<b>Replies</b>", f"The admin works {ADMIN_FROM:02d}:00 to {ADMIN_TO:02d}:00; night and weekend chats wait until morning", f"About {pct(X['after_hours_share'])} of chats answered late"],
    ["<b>Founder</b>", "Teaches the admin almost every day how to close, follow up and show empathy", "2 to 4 hours a week on SOPs, instead of marketing and hiring"],
    ["<b>Schedule</b>", "Client and psychologist matched by hand in a spreadsheet", "Clients wait, and some give up"],
    ["<b>After the session</b>", "Nobody asks for feedback", "Good Google Maps reviews are lost"],
  ]) + """
  <div class="callout"><p><b>An admin is worth more than making the first match.</b> The admin should solve deadlocked schedules, close warm leads and look after clients.</p></div>""")

slide("a2-night", "A", f"A2. About {pct(X['after_hours_share'])} of chats arrive when the admin is away", viz.day_band(ADMIN_FROM, ADMIN_TO, X['after_hours_share']),
      note=f"Admin hours and the {pct(X['after_hours_share'])} after-hours share are estimates; we count them at TemanJourney in December 2026.")

slide("a3", "A", "A3. Indonesia is the biggest market, and we can reach it first", viz.hbars([
      ("All psychologists in Indonesia", tam, "c-cost", f"{rp(tam)} a year"),
      (f"In private practice (about {num(round(bureaus_market, -2))} bureaus)", sam, "c-ochre", f"{rp(sam)} a year"),
      (f"Year-1 plan: {num(Y1)} bureaus ({pct(y1_share)})", Y[2027]['revenue'], "c-rev", f"{rp(Y[2027]['revenue'])} in 2027"),
  ]) + table(["Why now", "What it means"], [
    ["Clients buy in chat", "The first bureau to reply wins."],
    ["Data law (UU PDP 27/2022)", "Client data on phones is a legal risk."],
    ["Psychology law (UU 23/2022)", "Bureaus need proper records."],
    ["AI can now draft", "The bureau still decides."],
  ]), note=f"Value = psychologists × {rpk(per_psy_month)} a month each. The bureau count is our estimate: HIMPSI does not publish one.")

# ------------------------------------------------------------------ B. what
divider("part-b", "B", "B", "One product first: the bureau tool", ["The MVP fixes each step of the loop", "More clients for less than one admin", "Only we sell 24 hours in the bureau's own style"])

slide("b1", "B", "B1. Our MVP breaks the loop (1 of 2): it wins more clients", pain_fix_win,
      note="Steps 1 to 3 of the loop: the leads, the replies and the founder's time. All live at TemanJourney by February 2027.")
slide("b1-keep", "B", "B1. Our MVP breaks the loop (2 of 2): it serves and keeps them", pain_fix_keep,
      note="Steps 4 to 6: the schedule, the feedback and the next visit. Every target is an estimate, checked at TemanJourney before we sell to others.")

slide("b1-parts", "B", "B1. Most of the MVP is already built, so we can deliver", viz.product_stack([
      ("Year 1", "ps-lead", "The bureau tool (new)", [("Sales agent on WhatsApp", "Feb 2027"), ("Lead sorting", "Feb 2027"), ("Feedback and review requests", "Jan 2027"), ("Self-serve set-up", "Mar 2027"), ("Scheduling for any bureau", "Mar 2027"), ("Sales agent on Instagram", "May 2027")]),
      ("Already built", "ps-built", "What the agent sells into", [("Booking and payment", "live 30 Nov 2026"), ("Chat counselling (Theta)", "live 30 Nov 2026"), ("Tests and reports", "live today"), ("Follow-up emails", "live today")]),
      ("Later", "ps-next", "Year 2: Psynthesia · Year 3: abroad", [("AI drafts the report", "Rp300,000 a seat, 2028"), ("Malaysia or Singapore", "2029")]),
  ]), note="Proof: 205 AI reports in our trial (scoring 60 → 10 minutes, writing 7 → 1 hour). TemanJourney already buys our tests.")

slide("b2", "B", "B2. The sales agent answers every chat and hands the best leads to the admin", table(["", "Admin away (night, weekend)", "Admin online"], [
    ["Who replies", "The agent, at once", "The agent first, then the admin"],
    ["What it does", "Answers, asks a few questions, offers times and a payment link", "Picks serious leads and sums each one up"],
    ["Who closes", "The client books and pays, or the lead waits for the morning", "The admin"],
    ["Style", "The bureau's own style, learned from its past chats", "The admin corrects it"],
    ["Safety", "No counselling or diagnosis. Crisis: hotline at once, and a person is alerted", "The same"],
  ]))

slide("b2-value", "B", "B2. A bureau gets more clients for less than the cost of one admin", viz.compare_bars([
      ("Chats a month", X['inquiries'], X['inquiries'], num(X['inquiries']), num(X['inquiries'])),
      ("Answered in minutes", X['inquiries'] - x_after, X['inquiries'], num(X['inquiries'] - x_after), num(X['inquiries'])),
      ("Paying clients", x_now, x_agent, f"{num(x_now)} ({pct(X['conversion_now'])})", f"{num(x_agent)} ({pct(X['conversion_agent'], 1)})"),
  ]) + f"""
  <div class="two">
    <div class="card"><div class="eyebrow">The bureau gets</div><p class="lead"><b>{num(x_extra)} more clients</b> a month = <b>{rp(x_extra_rev)}</b> more sales {EST}</p></div>
    <div class="card"><div class="eyebrow">The bureau pays</div><p class="lead"><b>{rp(lp_cost, 2)}</b> (Starter plan + payment fees) = about <b>{lp_return:.0f}×</b> back. The top plan never costs more than one admin ({rpk(top_plan)}).</p></div>
  </div>""", note=f"Example: {num(X['inquiries'])} chats a month, {SESSIONS_PER_CLIENT} paid sessions per client. Easy to start (set up in one day), easy to run (all in one place), tests included (no fee to other test sites, which charge Rp{TEST_FEE_LOW:,} to Rp{TEST_FEE_HIGH:,} per person).")

slide("b3", "B", "B3. Only CoreItera sells for a bureau 24 hours, in its own style, into its real calendar", pos, note="Placement is our own judgement from public product pages (October 2026).")
slide("b3-matrix", "B", "B3. Feature by feature, the alternatives each miss a step of the loop", table(["", "CoreItera", "Admin on WhatsApp", "Generic chatbots", "Global software", "Counselling apps"], [
    ["Replies 24 hours", yes, no, yes, no, "Only for the app"],
    ["Bureau's own style", yes, yes, "Set by hand", no, no],
    ["Books real free times", yes, "By hand", no, yes, yes],
    ["Asks for feedback and reviews", yes, "If someone remembers", no, "Partly", "For the app"],
    ["Client belongs to the bureau", yes, yes, yes, yes, no],
  ]), note="Only CoreItera has every row: the loop needs all of them together.")

slide("b3-moat", "B", "B3. Each new bureau makes the agent better, which brings the next bureau", viz.flywheel([("More bureaus", "join for clients"), ("More chats", "answered 24 h"), ("Smarter agent", "every correction"), ("More clients", "bureaus earn more"), ("Referrals", "bureaus bring bureaus")], ("More people helped", "per psychologist")),
      note=f"Hard to copy: each bureau's style lives in its agent; its clients and bookings live in CoreItera; TemanJourney, HIMPSI and universities are trust channels; we cost {rp(I.BASE_RUN_COST)} a month to run, so we can price low.")

# ------------------------------------------------------------------ C. how
divider("part-c", "C", "C", "A safe plan and deal", ["Year 1 bureaus, year 2 Psynthesia, year 3 abroad", "Cheap to run; steady growth", "Rp50M in total for 20%"])

slide("c1", "C", "C1. Roadmap: bureaus in year 1, 700 bureaus in year 2, abroad in year 3", road,
      note=f"Why {num(B28)}: Vibefam (Singapore) runs 700+ gyms and studios on one system; bureaus need the same things. We build up step by step in year 1 ({num(B26)}, then +{num(Q_NEW[0])}, +{num(Q_NEW[1])}, +{num(Q_NEW[2])}, +{num(Q_NEW[3])} a quarter).")

slide("c1-build", "C", "C1. Three engineers ship one main build at a time", eng,
      note="Proposed owners. Psynthesia work starts in September 2027; it is sold from January 2028.")

slide("c1-team", "C", "C1. We hire only when profit can pay the first salary", team_viz, note="Sales first, then support, then quality and engineering. Galih and Riskia are the deal makers until the first sales hire.")
slide("c1-team-roles", "C", "C1. Who we hire, when, and what they cost", table(["Role", "Mar 2027", "Jun", "Sep", "Dec 2027", "Dec 2028", "Dec 2031", "Pay and job"], team_rows, right=(1, 2, 3, 4, 5, 6)),
      note="Galih and Riskia are the deal makers until the first sales hire. Galih: strategy and money. Adit: sales agent, scheduling, chat. Zaki: screens, set-up, writing style, Psynthesia. Riskia: runs TemanJourney and opens doors.")

slide("c1-risks", "C", "C1. We know what could stop us, and what we do then", table(["Risk", "Early sign", "What we do"], [
    ["Agent says something clinical", "Chats about symptoms or crisis", "No counselling; crisis goes to the hotline and a person."],
    ["Meta limits messaging", "A warning on WhatsApp or Instagram", "Official Meta tools only; answers about the bureau's services only."],
    ["Agent does not add clients", "No rise at TemanJourney in February", "Stop. No growth spending until it works."],
    ["Bureaus leave after a few months", "Churn not below 5% a month by December 2027", "Set-up call in week 1; monthly extra-clients report."],
    [f"{num(B28)} in year 2 is too ambitious", f"Under {num(Y1)} bureaus paying by December 2027", "Slow hiring and sign-ups; raise the optional round only with traction."],
    ["Family and company mix", "Decisions only in conversation", "Write everything down; both companies approve."],
  ]))

slide("c2", "C", f"C2. We are cheap to run and grow on our own income, but cash is thin until {month_name('2027-04').split()[0]} 2027", f"""
  <div class="tiles">
    <div class="tile"><b>{rp(I.BASE_RUN_COST)}</b><span>a month to run today; no salaries</span></div>
    <div class="tile"><b>{starter_needed} bureaus</b><span>on the Rp1M plan cover it</span></div>
    <div class="tile"><b>{month_name(first_profit)}</b><span>first month with a profit</span></div>
    <div class="tile"><b>{rp(low27)}</b><span>lowest cash in 2027 ({month_name(low27_m)}): under 2 months of costs</span></div>
  </div>""" + viz.quarter_chart(plan['quarters']),
      note=f"New money = only the Rp30M from this deal (TemanJourney Rp20M, Riskia Rp10M) in Q4 2026. No other outside money is in this plan. The risk: if no bureau pays before April 2027, cash falls to {rp(late['cash'])} in {month_name(late['month'])}; then we slow sign-ups and spending.")

slide("c2-cashq", "C", "C2. Cash by quarter: income catches up with costs from mid-2027", table(["Quarter", "Revenue", "Costs", "New money", "Net cash", "Cash at end"], q_rows, right=(1, 2, 3, 4, 5)),
      note="New money = only the Rp30M from this deal. Costs include the cost of winning each new bureau, paid before it pays us.")

slide("c2-revenue", "C", "C2. Revenue grows from the bureau tool first, then Psynthesia seats", viz.revenue_chart(Y),
      note=f"2027 income comes only from the bureau tool. Psynthesia adds income from 2028 and brings {pct(Y[2031]['copilot'] / Y[2031]['revenue'])} of revenue by 2031. With half the bureaus we still make a profit every year from {slow_profit_year}. Income from abroad is not counted.")

slide("c2-churn", "C", "C2. Churn starts at 8% because we are new, and falls to 2% by year 2", churn_viz,
      note="Industry: small-business software loses 3% to 7% of customers a month (Vena 2025; Focus Digital 6.4%). Most churn happens in the first 90 days.")
slide("c2-churn-count", "C", "C2. Churn means we must keep signing new bureaus", viz.hbars([
      ("Signed in 2027", SIGNED27, "c-cost", f"{num(SIGNED27)} bureaus"),
      ("Left during 2027", LOST27, "c-ochre", f"{num(LOST27)} bureaus"),
      ("Still paying in December 2027", Y1, "c-rev", f"{num(Y1)} bureaus"),
      ("New bureaus needed in 2028 to reach 700", Y[2028]['new'], "c-cost", f"{num(Y[2028]['new'])} (because {num(Y[2028]['churned'])} leave)"),
  ]), note=f"Each bureau that leaves must be replaced: in 2028 we need {num(Y[2028]['new'])} new bureaus to end the year with {num(B28)}.")

slide("c2-churn-cases", "C", "C2. Getting churn down saves hundreds of bureaus of sales work", table(["Churn path", "Jan 2027 → Dec 2027 → Dec 2028", "Paying Dec 2027", "New needed in 2028", "Profit 2028", "Profit 2031"], churn_rows, right=(2, 3, 4, 5)),
      note="How we keep churn low: a set-up call in week 1, a monthly extra-clients report, and the first month free so only serious bureaus pay.")

slide("c2-angel", "C", f"C2. An optional {rp(rnd['amount'], 0)} round, only if we have traction by January", table(["", "Main plan (own income)", "With the optional round"], [
    ["When", "No round", f"Traction by January 2027; closed by {rnd_when} at the earliest"],
    ["Outside money", "Rp30M (this deal)", f"Rp30M + {rp(rnd['amount'], 0)} from angel investors"],
    ["Cash at the end of 2027", rp(Y[2027]['cash_end']), rp(WY[2027]['cash_end'])],
    ["What the money is for", "None needed to stay above zero", "Scale: hire sales and support earlier, push faster to 700"],
    ["Ownership of Riskia's side", "20%", f"{pct(stake_after, 0)} unless they join the round"],
  ], right=(1, 2)), note=f"Traction = TemanJourney's closing rate rising and at least {num(B27[(2027, 3)])} paying bureaus by March 2027. Angel investors through warm introductions or cold email. The round would cost about {pct(rnd['dilution'])} of CoreItera {EST}. We do not count on it.")

slide("c3", "C", "C3. Riskia's side pays Rp50M in total for 20%, and the 15% promise is kept", viz.cap_bars([
      ("Today", [(0.90, "seg-o", "Current owners 90%"), (0.10, "seg-r", "10%")]),
      ("Step 1: 15% kept", [(0.85, "seg-o", "Current owners 85%"), (0.15, "seg-r", "15%")]),
      ("Step 2: Rp30M in", [(0.80, "seg-o", "Current owners 80%"), (0.1608, "seg-r", "16.1%"), (0.0392, "seg-t", "")]),
  ]) + table(["Who", "Paid before", "Paid now", "Total", "Owns after"], [
    ["Galih (with the other founders)", "Rp50M", "None", "Rp50M", "80.0%"],
    ["Riskia", "Rp20M", "Rp10M", "Rp30M", "16.1%"],
    ["TemanJourney", "None", "Rp20M", "Rp20M", "3.9%"],
    ["<b>Total</b>", "<b>Rp70M</b>", "<b>Rp30M</b>", "<b>Rp100M</b>", "<b>100%</b>"],
  ], right=(1, 2, 3, 4)), note="Price: Rp480M before the new money, Rp510M after: 2.5× the price of Riskia's first 10%; her 16.1% is worth about Rp82M on paper. Founders own more for the same Rp50M because their money came first and they built the platform, mostly unpaid.")

slide("c3-gets", "C", "C3. The Rp30M goes straight into making TemanJourney succeed", viz.stacked([("Running costs to February 2027", running_part, "seg-r"), ("Notary and share registration", I.DEAL_ONE_OFFS['notary and share registration'], "seg-o"), ("Launch", I.DEAL_ONE_OFFS['launch costs'], "seg-t")], f"Total {rp(I.DEAL_CASH)}") + table(["Riskia's side gets", "Plan", "Slow case"], [
    ["Ownership", f"20%; {pct(stake_after, 0)} only if the optional round happens and they do not join", "The same"],
    ["TemanJourney's software", "Free for 12 months, then cost price", "The same"],
    [f"Dividends 2030 and 2031 ({pct(runner.DIVIDEND_SHARE)} of profit) {EST}", rp(sum(div_plan.values()), 0), rp(sum(div_slow.values()), 0)],
    ["Say", "Monthly update; first look at new tools; right to invest again", "The same"],
  ]), note="Young plans often miss. These numbers are estimates, not promises.")

slide("ask", "Ask", "What we ask Riskia and TemanJourney", f"""
  <ol class="ask-list">
    <li><b>Invest Rp30M</b> (TemanJourney Rp20M, Riskia Rp10M) for 20%: Riskia 16.1%, TemanJourney 3.9%.</li>
    <li><b>Run the sales agent</b> on TemanJourney's WhatsApp from February 2027, and share the chat numbers.</li>
    <li><b>If we have traction by January,</b> help us reach angel investors for an optional {rp(rnd['amount'], 0)} round by {rnd_when}, with the right to join it.</li>
    <li><b>Name one person</b> at TemanJourney to own the launch.</li>
    <li><b>Introduce 2 Jakarta bureaus</b> by January 2027.</li>
    <li><b>Sign the term sheet</b> by Friday 30 October 2026.</li>
  </ol>""", kind="ask")

# ------------------------------------------------------------------ appendix (one table per slide)
divider("appendix", "Appendix", "+", "Appendix: the detail behind the numbers", ["Profit and loss, costs, slow case", "Prices, tests, launch plan", "Assumptions and what to confirm"])
slide("x-pnl", "Appendix", "Five-year profit and loss (estimates)", table(["Year", "Bureaus", "Seats", "Revenue", "Costs", "Tax", "Profit", "Margin", "Cash (Dec)"], [[r[0], r[1], r[2], r[6], r[7], r[8], r[9], r[10], r[11]] for r in yr_rows], right=tuple(range(1, 9))),
      note="Bureaus and seats in December. Tax: company income tax under Indonesian rules (see the tax slide); a tax consultant must confirm. TemanJourney is not counted.")
slide("x-costs", "Appendix", "Main costs by year", table(["Cost"] + [str(y) for y in Y], cost_rows, right=tuple(range(1, 6)), cls="compact"),
      note=f"2027: founders and Zaki get Rp10M a month each, only in a month that makes a profit (from {month_name(plan['founders_paid_from'])}). From {month_name(plan['full_pay_from'])}: Galih {rp(fp['Galih'], 0)}, Riskia {rp(fp['Riskia'], 0)} and Adit {rp(fp['Adit'], 0)} a year.")
TP_ = plan["tax_plan"]
tax_rows = [[str(y), rp(Y[y]["revenue"]), rp(Y[y]["before_tax"]), TP_[y]["rule"], pct(TP_[y]["rate"], 1) + (" of turnover" if TP_[y]["rule"] == "0.5% final" else " of profit"), f"<b>{rp(TP_[y]['tax'])}</b>"] for y in Y]
slide("x-tax", "Appendix", "Tax: 0.5% of turnover in 2027, then 22% of profit with the small-company cut", table(["Year", "Turnover", "Profit before tax", "Rule", "Average rate", "Tax"], tax_rows, right=(1, 2, 4, 5)),
      note="Checked against DJP and tax-news sources (DDTC, IKPI), October 2026. Not tax advice: a tax consultant should confirm.")
slide("x-tax-rules", "Appendix", "Tax rules behind the numbers", table(["Rule", "What it means for PT CoreItera"], [
    ["0.5% of turnover (PP 55/2022)", f"Only for turnover up to Rp4.8 billion a year, and for a PT only for its first 3 tax years. Since PP 20/2026 (April 2026), a PT can no longer start using it, but PTs already using it keep it until their 3 years end. PT CoreItera was registered on {I.TAX_REGISTERED} and uses it today, so it keeps 0.5% for tax years 2025, 2026 and 2027."],
    ["22% of profit (UU HPP)", "The normal company tax, on profit after costs, not on turnover."],
    ["Pasal 31E cut", "If turnover is up to Rp50 billion, the part of profit that matches the first Rp4.8 billion of turnover pays half (11%). So the average rate rises from 11% as turnover grows."],
    ["Losses", "A loss year pays no company tax, and the loss lowers taxable profit for the next 5 years."],
    ["PTKP", "The personal allowance is for people, not for a PT. It lowers each founder's own salary tax (PPh 21), which is already inside their gross pay."],
    ["Cash timing", "Each month we pay last year's tax / 12 (PPh 25); the rest of last year's tax is paid in April (PPh 29)."],
  ]), note="Checked against DJP and tax-news sources (DDTC, IKPI), October 2026. Not tax advice: a tax consultant should confirm, especially that 2027 is the last 0.5% year. VAT (PPN) is charged on top of our prices once turnover passes Rp4.8 billion, so it is not a cost in this model.")
slide("x-slow", "Appendix", "Slow case: half the bureaus", table(["Year", "Bureaus", "Revenue", "Profit", "Margin", "Cash (Dec)"], slow_rows, right=(1, 2, 3, 4, 5)),
      note=f"Lowest cash in the slow case: {rp(slow_low['cash'])} in {month_name(slow_low['month'])}.")
slide("x-prices", "Appendix", "Prices", table(["Plan", "Psychologists", "Price a month", "Per psychologist"], [
    ["Starter", "1 to 3", rpk(I.PLAN_TIERS[0][3]), "Rp333,000 or more"],
    ["Growth", "4 to 10", rpk(I.PLAN_TIERS[1][3]), f"{rpk(I.PLAN_TIERS[1][3] / 10)} to {rpk(I.PLAN_TIERS[1][3] / 4)}"],
    ["Pro", "11 or more", rpk(I.PLAN_TIERS[2][3]) + " (one admin at Jakarta minimum wage 2026)", f"{rpk(I.PLAN_TIERS[2][3] / 11)} or less"],
  ], right=(2, 3)), note="Every plan has the sales agent, scheduling, booking, chat counselling, tests and feedback requests. Plus Rp5,000 per payment (Midtrans fee at cost). Psynthesia: Rp300,000 a seat from 2028. First month free.")
slide("x-tests", "Appendix", "Five tests this business must pass", table(["Test", "Evidence", "Confidence"], tests))
slide("x-launch", "Appendix", "Launch at TemanJourney on 30 November 2026", table(["Milestone", "Done by", "Owner"], [
    ["Riskia reads this proposal; numbers agreed", "Fri 23 Oct", "Galih, Riskia"],
    ["Term sheet signed by both companies", "Fri 30 Oct", "Galih, Riskia"],
    ["Shareholder approval in both companies (RUPS); notary booked", "Fri 6 Nov", "Galih, Riskia"],
    ["Notary deed signed; Rp30M paid in", "Fri 13 Nov", "Riskia, TemanJourney"],
    ["At least 20 psychologist profiles and licences (SIPP) entered", "Fri 13 Nov", "TemanJourney admin"],
    ["Prices and weekly time blocks entered", "Fri 20 Nov", "TemanJourney admin, Adit"],
    ["Release practised on a copy of the live database", "Fri 20 Nov", "Galih, Adit, Zaki"],
    ["Test week: real bookings, payments and chats", "Fri 27 Nov", "All"],
    ["<b>Theta live for TemanJourney clients</b>", "<b>Mon 30 Nov</b>", "All"],
  ], right=(1,)))
slide("x-assume", "Appendix", "Main assumptions (1 of 2)", table(["Assumption", "Value", "Status"], assump[:8], cls="compact"))
slide("x-assume-2", "Appendix", "Main assumptions (2 of 2)", table(["Assumption", "Value", "Status"], assump[8:], cls="compact"),
      note="All in model/inputs.py; rebuild with model/build_page.py.")
slide("x-confirm", "Appendix", "To confirm before signing", table(["Number", "Used here", "Why it matters"], [
    ["CoreItera's share count (from the deed)", "1,000,000 as an example", "Sets the exact number of new shares"],
    ["How the earlier Rp50M and Rp20M are recorded", "Galih and Riskia", "The deed must match the cash table"],
    ["Number of private bureaus and practices in Indonesia", f"About {num(round(bureaus_market, -2))}", "HIMPSI does not publish it"],
    ["Optional round: size, price and timing", f"{rp(rnd['amount'], 0)} for {pct(rnd['dilution'])}, {rnd_when} at the earliest", "Only with traction; the price is an estimate"],
    ["Closing rate with the agent, and after-hours share", f"{pct(X['conversion_agent'], 1)}; {pct(X['after_hours_share'])}", "The lead product's whole promise"],
    ["Other shareholders of TemanJourney", "Not known", "They must approve the Rp20M"],
  ]))
slide("x-alt", "Appendix", "The other way to price the deal", table(["", "Recommended (15% kept, then cash)", "Cash only"], [
    ["Riskia at the end", "16.1%", "12.6%"], ["TemanJourney at the end", "3.9%", "7.4%"], ["Together", "20.0%", "20.0%"],
    ["Value of CoreItera before the new money", "Rp480M", "Rp240M"], ["Is the 15% promise kept?", "Yes", "No"],
  ], right=(1, 2)), note="The recommended deal keeps the promise and values CoreItera at twice the price.")


# ------------------------------------------------------------------ render
TRACK = [("Start", "Summary"), ("A", "A. Need"), ("B", "B. Product"), ("C", "C. Plan and deal"), ("Ask", "Ask")]
AGENDA = [("A", "The need is big", ["Too few psychologists", "Founders and admins stuck in a loop", "Indonesia is the market to start"]),
          ("B", "One product first: the bureau tool", ["The MVP fixes each step of the loop", "More clients for less than one admin", "Only we sell 24 hours in the bureau's own style"]),
          ("C", "A safe plan and deal", ["Year 1 bureaus, year 2 Psynthesia, year 3 abroad", "Cheap to run; own income", "Rp50M in total for 20%"]),
          ("Ask", "What we ask", [])]


def tracker(sec):
    if sec == "Appendix":
        return '<div class="trk"><span class="on">Appendix</span></div>'
    return '<div class="trk">' + "".join(f'<span class="{"on" if k == sec else ""}">{lab}</span>' for k, lab in TRACK) + "</div>"


def render(sl, n, total):
    foot = f'<div class="sfoot"><span class="src">{sl.get("note", "")}</span><span class="pg">CoreItera · {n}</span></div>'
    if sl["kind"] == "divider":
        if sl["sec"] == "Appendix":
            rows = f'<div class="ag on"><b>+</b><span>{sl["title"]}</span><ul>' + "".join(f"<li>{p}</li>" for p in sl["points"]) + "</ul></div>"
        else:
            rows = "".join(f'<div class="ag {"on" if k == sl["sec"] else ""}"><b>{k if k != "Ask" else "→"}</b><span>{t}</span>'
                           + (("<ul>" + "".join(f"<li>{p}</li>" for p in pts) + "</ul>") if k == sl["sec"] and pts else "") + "</div>"
                           for k, t, pts in AGENDA)
        return (f'<section class="slide divider" id="{sl["id"]}" data-sec="{sl["sec"]}"><div class="eyebrow">Agenda</div>'
                f'<div class="agenda">{rows}</div>{foot}</section>')
    lead = f'<p class="slead">{sl["lead"]}</p>' if sl.get("lead") else ""
    cls = "slide " + {"cover": "cover", "ask": "askslide"}.get(sl["kind"], "")
    if sl["kind"] == "cover":
        return (f'<section class="{cls}" id="{sl["id"]}" data-sec="{sl["sec"]}"><div class="eyebrow">CoreItera</div>'
                f'<h1 class="stitle">{sl["title"]}</h1>{lead}<div class="sbody">{sl["body"]}</div></section>')
    return (f'<section class="{cls}" id="{sl["id"]}" data-sec="{sl["sec"]}">{tracker(sl["sec"])}'
            f'<h2 class="stitle">{sl["title"]}</h2>{lead}<div class="sbody">{sl["body"]}</div>'
            + (f'<p class="snote">{sl["note"]}</p>' if sl.get("note") else "") + f'{foot}</section>')


S = [render(sl, i + 1, len(SL)) for i, sl in enumerate(SL)]

# deck export: structure for the later slide deck (titles, notes, kind, section, and the body HTML)
import json, re as _re
DECK_JSON = os.path.join(HERE, "..", "output", "slides.json")
with open(DECK_JSON, "w") as _f:
    json.dump([{"n": i + 1, "id": sl["id"], "section": sl["sec"], "kind": sl["kind"], "title": sl["title"],
                "lead": sl.get("lead", ""), "points": sl.get("points", []),
                "note": _re.sub(r"<[^>]+>", "", sl.get("note", "")), "body_html": sl.get("body", "")} for i, sl in enumerate(SL)],
              _f, indent=1, ensure_ascii=False)

footer = """
<footer>
  Sources: CoreItera trial data (205 reports rated over 6 months); TemanJourney 2027 financial model (October 2026); WHO (2025, 2026), Mental Health Atlas 2024; UN World Population Prospects 2024; GBD 2021 ASEAN Mental Disorders Collaborators (2025); HIMPSI (2026, 2021); Rettew et al. (2009); Augment proposal (IS5007, October 2026) for the typical week, assessment hours, the 6 to 15 a week calculation and Law 23/2022; Kemenkes, Survei Kesehatan Indonesia 2023; SimplePractice prices from third-party guides (2026); Riliv from Fortune Indonesia (2023); Klinik Pintar from Dealroom and e27; Vibefam from its June 2026 funding release and Product Hunt (700+ locations, 26 countries); Jakarta minimum wage 2026 (Rp5,729,876) from Bisnis.com; KantorKu test prices via the Augment proposal. US$1 = Rp16,500. A notary and a tax consultant must confirm the legal and tax steps.
</footer>"""


FIT_JS = """
(function(){
  function fit(){
    var bar = document.querySelector('.pbar');
    var w = document.documentElement.clientWidth - 32;
    var h = bar ? window.innerHeight - bar.offsetHeight - 20 : Infinity;
    var z = Math.max(0.2, Math.min(w / 1280, h / 720));
    document.querySelectorAll('.slide').forEach(function(s){ s.style.zoom = z; });
  }
  var q = location.search;
  var only = (q.match(/only=([A-Za-z0-9_-]+)/) || [])[1];
  if (only) {                       // one slide at 1280 x 720, for screenshots
    var bar = document.querySelector('.pbar'); if (bar) bar.remove();
    document.body.style.padding = '0'; document.body.style.margin = '0';
    document.querySelectorAll('.slide').forEach(function(s){ if (s.id !== only) s.remove(); else { s.style.zoom = 1; s.style.border = '0'; s.style.boxShadow = 'none'; } });
    var f = document.querySelector('footer'); if (f) f.remove();
    document.querySelector('main').style.gap = '0';
    return;
  }
  fit(); window.addEventListener('resize', fit);
  if (q.indexOf('check') >= 0) {    // list slides whose content does not fit
    var bad = [];
    document.querySelectorAll('.slide').forEach(function(s, i){
      var b = s.querySelector('.sbody'); var over = b ? b.scrollHeight - b.clientHeight : 0;
      var wide = Array.prototype.some.call(s.querySelectorAll('.tablewrap,.chart'), function(e){ return e.scrollWidth > e.clientWidth + 2; });
      if (over > 2 || wide) bad.push({n: i + 1, id: s.id, over: over, wide: wide});
    });
    var pre = document.createElement('pre'); pre.id = 'fitcheck'; pre.textContent = JSON.stringify(bad); document.body.appendChild(pre);
  }
})();
"""
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
<script>{FIT_JS}</script>
"""

with open(OUT_HTML, "w") as f:
    f.write(html)

import present
OUT_PRESENT = os.path.join(HERE, "..", "proposal", "theta-partnership-proposal-present.html")
with open(OUT_PRESENT, "w") as f:
    f.write(present.standalone("Theta Partnership Proposal", css, f"<main>\n{''.join(S)}\n{footer}\n</main><script>{FIT_JS}</script>"))
print("wrote", os.path.normpath(OUT_PRESENT))
runner.main()
print("wrote", os.path.normpath(OUT_HTML))
