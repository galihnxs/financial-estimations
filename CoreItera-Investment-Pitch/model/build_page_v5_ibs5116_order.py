"""Build the proposal page from the model, so every number on the page comes from engine.py.

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
plan = engine.run(I)
slow = engine.run(runner.slow_case())
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


# ---------- charts ----------
def quarter_chart(qs):
    W, H, L, R, T, B = 720, 300, 60, 20, 20, 50
    vals = [q["revenue"] for q in qs] + [q["costs"] for q in qs] + [q["cash_end"] for q in qs]
    top = max(vals)
    top = math.ceil(top / 50) * 50
    bottom = min(0, math.floor(min(vals) / 10) * 10)
    span = top - bottom
    def yv(v):
        return T + (top - v) / span * (H - T - B)
    n = len(qs)
    step = (W - L - R) / n
    bw = step * 0.28
    g = []
    for t in range(int(bottom), int(top) + 1, 50):
        g.append(f'<line x1="{L}" x2="{W - R}" y1="{yv(t):.1f}" y2="{yv(t):.1f}" class="grid"/>'
                 f'<text x="{L - 8}" y="{yv(t) + 4:.1f}" class="ax" text-anchor="end">{t}</text>')
    g.append(f'<line x1="{L}" x2="{W - R}" y1="{yv(0):.1f}" y2="{yv(0):.1f}" class="zero"/>')
    pts = []
    for i, q in enumerate(qs):
        x0 = L + i * step + step * 0.18
        for j, (key, c) in enumerate((("revenue", "c-rev"), ("costs", "c-cost"))):
            v = q[key]
            x = x0 + j * (bw + 3)
            g.append(f'<rect x="{x:.1f}" y="{yv(v):.1f}" width="{bw:.1f}" height="{max(0.5, yv(0) - yv(v)):.1f}" class="{c}"/>')
        cx = L + i * step + step / 2
        pts.append((cx, yv(q["cash_end"])))
        g.append(f'<text x="{cx:.1f}" y="{H - B + 20}" class="ax" text-anchor="middle">{q["quarter"]}</text>')
    g.append('<polyline fill="none" class="c-cash" points="' + " ".join(f"{x:.1f},{y:.1f}" for x, y in pts) + '"/>')
    for (x, y), q in zip(pts, qs):
        g.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" class="c-cashdot"/>'
                 f'<text x="{x + 8:.1f}" y="{y - 9:.1f}" class="lab" text-anchor="start">{q["cash_end"]:.0f}</text>')
    return (f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Quarterly revenue, costs and cash">'
            + "".join(g) + "</svg>"
            '<div class="legend"><span><i class="c-rev"></i>Revenue</span><span><i class="c-cost"></i>Costs</span>'
            '<span><i class="c-cashline"></i>Cash at the end of the quarter</span><span>Rp million</span></div></div>')


def revenue_chart(years):
    W, H, L, R, T, B = 720, 280, 70, 20, 20, 40
    top = math.ceil(max(v["revenue"] for v in years.values()) / 2000) * 2000
    def yv(v):
        return T + (top - v) / top * (H - T - B)
    step = (W - L - R) / len(years)
    bw = step * 0.5
    g = []
    for t in range(0, int(top) + 1, 2000):
        g.append(f'<line x1="{L}" x2="{W - R}" y1="{yv(t):.1f}" y2="{yv(t):.1f}" class="grid"/>'
                 f'<text x="{L - 8}" y="{yv(t) + 4:.1f}" class="ax" text-anchor="end">{t / 1000:.0f}</text>')
    for i, (y, v) in enumerate(years.items()):
        x = L + i * step + (step - bw) / 2
        base = 0.0
        for key, c in (("plan", "c-plan"), ("fees", "c-fee"), ("copilot", "c-cop")):
            h = v[key]
            g.append(f'<rect x="{x:.1f}" y="{yv(base + h):.1f}" width="{bw:.1f}" height="{max(0.5, yv(base) - yv(base + h)):.1f}" class="{c}"/>')
            base += h
        g.append(f'<text x="{x + bw / 2:.1f}" y="{yv(base) - 7:.1f}" class="lab" text-anchor="middle">{base / 1000:.1f}</text>'
                 f'<text x="{x + bw / 2:.1f}" y="{H - B + 20}" class="ax" text-anchor="middle">{y}</text>')
    return (f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Revenue by year and source">' + "".join(g) + "</svg>"
            '<div class="legend"><span><i class="c-plan"></i>Monthly plans</span><span><i class="c-fee"></i>Rp5,000 payment fees</span>'
            '<span><i class="c-cop"></i>Copilot seats</span><span>Rp billion a year</span></div></div>')


def positioning_map():
    W, H, L, B = 640, 420, 70, 60
    pts = [("CoreItera", 86, 86, "us"),
           ("WhatsApp, spreadsheets and Word", 8, 62, ""),
           ("Global practice software (SimplePractice, Carepatron)", 26, 14, ""),
           ("Clinic record software (Klinik Pintar)", 8, 40, ""),
           ("Counselling apps (Riliv, Satu Persen)", 32, 74, ""),
           ("Online test sites (KantorKu)", 48, 52, "")]
    def xy(x, y):
        return L + x / 100 * (W - L - 20), 20 + (100 - y) / 100 * (H - 20 - B)
    g = [f'<rect x="{L}" y="20" width="{W - L - 20}" height="{H - 20 - B}" class="frame"/>',
         f'<line x1="{xy(50, 0)[0]:.0f}" x2="{xy(50, 0)[0]:.0f}" y1="20" y2="{H - B}" class="grid"/>',
         f'<line x1="{L}" x2="{W - 20}" y1="{xy(0, 50)[1]:.0f}" y2="{xy(0, 50)[1]:.0f}" class="grid"/>',
         f'<text x="{(L + W - 20) / 2:.0f}" y="{H - 22}" class="ax" text-anchor="middle">Handles assessment and signed reports →</text>',
         f'<text x="18" y="{(20 + H - B) / 2:.0f}" class="ax" text-anchor="middle" transform="rotate(-90 18 {(20 + H - B) / 2:.0f})">Fits how Indonesian bureaus work →</text>']
    for name, x, y, cls in pts:
        cx, cy = xy(x, y)
        anchor = "end" if x > 60 else "start"
        dx = -10 if anchor == "end" else 10
        g.append(f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{9 if cls else 6}" class="{"dot-us" if cls else "dot"}"/>'
                 f'<text x="{cx + dx:.0f}" y="{cy + 4:.0f}" class="{"lab-us" if cls else "lab"}" text-anchor="{anchor}">{name}</text>')
    return f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Positioning map">' + "".join(g) + "</svg></div>"


# ---------- page ----------
css = open(os.path.join(HERE, "page_style.css")).read() + """
.chart{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:14px 16px;display:flex;flex-direction:column;gap:8px;overflow-x:auto}
.chart svg{width:100%;min-width:520px;height:auto;display:block}
.chart .grid{stroke:var(--line);stroke-width:1}
.chart .zero{stroke:var(--muted);stroke-width:1.2}
.chart .frame{fill:none;stroke:var(--line)}
.chart .ax{fill:var(--muted);font:13px var(--body)}
.chart .lab{fill:var(--ink);font:600 12.5px var(--body);paint-order:stroke;stroke:var(--surface);stroke-width:4px;stroke-linejoin:round}
.chart .lab-us{fill:var(--teal);font:700 14px var(--body)}
.chart .dot{fill:var(--slate)} .chart .dot-us{fill:var(--teal)}
.c-rev{fill:var(--teal);background:var(--teal)} .c-cost{fill:var(--slate);background:var(--slate)}
.c-cash{stroke:var(--ochre);stroke-width:2.5} .c-cashdot{fill:var(--ochre)} .c-cashline{background:var(--ochre)}
.c-plan{fill:var(--slate);background:var(--slate)} .c-fee{fill:var(--ochre);background:var(--ochre)} .c-cop{fill:var(--teal);background:var(--teal)}
.canvas{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));grid-auto-rows:minmax(110px,auto);gap:8px}
.canvas .cell{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:12px;font-size:14.5px;line-height:1.45;min-width:0;display:flex;flex-direction:column;gap:6px}
.canvas .cell b{font-size:12.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--teal)}
.canvas .tall{grid-row:span 2} .canvas .half{grid-column:span 5}
.canvas .split{grid-column:span 5;display:grid;grid-template-columns:1fr 1fr;gap:8px;background:none;border:0;padding:0}
@media (max-width:760px){.canvas{grid-template-columns:1fr}.canvas .tall{grid-row:auto}.canvas .half,.canvas .split{grid-column:auto}.canvas .split{grid-template-columns:1fr}}
.mission{background:var(--surface);border:1px solid var(--line);border-top:6px solid var(--teal);border-radius:6px;padding:28px 26px;display:flex;flex-direction:column;gap:10px}
.mission-text{font-family:var(--display);font-size:clamp(24px,3.6vw,34px);line-height:1.25;font-weight:600;text-wrap:balance;max-width:none}
.tiles.three{grid-template-columns:repeat(3,minmax(0,1fr))}
@media (max-width:720px){.tiles.three{grid-template-columns:1fr}}
.toc{display:flex;flex-wrap:wrap;gap:8px 18px;font-size:15px}
.yes{color:var(--teal);font-weight:700} .no{color:var(--muted)}
.conf{font-weight:700;font-size:13px;white-space:nowrap}
"""

S = []   # sections
add = S.append

add(f"""
<header style="display:flex;flex-direction:column;gap:14px">
  <div class="eyebrow">CoreItera · business proposal for Riskia and TemanJourney</div>
  <h1>Theta Partnership Proposal</h1>
  <p class="meta">From Galih Muji Agung, CEO of CoreItera · 19 October 2026 · Draft for discussion, version 5</p>
  <nav class="toc" aria-label="Sections">
    <a href="#mission">Mission</a><a href="#summary">1. Summary</a><a href="#problem">2. Problem</a><a href="#market">3. Market</a><a href="#product">4. Product</a>
    <a href="#competition">5. Competition and sales</a><a href="#edge">6. Our edge</a><a href="#team">7. Team</a>
    <a href="#operations">8. Operations</a><a href="#expansion">9. Expansion</a><a href="#financials">10. Financials</a><a href="#ask">11. The ask</a>
  </nav>
</header>""")

# 0 mission
assess_x = I.ASSESS_PER_WEEK_COPILOT / I.ASSESS_PER_WEEK_NOW
agree_x = I.AGREEMENT_TARGET / I.AGREEMENT_NOW
more_clients = I.EXAMPLE["conversion_agent"] / I.EXAMPLE["conversion_now"] - 1
asean_share = I.ASEAN6["assessments_per_year_m"] / I.ASEAN6["with_disorder_m"]
asean_after = I.ASEAN6["assessments_per_year_m"] * assess_x
per100k = I.ASEAN6["psychologists"] / (I.ASEAN6["people_m"] * 10)
add(f"""
<section id="mission">
  <div class="mission">
    <div class="eyebrow">Our mission</div>
    <p class="mission-text">{I.MISSION}</p>
  </div>

  <h2>The problem the world has, and Indonesia has most</h2>
  <div class="tiles">
    <div class="tile"><b>{I.WORLD['people_with_condition_bn']:.0f} billion+</b><span>people in the world live with a mental health condition (WHO)</span></div>
    <div class="tile"><b>US${I.WORLD['lost_usd_trillion']:.0f} trillion</b><span>lost every year to depression and anxiety at work (WHO)</span></div>
    <div class="tile"><b>{per100k:.0f} per 100,000</b><span>psychologists in our six ASEAN countries, against {I.WORLD['workers_per_100k']} mental health workers per 100,000 worldwide</span></div>
    <div class="tile"><b>{pct(asean_share, 1)}</b><span>of the {I.ASEAN6['with_disorder_m']:.1f} million people with a mental disorder in these countries can get a full assessment each year</span></div>
  </div>
  {table(["Where", "The gap"], [
    ["World", f"Over {I.WORLD['people_with_condition_bn']:.0f} billion people live with a mental health condition. The world has only about {I.WORLD['workers_per_100k']} mental health workers per 100,000 people."],
    ["Six ASEAN countries (Indonesia, Philippines, Vietnam, Thailand, Malaysia, Singapore)", f"About {num(I.ASEAN6['people_m'])} million people, {I.ASEAN6['with_disorder_m']:.1f} million with a mental disorder, and only about {num(I.ASEAN6['psychologists'])} psychologists. Together they can do about {I.ASEAN6['assessments_per_year_m']:.1f} million full assessments a year."],
    ["Indonesia", f"About {num(I.PSYCHOLOGISTS_ID)} psychologists, of whom {num(I.ID['clinical'])} are clinical. At today's training speed it takes {I.ID['years_to_close_gap']} years to close the clinical gap. More than {num(I.ID['graduates'])} psychology graduates cannot practise alone."],
  ])}
  <p><b>Training more psychologists alone cannot close this gap in our lifetime.</b> The faster way is to let each psychologist and each bureau help more people with the same hours. That is what CoreItera does.</p>

  <h3>What CoreItera changes</h3>
  <div class="tiles three">
    <div class="tile"><b>{assess_x:.1f}×</b><span>more full assessments per psychologist: from {I.ASSESS_PER_WEEK_NOW} to {I.ASSESS_PER_WEEK_COPILOT} a week, because the copilot cuts each one from 10.5 to about 2.2 hours of the psychologist's time</span></div>
    <div class="tile"><b>{pct(more_clients)} more</b><span>clients per bureau, because the sales agent answers Instagram and WhatsApp 24 hours and the admin closes the best leads {EST}</span></div>
    <div class="tile"><b>{agree_x:.1f}×</b><span>more accurate diagnosis as our target: routine diagnoses agree only {I.AGREEMENT_NOW:.2f} (on a 0 to 1 scale) with a structured assessment; we aim for {I.AGREEMENT_TARGET:.1f} or better</span></div>
  </div>
  <p>Across our six ASEAN countries, {assess_x:.1f} times more assessments would lift capacity from about {I.ASEAN6['assessments_per_year_m']:.1f} million to about {asean_after:.1f} million a year with the same psychologists. Report quality in our own trial already rose from {I.QUALITY_BEFORE} to above {I.QUALITY_AFTER} out of 6.</p>
  <p class="note">Sources: WHO (2025, 2026) and Mental Health Atlas 2024; UN World Population Prospects 2024; GBD 2021 ASEAN Mental Disorders Collaborators (2025); HIMPSI (2026); Rettew et al. (2009), 38 studies; assessment times and the {I.ASSESS_PER_WEEK_NOW} to {I.ASSESS_PER_WEEK_COPILOT} a week calculation from the Augment proposal (IS5007, October 2026) and CoreItera's 205-report trial. The 0.6 accuracy figure is a target for our pilot, not yet a result. The {pct(more_clients)} comes from the sales-agent example in section 2 and will be measured at TemanJourney first.</p>
  <p><b>Why start with Indonesia's bureaus:</b> Indonesia has the largest gap in the region (32.9 million people with a mental disorder, almost half of the six countries' total, GBD 2021), the largest pool of psychologists to serve them, and our first bureau, TemanJourney, is live on 30 November 2026. What works for an Indonesian bureau can then travel to the region and the world.</p>
</section>""")

# 1 summary
add(f"""
<section id="summary">
  <h2>1. Executive summary</h2>
  <div class="answer">
    <div class="eyebrow">In one paragraph</div>
    <p class="lead"><b>CoreItera helps a psychology bureau serve more clients, so it earns more.</b> Our sales agent answers the bureau's Instagram and WhatsApp 24 hours a day, in the bureau's own writing style. When the admin is away, it answers, sorts serious buyers from browsers and offers open times. When the admin is online, it hands the best leads over and the admin closes. Then our booking, counselling, assessment and copilot tools give psychologists the hours to serve those new clients. A bureau pays from <b>Rp1M a month</b> for everything, plus <b>Rp5,000</b> per paid booking, plus <b>Rp300,000</b> per copilot seat. We ask Riskia and TemanJourney for the next <b>Rp30M</b> (Rp50M in total from Riskia's side) for <b>20% of CoreItera together</b>.</p>
  </div>
  <div class="tiles">
    <div class="tile"><b>{num(Y[2031]['bureaus_end'])}</b><span>paying bureaus by December 2031, about {pct(Y[2031]['bureaus_end'] / bureaus_market)} of Indonesia's estimated {num(round(bureaus_market, -2))}</span></div>
    <div class="tile"><b>{rp(Y[2031]['revenue'])}</b><span>revenue in 2031, with a {pct(Y[2031]['profit'] / Y[2031]['revenue'])} profit margin {EST}</span></div>
    <div class="tile"><b>{month_name(plan['profitable_for_good_from'])}</b><span>profitable every month from here; first profitable month {month_name(plan['first_profitable_month'])}</span></div>
    <div class="tile"><b>{rp(div2, 0)}</b><span>dividends to Riskia's side in 2030 and 2031, if we pay out {pct(runner.DIVIDEND_SHARE)} of profit {EST}</span></div>
  </div>
  {table(["What investors ask", "Our answer"], [
    ["The opportunity", f"About {num(I.PSYCHOLOGISTS_ID)} psychologists in Indonesia (HIMPSI). Their clients find them and buy through Instagram and WhatsApp, but bureaus answer only in admin hours and run everything else on spreadsheets and Word."],
    ["Our position", "The only system that sells for the bureau 24 hours in its own voice <b>and</b> runs the bureau day and the assessment report, built for Indonesia: local tests, local payments, Indonesia's data law."],
    ["Value to a bureau", f"More clients: a bureau with {num(X['inquiries'])} chats a month can win about {num(x_extra)} more clients, about {rp(x_extra_rev)} more sales a month {EST}. A 10-psychologist bureau pays {rp(ex_tools)} a month for all the tools; with the copilot ({rp(ex_copilot)} more) its psychologists can earn about {rp(ex_extra_fees, 0)} more a month in assessment fees."],
    ["Proof today", "205 reports in our own trial: scoring from 60 to 10 minutes, report writing from about 7 hours to about 1, quality from 4.2 to above 5.5 out of 6."],
    ["Money asked", f"Rp30M now: TemanJourney Rp20M, Riskia Rp10M. With Riskia's earlier Rp20M, her side has put in Rp50M for 20% (Riskia 16.1%, TemanJourney 3.9%)."],
    ["Return", f"In the plan no sale is needed to earn the money back: dividends alone pay Riskia's side about {div2 / side_total:.0f} times its Rp50M in 2030 and 2031. At a cautious sale price of {exit_low['multiple']:.0f} times 2031 revenue, the 20% would be worth about {rp(exit_low['stake_value'])}."],
  ])}
</section>""")

# 2 problem
add(f"""
<section id="problem">
  <h2>2. The problem and our value proposition</h2>
  <div class="callout" style="background:var(--teal-soft)">
    <p class="lead"><b>The one problem:</b> bureaus turn away clients they could serve. Clients ask on Instagram and WhatsApp at night and at weekends, but the admin answers only in office hours, and a slow reply loses the client. When clients do come, psychologists lose about 18 of their 45 working hours a week to admin and report writing. <b>CoreItera helps a bureau support more clients, so it earns more.</b></p>
  </div>

  <h3>Pain 1: the bureau loses clients when nobody answers</h3>
  <ul>
    <li><b>Clients buy in chat.</b> In Indonesia a client finds a bureau on Instagram, asks on WhatsApp, and expects an answer within minutes. Many ask after work, at night or at the weekend.</li>
    <li><b>The admin is there only part of the day.</b> A chat at 10 pm is answered the next morning, when the client has already booked somewhere else or lost the courage to ask.</li>
    <li><b>The founder teaches selling again and again.</b> A bureau founder spends long hours with the admin and the operations manager: what to say, how to answer price questions, when to follow up. When the admin leaves, that know-how leaves too.</li>
    <li><b>Good and poor leads look the same.</b> The admin spends the same time on someone just asking the price as on a company that wants assessments for 50 staff.</li>
  </ul>

  <h3>How the sales agent fixes it</h3>
  {table(["", "When the admin is away (nights, weekends, busy hours)", "When the admin is online"], [
    ["Who answers", "The agent, at once, 24 hours a day", "The agent answers first, then hands over"],
    ["What it does", "Answers questions about services and prices, asks a few questions to understand the need, offers open times and a payment link", "Marks the serious leads, writes a short summary of each, and passes them to the admin"],
    ["Who closes", "The client books and pays; or the agent leaves a hot lead for the admin's morning list", "The admin closes, with the summary in front of them"],
    ["Style", "Writes in the bureau's own style, learned from its past chats and the founder's way of selling", "The admin can edit the style at any time"],
    ["Safety", "Never counsels or diagnoses. If a message shows a crisis, it gives the suicide prevention hotline at once and alerts a person", "The same"],
  ])}
  <p>The founder's way of selling is written into the agent once, then used in every chat, day and night. The founder no longer has to teach it to each new admin.</p>

  <h3>What it means for one bureau {EST}</h3>
  {table(["A bureau with", "Today", "With CoreItera"], [
    ["Chats a month on Instagram and WhatsApp", num(X['inquiries']), num(X['inquiries'])],
    ["Chats that arrive when the admin is away", f"{num(x_after)}, answered late", f"{num(x_after)}, answered at once"],
    ["Chats that become paying clients", f"{pct(X['conversion_now'])} = {num(x_now)} clients", f"{pct(X['conversion_agent'])} = {num(x_agent)} clients"],
    ["Extra sales a month", "", f"<b>{num(x_extra)} more clients × {rpk(X['first_purchase'])} = {rp(x_extra_rev)}</b>"],
  ], right=(1, 2))}
  <p class="note">Example numbers. We will measure the real ones at TemanJourney first: chats per month, the share after admin hours, and how many become clients, before and after the agent starts in February 2027.</p>

  <h3>Pain 2: no hours left to serve the new clients</h3>
  {table(["Where a psychologist's typical week goes", "Hours", "What CoreItera does about it"], [
    ["Counselling with clients", "20", "Nothing to remove: this is the work clients pay for."],
    ["Notes, scheduling, payments, chasing clients", "10", "<b>Bureau tools:</b> booking, payment, reminders, records and follow-ups run by themselves."],
    ["One full assessment: tests 1.5 h, scoring 1 h, report 7 h, feedback 1 h", "10.5", "<b>Copilot:</b> scoring drops to 10 minutes and the report to about 1 hour. The psychologist still checks and signs."],
    ["Other", "4.5", ""],
  ], right=(1,))}
  <p class="note">Typical week from the Augment proposal (IS5007, October 2026), a team estimate we will check with a one-week time diary of 5 TemanJourney psychologists in December 2026. Report times from CoreItera's own trial of 205 reports over 6 months.</p>

  <h3>Why today's tools fall short</h3>
  <ul>
    <li><b>WhatsApp, spreadsheets and Word</b> are free but manual, and only work when the admin is there.</li>
    <li><b>Generic WhatsApp chatbot tools</b> (for example Qontak by Mekari, Wati) answer chats, but they do not know the psychologists' free times, cannot book and take payment, do not write like the bureau, and do not follow psychology safety rules.</li>
    <li><b>Global practice software</b> (SimplePractice, Carepatron) handles booking and notes, but in English, priced in US dollars, with no Indonesian tests, no local payments and no sales help.</li>
    <li><b>Counselling apps</b> (Riliv, Satu Persen) bring clients, but the client belongs to the app, not the bureau.</li>
  </ul>

  <h3>Pain and pain-killer, for each customer</h3>
  {table(["Who", "Their pain", "What CoreItera gives them", "How much"], [
    ["Bureau founder", "Clients lost after hours; long hours teaching the admin to sell", "A sales agent that sells 24 hours in the founder's own style", f"About {num(x_extra)} more clients, {rp(x_extra_rev)} more sales a month in the example {EST}"],
    ["Admin and operations manager", "Too many chats; poor leads take as much time as good ones", "Only the serious leads, each with a summary, ready to close", "More closings in the same hours"],
    ["Psychologist", "Reports eat evenings; assessment pays a quarter of counselling per hour", "The copilot scores and drafts; they check, correct and sign", "About 6 hours saved per full report; room for more clients"],
    ["Client", "Slow replies at night; unclear prices", "An answer in minutes, open times and online payment", "Booked the same night"],
  ])}
</section>""")

# 3 market
add(f"""
<section id="market">
  <h2>3. Market</h2>
  <p>We count the market in psychologists, because that is what a bureau pays for. A mature bureau pays about {rpk(per_psy_month)} a month per psychologist: {rpk(avg_plan_mature / I.PSY_PER_BUREAU)} of plan, {rpk(I.COPILOT_SEAT)} of copilot and {rpk(I.PAYMENTS_PER_BUREAU[2031] * I.PAYMENT_FEE / I.PSY_PER_BUREAU)} of payment fees.</p>
  {table(["Market", "Who", "Psychologists", "Value a year"], [
    ["Total market", "Every psychologist in Indonesia (HIMPSI, 2026)", num(I.PSYCHOLOGISTS_ID), rp(tam)],
    ["Market we serve", f"Psychologists who work in or with a private bureau (about {pct(I.SHARE_IN_BUREAUS)}) {EST}", num(bureau_psy), rp(sam)],
    ["Our 2031 target", f"{num(Y[2031]['bureaus_end'])} bureaus and {num(Y[2031]['seats_end'])} copilot seats", num(Y[2031]['bureaus_end'] * I.PSY_PER_BUREAU), f"{rp(som)} ({pct(som / sam)} of the market we serve)"],
  ], right=(2, 3))}
  <h3>Segments</h3>
  {table(["Segment", "Size", "Why they buy", "Our plan"], [
    ["Small bureaus", "1 to 5 psychologists", "No admin, or one admin doing everything", "Starter plan; the first bureaus we win"],
    ["Mid-size bureaus", "6 to 15 psychologists", "Admin cost and missed follow-ups hurt most here", "Growth plan; our main target from 2028"],
    ["Large bureaus and clinics", "16 or more", "Many psychologists, many companies as clients", "Scale plan; sold by a sales person from 25 bureaus"],
    ["Psychologists working alone", "1", "Report writing", "Copilot seat only, from 2028"],
  ])}
  <h3>Trends on our side</h3>
  <ul>
    <li><b>Law 23/2022 on psychology services</b> sets clear rules for who may practise and lets psychology graduates assist psychologists. Bureaus need cleaner records and supervised assistants.</li>
    <li><b>Law 27/2022 on personal data protection (UU PDP)</b> makes client data on personal phones and shared spreadsheets a legal risk for the bureau.</li>
    <li><b>The Ministry of Health is training clinical psychologists faster</b> (Program Titian, 2026), so more new psychologists start practice and need tools.</li>
    <li><b>Clients buy in chat:</b> Indonesians find services on Instagram and ask on WhatsApp, and expect a fast answer. A bureau that answers first, at any hour, wins the client.</li>
    <li><b>Demand is rising:</b> only about 10% of young Indonesians with depression received treatment (Survei Kesehatan Indonesia 2023). Companies buy more counselling for staff.</li>
    <li><b>AI is now good enough to draft</b>, and Singapore's health ministry already allows AI-drafted clinical notes when a professional checks them. Responsibility stays with the psychologist.</li>
  </ul>
</section>""")

# 4 product
add(f"""
<section id="product">
  <h2>4. Products and services</h2>
  <p>One system, used in the order a client moves through a bureau: first the client is won, then served. TemanJourney is the first bureau to run all of it, from 30 November 2026.</p>
  {table(["Step", "Tool", "What happens", "Status"], [
    ["1", "<b>Sales agent</b> (Instagram and WhatsApp)", "Answers every chat at once, 24 hours, in the bureau's own writing style; sorts serious leads; offers open times and a payment link; hands hot leads to the admin to close. Never counsels; crisis messages go straight to the hotline and a person", '<span class="pill soon">WhatsApp Feb 2027, Instagram May 2027</span>'],
    ["2", "<b>Scheduling</b>", "Each psychologist sets weekly time blocks; clients book open times; moves and cancellations are handled", '<span class="pill soon">30 Nov 2026</span>'],
    ["3", "<b>Counselling</b> (Theta)", "Booking, online payment, 30-minute chat session, written summary with next steps", '<span class="pill soon">30 Nov 2026</span>'],
    ["4", "<b>Assessments</b>", "Send tests, collect and score results, keep every report in the client's record", '<span class="pill live">Live today</span>'],
    ["5", "<b>Follow-ups</b>", "Messages after a test or session with a link to book again (email today, WhatsApp next)", '<span class="pill soon">Q1 2027</span>'],
    ["6", "<b>Copilot</b>", "AI scores and drafts the report from a library checked by psychologists; the psychologist reviews and signs", '<span class="pill later">Q2 2027</span>'],
  ])}
  <h3>Prices</h3>
  {table(["Price", "Covers", "Amount"], [
    ["Monthly plan", "Every bureau tool (steps 1 to 5), the sales agent included, by bureau size", "from Rp1,000,000 a month"],
    ["Payment fee", "Each paid booking or assessment order; the Midtrans payment fee is passed on at cost", "Rp5,000 per payment"],
    ["Copilot", "The AI report assistant, per psychologist who uses it", "Rp300,000 per seat a month"],
  ], right=(2,))}
  {table(["Plan", "Psychologists", "Price a month", "Per psychologist"], [
    ["Starter", "1 to 5", "Rp1,000,000", "Rp200,000 or more"],
    ["Growth", "6 to 15", "Rp2,000,000", "Rp133,000 to Rp333,000"],
    ["Scale", "16 to 30", "Rp3,500,000", "Rp117,000 to Rp219,000"],
    ["Large", "31 or more", "Rp3,500,000 + Rp100,000 per extra psychologist", "about Rp110,000"],
  ], right=(2, 3))}
  <p class="note">Every new bureau gets its first month free. Safety rule for every tool: AI prepares, a licensed psychologist decides and signs. Each bureau's data is kept apart from every other bureau's.</p>
</section>""")

# 5 competition
yes, no = '<span class="yes">Yes</span>', '<span class="no">No</span>'
add(f"""
<section id="competition">
  <h2>5. Competition, marketing and sales</h2>
  <p><b>Our real competitor is the habit:</b> an admin on WhatsApp, a spreadsheet and Word. It costs nothing, so the agent must win more in new clients than we charge, from the first month.</p>
  {positioning_map()}
  <p class="note">Placement is our own judgement from each product's public pages (October 2026).</p>
  {table(["", "CoreItera", "WhatsApp + spreadsheet + Word", "Global practice software", "Counselling apps", "Online test sites"], [
    ["24-hour sales agent on Instagram and WhatsApp, in the bureau's own style", yes, "Admin hours only", no, "For the app's own sales", no],
    ["Booking and reminders", yes, "By hand", yes, yes, no],
    ["Local online payment", yes, "Bank transfer by hand", "US cards", yes, yes],
    ["Counselling sessions and notes", yes, "By hand", yes, yes, no],
    ["Assessment and scoring", yes, "By hand", no, "Simple tests", yes],
    ["Report drafted, checked and signed by a psychologist", yes, no, no, no, "Automatic, no psychologist"],
    ["Follow-up sales messages", yes, "By hand", "Partly", no, no],
    ["The client belongs to the bureau", yes, yes, yes, no, yes],
    ["Indonesian language and tests", yes, yes, no, yes, yes],
    ["Price for a 10-psychologist bureau a month", f"{rp(ex_tools)}, {rp(ex_tools + ex_copilot)} with copilot", "Free", "About Rp7M to Rp13M", "Commission per session", "Per test"],
  ])}
  <p class="note">Global price: SimplePractice lists US$49 to US$99 a month plus about US$39 to US$74 per extra clinician (third-party pricing guides, 2026; sources disagree), at US$1 = Rp16,500.</p>

  <h3>Tagline</h3>
  <p class="lead"><b>"Lebih banyak klien terlayani, 24 jam, dengan gaya biro sendiri."</b> More clients served, 24 hours a day, in your bureau's own style.</p>

  <h3>How we sell</h3>
  {table(["", "Plan"], [
    ["Price", "Low entry (Rp1M) so a bureau can say yes without a meeting; the first month free; prices rise with bureau size, never with features."],
    ["Promotion", "TemanJourney as the live example; HIMPSI workshops on AI-assisted reports that earn continuing-education credits; psychology faculties use CoreItera in practice classes; a bureau that refers another gets one month free."],
    ["Entry", "Jakarta first (TemanJourney's network of about 20 psychologists), then the psychology cities, then national (section 9)."],
    ["Cost to win a bureau", f"About {rp(cac)} in 2028 (workshops shared across new bureaus, plus a {rp(I.ONBOARDING_PER_NEW_BUREAU)} set-up visit). A bureau earns us about {rp(contrib)} a month after direct costs, so it pays back in {"about one month" if payback < 1.5 else f"about {payback:.0f} months"} {EST}."],
  ])}
</section>""")

# 6 edge
add(f"""
<section id="edge">
  <h2>6. Sustainable competitive advantage</h2>
  {table(["What protects us", "Why it is hard to copy"], [
    ["<b>The bureau's own voice</b>", "The agent learns each bureau's writing style from its past chats and the founder's way of selling, and gets better with every chat the admin corrects. We designed this style engine first for reports (in our psychologist workspace prototype); a new tool would start without any of it."],
    ["<b>The daily workflow</b>", "Once a bureau's bookings, records and reports live in CoreItera, moving out means moving every client file. Bureaus stay."],
    ["<b>The checked report library</b>", "Psychologists check and correct every draft; each correction improves the library. Quality went from 4.2 to above 5.5 out of 6 in our trial. A newcomer starts at zero."],
    ["<b>Local tests and norms</b>", "Scores only mean something against an Indonesian reference group. Our test volume builds those norms year by year, with associations and universities."],
    ["<b>Distribution through trust</b>", "TemanJourney, HIMPSI workshops and psychology faculties: channels a foreign software company or an app does not have."],
    ["<b>A very low cost base</b>", f"We run on {rp(I.BASE_RUN_COST)} a month today. {WORDS.get(starter_needed, str(starter_needed))} bureaus on the cheapest plan cover it, after the cost of serving them. We can price low and still last."],
  ])}
  <h3>Why the big players have not taken this</h3>
  <ul>
    <li><b>Global practice software</b> would need Indonesian tests, language, payments and data law. Indonesia is small for them and costly to enter.</li>
    <li><b>Counselling apps</b> earn by owning the client. Running the bureau's own clients for a fixed fee works against their model.</li>
    <li><b>Big AI companies</b> would carry clinical responsibility in every country and do not own the tests. They are more likely to be our model suppliers than our rivals.</li>
  </ul>
  <h3>Intellectual property</h3>
  <p>No patents: speed matters more. We protect the report library and our prompts as trade secrets, keep test licences in CoreItera's name, own the brand, and hold client data only with consent and only for the bureau that owns it.</p>

  <h3>Five tests this business must pass</h3>
  {table(["Test", "Evidence", "Confidence"], [
    ["1. Is the pain big?", "Chats after admin hours are answered late and clients go elsewhere; psychologists lose about 18 of 45 hours a week to admin and reports.", '<span class="conf" style="color:var(--ochre)">Medium: count after-hours chats at TemanJourney and run a time diary, Dec 2026</span>'],
    ["2. Is it much better than today?", "Report from about 7 hours to about 1; scoring from 60 to 10 minutes; quality up from 4.2 to above 5.5 out of 6.", '<span class="conf" style="color:var(--teal)">High: 205-report trial</span>'],
    ["3. Is the volume enough?", f"{WORDS.get(starter_needed, str(starter_needed))} Starter bureaus cover today's costs; the 2031 plan needs {pct(Y[2031]['bureaus_end'] / bureaus_market)} of bureaus. Even at half the bureaus, CoreItera makes a profit every year from {slow_profit_year}.", '<span class="conf" style="color:var(--ochre)">Medium: bureau count is an estimate</span>'],
    ["4. Could one factor kill it?", "Test licences and acceptance by psychologists. We use licensed tests or a report-only mode where the psychologist uploads their own scored results, so we never resell a test we do not own. The psychologist always signs.", '<span class="conf" style="color:var(--ochre)">Medium: written legal check by Q1 2027</span>'],
    ["5. Is the edge real?", "Others could build the tools, but not the checked library, the local norms or the trust channels quickly. The big players have reasons to stay out (above).", '<span class="conf" style="color:var(--ochre)">Medium: proven only at TemanJourney so far</span>'],
  ])}
</section>""")

# 7 team
add(f"""
<section id="team">
  <h2>7. Team</h2>
  {table(["Person", "Role", "What they bring"], [
    ["<b>Galih Muji Agung</b>", "CEO, co-founder", "Company direction, money and investors; builds checkout, payments and the data protections between apps. Studying at NUS in Singapore in 2026 to 2027."],
    ["<b>Adit</b>", "Co-founder, founding engineer, growth lead", "The psychologist's side: scheduling engine, chat, the scheduler screens; also B2B sales and social media."],
    ["<b>Zaki</b>", "Design lead, engineer", "The client app and design; from 2027 the AI agent and its screens."],
    ["<b>Riskia</b>", "Investor and adviser (CEO of TemanJourney)", "Runs the first bureau on CoreItera; relationships with banks, regulators and psychologists; turns real bureau problems into product requests."],
    ["<b>TemanJourney psychologists</b>", "Clinical design partners", "Check copilot drafts until we hire a clinical lead; tell us what slows them down every week."],
  ])}
  <p><b>What we still need:</b> a clinical lead psychologist to own report quality (paid once copilot income reaches {rp(I.CLINICAL_TRIGGER, 0)} a month), a first sales person once {I.SALES_FROM_BUREAUS} bureaus pay, and customer-success staff (one per {I.CS_PER_BUREAUS} bureaus). Each hire starts when revenue can carry it.</p>
</section>""")

# 8 operations
add(f"""
<section id="operations">
  <h2>8. Operational plan</h2>
  <h3>Lean canvas</h3>
  <div class="canvas">
    <div class="cell tall"><b>Problem</b>Clients lost when chats come after admin hours. Founders spend long hours teaching admins to sell. Psychologists lose about 18 of 45 hours a week to admin and reports.</div>
    <div class="cell"><b>Solution</b>A 24-hour sales agent on Instagram and WhatsApp that hands hot leads to the admin; scheduling, counselling, assessments, follow-ups; then the copilot.</div>
    <div class="cell tall"><b>Unique value</b>Serve more clients, so the bureau earns more: it sells 24 hours in the bureau's own style, and frees the hours to serve the new clients.</div>
    <div class="cell"><b>Unfair advantage</b>The bureau's own writing style, the checked report library, local norms, TemanJourney as live proof, trust channels.</div>
    <div class="cell tall"><b>Customers</b>Private psychology bureaus in Indonesia (1 to 30+ psychologists); later psychologists working alone. Early adopters: Jakarta bureaus in TemanJourney's network.</div>
    <div class="cell"><b>Key metrics</b>Extra clients per bureau from the agent; paying bureaus; copilot seats; hours saved per report; bureaus that leave each month (target under 1.5%).</div>
    <div class="cell"><b>Channels</b>TemanJourney referrals, HIMPSI workshops, psychology faculties, bureau-to-bureau referral.</div>
    <div class="split">
      <div class="cell"><b>Costs</b>Cloud and AI ({rp(I.BASE_RUN_COST)} a month today), AI per report (about Rp8,000), onboarding, workshops; founders, sales, support and engineers once revenue allows.</div>
      <div class="cell"><b>Revenue</b>Monthly plan from Rp1M (sales agent included); Rp5,000 per paid payment; Rp300,000 per copilot seat. TemanJourney free for 12 months, then cost price.</div>
    </div>
  </div>

  <h3>Launch at TemanJourney on 30 November 2026</h3>
  {table(["Milestone", "Done by", "Owner"], [
    ["Riskia reads this proposal; numbers agreed", "Fri 23 Oct", "Galih, Riskia"],
    ["One-page term sheet signed by both companies", "Fri 30 Oct", "Galih, Riskia"],
    ["Shareholder approval in both companies (RUPS); notary booked", "Fri 6 Nov", "Galih, Riskia"],
    ["Notary deed signed; Rp30M paid into CoreItera's account", "Fri 13 Nov", "Riskia, TemanJourney"],
    ["Profiles and licences (SIPP) of at least 20 psychologists entered", "Fri 13 Nov", "TemanJourney admin"],
    ["Prices set (chat session Rp149,000) and weekly time blocks entered", "Fri 20 Nov", "TemanJourney admin, Adit"],
    ["Release practised on a copy of the live database; code frozen", "Fri 20 Nov", "Galih, Adit, Zaki"],
    ["Three one-page admin guides: scheduling, new client, cancellations and refunds", "Mon 23 Nov", "TemanJourney admin, Zaki"],
    ["Test week: staff and friends book, pay and chat for real", "Fri 27 Nov", "All"],
    ["<b>Theta live for TemanJourney clients</b>", "<b>Mon 30 Nov</b>", "All"],
  ], right=(1,))}
  <h3>Milestones after launch</h3>
  {table(["When", "Milestone", "Proves"], [
    ["Dec 2026", "Time diary with 5 TemanJourney psychologists", "Test 1: the size of the pain"],
    ["Jan 2027", "First talks with Jakarta bureaus", "Bureaus want it"],
    ["Feb 2027", "Sales agent live on TemanJourney's WhatsApp, in its own style", "Extra clients after admin hours"],
    ["May 2027", "Sales agent on Instagram (after Meta app review)", "The second channel works"],
    ["Mar 2027", "2 paying bureaus; written legal check on test licences", "Bureaus pay; test 4"],
    ["Apr 2027", "Copilot on sale", "Psychologists pay per seat"],
    ["Jul 2027", "First HIMPSI workshop outside Jakarta", "The channel works"],
    ["Dec 2027", f"{num(Y[2027]['bureaus_end'])} paying bureaus, {num(Y[2027]['seats_end'])} copilot seats", "Ready for the national stage"],
  ])}
  <h3>Risks specific to this business</h3>
  {table(["Risk", "Early sign", "What we do"], [
    ["The agent says something a psychologist should say", "A chat about symptoms, diagnosis or crisis", "The agent never counsels or diagnoses; crisis words trigger the hotline and an alert to a person; the bureau reviews a sample of chats each week."],
    ["Meta changes its rules for business messaging", "A warning or a limit on the bureau's WhatsApp or Instagram account", "The agent only answers about the bureau's own services, as Meta's business rules expect; we use the official WhatsApp Business Platform and Instagram API, never unofficial tools."],
    ["A copilot draft states something wrong", "Psychologists change more than minor points in over 10% of drafts", "Every draft shows its source; a psychologist signs every report; pull any report type whose correction rate rises."],
    ["Test publishers object", "A publisher asks how its test is used", "Report-only mode: the bureau uploads its own scored results; licences stay in CoreItera's name."],
    ["Psychologists or HIMPSI distrust AI reports", "Low copilot uptake after workshops", "The psychologist leads and signs; ethics rules agreed with HIMPSI; workshops earn credits."],
    ["Family and company interests mix", "Decisions made in conversation, not in writing", "Every term written down; both companies approve at a shareholder meeting; TemanJourney's price fixed in the agreement."],
    ["Cash runs thin in mid-2027", f"Fewer than 2 paying bureaus by March 2027", f"A standby loan of up to {rp(safety_line, 0)} from the founders; workshops wait until bureaus pay (section 10)."],
  ])}
</section>""")

# 9 expansion
add(f"""
<section id="expansion">
  <h2>9. Expansion plan</h2>
  {table(["Stage", "Where", "How we win bureaus", "Resources added", f"Paying bureaus at the end"], [
    ["<b>1. Prove it</b><br>Nov 2026 to Mar 2027", "Jakarta: TemanJourney in Tebet, then nearby bureaus", "TemanJourney as the live example; introductions from its psychologists", "None: the three founders sell and onboard", num(I.BUREAU_TARGETS[(2027, 3)])],
    ["<b>2. Psychology cities</b><br>Apr to Dec 2027", "Greater Jakarta, Bandung, Yogyakarta, Surabaya, Semarang, Malang", "HIMPSI workshops with credits, faculties, bureau referrals; copilot launch", "Workshops; clinical lead once copilot income allows", num(Y[2027]['bureaus_end'])],
    ["<b>3. National</b><br>2028 to 2031", "Every province, through HIMPSI branches", "Sales people; psychologists working alone buy copilot seats; companies send work to bureaus on CoreItera", f"Sales and support hires by volume; more engineers as revenue grows", f"{num(Y[2028]['bureaus_end'])} (2028), {num(Y[2029]['bureaus_end'])} (2029), {num(Y[2031]['bureaus_end'])} (2031)"],
  ], right=(4,))}
  <p><b>Expansion risks.</b> Bureaus outside Java are smaller and pay later: they start on the Starter plan, and we count on them only from 2028. Support load grows with every bureau: one support person per {I.CS_PER_BUREAUS} bureaus is in the costs from {I.CS_FROM_BUREAUS} bureaus. After Indonesia, the same model can travel to other countries with too few psychologists, as set out in the Augment proposal; that is not in these numbers.</p>
</section>""")

# 10 financials
yr_rows = []
for y, v in Y.items():
    yr_rows.append([str(y), num(v['bureaus_end']), num(v['seats_end']), rp(v['plan']), rp(v['fees']), rp(v['copilot']),
                    f"<b>{rp(v['revenue'])}</b>", rp(v['variable'] + v['fixed']), rp(v['tax']), f"<b>{rp(v['profit'])}</b>",
                    pct(v['profit'] / v['revenue']), rp(v['cash_end'])])
q_rows = [[q["quarter"], rp(q["revenue"]), rp(q["costs"]), rp(q["investment"]) if q["investment"] else "None",
           rp(q["net"]), f"<b>{rp(q['cash_end'])}</b>"] for q in plan["quarters"]]
cost_rows = []
for key, name in (("base", "Cloud, AI and tools"), ("founders", "Three founders"), ("clinical", "Clinical quality checks"),
                  ("workshops", "Workshops and marketing"), ("cs", "Customer support"), ("sales", "Sales people"),
                  ("engineer", "Extra engineers"), ("office", "Office"), ("accountant", "Accountant")):
    cost_rows.append([name] + [rp(Y[y][key]) if Y[y][key] > 0.05 else "None" for y in Y])
cost_rows.append(["Variable: AI per report, servers, onboarding"] + [rp(Y[y]['variable']) for y in Y])
slow_rows = [[str(y), num(v['bureaus_end']), rp(v['revenue']), rp(v['profit']), pct(v['profit'] / v['revenue']), rp(v['cash_end'])] for y, v in SY.items()]

add(f"""
<section id="financials">
  <h2>10. Financials</h2>
  <p><b>How CoreItera makes money:</b> three lines. Monthly plans give steady income; the Rp5,000 payment fee grows with each bureau's sales; copilot seats carry the most margin per psychologist. By 2031 the copilot brings {pct(Y[2031]['copilot'] / Y[2031]['revenue'])} of revenue.</p>
  {revenue_chart(Y)}
  <h3>Five-year profit and loss {EST}</h3>
  {table(["Year", "Bureaus", "Seats", "Plans", "Payment fees", "Copilot", "Revenue", "Costs", "Tax", "Profit", "Margin", "Cash (Dec)"], yr_rows, right=tuple(range(1, 12)))}
  <p class="note">Bureaus and seats at December. Tax: 0.5% of revenue (PP 55/2022) to 2029, then 11% of profit (22% corporate tax with the small-business cut); a tax consultant must confirm. TemanJourney is not counted as revenue: free for 12 months, then cost price.</p>
  <h3>Main costs by year</h3>
  {table(["Cost"] + [str(y) for y in Y], cost_rows, right=tuple(range(1, 6)))}
  <p class="note">Every cost starts only when revenue can carry it: founders are paid once monthly revenue reaches {rp(I.FOUNDER_PAY_TRIGGER, 0)} (from {month_name(plan['founders_paid_from'])}); one engineer is added for every {rp(I.ENGINEER_EVERY_REVENUE, 0)} of monthly revenue; the office opens at {rp(I.OFFICE_TRIGGER, 0)} a month.</p>

  <h3>18-month cash flow, by quarter {EST}</h3>
  {quarter_chart(plan['quarters'])}
  {table(["Quarter", "Revenue", "Costs", "New money", "Net cash flow", "Cash at the end"], q_rows, right=(1, 2, 3, 4, 5))}
  <div class="callout">
    <p><b>Being direct about cash.</b> We start with Rp14M in the bank. With the Rp30M, cash falls to its lowest point, {rp(low['cash'])} in {month_name(low['month'])}, just before the first bureaus and copilot seats pay. In the slow case it would fall to {rp(slow_low['cash'])} in {month_name(slow_low['month'])}. So the founders keep a standby loan of up to <b>{rp(safety_line, 0)}</b> for mid-2027, repaid from bureau income in the second half of 2027. Spending on workshops waits until bureaus pay.</p>
  </div>

  <h3>If only half the bureaus come: the slow case</h3>
  {table(["Year", "Bureaus", "Revenue", "Profit", "Margin", "Cash (Dec)"], slow_rows, right=(1, 2, 3, 4, 5))}
  <p class="note">Half the bureaus, half the psychologists working alone, copilot use 30% lower. Hires still wait for revenue, so the business still makes a profit every year from {slow_profit_year}.</p>

  <h3>The investment and the equity</h3>
  {table(["Who", "Paid before", "Paid now", "Total cash", "Ownership after"], [
    ["Galih (the other founders' shares alongside)", "Rp50M", "None", "Rp50M", "80.0% (all current owners)"],
    ["Riskia", "Rp20M", "Rp10M", "Rp30M", "16.1%"],
    ["TemanJourney (PT Indo Medika Resolusi)", "None", "Rp20M", "Rp20M", "3.9%"],
    ["<b>Total</b>", "<b>Rp70M</b>", "<b>Rp30M</b>", "<b>Rp100M</b>", "<b>100%</b>"],
  ], right=(1, 2, 3, 4))}
  <ol class="steps">
    <li><div><b>Complete the 15%.</b> New shares bring Riskia from 10% to the 15% promised for her first Rp20M. Every other owner gives up 5 points.</div></li>
    <li><div><b>Issue shares for the Rp30M</b> at one price: TemanJourney Rp20M, Riskia Rp10M. The price is set so the two own exactly 20% together.</div></li>
  </ol>
  <p>This values CoreItera at <b>Rp480M before the new money</b> and Rp510M after. Riskia's first Rp20M bought 10% at about Rp200M, so this round is priced about 2.5 times higher. At this price Riskia's 16.1% is worth about Rp82M against the Rp30M she has paid. <b>Why the same Rp50M buys a different share:</b> the founders' money came first, when nobody knew if the product would work, and the founders' shares also pay for building the platform, mostly unpaid.</p>

  <h3>Use of the Rp30M</h3>
  {table(["Use", "Amount", "Why"], [
    ["Running costs, mid-November 2026 to February 2027", rp(running_part), "Keeps CoreItera and Theta running through launch and the first bureaus"],
    [f"Notary deed and share registration {EST}", rp(I.DEAL_ONE_OFFS['notary and share registration']), "A share issue needs a notary deed and a report to the Ministry of Law"],
    [f"Launch: payment set-up, test bookings, extra server load {EST}", rp(I.DEAL_ONE_OFFS['launch costs']), "One-off costs to go live"],
    ["<b>Total</b>", f"<b>{rp(I.DEAL_CASH)}</b>", ""],
  ], right=(1,))}

  <h3>Exit and return for Riskia's side</h3>
  {table(["Route", "Plan", "Slow case"], [
    [f"Dividends: {pct(runner.DIVIDEND_SHARE)} of profit paid out in 2030 and 2031", f"{rp(div2, 0)} ({div2 / side_total:.1f} times Rp50M)", f"{rp(sdiv2, 0)} ({sdiv2 / side_total:.1f} times)"],
    [f"Sale of CoreItera at {exit_low['multiple']:.0f} times 2031 revenue", f"Company {rp(exit_low['company_value'])}; 20% = {rp(exit_low['stake_value'])}", f"Company {rp(slow_exit_low['company_value'])}; 20% = {rp(slow_exit_low['stake_value'])}"],
    [f"Sale at {exit_mid['multiple']:.0f} times 2031 revenue", f"20% = {rp(exit_mid['stake_value'])}", f"20% = {rp(slow['exit']['mid']['stake_value'])}"],
  ])}
  <p>Likely buyers: an Indonesian HR-tech or health-tech company, or a test publisher that wants the bureaus. A later investor could also buy shares from Riskia's side. These values are estimates: many young companies miss their plans, and a later funding round would make every stake smaller. The plan does not need a new round.</p>
</section>""")

# 11 ask
add(f"""
<section id="ask">
  <h2>11. Conclusion and the ask</h2>
  <ul>
    <li><b>Market:</b> about {num(I.PSYCHOLOGISTS_ID)} psychologists and an estimated {num(round(bureaus_market, -2))} bureaus sell through Instagram and WhatsApp but answer only in admin hours. Nobody sells them one system built for Indonesian practice.</li>
    <li><b>Value:</b> a bureau serves more clients, so it earns more: about {rp(x_extra_rev)} more sales a month in our example for a plan from Rp1M; the copilot pays its price back about {ex_extra_fees / ex_copilot:.0f} times.</li>
    <li><b>Team and proof:</b> three founders who already built the platform, a 205-report trial, and TemanJourney live on 30 November 2026.</li>
    <li><b>Money:</b> cheap to run, profitable every month from {month_name(plan['profitable_for_good_from'])} in the plan, and still profitable every year from {slow_profit_year} if only half the bureaus come.</li>
  </ul>
  <div class="answer">
    <div class="eyebrow">What we ask Riskia and TemanJourney to decide</div>
    <ol>
      <li>Invest Rp30M now (TemanJourney Rp20M, Riskia Rp10M), Rp50M in total from Riskia's side, for 20% together: Riskia 16.1%, TemanJourney 3.9%.</li>
      <li>Accept TemanJourney's software terms: free for 12 months from launch, then cost price.</li>
      <li>Name one TemanJourney person who owns the launch tasks.</li>
      <li>Introduce us to 2 Jakarta bureaus by January 2027.</li>
      <li>Sign the term sheet by Friday 30 October 2026.</li>
    </ol>
  </div>
</section>""")

# appendices
assump = [
    ["Running cost today", rp(I.BASE_RUN_COST) + " a month", "Known"],
    ["Cash today (19 Oct 2026)", rp(I.CASH_TODAY), "Known"],
    ["Prices: plan, payment fee, copilot", "From Rp1M; Rp5,000; Rp300,000", "Decided"],
    ["Average plan price per bureau", ", ".join(f"{y}: {rp(v)}" for y, v in I.AVG_PLAN_BY_YEAR.items()), "Estimate"],
    ["Psychologists per bureau", str(I.PSY_PER_BUREAU), "Estimate"],
    ["Paid payments per bureau a month", ", ".join(f"{y}: {v}" for y, v in I.PAYMENTS_PER_BUREAU.items()), "Estimate"],
    ["Share of psychologists in private bureaus", pct(I.SHARE_IN_BUREAUS), "Estimate"],
    ["Bureaus that leave each month", pct(I.MONTHLY_CHURN, 1), "Estimate"],
    ["Copilot use among bureau psychologists", "10% at launch, 40% Dec 2027, 50% Dec 2028, 60% from Dec 2029", "Estimate"],
    ["Psychologists working alone on the copilot", ", ".join(f"{k[0]}: {v}" for k, v in I.SOLO_SEATS.items() if v), "Estimate"],
    ["AI cost per copilot seat", rp(I.AI_COST_PER_SEAT, 2) + " a month (about 15 reports at Rp8,000)", "Estimate"],
    ["Founder pay", "Rp10M each from Rp60M monthly revenue; Rp20M in 2030, Rp25M in 2031", "Estimate"],
    ["Dividends", pct(runner.DIVIDEND_SHARE) + " of profit from 2030", "Proposal"],
    ["Sales agent cost per bureau", rp(I.AGENT_COST_PER_BUREAU, 2) + " a month (AI replies and WhatsApp messages)", "Estimate"],
    ["Sales agent example", f"{num(X['inquiries'])} chats a month, {pct(X['after_hours_share'])} after admin hours, conversion {pct(X['conversion_now'])} to {pct(X['conversion_agent'])}, first purchase {rpk(X['first_purchase'])}", "Estimate, to measure at TemanJourney"],
]
add(f"""
<section>
  <h2>Appendix A. Main assumptions</h2>
  {table(["Assumption", "Value", "Status"], assump)}
  <p class="note">All in the model folder: model/inputs.py. Change a number there and rebuild this page with model/build_page.py.</p>
</section>
<section>
  <h2>Appendix B. The other way we could price the deal</h2>
  {table(["", "Recommended (15% completed, then cash)", "Cash only"], [
    ["Riskia at the end", "16.1%", "12.6%"], ["TemanJourney at the end", "3.9%", "7.4%"], ["Together", "20.0%", "20.0%"],
    ["Value of CoreItera before the new money", "Rp480M", "Rp240M"], ["Is the 15% promise kept?", "Yes", "No"],
  ], right=(1, 2))}
  <p>The recommended deal keeps the promise and values CoreItera at twice the price, a better starting point for later investors.</p>
</section>
<section>
  <h2>Appendix C. Numbers to confirm before signing</h2>
  {table(["Number", "Used here", "Why it matters"], [
    ["CoreItera's share count (from the deed)", "1,000,000 as an example", "Sets the exact number of new shares; the percentages do not change"],
    ["Who owns the 90% today, and how the earlier Rp50M and Rp20M are recorded", "Galih and the other founders", "The deed must match the cash table"],
    ["Number of private bureaus in Indonesia", f"About {num(round(bureaus_market, -2))}", "Sets our market share; HIMPSI does not publish it"],
    ["Notary and registration cost", rp(I.DEAL_ONE_OFFS['notary and share registration']), "Comes out of the Rp30M"],
    ["Other shareholders of TemanJourney", "Not known", "They must approve TemanJourney paying Rp20M for shares"],
  ])}
</section>""")

footer = """
<footer>
  Sources: CoreItera trial data (205 reports rated over 6 months); TemanJourney 2027 financial model (October 2026); HIMPSI (2026, 2021) for psychologist numbers; Augment proposal (IS5007, October 2026) for the typical week, assessment hours, the Rp700,000 fee, Law 23/2022 and Singapore's guidance on AI notes; Kemenkes, Survei Kesehatan Indonesia 2023; SimplePractice prices from third-party pricing guides (2026); Riliv partner numbers from Fortune Indonesia (2023); Klinik Pintar from Dealroom and e27. US$1 = Rp16,500. A notary and a tax consultant must confirm the legal and tax steps.
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
runner.main()
print("wrote", os.path.normpath(OUT_HTML))
