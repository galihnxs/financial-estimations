"""Monthly engine for the CoreItera investment model. Pure Python, no libraries.

run(inputs, with_round) -> dict with monthly rows, yearly P&L, quarterly cash flow and summary.
"""
import math


def months(start, end):
    y, m = start
    while (y, m) <= end:
        yield (y, m)
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)


def idx(ym):
    return ym[0] * 12 + ym[1]


def interp(points, ym, before=0.0):
    """Straight line between dated points; flat after the last one."""
    keys = sorted(points)
    if idx(ym) < idx(keys[0]):
        return before
    for a, b in zip(keys, keys[1:]):
        if idx(a) <= idx(ym) <= idx(b):
            t = (idx(ym) - idx(a)) / (idx(b) - idx(a))
            return points[a] + t * (points[b] - points[a])
    return points[keys[-1]]


def run(I, with_round=False):
    rows = []
    cash = I.CASH_TODAY
    bureaus_prev = 0.0
    founders_on = False
    prev_revenue = 0.0
    prev_copilot = 0.0
    acct_on = clin_on = sales_on = office_on = False
    engineers = 0
    copilot_on = False
    for ym in months(I.START, I.END):
        y, m = ym
        r = {"ym": ym, "label": f"{y}-{m:02d}"}

        # --- volume ---
        target = interp(I.BUREAU_TARGETS, ym) if y >= 2027 else 0.0
        kept = bureaus_prev * (1 - I.MONTHLY_CHURN)
        new = max(0.0, target - kept)
        bureaus = kept + new
        paying = kept if I.FREE_FIRST_MONTH else bureaus     # new bureaus pay from month 2
        copilot_on = copilot_on or idx(ym) >= idx(I.COPILOT_LAUNCH)
        adoption = interp(I.COPILOT_ADOPTION, ym) if copilot_on else 0.0
        solo = interp(I.SOLO_SEATS, ym) if copilot_on else 0.0
        seats = adoption * I.PSY_PER_BUREAU * paying + solo
        r.update(bureaus=bureaus, new=new, paying=paying, seats=seats, solo=solo)

        # --- revenue ---
        if y >= 2027:
            plan = paying * I.AVG_PLAN_BY_YEAR[y]
            fees = paying * I.PAYMENTS_PER_BUREAU[y] * I.PAYMENT_FEE
        else:
            plan = fees = 0.0
        copilot = seats * I.COPILOT_SEAT
        revenue = plan + fees + copilot
        r.update(plan=plan, fees=fees, copilot=copilot, revenue=revenue)

        # --- variable cost ---
        ai = seats * I.AI_COST_PER_SEAT
        server = bureaus * I.SERVER_PER_BUREAU
        agent = bureaus * I.AGENT_COST_PER_BUREAU if idx(ym) >= idx(I.AGENT_WHATSAPP_LAUNCH) else 0.0
        onb_rate = I.ONBOARDING_FOUNDERS_COST if bureaus <= I.ONBOARDING_FOUNDERS_UNTIL else I.ONBOARDING_PER_NEW_BUREAU
        onboarding = new * onb_rate
        acquisition = new * I.CAC_PER_NEW_BUREAU
        variable = ai + server + agent + onboarding + acquisition

        # --- fixed cost ---
        growth = (1 + I.INFRA_GROWTH_PER_YEAR) ** max(0, y - 2027)
        base = I.BASE_RUN_COST * growth
        if ym == I.START:
            base *= I.OCT_DAYS_LEFT
        if prev_revenue >= I.ACCOUNTANT_TRIGGER:
            acct_on = True
        accountant = I.ACCOUNTANT if acct_on else 0.0
        if prev_copilot >= I.CLINICAL_TRIGGER:
            clin_on = True
        clinical = I.CLINICAL_BY_YEAR.get(y, 0.0) if clin_on else 0.0
        workshops = I.WORKSHOPS.get(y, 0.0) if idx(ym) >= idx(I.WORKSHOPS_FROM) else 0.0
        if prev_revenue >= I.FOUNDER_PAY_TRIGGER:
            founders_on = True
        founders = I.FOUNDERS * I.FOUNDER_PAY_EACH.get(y, 10.0) if founders_on else 0.0
        cs = math.ceil(bureaus / I.CS_PER_BUREAUS) * I.CS_PAY if bureaus >= I.CS_FROM_BUREAUS else 0.0
        if bureaus >= I.SALES_FROM_BUREAUS:
            sales_on = True
        sales = (1 + math.floor(bureaus / I.SALES_EVERY_BUREAUS)) * I.SALES_PAY if sales_on else 0.0
        engineers = max(engineers, math.floor(prev_revenue / I.ENGINEER_EVERY_REVENUE))
        engineer = engineers * I.ENGINEER_PAY
        if prev_revenue >= I.OFFICE_TRIGGER:
            office_on = True
        office = I.OFFICE if office_on else 0.0
        ga = I.GA_SHARE_OF_REVENUE * prev_revenue
        fixed = base + accountant + clinical + workshops + founders + cs + sales + engineer + office + ga
        r.update(ai=ai, server=server, agent=agent, onboarding=onboarding, acquisition=acquisition, variable=variable, base=base,
                 accountant=accountant, clinical=clinical, workshops=workshops, founders=founders,
                 cs=cs, sales=sales, engineer=engineer, office=office, ga=ga, fixed=fixed)

        # --- profit and tax ---
        before_tax = revenue - variable - fixed
        if y <= I.FINAL_TAX_UNTIL:
            tax = revenue * I.FINAL_TAX_RATE
        else:
            tax = max(0.0, before_tax) * I.PROFIT_TAX_RATE
        profit = before_tax - tax
        r.update(gross=revenue - variable, before_tax=before_tax, tax=tax, profit=profit)

        # --- cash ---
        deal = I.DEAL_CASH if ym == I.DEAL_MONTH else 0.0
        one_offs = sum(I.DEAL_ONE_OFFS.values()) if ym == I.DEAL_MONTH else 0.0
        round_in = I.NEXT_ROUND["amount"] if (with_round and ym == I.NEXT_ROUND["month"]) else 0.0
        cash = cash + profit + deal + round_in - one_offs
        r.update(deal=deal, one_offs=one_offs, round_in=round_in, cash=cash)
        rows.append(r)
        bureaus_prev = bureaus
        prev_revenue = revenue
        prev_copilot = copilot

    return summarise(I, rows, with_round)


def summarise(I, rows, with_round):
    def total(rs, key):
        return sum(r[key] for r in rs)

    years = {}
    for y in range(2027, 2032):
        rs = [r for r in rows if r["ym"][0] == y]
        years[y] = {k: total(rs, k) for k in
                    ("plan", "fees", "copilot", "revenue", "variable", "gross", "base", "accountant",
                     "clinical", "workshops", "founders", "cs", "sales", "engineer", "office", "ga", "acquisition", "fixed",
                     "before_tax", "tax", "profit")}
        last = rs[-1]
        years[y].update(bureaus_end=last["bureaus"], seats_end=last["seats"], cash_end=last["cash"],
                        monthly_revenue_dec=last["revenue"])

    quarters = []
    q_start = [(2026, 10), (2027, 1), (2027, 4), (2027, 7), (2027, 10), (2028, 1)]
    names = ["Q4 2026", "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027", "Q1 2028"]
    for name, (y, m) in zip(names, q_start):
        rs = [r for r in rows if r["ym"][0] == y and m <= r["ym"][1] < m + 3]
        quarters.append({"quarter": name,
                         "revenue": total(rs, "revenue"),
                         "costs": total(rs, "variable") + total(rs, "fixed") + total(rs, "tax") + total(rs, "one_offs"),
                         "investment": total(rs, "deal") + total(rs, "round_in"),
                         "net": total(rs, "profit") + total(rs, "deal") + total(rs, "round_in") - total(rs, "one_offs"),
                         "cash_end": rs[-1]["cash"]})

    low = min(rows, key=lambda r: r["cash"])
    breakeven = next((r["label"] for r in rows
                      if r["ym"][0] >= 2027 and r["before_tax"] > 0
                      and all(x["before_tax"] > 0 for x in rows if idx(x["ym"]) > idx(r["ym"]))), None)
    first_positive = next((r["label"] for r in rows if r["ym"][0] >= 2027 and r["before_tax"] > 0), None)
    founders_from = next((r["label"] for r in rows if r["founders"] > 0), None)

    rev_2031 = years[2031]["revenue"]
    stake = sum(v for k, v in I.OWN_AFTER.items() if k != "current owners")
    stake_after = stake * (1 - I.NEXT_ROUND["dilution"]) if with_round else stake
    paid = I.PAID_BEFORE["Riskia"] + I.DEAL_CASH
    exits = {}
    for k, mult in I.REVENUE_MULTIPLE.items():
        value = rev_2031 * mult
        exits[k] = {"multiple": mult, "company_value": value, "stake_value": value * stake_after,
                    "times_money": value * stake_after / paid}

    return {"with_round": with_round, "rows": rows, "years": years, "quarters": quarters,
            "lowest_cash": {"month": low["label"], "cash": low["cash"]},
            "first_profitable_month": first_positive, "profitable_for_good_from": breakeven,
            "founders_paid_from": founders_from, "riskia_side": {"paid": paid, "stake": stake,
                                                                  "stake_after_round": stake_after},
            "exit": exits}
