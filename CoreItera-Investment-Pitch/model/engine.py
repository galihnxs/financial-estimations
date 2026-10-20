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


def yearly_tax(I, years_rev, years_pbt):
    """Company income tax per year (Rp M), with the 0.5% final scheme, Pasal 31E and loss carry-forward."""
    out, losses = {}, []          # losses: [year, amount left]
    for y in sorted(years_rev):
        T, P = years_rev[y], years_pbt[y]
        if y <= I.FINAL_TAX_LAST_YEAR and T <= I.FINAL_TAX_TURNOVER_LIMIT:
            out[y] = {"tax": T * I.FINAL_TAX_RATE, "rule": "0.5% final", "rate": I.FINAL_TAX_RATE}
            continue
        losses = [l for l in losses if y - l[0] <= I.LOSS_CARRY_YEARS]
        if P <= 0:
            losses.append([y, -P])
            out[y] = {"tax": 0.0, "rule": "loss, no tax", "rate": 0.0}
            continue
        taxable = P
        for l in losses:
            use = min(l[1], taxable)
            l[1] -= use
            taxable -= use
        losses = [l for l in losses if l[1] > 0]
        if T <= I.CORP_TAX_CUT_LIMIT:
            share = min(1.0, I.CORP_TAX_CUT_TURNOVER / T) if T > 0 else 1.0
            rate = share * I.CORP_TAX_RATE / 2 + (1 - share) * I.CORP_TAX_RATE
            rule = "22% with Pasal 31E cut"
        else:
            rate, rule = I.CORP_TAX_RATE, "22%"
        out[y] = {"tax": taxable * rate, "rule": rule, "rate": rate, "taxable": taxable}
    return out


def run(I, with_round=False, _tax=None):
    if _tax is None:
        first = run(I, with_round, _tax={})
        rows0 = first["rows"]
        years = sorted({r["ym"][0] for r in rows0})
        rev = {y: sum(r["revenue"] for r in rows0 if r["ym"][0] == y) for y in years}
        pbt = {y: sum(r["before_tax"] for r in rows0 if r["ym"][0] == y) for y in years}
        tax_plan = yearly_tax(I, rev, pbt)
        res = run(I, with_round, _tax=tax_plan)
        res["tax_plan"] = tax_plan
        return res
    rows = []
    cash = I.CASH_TODAY
    bureaus_prev = 0.0
    founders_on = False
    prev_revenue = 0.0
    prev_copilot = 0.0
    acct_on = clin_on = sales_on = office_on = False
    engineers = 0
    staff = {}
    cohorts = []
    prev_pbf_after = 0.0
    prev_pbf = 0.0
    copilot_on = False
    for ym in months(I.START, I.END):
        y, m = ym
        r = {"ym": ym, "label": f"{y}-{m:02d}"}

        # --- volume ---
        target = interp(getattr(I, 'TARGETS_OVERRIDE', None) or I.BUREAU_TARGETS, ym)
        # cohorts: bureaus in their first EARLY_MONTHS paid months churn faster
        young = sum(cohorts[-I.EARLY_MONTHS:]) if cohorts else 0.0
        old = bureaus_prev - young
        base_churn = interp(I.CHURN_CURVE, ym, before=I.CHURN_CURVE[min(I.CHURN_CURVE)])
        early = base_churn * I.EARLY_CHURN_MULTIPLIER
        churned = young * early + max(0.0, old) * base_churn
        kept = bureaus_prev - churned
        cohorts = [c * (1 - early) for c in cohorts[-I.EARLY_MONTHS:]]
        if ym in I.GROSS_ADDS:
            new = I.GROSS_ADDS[ym]
        else:
            new = max(0.0, target - kept)
        bureaus = kept + new
        cohorts.append(new)
        paying = kept if I.FREE_FIRST_MONTH else bureaus     # new bureaus pay from month 2
        copilot_on = copilot_on or idx(ym) >= idx(I.COPILOT_LAUNCH)
        adoption = interp(I.COPILOT_ADOPTION, ym) if copilot_on else 0.0
        solo = interp(I.SOLO_SEATS, ym) if copilot_on else 0.0
        seats = adoption * I.PSY_PER_BUREAU * paying + solo
        r.update(bureaus=bureaus, new=new, churned=churned, churn_rate=base_churn, paying=paying, seats=seats, solo=solo)

        # --- revenue ---
        if y >= 2026:
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
        acquisition = new * (I.CAC_PER_NEW_BUREAU + I.COMMISSION_FIRST_PAYMENT * I.AVG_PLAN_BY_YEAR.get(y, 1.0))
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
        def by_year(d):
            return d.get(y, d[max(k for k in d if k <= y)]) if y >= min(d) else 0.0
        def heads(role):
            plan = I.TEAM_PLAN[role]
            if y < 2027:
                return 0
            if y == 2027:
                return plan[(2027, (m - 1) // 3 + 1)]
            return plan[y]
        budget = prev_pbf_after                 # last month's profit after the team we already have
        for r_ in I.HIRE_ORDER:
            want = heads(r_)
            have = staff.get(r_, 0)
            pay_ = by_year(I.TEAM_PAY[r_])
            if want > have and budget >= pay_:
                have += 1                       # at most one hire a month per role, only if profit pays the salary
                budget -= pay_
            staff[r_] = have
        team = {r_: staff[r_] * by_year(I.TEAM_PAY[r_]) for r_ in I.TEAM_PLAN}
        cs, sales, quality = team["Support"], team["Sales"], team["Quality"]
        engineer = team["Engineers"]
        headcount = dict(staff)
        if prev_revenue >= I.OFFICE_TRIGGER:
            office_on = True
        office = I.OFFICE if office_on else 0.0
        abroad = I.ABROAD_COST if idx(ym) >= idx(I.ABROAD_FROM) else 0.0
        ga = by_year(I.GA_SHARE_OF_REVENUE) * prev_revenue
        fixed_before_pay = base + accountant + clinical + workshops + cs + sales + engineer + quality + office + ga + abroad
        r.update(ai=ai, server=server, agent=agent, onboarding=onboarding, acquisition=acquisition, variable=variable, base=base,
                 accountant=accountant, clinical=clinical, workshops=workshops,
                 cs=cs, sales=sales, office=office, ga=ga, abroad=abroad)

        # founders and Zaki: paid only from profit (Galih's rule)
        pbf = revenue - variable - fixed_before_pay
        low_pay = len(I.PAID_PEOPLE_2027) * I.FOUNDER_PAY_2027_MONTH
        full_pay = sum(I.FOUNDER_PAY_YEAR.values()) / 12 + by_year(I.CORE_TEAM_PAY_MONTH)
        if y >= 2028 and pbf >= full_pay:
            founders, core = sum(I.FOUNDER_PAY_YEAR.values()) / 12, by_year(I.CORE_TEAM_PAY_MONTH)
        elif y >= 2027 and pbf >= low_pay:
            founders, core = 3 * I.FOUNDER_PAY_2027_MONTH, I.FOUNDER_PAY_2027_MONTH
        else:
            founders, core = 0.0, 0.0
        engineer = engineer + core
        fixed = fixed_before_pay + founders + core
        r.update(founders=founders, engineer=engineer, quality=quality, fixed=fixed, headcount=headcount)
        prev_pbf_after = pbf - founders - core

        # --- profit and tax ---
        before_tax = revenue - variable - fixed
        n_months = 3 if y == 2026 else 12           # the model starts in October 2026
        tax = _tax.get(y, {}).get("tax", 0.0) / n_months   # booked evenly through the year (like monthly PPh 25 instalments)
        profit = before_tax - tax
        r.update(gross=revenue - variable, before_tax=before_tax, tax=tax, profit=profit)

        # --- cash ---
        deal = I.DEAL_CASH if ym == I.DEAL_MONTH else 0.0
        one_offs = sum(I.DEAL_ONE_OFFS.values()) if ym == I.DEAL_MONTH else 0.0
        round_in = I.NEXT_ROUND["amount"] if (with_round and ym == I.NEXT_ROUND["month"]) else 0.0
        # tax CASH (Indonesian practice): monthly PPh 25 instalments this year = last year's tax / 12;
        # the rest of last year's tax (PPh 29) is paid in April. The 0.5% final tax is paid in the same month.
        t_prev = _tax.get(y - 1, {}).get("tax", 0.0)
        t_prev2 = _tax.get(y - 2, {}).get("tax", 0.0)
        if _tax.get(y, {}).get("rule") == "0.5% final":
            cash_tax = tax
        else:
            final_prev = _tax.get(y - 1, {}).get("rule") == "0.5% final"
            cash_tax = 0.0 if final_prev else t_prev / 12
            if m == 4 and not final_prev:
                cash_tax += max(0.0, t_prev - (0.0 if _tax.get(y - 2, {}).get("rule") == "0.5% final" else t_prev2))
        cash = cash + before_tax - cash_tax + deal + round_in - one_offs
        r.update(deal=deal, one_offs=one_offs, round_in=round_in, cash=cash, cash_tax=cash_tax)
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
                     "new", "churned", "clinical", "workshops", "founders", "cs", "sales", "engineer", "quality", "office", "abroad", "ga", "acquisition", "onboarding", "agent", "ai", "server", "fixed",
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
    full_pay_from = next((r["label"] for r in rows if r["founders"] > 3 * I.FOUNDER_PAY_2027_MONTH + 0.01), None)

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
            "founders_paid_from": founders_from, "full_pay_from": full_pay_from, "riskia_side": {"paid": paid, "stake": stake,
                                                                  "stake_after_round": stake_after},
            "exit": exits}
