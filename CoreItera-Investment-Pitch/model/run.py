"""Run the CoreItera model: plan + slow case. Writes ../output/results.json and ../output/tables.md.

    python3 -I run.py
"""
import copy, json, os, sys, types
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import inputs as BASE
import engine

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
DIVIDEND_SHARE = 0.30        # E  share of yearly profit paid out from 2030
DIVIDEND_FROM = 2030


def slow_case():
    """Half the bureaus, half the solo seats, copilot adoption 30% lower."""
    I = types.SimpleNamespace(**{k: copy.deepcopy(v) for k, v in vars(BASE).items() if k.isupper()})
    I.BUREAU_TARGETS = {k: v * 0.5 for k, v in BASE.BUREAU_TARGETS.items()}
    I.SOLO_SEATS = {k: v * 0.5 for k, v in BASE.SOLO_SEATS.items()}
    I.COPILOT_ADOPTION = {k: v * 0.7 for k, v in BASE.COPILOT_ADOPTION.items()}
    return I


def slim(s):
    return {k: v for k, v in s.items() if k != "rows"} | {
        "monthly": [{k: (round(v, 3) if isinstance(v, float) else v) for k, v in r.items() if k != "ym"}
                    for r in s["rows"]]}


def dividends(s):
    stake = s["riskia_side"]["stake_after_round"]
    return {y: s["years"][y]["profit"] * DIVIDEND_SHARE * stake for y in s["years"] if y >= DIVIDEND_FROM}


def rp(x, d=1):
    return f"Rp{x:,.{d}f}M"


def main():
    plan = engine.run(BASE, BASE.USE_NEXT_ROUND)
    slow = engine.run(slow_case(), BASE.USE_NEXT_ROUND)
    res = {"plan": slim(plan), "slow": slim(slow),
           "dividends_riskia_side": {"plan": dividends(plan), "slow": dividends(slow)},
           "dividend_share": DIVIDEND_SHARE}
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, "results.json"), "w") as f:
        json.dump(res, f, indent=1, default=str)

    L = ["# CoreItera model results (Rp million)", ""]
    for name, s in (("Plan", plan), ("Slow case (half the bureaus)", slow)):
        L += [f"## {name}", "", f"- Lowest cash: {rp(s['lowest_cash']['cash'])} in {s['lowest_cash']['month']}",
              f"- First month with profit: {s['first_profitable_month']}; profitable every month from {s['profitable_for_good_from']}",
              f"- Founders paid from: {s['founders_paid_from']}", "",
              "| Year | Bureaus (Dec) | Copilot seats (Dec) | Revenue | Plan | Payment fees | Copilot | Variable cost | Fixed cost | Profit after tax | Margin | Cash (Dec) |",
              "|---|---|---|---|---|---|---|---|---|---|---|---|"]
        for y, v in s["years"].items():
            L.append(f"| {y} | {v['bureaus_end']:.0f} | {v['seats_end']:.0f} | {v['revenue']:,.1f} | {v['plan']:,.1f} | {v['fees']:,.1f} | "
                     f"{v['copilot']:,.1f} | {v['variable']:,.1f} | {v['fixed']:,.1f} | {v['profit']:,.1f} | "
                     f"{100 * v['profit'] / v['revenue']:.0f}% | {v['cash_end']:,.1f} |")
        L += ["", "| Quarter | Revenue | Costs | New money | Net cash flow | Cash at end |", "|---|---|---|---|---|---|"]
        for q in s["quarters"]:
            L.append(f"| {q['quarter']} | {q['revenue']:,.1f} | {q['costs']:,.1f} | {q['investment']:,.1f} | {q['net']:,.1f} | {q['cash_end']:,.1f} |")
        L += ["", "| Exit value (x 2031 revenue) | Company value | Riskia + TJ stake | Times the Rp50M |", "|---|---|---|---|"]
        for k, e in s["exit"].items():
            L.append(f"| {e['multiple']:.0f}x | {e['company_value']:,.0f} | {e['stake_value']:,.0f} | {e['times_money']:.0f}x |")
        L.append("")
    L += ["## Monthly, first 18 months (plan)", "", "| Month | Paying bureaus | Seats | Revenue | Costs | Cash |", "|---|---|---|---|---|---|"]
    for r in plan["rows"][:18]:
        L.append(f"| {r['label']} | {r['paying']:.1f} | {r['seats']:.0f} | {r['revenue']:.1f} | "
                 f"{r['variable'] + r['fixed'] + r['tax'] + r['one_offs']:.1f} | {r['cash']:.1f} |")
    L += ["", "## Dividends to Riskia's side (30% of profit paid out)", ""]
    for name, d in res["dividends_riskia_side"].items():
        L.append(f"- {name}: " + ", ".join(f"{y} {rp(v, 0)}" for y, v in d.items()))
    with open(os.path.join(OUT, "tables.md"), "w") as f:
        f.write("\n".join(L) + "\n")
    print("\n".join(L))


if __name__ == "__main__":
    main()
