"""
Teman Journey growth scenarios on top of model v13 (new rate card, base plan).

Run:  python analysis/growth_scenarios.py   (from financial-estimations/, needs openpyxl)
Reads TemanJourney_Model_2027_v13.xlsx (cached values) and changes nothing.
Two views:
  ladder()  - full-year run-rate on 2027 plan volume: each step adds a cost or a volume gate.
  phased()  - calendar 2027, month by month, with hires switched on at their gate month.
Every figure marked ESTIMATE below must be confirmed before it is used outside this deck.
"""
from pathlib import Path
import openpyxl

XLSX = Path(__file__).resolve().parent.parent / "TemanJourney_Model_2027_v13.xlsx"

UMP = 5.73        # ESTIMATE: Jakarta minimum wage 2026, Rp m / month (confirm)
ON = 1.18         # ESTIMATE: employer BPJS + THR on top of salary
CUR = dict(admin=3.0, socmed=1.5, leadmkt=2.5, psych=2.3, interns=2.0)        # Inputs sheet, 2027 budget
FAIR = dict(admin=UMP * ON, socmed=UMP * ON, leadmkt=7.0 * ON, psych=9.0 * ON, interns=2.0)  # ESTIMATE gross pay
ADS = 3.0         # extra IG ads and content, Rp m / month (founder budget choice)
EAP_R, EAP_C = 15.0, 5.0                      # Inputs: p_eap, c_eap
BANK_R, BANK_C, PMFEE = 51.0, 15.0, 2.5       # Rate_Card state bank Signature; cost 12 + Coreitera 3; ESTIMATE bench PM fee
COF, CTO, ADM2 = 8.0, 10.0, UMP * ON          # ESTIMATE cash pay: BD co-founder, CTO, second admin
TAX = 0.005


def _rows():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    def row(ws, label):
        for r in wb[ws].iter_rows():
            if r[0].value == label:
                return [c.value for c in r[4:19]]
        raise KeyError(label)
    return dict(R=row("PnL", "Total revenue"), D=row("PnL", "Total delivery cost"), F=row("PnL", "Total fixed cost"),
                T=row("PnL", "Income tax (final tax on revenue)"), cash=row("CashFlow", "Cash at end of month"),
                b2b=row("Volume", "Total B2B events (Flagship + KPW/LPS + KCP)"), NP=row("PnL", "Net profit"),
                fac=row("Capacity_OPEX", "Facilitators needed this month (placements / max events each)"), BI=row("PnL", "Revenue from BI (Flagship + BI share of KPW/LPS)"))


ROWS = _rows()
Y27 = slice(3, 15)
MONTHS = ["Oct-26", "Nov-26", "Dec-26"] + [m + "-27" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()]
BASE = dict(R=sum(ROWS["R"][Y27]), GP=sum(ROWS["R"][Y27]) - sum(ROWS["D"][Y27]), F=sum(ROWS["F"][Y27]), T=sum(ROWS["T"][Y27]))
BASE["N"] = BASE["GP"] - BASE["F"] - BASE["T"]
FAIR_DELTA = sum(FAIR.values()) - sum(CUR.values())          # Rp m / month


def scen(steps, extra_eap=2, groups=10):
    s = dict(BASE)
    if "fair" in steps: s["F"] += 12 * (FAIR_DELTA + ADS)
    if "eap" in steps: s["R"] += extra_eap * 12 * EAP_R; s["GP"] += extra_eap * 12 * (EAP_R - EAP_C)
    if "cof" in steps: s["F"] += 12 * COF
    if "bank" in steps: s["R"] += groups * BANK_R; s["GP"] += groups * (BANK_R - BANK_C - PMFEE)
    if "cto" in steps: s["F"] += 12 * CTO
    if "adm2" in steps: s["F"] += 12 * ADM2
    s["T"] = TAX * s["R"]; s["N"] = s["GP"] - s["F"] - s["T"]; s["m"] = s["N"] / s["R"]
    return s


LADDER = [("v13 rate card, today's pay", []),
          ("+ fair pay, psychologist, IG ads", ["fair"]),
          ("+ 2 more EAP contracts", ["fair", "eap"]),
          ("+ BD co-founder, no new deal", ["fair", "eap", "cof"]),
          ("+ co-founder brings 10 bank group-days", ["fair", "eap", "cof", "bank"]),
          ("+ CTO paid by Teman Journey", ["fair", "eap", "cof", "bank", "cto"]),
          ("+ second admin", ["fair", "eap", "cof", "bank", "cto", "adm2"])]
OLD = dict(R=1208.0, N=151.6, m=151.6 / 1208.0)    # Summary sheet snapshot, old prices base


def phased(groups_q4=1, cof_from=6, eap2_from=3, eap3_from=6):
    """Calendar 2027. Month index 0 = Jan-27. Fair pay for existing staff and IG ads from Jan; psychologist with EAP #2."""
    dR, dN, inc = [0.0] * 15, [0.0] * 15, [0.0] * 15
    for i in range(3, 15):
        mi = i - 3
        fixed = FAIR_DELTA + ADS - (FAIR["psych"] - CUR["psych"])
        if mi >= eap2_from: fixed += FAIR["psych"] - CUR["psych"]
        if mi >= cof_from: fixed += COF
        eap = (mi >= eap2_from) + (mi >= eap3_from)
        g = groups_q4 if mi >= 9 else 0
        dR[i] = eap * EAP_R + g * BANK_R
        dN[i] = eap * (EAP_R - EAP_C) + g * (BANK_R - BANK_C - PMFEE) - fixed - TAX * dR[i]
        if i + 1 < 15: inc[i + 1] += eap * EAP_R                      # EAP paid the next month
        for lag, sh in ((0, .2), (1, .5), (2, .3)):                    # bank collection mix (Inputs)
            if i + lag < 15: inc[i + lag] += g * BANK_R * sh
    c, cash = 0.0, []
    for i in range(15):
        c += inc[i] - (dR[i] - dN[i]); cash.append(ROWS["cash"][i] + c)
    R = BASE["R"] + sum(dR[3:]); N = BASE["N"] + sum(dN[3:])
    lo = min(cash)
    net = [ROWS['NP'][i] + dN[i] for i in range(15)]
    return dict(R=R, N=N, m=N / R, low=lo, low_month=MONTHS[cash.index(lo)], end=cash[-1], cash=cash, net=net)


def eap_curve(n_max=4):
    out = []
    for n in range(n_max + 1):
        s = scen(["fair", "eap"], extra_eap=n); out.append((n, s["m"]))
    return out


if __name__ == "__main__":
    print("v13 base", {k: round(v, 1) for k, v in BASE.items()}, f"{BASE['N']/BASE['R']:.1%}")
    print("fair pay delta / month", round(FAIR_DELTA, 2), "+ ads", ADS)
    for name, st in LADDER:
        s = scen(st); print(f"{name:42s} R {s['R']:7.1f}  F {s['F']:7.1f}  N {s['N']:6.1f}  {s['m']:.1%}")
    print("EAP curve", [(n, f"{m:.1%}") for n, m in eap_curve()])
    for a in [(1, 6), (0, 6), (0, 99)]:
        p = phased(*a); print("phased", a, f"R {p['R']:.1f} N {p['N']:.1f} {p['m']:.1%} low {p['low']:.1f} {p['low_month']} end {p['end']:.1f}")
    print("PM load: B2B events per month", dict(zip(MONTHS, ROWS["b2b"])))


def price_only_uplift(target=0.15):
    """Rise in all event prices (Flagship, KPW/LPS, bank) needed to hit the target margin after fair pay, with no new EAP."""
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ev = 0.0
    for r in wb["PnL"].iter_rows():
        if r[0].value in ("Flagship regulator ToT", "KPW / LPS regional ToT", "Bank KCP events") and r[1].value == "Rp m":
            ev += sum(c.value for c in r[7:19])
    s = scen(["fair"])
    x = (target * s["R"] - s["N"]) / (ev * (1 - TAX) - target * ev)
    return x, ev
