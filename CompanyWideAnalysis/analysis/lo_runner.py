"""
Recalculate the v16 model with LibreOffice (headless) and read the results for several input settings.

Run from financial-estimations/CompanyWideAnalysis:   python analysis/lo_runner.py [--save]
Same job as excel_runner.py (Excel for Mac), for machines with LibreOffice instead of Excel. LibreOffice gives the
same numbers as Excel on v15 (checked to 12 digits on revenue, net profit and cash). Each config writes a few cells
into a copy (openpyxl keeps the formulas), LibreOffice recalculates it, and the values are read back.
Writes output/v16_runs.json. With --save, TemanJourney_Model_2027_v16.xlsx is replaced by the recalculated copy of
the RECOMMENDED config, so the saved file carries values (needed by viewers that do not calculate).
"""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "TemanJourney_Model_2027_v16.xlsx"
OUTJ = ROOT / "output" / "v16_runs.json"
MON = ["Oct-26", "Nov-26", "Dec-26"] + [m + "-27" for m in "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()]
COLS = "EFGHIJKLMNOPQRS"

KPW_PLAN = [0, 1, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, 2, 1, 0]          # Channels row 15 (default)
SCH_PLAN = [0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0]          # Channels row 16 (default)


def row(sheet, r, vals):
    return {f"{sheet}!{c}{r}": v for c, v in zip(COLS, vals)}


# Keys in square brackets are Inputs keys (column A); plain refs are cells. v15 cells keep their v15 meaning.
BASE = {"C6": 1, "C101": 1, "C105": 1, "C107": "Feb-27", "C108": "Jul-27", "C111": "Oct-27", "C112": 1, "C117": 0,
        "[v16_on]": 1, "[b2c_p_new]": "=B2C_Pricing!$B$18", "[b2c_f_new]": "=B2C_Pricing!$B$21", "[b2c_pack]": 0,
        "[ev_p_new]": 4.144, "[ev_c_new]": 2.73, "[commit_s]": 0.4, "[commit_d]": 0.05, "[renew_r]": 0.3, "[upsell]": 0.25,
        "[lps_rate]": 0.34, "B2C!E8": 1.25, "B2C!F8": 1.4, "B2C!G8": 1.4,
        "Channels!B6": 14.0, "Channels!B7": 23.0, "Channels!B8": 20.0, "Channels!B9": 20.0,
        "Pay_Plan!J11": 25, "Pay_Plan!K11": 25, "Pay_Plan!H12": 3.5, "Pay_Plan!I12": 3.5,
        **{f"Pay_Plan!{c}12": 3.5 for c in "JKLMNOPQRS"},
        **row("Channels", 15, KPW_PLAN), **row("Channels", 16, SCH_PLAN)}
CONFIGS = {
    "v15_check": {"[v16_on]": 0},                                   # must equal v15: 2,040 sales, 26.1% margin
    "plan": {},
    "plan_b2c_old_price": {"[b2c_p_new]": 1.0, "[b2c_f_new]": 0.4},  # v16 with v15's Rp1m / 40% counseling
    "plan_b2c_old_events": {"[ev_p_new]": 20, "[ev_c_new]": 8},
    "b2c_only_new_price": {"[v16_on]": 0, "C14": "=B2C_Pricing!$B$18", "C21": "=B2C_Pricing!$B$21"},  # v15 + new B2C price only
    "b2c_pack": {"[b2c_pack]": 0.5, "B2C!E8": 1.6, "B2C!F8": 1.9, "B2C!G8": 1.9},
    "no_kpw": row("Channels", 15, [0] * 15),
    "no_schools": row("Channels", 16, [0] * 15),
    "no_scheme": {"[commit_s]": 0},
    "no_extra_ads": {f"Pay_Plan!{c}12": 0 for c in "HIJKLMNOPQRS"},   # ads do not pay at the new prices (LTV/CAC < 1)
    "scheme_10pct": {"[commit_d]": 0.10},
    "kpw_half": row("Channels", 15, [0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 1, 0, 1, 1, 0]),
    "no_lps": {"[lps_rate]": 0},
    "plan_dp30": {"C117": 0.3},
    # Recommended pricing: standard KPw packages at a 30% margin (Sapa stays the intro); Commit-2 then pays
    "kpw_repriced": {"Channels!B6": 15.5, "Channels!B7": 24, "Channels!B8": 24, "Channels!B9": 24},
    "kpw_repriced_no_scheme": {"Channels!B6": 15.5, "Channels!B7": 24, "Channels!B8": 24, "Channels!B9": 24, "[commit_s]": 0},
    "kpw_repriced_10pct": {"Channels!B6": 15.5, "Channels!B7": 24, "Channels!B8": 24, "Channels!B9": 24, "[commit_d]": 0.10},
    # Cash guard (README): second founder's pay from May, extra ads from March, 30% Flagship down payment
    "plan_guards": {"Pay_Plan!J11": 10, "Pay_Plan!K11": 10, "Pay_Plan!H12": 0, "Pay_Plan!I12": 0, "C117": 0.3},
    "recommended": {"Channels!B6": 15.5, "Channels!B7": 24, "Channels!B8": 24, "Channels!B9": 24,
                    "Pay_Plan!J11": 10, "Pay_Plan!K11": 10, "Pay_Plan!H12": 0, "Pay_Plan!I12": 0, "C117": 0.3},
    "conservative": {"C6": 3, "C112": 0, "[upsell]": 0.1, "[renew_r]": 0.2, "[commit_s]": 0.2,
                     **row("Channels", 15, [0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 1, 0, 1, 1, 0]),
                     **row("Channels", 16, [0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0])},
    "aggressive": {"C111": "Jul-27", "C112": 2, "[upsell]": 0.35,
                   **row("Channels", 15, [0, 1, 0, 1, 1, 1, 2, 1, 1, 2, 2, 2, 2, 2, 1]),
                   **row("Channels", 16, [0, 0, 0, 0, 1, 1, 0, 0, 1, 1, 2, 2, 2, 1, 0])},
}
RECOMMENDED = "plan"
ROWS = {"sales": 5, "delivery": 6, "fixed": 7, "tax": 8, "cost": 9, "net": 10, "cash_in": 11, "cash_out": 12,
        "cash": 13, "floor": 14, "ok": 15}
EXTRA = {  # FY 2027 values (column D) read for the write-up
    "seg": ("Segments", [5, 6, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 19, 20, 21]),
    "b2c": ("B2C", [26, 28, 30, 31, 35, 36, 37, 39, 43]),
    "pnl16": ("PnL", [9, 10, 11, 12, 19, 20, 21, 23, 31, 35, 36, 40, 41, 47, 48, 49, 50, 51, 53, 54, 55, 56, 57]),
    "vol16": ("Volume", [26, 27, 28, 39, 40, 41, 43, 44, 45, 46, 47, 48, 49]),
}


def keymap(wb):
    ws = wb["Inputs"]
    return {ws.cell(r, 1).value: f"C{r}" for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str)}


def apply(wb, cfg):
    km = keymap(wb)
    for ref, v in cfg.items():
        if ref.startswith("["):
            sheet, cell = "Inputs", km[ref[1:-1]]
        else:
            sheet, cell = ref.split("!") if "!" in ref else ("Inputs", ref)
        wb[sheet][cell] = v


def recalc(src, outdir):
    subprocess.run(["soffice", "--headless", "--norestore", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", str(outdir), str(src)], check=True, capture_output=True, timeout=300)
    return outdir / src.name


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    O = wb["Outlook"]
    d = {k: [float(O.cell(r, c).value or 0) for c in range(2, 17)] for k, r in ROWS.items()}
    y = slice(3, 15)
    R, N = sum(d["sales"][y]), sum(d["net"][y])
    lo = min(d["cash"])
    d["fy27"] = dict(R=R, N=N, m=N / R if R else 0, cost=sum(d["cost"][y]), low=lo, low_idx=d["cash"].index(lo),
                     end=d["cash"][-1], below_floor=int(sum(1 for v in d["ok"] if v < 1)))
    for k, (sh, rows) in EXTRA.items():
        d[k] = {f"{sh}!{r}": [wb[sh].cell(r, 1).value, wb[sh].cell(r, 3).value, wb[sh].cell(r, 4).value] for r in rows}
    return d


def main():
    save = "--save" in sys.argv
    only = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = json.loads(OUTJ.read_text()) if (only and OUTJ.exists()) else {}
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        for name, cfg in CONFIGS.items():
            if only and name not in only:
                continue
            wb = openpyxl.load_workbook(BOOK)
            apply(wb, {**BASE, **cfg})
            src = tmp / f"{name}.xlsx"
            wb.save(src)
            (tmp / "calc").mkdir(exist_ok=True)
            res = recalc(src, tmp / "calc")
            out[name] = read(res)
            f = out[name]["fy27"]
            print(f"{name:22s} sales {f['R']:7.0f}  net {f['N']:6.0f}  {f['m']:6.1%}  low cash {f['low']:6.0f} "
                  f"({MON[f['low_idx']]})  end {f['end']:5.0f}  months below floor {f['below_floor']}", flush=True)
            if save and name == RECOMMENDED:
                shutil.copy(res, BOOK)
    OUTJ.parent.mkdir(exist_ok=True)
    OUTJ.write_text(json.dumps(out, indent=1))


if __name__ == "__main__":
    main()
