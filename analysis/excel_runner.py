"""
Recalculate the model workbook (v15) in Microsoft Excel and read monthly results for several input settings.

Run from financial-estimations/:   python analysis/excel_runner.py [--save]
openpyxl cannot calculate formulas and there is no LibreOffice here, so Excel does the maths through AppleScript.
Each config sets a few Inputs cells, recalculates, and reads Outlook rows 5 to 15 (Oct-26 .. Dec-27).
Writes output/model_runs.json. Also reads the running cost by group (Pay_Plan rows 19 to 24) and the operations load (Capacity_OPEX rows 43 to 46). With --save, the workbook is saved with the RECOMMENDED config, values calculated.
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "TemanJourney_Model_2027_v15.xlsx"
SANDBOX = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents"
WORK = SANDBOX / "tj_model.xlsx"
OUTJ = ROOT / "output" / "model_runs.json"

# Inputs cells: C6 scenario, C51 v13 founder start, C101 timing switch, C105 growth plan, C107/C108 EAP 2/3 start,
# C112 bank days per month, C117 Flagship down payment. Pay steps live on the Pay_Plan sheet and are not changed here.
BASE = {"C6": 1, "C51": "Jan-27", "C101": 1, "C105": 1, "C107": "Feb-27", "C108": "Jul-27", "C111": "Oct-27", "C112": 1}
CONFIGS = {
    "v13_timing_today_pay": {"C101": 0, "C105": 0},
    "real_timing_today_pay": {"C105": 0},
    "plan": {},
    "plan_no_bank": {"C112": 0},
    "plan_bi_cut": {"C6": 3},
    "plan_eap_late": {"C107": "Apr-27", "C108": "Sep-27"},
    "plan_founders_full_may": {"Pay_Plan!J11": 10, "Pay_Plan!K11": 10},
    "plan_ads_from_mar": {"Pay_Plan!H12": 0, "Pay_Plan!I12": 0},
    "plan_both_guards": {"Pay_Plan!J11": 10, "Pay_Plan!K11": 10, "Pay_Plan!H12": 0, "Pay_Plan!I12": 0},
    # COO review: EAP 2 with a bank starts later; 3 regional events moved from Nov-Dec into Jul-Sep; 30% Flagship down payment
    "plan_eap2_bank": {"C107": "Apr-27"},
    "plan_q4_smooth": {"Volume!N16": 2, "Volume!O16": 2, "Volume!P16": 3, "Volume!R16": 2, "Volume!S16": 3},
    "plan_dp30": {"C117": 0.3},
    # Three named estimates for the deck: conservative / current (= plan) / aggressive
    "conservative": {"C6": 3, "C112": 0},
    "aggressive": {"C111": "Jul-27", "C112": 2},
    # Cash review: savings step 4 (30% down payment on the March head-office program) on top of the plan
    "save_now": {"C117": 0.3},
    "plan_all_levers": {"Pay_Plan!J11": 10, "Pay_Plan!K11": 10, "Pay_Plan!H12": 0, "Pay_Plan!I12": 0,
                        "Volume!N16": 2, "Volume!O16": 2, "Volume!P16": 3, "Volume!R16": 2, "Volume!S16": 3, "C117": 0.3},
    "plan_cfo_coo": {"Volume!N16": 2, "Volume!O16": 2, "Volume!P16": 3, "Volume!R16": 2, "Volume!S16": 3, "C117": 0.3},
}
PAY_DEFAULT = {"Pay_Plan!J11": 25, "Pay_Plan!K11": 25, "Pay_Plan!H12": 3.5, "Pay_Plan!I12": 3.5,   # reset after each config
               "Volume!N16": 1, "Volume!O16": 1, "Volume!P16": 2, "Volume!R16": 3, "Volume!S16": 5, "C117": 0}
RECOMMENDED = "plan"
LOAD = ["projects", "pm_needed", "pm_load", "trainer_load"]          # Capacity_OPEX rows 43-46 (COO review)
GROUPS = ["founders", "team", "marketing", "office", "sales_support", "running_total"]
ROWS = {"sales": 5, "delivery": 6, "fixed": 7, "tax": 8, "cost": 9, "net": 10, "cash_in": 11, "cash_out": 12,
        "cash": 13, "floor": 14, "ok": 15}


def q(v):
    return f'"{v}"' if isinstance(v, str) else str(v)


def script(configs, save):
    lines = ['tell application "Microsoft Excel"', f'open POSIX file "{WORK}"', "delay 2",
             "set wb to active workbook", 'set out to ""']
    for name, cfg in configs.items():
        full = {**BASE, **PAY_DEFAULT, **cfg}
        for cell, v in full.items():
            sheet, ref = cell.split("!") if "!" in cell else ("Inputs", cell)
            lines.append(f'set value of range "{ref}" of worksheet "{sheet}" of wb to {q(v)}')
        lines += ["calculate full",
                  f'set out to out & "@@{name}" & linefeed',
                  'set vals to (value of range "B5:P15" of worksheet "Outlook" of wb) & (value of range "E19:S24" of worksheet "Pay_Plan" of wb) & (value of range "E43:S46" of worksheet "Capacity_OPEX" of wb)',
                  "repeat with r in vals", "set AppleScript's text item delimiters to \";\"",
                  "set out to out & (r as text) & linefeed", "end repeat"]
    if save:
        for cell, v in {**BASE, **PAY_DEFAULT, **CONFIGS[RECOMMENDED]}.items():
            sheet, ref = cell.split("!") if "!" in cell else ("Inputs", cell)
            lines.append(f'set value of range "{ref}" of worksheet "{sheet}" of wb to {q(v)}')
        lines += ["calculate full", "save wb"]
    lines += ["close wb saving no", "return out", "end tell"]
    return "\n".join(lines)


def main():
    save = "--save" in sys.argv
    shutil.copy(BOOK, WORK)
    res = subprocess.run(["osascript", "-e", script(CONFIGS, save)], capture_output=True, text=True, timeout=600)
    if res.returncode:
        sys.exit(res.stderr)
    runs, cur = {}, None
    for line in res.stdout.splitlines():
        if line.startswith("@@"):
            cur = line[2:]; runs[cur] = []
        elif line.strip() and cur:
            runs[cur].append([float(x.replace(",", ".")) if x.strip() else 0.0 for x in line.split(";")])  # this Mac writes decimal commas
    out = {}
    for name, rows in runs.items():
        d = {k: rows[r - 5] for k, r in ROWS.items()}
        d.update({g: rows[11 + i] for i, g in enumerate(GROUPS)})
        y = slice(3, 15)
        R, Nn = sum(d["sales"][y]), sum(d["net"][y])
        lo = min(d["cash"])
        d.update({k: rows[17 + i] for i, k in enumerate(LOAD)})
        d["fy27"] = dict(R=R, N=Nn, m=Nn / R if R else 0, cost=sum(d["cost"][y]), low=lo,
                         low_idx=d["cash"].index(lo), end=d["cash"][-1], below_floor=int(sum(1 for v in d["ok"] if v < 1)))
        out[name] = d
    OUTJ.parent.mkdir(exist_ok=True)
    OUTJ.write_text(json.dumps(out, indent=1))
    if save:
        shutil.copy(WORK, BOOK)
    for n, d in out.items():
        f = d["fy27"]
        print(f"{n:30s} sales {f['R']:7.0f}  net {f['N']:6.0f}  {f['m']:6.1%}  low cash {f['low']:6.0f} (m{f['low_idx']})  "
              f"end {f['end']:5.0f}  months below floor {f['below_floor']}")


if __name__ == "__main__":
    main()
