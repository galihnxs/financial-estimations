"""
Recalculate the v14 workbook in Microsoft Excel and read monthly results for several input settings.

Run from financial-estimations/:   python analysis/excel_runner.py [--save]
openpyxl cannot calculate formulas and there is no LibreOffice here, so Excel does the maths through AppleScript.
Each config sets a few Inputs cells, recalculates, and reads Outlook rows 5 to 15 (Oct-26 .. Dec-27).
Writes output/v14_runs.json. With --save, the workbook is saved with the RECOMMENDED config, values calculated.
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "TemanJourney_Model_2027_v14.xlsx"
SANDBOX = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents"
WORK = SANDBOX / "tj_v14.xlsx"
OUTJ = ROOT / "output" / "v14_runs.json"

BASE = {"C6": 1, "C51": "Jan-27", "C101": 1, "C105": 1, "C119": 1}
CONFIGS = {
    "v13_timing_today_pay": {"C101": 0, "C105": 0},
    "real_timing_today_pay": {"C105": 0},
    "plan_founders_jan": {"C51": "Jan-27"},
    "plan_founders_mar": {"C51": "Mar-27"},
    "plan_founders_may": {"C51": "May-27"},
    "plan_founders_jul": {"C51": "Jul-27"},
    "plan_founders_may_no_bank": {"C51": "May-27", "C119": 0},
    "plan_founders_may_bi_cut": {"C51": "May-27", "C6": 3},
}
RECOMMENDED = "plan_founders_may"
ROWS = {"sales": 5, "delivery": 6, "fixed": 7, "tax": 8, "cost": 9, "net": 10, "cash_in": 11, "cash_out": 12,
        "cash": 13, "floor": 14, "ok": 15}


def q(v):
    return f'"{v}"' if isinstance(v, str) else str(v)


def script(configs, save):
    lines = ['tell application "Microsoft Excel"', f'open POSIX file "{WORK}"', "delay 2",
             "set wb to active workbook", 'set out to ""']
    for name, cfg in configs.items():
        full = {**BASE, **cfg}
        for cell, v in full.items():
            lines.append(f'set value of range "{cell}" of worksheet "Inputs" of wb to {q(v)}')
        lines += ["calculate full",
                  f'set out to out & "@@{name}" & linefeed',
                  'set vals to value of range "B5:P15" of worksheet "Outlook" of wb',
                  "repeat with r in vals", "set AppleScript's text item delimiters to \";\"",
                  "set out to out & (r as text) & linefeed", "end repeat"]
    if save:
        for cell, v in {**BASE, **CONFIGS[RECOMMENDED]}.items():
            lines.append(f'set value of range "{cell}" of worksheet "Inputs" of wb to {q(v)}')
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
        y = slice(3, 15)
        R, Nn = sum(d["sales"][y]), sum(d["net"][y])
        lo = min(d["cash"])
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
