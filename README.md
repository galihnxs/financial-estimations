# Teman Journey: financial estimations

The financial model and growth plan for Teman Journey (PT Indo Medika Resolusi), October 2026 to December 2027.
All money is in Rupiah million (Rp M) unless stated.

## The answer in one paragraph

Two goals: **create real value for our clients**, and **build a business that pays for itself**. Four impact pillars show
it is working: **Clients** (EAP renewals, before/after training scores), **People** (fair pay: minimum wage + BPJS + THR,
raised in steps as sales targets are met), **Company** (net margin of at least 15%, safe cash, less reliance on BI) and
**Community** (710 people in our trainings plus about 1,400 they reach, an estimate of 2 each; event guests; B2C counseling; referral). Then we are ready for investors and new units (Theta, Coreitera). The plan:
- keep the new prices;
- sell 2 more EAP (monthly counseling) contracts;
- hire only when signed work pays for the person.

We use October 2026 to February 2027 to pitch and sign. New work starts in March 2027. Today one founder is paid
Rp10M a month; from March the plan keeps Rp25M a month as room for both founders (the second founder's start is still
to decide). The project manager is paid per project (usually Rp2M). The admin is paid Rp1.8M + Rp17k a client + a weekly bonus
(minimum wage once pay step 2 unlocks). No full-time psychologist in 2027: two paid part-timers (from Jan 2027) plus a pool of 20 per-session psychologists; a
non-psychologist Lead Partnership joins in Oct 2027. **B2C** (sheet `B2C`): followers x booking rate x sessions per client,
set per quarter (3,439 followers on 5 Oct 2026, about 2 paid sessions a month today, 1.4 sessions per client from Q3 2027). Referral (B2C rows 50-59): referral share of new clients per quarter
(target 20% by Q4 2027), Rp100k credit to each side, and the unit economics per client (LTV, cost per client, LTV/CAC).
Instagram ads: Rp1M today, plus Rp3.5M moved from the lead-marketing budget from January (keep only if ads earn
at least 1.7x their cost; target 3x). **Pay rises unlock with sales, not dates:** when average sales of the last 3 months
reach Rp100M / Rp150M / Rp200M (Inputs C123, C126, C127; Pay_Plan rows 27-30), social media, lead marketing and admin
move up one step, and never back down. Result for 2027 (model v15):

| Measure | Value |
|---|---|
| Sales | Rp2,040M |
| Net margin | 26.1% |
| Lowest cash | Rp82M (February 2027) |
| Months below the safety floor | 1 |
| Cash at end of 2027 | Rp524M |

Cash is the weak point. The **guard rule**: if cash at the end of February is under Rp80M, delay the second founder's pay
until May and start the extra ads in March. With the guard rule, a 30% down payment on Flagship programs and 3 regional
events moved from Nov-Dec to Jul-Sep, the lowest cash is Rp107M and the margin 27.9%.

**Selling starts 4 to 6 months before the work** (COO review). Regulators and banks plan next year's budget in the
second half of this year, so BI, EAP 3 and bank days must be pitched in Q4 2026. See the **Roadmap** sheet.

**Three estimates** (deck and `output/model_runs.json`): **conservative** = BI buys 30% less and no bank deal (16.6%);
**current** = our plan (26.1%); **aggressive** = the bank deal starts in July with 2 training days a month (33.5%). The deck never
shows model version numbers.

## Files

| File | What it is |
|---|---|
| `TemanJourney_Model_2027_v15.xlsx` | **The current model.** Monthly sales, costs, profit and cash, with real timing, the pay plan, the growth plan and the roadmap. Open **Outlook** and **Roadmap** first. |
| `TemanJourney_Model_2027_v14.xlsx` | The previous model (fair pay in one step). Kept for comparison. |
| `TemanJourney_Model_2027_v13.xlsx` | The previous model. It assumes sales from October 2026. Kept unchanged for comparison. |
| `TemanJourney_Growth_Recommendation.pptx` / `.pdf` | The growth plan deck: 40 slides. Pyramid: summary first; each part (goal and problem, A. B2B, B. B2C, C. people, D. money, next steps, appendix) opens with a dark slide that asks the question it answers. |
| `projection guide.docx` | Written review of model v13. **Out of date:** its numbers are v13, not v15. |
| `analysis/` | Python scripts that build the model, run it, and build the deck (see below). |
| `output/model_runs.json` | Results of v15 under several settings, calculated by Excel. The deck and the Summary sheet read this file. |

## How the model works (v15)

The model is built bottom-up: number of events × price (or cost). Every result cell is a formula.

| Sheet | What it holds |
|---|---|
| Summary | The answers: profit, cash, delivery capacity, investor questions, and a scenario table. |
| Outlook | Sales, total cost, net profit and cash for each month, with one chart: bars for sales and costs, lines for cash and the safety floor. |
| Roadmap | Worked backward from each start date: lead times by client type, the latest date to start each deal and each hire, status today, the cash clock, and a sales speed log to fill in after each deal. |
| B2C | The B2C funnel per quarter (followers, booking rate, sessions per client), B2B vs B2C revenue, psychologist cover and capacity. Feeds Volume row 28. |
| Pay_Plan | Gross pay for every role, month by month (blue = founders' decisions), BPJS + THR, and the running cost by group. |
| Inputs | Every assumption, marked Known, Your plan or Estimate. Section 11 is timing, 12 the growth plan, 13 operations (COO review). |
| Volume | Events per month. Rows 34 to 36 hold the timing flag, the new EAP contracts and the bank days. |
| Unit_Economics, Rate_Card | Price, cost and profit per event; the new price list. |
| Capacity_OPEX | Trainers needed, hiring triggers and fixed costs. Rows 35 to 40 hold the growth-plan costs; rows 42 to 47 the operations load (projects per month, project managers needed, trainers certified 2 months ahead). |
| PnL, CashFlow | Monthly profit and cash. Cash follows the clients' payment delays (0 to 60 days). |
| To_Fill | Every estimate that still needs a real number. |

**Main switches** on the Inputs sheet:

| Cell | Switch | Our plan |
|---|---|---|
| C101 | No new sales before the start month (1 = yes) | 1 |
| C102 | First month of new work | Mar-27 |
| C105 | Growth plan on (1 = yes) | 1 |
| C107 / C108 | EAP contract 2 / 3 start | Feb-27 / Jul-27 |
| C112 | Bank training days per month (0 = no bank deal) | 1 |
| C114 | Project manager fee per project | Rp2M |
| C116 | Certify trainers this many months before the need | 2 |
| C117 | Down payment on Flagship programs | 0% (test 30%) |
| C6 | Stress test: 1 = base, 3 = BI buys 30% less | 1 |

Pay steps and founder pay live on the **Pay_Plan** sheet. With C101 = 0 and C105 = 0, v15 gives exactly the v13
results (25.2% margin, lowest cash Rp204M). This check proves that v15 only adds to v13.

## How to rebuild everything

The scripts use the conda environment `ind5003` (it has openpyxl and python-pptx) and **Microsoft Excel and
PowerPoint for Mac**. openpyxl cannot calculate formulas, and this Mac has no LibreOffice, so Excel does the maths
through AppleScript.

Run from this folder, in this order:

```bash
PY=/opt/homebrew/Caskroom/miniforge/base/envs/ind5003/bin/python
$PY analysis/build_model.py            # 1. build v15 from v13 (v13 is never changed)
$PY analysis/excel_runner.py --save    # 2. Excel recalculates each setting, writes output/model_runs.json, saves v15
$PY analysis/build_model.py            # 3. rebuild so the Summary scenario table uses the new runs
$PY analysis/excel_runner.py --save    # 4. recalculate and save again
$PY analysis/build_growth_deck.py      # 5. build the deck (reads model_runs.json and the saved v15 values)
```

| Script | Job |
|---|---|
| `analysis/build_model.py` | Copies v13 and adds the timing switch, the growth plan, Pay_Plan, Outlook, Roadmap, the operations rows and new To_Fill rows. |
| `analysis/excel_runner.py` | Opens a copy of v15 in Excel, sets each scenario (Inputs, Pay_Plan or Volume cells), recalculates and reads the monthly results and the operations load. To add a scenario, add a line to `CONFIGS`. |
| `analysis/growth_scenarios.py` | The "normal full year" steps used in the deck (slides 5, 9 and appendix slide 20). It reads v13. |
| `analysis/build_growth_deck.py` | Builds the deck. All numbers come from the sources above; none are typed by hand. |

**Deck checks.** After building the deck, run the text-fit and overlap checks from the DOS5024 project, then the pptx
skill's validator:

```bash
$PY ../NUS-Learning/DOS5024/analysis/deck_fit_check.py TemanJourney_Growth_Recommendation.pptx
$PY ../NUS-Learning/DOS5024/analysis/deck_collision_check.py TemanJourney_Growth_Recommendation.pptx
```

Two fit warnings are expected: the rotated labels on slides 13 and 19. To make the PDF, open the deck in PowerPoint and
save it as PDF. With AppleScript, the file must sit in `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/`.

**Gotchas**
- This Mac writes decimal commas in AppleScript output. `excel_runner.py` already handles this.
- Excel and PowerPoint can only open files from inside their own `Containers/.../Documents` folders when you use
  AppleScript. The scripts copy files there and back.
- If you open v15 and edit it by hand, the next `build_model.py` run overwrites it. Make lasting changes in the
  script, or save your file under a new name.

## Still open

1. When does the second founder start being paid? (The plan keeps Rp15M a month of room from March.)
2. Owners: Kia (CEO) leads BI and banks, Galih leads EAP 2 and people, until a BD co-founder joins (proposal).
3. Is our 2027 BI work already in BI's 2027 budget? Who is the contact?
4. EAP contract 2: a warm lead exists; sign by mid-January to start in February.
5. Is our current EAP contract active, and for how many staff? Will BI accept the new 2027 prices?
6. Lead Partnership pay (estimate: Rp6M from Oct 2027); BPJS + THR cost (estimate: 18%); is the admin a formal employee?
7. Theta and CTO costs are in no model yet.
8. Tax: check that the 0.5% small-business tax (PP 55/2022) still applies as sales grow. Ask a tax consultant.
9. Lead times on the Roadmap sheet are estimates. Log real deals in its section E.
10. Update `projection guide.docx` to v15.

The full list of estimates is on the **To_Fill** sheet and on slide 24 of the deck.
