# v16: projection update and proposal check

Written 16 Oct 2026, for the founders' feedback of the same day. Model: `TemanJourney_Model_2027_v16.xlsx`.
All money in Rp million (Rp M) unless stated. "Estimate" means a number we still need to confirm (sheet `To_Fill`).

## 1. The answer

1. **The new plan grows sales but earns a lower margin.** In the v16 plan, 2027 sales are Rp2,433M (v15: Rp2,040M),
   net profit is Rp557M, and the net margin falls from 26.1% to 22.9%. Two things pull the margin down. The latest B2C
   prices earn about 72% less per session. The new EAP packages for BI regional offices (BI KPw) earn little per
   package at the Sibolga prices.
2. **With the recommended fixes, the margin is 25.0% and cash is safe:** reprice the standard KPw packages, sell
   with the Commit-2 scheme below, and apply the cash guard now. Lowest cash is then Rp107M, with 1 month below the
   safety floor. Without the guard, cash falls to Rp70M in February 2027, which triggers our own guard rule (under Rp80M).
3. **B2C at the latest prices does not pay for Instagram ads.** Profit per client is about Rp240k. Winning one client
   through Instagram costs Rp450-840k, so ads return 0.1-0.5x their cost against our 1.7x rule. B2C pays when clients
   come through EAP, schools and referral, so the EAP-as-channel focus is the right call.
4. **Focusing on EAP does not lower our dependence on BI yet.** Most EAP growth comes from BI KPw offices, which is BI
   money. All BI money is 56% of 2027 revenue, compared with about 54% in v15 on the same basis. To really lower it,
   EAP contracts 2 and 3 must go to non-BI clients (banks, LPS/OJK, companies), and schools must grow.
5. **The proposals hold together, with four fixes.** These are covered in section 3: Sibolga margins, the psychologist
   fee rule, PPh 23 and Berbagi.

| 2027 | v15 plan | v16 plan | v16 recommended | v16 conservative | v16 aggressive |
|---|---|---|---|---|---|
| Sales | 2,040 | 2,433 | 2,451 | 1,739 | 3,300 |
| Net profit | 532 | 557 | 612 | 236 | 983 |
| Net margin | 26.1% | 22.9% | 25.0% | 13.6% | 29.8% |
| Lowest cash (month) | 82 (Feb) | 70 (Feb) | 107 (Jan) | 55 (Apr) | 69 (Feb) |
| Months below the safety floor | 1 | 5 | 1 | 7 | 5 |
| Cash at end of 2027 | 524 | 542 | 598 | 290 | 907 |

- **Recommended** = the plan, plus standard KPw packages at a 30% margin, the second founder's pay from May, the
  extra ads from March, and a 30% down payment on Flagship programs.
- **Conservative** = BI buys 30% less, no bank deal, half the KPw packages and schools, a weaker upsell.
- **Aggressive** = the bank deal from July at 2 days a month, 50% more KPw packages, 9 schools.

Check: with the v16 switch off (Inputs `v16_on` = 0), the workbook gives exactly v15 (Rp2,040M, 26.1%).

## 2. How we got there (thinking framework)

From first principles, profit = **volume x (price - variable cost) - fixed cost**, and cash = profit shifted by
**when each client pays**. Each piece of feedback changes one of these levers:

| Feedback | Lever | Where in v16 |
|---|---|---|
| 1. Pricing scheme | Price, and repeat volume | `Pricing_Scheme`, Inputs `commit_*` |
| 2. Targets (EAP channel, BI BAU, LPS/OJK, banks) | Volume by segment | `Segments`, `Channels`, Volume rows 38-50 |
| 2b. Schools | New volume line | `Channels` row 16, Inputs `sch_*` |
| 3. Proposal check | Price and cost per unit vs the model | Section 3 below |
| 4. Latest B2C prices | B2C price and cost per session | `B2C_Pricing`, Inputs `b2c_*` |

**What each piece is worth in the v16 plan** (net profit 2027; one at a time, so they do not add up exactly):

| Piece | Net profit effect |
|---|---|
| Latest B2C prices (vs Rp1M at 40%) | -96 |
| Public events at the real Lyfe at 30s format (vs a Rp20M mass event) | -42 |
| EAP channel: BI KPw packages, their add-on sessions and trainings | +62 |
| Schools | +47 |
| LPS / OJK programs from July | +33 |
| Commit-2 scheme (5% off + 2 free penyuluhan) | +2 |
| Stop the extra Rp3.5M a month of Instagram ads | +42 (upper bound: the model does not cut clients when ads stop) |

## 3. Proposal check (feedback 3)

### A. BI Sibolga EAP packages (costing v4)

| # | Check | Finding | Fix |
|---|---|---|---|
| A1 | Price vs cost | With every session paid at Rp350k, margins are 21% (Sapa), 24% (Tumbuh), 28% (Seperjalanan), 6% (Berbagi) and 17% (Penyuluhan). The average package is Rp15.5M at a 23% margin. | Next offices: Tumbuh at Rp15.5M and Seperjalanan / Penyuluhan at Rp24M (30% margin; `Channels` column H). Keep Sapa as the small intro. Sibolga stays as proposed. |
| A2 | Fit with the model | v15's EAP is Rp15M a month for 10 sessions (Rp1.5M a session); Sibolga is Rp0.5-0.64M a session. They are different products. | v16 models KPw packages as their own line. Never quote the per-session KPw price to BI head office: under Rate_Card rule 7 it would become the head-office price. |
| A3 | Psychologist fee | Sibolga pays a flat Rp350k. The new pricing file pays Rp300k / 325k / 471k / 595k by tier (online). A Tier 3-4 psychologist on BI sessions costs more than the costing assumes. | Set one EAP fee table by tier (e.g. Rp300-400k) and use it in both files. Give KPw quota sessions to Tier 1-2 psychologists. |
| A4 | "In-house = Rp0 honor" view | This only holds while the paid part-timers have spare time. Their load is 32-69% from Oct-26 to Mar-27, so Sibolga (Nov-Jan) can be delivered in-house. From Jul-27 the load is 122-149%, and the pool of per-session psychologists takes the overflow. | Fine for Sibolga. For later offices, cost every session at the honor (as v16 does). |
| A5 | Tax and cash | BI withholds 2% PPh 23, which the model does not show. With a Surat Keterangan PP 55, the buyer withholds only the 0.5% final tax: 1.5% more cash on every BI invoice. The 50% down payment at SPK is good. | Ask the tax consultant for the Surat Keterangan. Make the 50% down payment standard (Commit-2 rule 3). |
| A6 | Berbagi | Same content as Seperjalanan, but Rp3M cheaper, and TJ pays Rp2.3M for the webinar. 6% margin. | Drop it, or price it like Seperjalanan. |
| A7 | Ladder shape | Sapa = small scope at the full Rp500k a session: exactly the intro rule. | Keep it. |

### B. Lyfe at 30s x Kanteen n Co (B2C event, 31 Oct)

| # | Check | Finding | Fix |
|---|---|---|---|
| B1 | Event economics vs the model | The real format earns about Rp1M of cash profit at 12-16 people. v15 assumed a Rp20M mass event with Rp12M of contribution. | v16 uses the Lyfe numbers (Rp4.1M revenue, Rp2.7M cost). This takes Rp42M of profit out of 2027. |
| B2 | Voucher value | The proposal counts "~Rp600k contribution per conversion". That is the old Rp1M x 60%. At the latest prices it is about Rp170k a session, or about Rp240k a client. | Judge the event by referral-code sign-ups, package sales and Mega's content, not by counseling profit. |
| B3 | As a cost of winning clients | TJ's full cost (~Rp4.5M incl. the team-week) minus ~Rp1M of profit = ~Rp3.5M. Counseling alone would need ~15 new paying clients to earn it back. | Keep it as a partnership and content play. No series commitment until the 14 Nov review (already in the proposal). |
| B4 | Terms | Sign Kanteen first; cap ads; H-5 gate. | No change. |

### C. Latest B2C pricing file

It is internally consistent (check cell "OK") and already flags the gap with v15. v16 uses it as is (tier mix 30/35/25/10%,
70% online, Rp50k platform cost). Average price paid Rp615k, psychologist share 64%, **TJ margin Rp169k a session**
(v15: Rp600k).

## 4. Pricing scheme (feedback 1)

**Principle:** a discount must buy something worth more than it costs: a second sale we would otherwise lose, cash
earlier, and one less sales cycle.

| Step | What it is | Rule |
|---|---|---|
| Intro | Small scope at the **full unit price** (Sapa 15 sessions Rp7.5M; a 2-hour penyuluhan Rp5M; a half-day bank workshop; one school class) | Never a lower unit price: regulators will not accept a price rise later (rule 7). The intro fee is credited to Commit-2 if the client signs within 60 days. |
| List | One term or one event at the rate card | Floor: at least 30% margin (EAP), 3x the external cost (training). |
| Commit-2 | Two terms or events in one SPK, both dates fixed, 50% paid at signing | "Cheaper overall" = **5% off + 2 free penyuluhan**. The client sees 6-11% of value; TJ gives about 6% (the penyuluhan cost about Rp100k each). |
| BI head office | Flagship | No Commit discount. Ask for a 30% down payment instead. |
| B2C | Tier 1-2 session Rp500k online as the entry; 4-session package 10% off | Offer the package after the first session. |

Why 5% and not 10%: on a KPw package the largest cash discount that still beats selling one term at a time is **7%**
(at a 30% renewal rate). A 10% discount loses money at today's prices: net profit Rp551M at 10% off vs Rp557M at
5% off + bonus. Training and bank days can take up to 22-25%, and schools up to 13% (`Pricing_Scheme` column K).

## 5. Targets for 2027 (feedback 2)

| Priority | Target | Plan 2027 | Revenue | Share |
|---|---|---|---|---|
| 1. EAP as the channel | 10 new BI KPw packages (Sibolga first, Nov-26); 4.6 second terms or renewals; EAP contracts 2 and 3 (Feb, Jul); 2.3 trainings sold to EAP clients; 66 paid channel sessions | | 573 | 23.5% |
| BI BAU premium | Flagship 5 + BI regional ToT 14.5 + existing EAP: same as today | | 1,092 | 44.9% |
| 2. LPS / OJK | ~2 programs from July (mid-year budget revisions) + non-BI regional ToT | | 131 | 5.4% |
| 3. Banks (private last) | Bank events and training days, as v15 | | 375 | 15.4% |
| New: schools | 6 schools, 480 students (Feb pilot, Jul grade 10, Aug-Nov grade 12) at Rp300k each | | 144 | 5.9% |
| B2C direct | Instagram, referral, vouchers, events | | 119 | 4.9% |
| **All BI money** | BI BAU + KPw packages + their trainings | | 1,355 | **55.7%** |

**Schools** at Rp300k per student (group test with a short individual report): about 38% margin after the test tool
and the 50:50 assessor share. Plus a 1:1 follow-up for 5% of students at the B2C price. The individual Tes Minat
Bakat (Rp1.625M) stays the upsell. Price, tool cost and who pays (school or parents) are all estimates.

**Capacity:** psychologist load reaches 122-149% from July 2027. The per-session pool of 20 covers 16-35 sessions a
month. This costs nothing extra (paid per session), but book the pool early and keep the Lead Partnership hire in
October.

## 6. B2C at the latest prices (feedback 4)

| FY 2027 | v15 | v15 volume, latest prices | v16 plan |
|---|---|---|---|
| Paid sessions | 166 | 166 | 232 (66 via EAP and schools) |
| Counseling revenue | 166.0 | 102.0 | 142.7 |
| Public event revenue | 80.0 | 80.0 | 16.6 |
| **B2C revenue** | **246.0** | **182.0** | **159.3** |
| Gross profit (after psychologist fees, platform, event costs) | 147.6 | 76.0 | 44.8 |
| Paid part-time psychologists (shared with EAP) | 65.1 | 65.1 | 65.1 |
| Extra Instagram ads | 42.0 | 42.0 | 42.0 |
| **B2C profit after psychologists, ads, referral** | **+37.8** | **-33.8** | **-65.0** |

| Per client, Dec 2027 | v15 | Latest prices |
|---|---|---|
| Price per session | Rp1,000k | Rp615k |
| TJ margin per session (after psychologist and platform) | Rp600k | Rp169k |
| Profit per client (1.4 sessions) | Rp840k | Rp236k |
| Profit per client / cost to win one, Instagram | 1.9x | 0.5x (0.1-0.5x through 2027) |
| Profit per client / cost to win one, referral (Rp200k credit) | 3.3x | 1.2x |

What to do:
- Run the Rp3.5M extra ads for one quarter with source tracking. Stop them if the return is under 1.7x.
- Cut the referral credit to Rp50k each side (or one free 30-minute check-in).
- Push the 4-session package. On its own it adds only about Rp3M (+36 sessions); B2C works as an add-on to EAP and
  schools, not as a profit line on its own.

## 7. What changed in the model (v16)

New sheets:
- `Segments`: revenue by priority, plus the 2027 target list.
- `B2C_Pricing`: tiers, profit share, mix.
- `Channels`: KPw package catalog from the Sibolga v4 file; monthly plan for packages and schools.
- `Pricing_Scheme`: intro / list / Commit-2 for each product, with the break-even rule.

New rows:
- Inputs section 14 (rows 130-160).
- Volume rows 38-50, PnL rows 46-57, CashFlow rows 31-35.

Existing cells that changed:
- Inputs C13 / C14 / C20 / C21 now switch between the v15 and v16 values.
- Capacity, invoices, project-manager fees, admin per-client pay and psychologist load now include the new lines.
- Summary rows 34-51 show the v16 scenarios and the segment mix.

Scenario runs: `output/v16_runs.json` (`analysis/lo_runner.py`).

The deck (`TemanJourney_Growth_Recommendation.pptx`) and `projection guide.docx` still show v15.

## 8. Still to confirm

1. Sibolga: which package did BI KPw choose, and is the SPK signed?
2. Schools: price per student, who pays (school or parents), which test and what it costs, students per school.
3. Is today's EAP contract with BI? (Inputs `eap_bi`)
4. EAP psychologist fee by tier (flat Rp350k vs the new tier fees).
5. Commit-2 take-up (40%) and renewal without it (30%): replace with real numbers after the first 3 offices.
6. LPS / OJK: is a mid-2027 program realistic, or does it move to 2028 budgets (Roadmap sheet)?
7. Surat Keterangan PP 55 (tax consultant).
