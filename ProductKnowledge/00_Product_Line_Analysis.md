# Product line analysis: input for the sales deck and master pricing sheet

Date: 19 Nov 2026. Source: every proposal, RAB, model and contract in this repo (11 folders, about 45 documents).
Money in Rupiah. "M" = million. Prices exclude VAT (TJ and Coreitera are both Non-PKP).

## 1. The answer in one paragraph

We sell **about 30 products in 6 families**, but no single file lists them all. Prices live in at least 7 places
(model v15 Rate_Card, B2C pricing sheet, partner contract Lampiran C, the Program Mitra deck, and each client RAB),
and they **conflict in 9 places**, the biggest being B2C counseling (the plan assumes Rp1M a session at 60% margin;
the real price list gives Rp500k-850k at 30-40%). Each RAB also uses a different margin rule. The best starting point
for one master list is **Lampiran C of the partner contract** (already a clean B2B price list) plus the **B2C Catalog sheet**,
with one cost table (Lampiran B, extended) and one margin rule applied to all.

## 2. Where the content is

| Folder | What it holds | Product family | Use for master list? |
|---|---|---|---|
| `CompanyWideAnalysis/` | 2027 model v15 (Rate_Card, Unit_Economics), growth deck, **Program Mitra deck** | All TJ B2B + B2C | Yes: Rate_Card, Unit_Economics, Mitra deck |
| `B2C Pricing/` | Psychologist tiers, revenue share, B2C catalog | B2C counseling, therapy, tests | Yes: Catalog and Profit_Share sheets |
| `SPKnProposalAnalysis/Kemitraan ... Iza/` | Partner contract with **Lampiran B (standard costs)** and **Lampiran C (price list)** | B2B training, EAP | Yes: closest thing to an official price list |
| `SPKnProposalAnalysis/BI Sibolga/` | 4 EAP package versions + margin models + proposal | EAP (Jeda, Tumbuh, Seperjalanan, Berbagi, Penyuluhan, Sapa) | Yes: final v2-correct-files only |
| `SPKnProposalAnalysis/lyfeat30s and canteenco/` | Ticketed community event, venue SPK | Public / community events | Yes, as a template |
| `BI Fit and Proper/` | Selection assessment for 6 BI candidates | B2B assessment (selection) | Yes |
| `BI Jampurnas/` | Jury service for BI national event | Expert jury / penjurian | Yes |
| `BSN Pemetaan Karyawan/` | 2,000-employee BFI + DISC mapping, 5 packages | Mass assessment + talent mapping | Yes |
| `Sample Docs or Data/` | MHCU for LPDP community, CBP Rupiah training, old RAB | MHCU, training | Yes (MHCU); old RAB is superseded |
| `CoreItera-Investment-Pitch/` | Theta bureau software pitch, model inputs | Theta SaaS plans, copilot seats | Only if Coreitera is in scope |
| `CoreItera-Strategy/` | Synthesia proposal v0.2 | AI interpretation for bureaus | Only if Coreitera is in scope |
| `TJ-Spinoff/` | Shareholder agreement, funding scenarios | None (corporate) | No |
| `_brand/` | Brand assets | n/a | Yes, for the deck design |

## 3. The product catalog (what sales can sell today)

### A. Training and capacity building (B2B, Teman Journey)

| Product | Unit | List price | Delivery cost | Margin | Source |
|---|---|---|---|---|---|
| HQ Trainer Program (Flagship ToT), 3 days, up to 40 pax, incl. assessment + impact report | program | Rp100M | Rp28M | 72% | Rate_Card, Lampiran C |
| Regional Trainer Program (KPW/LPS ToT), 1 day, up to 30 pax | day | Rp28.5M | Rp10.25M | 64% | same |
| Bank program, state-owned (BUMN), 1 group, up to 40 pax | day | Rp51M | Rp15M | 71% | same |
| Bank program, private bank | day | Rp61M | Rp15M | 75% | same |
| Bank Elite (program + trainer licence + 3 months follow-up) | day | Rp66.3M / Rp79.3M | not costed | n/a | same |
| Capacity building, custom (e.g. Duta CBP Rupiah) | project | custom | custom | n/a | Sample Docs |

### B. Employee Assistance Program (EAP, B2B)

| Package | Contents | Term | Price | Cost | Net margin | Source |
|---|---|---|---|---|---|---|
| Jeda | 15 online sessions, 1 webinar | 3 months | Rp12M | Rp7.67M | 36% | Sibolga model, Lampiran C |
| Jeda Tahunan | per quarter: 15 sessions, MHCU, aggregate report, webinar | 12 months | Rp11M/quarter (Rp44M/yr) | Rp8.75M/quarter | **20%** | same |
| Tumbuh | 24 sessions, MHCU, aggregate report, webinar | 3 months | Rp20M | Rp12.6M | 37% | same |
| Seperjalanan | 36 sessions, 2x MHCU (pre/post), individual results, offline seminar | 3 months | Rp31.5M | Rp20.07M | 36% | same |
| Extra session | 60 min online | session | Rp500k | Rp385k | 23% | same |
| Monthly payment option | 3-month packages paid monthly | 3 months | +20% (Rp4.8M / 8M / 12.6M a month) | same | higher | Lampiran C |
| Generic bank EAP (planning figure only) | retainer | month | Rp15M | Rp5M | 67% | model v15 |

### C. Assessment (B2B)

| Product | Unit | Price | Cost | Net margin | Source |
|---|---|---|---|---|---|
| Selection / Fit and Proper (rubric, BFI + DISC, LGD, interview, report) | participant | about Rp2.5M (Rp15M for 6) | Rp1.46M | 36% | BI Fit and Proper |
| Mass profiling BFI + DISC, auto report, company only | project (2,000 pax) | Rp50M (Rp25k/pax) | Rp6.9M | 76% | BSN |
| Mass profiling, company + employee reports | project | Rp70M (Rp35k/pax) | Rp6.9M | 79% | BSN |
| Talent mapping 3A / 3B / 3C (unit mapping, talent box, job fit) | project | Rp101.5M / 139.5M / 210.5M | Rp19M-62M | 62-71% | BSN |
| Simultaneous-run add-on | participant | +Rp15k | Rp9.3M per project | 69% | BSN |
| Assessment + impact report add-on | participant | Rp150k | Rp70k-120k | n/a | Rate_Card |
| MHCU Flash (5 instruments + 30 min counseling) | participant | Rp242k | about Rp245k | **about 0%** | MHCU LPDP |
| MHCU Deep (5 instruments + 60 min counseling) | participant | Rp380k | about Rp420k | **negative** | MHCU LPDP |
| Individual MHCU result add-on | employee | Rp70k | Rp70k | **0%** | Sibolga |

MHCU cost assumes the B2B psychologist rate (Rp350k/hour) plus the Rp70k tool. If the LPDP sessions used cheaper
psychologists, the margin is positive, but this must be checked.

### D. Expert services and psychoeducation (B2B)

| Product | Unit | Price | Cost | Net margin | Source |
|---|---|---|---|---|---|
| Expert jury (psychologist + specialist + PM) | event | Rp15M | Rp7.4M | 44% | BI Jampurnas |
| Webinar by psychologist | session | not listed alone | Rp1.5M | n/a | Lampiran B |
| Offline seminar (Jakarta) | session | not listed alone | Rp3M | n/a | Lampiran B |
| Public psychoeducation event | event | Rp24M | Rp8M | 67% | Lampiran C |

### E. B2C (individuals, Teman Journey)

Psychologist tiers: Tier 1 (0-2 yrs), Tier 2 (2-5), Tier 3 (5-10), Tier 4 (10+). Psychologist share 60/65/65/70%.

| Service | Tier 1 | Tier 2 | Tier 3 | Tier 4 | TJ gross margin |
|---|---|---|---|---|---|
| Individual counseling, online 60 min | Rp500k | Rp500k | Rp725k | Rp850k | 30-40% |
| Individual counseling, offline 60 min | Rp580k | Rp580k | Rp800k | Rp925k | 30-40% |
| Couples counseling, offline 90 min | Rp800k | Rp950k | Rp1.1M | Rp1.425M | 30-40% |
| Pre-marriage check-up (3 sessions) | Rp2.2M | Rp2.6M | Rp2.9M | Rp3.8M | 30-40% |
| Marriage check-up (2 sessions, estimate) | Rp1.4M | Rp1.7M | Rp2.0M | Rp2.6M | 30-40% |
| Clinical therapy (EMDR / hypnotherapy), offline | Rp750k | Rp900k | Rp1.0M | Rp1.325M | 30-40% |
| Clinical therapy, online | Rp675k | Rp800k | Rp900k | Rp1.2M | 30-40% |

| Test (single price, 50:50 split) | Price | TJ margin |
|---|---|---|
| Child development screening (0-5) | Rp625k | 50% |
| IQ test (estimate) | Rp725k | 50% |
| School readiness (4-6) | Rp800k | 50% |
| Interests and aptitude (14+) | Rp1.625M | 50% |
| Profiling (6-13) | Rp1.225M | 50% |
| Adult mental health check (MHCU) | Rp375k | 31% |
| Ticketed community event (e.g. Lyfe at 30s) | Rp259k/pax | about Rp1M profit per event at 12 pax |

After the Rp50k platform fee, a counseling session earns TJ Rp125k-205k. A client won through Instagram ads
**loses** money on counseling alone (Rp115k-195k per session), so B2C is a funnel, not a profit centre yet.

### F. Coreitera and Theta (technology, sold to psychology bureaus)

| Product | Unit | Price | Status |
|---|---|---|---|
| Theta bureau plan: Starter (1-3 psychologists) | month | Rp1M | Live at TJ from 30 Nov 2026 |
| Theta: Growth (4-10) | month | Rp2.9M | Plan |
| Theta: Pro (11+) | month | Rp5.73M | Plan |
| Payment fee | transaction | Rp5k + Midtrans at cost | Plan |
| Synthesia / Psynthesia copilot | psychologist seat a month | Rp300k (+ fee per signed report, to set) | From 2028 |
| White-label test engine | participant | about cost | Plan |
| Self-test platform with "get it signed" upgrade | participant | matches market (Rp15k-45k) | Last phase |

These have a different buyer (bureau owners, not HR or regulators) and are mostly not live.

## 4. Conflicts the master sheet must settle

| # | Conflict | Where | Effect |
|---|---|---|---|
| 1 | **B2C counseling price and split**: plan uses Rp1M/session and 40% to psychologist; price list is Rp500k-850k and 60-70% | Model v15 B2C vs B2C Pricing | Profit per session 64% lower than plan; B2C profit in model v15 is overstated |
| 2 | **Psychologist cost basis**: Rp350k/hr face-to-face + Rp110k/hr document work (B2B) vs % share (B2C) vs Rp5M/month retainer (model EAP) | RABs, B2C Pricing, model | Same session costs differently by channel |
| 3 | **Margin rules**: "price at least 3x external cost" vs "net margin 15-26%" vs "net margin 35-40%" vs "partner deals keep 15% after commission" | Rate_Card, Sibolga, BI RABs, Lampiran/PKS | Sales cannot tell if a discount is allowed |
| 4 | **Assessment price per person**: Rp25k-50k (BSN mass), Rp150k (rate card add-on), Rp200k (BFI+DISC in Fit and Proper), cost Rp120k BFI+DISC vs Rp75k Coreitera fee estimate | BSN, Rate_Card, Fit and Proper, model | Same test, four prices |
| 5 | **MHCU** sold at or below cost (Rp70k add-on = cost; Flash/Deep near or below cost) | Sibolga, MHCU LPDP | Lost margin on a product we push in every EAP |
| 6 | **Tax wording**: "includes PPN 11%" in MHCU LPDP and an old BI offer, but TJ is Non-PKP | MHCU LPDP, BI Fit and Proper | Legal and invoice risk; fix the template |
| 7 | **Public event**: Rp20M in model vs Rp24M in price list | Model Inputs vs Lampiran C | Small, but the plan should use the list |
| 8 | **Jeda Tahunan** at 20% net, below the Rp15M/month planning EAP (67%) | Sibolga vs model | The EAP growth in the 2027 plan uses a much richer EAP than we actually sell |
| 9 | **Talent mapping A/B/C** first-order prices at about 2.5x cost, below the 3x rule | BSN | Allowed as discount, but needs a written rule |

Also: a flat **10% commission on gross profit** sits in every RAB, with no named recipient, while the partner contract
defines commission differently (10-12.5% of net, EAP fixed fees). Pick one.

## 5. What this means for the two deliverables

**Product knowledge deck (for sales).** Organise by the buyer's problem, not by our folder:
1. Who we are (TJ + Coreitera platform, licences, psychologist network).
2. Six families above, one slide each: what it is, who buys it, what is included, price from, proof (client).
3. Package ladders (EAP Jeda to Seperjalanan; bank Essential to Elite; mapping 1 to 3C).
4. Selling rules: list price floor, max discount, travel at cost, payment terms, what needs Director approval.
5. Case studies: BI (Flagship, Sibolga EAP, Fit and Proper, Jampurnas), BSN, LPDP community.

**Master spreadsheet.** One workbook, built like the existing models (inputs in blue, every price a formula):
- `Cost_Rates`: one rate per resource (psychologist hour by type and tier, tool per test, platform, admin, contract, travel).
- `Catalog`: every product with unit, list price, bill of materials (qty x cost rate), margin, multiple, and the rule check.
- `Discount_Rules`: floor, partner commission, volume and pilot discounts, approval levels.
- `Quote_Builder`: pick products and quantities, get price, cost, net margin after PPh 23 and commission.
- `Source_Map`: which proposal or RAB each number came from, and which conflicts were settled how.

## 6. Decisions needed before building

1. **Scope**: DECIDED 19 Nov 2026: include Coreitera / Theta / Synthesia alongside Teman Journey.
2. **Margin rule**: DECIDED (revised): margin only. At least 30% net margin after PPh 23 and commission. The 3x rule is dropped (it pushed EAP, selection and jury prices up 40-120% and MHCU up nearly 300%).
3. **Psychologist cost basis**: DECIDED (revised): B2B sessions (EAP, MHCU) cost a flat Rp350k for every tier; B2C uses the tier share (60/65/65/70%).
4. **MHCU and Jeda Tahunan**: DECIDED: MHCU Flash Rp500k, Deep Rp750k (fixed). MHCU test cost corrected to Rp7k max (Galih, 19 Nov 2026), so the individual add-on stays Rp70k (44%). Jeda Tahunan modelled with MHCU for 70 employees, as the proposal promises. EAP (revised): all EAP packages keep today's price ("Harga tetap"); Jeda Tahunan (18%) and the extra session (19%) are below 30% and need Director approval when sold.
5. **Language**: DECIDED: Bahasa Indonesia for both the deck and the spreadsheet.

## 7. Final outputs (19 Nov 2026)

- `Master_Produk_dan_Harga.xlsx`: 69 products, 186 cost lines, quote simulator. Built by `build/build_master_sheet.py`.
- `01_Daftar_Produk_Final.md`: the final product list with prices and margins, exported by `build/export_product_list.py`. Input for the sales deck.
- Costs and net profits match the source RABs exactly for BI Fit and Proper, BI Jampurnas and BSN.
