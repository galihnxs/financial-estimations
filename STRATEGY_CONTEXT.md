# Strategy context: where to start reading

Index for a new session on company strategy (Teman Journey, Coreitera, Theta). Written 25 Nov 2026.
Read the "Tier 1" files first. Everything else is supporting evidence.

## Tier 1: read first (the current strategy and numbers)

| # | File | What it gives you |
|---|---|---|
| 1 | `CompanyWideAnalysis/README.md` | The TJ 2027 plan in one page: goals, 4 impact pillars, pay steps, model v15 results (sales Rp2,040M, net margin 26.1%, lowest cash Rp82M in Feb 2027), guard rule, how to rebuild. |
| 2 | `CompanyWideAnalysis/TemanJourney_Model_2027_v15.xlsx` | **The current financial model.** Open Outlook, Roadmap, Summary, Inputs. |
| 3 | `CompanyWideAnalysis/TemanJourney_Growth_Recommendation.pdf` | 43-slide growth plan (B2B, B2C, people, money). Pyramid order. |
| 4 | `CompanyWideAnalysis/TJ_Portfolio_Focus.xlsx` | Nov 2026 portfolio review: which of 24 product families earn their place. Open Read_Me, then Portfolio. |
| 5 | `ProductKnowledge/06_Strategi_ke_Produk_dan_Harga.md` | Links company goals to products and segment pricing. Key finding: the EAP catalogue does not support the 2027 plan (margin falls from 26% to about 17%). |
| 6 | `TJ-Spinoff/TJ_Funding_Scenarios_and_Asking_Price.docx` | New TJ company funding: recommended "Build" scenario, Rp350M at Rp1,500M pre-money. Indonesian version: `TJ_Skenario_Pendanaan_dan_Harga_Penawaran.docx` (has a disclaimer; revenue figures are placeholders waiting for correction). |
| 7 | `CoreItera-Strategy/proposal.md` | Coreitera Synthesia v0.2: sell AI interpretation to psychology bureaus, do not join the test price war. |
| 8 | `CoreItera-Investment-Pitch/README.md` | Theta partnership pitch: Rp30M for 20% of Coreitera (pre-money Rp480M), model in `model/inputs.py`. |

## Tier 2: products and pricing (ProductKnowledge/)

| File | Use |
|---|---|
| `00_Product_Line_Analysis.md` | About 30 products in 6 families; where prices live and the 9 conflicts. Has a folder map of the repo. |
| `01_Daftar_Produk_Final.md` | Final product and price list (TJ + Coreitera). |
| `05_Struktur_Produk_dan_Segmen.md` | Product structure by segment. |
| `07_Benchmark_Kompetitor.md` | Does our price hypothesis survive competitor prices? |
| `08_Kalkulasi_Biaya_Layanan.md` | Service cost build-up. |
| `03_Cerita_Produk_Lengkap.md` | Buyers, their problems, how to sell. |
| `Master_Produk_dan_Harga.xlsx`, `Kalkulator_Sales.xlsx` | Master price sheet and sales calculator. |
| `04_Inventaris_Canva.md` | Canva deck inventory and change log. Only for deck work. |

## Tier 3: market and competitor evidence

| Path | Use |
|---|---|
| `CoreItera-Strategy/psychometric-competitor-research/` | Psychometric competitor report (21 Nov), UGM / UI / Unpad network notes, dean briefing, instrument requests. |
| `.../faxtor-offer-2026-11/` | Faxtor catalogue and 2026 price list (internal only). |
| `CoreItera-Strategy/[2026] Teman_Journey_ME_KPI_Q4_2026.xlsx` | Q4 2026 KPI tracking. |
| `CoreItera-Strategy/Marketing Monitoring TJ 2026.xlsx`, `Dokumen Rekapitulasi LEADS Marketing.xlsx`, content and LinkedIn recaps | Marketing and leads data. |
| `CompanyWideAnalysis/TemanJourney_Program_Mitra.pdf` | Partner program deck. |
| `B2C Pricing/TJ_Pricing_Tier_Psikolog_Profit_Sharing.xlsx` | B2C prices and psychologist revenue share. |

## Tier 4: real deals (evidence of what clients buy and pay)

| Folder | Deal |
|---|---|
| `BI Fit and Proper/` | Selection assessment for BI. |
| `BI Jampurnas/` | Jury service for a BI national event. |
| `BSN Pemetaan Karyawan/` | 2,000-employee BFI + DISC mapping (latest: "Pembaruan Akhir"). Confidential RAB. |
| `SPKnProposalAnalysis/` | BI Sibolga EAP packages, Lyfe at 30s and Kanteen n Co events, partner contract (Lampiran B costs, Lampiran C price list). |
| `Sample Docs or Data/` | Older proposals (MHCU LPDP, CBP Rupiah). The old RAB is superseded. |

## Watch out: stale or conflicting

- **Model versions:** v15 is current. v13 and v14 are kept only for comparison. `projection guide.docx` reviews v13 and is out of date.
- **Older reference numbers:** some notes (and the fractional-cfo skill's company reference) still quote model v11 (FY2027 revenue Rp1,208M, 12% margin). Use v15.
- **EAP price gap:** the plan assumes bank EAP at Rp15M per month; real packages sell at Rp6.7M to Rp10.5M (see `06_...md`).
- **B2C counseling:** the plan assumes Rp1M per session at 60% margin; the real price list is Rp500k to 850k at 30-40% (see `00_...md`).
- **School revenue share "maks 10%":** removed from the Canva client deck on 24 Nov 2026. It is still in `Master_Produk_dan_Harga.xlsx` and the 05/06 notes, waiting for a decision.
- **Spinoff revenue figures** (last year Rp350M, this year Rp800-900M) are placeholders until Riskia corrects them.

## Rebuild tools

- Python: `/opt/homebrew/Caskroom/miniforge/base/envs/ind5003/bin/python` (openpyxl, python-pptx).
- Excel for Mac recalculates the model through AppleScript (`CompanyWideAnalysis/analysis/excel_runner.py`). There is no LibreOffice.
