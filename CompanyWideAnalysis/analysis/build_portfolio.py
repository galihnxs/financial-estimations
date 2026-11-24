"""
Build TJ_Portfolio_Focus.xlsx: the product portfolio evaluation for Teman Journey and Coreitera.

Run from financial-estimations/:   python3 CompanyWideAnalysis/analysis/build_portfolio.py
Then recalculate in Excel (openpyxl cannot calculate formulas).

Sheets
  Read_Me             how to read and fill it, label legend
  Portfolio           FORMULAS ONLY: the metrics, the suggested call and the data confidence
  Settings            roles (pay, days available, day rate) and effort assumptions
  Product_Breakdown   INPUTS per product family: price, cost, units, person-days per role, strategy flags
  Past_Activity       what we actually delivered (2025-2026), with revenue and cost where known
  Acquisition_Effort  what we spent to win clients in 2026: content, outreach, ads, equipment

Every input carries a label: Actual (from records), Plan (our price list or model v15),
Proxy (estimate from indirect data), Missing (no data yet).

Sources (23 Nov 2026): ProductKnowledge/Kalkulator_Sales.xlsx; model v15; CoreItera-Strategy/*.xlsx
(KPI Q4, content plans, leads, LinkedIn recap); Downloads/BUDGET - Marketing & Operations.xlsx;
Downloads/_Data Konseling 2026.xlsx (aggregated here; no client names are copied).
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "TJ_Portfolio_Focus.xlsx"

KORP = 0.9 / (0.882 - 0.50)   # catalog pricing rule, corporate 50% target
EDU = 0.9 / (0.882 - 0.30)    # catalog pricing rule, education 30% target

# ---------------------------------------------------------------- roles
ROLES = [  # name, people, pay or value of time per month (Rp M), pay label, days available per month, source
    ("Kia", "Kia (co-founder, lead Teman Journey)", 10, "Plan", 15, "Pay: model v15 founder salary. Days: estimate."),
    ("Galih", "Galih (founder, Coreitera / tech)", 10, "Plan", 8, "Pay: model v15 f_sal2. Days: estimate for TJ + Coreitera."),
    ("Partnership / sales", "Iza (partnership, commission based)", 6, "Plan", 10, "Value of time: v15 Lead Partnership pay. Paid by commission today."),
    ("Lead marketing", "Desya (lead marketing)", 2.5, "Plan", 8, "Pay: v15 o_lead_mkt."),
    ("Content team", "Social media + interns (Azka, Dzakki, Rana, Qisty)", 3.5, "Plan", 16, "Pay: v15 social media 1.5 + interns 2.0."),
    ("Admin", "Admin", 2.0, "Plan", 10, "Pay: v15 Pay_Plan Oct-26."),
    ("PIC psychologist", "Vania (PIC psychologist, non-session work)", 2.3, "Plan", 6, "Pay: v15 o_pic. Session fees are in variable cost."),
    ("Project manager", "Project manager (paid per project)", 4.0, "Plan", 8, "Rp2M per project, about 2 projects a month (v15 pm_fee, pm_cap)."),
]
ROLE_ROW0 = 5  # Settings rows 5..12

# ---------------------------------------------------------------- products
# id, family, business, seg, funnel, codes, unit,
# price, price label, cost, cost label, units/yr, units label,
# days per unit by role (ROLES order), fixed days/yr, fixed role, lead time days, parallel?, repeat?, leads?, leads to, tech?, note
P = [
    ("P01", "Flagship regulator ToT", "TJ", "B2B", "Revenue", "TJ-PLT-01", "program (3 days)",
     100, "Plan", 28, "Plan", 5, "Plan", [8, 1, 3, 0, 2, 2, 4, 6], 0, "", 30, "N", "Y", "Y", "Regional ToT, bank programs", "N",
     "Units: v15 plan 2027. Kia leads delivery, so it cannot run in parallel with another Flagship."),
    ("P02", "Regional ToT (KPW / LPS)", "TJ", "B2B", "Revenue", "TJ-PLT-02", "training day",
     28.5, "Plan", 10.25, "Plan", 17, "Plan", [2, 0, 1, 0, 0.5, 1, 1, 2], 0, "", 14, "Y", "Y", "Y", "Bank programs, EAP", "N", "Units: v15 plan 2027."),
    ("P03", "Bank training programs", "TJ", "B2B", "Revenue", "TJ-PLT-03 to 09", "training day",
     51, "Plan", 15, "Plan", 7, "Plan", [4, 0, 3, 0, 0.5, 1, 1, 2], 0, "", 30, "Y", "N", "Y", "EAP, employee mapping", "N",
     "Units: v15 bank days 2027. Price: BUMN Signature. Long sales cycle."),
    ("P04", "EAP retainer: ongoing counseling for a company (Jeda, Tumbuh, Seperjalanan, Korporat)", "TJ", "B2B", "Retention", "TJ-EAP-01 to 06", "contract-quarter",
     20, "Plan", 11.7, "Plan", 12, "Plan", [1, 0, 0.5, 0, 0, 1.5, 1, 0.5], 0, "", 90, "Y", "Y", "Y", "Bank programs, renewal", "Y",
     "Ongoing support over 3-12 months: a pool of counseling sessions, webinars, aggregate reports. MHCU rounds are INCLUDED in "
     "Jeda Tahunan, Tumbuh, Seperjalanan and Korporat; one-off MHCU is P10. Units: 3 contracts x 4 quarters (v15). Price: EAP Tumbuh."),
    ("P05", "Prepaid counseling packages for a company (10 / 20 sessions, extra sessions)", "TJ", "B2B", "Revenue", "TJ-EAP-05, TJ-KON-01 to 05", "10-session package",
     round(4.345 * KORP, 1), "Plan", 4.345, "Plan", None, "Missing", [0.2, 0, 0.3, 0, 0, 0.3, 0.2, 0], 0, "", 180, "Y", "N", "Y", "EAP retainer","Y",
     "New 20 Nov 2026. Price: catalog rule, corporate."),
    ("P06", "In-house psychologist", "TJ", "B2B", "Retention", "TJ-KON-06, 07", "day on site",
     round(2.47 * KORP, 1), "Plan", 2.47, "Plan", None, "Missing", [0, 0, 0.2, 0, 0, 0.2, 0.1, 0], 0, "", 1, "Y", "Y", "Y", "EAP retainer","N",
     "New 20 Nov 2026."),
    ("P07", "Selection / fit-and-proper assessment", "TJ", "B2B", "Revenue", "TJ-ASM-01", "project (6 people)",
     15, "Plan", 8.77, "Plan", 1, "Plan", [2, 0.5, 0.5, 0, 0, 1, 1, 1], 0, "", 21, "Y", "N", "Y", "Managerial assessment, mapping", "Y",
     "Sold once to BI for Rp15M (Actual, 2026). Tests run on Coreitera."),
    ("P08", "Mass profiling and employee mapping", "TJ", "B2B", "Revenue", "TJ-ASM-02 to 07", "project (2,000 people)",
     70, "Plan", 6.89, "Plan", None, "Missing", [2, 3, 1, 0, 0, 1, 1, 1], 0, "", 30, "Y", "N", "Y", "EAP, MHCU", "Y",
     "Price: Paket 2. BSN proposal in progress."),
    ("P09", "Managerial and executive assessment", "TJ", "B2B", "Revenue", "TJ-ASM-13, 14", "batch of 10 people",
     round(8.35 * KORP, 1), "Plan", 8.35, "Plan", None, "Missing", [1, 0.5, 0.5, 0, 0, 0.5, 0.5, 0], 0, "", 21, "Y", "N", "Y", "Employee mapping", "Y",
     "New 20 Nov 2026."),
    ("P10", "MHCU check-up, one-off for a company (Flash, Deep, L0-L2)", "TJ", "B2B", "Lead gen", "TJ-ASM-08 to 12, 19", "check-up round (50 people)",
     12.1, "Plan", 12.2, "Plan", None, "Missing", [0.5, 0.3, 0.5, 0, 0, 0.5, 0.5, 0], 0, "", 30, "Y", "Y", "Y", "EAP retainer", "Y",
     "One-off screening, priced per person: 5 instruments, psychologist review, personal report, plus 1 counseling session "
     "(30 min Flash, 60 min Deep). NOT a retainer: ongoing support is P04. Price: Flash 242k x 50. Today's price is BELOW cost "
     "because of the session inside. Telering is the live client."),
    ("P11", "School test packages and child assessment", "TJ", "B2B", "Revenue", "TJ-ASM-15 to 18", "school (100 students)",
     round(16.57 * EDU, 1), "Plan", 16.57, "Plan", None, "Missing", [2, 0.5, 1, 0, 0, 1, 1.5, 1], 0, "", 30, "Y", "Y", "Y", "School workshops, B2C tests", "Y",
     "New 20 Nov 2026. Price: catalog rule, education."),
    ("P12", "Individual tests and Career Compass (B2C)", "TJ", "B2C", "Revenue", "TJ-TES-*, Career Compass, Career Reading", "test",
     1.625, "Plan", 0.8125, "Plan", None, "Missing", [0, 0, 0, 0, 0, 0.1, 0.1, 0], 0, "", 7, "Y", "N", "Y", "B2C counseling", "Y",
     "Price: Minat Bakat. Assessor gets 50%. Career Compass bundles sold at Rp326-529k in 2026."),
    ("P13", "B2C counseling and therapy", "TJ", "B2C", "Revenue", "TJ-B2C-KON-*, TRP-*", "session",
     0.5, "Plan", 0.375, "Plan", 166, "Plan", [0, 0, 0, 0, 0, 0.08, 0.02, 0], 60, "Content team", 7, "Y", "Y", "Y", "Proof for EAP; couples", "Y",
     "Units: v15 B2C plan 2027 (2026 actual was 10 paid sessions). Fixed days = Instagram and content."),
    ("P14", "Couples and marriage check-up packages", "TJ", "B2C", "Revenue", "TJ-B2C-PSG, PMC, MC", "package",
     2.6, "Plan", 1.95, "Proxy", None, "Missing", [0, 0, 0, 0, 0, 0.2, 0.1, 0], 0, "", 21, "Y", "N", "Y", "B2C counseling", "Y",
     "Price: Pre-Marriage T2. Cost: estimate, 75% psychologist share."),
    ("P15", "Speaker / webinar", "TJ", "B2B", "Lead gen", "TJ-SPK-*", "2-hour session",
     round(2.65 * KORP, 1), "Plan", 2.65, "Plan", None, "Missing", [0.5, 0, 0.3, 0, 0.2, 0.3, 0.2, 0], 0, "", 7, "Y", "N", "Y", "Workshops, EAP", "N",
     "Price: Tier 2, catalog rule. Own webinars (Webinar Guru) are free lead generation."),
    ("P16", "Workshops (corporate, school, team-building)", "TJ", "B2B", "Revenue", "TJ-WS-01 to 05", "2-hour workshop",
     round(3.415 * KORP, 1), "Plan", 3.415, "Plan", None, "Missing", [1, 0, 0.3, 0, 0, 0.4, 0.3, 0], 0, "", 7, "Y", "N", "Y", "EAP, school tests", "N",
     "New 20 Nov 2026."),
    ("P17", "Public psychoeducation event (large)", "TJ", "B2C", "Awareness", "TJ-EDU-04", "event",
     24, "Plan", 8.15, "Plan", 4, "Plan", [4, 0, 0, 2, 3, 2, 1, 0], 0, "", 45, "N", "N", "Y", "B2C counseling", "N",
     "Units: v15 B2C mass events 2027."),
    ("P18", "Ticketed community event (Mamaya, Halloween, Deen Haus)", "TJ", "B2C", "Awareness", "TJ-EDU-05", "event",
     3.1, "Plan", 1.76, "Plan", None, "Missing", [2, 0, 0, 1, 2, 1.5, 0.5, 0], 0, "", 30, "N", "N", "Y", "B2C counseling", "N",
     "Plan: 259k x 12 people. 2026 actual (Mamaya): 10 tickets, Rp2.37M."),
    ("P19", "Teman Kecil (children and parents)", "TJ", "B2C", "Awareness", "TJ-EXP-01, 02", "event (20 pairs)",
     round(4.095 * EDU, 1), "Plan", 4.095, "Plan", None, "Missing", [2, 0, 0, 0.5, 1, 1, 0.5, 0], 0, "", 30, "N", "N", "Y", "Child tests", "N",
     "New 20 Nov 2026. JLL proposal in progress."),
    ("P20", "Event service for partners", "TJ", "B2B", "Revenue", "TJ-EXP-03", "event",
     round(1.31 * EDU, 1), "Plan", 1.31, "Plan", None, "Missing", [1, 0, 0.5, 0, 0, 1, 0.5, 0], 0, "", 14, "Y", "N", "N", "", "N",
     "New 20 Nov 2026."),
    ("P21", "Competition judge (BI Duta, Jampurnas)", "TJ", "B2B", "Revenue", "TJ-EDU-01", "competition",
     15, "Plan", 7.4, "Plan", 1, "Plan", [1, 0, 0, 0, 0, 0.5, 0.5, 1], 0, "", 30, "Y", "Y", "N", "", "N",
     "Units: BI Jampurnas 2026 proposal."),
    ("P22", "Consulting day and FGD", "TJ", "B2B", "Revenue", "TJ-KNS-01, 02", "day",
     round(1.99 * KORP, 1), "Plan", 1.99, "Plan", None, "Missing", [1, 0, 0, 0, 0, 0, 0, 0], 0, "", 7, "Y", "N", "Y", "Employee mapping", "N",
     "New 20 Nov 2026. Senior psychologist."),
    ("P23", "Theta subscription (booking and admin for bureaus)", "Coreitera", "B2B", "Revenue", "CI-THT-01 to 04", "bureau-month",
     2.9, "Plan", 0.175, "Plan", None, "Missing", [0, 0.2, 0, 0, 0, 0.1, 0, 0], 30, "Galih", 30, "Y", "Y", "Y", "Synthesia copilot", "Y",
     "Price: Growth plan. Live in TJ 30 Nov 2026. Fixed days = sales and onboarding."),
    ("P24", "Synthesia copilot and white-label tests", "Coreitera", "B2B", "Revenue", "CI-SYN-*, CI-ASM-*", "seat-month",
     0.3, "Plan", 0.12, "Plan", None, "Missing", [0, 0.05, 0, 0, 0, 0, 0, 0], 40, "Galih", 30, "Y", "Y", "Y", "Theta Pro", "Y",
     "Planned 2028. Fixed days = build."),
]

# ---------------------------------------------------------------- past activity
# year, period, activity, client, product, status, units, revenue, rev label, cost, cost label, check, source
LI = "REKAP KONTEN LINKEDIN"
KS = "Data Konseling 2026"
PAST = [
    (2025, "2025", "Strategi Edukasi CBP Rupiah", "Bank Indonesia", "P01", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2025, "2025", "TDW Batch 2", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2025, "2025", "Penjurian Duta Guru dan Duta Muda", "Bank Indonesia", "P21", "Done", 1, None, "Missing", None, "Missing", "OK", LI),
    (2025, "2025", "Pelatihan Strategi Edukasi", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2025, "2025", "Pelatihan Kesejahteraan Psikologis Guru", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2025, "Sep 2025", "Young Health Summit (speaker)", "Universitas Yarsi", "P15", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2026, "2026", "Duta CBP Rupiah (selection)", "Bank Indonesia", "P21", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2026, "2026", "TDW 2 Bali", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "OK", LI),
    (2026, "2026", "TDW 3 Batam", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "OK", LI),
    (2026, "2026", "Duta Guru Manado", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2026, "2026", "Membangun Komunitas sebagai agen perubahan", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2026, "2026", "Strategi pendekatan komunitas ultra mikro", "Bank Indonesia", "P02", "Done", 1, None, "Missing", None, "Missing", "Check", LI),
    (2026, "Jun 2026 - now", "MHCU check-up for 25 people (one-off screening, still running)", "Telering", "P10", "Done", 0.5, None, "Missing", None, "Missing", "OK", "KPI Q4 Partnership: Closed Won"),
    (2026, "2026", "Fit-and-proper assessment, 6 people", "Bank Indonesia", "P07", "Done", 1, 15, "Actual", None, "Missing", "OK",
     "Galih, 23 Nov 2026: paid Rp15M. Planned cost Rp8.77M (RAB BI Fit and Proper); actual cost not recorded."),
    (2026, "Jan 2026", "Counseling vouchers, 3 sessions", "BI KPw Sibolga", "P04", "Done", 1, None, "Missing", None, "Missing", "Check", KS + " (Januari)"),
    (2026, "Jan-Feb 2026", "Career Compass event at Mamaya", "Public (venue Mamaya)", "P18", "Done", 1, 2.366, "Actual", 1.2005, "Actual", "OK",
     "BUDGET file: 10 paid tickets; variable cost excl. Rp1.8M ads (ads and equipment are in Acquisition_Effort)"),
    (2026, "19 Sep 2026", "Asesmen potensi-minat group (quota 6)", "Public", "P12", "Unclear", None, None, "Missing", None, "Missing", "Check", "KPI Q4 calendar"),
    (2026, "Sep 2026", "Asesmen penjurusan group (quota 5)", "Public", "P12", "Unclear", None, None, "Missing", None, "Missing", "Check", "KPI Q4 calendar"),
    (2026, "17 Oct 2026", "Career Maxxing", "Public", "P18", "Unclear", None, None, "Missing", None, "Missing", "Check", "Leads campaign list, content plan"),
    (2026, "31 Oct 2026", "Halloween Special", "TMG Hotel, Lyfe at 30s", "P18", "Done", 1, None, "Missing", None, "Missing", "OK", "KPI Q4: LPJ and expense recap exist, not in folder"),
    (2026, "14 Nov 2026", "TJ x Deen Haus, Bintaro", "Deen Haus", "P18", "Unclear", None, None, "Missing", None, "Missing", "OK", "KPI Q4: to be cancelled if quota not met"),
    (2026, "2026", "Event in Tuban", "Unknown", "?", "Done", 1, None, "Missing", None, "Missing", "Check", "KPI Q4 Operations: expense recap made"),
    (2026, "28 Nov 2026", "Webinar Guru Sejahtera (own, free)", "Teachers", "P15", "Upcoming", None, None, "Missing", None, "Missing", "OK", "Content plan Nov"),
    (2026, "28 Nov 2026", "Innerchild offline event", "Public", "P18", "Upcoming", None, None, "Missing", None, "Missing", "OK", "Content plan Nov"),
    (2026, "Jan 2026", "Career Compass bundle with counseling", "B2C clients", "P12", "Done", 1, 0.326, "Actual", None, "Missing", "OK", KS),
    (2026, "Feb 2026", "Career Compass bundle with counseling", "B2C clients", "P12", "Done", 1, 0.529, "Actual", None, "Missing", "OK", KS),
    (2026, "Apr, Jul 2026", "Career Reading sessions (free)", "B2C clients", "P12", "Done", 4, 0, "Actual", None, "Missing", "OK", KS),
    (2026, "Feb 2026", "B2C counseling, paid", "B2C clients", "P13", "Done", 1, 0.414, "Actual", None, "Missing", "OK", KS),
    (2026, "May 2026", "B2C counseling, paid", "B2C clients", "P13", "Done", 1, 0.299, "Actual", None, "Missing", "OK", KS),
    (2026, "Jun 2026", "B2C counseling, paid", "B2C clients", "P13", "Done", 2, 0.598, "Actual", None, "Missing", "OK", KS),
    (2026, "Jul 2026", "B2C counseling, paid", "B2C clients", "P13", "Done", 2, 0.728, "Actual", None, "Missing", "OK", KS),
    (2026, "Sep 2026 sheet", "B2C counseling, no amount recorded", "B2C clients", "P13", "Unclear", 2, None, "Missing", None, "Missing", "Check", KS + ": dated November in the September sheet"),
    (2026, "Oct 2026", "B2C counseling, paid", "B2C clients", "P13", "Done", 4, 2.08, "Actual", None, "Missing", "OK", KS),
    (2026, "Nov 2026", "B2C counseling (not in recap)", "B2C clients", "P13", "Unclear", None, None, "Missing", None, "Missing", "Check", "KPI Ops reports 9 paying clients in a November week; recap sheet empty"),
]

# ---------------------------------------------------------------- acquisition effort 2026
CONTENT = [["2026-01", "BRAND", "Unrated", 6], ["2026-01", "P12", "High", 10], ["2026-01", "P12", "Low", 4], ["2026-01", "P13", "High", 4], ["2026-01", "P13", "Low", 2], ["2026-01", "P15", "Low", 1], ["2026-02", "BRAND", "High", 1], ["2026-02", "BRAND", "Low", 1], ["2026-02", "BRAND", "Unrated", 1], ["2026-02", "P12", "High", 3], ["2026-02", "P12", "Low", 10], ["2026-02", "P12", "Unrated", 1], ["2026-02", "P13", "High", 3], ["2026-02", "P13", "Low", 15], ["2026-03", "BRAND", "High", 4], ["2026-03", "BRAND", "Low", 17], ["2026-03", "P12", "Low", 2], ["2026-03", "P13", "High", 2], ["2026-03", "P13", "Low", 6], ["2026-04", "BRAND", "High", 4], ["2026-04", "BRAND", "Low", 16], ["2026-04", "P12", "Low", 2], ["2026-04", "P13", "High", 2], ["2026-04", "P13", "Low", 6], ["2026-06 to 10", "BRAND", "Unrated", 32], ["2026-06 to 10", "INT", "Unrated", 1], ["2026-06 to 10", "P10", "Unrated", 4], ["2026-06 to 10", "P12", "Unrated", 15], ["2026-06 to 10", "P13", "Unrated", 12], ["2026-06 to 10", "P18", "Unrated", 11], ["2026-06 to 10", "P19", "Unrated", 3], ["2026-11", "BRAND", "Unrated", 12], ["2026-11", "P04", "Unrated", 1], ["2026-11", "P12", "Unrated", 1], ["2026-11", "P13", "Unrated", 10], ["2026-11", "P15", "Unrated", 6], ["2026-11", "P18", "Unrated", 3]]
CSRC = {"2026-01": "Content Planning: Jan 2026", "2026-02": "Content Planning: Feb 2026", "2026-03": "Content Planning: Mar 2026",
        "2026-04": "Content Planning: Apr 2026", "2026-06 to 10": "Marketing Monitoring: ALWAYS ON", "2026-11": "Marketing Monitoring: Content Plan"}
PK = "KPI Q4 Partnership sheet"
# period, type, activity, product, role, count, one-off hours, cash, label, source
OTHER = [
    ("Q4 2026", "Outreach: full", "Jones Lang LaSalle: Teman Kecil art therapy, proposal sent", "P19", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: full", "Mr. DIY: MHCU, proposal sent", "P10", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: full", "Sekolah Pelita Bangsa: IQ test, proposal and RAB sent", "P11", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: full", "Bank Syariah Nasional: employee mapping, proposal sent", "P08", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("2026", "Outreach: full", "Telering: MHCU (won)", "P10", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "BI Sibolga: EAP, first contact and costing v1-v4", "P04", "Partnership / sales", 1, 6, 0, "Proxy", PK + "; SPKnProposalAnalysis/BI Sibolga"),
    ("Q4 2026", "Outreach: contact+meeting", "Kementerian Pariwisata: waiting for TJ proposal", "P04", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "Permata Bank: waiting for TJ proposal", "P04", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "BNI: waiting for TJ proposal", "P04", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "Bidakara Medical Centre: partner idea", "PARTNER", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "Jakarta Respiratory Clinic: partner idea", "PARTNER", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "FNS Lawfirm: partner idea", "PARTNER", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Q4 2026", "Outreach: contact+meeting", "Deen Haus: event partner and white-label idea", "P18", "Partnership / sales", 1, 0, 0, "Proxy", PK),
    ("Sep-Oct 2026", "Outreach: contact", "10 schools: assessment proposal sent, follow-ups", "P11", "PIC psychologist", 10, 6, 0, "Proxy", "KPI Q4 Operations"),
    ("Sep-Oct 2026", "Outreach: contact", "15 brands / cafes listed, 8 proposals for monthly events", "P18", "Lead marketing", 15, 6, 0, "Proxy", "KPI Q4 Marketing"),
    ("2026", "Outreach: full", "BI Jampurnas: jury proposal and RAB", "P21", "Kia", 1, 0, 0, "Proxy", "BI Jampurnas folder"),
    ("2026", "Outreach: full", "BI Fit and Proper: proposal and RAB (won)", "P07", "Kia", 1, 0, 0, "Proxy", "BI Fit and Proper folder"),
    ("Jan 2026", "Cash", "Meta ads for the Mamaya event", "P18", "", 0, 0, 1.8, "Actual", "BUDGET file"),
    ("Jan 2026", "Cash", "Event equipment: tent, pop-up table, banner, frames (reusable)", "P18", "", 0, 0, 2.697, "Actual", "BUDGET file, FIXED sheet incl. delivery fees"),
    ("Jan 2026", "Cash", "Linktree", "BRAND", "", 0, 0, 0.888, "Actual", "BUDGET file"),
    ("Feb-Nov 2026", "Cash", "Instagram ads, about Rp1M a month", "BRAND", "", 0, 0, 10.0, "Plan", "README / model v15: IG ads Rp1M today. Replace with real spend."),
]

# ---------------------------------------------------------------- styles
F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F, color="000000")
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
HEAD = Font(name=F, bold=True, color="FFFFFF")
TITLE = Font(name=F, bold=True, size=14)
NOTE = Font(name=F, italic=True, color="595959", size=9)
NAVY = PatternFill("solid", fgColor="1F3864")
DGREEN = PatternFill("solid", fgColor="375623")
METRIC = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
LABEL_FILL = {"Actual": "C6EFCE", "Plan": "DDEBF7", "Proxy": "FCE4D6", "Missing": "F8CBAD"}


def header(ws, row, cols, fill=NAVY, height=60):
    for j, (name, w) in enumerate(cols, 1):
        c = ws.cell(row, j, name)
        c.font = HEAD; c.fill = fill; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[row].height = height


def put(ws, r, j, v, font=BLACK, fmt=None, fill=None):
    c = ws.cell(r, j, v)
    c.font = font; c.border = BOX; c.alignment = WRAP
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    return c


def label_colours(ws, rng):
    for lab, col in LABEL_FILL.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))


def label_dv(ws, rng):
    dv = DataValidation(type="list", formula1='"Actual,Plan,Proxy,Missing"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)


def yn_dv(ws, rng):
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)


def main():
    wb = Workbook()

    # ============================================================ Read_Me
    rm = wb.active; rm.title = "Read_Me"; rm.sheet_view.showGridLines = False
    rm.column_dimensions["A"].width = 22; rm.column_dimensions["B"].width = 110
    lines = [
        ("Portfolio focus: Teman Journey + Coreitera", None),
        ("Built 23 Nov 2026. Money in Rp million. One row per product family (105 catalog items grouped into 24).", None),
        ("", None),
        ("THE QUESTION", "Which products earn their place when 7-8 part-time people run two companies, and which take effort without building anything?"),
        ("", None),
        ("SHEETS", None),
        ("Portfolio", "FORMULAS ONLY. Plan metrics, 2026 actuals, cost to win a client, the suggested call, and how solid the data is."),
        ("Settings", "Roles: pay (or value of time), days available a month, day rate. Effort assumptions for content and outreach."),
        ("Product_Breakdown", "YOU FILL: price, cost, units a year, person-days per role per unit, lead time, strategy flags. Everything else adds up from here."),
        ("Past_Activity", "What we actually delivered in 2025-2026, with revenue and cost where a record exists."),
        ("Acquisition_Effort", "What we spent in 2026 to win clients: every content post (mapped to a product by its CTA), outreach, ads, equipment."),
        ("", None),
        ("LABELS", "Every input has a label. Fix the weakest labels first."),
        ("Actual", "From a record: receipts, ticket sales, the counseling recap, the leads file."),
        ("Plan", "From our price list (Master_Produk / Kalkulator_Sales) or model v15."),
        ("Proxy", "An estimate from indirect data, e.g. content hours from High/Low effort, outreach hours per stage, person-days per role."),
        ("Missing", "No data yet. The formula falls back to Plan where it can and says so."),
        ("", None),
        ("COLOURS", "Blue text = input you may change. Black = formula, do not overwrite. Green = taken from another sheet."),
        ("", None),
        ("FILL FIRST", "1. Product_Breakdown: person-days per role for the products you actually sell (columns N-U)."),
        ("", "2. Settings: days each person really has a month (column E)."),
        ("", "3. Past_Activity: revenue and cost for the BI events (amounts from invoices / SPK) and the Halloween and Tuban events (LPJ)."),
        ("", "4. Acquisition_Effort: real Instagram ad spend for Feb-Nov 2026 (row marked Plan)."),
        ("", "5. Past_Activity: psychologist fee actually paid per counseling session (cost column)."),
        ("", None),
        ("PRIVACY", "No client names or contact details are copied here. The counseling and ticket files are aggregated per month."),
    ]
    for i, (a, b) in enumerate(lines, 1):
        rm.cell(i, 1, a).font = TITLE if i == 1 else (BOLD if a and a.isupper() else BLACK)
        if b:
            rm.cell(i, 2, b).font = BLACK; rm.cell(i, 2).alignment = WRAP
        if a in LABEL_FILL:
            rm.cell(i, 1).fill = PatternFill("solid", fgColor=LABEL_FILL[a])

    # ============================================================ Settings
    st = wb.create_sheet("Settings"); st.sheet_view.showGridLines = False
    st["A1"] = "Settings: roles, time and effort assumptions"; st["A1"].font = TITLE
    st["A2"] = "Day rate = pay (or value of time) a month / days available a month. It prices team time in the Portfolio."; st["A2"].font = NOTE
    st["A3"] = "ROLES"; st["A3"].font = BOLD
    header(st, 4, [("Role", 20), ("People", 40), ("Pay or value of time a month (Rp M)", 14), ("Label", 9),
                   ("Days available a month", 12), ("Label", 9), ("Day rate (Rp M / day)", 12), ("Usable days a year (after buffer)", 14), ("Source", 60)])
    for i, (role, people, pay, plab, days, src) in enumerate(ROLES):
        r = ROLE_ROW0 + i
        put(st, r, 1, role); put(st, r, 2, people)
        put(st, r, 3, pay, BLUE, "#,##0.0"); put(st, r, 4, plab, BLUE)
        put(st, r, 5, days, BLUE, "0"); put(st, r, 6, "Proxy", BLUE)
        put(st, r, 7, f"=IF(E{r}>0,C{r}/E{r},0)", fmt="#,##0.00")
        put(st, r, 8, f"=E{r}*12*(1-$C$17)", fmt="#,##0")
        put(st, r, 9, src, NOTE)
    RT = ROLE_ROW0 + len(ROLES)  # 13
    put(st, RT, 1, "Total", BOLD)
    put(st, RT, 3, f"=SUM(C{ROLE_ROW0}:C{RT-1})", BOLD, "#,##0.0")
    put(st, RT, 5, f"=SUM(E{ROLE_ROW0}:E{RT-1})", BOLD, "0")
    put(st, RT, 8, f"=SUM(H{ROLE_ROW0}:H{RT-1})", BOLD, "#,##0")
    label_colours(st, f"D{ROLE_ROW0}:D{RT-1}"); label_colours(st, f"F{ROLE_ROW0}:F{RT-1}")
    label_dv(st, f"D{ROLE_ROW0}:D{RT-1}"); label_dv(st, f"F{ROLE_ROW0}:F{RT-1}")

    st["A15"] = "GENERAL"; st["A15"].font = BOLD
    general = [  # row, label, value, fmt, label, source
        (16, "Fixed running cost a month (Rp M)", 31.69, "#,##0.00", "Plan", "Model v15 Pay_Plan row 24, Oct-26 (team, founders, marketing tools, office)."),
        (17, "Buffer kept free (share of team time)", 0.2, "0%", "Plan", "Protects against burnout: plan to 80% load."),
        (18, "Sales commission (share of gross profit, B2B)", 0.1, "0%", "Plan", "Catalog rule."),
        (19, "Hours in a working day", 8, "0", "Plan", ""),
        (20, "Hours per High-effort post", 4, "0.0", "Proxy", "Assumption. Difficulty is recorded only for Jan-Apr 2026."),
        (21, "Hours per Low-effort post", 1.5, "0.0", "Proxy", "Assumption."),
        (22, "Hours per unrated post", 2.5, "0.0", "Proxy", "Average of High and Low. Used for Jun-Nov 2026."),
        (23, "Hours: first contact", 0.5, "0.0", "Proxy", "Assumption."),
        (24, "Hours: meeting (incl. prep and travel)", 3, "0.0", "Proxy", "Assumption."),
        (25, "Hours: proposal and RAB", 6, "0.0", "Proxy", "Assumption."),
        (26, "Year of actuals", 2026, "0", "Plan", "Portfolio actual columns read this year from Past_Activity."),
    ]
    for r, lab, val, fmt, lb, src in general:
        put(st, r, 1, lab); st.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        put(st, r, 3, val, BLUE, fmt); put(st, r, 4, lb, BLUE); put(st, r, 9, src, NOTE)
    label_colours(st, "D16:D26"); label_dv(st, "D16:D26")
    put(st, 27, 1, "Blended day rate (Rp M / day)", BOLD); st.merge_cells("A27:B27")
    put(st, 27, 3, f"=C{RT}/E{RT}", BOLD, "#,##0.00")
    put(st, 28, 1, "Break-even contribution per team-day (Rp M)", BOLD); st.merge_cells("A28:B28")
    put(st, 28, 3, f"=C16*12/H{RT}", BOLD, "#,##0.00")
    put(st, 28, 9, "Each planned team-day must earn at least this to pay the fixed cost with the team at planned load.", NOTE)
    DAYRATE = [f"Settings!$G${ROLE_ROW0 + i}" for i in range(len(ROLES))]
    ROLE_NAMES = f"Settings!$A${ROLE_ROW0}:$A${RT-1}"
    ROLE_RATES = f"Settings!$G${ROLE_ROW0}:$G${RT-1}"

    # ============================================================ Product_Breakdown
    pb = wb.create_sheet("Product_Breakdown"); pb.sheet_view.showGridLines = False
    pb["A1"] = "Product breakdown: the inputs (fill this; the Portfolio adds it up)"; pb["A1"].font = TITLE
    pb["A2"] = ("Per unit unless stated. Person-days = days of OUR team (not freelance psychologist sessions, which are in variable cost), "
                "including selling, preparing, delivering and reporting."); pb["A2"].font = NOTE
    role_cols = [(f"Days per unit: {r[0]}", 9) for r in ROLES]
    cols = ([("ID", 6), ("Product family", 30), ("Business", 9), ("B2B / B2C", 7), ("Funnel role", 10), ("Catalog codes", 16), ("Unit", 14),
             ("Price per unit", 9), ("Label", 8), ("Variable cost per unit", 9), ("Label", 8), ("Units a year (plan)", 8), ("Label", 8)]
            + role_cols + [("Label (days)", 8), ("Fixed team-days a year (build, marketing)", 10), ("Fixed days: role", 13),
                           ("Lead time: calendar days from signing to done", 10), ("Can run in parallel with other work? (Y/N)", 10),
                           ("Repeat or recurring revenue? (Y/N)", 9), ("Leads to a bigger product? (Y/N)", 9), ("Leads to", 18),
                           ("Tech can do most of the delivery? (Y/N)", 9), ("Team energy 1-5 (team fills)", 8), ("Source / note", 50)])
    header(pb, 4, cols, height=78)
    pb.cell(3, 8, "MONEY").font = BOLD; pb.cell(3, 14, "WHO DOES IT: PERSON-DAYS PER UNIT").font = BOLD; pb.cell(3, 23, "TIME AND STRATEGY").font = BOLD
    R0 = 5
    for i, p in enumerate(P):
        r = R0 + i
        (pid, fam, bu, seg, fun, codes, unit, price, plab, cost, clab, units, ulab, days, fixed, frole, elapsed, par, rep, leads, into, tech, note) = p
        for j, v in enumerate([pid, fam, bu, seg, fun, codes, unit], 1):
            put(pb, r, j, v, BLUE if j in (4, 5) else BLACK)
        put(pb, r, 8, price, BLUE, "#,##0.00"); put(pb, r, 9, plab, BLUE)
        put(pb, r, 10, cost, BLUE, "#,##0.00"); put(pb, r, 11, clab, BLUE)
        put(pb, r, 12, units, BLUE, "#,##0"); put(pb, r, 13, ulab, BLUE)
        for k, d in enumerate(days):
            put(pb, r, 14 + k, d if d else None, BLUE, "0.0#")
        put(pb, r, 22, "Proxy", BLUE)
        put(pb, r, 23, fixed or None, BLUE, "0"); put(pb, r, 24, frole or None, BLUE)
        put(pb, r, 25, elapsed, BLUE, "0"); put(pb, r, 26, par, BLUE)
        put(pb, r, 27, rep, BLUE); put(pb, r, 28, leads, BLUE); put(pb, r, 29, into, BLUE); put(pb, r, 30, tech, BLUE)
        put(pb, r, 31, None, BLUE); put(pb, r, 32, note, NOTE)
    R1 = R0 + len(P) - 1
    for col in ("I", "K", "M", "V"):
        label_colours(pb, f"{col}{R0}:{col}{R1}"); label_dv(pb, f"{col}{R0}:{col}{R1}")
    yn_dv(pb, f"Z{R0}:AB{R1}"); yn_dv(pb, f"AD{R0}:AD{R1}")
    dvr = DataValidation(type="list", formula1=f"={ROLE_NAMES}", allow_blank=True)
    pb.add_data_validation(dvr); dvr.add(f"X{R0}:X{R1}")
    dvs = DataValidation(type="list", formula1='"B2B,B2C"'); pb.add_data_validation(dvs); dvs.add(f"D{R0}:D{R1}")
    dvf = DataValidation(type="list", formula1='"Awareness,Lead gen,Revenue,Retention"'); pb.add_data_validation(dvf); dvf.add(f"E{R0}:E{R1}")
    dve = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True); pb.add_data_validation(dve); dve.add(f"AE{R0}:AE{R1}")
    pb.freeze_panes = "C5"

    # ============================================================ Past_Activity
    pa = wb.create_sheet("Past_Activity"); pa.sheet_view.showGridLines = False
    pa["A1"] = "Past activity: what we actually delivered"; pa["A1"].font = TITLE
    pa["A2"] = ("One row per event, project or month of counseling. Revenue and cost in Rp M. Blank revenue = Missing: the Portfolio then "
                "uses the list price and marks it Plan. 'Check' = my mapping to a product family needs your confirmation."); pa["A2"].font = NOTE
    header(pa, 4, [("No", 5), ("Year", 6), ("Period", 13), ("Activity", 38), ("Client / partner", 22), ("Product ID", 8),
                   ("Product family", 28), ("B2B / B2C", 7), ("Status", 9), ("Units delivered", 8), ("Revenue (Rp M)", 9), ("Label", 8),
                   ("Variable cost (Rp M)", 9), ("Label", 8), ("Team-days used", 8), ("Label", 8), ("Mapping", 8), ("Source", 48)])
    for i, row in enumerate(PAST):
        r = 5 + i
        (yr, per, act, cli, pid, status, units, rev, rlab, cost, clab, chk, src) = row
        put(pa, r, 1, i + 1); put(pa, r, 2, yr, BLUE); put(pa, r, 3, per, BLUE); put(pa, r, 4, act, BLUE); put(pa, r, 5, cli, BLUE)
        put(pa, r, 6, pid, BLUE)
        put(pa, r, 7, f'=IFERROR(VLOOKUP(F{r},Product_Breakdown!$A${R0}:$B${R1},2,FALSE),"(not mapped)")', GREEN)
        put(pa, r, 8, f'=IFERROR(VLOOKUP(F{r},Product_Breakdown!$A${R0}:$D${R1},4,FALSE),"")', GREEN)
        put(pa, r, 9, status, BLUE); put(pa, r, 10, units, BLUE, "0.0#")
        put(pa, r, 11, rev, BLUE, "#,##0.000"); put(pa, r, 12, rlab, BLUE)
        put(pa, r, 13, cost, BLUE, "#,##0.000"); put(pa, r, 14, clab, BLUE)
        put(pa, r, 15, None, BLUE); put(pa, r, 16, "Missing", BLUE)
        put(pa, r, 17, chk, BLUE); put(pa, r, 18, src, NOTE)
    PA1 = 200
    for col in ("L", "N", "P"):
        label_colours(pa, f"{col}5:{col}{PA1}"); label_dv(pa, f"{col}5:{col}{PA1}")
    dvst = DataValidation(type="list", formula1='"Done,Upcoming,Cancelled,Unclear"'); pa.add_data_validation(dvst); dvst.add(f"I5:I{PA1}")
    pa.conditional_formatting.add(f"Q5:Q{PA1}", CellIsRule(operator="equal", formula=['"Check"'], fill=PatternFill("solid", fgColor="FFEB9C")))
    pa.freeze_panes = "E5"

    # ============================================================ Acquisition_Effort
    ae = wb.create_sheet("Acquisition_Effort"); ae.sheet_view.showGridLines = False
    ae["A1"] = "Acquisition effort 2026: what we spent to win clients"; ae["A1"].font = TITLE
    ae["A2"] = ("Content posts are mapped to a product by their CTA or topic (BRAND = no product CTA; it is shared across B2C products in the Portfolio). "
                "Hours per item come from Settings. Time cost = hours / 8 x the role's day rate."); ae["A2"].font = NOTE
    header(ae, 4, [("No", 5), ("Period", 12), ("Type", 18), ("Activity", 46), ("Product ID", 9), ("Role", 16), ("Count", 7),
                   ("Hours each", 8), ("One-off hours", 8), ("Total hours", 8), ("Time cost (Rp M)", 9), ("Cash (Rp M)", 9),
                   ("Total cost (Rp M)", 9), ("Label", 8), ("Source", 44)])
    rows = []
    for per, pid, diff, n in CONTENT:
        rows.append((per, f"Content {diff}", f"{n} posts ({diff.lower()} effort)", pid, "Content team", n, 0, 0, "Proxy", CSRC[per]))
    rows += OTHER
    for i, row in enumerate(rows):
        r = 5 + i
        (per, typ, act, pid, role, n, one, cash, lab, src) = row
        put(ae, r, 1, i + 1); put(ae, r, 2, per, BLUE); put(ae, r, 3, typ, BLUE); put(ae, r, 4, act, BLUE); put(ae, r, 5, pid, BLUE)
        put(ae, r, 6, role or None, BLUE); put(ae, r, 7, n or None, BLUE, "0")
        put(ae, r, 8, (f'=IF(C{r}="Content High",Settings!$C$20,IF(C{r}="Content Low",Settings!$C$21,IF(C{r}="Content Unrated",Settings!$C$22,'
                       f'IF(C{r}="Outreach: contact",Settings!$C$23,IF(C{r}="Outreach: contact+meeting",Settings!$C$23+Settings!$C$24,'
                       f'IF(C{r}="Outreach: full",Settings!$C$23+Settings!$C$24+Settings!$C$25,0))))))'), fmt="0.0")
        put(ae, r, 9, one or None, BLUE, "0.0")
        put(ae, r, 10, f"=N(G{r})*H{r}+N(I{r})", fmt="#,##0.0")
        put(ae, r, 11, f'=IF(F{r}="",0,J{r}/Settings!$C$19*IFERROR(INDEX({ROLE_RATES},MATCH(F{r},{ROLE_NAMES},0)),0))', fmt="#,##0.00")
        put(ae, r, 12, cash or None, BLUE, "#,##0.000")
        put(ae, r, 13, f"=K{r}+N(L{r})", fmt="#,##0.00")
        put(ae, r, 14, lab, BLUE); put(ae, r, 15, src, NOTE)
    AE1 = 4 + len(rows)
    AEMAX = 300
    label_colours(ae, f"N5:N{AE1}"); label_dv(ae, f"N5:N{AE1}")
    dvt = DataValidation(type="list", formula1='"Content High,Content Low,Content Unrated,Outreach: contact,Outreach: contact+meeting,Outreach: full,Cash"')
    ae.add_data_validation(dvt); dvt.add(f"C5:C{AE1}")
    dvr2 = DataValidation(type="list", formula1=f"={ROLE_NAMES}", allow_blank=True); ae.add_data_validation(dvr2); dvr2.add(f"F5:F{AE1}")
    t = AE1 + 1
    put(ae, t, 4, "Total 2026", BOLD)
    for col, j, fmt in (("J", 10, "#,##0"), ("K", 11, "#,##0.0"), ("L", 12, "#,##0.0"), ("M", 13, "#,##0.0")):
        put(ae, t, j, f"=SUM({col}5:{col}{AE1})", BOLD, fmt)
    put(ae, t + 1, 4, "of which BRAND (no product CTA)")
    put(ae, t + 1, 10, f'=SUMIFS(J5:J{AE1},E5:E{AE1},"BRAND")', fmt="#,##0")
    put(ae, t + 1, 13, f'=SUMIFS(M5:M{AE1},E5:E{AE1},"BRAND")', fmt="#,##0.0")
    put(ae, t + 2, 4, "Recorded content work in team-days | content team days available Jan-Nov | share")
    put(ae, t + 2, 10, f'=SUMIFS(J5:J{AE1},F5:F{AE1},"Content team")/Settings!$C$19', fmt="#,##0")
    put(ae, t + 2, 11, f"=Settings!$E${ROLE_ROW0 + 4}*11", fmt="#,##0")
    put(ae, t + 2, 12, f"=J{t+2}/K{t+2}", BOLD, "0%")
    put(ae, t + 2, 15, "Share of the content team's 2026 time that the recorded posts explain (May 2026 has no calendar).", NOTE)

    f0 = t + 5
    ae.cell(f0, 4, "B2C FUNNEL 2026 (from records)").font = BOLD
    funnel = [
        ("Leads: people who contacted us (1 Jan - 20 Nov)", 175, "Actual", "Rekap LEADS, Ringkasan"),
        ("Filled the registration form", 34, "Actual", "Rekap LEADS, Ringkasan"),
        ("Paid counseling or Career Compass sessions (Jan-Oct)", 12, "Actual", KS),
        ("Paid event tickets (Mamaya)", 10, "Actual", "BUDGET file, PEMASUKAN EVENT"),
        ("Leads with no campaign recorded", 115, "Actual", "Rekap LEADS: 66% of leads cannot be traced to a post"),
    ]
    for k, (lab, v, lb, src) in enumerate(funnel):
        r = f0 + 1 + k
        put(ae, r, 4, lab); put(ae, r, 7, v, BLUE, "#,##0"); put(ae, r, 14, lb, BLUE); put(ae, r, 15, src, NOTE)
    put(ae, f0 + 6, 4, "Lead to form", BOLD); put(ae, f0 + 6, 7, f"=G{f0+2}/G{f0+1}", BOLD, "0%")
    put(ae, f0 + 7, 4, "Lead to paid (sessions + tickets)", BOLD); put(ae, f0 + 7, 7, f"=(G{f0+3}+G{f0+4})/G{f0+1}", BOLD, "0%")
    put(ae, f0 + 8, 4, "B2C acquisition cost 2026 per paid B2C sale (Rp M)", BOLD)
    put(ae, f0 + 8, 7, f'=SUMIFS(Portfolio!$W${R0}:$W${R1},Portfolio!$C${R0}:$C${R1},"B2C")/(G{f0+3}+G{f0+4})', BOLD, "#,##0.00")
    put(ae, f0 + 9, 4, "B2C revenue 2026 recorded per paid B2C sale (Rp M)", BOLD)
    put(ae, f0 + 9, 7, f'=SUMIFS(Portfolio!$R${R0}:$R${R1},Portfolio!$C${R0}:$C${R1},"B2C")/(G{f0+3}+G{f0+4})', BOLD, "#,##0.00")
    label_colours(ae, f"N{f0+1}:N{f0+5}")
    ae.freeze_panes = "E5"

    # ============================================================ Portfolio
    po = wb.create_sheet("Portfolio"); po.sheet_view.showGridLines = False
    po["A1"] = "Portfolio: what each product earns, costs, takes and needs to win"; po["A1"].font = TITLE
    po["A2"] = ("Formulas only. PLAN = a normal year at our prices. ACTUAL 2026 = what the records show. Team time is priced at each role's day rate "
                "(Settings). Last column says how solid each row's data is."); po["A2"].font = NOTE
    for c, g in [(1, "PRODUCT"), (5, "PLAN: A NORMAL YEAR"), (17, "ACTUAL 2026: DELIVERED, AND WHAT IT TOOK TO WIN"), (26, "DECISION")]:
        po.cell(3, c, g).font = BOLD
    pcols = [("ID", 6), ("Product family", 30), ("B2B / B2C", 7), ("Funnel role", 10),
             ("Units a year", 8), ("Revenue a year", 10), ("Variable cost a year", 10), ("Contribution a year (after B2B commission)", 11),
             ("Contribution margin", 9), ("Team-days per unit", 8), ("Team-days a year", 9), ("Kia days a year", 8),
             ("Team time cost a year (day rates)", 10), ("Profit after team time", 10), ("CONTRIBUTION PER TEAM-DAY", 11), ("Lead time (days)", 8),
             ("Deliveries done", 8), ("Revenue: actual", 9), ("Revenue: deliveries w/o record, at list price (Plan)", 11),
             ("Acquisition hours (direct)", 9), ("Acquisition cost: direct (time + cash)", 10), ("Acquisition cost: share of BRAND content and ads", 10),
             ("ACQUISITION COST 2026", 10), ("COST TO WIN ONE SALE (CAC)", 10), ("Acquisition cost / revenue", 9),
             ("Strategy score (0-3)", 8), ("Suggested call", 34), ("Data confidence", 52)]
    header(po, 4, pcols, height=80)
    for j in range(5, 17):
        po.cell(4, j).fill = DGREEN
    for j in (15, 23, 24):
        po.cell(4, j).fill = PatternFill("solid", fgColor="7F6000")
    PBR = "Product_Breakdown!"
    for i in range(len(P)):
        r = R0 + i
        g = lambda col, r=r: f"{PBR}{col}{r}"
        put(po, r, 1, f"={g('A')}", GREEN); put(po, r, 2, f"={g('B')}", GREEN); put(po, r, 3, f"={g('D')}", GREEN); put(po, r, 4, f"={g('E')}", GREEN)
        put(po, r, 5, f'=IF({g("L")}="","",{g("L")})', GREEN, "#,##0")
        put(po, r, 6, f'=IF(E{r}="","",E{r}*{g("H")})', fmt="#,##0.0")
        put(po, r, 7, f'=IF(E{r}="","",E{r}*{g("J")})', fmt="#,##0.0")
        comm = f'IF(C{r}="B2B",1-Settings!$C$18,1)'
        put(po, r, 8, f'=IF(E{r}="","",(F{r}-G{r})*{comm})', fmt="#,##0.0")
        put(po, r, 9, f'=IF(N({g("H")})=0,"",({g("H")}-{g("J")})*{comm}/{g("H")})', fmt="0%")
        put(po, r, 10, f"=SUM({PBR}N{r}:U{r})", fmt="0.0#")
        put(po, r, 11, f'=IF(E{r}="","",E{r}*J{r}+N({g("W")}))', fmt="#,##0")
        put(po, r, 12, f'=IF(E{r}="","",E{r}*N({g("N")})+IF({g("X")}="Kia",N({g("W")}),0))', fmt="#,##0")
        unit_tc = "+".join(f"N({PBR}{L(14 + k)}{r})*{DAYRATE[k]}" for k in range(len(ROLES)))
        fixed_tc = f'N({g("W")})*IFERROR(INDEX({ROLE_RATES},MATCH({g("X")},{ROLE_NAMES},0)),0)'
        put(po, r, 13, f'=IF(E{r}="","",E{r}*({unit_tc})+{fixed_tc})', fmt="#,##0.0")
        put(po, r, 14, f'=IF(E{r}="","",H{r}-M{r})', fmt="#,##0.0")
        put(po, r, 15, f'=IF(AND(E{r}<>"",N(K{r})>0),H{r}/K{r},IF(J{r}>0,({g("H")}-{g("J")})*{comm}/J{r},""))', BOLD, "#,##0.00", METRIC)
        put(po, r, 16, f"={g('Y')}", GREEN, "0")
        yr = "Settings!$C$26"
        crit = f'Past_Activity!$F$5:$F${PA1},A{r},Past_Activity!$B$5:$B${PA1},{yr}'
        put(po, r, 17, f'=SUMIFS(Past_Activity!$J$5:$J${PA1},{crit},Past_Activity!$I$5:$I${PA1},"Done")', fmt="0.#")
        put(po, r, 18, f'=SUMIFS(Past_Activity!$K$5:$K${PA1},{crit},Past_Activity!$L$5:$L${PA1},"Actual")', fmt="#,##0.00")
        put(po, r, 19, f'=SUMIFS(Past_Activity!$J$5:$J${PA1},{crit},Past_Activity!$I$5:$I${PA1},"Done",Past_Activity!$L$5:$L${PA1},"Missing")*{g("H")}', fmt="#,##0.0")
        put(po, r, 20, f"=SUMIFS(Acquisition_Effort!$J$5:$J${AEMAX},Acquisition_Effort!$E$5:$E${AEMAX},A{r})", fmt="#,##0")
        put(po, r, 21, f"=SUMIFS(Acquisition_Effort!$M$5:$M${AEMAX},Acquisition_Effort!$E$5:$E${AEMAX},A{r})", fmt="#,##0.00")
        put(po, r, 22, (f'=IF(C{r}="B2C",IFERROR(SUMIFS(Acquisition_Effort!$M$5:$M${AEMAX},Acquisition_Effort!$E$5:$E${AEMAX},"BRAND")'
                        f'*T{r}/SUMIFS($T${R0}:$T${R1},$C${R0}:$C${R1},"B2C"),0),0)'), fmt="#,##0.00")
        put(po, r, 23, f"=U{r}+V{r}", BOLD, "#,##0.00", METRIC)
        put(po, r, 24, f'=IF(Q{r}>0,W{r}/Q{r},IF(W{r}>0,"no sale yet",""))', BOLD, "#,##0.00", METRIC)
        put(po, r, 25, f'=IF(R{r}+S{r}>0,W{r}/(R{r}+S{r}),"")', fmt="0%")
        put(po, r, 26, f'=COUNTIF({g("AA")},"Y")+COUNTIF({g("AB")},"Y")+COUNTIF({g("AD")},"Y")', fmt="0")
        um = f'({g("H")}-{g("J")})*{comm}'
        call = (f'=IF({um}<0,"Stop or reprice: loses money on every unit",'
                f'IF(AND(Q{r}=0,W{r}>=1,E{r}=""),"Effort without sales: pause promotion until one sells",'
                f'IF(AND(E{r}="",Q{r}=0),"Not proven: sell once before promoting it",'
                f'IF(D{r}="Awareness","Keep only if it brings paying clients",'
                f'IF(AND(O{r}>=Settings!$C$28,Z{r}>=2),"Grow",'
                f'IF(O{r}>=Settings!$C$28,"Keep for cash, do not expand",'
                f'IF(Z{r}>=2,"Fix: raise price, cut days or automate","Stop or fold into another product")))))))'
                f'&IF(AND(ISNUMBER(X{r}),X{r}>{um}),"  | costs more to win than one sale earns","")'
                f'&IF(AND({g("AE")}<>"",N({g("AE")})<=2),"  | drains the team","")')
        put(po, r, 27, call)
        conf = (f'="Price "&{g("I")}&" · Cost "&{g("K")}&" · Units "&{g("M")}&" · Days "&{g("V")}'
                f'&" · 2026 revenue "&IF(R{r}>0,IF(S{r}>0,"part Actual","Actual"),IF(Q{r}>0,"Missing (list price used)","no deliveries"))')
        put(po, r, 28, conf, NOTE)
    tr = R1 + 1
    put(po, tr, 2, "Total", BOLD)
    for col in ("F", "G", "H", "K", "L", "M", "N", "Q", "R", "S", "T", "U", "V", "W"):
        put(po, tr, " ABCDEFGHIJKLMNOPQRSTUVW".index(col), f"=SUM({col}{R0}:{col}{R1})", BOLD, "#,##0.0")
    put(po, tr, 15, f'=IF(N(K{tr})>0,H{tr}/K{tr},"")', BOLD, "#,##0.00")

    po.conditional_formatting.add(f"O{R0}:O{R1}", FormulaRule(formula=[f'AND(ISNUMBER(O{R0}),O{R0}<Settings!$C$28)'], fill=PatternFill("solid", fgColor="F8CBAD")))
    for start, colr in (("Grow", "C6EFCE"), ("Stop", "F8CBAD"), ("Fix", "FFEB9C"), ("Effort", "F8CBAD"), ("Keep", "DDEBF7")):
        po.conditional_formatting.add(f"AA{R0}:AA{R1}", FormulaRule(formula=[f'LEFT(AA{R0},{len(start)})="{start}"'], fill=PatternFill("solid", fgColor=colr)))

    # load check by role
    lc = tr + 3
    po.cell(lc, 1, "LOAD CHECK BY ROLE (plan year vs time available after buffer)").font = BOLD
    for j, name in enumerate(["Role", None, "Days needed a year", "Days available a year", "Load"], 1):
        c = po.cell(lc + 1, j, name); c.font = HEAD; c.fill = NAVY; c.alignment = Alignment(wrap_text=True)
    po.merge_cells(start_row=lc + 1, start_column=1, end_row=lc + 1, end_column=2)
    for k in range(len(ROLES)):
        r = lc + 2 + k
        col = L(14 + k)
        put(po, r, 1, f"=Settings!A{ROLE_ROW0 + k}", GREEN); po.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        put(po, r, 3, (f"=SUMPRODUCT(Product_Breakdown!$L${R0}:$L${R1},Product_Breakdown!{col}${R0}:{col}${R1})"
                       f"+SUMIFS(Product_Breakdown!$W${R0}:$W${R1},Product_Breakdown!$X${R0}:$X${R1},A{r})"), fmt="#,##0")
        put(po, r, 4, f"=Settings!H{ROLE_ROW0 + k}", GREEN, "#,##0")
        put(po, r, 5, f"=IF(D{r}>0,C{r}/D{r},0)", BOLD, "0%")
    le = lc + 2 + len(ROLES)
    put(po, le, 1, "All roles", BOLD); po.merge_cells(start_row=le, start_column=1, end_row=le, end_column=2)
    put(po, le, 3, f"=SUM(C{lc+2}:C{le-1})", BOLD, "#,##0"); put(po, le, 4, f"=SUM(D{lc+2}:D{le-1})", BOLD, "#,##0")
    put(po, le, 5, f"=C{le}/D{le}", BOLD, "0%")
    po.conditional_formatting.add(f"E{lc+2}:E{le}", CellIsRule(operator="greaterThan", formula=["1"], fill=PatternFill("solid", fgColor="F8CBAD")))
    po.conditional_formatting.add(f"E{lc+2}:E{le}", CellIsRule(operator="between", formula=["0.85", "1"], fill=PatternFill("solid", fgColor="FFEB9C")))
    po.cell(le + 1, 1, "Only products with 'Units a year' filled count. Above 100% = more work than the planned time: the burnout signal.").font = NOTE

    s0 = le + 3
    po.cell(s0, 1, "PLAN YEAR: PROFIT CHECK").font = BOLD
    for k, (lab, fml, fmt, font) in enumerate([
        ("Contribution a year (products with units)", f"=H{tr}", "#,##0.0", BLACK),
        ("Fixed running cost a year", "=Settings!C16*12", "#,##0.0", GREEN),
        ("Left after fixed cost", f"=C{s0+1}-C{s0+2}", "#,##0.0", BOLD),
        ("Break-even contribution per team-day", "=Settings!C28", "#,##0.00", GREEN),
    ]):
        r = s0 + 1 + k
        put(po, r, 1, lab, BOLD if font is BOLD else BLACK); po.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        put(po, r, 3, fml, font, fmt)

    po.freeze_panes = "C5"
    wb.move_sheet("Portfolio", offset=-4)  # Read_Me, Portfolio, Settings, ...
    wb.save(OUT)
    print(OUT, "products", R0, R1, "past rows", len(PAST), "effort rows", len(rows), "AE total row", t, "funnel", f0, "load", lc)


if __name__ == "__main__":
    main()
