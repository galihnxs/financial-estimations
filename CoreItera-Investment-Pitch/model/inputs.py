"""Every assumption of the CoreItera investment model, in one place.

Money is in Rupiah million (Rp M) unless the name says otherwise.
Status: K = known (Galih said it), P = Galih's plan / price decision, E = estimate (to confirm).
Change a number here, then run `python3 run.py` to rebuild output/results.json and output/tables.md.
"""

# ---- Timeline ---------------------------------------------------------------
START = (2026, 10)          # K  first model month: October 2026 (model opens 19 Oct)
END = (2031, 12)            # P  five full years after launch: 2027-2031

# ---- Cash and the deal -------------------------------------------------------
CASH_TODAY = 14.0           # K  CoreItera bank balance, 19 Oct 2026
OCT_DAYS_LEFT = 12 / 31     # K  share of October still to spend (19-31 Oct)
DEAL_CASH = 30.0            # K  new money: TemanJourney 20 + Riskia 10
DEAL_MONTH = (2026, 11)     # P  notary deed and payment by 13 Nov 2026
DEAL_ONE_OFFS = {"notary and share registration": 4.0,   # E
                 "launch costs": 2.9}                     # E
PAID_BEFORE = {"Galih": 50.0, "Riskia": 20.0}            # K

# Ownership (end state agreed in chat, 19 Oct 2026)
OWN_BEFORE = {"current owners": 0.90, "Riskia": 0.10}                    # K
OWN_AFTER = {"current owners": 0.80, "Riskia": 0.1608, "TemanJourney": 0.0392}  # P (cap_table_calc.py)
POST_MONEY = 510.0          # P  Rp M, price of the new shares (pre-money 480)

# ---- Running cost today -------------------------------------------------------
BASE_RUN_COST = 6.6         # K  cloud, AI, tools; nobody is paid a salary

# ---- Prices (Galih, 19 Oct 2026) ----------------------------------------------
ADMIN_COST = 5.73          # K  one full-time admin at Jakarta minimum wage (UMP DKI 2026: Rp5,729,876)
PLAN_TIERS = [              # P  Galih, 19 Oct 2026: lowest plan Rp1M, top plan = the cost of one admin; never more
    ("Starter", 1, 3, 1.0),
    ("Growth", 4, 10, 2.9),  #    about half an admin
    ("Pro", 11, None, 5.73), #    one admin, however many psychologists
]
PAYMENT_FEE = 0.005         # K  Rp5,000 per paid transaction; Midtrans fee passed on at cost
COPILOT_SEAT = 0.3          # K  Rp300,000 per psychologist seat per month
COPILOT_LAUNCH = (2028, 1)  # P  phase 2: copilot sold only after the proof gate (20 paying bureaus, Dec 2027)

# Average plan price per paying bureau (mix of tiers). Small bureaus join first.
AVG_PLAN_BY_YEAR = {2026: 1.0, 2027: 1.4, 2028: 1.6, 2029: 1.7, 2030: 1.75, 2031: 1.8}   # E
PSY_PER_BUREAU = 3          # E  average psychologists per paying bureau
PAYMENTS_PER_BUREAU = {2026: 40, 2027: 50, 2028: 60, 2029: 70, 2030: 75, 2031: 80}   # E  paid bookings/orders a month

# ---- Mission and impact (Augment proposal IS5007 Oct 2026 + CoreItera trial) -------
VISION = "CoreItera is a global technology company that provides AI services for the next million responsible businesses."  # K (Galih, 20 Oct 2026)
FIRST_STEP = "We start by helping mental health firms grow from 10 to their first 1,000 clients in 2027, with minimum resources."  # K
MISSION = "2027 to 2031: reduce mental health stress by helping expert firms operate at scale."  # K
WORLD = {"people_with_condition_bn": 1.0,      # K  WHO 2025: over 1 billion people
         "lost_usd_trillion": 1.0,             # K  WHO 2026: US$1 trillion a year lost to depression and anxiety at work
         "workers_per_100k": 13}               # K  WHO Mental Health Atlas 2024, world median
ASEAN6 = {"people_m": 610, "psychologists": 17_800, "with_disorder_m": 70.7,   # K  UN WPP 2024; HIMPSI etc.; GBD 2021 ASEAN
          "assessments_per_year_m": 0.4}       # K  Augment A8: ~17,800 x 5 h a week / 10.5 h x 46 weeks
ID = {"clinical": 3_059, "years_to_close_gap": "16 to 26", "graduates": 300_000}   # K  HIMPSI 2026; Augment
ASSESS_PER_WEEK_NOW = 6      # K  Augment: one psychologist doing only full assessments (10.5 h each)
ASSESS_PER_WEEK_COPILOT = 15 # K  Augment: 2.2 h each with the copilot (10 min score check, 1 h review, 1 h feedback)
AGREEMENT_NOW = 0.27         # K  Rettew et al. 2009: routine diagnosis vs structured assessment (0 to 1)
AGREEMENT_TARGET = 0.60      # P  pilot target, "good" agreement with an independent senior psychologist
QUALITY_BEFORE, QUALITY_AFTER = 4.2, 5.5   # K  CoreItera trial, out of 6

# ---- Market (for share checks) --------------------------------------------------
PSYCHOLOGISTS_ID = 14_000   # K  HIMPSI 2026, via Augment proposal (3,059 clinical)
SHARE_IN_BUREAUS = 0.5      # E  share of psychologists in private practice (a bureau or their own practice)
# -> about 7,000 psychologists / 3 per practice = about 2,300 bureaus and practices (to confirm with HIMPSI)

# ---- Growth plan: paying bureaus at the end of each period ------------------------
# Net targets (after churn). TemanJourney is not counted: free year, then cost price.
BUREAU_TARGETS = {          # K  Galih, 20 Oct 2026: NEW paying bureaus per quarter; TemanJourney not counted (free year)
    # Q4 2026: 2 (with TemanJourney = 3 bureaus) | Q1 2027: +21 (7 a month) | Q2: +40 | Q3: +80 | Q4: +100
    (2026, 11): 0, (2026, 12): 2,
    (2027, 3): 23, (2027, 6): 63, (2027, 9): 143, (2027, 12): 243,
    (2028, 12): 700,                                  # year 2: 700 bureaus
    (2029, 12): 1000, (2030, 12): 1250, (2031, 12): 1450,   # Indonesia only; Malaysia or Singapore from year 3 not counted
}
STRETCH_TARGETS = {         # P  Galih's 700-in-year-1 goal: only with the growth round (scenario)
    (2026, 11): 0, (2026, 12): 1, (2027, 3): 20, (2027, 6): 150, (2027, 9): 400, (2027, 12): 700,
    (2028, 12): 1000, (2029, 12): 1150, (2030, 12): 1300, (2031, 12): 1450,
}
# ---- Churn: bureaus that stop paying (logo churn) ---------------------------------------
# Industry: small-business software loses 3% to 7% of customers a month (Vena 2025, from Recurly and Paddle data);
# Focus Digital (2025-26, 15 verticals) puts SMB SaaS at 6.4% a month. Good SMB software aims for under 2%.
# Most SMB churn happens in the first 90 days, so the first 3 paid months churn faster.
# Galih, 20 Oct 2026: we start high because we are new (8%), fall to the industry middle (5%) by the end of
# year 1, and reach the good-software level (2%) by the end of year 2. Straight line between the points.
CHURN_CURVE = {(2026, 12): 0.08, (2027, 12): 0.05, (2028, 12): 0.02}   # P  monthly churn; flat after the last point
MONTHLY_CHURN = 0.05        # K  industry middle, used only as a comparison
EARLY_CHURN_MULTIPLIER = 1.0   # the curve already includes "new product" churn, so no extra first-months factor
EARLY_MONTHS = 3
CHURN_SCENARIOS = {         # name: curve
    "Our plan: 8% falling to 2%": {(2026, 12): 0.08, (2027, 12): 0.05, (2028, 12): 0.02},
    "Stays at 5% (industry middle)": {(2026, 12): 0.05},
    "Stays at 8% (we never improve)": {(2026, 12): 0.08},
    "Faster: 6% falling to 2% in year 1": {(2026, 12): 0.06, (2027, 12): 0.02},
}
# New bureaus are Galih's GROSS adds per quarter in year 1; from 2028 the targets are paying bureaus at year end.
GROSS_ADDS = {(2026, 12): 2}
GROSS_ADDS.update({(2027, m): 7 for m in (1, 2, 3)})
GROSS_ADDS.update({(2027, m): 40 / 3 for m in (4, 5, 6)})
GROSS_ADDS.update({(2027, m): 80 / 3 for m in (7, 8, 9)})
GROSS_ADDS.update({(2027, m): 100 / 3 for m in (10, 11, 12)})
FREE_FIRST_MONTH = True     # P  a new bureau pays from its second month

# Psynthesia (psychologist copilot) seats from year 2 = adoption x psychologists in paying bureaus + psychologists working alone
COPILOT_ADOPTION = {(2028, 1): 0.10, (2028, 12): 0.40, (2029, 12): 0.55, (2030, 12): 0.60}   # E
SOLO_SEATS = {(2028, 1): 0, (2028, 12): 40, (2029, 12): 150, (2030, 12): 320, (2031, 12): 520}  # E

# ---- Sales agent on Instagram and WhatsApp (Galih, 19 Oct 2026) -------------------
# Included in the monthly plan. Answers 24 hours in the bureau's own writing style; qualifies leads
# and offers times and a payment link when the admin is away; hands hot leads to the admin to close
# when the admin is online.
AGENT_WHATSAPP_LAUNCH = (2027, 2)   # E  with WhatsApp follow-ups (Q1 2027)
AGENT_INSTAGRAM_LAUNCH = (2027, 5)  # E  Instagram messages need Meta app review first
AGENT_COST_PER_BUREAU = 0.10        # E  AI replies + WhatsApp template messages, about 300 chats a month

# Worked example for one bureau (per month) - estimates, to be measured at TemanJourney first
EXAMPLE = {
    "inquiries": 300,          # K/E  Galih: hundreds of leads a month
    "after_hours_share": 0.40, # E  share that arrive when the admin is not working
    "conversion_now": 0.05,    # K  Galih: closing rate today is under 5%
    "conversion_agent": 0.075, # E  with replies in minutes, 24 hours, and the admin closing hot leads (+50%)
    "first_purchase": 0.4,     # E  Rp400,000 average first purchase (session or assessment)
}

# ---- Variable costs -----------------------------------------------------------------
AI_COST_PER_SEAT = 0.12     # E  ~15 reports a month x Rp8,000 (target under US$0.50 a report)
SERVER_PER_BUREAU = 0.05    # E  extra cloud, storage, email/WhatsApp messages per bureau a month
ONBOARDING_PER_NEW_BUREAU = 0.3   # E  self-serve set-up in one day; a support call and data import
ONBOARDING_FOUNDERS_UNTIL = 25    # P  founders onboard the first 25 bureaus themselves
ONBOARDING_FOUNDERS_COST = 0.2    # E  transport only

# ---- Fixed cost steps (switch on once and never off) ----------------------------------
# Lean rule for 2027: a cost starts only when revenue can carry it.
ACCOUNTANT = 1.5            # E  small PT accountant/tax consultant
ACCOUNTANT_TRIGGER = 10.0   # P/E start once monthly revenue reaches Rp10M
CLINICAL_BY_YEAR = {2027: 3.0, 2028: 6.0, 2029: 10.0, 2030: 15.0, 2031: 20.0}   # E  psychologists who check copilot quality
CLINICAL_TRIGGER = 10.0     # P  until copilot income reaches Rp10M a month, TemanJourney psychologists review drafts as design partner
WORKSHOPS = {2027: 5.0, 2028: 10.0, 2029: 15.0, 2030: 20.0, 2031: 25.0}   # E  HIMPSI workshops, faculties, ads per month
WORKSHOPS_FROM = (2027, 7)   # P  workshops only once the business makes a profit
CAC_PER_NEW_BUREAU = 0.3   # E  ads and the referral credit for each new bureau
COMMISSION_FIRST_PAYMENT = 0.10   # K  Galih: the person who closes the bureau (shared CS or sales) gets 10% of its FIRST plan payment only
GA_SHARE_OF_REVENUE = {2027: 0.03, 2028: 0.08}  # E  admin, legal, HR, tools; founders do most of it in 2027   # P  after the first paying bureaus
# Pay (Galih, 19 Oct 2026). 2027: Rp10M a month each, and only in a month that makes a profit.
# From 2028: the full pay below, in months whose profit covers it (else Rp10M each if profit covers that).
FOUNDER_PAY_2027_MONTH = 10.0                                          # K  per person per month
FOUNDER_PAY_YEAR = {"Galih": 700.0, "Riskia": 700.0, "Adit": 500.0}   # K  Rp million a year, from 2028
CORE_TEAM_PAY_MONTH = {2027: 10.0, 2028: 20.0}                         # E  Zaki, design lead: same rule as founders (Galih to confirm)
PAID_PEOPLE_2027 = ["Galih", "Riskia", "Adit", "Zaki"]

# Team plan: the MOST people we would want at the end of each period (P, Galih to adjust).
# A hire only happens when last month's profit can pay that person (Galih: pay the first salary first).
# Support today = one customer-service person shared with TemanJourney, paid by TemanJourney; she also sells.
HIRE_ORDER = ["Sales", "Support", "Quality", "Engineers"]   # P  who is hired first when profit allows
TEAM_PLAN = {
    #              2027: Q1 Q2 Q3 Q4    then December of each year
    "Sales":       {(2027, 1): 0, (2027, 2): 1, (2027, 3): 2, (2027, 4): 3, 2028: 4, 2029: 5, 2030: 5, 2031: 6},
    "Support":     {(2027, 1): 0, (2027, 2): 0, (2027, 3): 1, (2027, 4): 2, 2028: 5, 2029: 8, 2030: 9, 2031: 10},
    "Engineers":   {(2027, 1): 0, (2027, 2): 1, (2027, 3): 2, (2027, 4): 2, 2028: 4, 2029: 5, 2030: 6, 2031: 7},
    "Quality":     {(2027, 1): 0, (2027, 2): 0, (2027, 3): 1, (2027, 4): 1, 2028: 2, 2029: 2, 2030: 3, 2031: 3},
}
TEAM_PAY = {               # E  Rp million a month incl. BPJS and THR
    "Sales":     {2027: 2.0, 2028: 3.0},    #    K  Galih: first sales salary Rp2M a month + 10% of first payment
    "Support":   {2027: 6.0, 2028: 7.0},
    "Engineers": {2027: 15.0, 2028: 20.0},  #    beyond Adit and Zaki
    "Quality":   {2027: 12.0, 2028: 15.0},  #    test engineer who checks every release and the sales agent's answers
}
OFFICE_TRIGGER = 200.0      # E  small Jakarta office once monthly revenue reaches Rp200M
OFFICE = 10.0               # E
INFRA_GROWTH_PER_YEAR = 0.15  # E  base cloud/tools cost grows 15% a year from 2028

# ---- Tax (confirm with a tax consultant) -----------------------------------------------
# Indonesian company income tax for PT CoreItera (checked 20 Oct 2026; a tax consultant must confirm):
# 1) Final 0.5% of turnover (PP 55/2022) only for gross turnover up to Rp4.8 billion a year, and for a PT only for
#    3 tax years from registration. PP 20/2026 (from April 2026): PTs are no longer NEW users of the 0.5% scheme;
#    PTs already using it keep it until their 3 years end.
# 2) Otherwise: 22% of taxable profit (UU HPP). Pasal 31E: if turnover is up to Rp50 billion, the share of profit that
#    matches the first Rp4.8 billion of turnover is taxed at half (11%).
# 3) Losses can be carried forward for 5 years. PTKP (the personal allowance) is for people, not for a PT:
#    it applies to each founder's own salary tax (PPh 21), which is inside their gross pay.
FINAL_TAX_RATE = 0.005
FINAL_TAX_TURNOVER_LIMIT = 4800.0   # Rp M a year
TAX_REGISTERED = "10 October 2025"  # K  Galih, 20 Oct 2026: PT CoreItera registered then, and uses the 0.5% scheme today
FINAL_TAX_LAST_YEAR = 2027          # K  a PT keeps 0.5% for 3 tax years from registration: 2025, 2026, 2027 (registered before PP 20/2026)
CORP_TAX_RATE = 0.22
CORP_TAX_CUT_TURNOVER = 4800.0      # Pasal 31E: 50% cut on the profit share of the first Rp4.8 billion of turnover
CORP_TAX_CUT_LIMIT = 50000.0        # ... only if turnover is up to Rp50 billion
LOSS_CARRY_YEARS = 5

USE_NEXT_ROUND = False      # P  Galih, 20 Oct 2026: the main plan runs on own income; the Rp150M round is an option

# ---- Year 3: first country outside Indonesia (Galih, 20 Oct 2026) ----------------
ABROAD_FROM = (2029, 1)     # P  Malaysia or Singapore
ABROAD_COST = 25.0          # E  Rp million a month: company set-up, local partner, data-law (PDPA) work, travel
# Revenue from abroad is NOT counted in the numbers (upside).

# ---- Exit and return for Riskia's side ---------------------------------------------------
NEXT_ROUND = {"month": (2027, 4), "amount": 150.0, "dilution": 0.10}   # P/E  OPTIONAL: only if we have traction by January 2027; closed by April 2027 at the earliest; price is an estimate
REVENUE_MULTIPLE = {"low": 2.0, "mid": 4.0, "high": 6.0}   # E  company value = x times yearly revenue in 2031
