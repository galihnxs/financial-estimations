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
AVG_PLAN_BY_YEAR = {2027: 1.4, 2028: 1.6, 2029: 1.7, 2030: 1.75, 2031: 1.8}   # E
PSY_PER_BUREAU = 3          # E  average psychologists per paying bureau
PAYMENTS_PER_BUREAU = {2027: 50, 2028: 60, 2029: 70, 2030: 75, 2031: 80}   # E  paid bookings/orders a month

# ---- Mission and impact (Augment proposal IS5007 Oct 2026 + CoreItera trial) -------
MISSION = "CoreItera: enabling scalable mental health support worldwide, starting with Indonesia's psychology bureaus."  # K (Galih, 19 Oct 2026)
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
BUREAU_TARGETS = {          # P  Galih's goal: 700 bureaus in the first year (benchmark: Vibefam, 700+ gym locations)
    (2027, 2): 0, (2027, 3): 20, (2027, 6): 150, (2027, 9): 400, (2027, 12): 700,
    (2028, 12): 1000, (2029, 12): 1150, (2030, 12): 1300, (2031, 12): 1450,
}
MONTHLY_CHURN = 0.015       # E  1.5% of paying bureaus leave each month
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
    "inquiries": 200,          # E  chats from Instagram and WhatsApp a month
    "after_hours_share": 0.40, # E  share that arrive when the admin is not working
    "conversion_now": 0.15,    # E  share of chats that become paying clients today
    "conversion_agent": 0.20,  # E  with instant replies 24 hours and admin closing hot leads
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
WORKSHOPS_FROM = (2027, 3)
CAC_PER_NEW_BUREAU = 1.0   # E  ads, referral credit and sales commission for each new bureau (the free first month is on top)
GA_SHARE_OF_REVENUE = 0.08 # E  admin, legal, HR, tools and payment operations as the company grows   # P  after the first paying bureaus
FOUNDER_PAY_EACH = {2028: 10.0, 2029: 10.0, 2030: 20.0, 2031: 25.0}       # E  3 founders (Galih, Adit, Zaki)
FOUNDERS = 3
FOUNDER_PAY_TRIGGER = 60.0  # P/E founders paid once monthly revenue reaches Rp40M (ratchet)
CS_PER_BUREAUS = 80         # E  one customer-success person per 50 paying bureaus
CS_FROM_BUREAUS = 25        # E  founders onboard bureaus themselves until 25
CS_PAY = 7.0                # E  incl. BPJS and THR
SALES_FROM_BUREAUS = 25     # E  first sales person once 25 bureaus pay (founders sell until then)
SALES_PAY = 8.0             # E
SALES_EVERY_BUREAUS = 150    # E  one more sales person per 80 paying bureaus
ENGINEER_EVERY_REVENUE = 150.0   # E  one more engineer for every Rp150M of monthly revenue (beyond the founding three)
ENGINEER_PAY = 15.0         # E
OFFICE_TRIGGER = 200.0      # E  small Jakarta office once monthly revenue reaches Rp200M
OFFICE = 10.0               # E
INFRA_GROWTH_PER_YEAR = 0.15  # E  base cloud/tools cost grows 15% a year from 2028

# ---- Tax (confirm with a tax consultant) -----------------------------------------------
FINAL_TAX_RATE = 0.005      # E  PP 55/2022: 0.5% of revenue while eligible (PT: first 4 tax years)
FINAL_TAX_UNTIL = 2029
PROFIT_TAX_RATE = 0.11      # E  22% corporate tax with the 50% small-business cut, from 2030

USE_NEXT_ROUND = True       # P  the 700-bureau goal needs money before bureau income catches up

# ---- Exit and return for Riskia's side ---------------------------------------------------
NEXT_ROUND = {"month": (2027, 3), "amount": 150.0, "dilution": 0.10}   # E/P  growth round after the phase 1 gate: needed for the 700-bureau push (Galih to decide)
REVENUE_MULTIPLE = {"low": 2.0, "mid": 4.0, "high": 6.0}   # E  company value = x times yearly revenue in 2031
