# CoreItera investment pitch (Theta partnership)

Everything for the CoreItera pitch to Riskia and TemanJourney lives here: the model, the proposal page and the reference documents.
Money is in Rupiah million (Rp M) unless stated.

## The offer in one line

The next Rp30M (TemanJourney Rp20M, Riskia Rp10M; Rp50M in total from Riskia's side) for 20% of CoreItera together: Riskia 16.1%, TemanJourney 3.9%. Pre-money Rp480M, post-money Rp510M.

## Folders

| Path | What it is |
|---|---|
| `model/inputs.py` | Every assumption, marked K (known), P (Galih's plan or price decision) or E (estimate). **Edit numbers here only.** |
| `model/engine.py` | Monthly engine, October 2026 to December 2031: bureaus, seats, revenue, costs (fixed costs switch on by revenue or volume), tax, cash. |
| `model/run.py` | Runs the plan and the slow case (half the bureaus); writes `output/results.json` and `output/tables.md`. |
| `model/build_page.py` | Builds `proposal/theta-partnership-proposal.html` from the model (v7: pyramid order). Prose lives in this script; numbers come from the engine. |
| `model/check_slides.sh` | Headless-Chrome check: `./check_slides.sh` lists slides whose content does not fit the 1280x720 page; `./check_slides.sh shots b1 c2` writes PNGs of single slides to `output/shots/`. Run it after every change. |
| `model/viz.py` | Every visual (pyramid, bars, 24-hour band, product stack, roadmap, Gantt, flywheel, positioning map, cash and revenue charts), drawn to scale from model numbers. |
| `model/build_page_v5_ibs5116_order.py` | The older builder in the IS5116 11-section order, kept for reference. |
| `model/page_style.css` | Page styles (TemanJourney slate blue, CoreItera teal). |
| `model/cap_table_calc.py` | The two-step share issue (15% completed, then cash to 20% combined) and the cash-only alternative. |
| `output/` | Model results. Generated: do not edit by hand. `output/slides.json` = the deck outline (one entry per slide: section, title, lead, note, body HTML) for the later slide deck. |
| `proposal/` | The proposal page. Generated: do not edit by hand. Published at https://claude.ai/artifact/RfDM9C2LhJpqV5XF5L4KgN |
| `proposal/theta-partnership-proposal-present.html` | Offline presenter version (double-click to open): full HTML page, fonts embedded, presenter bar, ← → steps through 16 sections, F full screen, T light/dark. Built by `model/present.py`. |
| `model/fonts_embedded.css` | Gelasio, Source Sans 3 and IBM Plex Mono as embedded files, so the presenter version needs no internet. |
| `reference/` | IS5116 final project brief (the benchmark for a good proposal) and the Augment proposal (IS5007) that supplies the problem evidence. |

## Rebuild

```bash
cd model && python3 -I build_page.py
```

## Structure of the proposal (v9: deck-ready, one message per slide)

39 slides: cover, answer (pyramid), four numbers; A divider + A1, A2 (loop, cost, night chats), A3; B divider + B1 (pain to fix, MVP parts), B2 (how the agent works, value), B3 (positioning, flywheel); C divider + C1 (roadmap, build plan, team, risks), C2 (cash, angel round, churn, churn cases, revenue), C3 (deal, use of money); the ask; appendix divider + 9 detail slides. Every slide title is its message.

### Older structure (v7: MECE pyramid)

Mission → the answer (governing thought + pyramid picture) → A. The need (A1 world/ASEAN gap, A2 bureaus lose clients, A3 Indonesia market) → B. One product first (B1 year-1 bureau use case: easy to start, easy to run, tests included, led by the sales agent; B2 already built; B3 why we win) → C. Plan and deal (C1 roadmap with proof gates, one build at a time, team, risks; C2 money incl. the growth round; C3 the deal) → the ask → appendix (5-year P&L, costs, slow case, prices, five tests, launch workback, assumptions, sale value, alternative pricing, numbers to confirm). The IS5116 sections are all still covered.

## Facts from Galih (19 Oct 2026)

- Cash paid before: Galih Rp50M, Riskia Rp20M (for her 10%). Promised 15% is completed in this deal.
- CoreItera cash today Rp14M; running cost Rp6.6M a month; nobody paid a salary.
- Prices: bureau plan from Rp1M a month, tiered by psychologists; Rp5,000 per paid transaction on top of the Midtrans fee (passed on at cost); copilot Rp300,000 per seat a month, later.
- Theta live at TemanJourney on 30 November 2026. TemanJourney software free for 12 months, then cost price.

## Year-1 plan (Galih, 19 Oct 2026)

Year 1 = bureau use case only, TemanJourney first; goal 700 paying bureaus by Dec 2027 (benchmark Vibefam, 700+ fitness locations). Plans Rp1M (1 to 3 psychologists), Rp2.9M (4 to 10), Rp5.73M (11+, one admin at Jakarta UMP 2026). Psynthesia (report copilot) seats only from year 2 (Jan 2028). The 700 push needs a growth round (model: Rp150M for 10% in March 2027).

## Open numbers

Share count from the deed; how the earlier Rp50M and Rp20M are recorded; number of private bureaus and practices in Indonesia (about 2,300 is our estimate; the 700 goal is 30% of it); growth round size and price; notary cost; TemanJourney's other shareholders; tier boundaries.
