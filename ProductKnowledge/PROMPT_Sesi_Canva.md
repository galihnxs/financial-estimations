# Prompt for the next session (new Canva account)

Copy everything below the line into a new Claude Code session opened in
`/Users/galihmujiagung/Documents/GitHub/core-itera/financial-estimations`.
Fill in the Canva links first.

---

You are continuing work on one product knowledge system for **Teman Journey (TJ)**, a psychology bureau under
PT Indo Medika Resolusi, and **Coreitera** (Theta, Synthesia, white-label test engine). Our sales team struggles
because every proposal is a separate deck with different numbers. All output is in **Bahasa Indonesia**.

Work in four phases. **Stop at the end of each phase**, show me a short summary and wait for my OK before you start
the next one.

## 0. Read first (do not skip)

Read these local files completely before you open Canva:

| File | What it holds |
|---|---|
| `ProductKnowledge/00_Product_Line_Analysis.md` | Folder map, 6 product families, conflicts, every pricing decision |
| `ProductKnowledge/01_Daftar_Produk_Final.md` | Final list prices (69 products) |
| `ProductKnowledge/02_Template_Analysis.md` | Slide-by-slide review of the BI proposal template |
| `ProductKnowledge/03_Cerita_Produk_Lengkap.md` | The full story: buyers, journey, how to sell, proposed deck outline |
| `ProductKnowledge/Master_Produk_dan_Harga.xlsx` | Central sheet; built by `ProductKnowledge/build/build_master_sheet.py` |
| `ProductKnowledge/build/*.py` | Build scripts. Change numbers in the scripts, then rebuild; never hand-edit the outputs |
| `ProductKnowledge/Katalog_Produk_TJ_Coreitera.pptx` / `.pdf` | First deck draft (34 slides). I found it ugly: too many tables |

**Pricing rules already decided (do not reopen them unless I ask):**
- Margin rule only: net margin at least 30% after PPh 23 (2%) and 10% sales commission on gross profit. No "3x" rule.
- B2B counseling (EAP, MHCU) costs a flat Rp350k per session. B2C uses tier revenue share 60/65/65/70%.
- EAP packages keep today's prices ("Harga tetap"); those under 30% need Director approval when sold.
- MHCU Flash Rp500k and Deep Rp750k (fixed). MHCU test cost at most Rp7k. Jeda Tahunan promises MHCU for 70 employees.
- Status "[cek]": Sibolga, JAMPURNAS, Telering contracts. Discount approval levels are still a proposal.

## Phase 1. Inventory every Canva proposal

Canva designs to analyze (search the account too; these are the ones I know about):

- [paste Canva link 1]
- [paste Canva link 2]
- [paste Canva link 3]
- (add more)

For each design: read every page, then record client, date, products offered, scope (sessions, participants,
duration, deliverables), every price and how it is packaged.

The numbers in Canva will differ from our sheet. That is expected. The goal of this phase is to **find products,
packages, add-ons and scope promises that exist in Canva but not in our data**.

Output: `ProductKnowledge/04_Inventaris_Canva.md` with
1. a table of designs (name, link, client, date, purpose)
2. a product-by-product table: Canva name, mapped code in our Katalog (or "TIDAK ADA"), Canva price, our list price
3. a list of new products or variants not in our data, with the evidence (design + page)
4. reusable slides and visual components worth keeping for the new decks

Do **not** edit any Canva design in this phase.

## Phase 2. Resolve each number problem

For every row where Canva and our sheet disagree, classify it:

| Class | Meaning | Default action |
|---|---|---|
| A. Old price | Canva shows an older price | Our list price wins |
| B. Deal price | A discount for one client | Keep list price; note the discount as a precedent |
| C. Different scope | Same name, different content | Split into a new code or variant |
| D. Missing product | Not in our data | Build its cost lines with the same rates and rules, then price it |
| E. Unclear | Cannot tell | Ask me |

Rules: cost assumptions come only from the `Tarif_Biaya` rates in the sheet. If a new cost is needed, ask me for
it. Do not invent one. Check every proposed price against the 30% net margin rule and show the margin.

Output: section "Resolusi angka" in `04_Inventaris_Canva.md`, one row per problem with class, decision and reason.
Put class E items in one list of questions for me.

## Phase 3. Update the local plan

After I approve Phase 2:
1. Add new products and variants to `build/build_master_sheet.py` (Katalog, cost lines, sources in `Sumber_Konflik`).
2. Rebuild the workbook, recalculate it (`formulas` package; no LibreOffice on this Mac), and confirm zero formula
   errors.
3. Regenerate `01_Daftar_Produk_Final.md` with `build/export_product_list.py <recalculated.xlsx>`.
4. Update `00_Product_Line_Analysis.md` (decisions, new families) and `03_Cerita_Produk_Lengkap.md` (story, buyers,
   deck outline) so all four documents agree.
5. Report every changed number: old, new, why.

## Phase 4. Build the deliverables

Build both decks in **Canva**. **Never edit an original design.** Copy it first and work only on the copy. Use the
TJ look from the existing proposals: navy gradient, white text, sand cards, the TJ logo, photos of our
psychologists. Avoid table-heavy slides: use cards, big numbers, journeys and photos. Use the
`temanjourney-brand` skill for copy and visuals, and `humanizer-indonesia` to edit all Indonesian text.

### Deck A: client proposal (external, ready to send)
- Beautiful, story first: the client's problem, our approach, products as a journey, proof, then investment.
- Modular. Every product family has its own slides so sales can **hide** what the client doesn't need, change
  the client name and date, and send.
- Mark every editable field clearly (client name, logo, date, participant count, package choice).
- Show list prices only. No cost, margin, commission or internal rules.
- First slide (hidden): "Cara memakai deck ini", a checklist of what to change before sending.

### Deck B: product knowledge (internal only)
- For sales to understand every product: what it is, who buys it, when to offer it, entry product, price,
  objections and answers, what's included and what isn't.
- Selling rules: list price, minimum price, discount and approval, payment, tax.
- A section "Cara membaca file Excel": which sheet and which columns to check to understand the margin of a
  deal (for example Katalog: list price, minimum price, net margin, check column; Simulasi_Penawaran for a full
  offer). Use screenshots or clear callouts.
- Mark every page "INTERNAL".

### File C: sales spreadsheet (`ProductKnowledge/Kalkulator_Sales.xlsx`)
Build it from a new script `build/build_sales_sheet.py` that reads the master workbook, so both stay in sync.
- **No detailed cost breakdown.** Do not include `Rincian_Biaya`, `Tarif_Biaya` or any cost lines. If a cost per
  product is needed for the formulas, keep only one total per product on a hidden, protected sheet.
- Per product: list price, minimum price (lowest price that still gives 30% net margin), and the margin status.
- An offer builder: sales pick products, quantities and their proposed price or discount. The sheet shows the
  margin band (green: at or above 30%; yellow: needs approval; red: not allowed, following `Aturan_Penjualan`)
  and **their commission (10% of gross profit)**.
- Make the message obvious: a higher price means a higher margin for the company and more commission for sales. Add
  a small "if you give 10% less discount, your commission goes up by Rp X" line.
- A one-page Panduan: which cells to edit (blue), a worked example row.
- Arial, formulas only (no pasted results), zero formula errors after recalculation.

## Working rules
- Ask before any choice that changes a price or a promise to a client.
- Keep a short log of what you changed at the end of `04_Inventaris_Canva.md`.
- Don't commit or push unless I ask.
- At the end, give me: the Canva links of Deck A and Deck B, the path of the sales spreadsheet, and a list of
  open questions.
