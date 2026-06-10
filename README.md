# Queenstown Harbor — Land Development Partnership

A pitch deck prepared for **Robert A. Connell, CFP** (Apex Financial Advisors /
Accountable Equity · Capital H6), proposing **JAL Strategies** as development &
capital partner to **monetize the developable land** across the Capital H6
golf-course portfolio (Queenstown Harbor first). **Capital-light entitled-lot
basis:** entitle the non-conservation acreage and sell **entitled lots** to
builders — the builder funds the horizontal. Costs are repaid, then profit
splits a simple **50/50** — half returning capital to the investors, half to
the development company. (Reflects Bob's 6/9 call: capital-light, a straight
50/50, no waterfall, and a repeatable pipeline — Queenstown is opportunity 1
(this proposal); Renault is opportunity 2 and the in-person site visit.)

Queenstown Harbor is the 36-hole Eastern Shore waterfront resort that
Accountable Equity / VIVÂMEE acquired via **Capital H6** in May 2026 (with The
Golf Club at South River, ~$25M, from The Brick Companies).

It reuses the visual design system of JAL's *Republic Square* deck:

- 16:9 (13.33" × 7.5")
- Montserrat (headings) / DM Sans (body)
- Palette — navy `#0B163C`, aubergine `#3A243A`, plum `#6B5A6B`, mauve `#C4B8C4`,
  cream `#F4F2ED`, slate `#8C9BB5`

> **Status: discussion draft — not final.** Framed throughout as a preliminary
> discussion document; figures are compiled from public sources and independent
> research and are subject to change and confirmation in further conversations.

## Build

```bash
npm install
npm run all        # deck (.pptx + .pdf) AND the backup model (.xlsx)
# or individually:
npm run build      # JAL_Queenstown_Harbor_Proposal.pptx + .pdf
npm run model      # JAL_Queenstown_Harbor_Model.xlsx
```

`generate.js` records the deck once (a small render-agnostic proxy) and replays
it to **both** [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) and
[PDFKit](https://pdfkit.org/), so the `.pptx` and `.pdf` stay identical. The PDF
embeds the brand fonts from `fonts/` (Montserrat & DM Sans, both
[SIL OFL](https://openfontlicense.org/)); if absent it falls back to Helvetica.

**The deck is dynamically linked to the model.** `generate.js` does
`require("./model")` and renders **every model-derived figure** — totals, the
50/50 split, JAL's carry/all-in, the cash-flow table, and the sensitivity
grid — from the engine (via a small `MV` formatting layer). Nothing financial
is hardcoded in the deck, so the presentation and the workbook **cannot drift**:
change an input in `model.js` and the next `npm run all` updates both. (`model.js`
exports its engine and only writes the `.xlsx` when run directly.)

`model.js` writes a backup financial model — `JAL_Queenstown_Harbor_Model.xlsx`
— that is **fully dynamic and calculated**: a single `INPUTS` object drives a
pure calc engine (`calcP1`) that derives *every* figure — totals, the
year-by-year cash flow, the project IRR, the 50/50 split and the sensitivity grid.
There are no hardcoded results. In the workbook the **INPUTS block plus the
cash-flow timing vectors are the only typed numbers**; every other cell is a
**live Excel formula** that references them (and the file is flagged to
recalculate on open, so spreadsheet and engine always agree). Edit any input —
in the code or in the sheet — and everything recomputes. One sheet: the
**entitled-lot, capital-light** model with the straight 50/50 split and a
finished-lot comparison. Outputs:

- `JAL_Queenstown_Harbor_Proposal.pptx` · `JAL_Queenstown_Harbor_Proposal.pdf`
- `JAL_Queenstown_Harbor_Model.xlsx`

## Slides

1. Cover
2. The Opportunity — monetize the land; return capital to investors
3. The Asset — Queenstown Harbor, by the numbers
4. The Land Bank — developable vs. conservation acreage
5. The Playbook — entitle → sell entitled lots → premium → recycle
6. The Math — 700 ac → ~280 entitled lots → ~$19.6M revenue *(illustrative)*
7. Why It Sells — frontage, Bay Bridge demand, builder appetite
8. South River — the sister H6 asset (Annapolis-side private club)
9. Renault — Opportunity 2 & the site visit (NJ, near Atlantic City)
10. The Pipeline — one playbook across the portfolio (not a one-off)
11. Development Plan — entitle, sell entitled lots, stay capital-light
12. Pro Forma & the 50/50 — entitled-lot budget, profit, return on cost *(illustrative)*
13. Basis & Sources — every assumption, its reasoning and source
14. The Deal — a straight 50/50, no waterfall
15. Profit Sensitivity — net profit across entitled-lot price × lot yield
16. Buyer Universe / Liquidity — named regional + national lot buyers (incl. Cole)
17. Cash Flow — entitled-lot project cash flow, year by year *(illustrative)*
18. Capital Access — JAL's debt & equity network
19. The Structure — a development company under Capital H6 (50/50)
20. Engagement & Compensation — $15K/mo retainer + reimbursed expenses, 1% debt / 2% equity placement, 30% of the dev-co 50% (~$2.1M carry; ~$2.7M all-in)
21. Why JAL — Justin A. Levine
22. The Path — diligence → Renault site visit (late Jul / early Aug) → mandate
23. Thank You

## Verified facts (public sources)

- 36-hole waterfront resort, **310 Links Lane, Queenstown, MD** (Eastern Shore,
  Chester River → Chesapeake Bay)
- **River Course** (Lindsay Ervin, par 72, 7,096 yds, 1991) +
  **Lakes Course** (par 71, 6,569 yds, 1996)
- **River House & Tavern** event venue + on-site cottages
- Acquired from **The Brick Companies** via **Capital H6** (May 2026, ~$25M,
  with The Golf Club at South River); ~**870+ acres** across the two properties,
  incl. **198 acres** under permanent conservation
- ~1 hr to D.C. / Baltimore; minutes over the Bay Bridge to Annapolis
- **The Golf Club at South River** (Edgewater, MD): private 18-hole course
  (Brian Ault, 1996), ~**165 acres**, **600+ member families**, 8 lakes + 14
  protected areas, clubhouse (The Bistro), fitness, simulator lounge, range

Sources: qhgolf.com, VisitMaryland.org, Club + Resort Business, Connect CRE,
Eye On Annapolis, GolfDigest, golfclubsr.com, foretee.com, Queen Anne's County,
MD DNR, Redfin, Land.com, Zillow, HomeGuide.

Each data slide (2, 3, 4, 6, 7, 8, 13) carries a small **hyperlinked source
footnote** just above the footer, citing where its figures come from. The
**Basis & Sources** slide (13) backs every assumption with its
reasoning and a citation.

## Accuracy note

The model is a **capital-light entitled-lot development** (slides 6, 11–15): we
fund only the **soft costs** to entitle the land and sell **entitled lots**; the
**builder funds the horizontal** (roads, sewer, utilities). Illustrative but
derived and internally consistent:

- ~700 gross acres − **198-ac** conservation − golf/range/lodging (~300 ac) −
  **Chesapeake Bay Critical Area** + wetlands (~60 ac) ≈ **~140 developable
  acres** → **~280 entitled lots**
- 280 × **~$70K** entitled lot (≈ finished ~$150K less the builder's ~$80K to
  finish) = **~$19.6M** revenue
- less **~$5.3M** soft cost (land contributed at **~zero basis**; no horizontal —
  the builder funds it) = **~$14.3M** profit → **~2.7× on cost**, ~73% margin.
  Costs repaid, then a straight **50/50**: **~$7.2M** to the property (Josh
  ~$3.6M → returns capital to his investors against their 8% pref) and
  **~$7.2M** to the development company (**Accountable Equity** / Capital H6 ·
  Bob · JAL + partners). **No pref, no promote, no waterfall.** JAL earns **30%
  of the dev-co 50%** (≈ **~$2.1M**, ~15% of profit) plus a **$15K/mo** retainer
  (non-creditable floor) → **~$2.7M all-in** (~19% of profit, mostly contingent
  carry); placement fees only on outside capital raised; co-invest optional (TBD).
  Sensitivity (slide 15): net profit ~$8–22M across $55–85K entitled lots ×
  240–320 lots.
- **Why capital-light:** finishing the lots ourselves would gross more (~$42M)
  but needs **~5× the capital** (~$25M) and all the horizontal execution +
  absorption risk for only a few $M more profit — so we stay light, and it's
  what the sponsor wants.
- **Year-by-year cash flow** (slide 17): the entitled-lot project cash flow nets
  ~$14.3M on ~$5.3M soft cost (~86% unlevered IRR on a near-zero land basis); it
  ties cell-for-cell to the model.

`JAL_Queenstown_Harbor_Model.xlsx` is the live backup model behind these
numbers — an **entitled-lot, capital-light** development: **~$14.3M profit on
~$5.3M cost, ~2.7× on cost**, split a straight 50/50 (no pref, no promote), with
a profit sensitivity over entitled-lot price × lot yield and a finished-lot
comparison. The **entitled-lot price (~$70K) is the key assumption** — confirm
against builder bids. The 700 ac / density / haircuts are estimates — recalibrate
against survey, easement and entitlement (Queen Anne's County) data before
sharing externally.

> *Note: an earlier hospitality/hotel "Phase 2" has been removed — this proposal
> is the land (entitled-lot) opportunity only.*

> Confidential — for discussion purposes only.
