# Queenstown Harbor — Land Development Partnership

A pitch deck prepared for **Robert A. Connell, CFP** (Apex Financial Advisors /
Accountable Equity · Capital H6), proposing **JAL Strategies** as development &
capital partner to **monetize the developable land around Queenstown Harbor** —
entitle and develop the non-conservation acreage and sell **finished lots** to
homebuilders (the JV funds the horizontal; builders just build the homes), then
recycle the proceeds.

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

`model.js` writes a backup financial model — `JAL_Queenstown_Harbor_Model.xlsx`
— with **live formulas** (every output references the INPUTS block) so it shows
exactly how each number is derived. Outputs:

- `JAL_Queenstown_Harbor_Proposal.pptx` · `JAL_Queenstown_Harbor_Proposal.pdf`
- `JAL_Queenstown_Harbor_Model.xlsx`

## Slides

1. Cover
2. The Opportunity — you just bought it; now monetize the land
3. The Asset — Queenstown Harbor, by the numbers
4. The Land Bank — developable vs. conservation acreage
5. The Playbook — entitle → develop finished lots → premium → recycle
6. The Math — 700 ac → ~280 finished lots → $42M revenue *(illustrative)*
7. Why It Sells — frontage, Bay Bridge demand, builder appetite
8. South River — the sister H6 asset (Annapolis-side private club)
9. Development Plan — entitle, develop, deliver finished lots
10. Pro Forma & Returns — finished-lot budget, profit, IRR / equity multiple *(illustrative)*
11. Basis & Sources — every pro-forma assumption, its reasoning and source
12. Buyer Universe / Liquidity — named regional + national lot buyers
13. Capital Access — JAL's debt & equity network
14. The Structure — a land vehicle under Capital H6
15. Engagement & Compensation — hybrid, skin in the game
16. Why JAL — Justin A. Levine
17. The Path — diligence → July site visit → mandate
18. Thank You

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

Each data slide (2, 3, 4, 6, 7, 8, 11) carries a small **hyperlinked source
footnote** just above the footer, citing where its figures come from. The
**Basis & Sources** slide (11) backs every pro-forma assumption with its
reasoning and a citation.

## Accuracy note

The model is a **finished-lot development** (slides 6, 9–11): the JV funds the
horizontal — roads, sewer, utilities — and sells **finished** lots; the
homebuilder just builds the homes. Illustrative but derived and internally
consistent:

- ~700 gross acres − **198-ac** conservation − golf/range/lodging (~300 ac) −
  **Chesapeake Bay Critical Area** + wetlands (~60 ac) ≈ **~140 developable
  acres** → **~280 finished lots**
- 280 × **$150K** finished lot = **$42.0M** revenue
- less **$27.5M** cost (incl. **$16.8M** JV-funded horizontal at $60K/lot) =
  **~$14.5M** profit → **~2.3× equity**, **~24% project IRR**

`JAL_Queenstown_Harbor_Model.xlsx` is the live backup model behind these
numbers. Only the 198-ac easement and the 9-ac range are sourced; the 700 ac is
the owner's verbal figure and the haircuts, density, price and cost loads are
estimates — recalibrate against survey, easement and entitlement (Queen Anne's
County) data before sharing externally.

> Confidential — for discussion purposes only.
