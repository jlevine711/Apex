# Queenstown Harbor — Land Development Partnership

A pitch deck prepared for **Robert A. Connell, CFP** (Apex Financial Advisors /
Accountable Equity · Capital H6), proposing **JAL Strategies** as development &
capital partner to **monetize the developable land around Queenstown Harbor** —
entitle the non-conservation acreage and sell finished lots to homebuilders
(builders fund the horizontal infrastructure), then recycle the proceeds.

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
npm run build      # writes both the .pptx and the .pdf
```

`generate.js` records the deck once (a small render-agnostic proxy) and replays
it to **both** [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) and
[PDFKit](https://pdfkit.org/), so the two outputs stay identical:

- `JAL_Queenstown_Harbor_Proposal.pptx`
- `JAL_Queenstown_Harbor_Proposal.pdf`

The PDF embeds the brand fonts from `fonts/` (Montserrat & DM Sans, both
[SIL OFL](https://openfontlicense.org/)); if those files are absent it falls
back to Helvetica.

## Slides

1. Cover
2. The Opportunity — you just bought it; now monetize the land
3. The Asset — Queenstown Harbor, by the numbers
4. Provenance & History — four centuries to a 36-hole resort
5. The Land Bank — developable vs. conservation acreage
6. The Playbook — entitle → sell to builders → golf/water premium → recycle
7. The Math — 700 ac → developable envelope → lot value *(illustrative)*
8. Why It Sells — frontage, Bay Bridge demand, builder appetite
9. South River — the sister H6 asset (Annapolis-side private club)
10. Capital — capital-light by design
11. Capital Access — JAL's debt & equity network
12. The Structure — a land vehicle under Capital H6
13. Engagement & Compensation — hybrid, skin in the game
14. Why JAL — Justin A. Levine
15. The Path — diligence → July site visit → mandate
16. Thank You

## Verified facts (public sources)

- 36-hole waterfront resort, **310 Links Lane, Queenstown, MD** (Eastern Shore,
  Chester River → Chesapeake Bay)
- **River Course** (Lindsay Ervin, par 72, 7,096 yds, 1991) +
  **Lakes Course** (par 71, 6,569 yds, 1996)
- **Provenance:** land traces to a **1600s** colonial grant to surveyor
  **Henry deCoursey** ("My Lord's Gift" / the "thumb grant," from Cecil Calvert,
  2nd Lord Baltimore); developed from a **1970s dairy farm** by **The Brick
  Companies** (founded **1892**), which held it ~30 years until the 2026 sale
  (a construction-era archaeological dig preserved artifacts now with the
  Maryland Historical Trust)
- **River House & Tavern** event venue + on-site cottages
- Acquired from **The Brick Companies** via **Capital H6** (May 2026, ~$25M,
  with The Golf Club at South River); ~**870+ acres** across the two properties,
  incl. **198 acres** under permanent conservation
- ~1 hr to D.C. / Baltimore; minutes over the Bay Bridge to Annapolis
- **The Golf Club at South River** (Edgewater, MD): private 18-hole course
  (Brian Ault, 1996), ~**165 acres**, **600+ member families**, 8 lakes + 14
  protected areas, clubhouse (The Bistro), fitness, simulator lounge, range

Sources: qhgolf.com (incl. /who-we-are), thebrickcompanies.com, Where2Golf,
VisitMaryland.org, Club + Resort Business, Connect CRE, Eye On Annapolis,
GolfDigest, golfclubsr.com, foretee.com, Queen Anne's County, MD DNR.

Each data slide (2, 3, 4, 6, 7, 8) carries a small **hyperlinked source
footnote** just above the footer, citing where its figures come from.

## Accuracy note

Figures on **slide 6 (THE MATH)** are **illustrative** but now *derived from
real constraints* rather than guessed: ~700 gross acres, less the **198-acre**
conservation easement, less the golf/range/lodging footprint (~300 ac), less
Maryland's **Chesapeake Bay Critical Area** buffer and tidal wetlands (~60 ac)
≈ **~140 developable acres** → ~280 lots → ≈ $29M net. Recalibrate against
survey, easement and entitlement (Queen Anne's County) data before sharing
externally.

> Confidential — for discussion purposes only.
