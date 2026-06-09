/**
 * JAL Strategies — Queenstown Harbor Land Development Partnership
 * Prepared for Robert A. Connell, CFP (Apex Financial Advisors / Accountable Equity · Capital H6)
 *
 * Scope: monetize the developable LAND around Queenstown Harbor — the 36-hole
 * Eastern Shore resort that Accountable Equity / VIVÂMEE acquired via Capital H6
 * (May 2026, with The Golf Club at South River, $25M, from The Brick Companies).
 * Capital-light entitled-lot basis: entitle the non-conservation acreage and sell
 * ENTITLED lots to builders (the builder funds the horizontal). Costs repaid, then a
 * straight 50/50 — half returns capital to the investors, half to the development company.
 *
 * Design system reverse-engineered from the JAL "Republic Square" pitch deck:
 *   - 16:9 (13.33" x 7.5"); Montserrat (headings) / DM Sans (body)
 *   - Palette: navy #0B163C, aubergine #3A243A, plum #6B5A6B,
 *     mauve #C4B8C4, cream #F4F2ED, slate #8C9BB5, white
 *
 * Verified facts from public sources (qhgolf.com, VisitMaryland, Club+Resort
 * Business, Eye On Annapolis, GolfDigest). Figures on the "THE MATH" slide are
 * clearly labelled ILLUSTRATIVE — to be recalibrated against survey, conservation
 * easement and entitlement data.
 */

const pptxgen = require("pptxgenjs");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const nodePath = require("path");

// ----- Design tokens -------------------------------------------------------
const C = {
  navy: "0B163C",
  aubergine: "3A243A",
  plum: "6B5A6B",
  mauve: "C4B8C4",
  cream: "F4F2ED",
  slate: "8C9BB5",
  white: "FFFFFF",
  hair: "E6E1DB",
};
const HEAD = "Montserrat";
const BODY = "DM Sans";

const PAGE_W = 13.333;
const PAGE_H = 7.5;
const ML = 0.5;
const CW = 12.33;
const TOTAL = 26;
let PAGENO = 2; // auto page counter — cover is 1, chrome slides number themselves 2..N-1, Thank You is N

const DECK_LABEL = "QUEENSTOWN HARBOR  ·  CONFIDENTIAL  ·  DISCUSSION DRAFT";
const FOOTER_LEFT =
  "JAL Strategies  |  Justin A. Levine, Founder & CEO  |  jlevine@jalstrategies.com";

// ----- Render-agnostic deck proxy ------------------------------------------
// Content/helpers call the familiar pptxgenjs slide API (addText / addShape /
// addNotes / background). We record those calls and replay them to BOTH a
// pptxgenjs deck and a pdfkit document, so the .pptx and .pdf stay identical.
class SlideProxy {
  constructor() { this.calls = []; this.background = undefined; this.notes = ""; }
  addText(text, opts) { this.calls.push(["text", text, opts]); }
  addShape(type, opts) { this.calls.push(["shape", type, opts]); }
  addNotes(t) { this.notes = t; }
}
class Deck {
  constructor() { this.slides = []; this.props = {}; }
  defineLayout() {}
  set layout(v) {}
  set author(v) { this.props.author = v; }
  set company(v) { this.props.company = v; }
  set title(v) { this.props.title = v; }
  addSlide() { const s = new SlideProxy(); this.slides.push(s); return s; }
}

const pptx = new Deck();
pptx.defineLayout({ name: "W16", width: PAGE_W, height: PAGE_H });
pptx.layout = "W16";
pptx.author = "JAL Strategies";
pptx.company = "JAL Strategies";
pptx.title =
  "Queenstown Harbor — Land Development Partnership (for Robert A. Connell)";

// ----- Low-level helpers ---------------------------------------------------
const rect = (s, x, y, w, h, color, opts = {}) =>
  s.addShape("rect", { x, y, w, h, fill: { color }, line: { type: "none" }, ...opts });

const hline = (s, x, y, w, color, h = 0.014) =>
  s.addShape("rect", { x, y, w, h, fill: { color }, line: { type: "none" } });

const txt = (s, text, x, y, w, h, o = {}) =>
  s.addText(text, {
    x, y, w, h,
    fontFace: o.font || BODY,
    fontSize: o.size || 11,
    bold: !!o.bold,
    color: o.color || C.navy,
    align: o.align || "left",
    valign: o.valign || "middle",
    charSpacing: o.spc || 0,
    lineSpacing: o.lh || undefined,
    margin: o.margin || 0,
    wrap: o.wrap !== false,
  });

function chrome(s, { eyebrow, title, desc }) {
  const page = PAGENO++;
  s.background = { color: C.cream };
  rect(s, 0, 0, PAGE_W, 0.04, C.navy);
  txt(s, "JAL STRATEGIES", ML, 0.12, 3.0, 0.3, {
    font: HEAD, size: 9, bold: true, color: C.navy, spc: 3,
  });
  txt(s, DECK_LABEL, 6.0, 0.12, 6.83, 0.3, {
    font: HEAD, size: 9, color: C.plum, spc: 2, align: "right",
  });
  hline(s, ML, 7.05, 12.33, C.mauve);
  txt(s, FOOTER_LEFT, ML, 7.15, 9.0, 0.3, { font: BODY, size: 8, color: C.plum });
  txt(s, `${page} / ${TOTAL}`, 12.0, 7.15, 0.83, 0.3, {
    font: BODY, size: 8, color: C.plum, align: "right",
  });
  txt(s, eyebrow, ML, 0.5, 9.0, 0.3, {
    font: HEAD, size: 10, bold: true, color: C.aubergine, spc: 6,
  });
  txt(s, title, ML, 0.84, CW, 0.6, { font: HEAD, size: 25, bold: true, color: C.navy });
  if (desc)
    txt(s, desc, ML, 1.46, CW, 0.5, { font: BODY, size: 13, color: C.plum, lh: 16 });
}

function callout(s, text, y = 6.5) {
  rect(s, ML, y, CW, 0.46, C.aubergine);
  txt(s, text, ML + 0.2, y, CW - 0.4, 0.46, {
    font: BODY, size: 11, bold: true, color: C.white, align: "center",
  });
}

// Source map — verified public sources for the figures used in the deck.
const SRC = {
  cr:    { label: "Club + Resort Business", url: "https://clubandresortbusiness.com/vivamee-hospitality-acquires-two-maryland-golf-locations-for-25m/" },
  eoa:   { label: "Eye On Annapolis", url: "https://www.eyeonannapolis.net/2026/05/vivamee-hospitality-acquires-golf-club-at-south-river-queenstown-harbor-in-25m-deal/" },
  qh:    { label: "qhgolf.com", url: "https://qhgolf.com/golf/" },
  qhwho: { label: "qhgolf.com/who-we-are", url: "https://qhgolf.com/who-we-are/" },
  gd:    { label: "GolfDigest", url: "https://www.golfdigest.com/courses/md/queenstown-harbor-the-river" },
  vm:    { label: "VisitMaryland", url: "https://www.visitmaryland.org/listing/sports/queenstown-harbor-golf-links" },
  qac:   { label: "Queen Anne's Co. — Critical Area", url: "https://www.qac.org/398/Chesapeake-Bay-Critical-Area" },
  dnr:   { label: "MD DNR — Critical Area", url: "https://dnr.maryland.gov/criticalarea/Pages/compliance.aspx" },
  dist:  { label: "distance-cities.com", url: "https://www.distance-cities.com/distance-queenstown-md-to-washington-dc" },
  sr:    { label: "golfclubsr.com", url: "https://www.golfclubsr.com/faq" },
  ft:    { label: "foretee.com", url: "https://foretee.com/courses/maryland/edgewater/usa/the-golf-club-at-south-river/6925" },
  qacLots: { label: "Land.com — QAC lots", url: "https://www.land.com/Queen-Annes-County-MD/" },
  newcon:  { label: "Zillow — QAC new homes", url: "https://www.zillow.com/queen-annes-county-md/new-homes/" },
  devcost: { label: "HomeGuide — land dev cost", url: "https://homeguide.com/costs/cost-to-develop-land" },
  fillat:  { label: "FILLAT+ — resort design", url: "https://www.pfarc.com/queenstown-harbor-resort-spa/" },
  lax:     { label: "Lacrosse Homes (Centreville, QAC)", url: "https://www.laxhomes.com/communities/" },
  caruso:  { label: "Caruso Homes", url: "https://www.carusohomes.com/new-homes/md/maryland/floor-plans/" },
  gem:     { label: "Gemcraft Homes", url: "https://www.gemcrafthomes.com/" },
  zcom:    { label: "Zillow — MD new communities", url: "https://www.zillow.com/browse/communities/md/" },
  hvs:     { label: "HVS — hotel dev cost 2025", url: "https://www.hvs.com/article/10219-hvs-us-hotel-development-cost-survey-2025" },
  hcap:    { label: "hotel cap rates 2025", url: "https://www.mmcginvest.com/post/u-s-hotel-cap-rates-in-2025-trends-drivers-and-segment-analysis" },
  rcost:   { label: "restaurant build cost", url: "https://blog.eb3construction.com/construction/food-beverage/restaurant-build-out-estimate/" },
};

// Small hyperlinked source footnote, just above the footer rule.
function footnote(s, y, items) {
  const runs = [{ text: "Sources (own research; subject to change & confirmation):  ", options: { color: C.plum, bold: true } }];
  items.forEach((it, i) => {
    if (i) runs.push({ text: "    ·    ", options: { color: C.mauve } });
    runs.push({ text: it.label, options: { color: C.aubergine, underline: true, hyperlink: { url: it.url, tooltip: it.label } } });
  });
  s.addText(runs, { x: ML, y, w: CW, h: 0.26, fontFace: BODY, fontSize: 7, color: C.plum, align: "left", valign: "middle", margin: 0 });
}

function statCard(s, x, y, num, label, sub, o = {}) {
  const w = o.w || 6.0;
  const auto = num.length > 9 ? 19 : num.length > 6 ? 22 : 27;
  rect(s, x, y, w, 0.95, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, num, x + 0.2, y + 0.1, 2.2, 0.75, {
    font: HEAD, size: o.numSize || auto, bold: true, color: o.numColor || C.navy,
  });
  txt(s, label, x + 2.4, y + 0.12, w - 2.6, 0.36, {
    font: HEAD, size: 12, bold: true, color: C.navy,
  });
  txt(s, sub, x + 2.4, y + 0.5, w - 2.6, 0.36, { font: BODY, size: 10, color: C.plum });
}

function sidePanel(s, x, y, w, h, heading, lead, items) {
  rect(s, x, y, w, h, C.navy);
  rect(s, x, y, 0.12, h, C.aubergine);
  txt(s, heading, x + 0.4, y + 0.18, w - 0.6, 0.3, {
    font: HEAD, size: 10, bold: true, color: C.mauve, spc: 4,
  });
  txt(s, lead, x + 0.4, y + 0.56, w - 0.62, 0.82, {
    font: HEAD, size: 14.5, bold: true, color: C.white, lh: 18,
  });
  let iy = y + 1.5;
  const step = (h - 1.55) / items.length;
  items.forEach((it) => {
    txt(s, it, x + 0.4, iy, w - 0.62, step - 0.04, {
      font: BODY, size: 10, color: C.cream, lh: 12.5, valign: "top",
    });
    iy += step;
  });
}

function table(s, x, y, colDefs, header, rows, o = {}) {
  const rowH = o.rowH || 0.5;
  const headH = o.headH || 0.4;
  let cx = x;
  header.forEach((h, i) => {
    rect(s, cx, y, colDefs[i].w, headH, C.navy);
    txt(s, h, cx, y, colDefs[i].w, headH, {
      font: HEAD, size: o.headSize || 9, bold: true, color: C.white,
      align: colDefs[i].align || "left", margin: [2, 6, 2, 6],
    });
    cx += colDefs[i].w;
  });
  let ry = y + headH;
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? C.white : C.cream;
    let rx = x;
    row.forEach((cell, ci) => {
      const d = colDefs[ci];
      rect(s, rx, ry, d.w, rowH, fill);
      const cfg = typeof cell === "object" ? cell : { text: cell };
      txt(s, cfg.text, rx, ry, d.w, rowH, {
        font: cfg.font || d.font || BODY,
        size: cfg.size || d.size || 10,
        bold: cfg.bold !== undefined ? cfg.bold : d.bold || false,
        color: cfg.color || d.color || C.navy,
        align: cfg.align || d.align || "left",
        lh: cfg.lh || d.lh,
        margin: [2, 6, 2, 6],
        valign: "middle",
      });
      rx += d.w;
    });
    ry += rowH;
  });
  return ry;
}

// ============================================================ SLIDE 1 — COVER
function cover() {
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  rect(s, 0, 0, 0.25, PAGE_H, C.aubergine);
  txt(s, "JAL STRATEGIES", 0.75, 0.6, 6.0, 0.4, {
    font: HEAD, size: 13, bold: true, color: C.mauve, spc: 6,
  });
  txt(s, "Disciplined Real Estate Investing, Development & Capital Formation",
    0.75, 1.0, 9.5, 0.3, { font: BODY, size: 11, color: C.slate });

  txt(s, "QUEENSTOWN", 0.72, 2.18, 12.2, 1.05, {
    font: HEAD, size: 64, bold: true, color: C.white, spc: 1,
  });
  txt(s, "HARBOR", 0.72, 3.12, 12.2, 1.05, {
    font: HEAD, size: 64, bold: true, color: C.mauve, spc: 1,
  });

  hline(s, 0.78, 4.42, 3.5, C.slate, 0.03);
  txt(s, "Land Development & Lot-Monetization Partnership", 0.75, 4.58, 12.0, 0.45, {
    font: HEAD, size: 22, bold: true, color: C.white,
  });
  txt(s, "Entitling and monetizing the developable land around the River & Lakes courses",
    0.75, 5.12, 12.0, 0.35, { font: BODY, size: 14, color: C.slate });
  txt(s,
    "Queenstown, Maryland  ·  Eastern Shore  ·  36 holes on the Chesapeake",
    0.75, 5.52, 12.0, 0.35, { font: BODY, size: 12, color: C.mauve });
  txt(s,
    "Preliminary discussion document — not a final proposal or commitment. Figures compiled from public sources and independent research; subject to change and confirmation in further conversations.",
    0.75, 6.0, 11.6, 0.6, { font: BODY, size: 9.5, color: C.slate, lh: 12, valign: "top" });

  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, "PREPARED FOR ROBERT A. CONNELL, CFP   |   ACCOUNTABLE EQUITY · CAPITAL H6",
    0.5, 7.0, 9.6, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5 });
  txt(s, "JUNE 2026  ·  DISCUSSION DRAFT", 9.7, 7.0, 3.13, 0.4, {
    font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.2, align: "right",
  });
  s.addNotes(
    "Cover. Scoped to the land around Queenstown Harbor — the asset Capital H6 just acquired. " +
    "Bob is a GP in H6, so this is squarely his lane: entitle the developable acreage, sell lots, recycle."
  );
}

// ===================================================== SLIDE 2 — OPPORTUNITY
function opportunity() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE OPPORTUNITY",
    title: "Monetize the land. Return capital to your investors.",
    desc: "Bob — Capital H6 bought these courses with ~1,000 acres at a near-zero land basis. The resort runs itself; the upside is the land — entitle it and sell entitled lots to builders (they fund the horizontal), returning capital to the investors who are owed their pref.",
    page: 2,
  });
  const cards = [
    ["$25M", "Just acquired (May 2026)", "Queenstown + South River, via Capital H6"],
    ["~1,000 ac", "Near-zero land basis", "Land across the portfolio — that's the upside"],
    ["Capital-light", "Entitle & sell lots", "Builder funds the horizontal; we fund soft costs"],
    ["Liquidity", "For your investors", "Lot sales return capital against the 8% pref"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "WHY NOW",
    "Monetize the land to get your investors their capital back.",
    [
      "Near-zero basis: ~1,000 acres came with the golf-course buys — the land value is pure upside.",
      "Capital-light: entitle the parcels and sell entitled lots; the builder funds the roads, sewer & utilities.",
      "Your investors are owed an 8% pref and have seen no cash yet — lot sales return their capital.",
      "And it repeats — a steady pipeline of family-owned courses keeps coming to Josh.",
    ]);
  footnote(s, 6.66, [SRC.cr, SRC.eoa]);
  s.addNotes(
    "Reframed to Bob's own words: ~1,000 acres at no cost basis; the goal is liquidity for Josh's syndication " +
    "investors (owed an 8% pref, no cash yet). Capital-light: we entitle and sell entitled lots; the builder funds the horizontal."
  );
}

// ========================================================== SLIDE 3 — THE ASSET
function theAsset() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE ASSET",
    title: "Queenstown Harbor, by the numbers.",
    desc: "A 36-hole waterfront destination on Maryland's Eastern Shore — acquired from The Brick Companies in May 2026, with established golf, an event venue and cottages already in place.",
    page: 3,
  });
  const cards = [
    ["River", "Course · par 72", "Lindsay Ervin design · 7,096 yds · opened 1991"],
    ["Lakes", "Course · par 71", "Lindsay Ervin design · 6,569 yds · opened 1996"],
    ["310", "Links Lane, Queenstown MD", "Chester River at the Chesapeake Bay"],
    ["$25M", "Acquired with South River", "From The Brick Companies · via Capital H6"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "ALREADY IN PLACE",
    "An operating resort — not a raw land play.",
    [
      "River House & Tavern — newly renovated wedding, event and corporate venue.",
      "On-site cottages — existing overnight lodging and a hospitality base.",
      "36 holes of championship golf since 1991 — traffic, brand and demand already here.",
      "198 acres under permanent conservation — stewardship credibility, and an envelope to plan around.",
    ]);
  footnote(s, 6.66, [SRC.qh, SRC.gd, SRC.vm]);
  s.addNotes(
    "Shows the homework Bob asked for. Accurate specs: River (par 72, 7,096) + Lakes (par 71, 6,569), " +
    "Lindsay Ervin, River House & Tavern, cottages. Conservation acreage is real — flag it honestly."
  );
}

// ====================================================== SLIDE 4 — THE LAND BANK
function theLandBank() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE LAND BANK",
    title: "The upside is the land around the golf.",
    desc: "The resort operates; the value-creation is entitling the non-conservation acreage and selling entitled lots to builders — capital-light, the builder funds the horizontal — additive to the hospitality.",
    page: 4,
  });
  const cards = [
    ["870+ ac", "Portfolio land", "Across Queenstown + South River"],
    ["198 ac", "Under conservation", "Permanent easement at Queenstown — plan around it"],
    ["~140 ac", "Developable land", "Net of conservation, wetlands & course (to verify)"],
    ["~280", "Entitled lots", "Sold to regional homebuilders"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "THE THESIS",
    "Monetize the land the golf doesn't need.",
    [
      "Identify developable parcels at the course edges — away from conservation, wetlands and play.",
      "Entitle and sell entitled lots — the builder funds the horizontal (roads, sewer, utilities) and builds.",
      "Golf-frontage and water-view lots command premium pricing and faster absorption.",
      "Recycle proceeds into the broader Capital H6 program — including South River.",
    ]);
  footnote(s, 6.66, [SRC.cr, SRC.qhwho]);
  s.addNotes(
    "Honest framing: 870+ ac, minus 198 ac conservation (plus wetlands/course), leaves the developable " +
    "envelope — the first diligence item. The land bank is upside on top of an operating resort."
  );
}

// ======================================================== SLIDE 5 — PLAYBOOK
function playbook() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PLAYBOOK",
    title: "Four moves that turn land into cash.",
    desc: "Capital-light: entitle the land (soft costs only), sell entitled lots to builders who fund the horizontal, then recycle into the next parcel — and the next property.",
    page: 5,
  });
  const items = [
    ["01", "Entitle", "Approvals for lots",
      "Secure zoning and subdivision approvals for the developable acreage around the fairways — soft costs only."],
    ["02", "Sell Entitled Lots", "Entitled lots → builders",
      "Sell entitled (approved) lots to builders — they fund the horizontal (roads, sewer, utilities) and build the homes."],
    ["03", "Golf & Water Premium", "Frontage sells higher",
      "Lots on a course or with Chesapeake water views command a premium and absorb faster."],
    ["04", "Recycle & Repeat", "Parcel — then property",
      "Roll proceeds into the next parcel, and run the same playbook across the Capital H6 pipeline."],
  ];
  const cardW = 2.94, gap = 0.18;
  let x = ML;
  items.forEach(([num, head, sub, body]) => {
    rect(s, x, 2.1, cardW, 4.1, C.navy);
    rect(s, x, 2.1, cardW, 0.1, C.aubergine);
    txt(s, num, x + 0.25, 2.35, cardW - 0.5, 0.9, { font: HEAD, size: 44, bold: true, color: C.mauve });
    txt(s, head, x + 0.25, 3.35, cardW - 0.5, 0.4, { font: HEAD, size: 14, bold: true, color: C.white });
    txt(s, sub, x + 0.25, 3.78, cardW - 0.5, 0.3, { font: HEAD, size: 10, bold: true, color: C.slate, spc: 1 });
    hline(s, x + 0.25, 4.18, cardW - 0.5, C.aubergine, 0.02);
    txt(s, body, x + 0.25, 4.32, cardW - 0.5, 1.75, { font: BODY, size: 10.5, color: C.cream, lh: 14, valign: "top" });
    x += cardW + gap;
  });
  callout(s, "Move 01 is mostly soft costs — the fastest path to proof, and to first cash.", 6.5);
  s.addNotes(
    "Sequencing IS the pitch: entitle (soft costs) → sell to builders (they fund infra) → golf/water " +
    "premium → recycle. Matches Bob's 'let the developer handle the infrastructure.'"
  );
}

// ========================================== SLIDE 6 — THE MATH (ILLUSTRATIVE)
function mathLand() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE MATH  ·  ILLUSTRATIVE",
    title: "From 700 acres to ~$19.6M of entitled-lot revenue.",
    desc: "Capital-light entitled-lot basis: we entitle the land and sell entitled lots; the builder funds the horizontal. ~280 lots at ~$70K = ~$19.6M revenue. Net profit & the 50/50 split are on the pro forma. Illustrative — survey-dependent.",
    page: 6,
  });
  const cols = [
    { w: 3.2, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.4, color: C.navy, size: 9.5 },
    { w: 2.9, color: C.plum, size: 9, font: BODY },
    { w: 2.83, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "right" },
  ];
  const rows = [
    ["Gross Queenstown land", "per ownership (to confirm)", "—", "~700 ac"],
    ["Less permanent conservation", "recorded easement", "700 − 198", "502 ac"],
    ["Less golf, range, water & lodging", "two 18s + 9-ac range + cottages", "502 − ~300", "~200 ac"],
    ["Less Critical Area & wetlands", "Chesapeake 1,000-ft zone", "200 − ~60", "~140 ac"],
    ["Entitled lots", "~2.0 lots / acre", "140 × 2.0", "~280 lots"],
    ["Entitled lot price", "to a builder (they finish)", "≈ $150K − ~$80K", "$70K"],
    [{ text: "GROSS ENTITLED-LOT REVENUE", bold: true, font: HEAD, color: C.navy },
      { text: "net profit & 50/50 split on the pro forma", color: C.plum }, "280 × $70K", "$19.6M"],
  ];
  table(s, ML, 2.1, cols, ["Step", "Basis", "Calculation", "= Result"], rows, { rowH: 0.48 });
  rect(s, ML, 6.0, CW, 0.45, C.navy);
  txt(s,
    "~$19.6M entitled-lot revenue on ~$5.3M of soft cost → ~$14.3M profit — the builder funds the horizontal. Split & returns on the pro forma.",
    ML + 0.2, 6.0, CW - 0.4, 0.45,
    { font: HEAD, size: 11, bold: true, color: C.white, align: "center" });
  footnote(s, 6.62, [SRC.qhwho, SRC.qac, SRC.dnr]);
  s.addNotes(
    "Capital-light entitled-lot basis: we entitle and sell entitled lots; the builder funds the horizontal. 280 lots × $70K " +
    "= $19.6M revenue, ~$5.3M soft cost → ~$14.3M profit (~2.7× on cost). The entitled-lot price (~finished $150K less the " +
    "builder's ~$80K to finish) is the key assumption. Split & returns on the pro forma. Illustrative until survey / entitlement."
  );
}

// ===================================================== SLIDE 7 — WHY IT SELLS
function whySells() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "WHY IT SELLS",
    title: "These lots sell — here's why.",
    desc: "Course and water frontage, Bay Bridge proximity, and builder appetite for entitled lots all push price and absorption. We're meeting demand that already exists.",
    page: 7,
  });
  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10.5 },
    { w: 4.8, color: C.navy, size: 9.5, lh: 11 },
    { w: 4.53, color: C.plum, size: 9.5, lh: 11 },
  ];
  const rows = [
    ["Golf & water frontage", "Lots on the River / Lakes courses & Chester River",
      "Premium pricing and faster absorption vs. a raw subdivision"],
    ["Bay Bridge location", "Minutes to Annapolis; ~1 hr to D.C. & Baltimore",
      "Commuter, second-home and weekend demand from three metros"],
    ["Chesapeake scarcity", "Limited new waterfront-adjacent lots on the Shore",
      "High barriers → durable, premium lot values"],
    ["Builder appetite", "Regional homebuilders want entitled / finished lots",
      "They pay to avoid entitlement risk — and fund the infrastructure"],
    ["Fresh owned basis", "$25M for 870+ ac just closed", "Low land basis → development margin starts higher"],
  ];
  table(s, ML, 2.15, cols, ["Driver", "Detail", "Why it lifts value"], rows, { rowH: 0.62 });
  callout(s, "Demand is already here — we're entitling supply to meet it, not creating a market.", 6.5);
  footnote(s, 5.82, [SRC.dist, SRC.qac]);
  s.addNotes(
    "Mirrors the reference competitive-supply slide. Eastern-Shore + Bay-Bridge demand, Chesapeake scarcity, " +
    "builder appetite, and the fresh $25M basis all support lot sales."
  );
}

// ===================================================== SLIDE 8 — SOUTH RIVER
function southRiver() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE SISTER ASSET",
    title: "South River — the Annapolis-side anchor.",
    desc: "Acquired in the same $25M Capital H6 deal: an established 18-hole private club near Annapolis with 600+ member families — a membership and events engine, and the next place to run the playbook.",
    page: 8,
  });
  const cards = [
    ["18 holes", "Private club (Brian Ault, 1996)", "Edgewater, MD — on the South River"],
    ["600+", "Member families", "Recurring dues + a built-in events base"],
    ["165 ac", "Club grounds", "8 lakes · 14 environmentally protected areas"],
    ["Same deal", "Part of the $25M H6 buy", "One platform, two assets"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "THE ANGLE",
    "Membership and events first; land where diligence supports it.",
    [
      "600+ member families — recurring dues, plus an events / dining base to grow.",
      "Annapolis-side, west of the Bay — an affluent, supply-constrained catchment.",
      "The Bistro, fitness, simulator lounge & range already drive non-golf revenue.",
      "Assess limited parcel entitlement without disturbing the member experience.",
    ]);
  footnote(s, 6.66, [SRC.sr, SRC.ft, SRC.cr]);
  s.addNotes(
    "South River is the second H6 asset — but it's a built-out private club (165 ac, 8 lakes, 14 protected " +
    "areas, 600+ families). Frame honestly: membership + events upside first, land only where diligence supports it."
  );
}

// ===================================== SLIDE 9 — RENAULT (THE FIRST DEAL)
function renault() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE FIRST DEAL  ·  RENAULT WINERY",
    title: "Renault — the first deal, and the site visit.",
    desc: "Renault Winery (Egg Harbor City, NJ — ~25 min from Atlantic City) is the lead property and where we'll meet. The township has offered land for ~34 lots overlooking the course, a +100–200-room hotel expansion is planned, and the architect is already mapping lots.",
    page: 9,
  });
  const cards = [
    ["~34 lots", "Township-offered land", "Overlooking the golf course"],
    ["+100–200", "Hotel rooms planned", "Expansion off the back of the resort"],
    ["~25 min", "From Atlantic City", "Where I'll meet Bob & Josh"],
    ["In motion", "Architect mapping lots", "Underway since the last meeting"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "WHY RENAULT FIRST",
    "Josh's original property — the proof case for the playbook.",
    [
      "The township has offered additional land for ~34 lots; the architect is already boxing them out.",
      "A required hotel expansion (+100–200 rooms) layers hospitality on top of the lot sales.",
      "The winery already runs a lake & campground profitably — operating upside beyond golf.",
      "Closest asset to walk — ~25 minutes from Atlantic City airport; the natural place to start.",
    ]);
  s.addNotes(
    "Renault (NJ) is the lead / first deal and the site-visit location (Bob: 'an expansion down at Reno, which is where I'd meet you'). " +
    "Township offered ~34 lots; +100–200 hotel rooms planned; architect already mapping. The capital-light entitled-lot playbook applies here too."
  );
}

// ===================================== SLIDE 10 — THE PIPELINE
function pipeline() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PIPELINE  ·  NOT A ONE-OFF",
    title: "One playbook, a portfolio of properties.",
    desc: "This isn't a single deal. Family-owned golf courses keep coming to Josh — bought at or below appraisal, each with monetizable land. The same capital-light playbook runs across the portfolio; Bob is negotiating the development rights.",
    page: 10,
  });
  const cols = [
    { w: 3.4, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 2.6, color: C.aubergine, bold: true, font: HEAD, size: 9, align: "center" },
    { w: 6.33, color: C.plum, size: 9, lh: 11 },
  ];
  const rows = [
    ["Renault Winery", "NJ · first deal", "~34 township-offered lots + a +100–200-room hotel expansion"],
    ["Queenstown Harbor", "MD · this proposal", "~280 entitled lots around the 36 holes (~1,000 ac, ~zero basis)"],
    ["The Golf Club at South River", "MD · operating", "Profitable club — appraised ~2× the purchase price"],
    ["Kent Island", "MD · venue", "Clubhouse / event venue — Palantir rented it for a week"],
    ["Colorado course", "Off-market", "Not yet on the website; equity needed to acquire"],
    ["Virginia (1–2 courses)", "Offered to Josh", "Family-owned courses seeking a well-regarded operator"],
  ];
  table(s, ML, 2.1, cols, ["Property", "Status", "The land opportunity"], rows, { rowH: 0.55 });
  callout(s, "A steady train of undervalued courses → monetizable land. Same capital-light playbook each time — and the buys are extraordinary (South River at half appraisal; tax-structured for ~50% bonus depreciation).", 6.35);
  s.addNotes(
    "Bob: 'this isn't a one-off.' Pipeline — Renault (first), Queenstown, South River, Kent Island (Palantir), a Colorado " +
    "course (off-website), 1–2 in Virginia, plus family-owned courses regularly offered to Josh. He's negotiating development " +
    "rights across the properties. Buys are exceptional (South River at half appraisal; 90% to depreciable assets → ~50% bonus depreciation)."
  );
}

// ============================================== SLIDE 11 — CAPITAL: LIGHT
function capitalLight() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "DEVELOPMENT PLAN",
    title: "Entitle the land. Sell entitled lots.",
    desc: "Capital-light: we fund only the soft costs to entitle the land, then sell entitled lots to builders who fund the horizontal. ~$5M of cost, not ~$25M — and the builder carries the construction risk.",
    page: 9,
  });
  statCard(s, ML, 2.05, "~$5M", "Total cost (soft only)",
    "Entitlement, civil, environmental, legal, PM", { w: 6.05 });
  statCard(s, ML, 3.12, "$0", "Horizontal we fund",
    "The builder funds roads, sewer & utilities", { w: 6.05 });

  rect(s, 7.0, 2.05, 5.83, 2.02, C.navy);
  rect(s, 7.0, 2.05, 0.12, 2.02, C.aubergine);
  txt(s, "THE PRINCIPLE", 7.4, 2.2, 5.4, 0.3, { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 4 });
  txt(s, "Stay light; let the builder build.", 7.4, 2.5, 5.4, 0.35, { font: HEAD, size: 13, bold: true, color: C.white });
  txt(s, "We fund only the entitlement and sell entitled lots — the builder funds the horizontal and the homes. Far less capital and no construction risk, on land carried at ~zero basis.",
    7.4, 2.96, 5.4, 1.0, { font: BODY, size: 10.5, color: C.cream, lh: 14, valign: "top" });

  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.3, color: C.navy, size: 9.5 },
    { w: 2.2, color: C.aubergine, bold: true, font: HEAD, size: 10, align: "center" },
    { w: 3.83, color: C.plum, size: 9 },
  ];
  const rows = [
    ["Phase 1 — Entitle", "Approvals for lots", "~$4M soft", "We fund (soft costs only)"],
    ["Phase 2 — Sell lots", "Entitled lots → builders", "Builder", "Builder funds the horizontal"],
    ["Phase 3 — Recycle", "Costs repaid, then 50/50", "Self-funding", "Proceeds → investors + next parcel"],
  ];
  table(s, ML, 4.3, cols, ["Phase", "Approach", "Capital", "Who funds"], rows, { rowH: 0.5 });
  callout(s, "We stay capital-light — fund the entitlement, sell entitled lots, let the builder build. Costs repaid, then a simple 50/50.", 6.5);
  s.addNotes(
    "Capital-light entitled-lot plan (per Bob): fund only the soft costs to entitle, then sell entitled lots; the builder " +
    "funds the horizontal. ~$5M of cost vs ~$25M to build it out — far less capital, no construction risk, much higher return on cost."
  );
}

// ============================================ SLIDE 10 — PRO FORMA & RETURNS
function proForma() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "PRO FORMA  ·  ILLUSTRATIVE",
    title: "The budget, and the 50/50.",
    desc: "Capital-light entitled-lot pro forma: we fund ~$5.3M of soft costs to entitle the land and sell ~280 entitled lots at ~$70K (~$19.6M). The builder funds the horizontal. Costs repaid, then a straight 50/50. Illustrative.",
    page: 10,
  });
  const usesCols = [
    { w: 3.95, color: C.navy, size: 9.5 },
    { w: 1.55, color: C.navy, bold: true, font: HEAD, size: 9.5, align: "right" },
  ];
  const uses = [
    ["Soft costs — entitle the land", "$4.0M"],
    ["Marketing & brokerage (~3%)", "$0.6M"],
    ["Contingency (10% of soft)", "$0.4M"],
    ["Financing / carry", "$0.3M"],
    [{ text: "TOTAL DEVELOPMENT COST", bold: true, font: HEAD, color: C.navy },
      { text: "$5.3M", bold: true, font: HEAD, color: C.aubergine, align: "right" }],
  ];
  table(s, ML, 2.0, usesCols, ["USES OF CAPITAL  ·  SOFT ONLY", ""], uses, { rowH: 0.48 });

  rect(s, ML, 4.95, 5.5, 1.45, C.navy);
  rect(s, ML, 4.95, 0.1, 1.45, C.aubergine);
  txt(s, "KEY ASSUMPTIONS", ML + 0.25, 5.05, 5.1, 0.25, { font: HEAD, size: 8.5, bold: true, color: C.mauve, spc: 2 });
  txt(s, "$70K entitled lot  ·  280 lots  ·  builder funds the horizontal  ·  land at ~zero basis  ·  costs repaid, then a straight 50/50 (no pref, no promote)",
    ML + 0.25, 5.32, 5.05, 1.0, { font: BODY, size: 9, color: C.cream, lh: 12, valign: "top" });

  rect(s, 6.1, 2.0, 3.1, 4.4, C.navy);
  rect(s, 6.1, 2.0, 0.1, 4.4, C.aubergine);
  txt(s, "THE 50/50", 6.32, 2.13, 2.8, 0.3, { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 2 });
  txt(s, "$14.3M", 6.32, 2.5, 2.8, 0.55, { font: HEAD, size: 28, bold: true, color: C.white });
  txt(s, "net profit", 6.32, 3.06, 2.8, 0.25, { font: BODY, size: 9.5, color: C.mauve });
  txt(s, "Costs repaid first, then a straight split:", 6.32, 3.5, 2.66, 0.4, { font: BODY, size: 9.5, color: C.cream, lh: 12, valign: "top" });
  txt(s, "Property / H — 50%", 6.32, 4.0, 2.66, 0.25, { font: HEAD, size: 10.5, bold: true, color: C.white });
  txt(s, "$7.2M  ·  Josh ~$3.6M → returns capital to your investors", 6.32, 4.26, 2.66, 0.6, { font: BODY, size: 9, color: C.cream, lh: 11.5, valign: "top" });
  txt(s, "Development co — 50%", 6.32, 5.1, 2.66, 0.25, { font: HEAD, size: 10.5, bold: true, color: C.white });
  txt(s, "$7.2M  ·  Bob · JAL · partners (JAL earns a share)", 6.32, 5.36, 2.66, 0.6, { font: BODY, size: 9, color: C.cream, lh: 11.5, valign: "top" });

  rect(s, 9.3, 2.0, 3.53, 4.4, C.navy);
  rect(s, 9.3, 2.0, 0.12, 4.4, C.aubergine);
  txt(s, "RETURNS", 9.55, 2.13, 3.2, 0.3, { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 3 });
  const mets = [
    ["$19.6M", "Entitled-lot revenue"],
    ["$14.3M", "Net profit (73% margin)"],
    ["2.7×", "Return on cost (~$5.3M)"],
    ["$7.2M", "Development co (50%)"],
    ["$7.2M", "Property / investors (50%)"],
  ];
  let my = 2.55;
  mets.forEach(([v, l]) => {
    txt(s, v, 9.55, my, 3.1, 0.32, { font: HEAD, size: 17, bold: true, color: C.white });
    txt(s, l, 9.55, my + 0.31, 3.1, 0.22, { font: BODY, size: 8.5, color: C.mauve });
    my += 0.74;
  });

  callout(s, "Illustrative: ~$14.3M profit on ~$5.3M of soft cost (~2.7× on cost). Builder funds the horizontal. Costs repaid, then a straight 50/50 — half returns capital to your investors.", 6.55);
  s.addNotes(
    "Capital-light entitled-lot pro forma, consistent with slide 6. $19.6M revenue − $5.3M soft cost ≈ $14.3M profit " +
    "(~2.7× on cost, 73% margin) on a near-zero land basis; the builder funds the horizontal. Costs repaid first, then a " +
    "straight 50/50 — $7.2M to the property (Josh ~$3.6M, returning capital to his investors) and $7.2M to the development " +
    "company (Bob · JAL · partners). No pref, no promote. Model = source of truth. All illustrative."
  );
}

// ====================================== SLIDE 11 — PRO FORMA BASIS & SOURCES
function basis() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "PRO FORMA  ·  BASIS & SOURCES",
    title: "What the pro forma assumes — and why.",
    desc: "Every input, its value, and the reasoning. Market figures are from public sources; deal-specific figures (acreage, land basis, density) are to be confirmed in diligence.",
    page: 11,
  });
  const cols = [
    { w: 2.9, font: HEAD, bold: true, color: C.navy, size: 9.5 },
    { w: 1.7, color: C.aubergine, bold: true, font: HEAD, size: 9.5, align: "center" },
    { w: 7.73, color: C.plum, size: 9, lh: 11 },
  ];
  const rows = [
    ["Net developable acreage", "~140 ac", "700 gross − 198 conservation − ~300 golf − ~60 Critical Area; survey to confirm"],
    ["Lot density", "~2.0 / ac", "Blended for the developable (non-Critical-Area) land; Queen Anne's County zoning to confirm"],
    ["Entitled lot price", "$70K", "≈ finished lot $150K less the builder's ~$80K to finish (horizontal + margin); confirm vs builder bids"],
    ["Soft costs (entitle)", "~$4.0M", "Entitlement / approvals, civil & environmental, legal / survey, project management"],
    ["Horizontal", "Builder-funded", "Roads, sewer & utilities funded by the builder — not us (capital-light)"],
    ["Land basis", "~zero", "Contributed at a near-zero basis — the land came with the $25M H6 acquisition"],
    ["Structure", "Straight 50/50", "Costs repaid, then 50% property / 50% development co — no pref, no promote"],
  ];
  table(s, ML, 2.1, cols, ["Assumption", "Value", "Basis / reasoning"], rows, { rowH: 0.52 });
  footnote(s, 6.6, [SRC.qacLots, SRC.newcon, SRC.devcost, SRC.cr]);
  s.addNotes(
    "Backs the pro forma. The key assumption is the entitled-lot price (~$70K = finished ~$150K less the builder's ~$80K to " +
    "finish); QAC home/lot prices (Land.com, Zillow) support the finished anchor. Capital-light: the builder funds the horizontal. " +
    "Deal-specific inputs — gross acreage, the developable balance, density — are flagged 'to confirm' pending survey / easement / zoning."
  );
}

// ===================================== SLIDE 12 — THE 50/50 (was waterfall)
function waterfall() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE DEAL  ·  HOW PROFIT SPLITS",
    title: "A straight deal — 50/50, no waterfall.",
    desc: "Bob's structure, kept simple: revenue comes in, costs are repaid, then profit splits 50/50 — half to the property owner (the H entities; Josh owns 50%), half to the development company. No pref, no promote tiers.",
    page: 12,
  });
  const flow = [
    ["$19.6M", "Entitled-lot revenue", "~280 lots × ~$70K"],
    ["($5.3M)", "Less: costs repaid", "the soft costs we funded"],
    ["$14.3M", "Net profit to split", "≈ 2.7× on cost"],
  ];
  const fw = 3.97, fgap = 0.21; let fx = ML;
  flow.forEach(([n, l, sub], i) => {
    const hot = i === 2;
    rect(s, fx, 2.0, fw, 1.25, hot ? C.aubergine : C.white, hot ? {} : { line: { color: C.hair, width: 0.75 } });
    txt(s, n, fx + 0.25, 2.12, fw - 0.5, 0.55, { font: HEAD, size: 26, bold: true, color: hot ? C.white : C.navy });
    txt(s, l, fx + 0.25, 2.72, fw - 0.5, 0.3, { font: HEAD, size: 11, bold: true, color: hot ? C.mauve : C.navy });
    txt(s, sub, fx + 0.25, 2.99, fw - 0.5, 0.25, { font: BODY, size: 9, color: hot ? C.cream : C.plum });
    fx += fw + fgap;
  });

  rect(s, ML, 3.55, 6.06, 2.55, C.navy);
  rect(s, ML, 3.55, 0.12, 2.55, C.aubergine);
  txt(s, "PROPERTY / H ENTITIES", ML + 0.3, 3.7, 5.6, 0.3, { font: HEAD, size: 11, bold: true, color: C.mauve, spc: 2 });
  txt(s, "50%", ML + 0.3, 4.0, 5.6, 0.5, { font: HEAD, size: 22, bold: true, color: C.slate });
  txt(s, "$7.2M", ML + 0.3, 4.5, 5.6, 0.6, { font: HEAD, size: 34, bold: true, color: C.white });
  txt(s, "Josh owns 50% of each property → ~$3.6M (≈ 25% of profit). This is the liquidity that returns capital to his investors against their 8% pref.",
    ML + 0.3, 5.22, 5.6, 0.8, { font: BODY, size: 10, color: C.cream, lh: 13, valign: "top" });

  rect(s, 6.77, 3.55, 6.06, 2.55, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "DEVELOPMENT COMPANY", 7.05, 3.7, 5.6, 0.3, { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "50%", 7.05, 4.0, 5.6, 0.5, { font: HEAD, size: 22, bold: true, color: C.slate });
  txt(s, "$7.2M", 7.05, 4.5, 5.6, 0.6, { font: HEAD, size: 34, bold: true, color: C.navy });
  txt(s, "Accountable Equity / Capital H6 · Bob · JAL + partners. JAL earns a share of this 50% (plus a $15K/mo retainer) — sweat equity, real skin in the game.",
    7.05, 5.22, 5.6, 0.8, { font: BODY, size: 10, color: C.plum, lh: 13, valign: "top" });

  callout(s, "Costs repaid, then a straight 50/50 — half returns capital to your investors, half to the development company. No pref, no promote, no 50-page waterfall.", 6.5);
  s.addNotes(
    "Bob was emphatic: 'there is no waterfall — it's a straight deal.' Revenue in, costs repaid, then 50/50. " +
    "Property / H entities 50% (Josh ~25% of profit → investor liquidity); development company 50% (Accountable Equity / " +
    "Capital H6 · Bob · JAL + partners). JAL takes a share of the dev-co 50% plus a $15K/mo retainer. Simple and aligned."
  );
}

// ===================================== SLIDE 13 — PHASE 1 PROFIT SENSITIVITY
function sensitivity() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "SENSITIVITY  ·  PHASE 1 PROFIT",
    title: "How robust is the profit?",
    desc: "Net project profit ($M) across the entitled-lot price and the lot yield — the two swing factors. The base — $70K lots, ~280 of them — pencils to ~$14.3M; even conservative pricing holds double-digit millions. Split 50/50.",
    page: 13,
  });
  const cols = [
    { w: 2.93, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.13, color: C.navy, size: 12, align: "center" },
    { w: 3.13, color: C.navy, size: 12, align: "center" },
    { w: 3.13, color: C.navy, size: 12, align: "center" },
  ];
  const rows = [
    ["$55K / lot", "$8.1M", "$10.2M", "$12.4M"],
    ["$70K / lot  (base)", "$11.6M", { text: "$14.3M", bold: true, color: C.aubergine, font: HEAD }, "$17.0M"],
    ["$85K / lot", "$15.1M", "$18.4M", "$21.7M"],
  ];
  table(s, ML, 2.3, cols, ["NET PROFIT $M  (price / lots)", "240 lots", "280 lots", "320 lots"], rows, { rowH: 0.85, headSize: 9 });
  callout(s, "Base $70K × 280 → ~$14.3M profit, split 50/50. The downside still clears ~$8M; stronger pricing or yield pushes past $20M.", 5.55);
  txt(s, "Net project profit ($M); each cell re-runs the model. Half flows to the property (investor liquidity), half to the development company. Illustrative.",
    ML, 6.55, CW, 0.28, { font: BODY, size: 7.5, color: C.plum });
  s.addNotes(
    "Sensitivity of Phase 1 net profit to entitled-lot price × lot yield. Base $70K / 280 lots → ~$14.3M, split 50/50 " +
    "($7.2M each). Even at $55K / 240 lots it clears ~$8M; $85K or 320 lots pushes toward $18–22M. Each cell re-runs the model."
  );
}

// ===================================== SLIDE 14 — BUYER UNIVERSE / LIQUIDITY
function builders() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "LOT LIQUIDITY  ·  BUYER UNIVERSE",
    title: "A deep pool of lot buyers.",
    desc: "Regional and national homebuilders are actively building across Queen Anne's County and the Eastern Shore. Entitled, finished lots trade readily — multiple credible takeout buyers let us run a competitive process and de-risk the exit.",
    page: 14,
  });
  const cols = [
    { w: 3.1, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 2.3, color: C.aubergine, bold: true, font: HEAD, size: 9, align: "center" },
    { w: 6.93, color: C.plum, size: 9, lh: 11 },
  ];
  const rows = [
    ["NVR / Ryan Homes", "National · public", "Largest builder in Maryland; routine acquirer of finished / entitled lots"],
    ["D.R. Horton", "National · public", "#1 U.S. homebuilder; active Maryland operations"],
    ["Lennar", "National · public", "Top-3 U.S. builder; Maryland division"],
    ["Cole Homes", "Regional · local", "Named by Bob — active local to the Renault / Atlantic City market"],
    ["Lacrosse Homes", "Regional", "Centreville, Queen Anne's Co. — Meadow Creek, Kent Island Estates"],
    ["Caruso Homes", "Regional", "Eastern Shore communities + build-on-your-lot across Maryland"],
    ["Gemcraft Homes", "Regional", "Mid-Atlantic production builder (MD · DE · VA · PA)"],
    ["Bay to Beach · Baldwin Homes", "Regional · custom", "Eastern Shore custom / semi-custom builders"],
  ];
  table(s, ML, 2.1, cols, ["Builder", "Reach", "Footprint / relevance"], rows, { rowH: 0.42 });
  callout(s, "Multiple well-capitalized buyers = a liquid exit — entitled-lot takedowns or a bulk sale, marketed competitively.", 5.9);
  footnote(s, 6.55, [SRC.lax, SRC.caruso, SRC.gem, SRC.zcom]);
  s.addNotes(
    "Answers the exit-liquidity question: the entitled lots have many credible buyers. Nationals (NVR/Ryan, " +
    "D.R. Horton, Lennar) run land-light and buy finished lots; regionals (Lacrosse, Caruso, Gemcraft) are " +
    "literally building in Queen Anne's Co. / the Shore. We can run a competitive takedown or bulk sale."
  );
}

// ===================================== SLIDE 13 — PHASE 2: HOSPITALITY
function phase2() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "PHASE 2  ·  HOTEL REQUIRED  ·  ILLUSTRATIVE",
    title: "Phase 2 — the hotel (a required component).",
    desc: "The hotel is required (per Bob) — at Renault the plan is +100–200 rooms; at Queenstown a 120-key resort & spa is already designed (FILLAT+). VIVÂMEE operates; JAL leads capital formation. Build-to-core, recapitalizing at stabilization (~Yr 4), clears a mid-to-high-20s return.",
    page: 15,
  });
  const costCols = [
    { w: 3.95, color: C.navy, size: 9.5 },
    { w: 1.55, color: C.navy, bold: true, font: HEAD, size: 9.5, align: "right" },
  ];
  const cost = [
    ["Hotel (100 keys × $325K)", "$32.5M"],
    ["Restaurants (16K SF × $550)", "$8.8M"],
    ["Land — hospitality parcel", "$2.0M"],
    ["Soft / FF&E / pre-opening", "$3.7M"],
    [{ text: "TOTAL DEVELOPMENT COST", bold: true, font: HEAD, color: C.navy },
      { text: "$47.0M", bold: true, font: HEAD, color: C.aubergine, align: "right" }],
  ];
  table(s, ML, 2.0, costCols, ["DEVELOPMENT COST", ""], cost, { rowH: 0.48 });

  const noiCols = [
    { w: 1.95, color: C.navy, size: 9.5 },
    { w: 0.65, color: C.aubergine, bold: true, font: HEAD, size: 9, align: "right" },
    { w: 0.5, color: C.plum, size: 8, align: "right" },
  ];
  const noi = [
    ["Hotel NOI", "$2.8M", "8.6%"],
    ["Restaurant NOI", "$1.0M", "11.5%"],
    [{ text: "STABILIZED NOI", bold: true, font: HEAD, color: C.navy },
      { text: "$3.8M", bold: true, font: HEAD, color: C.aubergine, align: "right" },
      { text: "8.1%", color: C.plum, align: "right" }],
  ];
  table(s, 6.1, 2.0, noiCols, ["STABILIZED NOI", "", "YoC"], noi, { rowH: 0.48 });

  rect(s, 6.1, 3.92, 3.1, 1.84, C.navy);
  rect(s, 6.1, 3.92, 0.1, 1.84, C.aubergine);
  txt(s, "KEY METRICS", 6.32, 4.02, 2.8, 0.25, { font: HEAD, size: 8.5, bold: true, color: C.mauve, spc: 2 });
  txt(s, "Build-to-core: recap / sell ~Yr 4  ·  7.5% exit cap → ~$62M  ·  65% LTC  ·  premium NOI (VIVÂMEE)  ·  ~2.4-pt development spread  ·  land contributed",
    6.32, 4.3, 2.78, 1.4, { font: BODY, size: 8.5, color: C.cream, lh: 11.5, valign: "top" });

  rect(s, 9.3, 2.0, 3.53, 3.76, C.navy);
  rect(s, 9.3, 2.0, 0.12, 3.76, C.aubergine);
  txt(s, "RETURNS", 9.55, 2.13, 3.2, 0.3, { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 3 });
  const mets = [
    ["$47.0M", "Total project cost"],
    ["$4.7M", "Premium NOI (build-to-core)"],
    ["9.9%", "Yield on cost (premium)"],
    ["~2.3×", "Equity multiple (~4-yr)"],
    ["~28%", "Build-to-core IRR"],
  ];
  let my = 2.5;
  mets.forEach(([v, l]) => {
    txt(s, v, 9.55, my, 3.1, 0.32, { font: HEAD, size: 17, bold: true, color: C.white });
    txt(s, l, 9.55, my + 0.31, 3.1, 0.22, { font: BODY, size: 8.5, color: C.mauve });
    my += 0.63;
  });

  rect(s, ML, 5.95, CW, 0.45, C.aubergine);
  txt(s, "Built to core (recap at stabilization), the hotel clears ~28% IRR / ~2.3×. Held for income it's ~12% — the spread is the contributed land + VIVÂMEE's premium NOI.",
    ML + 0.2, 5.95, CW - 0.4, 0.45, { font: BODY, size: 10.5, bold: true, color: C.white, align: "center" });
  footnote(s, 6.55, [SRC.fillat, SRC.hvs, SRC.hcap, SRC.rcost]);
  s.addNotes(
    "Phase 2 = hospitality upside, VIVÂMEE-led; JAL leads capital formation. A 120-key resort & spa is already designed " +
    "(FILLAT+ Architecture), validating the program. Two cases in the model: hold for income (~8% YoC, " +
    "~12% IRR — the floor) or BUILD-TO-CORE — recap/sell at stabilization (~Yr 4) at a 7.5% cap, capturing the development " +
    "spread from the contributed land + premium NOI → ~28% IRR / ~2.3x. Build-to-core lets JAL earn the development return " +
    "while VIVÂMEE keeps and operates the resort. Model is source of truth."
  );
}

// ===================================== SLIDE 16 — PHASE 2 SENSITIVITY
function phase2sens() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "SENSITIVITY  ·  PHASE 2 BUILD-TO-CORE",
    title: "How robust is the hotel return?",
    desc: "Build-to-core IRR (65% LTC, recap at stabilization) across stabilized NOI and exit cap — the two biggest swing factors. Base ($4.65M NOI, 7.5% cap) → ~28%; it holds in the mid-20s+ as long as NOI lands and caps stay sub-8%.",
    page: 16,
  });
  const cols = [
    { w: 2.4, font: HEAD, bold: true, color: C.navy, size: 9.5 },
    { w: 2.475, color: C.navy, size: 12, align: "center" },
    { w: 2.475, color: C.navy, size: 12, align: "center" },
    { w: 2.475, color: C.navy, size: 12, align: "center" },
    { w: 2.475, color: C.navy, size: 12, align: "center" },
  ];
  const rows = [
    ["$4.2M NOI", "24.1%", "19.8%", "15.6%", "11.6%"],
    ["$4.65M NOI  (base)", "31.7%", { text: "27.6%", bold: true, color: C.aubergine, font: HEAD }, "23.7%", "19.9%"],
    ["$5.15M NOI", "38.8%", "34.8%", "31.1%", "27.5%"],
  ];
  table(s, ML, 2.3, cols, ["BUILD-TO-CORE IRR  (NOI / cap)", "7.0% cap", "7.5% cap", "8.0% cap", "8.5% cap"], rows, { rowH: 0.85, headSize: 9 });
  callout(s, "Base ~28%. Strong NOI + a sub-8% cap → low-to-high 30s; soft NOI or an 8.5% cap pulls it toward the high-teens / low-20s.", 5.55);
  txt(s, "Build-to-core project IRR (65% LTC, recap ~Yr 4). VIVÂMEE-led; JAL leads capital formation. Each cell re-runs the model. Illustrative.",
    ML, 6.55, CW, 0.28, { font: BODY, size: 7.5, color: C.plum });
  s.addNotes(
    "Phase 2 build-to-core IRR sensitized to stabilized NOI × exit cap (the development spread). Base $4.65M NOI / 7.5% cap " +
    "→ ~28%. Premium NOI ($5.15M) or a 7% cap pushes it to the mid-30s; an 8.5% cap or soft NOI drops it to ~12–20%. " +
    "65% LTC, recap at stabilization. Model is the source of truth."
  );
}

// ===================================== SLIDE 17 — PHASE 1 CASH FLOW
function phase1cf() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "PHASE 1  ·  CASH FLOW",
    title: "Phase 1 — entitled-lot cash flow.",
    desc: "Capital-light entitled-lot project cash flow ($M). We fund the soft costs to entitle (Yr 0–2); entitled lots sell as approvals land (Yr 2–4). Net profit ~$14.3M, then split 50/50. Illustrative.",
    page: 17,
  });
  const cf = (label, vals, opts = {}) => [
    opts.bold ? { text: label, bold: true, font: HEAD, color: C.navy } : label,
    ...vals.slice(0, 5).map((v) => (opts.bold ? { text: v, bold: true, font: HEAD, color: C.navy, align: "center" } : v)),
    { text: vals[5] || "", bold: true, font: HEAD, color: C.aubergine, align: "right" },
  ];
  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 9.5 },
    { w: 1.5, color: C.navy, size: 10, align: "center" }, { w: 1.5, color: C.navy, size: 10, align: "center" },
    { w: 1.5, color: C.navy, size: 10, align: "center" }, { w: 1.5, color: C.navy, size: 10, align: "center" },
    { w: 1.5, color: C.navy, size: 10, align: "center" }, { w: 1.8, color: C.aubergine, bold: true, font: HEAD, size: 10, align: "right" },
  ];
  const rows = [
    cf("Entitled-lot revenue", ["—", "—", "4.9", "7.8", "6.9", "19.6"]),
    cf("Less: soft costs (entitle)", ["(1.6)", "(2.4)", "(1.0)", "(0.3)", "—", "(5.3)"]),
    cf("Net project cash flow", ["(1.6)", "(2.4)", "3.9", "7.5", "6.9", "14.3"], { bold: true }),
    cf("Cumulative cash flow", ["(1.6)", "(4.0)", "(0.1)", "7.4", "14.3", ""]),
  ];
  table(s, ML, 2.25, cols, ["$M", "Yr 0", "Yr 1", "Yr 2", "Yr 3", "Yr 4", "Total"], rows, { rowH: 0.6, headSize: 9 });
  callout(s, "Net profit ~$14.3M on ~$5.3M of soft cost (~2.7× on cost); the builder funds the horizontal. Costs repaid, then a straight 50/50.", 5.8);
  txt(s, "Project cash flow; illustrative phasing. The builder funds the horizontal. Ties to the model's entitled-lot sheet. Figures subject to confirmation.",
    ML, 6.55, CW, 0.28, { font: BODY, size: 7.5, color: C.plum });
  s.addNotes("Phase 1 entitled-lot project cash flow — ties to the model: $19.6M revenue, $5.3M soft cost, $14.3M profit (~2.7× on cost). We fund only soft costs Yr 0–2; entitled lots sell Yr 2–4; the builder funds the horizontal. Costs repaid, then a straight 50/50.");
}

// ===================================== SLIDE 18 — PHASE 2 CASH FLOW
function phase2cf() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "PHASE 2  ·  BUILD-TO-CORE CASH FLOW",
    title: "Phase 2 — build-to-core equity cash flow.",
    desc: "Hospitality build-to-core, levered equity cash flow ($M). Build to ~$47M cost at 65% LTC; stabilize at the premium NOI; recap / sell at ~Yr 4 (~$62M @ a 7.5% cap). VIVÂMEE operates. Illustrative.",
    page: 18,
  });
  const cf = (label, vals, opts = {}) => [
    opts.bold ? { text: label, bold: true, font: HEAD, color: C.navy } : label,
    ...vals.slice(0, 5).map((v) => (opts.bold ? { text: v, bold: true, font: HEAD, color: C.navy, align: "center" } : v)),
    { text: vals[5] || "", bold: true, font: HEAD, color: C.aubergine, align: "right" },
  ];
  const cols = [
    { w: 3.4, font: HEAD, bold: true, color: C.navy, size: 9.5 },
    { w: 1.42, color: C.navy, size: 10, align: "center" }, { w: 1.42, color: C.navy, size: 10, align: "center" },
    { w: 1.42, color: C.navy, size: 10, align: "center" }, { w: 1.42, color: C.navy, size: 10, align: "center" },
    { w: 1.42, color: C.navy, size: 10, align: "center" }, { w: 1.78, color: C.aubergine, bold: true, font: HEAD, size: 10, align: "right" },
  ];
  const rows = [
    cf("Equity invested", ["(8.2)", "(8.2)", "—", "—", "—", "(16.4)"]),
    cf("Operating cash flow (after debt)", ["—", "—", "1.0", "2.5", "2.5", "6.0"]),
    cf("Recap equity (sale − loan)", ["—", "—", "—", "—", "31.5", "31.5"]),
    cf("Net equity cash flow", ["(8.2)", "(8.2)", "1.0", "2.5", "34.0", "21.1"], { bold: true }),
    cf("Cumulative cash flow", ["(8.2)", "(16.4)", "(15.4)", "(12.9)", "21.1", ""]),
  ];
  table(s, ML, 2.25, cols, ["$M", "Yr 0", "Yr 1", "Yr 2", "Yr 3", "Yr 4", "Total"], rows, { rowH: 0.6, headSize: 9 });
  callout(s, "~$16.4M equity → ~$21M net; build-to-core (65% LTC) levered IRR ~28% / ~2.3× at the Yr-4 recap (~$62M on ~$47M cost). Ties to the model's build-to-core sheet.", 5.8);
  txt(s, "Levered equity cash flow — the line that drives the ~28% IRR; recap / sale at stabilization. A 120-key resort & spa is already designed (FILLAT+). Illustrative.",
    ML, 6.55, CW, 0.28, { font: BODY, size: 7.5, color: C.plum });
  s.addNotes("Phase 2 build-to-core LEVERED EQUITY cash flow — ties cell-for-cell to the model's Phase 2 sheet: equity $16.4M (bcEq, split Yr 0–1), stabilized levered CF ~$2.5M (NOI − debt; Yr 2 ramps at ~40%), exit equity $31.5M (recap $62M − loan), net equity CF [-8.2, -8.2, 1.0, 2.5, 34.0], IRR ~28% / ~2.3×. Project develops to ~$47M cost, recaps ~$62M @ a 7.5% cap.");
}

// ===================================== SLIDE 19 — CAPITAL ACCESS
function capitalAccess() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL ACCESS",
    title: "I bring the capital relationships.",
    desc: "When a parcel does call for equity or debt — or to move faster — here's the access. Active mandates, not a paper Rolodex.",
    page: 19,
  });
  const cards = [
    ["$15M", "Equity being raised now", "Houston retail — family-office capital, closing summer 2026"],
    ["$19.5M", "Debt already secured", "Lender committed on the same transaction"],
    ["Mid–high 20s", "Target development IRR", "Investors expect mid-to-high 20s out-of-state"],
    ["$6.1B", "Career transaction volume", "Across debt & equity — Blackstone + Levcor"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "THE NETWORK",
    "Where I place capital — equity and debt.",
    [
      "Family offices (Texas) — my core equity base; they'll travel for the right risk-adjusted return.",
      "Institutional equity — private-equity fund relationships.",
      "National debt — relationships across national bank platforms and specialty construction lenders.",
      "The bar for out-of-state development is mid-to-high-20s IRR — the fresh land basis helps clear it.",
    ]);
  s.addNotes(
    "Bob's inbound was about capital access. Lead with the live Houston deal ($15M / $19.5M) as proof I close, " +
    "then the network. Even capital-light, this is why I'm useful."
  );
}

// =============================================== SLIDE 10 — THE STRUCTURE
function structure() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE STRUCTURE",
    title: "A vehicle inside the H6 platform.",
    desc: "Accountable Equity's Capital H6 owns the land. Spin the developable parcels into a development company — the property contributes the land, the dev co (Bob · JAL · partners) entitles and sells the lots, and profit splits 50/50.",
    page: 20,
  });
  rect(s, ML, 1.95, CW, 0.86, C.navy);
  rect(s, ML, 1.95, 0.12, 0.86, C.aubergine);
  txt(s, "QUEENSTOWN HARBOR LAND PARTNERS, LLC", ML + 0.35, 2.06, 8.7, 0.4, { font: HEAD, size: 16, bold: true, color: C.white });
  txt(s, "Working name", ML + 0.35, 2.46, 8.5, 0.3, { font: BODY, size: 10, color: C.mauve });
  txt(s, "Development company under Capital H6", 9.4, 1.95, 3.4, 0.86,
    { font: BODY, size: 10, color: C.mauve, align: "right", valign: "middle", margin: [2, 10, 2, 6] });

  const px = [ML, 4.62, 8.74], pw = 3.93;
  rect(s, px[0], 3.0, pw, 2.85, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "LAND = EQUITY", px[0] + 0.25, 3.18, pw - 0.5, 0.3, { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "Owned", px[0] + 0.25, 3.5, pw - 0.5, 0.6, { font: HEAD, size: 30, bold: true, color: C.navy });
  txt(s, "Developable acreage contributed by Capital H6 at appraised value — just acquired, at a low basis.",
    px[0] + 0.25, 4.2, pw - 0.5, 1.5, { font: BODY, size: 10.5, color: C.plum, lh: 14, valign: "top" });

  rect(s, px[1], 3.0, pw, 2.85, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "CAPITAL STACK", px[1] + 0.25, 3.18, pw - 0.5, 0.3, { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "Illustrative", px[1] + 0.25, 3.46, pw - 0.5, 0.25, { font: BODY, size: 9, color: C.plum });
  const stack = [
    ["Land", "Contributed"],
    ["Horizontal", "Builder-funded"],
    ["Soft costs", "Dev co funds"],
  ];
  let sy = 3.78;
  stack.forEach(([a, b]) => {
    txt(s, a, px[1] + 0.25, sy, 2.5, 0.34, { font: BODY, size: 10, color: C.navy });
    txt(s, b, px[1] + 2.6, sy, pw - 2.85, 0.34, { font: HEAD, size: 9.5, bold: true, color: C.aubergine, align: "right" });
    hline(s, px[1] + 0.25, sy + 0.36, pw - 0.5, C.mauve, 0.01);
    sy += 0.55;
  });

  rect(s, px[2], 3.0, pw, 2.85, C.navy);
  rect(s, px[2], 3.0, 0.1, 2.85, C.aubergine);
  txt(s, "ROLES", px[2] + 0.25, 3.18, pw - 0.5, 0.3, { font: HEAD, size: 11, bold: true, color: C.mauve, spc: 2 });
  const roles = [
    ["Josh McCallen / VIVÂMEE", "Resort owner-operator"],
    ["Bob Connell / Accountable Equity", "Development lead · Capital H6 GP"],
    ["JAL Strategies", "Dev-co member — entitlement, capital formation & sales"],
  ];
  let ry = 3.62;
  roles.forEach(([a, b]) => {
    txt(s, a, px[2] + 0.25, ry, pw - 0.5, 0.28, { font: HEAD, size: 10.5, bold: true, color: C.white });
    txt(s, b, px[2] + 0.25, ry + 0.27, pw - 0.5, 0.34, { font: BODY, size: 9.5, color: C.cream, lh: 11 });
    ry += 0.72;
  });

  callout(s, "JAL is a member of the development company — earning a share of its 50% (plus a $15K/mo retainer); intends to invest, sized in diligence.", 6.5);
  s.addNotes(
    "The structure (per Bob): the property contributes the land; the development company (Accountable Equity / Capital H6 · " +
    "Bob · JAL · partners) entitles and sells the lots; costs repaid, then a straight 50/50. JAL runs entitlement + capital " +
    "formation + lot sales for a share of the dev-co 50% plus a $15K/mo retainer, and intends to invest (amount TBD in diligence). VIVÂMEE operates the resort."
  );
}

// ==================================== SLIDE 11 — ENGAGEMENT & COMPENSATION
function engagement() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "ENGAGEMENT  ·  FOR DISCUSSION",
    title: "How I'd want to be engaged.",
    desc: "You asked how I'd like to be comped. My preference matches yours — a small base to fund the work, then real alignment on the value we create. For discussion only.",
    page: 21,
  });
  rect(s, ML, 1.95, CW, 0.95, C.navy);
  rect(s, ML, 1.95, 0.12, 0.95, C.aubergine);
  txt(s, "Hybrid", ML + 0.35, 2.02, 3.0, 0.8, { font: HEAD, size: 30, bold: true, color: C.mauve });
  txt(s, "$15K/mo + expenses, then a share of the dev-co profit", 3.6, 2.06, 5.85, 0.4, { font: HEAD, size: 13.5, bold: true, color: C.white });
  txt(s, "Aligned to value created — and I intend to contribute capital to the dev co myself (amount sized in diligence).", 3.6, 2.44, 5.85, 0.42, { font: BODY, size: 10, color: C.cream, lh: 12, valign: "top" });
  txt(s, "Aligned\nupside", 9.6, 1.95, 3.23, 0.95, { font: HEAD, size: 13, bold: true, color: C.slate, align: "right", valign: "middle", lh: 15, margin: [2, 10, 2, 6] });

  const cols = [
    { w: 3.4, font: HEAD, bold: true, color: C.navy, size: 11 },
    { w: 3.0, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "center" },
    { w: 5.93, color: C.plum, size: 9.5, lh: 12 },
  ];
  const rows = [
    ["Advisory retainer", "$15K / mo", "Funds entitlement, underwriting & capital sourcing during an initial term"],
    ["Expenses", "Reimbursed at cost", "Travel, survey, market & entitlement studies — billed separately, not netted from fees"],
    ["Capital placement fee", "1% debt · 2% equity", "Success fee — 2% of equity raised (1% of debt placed) for a parcel"],
    ["Profit share", "Share of dev-co 50%", "A share of the development company's 50% of profit (with Accountable Equity, Bob & partners)"],
  ];
  table(s, ML, 3.12, cols, ["Component", "Terms", "What it covers"], rows, { rowH: 0.7 });
  callout(s, "The structure flexes to the deal — the point is alignment: I win when you win.", 6.5);
  s.addNotes(
    "Answers Bob's comp question — he pushed for skin in the game over a big retainer. Hybrid: a modest $15K/mo retainer " +
    "(≈ one day/week) + expenses, then the real upside is a share of the development company's 50% of profit. Plus I intend " +
    "to put capital into the dev co (amount TBD). Lead with alignment — the retainer is a floor, the profit share is the prize."
  );
}

// =============================================== SLIDE 12 — WHY JAL (BIO)
function whyJAL() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "WHY JAL AS YOUR PARTNER",
    title: "Pedigree. Relationships. Operator.",
    desc: "Institutional capital-markets pedigree, an active investor network, and hands-on development & asset-management experience.",
    page: 22,
  });
  rect(s, ML, 1.85, 5.0, 4.75, C.navy);
  rect(s, ML, 1.85, 0.12, 4.75, C.aubergine);
  txt(s, "JUSTIN A. LEVINE", ML + 0.3, 2.03, 4.5, 0.3, { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 4 });
  txt(s, "Founder & CEO, JAL Strategies", ML + 0.3, 2.36, 4.5, 0.5, { font: HEAD, size: 19, bold: true, color: C.white });
  const stats = [
    ["$6.1B", "career transaction volume ($5.5B Blackstone + $600M Levcor)"],
    ["$600M", "debt + equity allocated across 26 CRE deals at Levcor"],
    ["3.9M SF", "TX + NC retail managed at Levcor (2014–2025)"],
    ["20+ years", "CRE operating, capital markets & asset management"],
  ];
  let yy = 3.0;
  stats.forEach(([n, d]) => {
    txt(s, n, ML + 0.3, yy, 2.05, 0.55, { font: HEAD, size: 21, bold: true, color: C.slate });
    txt(s, d, ML + 2.35, yy + 0.03, 2.45, 0.72, { font: BODY, size: 9, color: C.cream, lh: 11.5 });
    yy += 0.85;
  });

  rect(s, 5.8, 1.85, 7.03, 4.75, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "CAREER ARC", 6.05, 2.03, 6.6, 0.3, { font: HEAD, size: 10, bold: true, color: C.aubergine, spc: 3 });
  const arc = [
    ["LEVCOR, INC.", "President 2024–2025; CIO 2021–2024; VP 2014–2021. $600M / 3.9M SF TX + NC retail portfolio; realized 4.13x / 16.1% IRR on six dispositions."],
    ["THE BLACKSTONE GROUP", "Analyst / Associate 2006–2011, NYC. Capital raising for U.S., Asia & India funds; $2.5B raised; ~$5.5B aggregate transaction volume."],
    ["JAL-JCP REAL ESTATE", "Co-Founder. Partnership with JCP Investment Management ($250M+ AUM, SEC-registered). Below-replacement-cost unanchored strip centers."],
    ["ULI HOUSTON", "Chair, District Council Management Committee 2022–2024. ULI Americas Mixed-Use & Retail Councils. Member since 2015."],
    ["EDUCATION", "M.B.A., The Wharton School, University of Pennsylvania. B.S. Economics + Communication Studies, Northwestern University."],
  ];
  let ay = 2.4;
  arc.forEach(([org, d], i) => {
    txt(s, org, 6.05, ay, 6.55, 0.28, { font: HEAD, size: 10, bold: true, color: C.navy, spc: 1 });
    txt(s, d, 6.05, ay + 0.28, 6.55, 0.5, { font: BODY, size: 9, color: C.plum, lh: 11, valign: "top" });
    if (i < arc.length - 1) hline(s, 6.05, ay + 0.78, 6.55, C.mauve, 0.01);
    ay += 0.84;
  });
  s.addNotes(
    "Bob's email said 'review your resume.' This is it. Capital-markets pedigree (Blackstone), " +
    "operator/development (Levcor, JAL-JCP), Wharton, ULI leadership."
  );
}

// =============================================== SLIDE 13 — THE PATH
function path() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PATH",
    title: "From this call to a signed mandate.",
    desc: "Light and fast — diligence now, a Zoom shortly, then an in-person at Renault (near Atlantic City) in late July or early August, working around my July 15 surgery; then a defined mandate.",
    page: 23,
  });
  const stages = [
    ["STAGE 1", "DILIGENCE & DATA", "Now – July", "Review & analysis",
      ["Confirm developable vs. conservation acreage (survey + easements)",
       "Pull entitlement & zoning status (Queen Anne's Co. + Renault township)",
       "Model entitled-lot yield, pricing & absorption",
       "Confirm the property / dev-co contribution & 50/50 mechanics"]],
    ["STAGE 2", "IN PERSON", "Late Jul / early Aug", "Walk Renault",
      ["Zoom shortly to align on structure & the first parcels",
       "Meet Bob & Josh; walk Renault (near Atlantic City)",
       "Around my July 15 surgery — I'm in NY through Aug 14",
       "See the other assets in the pipeline"]],
    ["STAGE 3", "MANDATE", "H2 2026", "Engage & execute",
      ["Stand up the development company",
       "Launch entitlements + soft-cost budget (capital-light)",
       "Open builder conversations (incl. Cole)",
       "First entitled-lot sales underwritten"]],
  ];
  const colW = 3.97, gap = 0.21;
  let x = ML;
  stages.forEach(([tag, name, when, sub, bullets]) => {
    rect(s, x, 1.95, colW, 1.06, C.navy);
    rect(s, x, 1.95, 0.1, 1.06, C.aubergine);
    txt(s, tag, x + 0.25, 2.04, colW - 0.5, 0.32, { font: HEAD, size: 12, bold: true, color: C.mauve, spc: 3 });
    txt(s, name, x + 0.25, 2.34, colW - 0.5, 0.32, { font: HEAD, size: 13.5, bold: true, color: C.white });
    txt(s, `${when}  ·  ${sub}`, x + 0.25, 2.68, colW - 0.5, 0.28, { font: BODY, size: 9.5, color: C.slate });
    rect(s, x, 3.05, colW, 2.9, C.white, { line: { color: C.hair, width: 0.75 } });
    let by = 3.22;
    bullets.forEach((b) => {
      txt(s, "—", x + 0.22, by, 0.25, 0.5, { font: HEAD, size: 10, bold: true, color: C.aubergine, valign: "top" });
      txt(s, b, x + 0.5, by, colW - 0.72, 0.62, { font: BODY, size: 10, color: C.navy, lh: 12, valign: "top" });
      by += 0.66;
    });
    x += colW + gap;
  });
  callout(s,
    "What I need: surveys / site plans (you have Renault's; Josh has the rest) · entitlement & zoning status · the 50/50 terms · target timeline. Happy to sign an NDA.",
    6.5);
  s.addNotes(
    "Logistics from the call: Zoom soon, in-person at Renault (near Atlantic City) late July / early August around the " +
    "July 15 foot surgery (non-weight-bearing 4 wks, not the driving foot); I'm in NY through Aug 14; Bob cruises Jun 25–Jul 2 " +
    "and leaves for Utah Aug 8. Bob to send surveys / site plans (has Renault's; Josh has the rest). Restate the data ask + offer the NDA."
  );
}

// =============================================== SLIDE 14 — THANK YOU
function thankYou() {
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  rect(s, 0, 0, 0.25, PAGE_H, C.aubergine);
  txt(s, "JAL STRATEGIES", 0.75, 0.6, 6.0, 0.4, { font: HEAD, size: 13, bold: true, color: C.mauve, spc: 6 });
  txt(s, "THANK", 0.72, 2.45, 12.0, 1.05, { font: HEAD, size: 72, bold: true, color: C.white });
  txt(s, "YOU", 0.72, 3.38, 12.0, 1.05, { font: HEAD, size: 72, bold: true, color: C.mauve });
  hline(s, 0.78, 4.62, 3.5, C.slate, 0.03);
  txt(s, "Bob — appreciate the call, and the résumé read.", 0.75, 4.78, 12.0, 0.45, { font: HEAD, size: 21, bold: true, color: C.white });
  txt(s, "Looking forward to monetizing the land with you, Josh and the H6 team — let's find time at Renault in late July or early August.",
    0.75, 5.3, 12.0, 0.4, { font: BODY, size: 14, color: C.slate });
  txt(s, "Justin A. Levine  |  jlevine@jalstrategies.com  |  JAL Strategies", 0.75, 5.74, 12.0, 0.4, { font: BODY, size: 12, color: C.mauve });
  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, DECK_LABEL, 0.5, 7.0, 9.0, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 2 });
  txt(s, `${TOTAL} / ${TOTAL}`, 9.7, 7.0, 3.13, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 2, align: "right" });
  s.addNotes("Close warm. Reference his email ('review your resume'). Lock the July meeting at Queenstown.");
}

// ----- Build ---------------------------------------------------------------
cover();
opportunity();
theAsset();
theLandBank();
playbook();
mathLand();
whySells();
southRiver();
renault();
pipeline();
capitalLight();
proForma();
basis();
waterfall();
sensitivity();
builders();
phase2();
phase2sens();
phase1cf();
phase2cf();
capitalAccess();
structure();
engagement();
whyJAL();
path();
thankYou();

const OUT_PPTX = "JAL_Queenstown_Harbor_Proposal.pptx";
const OUT_PDF = "JAL_Queenstown_Harbor_Proposal.pdf";

// ----- Renderer: PPTX (replay recorded calls to pptxgenjs) -----------------
function renderPptx(deck) {
  const p = new pptxgen();
  p.defineLayout({ name: "W16", width: PAGE_W, height: PAGE_H });
  p.layout = "W16";
  p.author = deck.props.author;
  p.company = deck.props.company;
  p.title = deck.props.title;
  deck.slides.forEach((S) => {
    const s = p.addSlide();
    if (S.background) s.background = S.background;
    S.calls.forEach(([kind, a, b]) => {
      if (kind === "shape") s.addShape(a, b);
      else if (kind === "text") s.addText(a, b);
    });
    if (S.notes) s.addNotes(S.notes);
  });
  return p.writeFile({ fileName: OUT_PPTX });
}

// ----- Renderer: PDF (replay the same calls to pdfkit) ---------------------
const PT = 72; // inches -> points
function loadFonts(doc) {
  const dir = nodePath.join(__dirname, "fonts");
  const reg = (key, file, builtin) => {
    const fp = nodePath.join(dir, file);
    if (fs.existsSync(fp)) { doc.registerFont(key, fp); return key; }
    return builtin; // fall back to a PDF base font if the TTF is missing
  };
  return {
    MontReg: reg("MontReg", "Montserrat-Regular.ttf", "Helvetica"),
    MontBold: reg("MontBold", "Montserrat-Bold.ttf", "Helvetica-Bold"),
    DMReg: reg("DMReg", "DMSans-Regular.ttf", "Helvetica"),
    DMBold: reg("DMBold", "DMSans-Bold.ttf", "Helvetica-Bold"),
  };
}
const pickFont = (F, face, bold) =>
  face === HEAD ? (bold ? F.MontBold : F.MontReg) : (bold ? F.DMBold : F.DMReg);
const normMargin = (m) =>
  m == null ? [0, 0, 0, 0] : typeof m === "number" ? [m, m, m, m] : m; // [t,r,b,l]

function drawRect(doc, o) {
  const x = o.x * PT, y = o.y * PT, w = o.w * PT, h = o.h * PT;
  if (o.fill && o.fill.color) doc.rect(x, y, w, h).fill("#" + o.fill.color);
  if (o.line && o.line.color && o.line.type !== "none")
    doc.rect(x, y, w, h).lineWidth(o.line.width || 1).stroke("#" + o.line.color);
}
function drawText(doc, F, text, o) {
  const m = normMargin(o.margin);
  const x = o.x * PT + m[3];
  const w = o.w * PT - m[1] - m[3];
  const yb = o.y * PT + m[0];
  const hb = o.h * PT - m[0] - m[2];
  const size = o.fontSize || 11;
  if (Array.isArray(text)) return drawRich(doc, F, text, o, x, w, yb, hb, size);
  doc.font(pickFont(F, o.fontFace, o.bold)).fontSize(size).fillColor("#" + (o.color || "000000"));
  const opts = { width: w, align: o.align || "left", characterSpacing: o.charSpacing || 0, lineBreak: o.wrap !== false };
  if (o.lineSpacing) opts.lineGap = Math.max(0, o.lineSpacing - size);
  if (o.underline) opts.underline = true;
  if (o.hyperlink && o.hyperlink.url) opts.link = o.hyperlink.url;
  let ty = yb;
  const va = o.valign || "middle";
  if (va !== "top") {
    const th = doc.heightOfString(text, { width: w, characterSpacing: opts.characterSpacing, lineGap: opts.lineGap || 0 });
    const slack = Math.max(0, hb - th);
    ty = yb + (va === "bottom" ? slack : slack / 2);
  }
  doc.text(text, x, ty, opts);
}
function drawRich(doc, F, runs, o, x, w, yb, hb, size) {
  const ty = yb + Math.max(0, (hb - size) / 2);
  runs.forEach((run, i) => {
    const ro = run.options || {};
    doc.font(pickFont(F, o.fontFace, ro.bold)).fontSize(size).fillColor("#" + (ro.color || o.color || "000000"));
    const opts = { continued: i < runs.length - 1, characterSpacing: o.charSpacing || 0, underline: !!ro.underline };
    if (ro.hyperlink && ro.hyperlink.url) opts.link = ro.hyperlink.url;
    if (i === 0) doc.text(run.text, x, ty, opts);
    else doc.text(run.text, opts);
  });
}
function renderPdf(deck) {
  return new Promise((resolve, reject) => {
    const sz = [PAGE_W * PT, PAGE_H * PT];
    const doc = new PDFDocument({ size: sz, margin: 0, autoFirstPage: false,
      info: { Title: deck.props.title, Author: deck.props.author } });
    const stream = fs.createWriteStream(OUT_PDF);
    doc.pipe(stream);
    const F = loadFonts(doc);
    deck.slides.forEach((S) => {
      doc.addPage({ size: sz, margin: 0 });
      doc.rect(0, 0, sz[0], sz[1]).fill("#" + ((S.background && S.background.color) || "FFFFFF"));
      S.calls.forEach(([kind, a, b]) => {
        if (kind === "shape" && a === "rect") drawRect(doc, b);
        else if (kind === "text") drawText(doc, F, a, b);
      });
    });
    doc.end();
    stream.on("finish", resolve);
    stream.on("error", reject);
  });
}

Promise.all([renderPptx(pptx), renderPdf(pptx)])
  .then(() => console.log("Wrote", OUT_PPTX, "and", OUT_PDF))
  .catch((e) => { console.error(e); process.exit(1); });
