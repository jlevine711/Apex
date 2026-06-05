/**
 * JAL Strategies — Queenstown Harbor Land Development Partnership
 * Prepared for Robert A. Connell, CFP (Apex Financial Advisors / Accountable Equity · Capital H6)
 *
 * Scope: monetize the developable LAND around Queenstown Harbor — the 36-hole
 * Eastern Shore resort that Accountable Equity / VIVÂMEE acquired via Capital H6
 * (May 2026, with The Golf Club at South River, $25M, from The Brick Companies).
 * Entitle the developable, non-conservation acreage and sell finished lots to
 * homebuilders (builders fund horizontal infrastructure); recycle proceeds.
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
const TOTAL = 15;

const DECK_LABEL = "QUEENSTOWN HARBOR  |  CONFIDENTIAL";
const FOOTER_LEFT =
  "JAL Strategies  |  Justin A. Levine, Founder & CEO  |  jlevine@jalstrategies.com";

const pptx = new pptxgen();
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

function chrome(s, { eyebrow, title, desc, page }) {
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
};

// Small hyperlinked source footnote, just above the footer rule.
function footnote(s, y, items) {
  const runs = [{ text: "Sources:  ", options: { color: C.plum, bold: true } }];
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

  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, "PREPARED FOR ROBERT A. CONNELL, CFP   |   ACCOUNTABLE EQUITY · CAPITAL H6",
    0.5, 7.0, 9.6, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5 });
  txt(s, "JUNE 2026  |  CONFIDENTIAL", 9.7, 7.0, 3.13, 0.4, {
    font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5, align: "right",
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
    title: "You just bought it. Now monetize the land.",
    desc: "Bob — Capital H6 just closed Queenstown Harbor. The resort runs itself; the upside is the developable land around the 36 holes — entitle it, sell lots to builders, recycle the proceeds.",
    page: 2,
  });
  const cards = [
    ["$25M", "Just acquired (May 2026)", "Queenstown + South River, via Capital H6"],
    ["36 holes", "River & Lakes courses", "Waterfront resort on the Eastern Shore"],
    ["870+ ac", "Across the two properties", "Waterfront + conservation land (to verify)"],
    ["Soft costs", "To begin", "Builders fund the horizontal infrastructure"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "WHY NOW",
    "You own it free and clear — Capital H6 just closed the purchase.",
    [
      "Fresh, low basis: $25M for two courses across 870+ acres — the land value is the upside.",
      "Builders fund sewer, roads and utilities — you carry approvals, not construction risk.",
      "Minutes over the Bay Bridge to Annapolis; ~1 hour to D.C. & Baltimore.",
      "Golf and water frontage on the Chesapeake commands a premium — and sells quickly.",
    ]);
  footnote(s, 6.66, [SRC.cr, SRC.eoa]);
  s.addNotes(
    "Anchor on the fresh acquisition (H6 — Bob's fund). The land bank is additive to the hospitality: " +
    "monetize the developable acreage the golf doesn't need, capital-light."
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
    desc: "The resort operates; the value-creation is entitling the developable, non-conservation acreage and selling finished lots to homebuilders. Capital-light, and additive to the hospitality.",
    page: 4,
  });
  const cards = [
    ["870+ ac", "Portfolio land", "Across Queenstown + South River"],
    ["198 ac", "Under conservation", "Permanent easement at Queenstown — plan around it"],
    ["Net", "Developable land", "Balance after conservation, wetlands & course (to verify)"],
    ["Lots", "Entitle & sell", "Finished lots to regional homebuilders"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "THE THESIS",
    "Monetize the land the golf doesn't need.",
    [
      "Identify developable parcels at the course edges — away from conservation, wetlands and play.",
      "Entitle for residential; sell finished lots to builders who fund the horizontal work.",
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
    desc: "Capital-light by sequence — soft costs first, then builders carry the infrastructure and construction risk, and the proceeds recycle into the next parcel.",
    page: 5,
  });
  const items = [
    ["01", "Entitle", "Approvals for lots",
      "Secure zoning and subdivision approvals for the developable acreage around the fairways."],
    ["02", "Sell to Builders", "Finished lots → homebuilders",
      "Sell entitled lots to regional builders and let them fund sewer, roads and utilities."],
    ["03", "Golf & Water Premium", "Frontage sells higher",
      "Lots on a course or with Chesapeake water views command a premium and absorb faster."],
    ["04", "Recycle & Repeat", "Parcel by parcel",
      "Roll proceeds into the next parcel — and seed the broader Capital H6 program."],
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
    title: "From 700 acres to the developable envelope.",
    desc: "Built down from real constraints — the recorded conservation easement, the golf footprint, and Maryland's Chesapeake Bay Critical Area. Developable acreage is survey-dependent; figures illustrative.",
    page: 6,
  });
  const cols = [
    { w: 3.2, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.4, color: C.navy, size: 9.5 },
    { w: 2.9, color: C.plum, size: 9, font: BODY },
    { w: 2.83, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "right" },
  ];
  const rows = [
    ["Queenstown land (gross)", "per ownership estimate", "—", "~700 ac"],
    ["Less permanent conservation", "recorded easement", "700 − 198", "502 ac"],
    ["Less golf, range, water & lodging", "two 18s + 9-ac range + cottages", "502 − ~300", "~200 ac"],
    ["Less Critical Area & wetlands", "Chesapeake 1,000-ft zone", "200 − ~60", "~140 ac"],
    ["Net developable → lots", "~2.0 lots / acre", "140 × 2.0", "~280 lots"],
    [{ text: "FINISHED LOT VALUE (net, illustrative)", bold: true, font: HEAD, color: C.navy },
      { text: "$150K / lot, less ~30% costs", color: C.plum }, "280 × $150K × 0.70", "≈ $29M"],
  ];
  table(s, ML, 2.12, cols, ["Step", "Basis", "Calculation", "= Result"], rows, { rowH: 0.55 });
  rect(s, ML, 5.9, CW, 0.5, C.navy);
  txt(s,
    "≈ $29M illustrative net from ~140 developable acres — net of conservation & Critical Area, additive to the resort.",
    ML + 0.2, 5.9, CW - 0.4, 0.5,
    { font: HEAD, size: 12, bold: true, color: C.white, align: "center" });
  footnote(s, 6.62, [SRC.qhwho, SRC.qac, SRC.dnr]);
  s.addNotes(
    "Now derived, not guessed: 700 gross − 198 conservation − ~300 golf complex − ~60 Critical Area/wetlands " +
    "≈ 140 developable acres. Maryland's Chesapeake Bay Critical Area (1,000-ft zone, ~1 unit/20 ac in RCA) is the " +
    "real constraint on the waterfront land. All figures illustrative until survey / easement review."
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

// ============================================== SLIDE 9 — CAPITAL: LIGHT
function capitalLight() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL STRATEGY",
    title: "Capital-light by design.",
    desc: "The land bank barely needs capital: soft costs to entitle, then builders fund the horizontal work and buy the lots. Cash in is small; proceeds come fast.",
    page: 9,
  });
  statCard(s, ML, 2.05, "$1–3M", "Soft costs to start",
    "Entitlements, planning & approvals", { w: 6.05 });
  statCard(s, ML, 3.12, "Land", "= the equity",
    "Acreage held by Capital H6 — already acquired", { w: 6.05 });

  rect(s, 7.0, 2.05, 5.83, 2.02, C.navy);
  rect(s, 7.0, 2.05, 0.12, 2.02, C.aubergine);
  txt(s, "THE PRINCIPLE", 7.4, 2.2, 5.4, 0.3, { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 4 });
  txt(s, "You carry approvals, not construction.", 7.4, 2.5, 5.4, 0.35, { font: HEAD, size: 13, bold: true, color: C.white });
  txt(s, "Builders take the infrastructure burden — sewer, roads, utilities — and pay for finished lots. Your cash at risk stays small and staged.",
    7.4, 2.96, 5.4, 1.0, { font: BODY, size: 10.5, color: C.cream, lh: 14, valign: "top" });

  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.3, color: C.navy, size: 9.5 },
    { w: 2.2, color: C.aubergine, bold: true, font: HEAD, size: 10, align: "center" },
    { w: 3.83, color: C.plum, size: 9 },
  ];
  const rows = [
    ["Phase 1 — Entitle", "Approvals for lots", "$1–3M", "JAL + H6 dev co (skin in the game)"],
    ["Phase 2 — Builder takedown", "Horizontal + lot purchase", "~$0 to you", "Regional homebuilder funds & buys"],
    ["Phase 3 — Recycle", "Proceeds redeployed", "Self-funding", "Next parcel / South River / H6 program"],
  ];
  table(s, ML, 4.3, cols, ["Phase", "Approach", "Your capital", "Who funds"], rows, { rowH: 0.5 });
  callout(s, "Minimal capital at risk — and the land bank seeds everything that comes after.", 6.5);
  s.addNotes(
    "For the land bank the honest story is capital-light: soft costs in, builders fund infra and buy lots. " +
    "Land is already owned via H6 — no acquisition capital needed."
  );
}

// =============================================== SLIDE 9 — CAPITAL ACCESS
function capitalAccess() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL ACCESS",
    title: "I bring the capital relationships.",
    desc: "When a parcel does call for equity or debt — or to move faster — here's the access. Active mandates, not a paper Rolodex.",
    page: 10,
  });
  const cards = [
    ["$14M", "Equity being raised now", "Houston retail — family-office capital, closing summer 2026"],
    ["$19.5M", "Debt already secured", "Lender committed on the same transaction"],
    ["Mid-20s%", "Target development IRR", "Investors expect mid-to-high 20s out-of-state"],
    ["$6.1B", "Career transaction volume", "Across debt & equity — Blackstone + Levcor"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "THE NETWORK",
    "Where I place capital — equity and debt.",
    [
      "Family offices (Texas) — my core equity base; they'll travel for the right risk-adjusted return.",
      "Institutional equity — Angelo Gordon and peer private-equity relationships.",
      "National debt — Goldman Sachs (met two weeks ago; national platform, Dallas team) + specialty lenders.",
      "The bar for out-of-state development is mid-to-high-20s IRR — the fresh land basis helps clear it.",
    ]);
  s.addNotes(
    "Bob's inbound was about capital access. Lead with the live Houston deal ($14M / $19.5M) as proof I close, " +
    "then the network. Even capital-light, this is why I'm useful."
  );
}

// =============================================== SLIDE 10 — THE STRUCTURE
function structure() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE STRUCTURE",
    title: "A vehicle inside the H6 platform.",
    desc: "Capital H6 owns the land. Spin the developable parcels into a land-development entity — H6 contributes the acreage as equity; JAL co-invests and runs entitlement, capital and lot sales alongside your team.",
    page: 11,
  });
  rect(s, ML, 1.95, CW, 0.86, C.navy);
  rect(s, ML, 1.95, 0.12, 0.86, C.aubergine);
  txt(s, "QUEENSTOWN HARBOR LAND PARTNERS, LLC", ML + 0.35, 2.06, 8.7, 0.4, { font: HEAD, size: 16, bold: true, color: C.white });
  txt(s, "Working name", ML + 0.35, 2.46, 8.5, 0.3, { font: BODY, size: 10, color: C.mauve });
  txt(s, "Land-development vehicle under Capital H6", 9.4, 1.95, 3.4, 0.86,
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
    ["Land equity (H6)", "Contributed"],
    ["Builder-funded infra", "Horizontal"],
    ["Cash equity (JAL + LP)", "Soft costs"],
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
    ["Bob Connell / Capital H6", "Development lead (H6 GP)"],
    ["JAL Strategies", "Capital, entitlement & sales — co-invests"],
  ];
  let ry = 3.62;
  roles.forEach(([a, b]) => {
    txt(s, a, px[2] + 0.25, ry, pw - 0.5, 0.28, { font: HEAD, size: 10.5, bold: true, color: C.white });
    txt(s, b, px[2] + 0.25, ry + 0.27, pw - 0.5, 0.34, { font: BODY, size: 9.5, color: C.cream, lh: 11 });
    ry += 0.72;
  });

  callout(s, "JAL takes skin in the game — co-investing alongside, not just advising.", 6.5);
  s.addNotes(
    "Nails Bob's exact role: H6 GP / development lead. H6 contributes the developable land; JAL co-invests " +
    "and runs entitlement + capital + lot sales. VIVÂMEE keeps operating the resort."
  );
}

// ==================================== SLIDE 11 — ENGAGEMENT & COMPENSATION
function engagement() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "ENGAGEMENT  ·  FOR DISCUSSION",
    title: "How I'd want to be engaged.",
    desc: "You asked how I'd like to be comped. My preference matches yours — a small base to fund the work, then real alignment on the value we create. For discussion only.",
    page: 12,
  });
  rect(s, ML, 1.95, CW, 0.95, C.navy);
  rect(s, ML, 1.95, 0.12, 0.95, C.aubergine);
  txt(s, "Hybrid", ML + 0.35, 2.02, 3.0, 0.8, { font: HEAD, size: 30, bold: true, color: C.mauve });
  txt(s, "$15K/mo + expenses, then success fees + carry", 3.6, 2.06, 5.85, 0.4, { font: HEAD, size: 13.5, bold: true, color: C.white });
  txt(s, "Aligned to lot-sale value created — not a flat consulting check.", 3.6, 2.46, 5.85, 0.35, { font: BODY, size: 10, color: C.cream });
  txt(s, "Skin in\nthe game", 9.6, 1.95, 3.23, 0.95, { font: HEAD, size: 13, bold: true, color: C.slate, align: "right", valign: "middle", lh: 15, margin: [2, 10, 2, 6] });

  const cols = [
    { w: 3.4, font: HEAD, bold: true, color: C.navy, size: 11 },
    { w: 3.0, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "center" },
    { w: 5.93, color: C.plum, size: 9.5, lh: 12 },
  ];
  const rows = [
    ["Advisory retainer", "$15K / mo", "Funds entitlement, underwriting & capital sourcing during an initial term; creditable against success fees"],
    ["Expenses", "Reimbursed at cost", "Travel, survey, market & entitlement studies — billed separately, not netted from fees"],
    ["Capital placement fee", "1% debt · 2–3% equity", "Success fee if and when we place debt or equity for a parcel"],
    ["Carried interest", "Co-GP promote", "Share of value created on lot sales; JAL co-invests its own capital alongside"],
  ];
  table(s, ML, 3.12, cols, ["Component", "Terms", "What it covers"], rows, { rowH: 0.7 });
  callout(s, "The structure flexes to the deal — the point is alignment: I win when you win.", 6.5);
  s.addNotes(
    "Directly answers Bob's comp question. Hybrid: $15K/mo creditable retainer (≈ one day/week) + expenses " +
    "reimbursed, then placement fees and the promote. Lead with alignment — the retainer is a floor, the carry is the prize."
  );
}

// =============================================== SLIDE 12 — WHY JAL (BIO)
function whyJAL() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "WHY JAL AS YOUR PARTNER",
    title: "Pedigree. Relationships. Operator.",
    desc: "Institutional capital-markets pedigree, an active investor network, and hands-on development & asset-management experience.",
    page: 13,
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
    desc: "Light and fast to start — diligence now, an in-person while I'm in the Northeast in July, then a defined engagement on the first parcels.",
    page: 14,
  });
  const stages = [
    ["STAGE 1", "DILIGENCE & DATA", "Now – July", "Review & analysis",
      ["Confirm developable vs. conservation acreage (survey + easements)",
       "Pull entitlement & zoning status (Queen Anne's County)",
       "Model lot yield & absorption for the developable parcels",
       "Confirm the H6 ownership & contribution mechanics"]],
    ["STAGE 2", "IN PERSON", "July 2026", "Walk the site",
      ["Meet while I'm in the NY / Philly area — train down or drive to the Shore",
       "Walk Queenstown (River House & the course edges)",
       "Meet Josh / VIVÂMEE",
       "Align on first parcels, capital plan & comp"]],
    ["STAGE 3", "MANDATE", "H2 2026", "Engage & execute",
      ["Stand up the land vehicle under Capital H6",
       "Launch Phase 1 entitlements + soft-cost budget",
       "Open regional-homebuilder conversations",
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
    "What I need to sharpen the model: survey / plat & conservation easements · entitlement & zoning status · H6 structure · target timeline.",
    6.5);
  s.addNotes(
    "Close on logistics. The July in-person is real (wife's family on Long Island) — Queenstown is a doable " +
    "drive from Philly. The callout is the data ask — restate it so Bob knows exactly what to send."
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
  txt(s, "Looking forward to monetizing the land around Queenstown Harbor with you, Josh and the H6 team — let's find time in July.",
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
capitalLight();
capitalAccess();
structure();
engagement();
whyJAL();
path();
thankYou();

const OUT = "JAL_Queenstown_Harbor_Proposal.pptx";
pptx.writeFile({ fileName: OUT }).then((f) => console.log("Wrote", f));
