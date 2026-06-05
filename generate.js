/**
 * JAL Strategies — Development & Capital Partnership Proposal
 * Prepared for Robert A. Connell, CFP (Apex Financial Advisors / Accountable Equity)
 *
 * Design system reverse-engineered from the JAL "Republic Square" pitch deck:
 *   - 16:9 (13.33" x 7.5")
 *   - Montserrat (headings) / DM Sans (body)
 *   - Palette: navy #0B163C, aubergine #3A243A, plum #6B5A6B,
 *     mauve #C4B8C4, cream #F4F2ED, slate #8C9BB5, white
 *
 * Content is grounded in the 6/5/2026 call between Justin Levine and Bob Connell
 * (Plaud transcript). Financial figures on "THE MATH" / capital slides are
 * clearly labelled ILLUSTRATIVE — to be recalibrated against Accountable Equity data.
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
  hair: "E6E1DB", // faint card hairline on cream
};
const HEAD = "Montserrat";
const BODY = "DM Sans";

const PAGE_W = 13.333;
const PAGE_H = 7.5;
const ML = 0.5;          // left margin
const CW = 12.33;        // content width
const TOTAL = 14;        // total slides (for "n / 14")

const DECK_LABEL = "ACCOUNTABLE EQUITY  |  CONFIDENTIAL";
const FOOTER_LEFT =
  "JAL Strategies  |  Justin A. Levine, Founder & CEO  |  jlevine@jalstrategies.com";

const pptx = new pptxgen();
pptx.defineLayout({ name: "W16", width: PAGE_W, height: PAGE_H });
pptx.layout = "W16";
pptx.author = "JAL Strategies";
pptx.company = "JAL Strategies";
pptx.title =
  "Accountable Equity — Land Development & Capital Partnership (for Robert A. Connell)";

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

// Content-slide chrome: cream bg, header band, eyebrow/title/desc, footer.
function chrome(s, { eyebrow, title, desc, page }) {
  s.background = { color: C.cream };
  rect(s, 0, 0, PAGE_W, 0.04, C.navy);
  txt(s, "JAL STRATEGIES", ML, 0.12, 3.0, 0.3, {
    font: HEAD, size: 9, bold: true, color: C.navy, spc: 3,
  });
  txt(s, DECK_LABEL, 6.0, 0.12, 6.83, 0.3, {
    font: HEAD, size: 9, color: C.plum, spc: 2, align: "right",
  });
  // footer
  hline(s, ML, 7.05, 12.33, C.mauve);
  txt(s, FOOTER_LEFT, ML, 7.15, 9.0, 0.3, { font: BODY, size: 8, color: C.plum });
  txt(s, `${page} / ${TOTAL}`, 12.0, 7.15, 0.83, 0.3, {
    font: BODY, size: 8, color: C.plum, align: "right",
  });
  // eyebrow / title / desc
  txt(s, eyebrow, ML, 0.5, 9.0, 0.3, {
    font: HEAD, size: 10, bold: true, color: C.aubergine, spc: 6,
  });
  txt(s, title, ML, 0.84, CW, 0.6, { font: HEAD, size: 25, bold: true, color: C.navy });
  if (desc)
    txt(s, desc, ML, 1.46, CW, 0.5, { font: BODY, size: 13, color: C.plum, lh: 16 });
}

// Bottom emphasis band (aubergine) with centered white text.
function callout(s, text, y = 6.5) {
  rect(s, ML, y, CW, 0.46, C.aubergine);
  txt(s, text, ML + 0.2, y, CW - 0.4, 0.46, {
    font: BODY, size: 11, bold: true, color: C.white, align: "center",
  });
}

// White stat card: big number + label + sublabel.
function statCard(s, x, y, num, label, sub, o = {}) {
  const w = o.w || 6.0;
  // auto-shrink long values so they never spill into the label column
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

// Navy side panel: mauve heading, white lead line, cream supporting items.
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

// Shape-grid table. colDefs: [{w, align, color, font, bold, size}]. rows: array of
// arrays; a cell is a string or {text, color, bold, font, size, align}.
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
  return ry; // bottom y
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

  txt(s, "ACCOUNTABLE", 0.72, 2.18, 12.2, 1.05, {
    font: HEAD, size: 67, bold: true, color: C.white, spc: 1,
  });
  txt(s, "EQUITY", 0.72, 3.12, 12.2, 1.05, {
    font: HEAD, size: 67, bold: true, color: C.mauve, spc: 1,
  });

  hline(s, 0.78, 4.42, 3.5, C.slate, 0.03);
  txt(s, "Land Development & Capital Partnership", 0.75, 4.58, 12.0, 0.45, {
    font: HEAD, size: 22, bold: true, color: C.white,
  });
  txt(s, "A proposal for JAL Strategies as development & capital partner",
    0.75, 5.12, 12.0, 0.35, { font: BODY, size: 14, color: C.slate });
  txt(s,
    "Queenstown Harbor  ·  Kent Island  ·  LBI National  ·  Bohemia Manor  ·  1,000+ acres",
    0.75, 5.52, 12.0, 0.35, { font: BODY, size: 12, color: C.mauve });

  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, "PREPARED FOR ROBERT A. CONNELL, CFP   |   APEX FINANCIAL ADVISORS · ACCOUNTABLE EQUITY",
    0.5, 7.0, 9.6, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5 });
  txt(s, "JUNE 2026  |  CONFIDENTIAL", 9.7, 7.0, 3.13, 0.4, {
    font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5, align: "right",
  });
  s.addNotes(
    "Cover. This is the discussion document Bob asked me to put together on our 6/5 call. " +
    "Frame: JAL as development + capital partner for Accountable Equity's land bank."
  );
}

// ===================================================== SLIDE 2 — OPPORTUNITY
function opportunity() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE OPPORTUNITY",
    title: "A 1,000-acre land bank — owned, and waiting.",
    desc: "Bob — here's how I'd think about turning Accountable Equity's land into cash, and exactly where I can help.",
    page: 2,
  });
  const cards = [
    ["1,000+", "Acres of developable land", "Across MD, NJ & VA — owned, not optioned"],
    ["700", "Acres at Queenstown Harbor", "Flagship golf resort near Annapolis & D.C."],
    ["~1.5M SF", "Stabilized commercial base", "Lockheed (Arlington TX), Las Vegas, Philadelphia"],
    ["Up to $100M", "Capital program contemplated", "Debt, equity & development partners"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "WHY THIS MOVES NOW",
    'The hotels are at capacity — turning business away and "not making a dime" on it.',
    [
      "Land is owned free and clear — the hardest, most expensive input is already on the balance sheet.",
      "Town relationships are real — one municipality has offered to contribute land to build on.",
      "Josh McCallan turns properties into destinations — a proven operator de-risks the demand side.",
      "What's missing is capital + bandwidth — precisely the gap JAL is built to fill.",
    ]);
  s.addNotes(
    "Open by mirroring Bob's framing back to him: lots of owned land, strong hospitality operator, " +
    "real town relationships, hotels turning away business. The constraint is capital + bandwidth."
  );
}

// ======================================================= SLIDE 3 — PORTFOLIO
function portfolio() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PORTFOLIO",
    title: "Five hospitality assets, one commercial base.",
    desc: "Land grouped by what it can become. Hospitality parcels drive the development upside; the commercial portfolio anchors cash flow and collateral.",
    page: 3,
  });
  const cols = [
    { w: 2.25, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 1.85, color: C.plum, size: 9.5 },
    { w: 3.55, color: C.navy, size: 9.5 },
    { w: 4.68, color: C.navy, size: 9.5, lh: 11 },
  ];
  const rows = [
    ["Queenstown Harbor", "Queenstown, MD", "700 ac · 3 courses · near Annapolis & D.C.",
      "Entitle & sell residential lots around the courses; add hotel keys for overflow"],
    ["Kent Island Resort", "Kent Island, MD", "Chesapeake Bay waterfront resort",
      "Expand keys + destination amenities; capture turned-away demand"],
    ["LBI National Golf", "Long Beach Island, NJ", "Coastal golf in a high-barrier shore market",
      "Entitled lot sales + boutique coastal hospitality"],
    ["Bohemia Manor", "Cecil County, MD", "Historic estate & event grounds",
      "Hospitality + events; winery-style destination template"],
    ["Winery HQ", "New Jersey", 'Transformed "dump" → destination venue',
      "Proof of concept — the destination playbook other sites replicate"],
    ["Commercial base", "TX · NV · PA", "~1.5M SF incl. Lockheed Martin (Arlington)",
      "Stabilized cash flow + collateral base for development debt"],
  ];
  table(s, ML, 2.05, cols, ["Asset", "Location", "Profile", "Development thesis"], rows,
    { rowH: 0.6 });
  callout(s,
    "We start where capital-light wins are fastest — entitle, sell to a homebuilder, recycle proceeds into the next parcel.",
    6.5);
  s.addNotes(
    "Walk the portfolio. Emphasize that these are owned assets with distinct paths. " +
    "Queenstown is the flagship; commercial base (Lockheed etc.) underwrites credit."
  );
}

// ======================================================== SLIDE 4 — PLAYBOOK
function playbook() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PLAYBOOK",
    title: "Four levers that turn land into cash flow.",
    desc: "Sequenced to be capital-light first — generate proceeds early, then recycle into the heavier hospitality build-out.",
    page: 4,
  });
  const items = [
    ["01", "Entitle & Sell", "Lots → homebuilders",
      "Secure approvals, sell finished lots, and let the builder fund the horizontal infrastructure — sewer, roads, utilities."],
    ["02", "Expand Keys", "Capture the overflow",
      "Add hotel rooms where demand is already being turned away. Convert lost referrals into owned, recurring revenue."],
    ["03", "Destination Build", "Amenitize to lift value",
      "Replicate the winery-to-destination template — F&B, events, experiential — to raise both land and room values."],
    ["04", "Recycle & Scale", "Proceeds → next parcel",
      "Roll lot-sale and refinancing proceeds into the next property; compound the playbook across the portfolio."],
  ];
  const cardW = 2.94, gap = 0.18;
  let x = ML;
  items.forEach(([num, head, sub, body]) => {
    rect(s, x, 2.1, cardW, 4.1, C.navy);
    rect(s, x, 2.1, cardW, 0.1, C.aubergine);
    txt(s, num, x + 0.25, 2.35, cardW - 0.5, 0.9, {
      font: HEAD, size: 44, bold: true, color: C.mauve,
    });
    txt(s, head, x + 0.25, 3.35, cardW - 0.5, 0.4, {
      font: HEAD, size: 15, bold: true, color: C.white,
    });
    txt(s, sub, x + 0.25, 3.78, cardW - 0.5, 0.3, {
      font: HEAD, size: 10, bold: true, color: C.slate, spc: 1,
    });
    hline(s, x + 0.25, 4.18, cardW - 0.5, C.aubergine, 0.02);
    txt(s, body, x + 0.25, 4.32, cardW - 0.5, 1.75, {
      font: BODY, size: 10.5, color: C.cream, lh: 14, valign: "top",
    });
    x += cardW + gap;
  });
  callout(s, "Lever 01 needs mostly soft costs — the fastest path to proof, and to first cash.", 6.5);
  s.addNotes(
    "The sequencing is the pitch: start capital-light (entitle + lot sales), then recycle into keys " +
    "and destination build. Aligns with Bob saying they may only need soft costs to start."
  );
}

// ============================================== SLIDE 5 — FLAGSHIP: QUEENSTOWN
function flagship() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "FLAGSHIP ASSET",
    title: "Queenstown Harbor — 700 acres near D.C.",
    desc: "The anchor of the program: an established 36-hole destination with developable land around the courses and drive-to demand from Annapolis, Baltimore & Washington.",
    page: 5,
  });
  const cards = [
    ["700", "Acres owned", "36 holes already in the ground"],
    ["~45 min", "To Annapolis", "≈ 90 min to D.C. & Baltimore metros"],
    ["3", "Golf courses", "Established destination, existing traffic"],
    ["Soft costs", "To begin", "Entitlement-led; minimal capital at risk in Phase 1"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "THE MOVE",
    "Entitle residential lots around the existing courses and sell to a regional homebuilder.",
    [
      "Let the builder fund sewer, roads and horizontal infrastructure — we keep the upside, not the burden.",
      "Add hotel keys to capture the overflow the resort turns away today.",
      "Premium private-retreat demand is real — a recent week-long corporate buyout signals pricing power.",
      "Capital-light entry: Phase 1 is approvals and lot sales, not vertical construction.",
    ]);
  s.addNotes(
    "Queenstown is the wedge. 700 owned acres, established golf, close to D.C. " +
    "Lead with entitlement + lot sales (capital-light), keys second. The corporate-buyout note " +
    "is the Palantir week Bob mentioned — kept generic on purpose."
  );
}

// ========================================== SLIDE 6 — THE MATH: LAND (ILLUS.)
function mathLand() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE MATH  ·  ILLUSTRATIVE",
    title: "How lot entitlement creates value.",
    desc: "Illustrative only — to be calibrated with Accountable Equity's survey, entitlement status and absorption data. The method matters here, not the placeholder numbers.",
    page: 6,
  });
  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.0, color: C.navy, size: 10 },
    { w: 3.3, color: C.plum, size: 9, font: BODY },
    { w: 3.0, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "right" },
  ];
  const rows = [
    ["Developable land", "~25% of 700 ac to lots", "700 × 0.25", "175 ac"],
    ["Lot yield", "~2.0 lots / acre", "175 × 2.0", "350 lots"],
    ["Finished lot price", "$120,000 / lot", "350 × $120,000", "$42.0M"],
    ["Less land-dev / soft costs", "~30% of gross", "× 0.70", "$29.4M"],
    [{ text: "NET LOT PROCEEDS (illustrative)", bold: true, font: HEAD, color: C.navy },
      { text: "before hotel-key upside", color: C.plum }, "", "≈ $29M"],
  ];
  table(s, ML, 2.15, cols, ["Step", "Assumption", "Calculation", "= Value"], rows,
    { rowH: 0.62 });
  rect(s, ML, 5.85, CW, 0.55, C.navy);
  txt(s,
    "≈ $29M net lot proceeds — before any hotel-key or destination upside — and recyclable into the next parcel.",
    ML + 0.2, 5.85, CW - 0.4, 0.55,
    { font: HEAD, size: 12, bold: true, color: C.white, align: "center" });
  s.addNotes(
    "Be explicit that these are illustrative placeholders. The point is the method: " +
    "developable acres × lot yield × price, net of land-dev cost. Swap in real survey/absorption data."
  );
}

// ===================================== SLIDE 7 — THE MATH: HOSPITALITY (ILLUS.)
function mathHotel() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE MATH  ·  ILLUSTRATIVE",
    title: "Turned-away guests → owned revenue.",
    desc: "The resort is sending business to other hotels today. Even a modest key expansion converts referrals into NOI. Illustrative — calibrate with your occupancy and turn-away logs.",
    page: 7,
  });
  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.0, color: C.navy, size: 10 },
    { w: 3.3, color: C.plum, size: 9, font: BODY },
    { w: 3.0, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "right" },
  ];
  const rows = [
    ["New keys added", "120 rooms", "—", "120 keys"],
    ["Stabilized occupancy", "65%", "120 × 365 × 0.65", "28,470 nights"],
    ["ADR", "$185 / night", "× $185", "$5.27M rooms"],
    ["Total revenue incl. F&B / events", "~1.4× room revenue", "× 1.4", "$7.4M gross"],
    ["Stabilized NOI margin", "~35%", "× 0.35", "$2.6M NOI"],
    [{ text: "VALUE AT 8.0% CAP", bold: true, font: HEAD, color: C.navy },
      { text: "$2.6M ÷ 0.08", color: C.plum }, "", "≈ $32M"],
  ];
  table(s, ML, 2.15, cols, ["Driver", "Assumption", "Calculation", "= Result"], rows,
    { rowH: 0.52 });
  rect(s, ML, 5.65, CW, 0.55, C.navy);
  txt(s,
    "≈ $2.6M incremental NOI ≈ $32M of value at an 8% cap — created from demand you already have.",
    ML + 0.2, 5.65, CW - 0.4, 0.55,
    { font: HEAD, size: 12, bold: true, color: C.white, align: "center" });
  s.addNotes(
    "This directly answers Bob's 'sending business off and not making a dime.' " +
    "Modest 120-key add converts existing overflow into ~$2.6M NOI / ~$32M value. Illustrative."
  );
}

// ============================================== SLIDE 8 — CAPITAL STRATEGY
function capitalStrategy() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL STRATEGY",
    title: "Debt-led, land-as-equity, phased.",
    desc: "The land is the equity. Most of the program is financeable with debt and developer partners — the cash ask is modest and staged, not a single day-one check.",
    page: 8,
  });
  statCard(s, ML, 2.05, "Up to ~$100M", "Total program capacity",
    "Phased across the portfolio — not all at once", { w: 6.05 });
  statCard(s, ML, 3.12, "76%+", "Funded by land & debt",
    "Owned land contributed as equity; senior debt on top", { w: 6.05 });
  rect(s, 7.0, 2.05, 5.83, 2.02, C.navy);
  rect(s, 7.0, 2.05, 0.12, 2.02, C.aubergine);
  txt(s, "THE PRINCIPLE", 7.4, 2.2, 5.4, 0.3,
    { font: HEAD, size: 10, bold: true, color: C.mauve, spc: 4 });
  txt(s, "Use the balance sheet you already have.", 7.4, 2.5, 5.4, 0.35,
    { font: HEAD, size: 13, bold: true, color: C.white });
  txt(s, "The land is the equity; debt and homebuilder takedowns do the heavy lifting. Cash equity stays small and staged — soft costs first, scale as the parcels prove out.",
    7.4, 2.96, 5.4, 1.0, { font: BODY, size: 10.5, color: C.cream, lh: 14, valign: "top" });
  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 3.3, color: C.navy, size: 9.5 },
    { w: 2.2, color: C.aubergine, bold: true, font: HEAD, size: 10, align: "center" },
    { w: 3.8, color: C.plum, size: 9 },
  ];
  const rows = [
    ["Phase 1 — Entitlements", "Soft costs only", "$1–3M", "JAL + Bob's dev co (skin in the game)"],
    ["Phase 2 — Lot development", "Builder-funded horizontal", "~$0", "Homebuilder takedown"],
    ["Phase 3 — Hospitality keys", "Construction debt + equity", "$30–60M", "Hospitality lender + family-office equity"],
    ["Portfolio — Recycle", "Proceeds + refinancing", "Self-funding", "Roll forward into next parcel"],
  ];
  table(s, ML, 4.3, cols, ["Phase", "Approach", "Capital", "Source"], rows, { rowH: 0.46 });
  callout(s,
    "Start capital-light: a $1–3M soft-cost raise funds entitlements and first lot sales — proof before the big check.",
    6.5);
  s.addNotes(
    "Answers Bob's 'equity or debt — either.' Position: debt-led, land-as-equity. " +
    "Emphasize the modest, staged cash ask — de-risks the conversation."
  );
}

// =============================================== SLIDE 9 — CAPITAL ACCESS
function capitalAccess() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL ACCESS",
    title: "I bring the capital relationships.",
    desc: "Active mandates, not a paper Rolodex — debt and equity, on call. Here's what's moving right now and who I place with.",
    page: 9,
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
      "National debt — Goldman Sachs (met two weeks ago; national platform, Dallas team) + hospitality credit lenders.",
      "The bar for out-of-state development is mid-to-high-20s IRR — owned land helps the deal clear it.",
    ]);
  s.addNotes(
    "Credibility slide. Lead with the live Houston deal ($14M equity / $19.5M debt) as proof I actually close. " +
    "Then the network: TX family offices, Angelo Gordon, Goldman, hospitality credit."
  );
}

// =============================================== SLIDE 10 — THE STRUCTURE
function structure() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE STRUCTURE",
    title: "A vehicle: your dev co + JAL.",
    desc: "Formalize the development entity you're already standing up. Accountable Equity contributes land; the vehicle raises debt + equity; JAL co-invests and runs capital formation alongside your team.",
    page: 10,
  });
  // NewCo banner
  rect(s, ML, 1.95, CW, 0.86, C.navy);
  rect(s, ML, 1.95, 0.12, 0.86, C.aubergine);
  txt(s, "ACCOUNTABLE DEVELOPMENT PARTNERS, LLC", ML + 0.35, 2.06, 8.5, 0.4, {
    font: HEAD, size: 17, bold: true, color: C.white,
  });
  txt(s, "Working name", ML + 0.35, 2.46, 8.5, 0.3, {
    font: BODY, size: 10, color: C.mauve,
  });
  txt(s, "Development vehicle formed with Accountable Equity",
    9.2, 1.95, 3.6, 0.86, { font: BODY, size: 10, color: C.mauve, align: "right",
      valign: "middle", margin: [2, 10, 2, 6] });

  // three panels
  const px = [ML, 4.62, 8.74], pw = 3.93;
  // Panel A — Land = equity
  rect(s, px[0], 3.0, pw, 2.85, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "LAND = EQUITY", px[0] + 0.25, 3.18, pw - 0.5, 0.3,
    { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "Owned", px[0] + 0.25, 3.5, pw - 0.5, 0.6,
    { font: HEAD, size: 30, bold: true, color: C.navy });
  txt(s, "Acreage contributed at appraised value forms the equity base — the hardest, most expensive input is already on the balance sheet.",
    px[0] + 0.25, 4.2, pw - 0.5, 1.5, { font: BODY, size: 10.5, color: C.plum, lh: 14, valign: "top" });

  // Panel B — Capital stack
  rect(s, px[1], 3.0, pw, 2.85, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "CAPITAL STACK", px[1] + 0.25, 3.18, pw - 0.5, 0.3,
    { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "Illustrative", px[1] + 0.25, 3.46, pw - 0.5, 0.25,
    { font: BODY, size: 9, color: C.plum });
  const stack = [
    ["Land equity (AE)", "Contributed"],
    ["Construction / acq. debt", "Senior"],
    ["Cash equity (JAL + LP)", "Co-invest"],
  ];
  let sy = 3.78;
  stack.forEach(([a, b]) => {
    txt(s, a, px[1] + 0.25, sy, 2.5, 0.34, { font: BODY, size: 10, color: C.navy });
    txt(s, b, px[1] + 2.6, sy, pw - 2.85, 0.34, { font: HEAD, size: 9.5, bold: true, color: C.aubergine, align: "right" });
    hline(s, px[1] + 0.25, sy + 0.36, pw - 0.5, C.mauve, 0.01);
    sy += 0.55;
  });

  // Panel C — Roles
  rect(s, px[2], 3.0, pw, 2.85, C.navy);
  rect(s, px[2], 3.0, 0.1, 2.85, C.aubergine);
  txt(s, "ROLES", px[2] + 0.25, 3.18, pw - 0.5, 0.3,
    { font: HEAD, size: 11, bold: true, color: C.mauve, spc: 2 });
  const roles = [
    ["Josh McCallan / AE", "Owner & hospitality operator"],
    ["Bob Connell", "Development lead"],
    ["JAL Strategies", "Capital + execution — co-invests"],
  ];
  let ry = 3.62;
  roles.forEach(([a, b]) => {
    txt(s, a, px[2] + 0.25, ry, pw - 0.5, 0.28, { font: HEAD, size: 11, bold: true, color: C.white });
    txt(s, b, px[2] + 0.25, ry + 0.27, pw - 0.5, 0.3, { font: BODY, size: 9.5, color: C.cream });
    ry += 0.72;
  });

  callout(s, "JAL takes skin in the game — co-investing alongside, not just advising.", 6.5);
  s.addNotes(
    "Bob said he's setting up a development company. Proposal: formalize it as a JV vehicle. " +
    "AE contributes land = equity; JAL co-invests + runs capital formation. Skin in the game throughout."
  );
}

// ==================================== SLIDE 11 — ENGAGEMENT & COMPENSATION
function engagement() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "ENGAGEMENT  ·  FOR DISCUSSION",
    title: "How I'd want to be engaged.",
    desc: "You asked how I'd like to be comped. My preference matches yours — a small base to fund the work, then real alignment on what we create together. For discussion only.",
    page: 11,
  });
  // summary band
  rect(s, ML, 1.95, CW, 0.95, C.navy);
  rect(s, ML, 1.95, 0.12, 0.95, C.aubergine);
  txt(s, "Hybrid", ML + 0.35, 2.02, 3.0, 0.8, { font: HEAD, size: 30, bold: true, color: C.mauve });
  txt(s, "Retainer + success fees + carried interest", 3.6, 2.06, 5.4, 0.4,
    { font: HEAD, size: 14, bold: true, color: C.white });
  txt(s, "Aligned to capital placed and value created — not a flat consulting check.",
    3.6, 2.46, 5.6, 0.35, { font: BODY, size: 10, color: C.cream });
  txt(s, "Skin in\nthe game", 9.6, 1.95, 3.23, 0.95, { font: HEAD, size: 13, bold: true,
    color: C.slate, align: "right", valign: "middle", lh: 15, margin: [2, 10, 2, 6] });

  const cols = [
    { w: 3.4, font: HEAD, bold: true, color: C.navy, size: 11 },
    { w: 3.0, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "center" },
    { w: 5.93, color: C.plum, size: 9.5, lh: 12 },
  ];
  const rows = [
    ["Advisory retainer", "$15–25K / mo", "Funds analysis, entitlement strategy & capital sourcing during an initial term (creditable against success fees)"],
    ["Capital placement fee", "1% debt · 2–3% equity", "Standard success fee on capital actually placed — debt and equity"],
    ["Carried interest", "Co-GP promote", "Meaningful share of value created; JAL co-invests its own capital alongside the deal"],
  ];
  table(s, ML, 3.15, cols, ["Component", "Terms", "What it covers"], rows, { rowH: 0.78 });
  callout(s, "The structure flexes to the deal — the point is alignment: I win when you win.", 6.5);
  s.addNotes(
    "Directly answers Bob's question on comp. He said his preference is skin in the game; I agreed and " +
    "proposed a hybrid. Retainer is modest and creditable — emphasize the carry/co-invest as the real alignment. " +
    "All numbers are opening positions for discussion."
  );
}

// =============================================== SLIDE 12 — WHY JAL (BIO)
function whyJAL() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "WHY JAL AS YOUR PARTNER",
    title: "Pedigree. Relationships. Operator.",
    desc: "Institutional capital-markets pedigree, an active investor network, and hands-on development & asset-management experience.",
    page: 12,
  });
  // left navy panel
  rect(s, ML, 1.85, 5.0, 4.75, C.navy);
  rect(s, ML, 1.85, 0.12, 4.75, C.aubergine);
  txt(s, "JUSTIN A. LEVINE", ML + 0.3, 2.03, 4.5, 0.3, {
    font: HEAD, size: 10, bold: true, color: C.mauve, spc: 4,
  });
  txt(s, "Founder & CEO, JAL Strategies", ML + 0.3, 2.36, 4.5, 0.5, {
    font: HEAD, size: 19, bold: true, color: C.white,
  });
  const stats = [
    ["$6.1B", "career transaction volume ($5.5B Blackstone + $600M Levcor)"],
    ["$600M", "debt + equity allocated across 26 CRE deals at Levcor"],
    ["3.9M SF", "TX + NC retail managed at Levcor (2014–2025)"],
    ["20+ years", "CRE operating, capital markets & asset management"],
  ];
  let yy = 3.0;
  stats.forEach(([n, d]) => {
    txt(s, n, ML + 0.3, yy, 2.05, 0.55, { font: HEAD, size: 21, bold: true, color: C.slate });
    txt(s, d, ML + 2.35, yy + 0.03, 2.45, 0.72, { font: BODY, size: 9, color: C.cream, lh: 11.5, valign: "middle" });
    yy += 0.85;
  });

  // right white panel — career arc
  rect(s, 5.8, 1.85, 7.03, 4.75, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "CAREER ARC", 6.05, 2.03, 6.6, 0.3, {
    font: HEAD, size: 10, bold: true, color: C.aubergine, spc: 3,
  });
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
    "Bob's email literally said 'review your resume.' This is it. Capital-markets pedigree (Blackstone), " +
    "operator/development (Levcor, JAL-JCP), Wharton, ULI leadership."
  );
}

// =============================================== SLIDE 13 — THE PATH
function path() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PATH",
    title: "From this call to a signed mandate.",
    desc: "Light and fast to start — diligence now, an in-person while I'm in the Northeast in July, then a defined engagement.",
    page: 13,
  });
  const stages = [
    ["STAGE 1", "DILIGENCE & DATA", "Now – July", "Review & analysis",
      ["Review AE sites (web, maps & visits)",
       "You send project data (entitlements, occupancy / turn-away, capital stack)",
       "JAL refines the monetization model per parcel",
       "Draft engagement & vehicle terms"]],
    ["STAGE 2", "IN PERSON", "July 2026", "Philadelphia / Yardley",
      ["Meet in person while I'm in the NY / Philly area",
       "Walk the priority sites (Queenstown / winery HQ)",
       "Align on first parcel, capital plan & comp",
       "Meet Josh & key stakeholders"]],
    ["STAGE 3", "MANDATE", "H2 2026", "Engage & execute",
      ["Stand up the development vehicle",
       "Launch Phase 1 entitlements + soft-cost raise",
       "Begin capital sourcing (debt + equity)",
       "First lot sales / key-expansion underwriting"]],
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
    "What I need to sharpen the model: entitlement status per parcel · occupancy & turn-away data · existing debt · the $100M use-of-funds · target timeline.",
    6.5);
  s.addNotes(
    "Close on logistics. The July in-person is real (wife's family on Long Island). " +
    "The callout is the data ask — restate it so Bob knows exactly what to send."
  );
}

// =============================================== SLIDE 14 — THANK YOU
function thankYou() {
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  rect(s, 0, 0, 0.25, PAGE_H, C.aubergine);
  txt(s, "JAL STRATEGIES", 0.75, 0.6, 6.0, 0.4, {
    font: HEAD, size: 13, bold: true, color: C.mauve, spc: 6,
  });
  txt(s, "THANK", 0.72, 2.45, 12.0, 1.05, { font: HEAD, size: 72, bold: true, color: C.white });
  txt(s, "YOU", 0.72, 3.38, 12.0, 1.05, { font: HEAD, size: 72, bold: true, color: C.mauve });
  hline(s, 0.78, 4.62, 3.5, C.slate, 0.03);
  txt(s, "Bob — appreciate the call, and the résumé read.", 0.75, 4.78, 12.0, 0.45, {
    font: HEAD, size: 21, bold: true, color: C.white,
  });
  txt(s, "Looking forward to building this with you and the Accountable Equity team — let's find time in July.",
    0.75, 5.3, 12.0, 0.4, { font: BODY, size: 14, color: C.slate });
  txt(s, "Justin A. Levine  |  jlevine@jalstrategies.com  |  JAL Strategies",
    0.75, 5.74, 12.0, 0.4, { font: BODY, size: 12, color: C.mauve });
  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, DECK_LABEL, 0.5, 7.0, 9.0, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 2 });
  txt(s, `${TOTAL} / ${TOTAL}`, 9.7, 7.0, 3.13, 0.4, {
    font: HEAD, size: 9, bold: true, color: C.mauve, spc: 2, align: "right",
  });
  s.addNotes("Close warm. Reference his email ('review your resume'). Lock the July meeting.");
}

// ----- Build ---------------------------------------------------------------
cover();
opportunity();
portfolio();
playbook();
flagship();
mathLand();
mathHotel();
capitalStrategy();
capitalAccess();
structure();
engagement();
whyJAL();
path();
thankYou();

const OUT = "JAL_Accountable_Equity_Proposal.pptx";
pptx.writeFile({ fileName: OUT }).then((f) => console.log("Wrote", f));
