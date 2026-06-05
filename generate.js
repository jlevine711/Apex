/**
 * JAL Strategies — Golf-Course Land Development Partnership
 * Prepared for Robert A. Connell, CFP (Apex Financial Advisors / Accountable Equity)
 *
 * Scope: the developable LAND BANK around Accountable Equity's golf courses —
 * entitle the acreage, sell finished lots to homebuilders (builders fund the
 * horizontal infrastructure), recycle proceeds. Led by Queenstown Harbor.
 *
 * Design system reverse-engineered from the JAL "Republic Square" pitch deck:
 *   - 16:9 (13.33" x 7.5"); Montserrat (headings) / DM Sans (body)
 *   - Palette: navy #0B163C, aubergine #3A243A, plum #6B5A6B,
 *     mauve #C4B8C4, cream #F4F2ED, slate #8C9BB5, white
 *
 * Content grounded in the 6/5/2026 Levine–Connell call (Plaud transcript).
 * Figures on "THE MATH" / capital slides are clearly labelled ILLUSTRATIVE —
 * to be recalibrated against Accountable Equity survey / entitlement data.
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
const TOTAL = 14;

const DECK_LABEL = "ACCOUNTABLE EQUITY  |  CONFIDENTIAL";
const FOOTER_LEFT =
  "JAL Strategies  |  Justin A. Levine, Founder & CEO  |  jlevine@jalstrategies.com";

const pptx = new pptxgen();
pptx.defineLayout({ name: "W16", width: PAGE_W, height: PAGE_H });
pptx.layout = "W16";
pptx.author = "JAL Strategies";
pptx.company = "JAL Strategies";
pptx.title =
  "Accountable Equity — Golf-Course Land Development Partnership (for Robert A. Connell)";

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

// Shape-grid table. colDefs: [{w, align, color, font, bold, size, lh}].
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

  txt(s, "ACCOUNTABLE", 0.72, 2.18, 12.2, 1.05, {
    font: HEAD, size: 67, bold: true, color: C.white, spc: 1,
  });
  txt(s, "EQUITY", 0.72, 3.12, 12.2, 1.05, {
    font: HEAD, size: 67, bold: true, color: C.mauve, spc: 1,
  });

  hline(s, 0.78, 4.42, 3.5, C.slate, 0.03);
  txt(s, "Golf-Course Land Development Partnership", 0.75, 4.58, 12.0, 0.45, {
    font: HEAD, size: 22, bold: true, color: C.white,
  });
  txt(s, "Entitling and selling the residential land around the courses you already own",
    0.75, 5.12, 12.0, 0.35, { font: BODY, size: 14, color: C.slate });
  txt(s,
    "Queenstown Harbor  ·  LBI National  ·  700+ acres of golf-frontage land",
    0.75, 5.52, 12.0, 0.35, { font: BODY, size: 12, color: C.mauve });

  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, "PREPARED FOR ROBERT A. CONNELL, CFP   |   APEX FINANCIAL ADVISORS · ACCOUNTABLE EQUITY",
    0.5, 7.0, 9.6, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5 });
  txt(s, "JUNE 2026  |  CONFIDENTIAL", 9.7, 7.0, 3.13, 0.4, {
    font: HEAD, size: 9, bold: true, color: C.mauve, spc: 1.5, align: "right",
  });
  s.addNotes(
    "Cover. The discussion document Bob asked for on our 6/5 call — scoped to the land bank " +
    "around the golf courses: entitle, sell lots to builders, recycle. Capital-light, JAL-led."
  );
}

// ===================================================== SLIDE 2 — OPPORTUNITY
function opportunity() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE OPPORTUNITY",
    title: "Sell the land around the courses you own.",
    desc: "Bob — the fastest, lowest-risk money here is the developable land wrapped around the golf courses: entitle it, sell finished lots to homebuilders, recycle the proceeds.",
    page: 2,
  });
  const cards = [
    ["700 ac", "At Queenstown Harbor", "3 courses · ~45 min to Annapolis & D.C."],
    ["~175 ac", "Developable to lots", "Around the fairways (illustrative)"],
    ["~350 lots", "Entitlement potential", "Sold to regional homebuilders (illustrative)"],
    ["Soft costs", "To begin", "Builders fund the horizontal infrastructure"],
  ];
  let cy = 2.05;
  cards.forEach(([n, l, sub]) => { statCard(s, ML, cy, n, l, sub); cy += 1.07; });

  sidePanel(s, 7.0, 2.05, 5.83, 4.5, "WHY THE LAND BANK FIRST",
    "The equity is already in the ground — you own the land free and clear.",
    [
      "Builders fund sewer, roads and utilities — you carry approvals, not construction risk.",
      "Golf frontage commands a premium — lots on a course sell faster and for more.",
      "One town has offered to contribute land — the entitlement tailwinds are real.",
      "Proceeds recycle — lot sales can seed the broader Accountable Equity program.",
    ]);
  s.addNotes(
    "Frame the land bank as the first move: owned land, builder-funded infrastructure, golf premium, " +
    "town tailwinds. Capital-light and fast — exactly the play Bob leaned toward on the call."
  );
}

// ========================================================== SLIDE 3 — THE LAND
function theLand() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE LAND",
    title: "Golf-course land, course by course.",
    desc: "Each course already amenitizes the land around it. We entitle the developable acreage and sell finished lots — and the same playbook repeats on every course in the portfolio.",
    page: 3,
  });
  const E = (t, color = C.navy) => ({ text: t, bold: true, color, font: HEAD });
  const cols = [
    { w: 2.5, font: HEAD, bold: true, color: C.navy, size: 10 },
    { w: 2.0, color: C.plum, size: 9.5 },
    { w: 3.6, color: C.navy, size: 9.5, lh: 11 },
    { w: 4.23, color: C.navy, size: 9.5, lh: 11 },
  ];
  const rows = [
    ["Queenstown Harbor", "Queenstown, MD", "700 ac · 3 courses · ~45 min to Annapolis / D.C.",
      "Entitle & sell residential lots around the fairways; builder-funded infrastructure"],
    ["LBI National Golf", "Long Beach Island, NJ", "Coastal golf in a high-barrier shore market",
      "Scarce Jersey-Shore lots; premium entitled-lot sales"],
    ["Recently acquired course", "Mid-Atlantic", "Newly added to the golf portfolio",
      "Repeat the entitle-and-sell template on fresh acreage"],
    [E("Across the portfolio"), E("MD · NJ", C.aubergine),
      E("700+ developable acres around owned courses"),
      E("One repeatable, capital-light playbook", C.aubergine)],
  ];
  table(s, ML, 2.05, cols, ["Course", "Location", "Setting", "Land-bank thesis"], rows,
    { rowH: 0.82 });
  callout(s,
    "Start at Queenstown — the largest parcel, closest to D.C. — then repeat the template course by course.",
    6.5);
  s.addNotes(
    "Keep it golf-only. Queenstown is the anchor (700 ac, 3 courses near D.C.); LBI National is the " +
    "shore-scarcity play; the recently-acquired course is pipeline. Same template each time."
  );
}

// ======================================================== SLIDE 4 — PLAYBOOK
function playbook() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE PLAYBOOK",
    title: "Four moves that turn land into cash.",
    desc: "Capital-light by sequence — soft costs first, then builders carry the infrastructure and the construction risk, and the proceeds recycle into the next course.",
    page: 4,
  });
  const items = [
    ["01", "Entitle", "Approvals for lots",
      "Secure zoning and subdivision approvals for the developable acreage around the fairways."],
    ["02", "Sell to Builders", "Finished lots → homebuilders",
      "Sell entitled lots to regional builders and let them fund sewer, roads and utilities."],
    ["03", "Golf Premium", "Frontage sells higher",
      "Lots on or near a course command a premium and absorb faster than a comparable raw subdivision."],
    ["04", "Recycle & Repeat", "Course by course",
      "Roll proceeds into the next course's land bank — and seed the broader Accountable Equity program."],
  ];
  const cardW = 2.94, gap = 0.18;
  let x = ML;
  items.forEach(([num, head, sub, body]) => {
    rect(s, x, 2.1, cardW, 4.1, C.navy);
    rect(s, x, 2.1, cardW, 0.1, C.aubergine);
    txt(s, num, x + 0.25, 2.35, cardW - 0.5, 0.9, { font: HEAD, size: 44, bold: true, color: C.mauve });
    txt(s, head, x + 0.25, 3.35, cardW - 0.5, 0.4, { font: HEAD, size: 15, bold: true, color: C.white });
    txt(s, sub, x + 0.25, 3.78, cardW - 0.5, 0.3, { font: HEAD, size: 10, bold: true, color: C.slate, spc: 1 });
    hline(s, x + 0.25, 4.18, cardW - 0.5, C.aubergine, 0.02);
    txt(s, body, x + 0.25, 4.32, cardW - 0.5, 1.75, { font: BODY, size: 10.5, color: C.cream, lh: 14, valign: "top" });
    x += cardW + gap;
  });
  callout(s, "Move 01 is mostly soft costs — the fastest path to proof, and to first cash.", 6.5);
  s.addNotes(
    "The sequencing IS the pitch: entitle (soft costs) → sell to builders (they fund infra) → " +
    "golf premium → recycle. Matches Bob's 'let the developer handle the infrastructure.'"
  );
}

// ============================================== SLIDE 5 — FLAGSHIP: QUEENSTOWN
function flagship() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "FLAGSHIP PARCEL",
    title: "Queenstown Harbor — 700 acres near D.C.",
    desc: "The anchor of the land bank: an established 36-hole destination with developable acreage around the courses and drive-to demand from Annapolis, Baltimore & Washington.",
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
      "Let the builder fund sewer, roads and infrastructure — we keep the upside, not the burden.",
      "Golf-frontage lots near Annapolis & D.C. command a premium and absorb quickly.",
      "Location pull is real — a recent week-long corporate buyout signals premium demand for the setting.",
      "Capital-light entry: Phase 1 is approvals and lot sales, not vertical construction.",
    ]);
  s.addNotes(
    "Queenstown is the wedge: 700 owned acres, established golf, an hour from D.C. " +
    "Lead with entitlement + lot sales. The corporate-buyout note is the Palantir week Bob mentioned — " +
    "kept generic; name it if Bob's comfortable."
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
      { text: "Queenstown alone", color: C.plum }, "", "≈ $29M"],
  ];
  table(s, ML, 2.15, cols, ["Step", "Assumption", "Calculation", "= Value"], rows, { rowH: 0.62 });
  rect(s, ML, 5.85, CW, 0.55, C.navy);
  txt(s,
    "≈ $29M net proceeds from Queenstown lots alone — recyclable into LBI National and the next course.",
    ML + 0.2, 5.85, CW - 0.4, 0.55,
    { font: HEAD, size: 12, bold: true, color: C.white, align: "center" });
  s.addNotes(
    "Be explicit these are illustrative placeholders. The method: developable acres × lot yield × " +
    "price, net of land-dev cost. Swap in real survey / absorption data. Builders typically fund the horizontal."
  );
}

// ===================================================== SLIDE 7 — WHY IT SELLS
function whySells() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "WHY IT SELLS",
    title: "Golf-course lots sell — here's why.",
    desc: "Course frontage, drive-to metro demand and builder appetite for entitled lots all push price and absorption. We're meeting demand that already exists.",
    page: 7,
  });
  const cols = [
    { w: 3.0, font: HEAD, bold: true, color: C.navy, size: 10.5 },
    { w: 4.8, color: C.navy, size: 9.5, lh: 11 },
    { w: 4.53, color: C.plum, size: 9.5, lh: 11 },
  ];
  const rows = [
    ["Golf-course frontage", "Lots on or adjacent to an established course",
      "Premium pricing and faster absorption vs. a raw subdivision"],
    ["Queenstown location", "~45 min to Annapolis; ~90 to D.C. & Baltimore",
      "Commuter + second-home demand from three metros"],
    ["LBI National scarcity", "Limited developable land on the Jersey Shore",
      "High barriers → durable, premium shore-market lot values"],
    ["Builder appetite", "Regional homebuilders want entitled / finished lots",
      "They pay to avoid entitlement risk — and fund the infrastructure"],
    ["Owned land basis", "Acreage already on the balance sheet",
      "No land-acquisition cost to recover — margin starts higher"],
  ];
  table(s, ML, 2.15, cols, ["Driver", "Detail", "Why it lifts value"], rows, { rowH: 0.62 });
  callout(s, "Demand is already here — we're entitling supply to meet it, not creating a market.", 6.5);
  s.addNotes(
    "Mirrors the reference's competitive-supply slide. Point: golf frontage + metro demand + builder " +
    "appetite + owned-land basis = lots that sell. Land development is a good business when the strategy is right."
  );
}

// ============================================== SLIDE 8 — CAPITAL: LIGHT
function capitalLight() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL STRATEGY",
    title: "Capital-light by design.",
    desc: "The land bank barely needs capital: soft costs to entitle, then builders fund the horizontal work and buy the lots. Cash in is small; proceeds come fast.",
    page: 8,
  });
  statCard(s, ML, 2.05, "$1–3M", "Soft costs to start",
    "Entitlements, planning & approvals", { w: 6.05 });
  statCard(s, ML, 3.12, "Land", "= the equity",
    "Owned acreage carries the deal — no acquisition cost", { w: 6.05 });

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
    ["Phase 1 — Entitle", "Approvals for lots", "$1–3M", "JAL + Bob's dev co (skin in the game)"],
    ["Phase 2 — Builder takedown", "Horizontal + lot purchase", "~$0 to you", "Regional homebuilder funds & buys"],
    ["Phase 3 — Recycle", "Proceeds redeployed", "Self-funding", "Next course / broader AE program"],
  ];
  table(s, ML, 4.3, cols, ["Phase", "Approach", "Your capital", "Who funds"], rows, { rowH: 0.5 });
  callout(s, "Minimal capital at risk — and the land bank seeds everything that comes after.", 6.5);
  s.addNotes(
    "Reframed from the $100M hospitality line. For the land bank, the honest story is capital-light: " +
    "soft costs in, builders fund infra and buy lots. Bob said as much — they may only need soft costs."
  );
}

// =============================================== SLIDE 9 — CAPITAL ACCESS
function capitalAccess() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "CAPITAL ACCESS",
    title: "I bring the capital relationships.",
    desc: "When a parcel does call for equity or debt — or to move faster — here's the access. Active mandates, not a paper Rolodex.",
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
      "National debt — Goldman Sachs (met two weeks ago; national platform, Dallas team) + specialty lenders.",
      "The bar for out-of-state development is mid-to-high-20s IRR — owned land helps the deal clear it.",
    ]);
  s.addNotes(
    "Bob's inbound was literally about capital access. Lead with the live Houston deal ($14M / $19.5M) " +
    "as proof I close, then the network. Even capital-light, this is why I'm useful."
  );
}

// =============================================== SLIDE 10 — THE STRUCTURE
function structure() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "THE STRUCTURE",
    title: "A vehicle: your dev co + JAL.",
    desc: "Formalize the development entity you're standing up — scoped to the course-adjacent land. Accountable Equity contributes the acreage as equity; JAL co-invests and runs entitlement, capital and lot sales alongside your team.",
    page: 10,
  });
  rect(s, ML, 1.95, CW, 0.86, C.navy);
  rect(s, ML, 1.95, 0.12, 0.86, C.aubergine);
  txt(s, "ACCOUNTABLE LAND PARTNERS, LLC", ML + 0.35, 2.06, 8.5, 0.4, { font: HEAD, size: 17, bold: true, color: C.white });
  txt(s, "Working name", ML + 0.35, 2.46, 8.5, 0.3, { font: BODY, size: 10, color: C.mauve });
  txt(s, "Land-development vehicle for the golf-course acreage", 9.2, 1.95, 3.6, 0.86,
    { font: BODY, size: 10, color: C.mauve, align: "right", valign: "middle", margin: [2, 10, 2, 6] });

  const px = [ML, 4.62, 8.74], pw = 3.93;
  rect(s, px[0], 3.0, pw, 2.85, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "LAND = EQUITY", px[0] + 0.25, 3.18, pw - 0.5, 0.3, { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "Owned", px[0] + 0.25, 3.5, pw - 0.5, 0.6, { font: HEAD, size: 30, bold: true, color: C.navy });
  txt(s, "Course-adjacent acreage contributed at appraised value forms the equity base — already owned, no acquisition cost.",
    px[0] + 0.25, 4.2, pw - 0.5, 1.5, { font: BODY, size: 10.5, color: C.plum, lh: 14, valign: "top" });

  rect(s, px[1], 3.0, pw, 2.85, C.white, { line: { color: C.hair, width: 0.75 } });
  txt(s, "CAPITAL STACK", px[1] + 0.25, 3.18, pw - 0.5, 0.3, { font: HEAD, size: 11, bold: true, color: C.aubergine, spc: 2 });
  txt(s, "Illustrative", px[1] + 0.25, 3.46, pw - 0.5, 0.25, { font: BODY, size: 9, color: C.plum });
  const stack = [
    ["Land equity (AE)", "Contributed"],
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
    ["Josh McCallan / AE", "Land owner"],
    ["Bob Connell", "Development lead"],
    ["JAL Strategies", "Capital, entitlement & sales — co-invests"],
  ];
  let ry = 3.62;
  roles.forEach(([a, b]) => {
    txt(s, a, px[2] + 0.25, ry, pw - 0.5, 0.28, { font: HEAD, size: 11, bold: true, color: C.white });
    txt(s, b, px[2] + 0.25, ry + 0.27, pw - 0.5, 0.34, { font: BODY, size: 9.5, color: C.cream, lh: 11 });
    ry += 0.72;
  });

  callout(s, "JAL takes skin in the game — co-investing alongside, not just advising.", 6.5);
  s.addNotes(
    "Formalize Bob's dev co as a land JV. AE contributes course-adjacent land = equity; JAL co-invests " +
    "and runs entitlement + capital + lot sales. Builders fund the horizontal."
  );
}

// ==================================== SLIDE 11 — ENGAGEMENT & COMPENSATION
function engagement() {
  const s = pptx.addSlide();
  chrome(s, {
    eyebrow: "ENGAGEMENT  ·  FOR DISCUSSION",
    title: "How I'd want to be engaged.",
    desc: "You asked how I'd like to be comped. My preference matches yours — a small base to fund the work, then real alignment on the value we create. For discussion only.",
    page: 11,
  });
  rect(s, ML, 1.95, CW, 0.95, C.navy);
  rect(s, ML, 1.95, 0.12, 0.95, C.aubergine);
  txt(s, "Hybrid", ML + 0.35, 2.02, 3.0, 0.8, { font: HEAD, size: 30, bold: true, color: C.mauve });
  txt(s, "Retainer + success fees + carried interest", 3.6, 2.06, 5.4, 0.4, { font: HEAD, size: 14, bold: true, color: C.white });
  txt(s, "Aligned to lot-sale value created — not a flat consulting check.", 3.6, 2.46, 5.6, 0.35, { font: BODY, size: 10, color: C.cream });
  txt(s, "Skin in\nthe game", 9.6, 1.95, 3.23, 0.95, { font: HEAD, size: 13, bold: true, color: C.slate, align: "right", valign: "middle", lh: 15, margin: [2, 10, 2, 6] });

  const cols = [
    { w: 3.4, font: HEAD, bold: true, color: C.navy, size: 11 },
    { w: 3.0, color: C.aubergine, bold: true, font: HEAD, size: 11, align: "center" },
    { w: 5.93, color: C.plum, size: 9.5, lh: 12 },
  ];
  const rows = [
    ["Advisory retainer", "$15–25K / mo", "Funds entitlement strategy, underwriting & capital sourcing during an initial term (creditable against success fees)"],
    ["Capital placement fee", "1% debt · 2–3% equity", "Success fee if and when we place debt or equity for a parcel"],
    ["Carried interest", "Co-GP promote", "Share of value created on lot sales; JAL co-invests its own capital alongside"],
  ];
  table(s, ML, 3.15, cols, ["Component", "Terms", "What it covers"], rows, { rowH: 0.78 });
  callout(s, "The structure flexes to the deal — the point is alignment: I win when you win.", 6.5);
  s.addNotes(
    "Directly answers Bob's comp question. He said his preference is skin in the game; I agree and propose " +
    "a hybrid — modest, creditable retainer + carry on lot-sale value. All numbers are opening positions."
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
    desc: "Light and fast to start — diligence now, an in-person while I'm in the Northeast in July, then a defined engagement on the first parcel.",
    page: 13,
  });
  const stages = [
    ["STAGE 1", "DILIGENCE & DATA", "Now – July", "Review & analysis",
      ["Review the courses (web, maps & visits)",
       "You send parcel data (survey, entitlement status, zoning, encumbrances)",
       "JAL refines the lot-monetization model per course",
       "Draft engagement & vehicle terms"]],
    ["STAGE 2", "IN PERSON", "July 2026", "Philadelphia / Yardley",
      ["Meet in person while I'm in the NY / Philly area",
       "Walk the priority parcels (Queenstown / LBI)",
       "Align on first course, capital plan & comp",
       "Meet Josh & key stakeholders"]],
    ["STAGE 3", "MANDATE", "H2 2026", "Engage & execute",
      ["Stand up the land-development vehicle",
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
    "What I need to sharpen the model: survey / plat & acreage · entitlement & zoning status · any debt on the land · target timeline.",
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
  txt(s, "JAL STRATEGIES", 0.75, 0.6, 6.0, 0.4, { font: HEAD, size: 13, bold: true, color: C.mauve, spc: 6 });
  txt(s, "THANK", 0.72, 2.45, 12.0, 1.05, { font: HEAD, size: 72, bold: true, color: C.white });
  txt(s, "YOU", 0.72, 3.38, 12.0, 1.05, { font: HEAD, size: 72, bold: true, color: C.mauve });
  hline(s, 0.78, 4.62, 3.5, C.slate, 0.03);
  txt(s, "Bob — appreciate the call, and the résumé read.", 0.75, 4.78, 12.0, 0.45, { font: HEAD, size: 21, bold: true, color: C.white });
  txt(s, "Looking forward to entitling and monetizing the course lands with you and the Accountable Equity team — let's find time in July.",
    0.75, 5.3, 12.0, 0.4, { font: BODY, size: 14, color: C.slate });
  txt(s, "Justin A. Levine  |  jlevine@jalstrategies.com  |  JAL Strategies", 0.75, 5.74, 12.0, 0.4, { font: BODY, size: 12, color: C.mauve });
  rect(s, 0, 6.9, PAGE_W, 0.6, C.aubergine);
  txt(s, DECK_LABEL, 0.5, 7.0, 9.0, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 2 });
  txt(s, `${TOTAL} / ${TOTAL}`, 9.7, 7.0, 3.13, 0.4, { font: HEAD, size: 9, bold: true, color: C.mauve, spc: 2, align: "right" });
  s.addNotes("Close warm. Reference his email ('review your resume'). Lock the July meeting.");
}

// ----- Build ---------------------------------------------------------------
cover();
opportunity();
theLand();
playbook();
flagship();
mathLand();
whySells();
capitalLight();
capitalAccess();
structure();
engagement();
whyJAL();
path();
thankYou();

const OUT = "JAL_Accountable_Equity_Proposal.pptx";
pptx.writeFile({ fileName: OUT }).then((f) => console.log("Wrote", f));
