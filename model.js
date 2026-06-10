/**
 * Queenstown Harbor — Entitled-Lot (Capital-Light) Development Model.
 *
 * CAPITAL-LIGHT basis (per Bob, 6/9): we fund only the SOFT costs to entitle the
 * land, then sell ENTITLED lots to builders — the builder funds all the
 * horizontal (roads, sewer, utilities). Costs are repaid first, then profit is
 * split a simple 50/50: half to the property owner (the "H" entities; Josh owns
 * 50%, so ~25% of profit) and half to the development company (Bob · JAL ·
 * partners). No pref, no promote waterfall — a straight deal.
 *
 * DYNAMIC & CALCULATED. A single INPUTS object (`IN`) feeds pure calc engines
 * (`calcP1`, `calcP2`) that derive every figure; in the workbook the INPUTS plus
 * the cash-flow timing vectors are the only typed numbers and every other cell
 * is a live Excel formula (recalcs on open). All figures illustrative.
 *
 *   npm run model   ->   JAL_Queenstown_Harbor_Model.xlsx
 */
const ExcelJS = require("exceljs");

// ---- IRR (bisection) — caches the value Excel shows before it recalculates --
function irr(cfs) {
  const npv = (r) => cfs.reduce((a, c, i) => a + c / Math.pow(1 + r, i), 0);
  let lo = -0.9, hi = 9;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (npv(mid) > 0) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

// ============================================================================
// INPUTS — the single source of truth.  Edit here (or in the sheet) and
// everything below recalculates.  Phasing vectors run Yr 0..4 and sum to 1.
// ============================================================================
const IN = {
  // ---- Phase 1 — entitled-lot (capital-light) ----
  grossAc: 700, consAc: 198, golfAc: 302, critAc: 60, density: 2.0,
  entitledPx: 70000,            // price of an ENTITLED (paper) lot to a builder
  // soft costs to entitle (we fund these; builder funds the horizontal)
  entitlement: 1200000,         // planning, zoning, PUD approvals
  civilEng: 1000000,            // civil engineering / construction plans
  environmental: 600000,        // wetlands, Critical Area, studies
  legalSurvey: 500000,          // legal, survey, title
  projMgmt: 700000,             // project management / consultants
  mktPct: 0.03, contPct: 0.10, fin: 300000,
  split: 0.50,                  // profit split — development company share = 1 − split
  joshOfProperty: 0.50,         // Josh's ownership of each property entity
  jalShareOfDevCo: 0.30,        // JAL's sweat-equity share of the dev co's 50% (~15% of profit)
  retainerMo: 15000, engMonths: 36, // $15K/mo retainer over the active engagement (illustrative)
  // finished-lot comparison (what building it out ourselves would look like)
  finishedPx: 150000, horizPerLot: 60000, finFin: 1400000,
  // cash-flow timing (% by year, Yr 0..4)
  revPhase:  [0, 0, 0.25, 0.40, 0.35],   // entitled-lot sales as approvals land, Yr 2–4
  costPhase: [0.30, 0.45, 0.18, 0.07, 0],// soft costs front-loaded Yr 0–2

  // ---- Phase 2 — hospitality (hotel + restaurants); hotel is REQUIRED ----
  hKeys: 100, hCostKey: 325000, hADR: 250, hOcc: 0.62, hMult: 1.6, hMgn: 0.31,
  rSF: 16000, rCostSF: 550, rSalesSF: 700, rMgn: 0.09,
  pLand: 2000000, pSoft: 3700000, pCap: 0.08, pLTC: 0.60, pRate: 0.07, pGrow: 0.03,
  eqSplit: 0.5, rampYr2: 0.5,
  bcUp: 0.22, bcCap: 0.075, bcLTC: 0.65, bcRampYr2: 0.4,
};

// ---- Phase 1 engine: derive everything from the inputs ----------------------
function calcP1(over = {}) {
  const i = { ...IN, ...over };
  const netAc = i.grossAc - i.consAc - i.golfAc - i.critAc;
  const lots = over.lots != null ? over.lots : netAc * i.density;
  const rev = lots * i.entitledPx;
  const softBase = i.entitlement + i.civilEng + i.environmental + i.legalSurvey + i.projMgmt;
  const mkt = rev * i.mktPct;
  const cont = softBase * i.contPct;
  const fin = i.fin;
  const totCost = softBase + mkt + cont + fin;           // we only fund soft costs
  const profit = rev - totCost;
  const propShare = profit * i.split;                    // property / H entities
  const devShare = profit * (1 - i.split);               // development company
  const joshShare = propShare * i.joshOfProperty;        // ~25% of profit
  const jalTake = devShare * i.jalShareOfDevCo;          // JAL's sweat-equity share of the dev co
  const jalOfProfit = jalTake / profit;                  // ≈ 17% of total profit
  const retainerTotal = i.retainerMo * i.engMonths;      // $15K/mo over the engagement
  const jalAllIn = jalTake + retainerTotal;              // carry + retainer (illustrative all-in)
  const roiCost = profit / totCost;
  const margin = profit / rev;
  // annual project cash flow, derived from totals × phasing
  const revArr = i.revPhase.map((p) => rev * p);
  const costArr = i.costPhase.map((p) => -totCost * p);
  const netCF = revArr.map((v, t) => v + costArr[t]);
  const projIRR = irr(netCF);
  // finished-lot comparison (build the horizontal ourselves; land contributed)
  const finRev = lots * i.finishedPx;
  const finHoriz = lots * i.horizPerLot;
  const finCost = finHoriz + softBase + finRev * i.mktPct + finHoriz * 0.07 + i.finFin;
  const finProfit = finRev - finCost;
  const finRoiCost = finProfit / finCost;
  return { netAc, lots, rev, softBase, mkt, cont, fin, totCost, profit, propShare,
    devShare, joshShare, jalTake, jalOfProfit, retainerTotal, jalAllIn, roiCost, margin, revArr, costArr, netCF, projIRR,
    finRev, finHoriz, finCost, finProfit, finRoiCost };
}

// sensitivity helpers reuse the engines (no duplicated constants)
const p1Profit = (entitledPx, lots) => calcP1({ entitledPx, lots }).profit;

const P1 = calcP1();

// The deck (generate.js) imports this engine, so deck and workbook can't drift.
module.exports = { IN, irr, calcP1, p1Profit, P1 };

// ============================================================================
// WORKBOOK  (rendered only when run directly: `node model.js` / npm run model)
// ============================================================================
if (require.main === module) {
const NAVY = "FF0B163C", AUB = "FF3A243A", CREAM = "FFF4F2ED", WHITE = "FFFFFFFF";
const wb = new ExcelJS.Workbook();
wb.creator = "JAL Strategies";
wb.title = "Queenstown Harbor — Entitled-Lot (Capital-Light) Model";
wb.calcProperties.fullCalcOnLoad = true;
const ws = wb.addWorksheet("Entitled-Lot Model", {
  views: [{ showGridLines: false }],
  properties: { defaultRowHeight: 16 },
});
ws.columns = [{ width: 42 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }];

const COLS = ["B", "C", "D", "E", "F"];
let r = 0;
const row = () => ws.getRow(++r);
const money = '$#,##0';
const pct = '0.0%';
const mult = '0.00"x"';

function band(text, sub) {
  const a = row();
  a.getCell(1).value = text;
  a.getCell(1).font = { name: "Calibri", bold: true, size: 16, color: { argb: WHITE } };
  for (let c = 1; c <= 6; c++) a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  a.height = 24;
  if (sub) {
    const b = row();
    b.getCell(1).value = sub;
    b.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: AUB } };
    ws.mergeCells(r, 1, r, 6);
  }
}
function header(text) {
  r++;
  const a = row();
  a.getCell(1).value = text;
  for (let c = 1; c <= 6; c++) {
    a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AUB } };
    a.getCell(c).font = { name: "Calibri", bold: true, size: 10, color: { argb: WHITE } };
  }
  a.getCell(1).alignment = { indent: 1 };
  return r;
}
function line(label, value, note, fmt, opts = {}) {
  const a = row();
  a.getCell(1).value = label;
  a.getCell(1).font = { name: "Calibri", size: 10, bold: !!opts.bold, color: { argb: opts.bold ? NAVY : "FF333333" } };
  a.getCell(1).alignment = { indent: 1 };
  const vc = a.getCell(2);
  vc.value = value;
  if (fmt) vc.numFmt = fmt;
  vc.font = { name: "Calibri", size: 10, bold: opts.bold !== false, color: { argb: opts.accent ? AUB : NAVY } };
  vc.alignment = { horizontal: "right" };
  if (note) {
    const nc = a.getCell(3);
    ws.mergeCells(r, 3, r, 6);
    nc.value = note;
    nc.font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } };
  }
  if (opts.fill) for (let c = 1; c <= 2; c++) a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };
  return r;
}
function yrHead() {
  const a = row();
  a.getCell(1).value = "$ by year";
  a.getCell(1).font = { name: "Calibri", bold: true, size: 9, color: { argb: "FF6B5A6B" } };
  a.getCell(1).alignment = { indent: 1 };
  ["Yr 0", "Yr 1", "Yr 2", "Yr 3", "Yr 4"].forEach((y, i) => {
    const c = a.getCell(2 + i); c.value = y;
    c.font = { name: "Calibri", bold: true, size: 9, color: { argb: "FF6B5A6B" } };
    c.alignment = { horizontal: "right" };
  });
  return r;
}
function hrow(label, cells, fmt, opts = {}) {
  const a = row();
  a.getCell(1).value = label;
  a.getCell(1).font = { name: "Calibri", size: 10, bold: !!opts.bold, color: { argb: opts.bold ? NAVY : "FF333333" } };
  a.getCell(1).alignment = { indent: 1 };
  cells.forEach((v, idx) => {
    const c = a.getCell(2 + idx);
    c.value = v;
    if (fmt) c.numFmt = fmt;
    c.font = { name: "Calibri", size: 10, bold: !!opts.bold, color: { argb: opts.accent ? AUB : NAVY } };
    c.alignment = { horizontal: "right" };
    if (opts.fill) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };
  });
  return r;
}
function phased(scalarRef, phaseRow, results, sign = 1) {
  return COLS.map((col, t) => ({ formula: `${sign < 0 ? "-" : ""}${scalarRef}*${col}${phaseRow}`, result: results[t] }));
}
function sumRows(rowA, rowB, results) {
  return COLS.map((col, t) => ({ formula: `${col}${rowA}+${col}${rowB}`, result: results[t] }));
}

band("Queenstown Harbor — Entitled-Lot (Capital-Light) Development Model",
  "Entitle the land and sell ENTITLED lots; the builder funds the horizontal. Costs repaid, then a simple 50/50.  ·  Dynamic — live formulas off INPUTS.  ·  ILLUSTRATIVE.");

// ---- INPUTS ----
header("INPUTS  (edit these — everything below recalculates)");
const grossAc = line("Gross Queenstown land (acres)", IN.grossAc, "owner's figure — to confirm by survey", "#,##0");
const consAc  = line("Less: permanent conservation (acres)", IN.consAc, "recorded easement (qhgolf.com)", "#,##0");
const golfAc  = line("Less: golf, range, water & lodging (acres)", IN.golfAc, "two 18s + 9-ac range + clubhouse/cottages (est.)", "#,##0");
const critAc  = line("Less: Critical Area & wetlands (acres)", IN.critAc, "Chesapeake Bay 1,000-ft zone (est.)", "#,##0");
const density = line("Lot density (lots / acre)", IN.density, "blended; Queen Anne's Co. zoning to confirm", "0.0");
const entPx   = line("Entitled lot price ($ / lot)", IN.entitledPx, "to a builder — ~finished $150K less the builder's ~$80K to finish", money);
const cEnt    = line("Entitlement / approvals ($)", IN.entitlement, "planning, zoning, PUD", money);
const cCiv    = line("Civil engineering ($)", IN.civilEng, "construction plans", money);
const cEnv    = line("Environmental studies ($)", IN.environmental, "wetlands, Critical Area", money);
const cLeg    = line("Legal / survey / title ($)", IN.legalSurvey, "", money);
const cPM     = line("Project management ($)", IN.projMgmt, "consultants over the entitlement", money);
const mktPct  = line("Marketing & brokerage (% of revenue)", IN.mktPct, "lot sales", pct);
const contPct = line("Contingency (% of soft cost)", IN.contPct, "on the soft costs", pct);
const fin     = line("Financing / carry ($)", IN.fin, "small bank line", money);
const splitIn = line("Profit split — each side", IN.split, "50/50 after costs (no pref, no promote)", pct);
const joshIn  = line("Josh ownership of property entity", IN.joshOfProperty, "→ Josh nets ~25% of profit", pct);
const jalIn   = line("JAL share of the dev co", IN.jalShareOfDevCo, "sweat equity — 30% of the dev-co 50% (~15% of profit)", pct);

const B = (i) => `B${i}`;

// ---- LAND -> LOTS ----
header("DEVELOPABLE LAND  →  ENTITLED LOTS");
const netAc = line("Net developable land (acres)",
  { formula: `${B(grossAc)}-${B(consAc)}-${B(golfAc)}-${B(critAc)}`, result: P1.netAc },
  "gross − conservation − golf − Critical Area", "#,##0", { bold: true });
const lots = line("Entitled lots",
  { formula: `${B(netAc)}*${B(density)}`, result: P1.lots },
  "net developable acres × density", "#,##0", { bold: true, accent: true });

// ---- REVENUE ----
header("REVENUE");
const rev = line("Gross entitled-lot revenue",
  { formula: `${B(lots)}*${B(entPx)}`, result: P1.rev },
  "entitled lots × entitled-lot price", money, { bold: true, accent: true });

// ---- DEVELOPMENT COST (SOFT ONLY) ----
header("DEVELOPMENT COST  (SOFT ONLY — builder funds the horizontal)");
const uMkt = line("Marketing & brokerage", { formula: `${B(rev)}*${B(mktPct)}`, result: P1.mkt }, "% of revenue", money);
const softBase = line("Soft costs (entitle the land)", { formula: `${B(cEnt)}+${B(cCiv)}+${B(cEnv)}+${B(cLeg)}+${B(cPM)}`, result: P1.softBase }, "entitlement + civil + env + legal + PM", money);
const uCont = line("Contingency", { formula: `${B(softBase)}*${B(contPct)}`, result: P1.cont }, "% of soft cost", money);
const uFin = line("Financing / carry", { formula: `${B(fin)}`, result: P1.fin }, "small bank line", money);
const totCost = line("TOTAL DEVELOPMENT COST", { formula: `${B(softBase)}+${B(uMkt)}+${B(uCont)}+${B(uFin)}`, result: P1.totCost }, "what we fund (land contributed at ~zero basis)", money, { bold: true, accent: true, fill: true });

// ---- PROFIT & 50/50 SPLIT ----
header("PROFIT  &  THE 50/50 SPLIT  (costs repaid first, then a straight deal)");
const profit = line("Net profit", { formula: `${B(rev)}-${B(totCost)}`, result: P1.profit }, "revenue − development cost", money, { bold: true, accent: true });
line("Profit margin (on revenue)", { formula: `${B(profit)}/${B(rev)}`, result: P1.margin }, "near-zero land basis", pct);
line("Return on cost", { formula: `${B(profit)}/${B(totCost)}`, result: P1.roiCost }, "profit / development cost", mult, { bold: true, accent: true });
const propShare = line("→ Property / H entities (50%)", { formula: `${B(profit)}*${B(splitIn)}`, result: P1.propShare }, "returns capital to Josh's investors", money, { bold: true });
line("    of which Josh (~25% of profit)", { formula: `${B(propShare)}*${B(joshIn)}`, result: P1.joshShare }, "Josh owns 50% of the property entity", money);
const devShareRow = line("→ Development company (50%)", { formula: `${B(profit)}*(1-${B(splitIn)})`, result: P1.devShare }, "Accountable Equity / Capital H6 · Bob · JAL · partners", money, { bold: true, accent: true });
const jalTakeRow = line("    of which JAL (30% of dev co)", { formula: `${B(devShareRow)}*${B(jalIn)}`, result: P1.jalTake }, "the carry — sweat equity; separate from any placement fee or co-invest", money, { bold: true });
line("    + JAL retainer ($15K/mo × ~36 mo)", { formula: `${P1.retainerTotal}`, result: P1.retainerTotal }, "a floor during the engagement (a deal cost)", money);
line("    = JAL all-in (illustrative)", { formula: `${B(jalTakeRow)}+${B(jalTakeRow + 1)}`, result: P1.jalAllIn }, "carry + retainer on the Queenstown base case", money, { bold: true, accent: true });

// ---- CASH FLOW ----
header("PROJECT CASH FLOW  (revenue × phasing − cost × phasing)");
yrHead();
const revPhaseRow  = hrow("Revenue phasing (entitled-lot sales)", IN.revPhase, pct);
const costPhaseRow = hrow("Cost phasing (entitlement soft costs)", IN.costPhase, pct);
const p1RevRow  = hrow("Entitled-lot revenue", phased(`$B$${rev}`, revPhaseRow, P1.revArr), money);
const p1CostRow = hrow("Development cost", phased(`$B$${totCost}`, costPhaseRow, P1.costArr, -1), money);
const p1NetRow  = hrow("Net project cash flow", sumRows(p1RevRow, p1CostRow, P1.netCF), money, { bold: true, accent: true });
line("Project IRR (unlevered)", { formula: `IRR(B${p1NetRow}:F${p1NetRow})`, result: P1.projIRR },
  "high — minimal capital on a near-zero-basis land", pct, { bold: true, accent: true });

// ---- CAPITAL-LIGHT vs FINISHED-LOT ----
header("WHY CAPITAL-LIGHT  (entitled vs. finishing the lots ourselves)");
const ch = row();
["", "Entitled lots", "Finished lots"].forEach((t, i) => { const c = ch.getCell(1 + i); c.value = t; c.font = { name: "Calibri", bold: true, size: 9, color: { argb: i ? NAVY : "FF6B5A6B" } }; c.alignment = { horizontal: i ? "right" : "left", indent: i ? 0 : 1 }; });
const cmp = (label, ent, finv, fmt) => {
  const a = row();
  a.getCell(1).value = label; a.getCell(1).font = { name: "Calibri", size: 10, color: { argb: "FF333333" } }; a.getCell(1).alignment = { indent: 1 };
  const e = a.getCell(2); e.value = ent; e.numFmt = fmt; e.font = { name: "Calibri", size: 10, bold: true, color: { argb: AUB } }; e.alignment = { horizontal: "right" };
  const f = a.getCell(3); f.value = finv; f.numFmt = fmt; f.font = { name: "Calibri", size: 10, color: { argb: NAVY } }; f.alignment = { horizontal: "right" };
};
cmp("Revenue", P1.rev, P1.finRev, money);
cmp("Development cost (capital deployed)", P1.totCost, P1.finCost, money);
cmp("Net profit", P1.profit, P1.finProfit, money);
cmp("Return on cost", P1.roiCost, P1.finRoiCost, mult);
{
  const a = row(); ws.mergeCells(r, 1, r, 6);
  a.getCell(1).value = "Finishing the lots adds ~$" + ((P1.finProfit - P1.profit) / 1e6).toFixed(1) + "M of profit but needs ~" + (P1.finCost / P1.totCost).toFixed(1) + "× the capital and all the horizontal execution + absorption risk. Capital-light wins on return-on-cost and risk — and it's what the sponsor wants.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } };
}

// ---- SENSITIVITY ----
header("PROFIT SENSITIVITY  ($M — entitled lot price  ×  lot yield)");
const sPrices = [55000, 70000, 85000];
const sLots = [240, 280, 320];
const shr = row(); shr.getCell(1).value = "Entitled $/lot  /  lots"; shr.getCell(1).font = { bold: true, size: 9, color: { argb: WHITE } }; shr.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
sLots.forEach((l, i) => { const c = shr.getCell(2 + i); c.value = l + " lots"; c.font = { bold: true, size: 10, color: { argb: WHITE } }; c.alignment = { horizontal: "center" }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
sPrices.forEach((p) => {
  const rr = row(); rr.getCell(1).value = "$" + (p / 1000) + "K / lot"; rr.getCell(1).font = { bold: true, size: 10, color: { argb: NAVY } }; rr.getCell(1).alignment = { indent: 1 };
  sLots.forEach((l, i) => { const v = p1Profit(p, l) / 1e6; const c = rr.getCell(2 + i); c.value = v; c.numFmt = '$#,##0.0,"M"'.replace(",,", ""); c.numFmt = '#,##0.0'; const base = (p === IN.entitledPx && l === 280); c.font = { size: 10, bold: base, color: { argb: base ? AUB : NAVY } }; c.alignment = { horizontal: "center" }; if (base) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } }; });
});
{
  const a = row(); ws.mergeCells(r, 1, r, 6);
  a.getCell(1).value = "Net project profit ($M) across entitled-lot price × lot yield. Base = $70K × 280 lots (highlighted). 50% flows to the property (Josh's investor liquidity), 50% to the development company.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } };
}

r++;
{
  const a = row(); ws.mergeCells(r, 1, r, 6);
  a.getCell(1).value = "Illustrative and for discussion only. Capital-light entitled-lot basis: we fund only the soft costs to entitle the land and sell entitled lots; the builder funds the horizontal. Costs repaid first, then a straight 50/50 (no pref, no promote). The entitled-lot price is the key assumption — confirm against builder bids in diligence.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 8.5, color: { argb: "FF6B5A6B" } };
  a.getCell(1).alignment = { wrapText: true, vertical: "top" };
  a.height = 42;
}

wb.xlsx.writeFile("JAL_Queenstown_Harbor_Model.xlsx")
  .then(() => console.log(
    "Wrote model  P1 (entitled): rev $" + (P1.rev / 1e6).toFixed(1) + "M  cost $" + (P1.totCost / 1e6).toFixed(1) +
    "M  profit $" + (P1.profit / 1e6).toFixed(1) + "M  (" + P1.roiCost.toFixed(1) + "x on cost, " + (P1.margin * 100).toFixed(0) + "% margin)  IRR ~" + (P1.projIRR * 100).toFixed(0) + "%" +
    "\n             50/50: $" + (P1.propShare / 1e6).toFixed(1) + "M property (Josh ~$" + (P1.joshShare / 1e6).toFixed(1) + "M) / $" + (P1.devShare / 1e6).toFixed(1) + "M dev co" +
    "\n             P1 netCF($M): [" + P1.netCF.map((v) => (v / 1e6).toFixed(1)).join(", ") + "]" +
    "\n             vs finished: profit $" + (P1.finProfit / 1e6).toFixed(1) + "M on $" + (P1.finCost / 1e6).toFixed(1) + "M cost (" + P1.finRoiCost.toFixed(1) + "x)"))
  .catch((e) => { console.error(e); process.exit(1); });
} // end require.main gate
