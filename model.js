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

// ---- Phase 2 engine (hospitality; hotel is a required component) ------------
function calcP2(over = {}) {
  const i = { ...IN, ...over };
  const hCost = i.hKeys * i.hCostKey;
  const hRoom = i.hKeys * 365 * i.hADR * i.hOcc;
  const hRev = hRoom * i.hMult;
  const hNOI = hRev * i.hMgn;
  const rCost = i.rSF * i.rCostSF;
  const rSales = i.rSF * i.rSalesSF;
  const rNOI = rSales * i.rMgn;
  const pCost = hCost + rCost + i.pLand + i.pSoft;
  const pNOI = hNOI + rNOI;
  const pYoC = pNOI / pCost;
  const pLoan = pCost * i.pLTC;
  const pEq = pCost - pLoan;
  const stabLevCF = pNOI - pLoan * i.pRate;
  const pCoC = stabLevCF / pEq;
  const pXNOI = pNOI * Math.pow(1 + i.pGrow, 4);
  const pXVal = pXNOI / i.pCap;
  const pXEq = pXVal - pLoan;
  const cf2 = [
    -pEq * i.eqSplit, -pEq * (1 - i.eqSplit),
    stabLevCF * i.rampYr2, stabLevCF,
    stabLevCF * Math.pow(1 + i.pGrow, 1), stabLevCF * Math.pow(1 + i.pGrow, 2),
    stabLevCF * Math.pow(1 + i.pGrow, 3), stabLevCF * Math.pow(1 + i.pGrow, 4) + pXEq,
  ];
  const irr2 = irr(cf2), em2 = cf2.slice(2).reduce((a, b) => a + b, 0) / -(cf2[0] + cf2[1]);
  const bcNOI = pNOI * (1 + i.bcUp);
  const bcVal = bcNOI / i.bcCap;
  const bcLoan = pCost * i.bcLTC;
  const bcEq = pCost - bcLoan;
  const bcStabLevCF = bcNOI - bcLoan * i.pRate;
  const bcYoC = bcNOI / pCost;
  const bcXEq = bcVal - bcLoan;
  const bcf = [
    -bcEq * i.eqSplit, -bcEq * (1 - i.eqSplit),
    bcStabLevCF * i.bcRampYr2, bcStabLevCF, bcStabLevCF + bcXEq,
  ];
  const bcIRR = irr(bcf), bcEM = bcf.slice(2).reduce((a, b) => a + b, 0) / -(bcf[0] + bcf[1]);
  return { hCost, hRoom, hRev, hNOI, rCost, rSales, rNOI, pCost, pNOI, pYoC, pLoan,
    pEq, stabLevCF, pCoC, pXNOI, pXVal, pXEq, cf2, irr2, em2,
    bcNOI, bcVal, bcLoan, bcEq, bcStabLevCF, bcYoC, bcXEq, bcf, bcIRR, bcEM };
}

// sensitivity helpers reuse the engines (no duplicated constants)
const p1Profit = (entitledPx, lots) => calcP1({ entitledPx, lots }).profit;
const p2IRR = (bcUp, bcCap, bcLTC = IN.bcLTC) => calcP2({ bcUp, bcCap, bcLTC }).bcIRR;

const P1 = calcP1();
const P2 = calcP2();

// ============================================================================
// WORKBOOK
// ============================================================================
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

// ============== PHASE 2 — HOSPITALITY (HOTEL + RESTAURANTS) ==============
const p2 = wb.addWorksheet("Phase 2 — Hospitality", { views: [{ showGridLines: false }] });
p2.columns = [{ width: 42 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }];
let q = 0;
const qrow = () => p2.getRow(++q);
const Q = (i) => `B${i}`;
function qband(t, sub) {
  const a = qrow(); a.getCell(1).value = t; a.getCell(1).font = { name: "Calibri", bold: true, size: 16, color: { argb: WHITE } };
  for (let c = 1; c <= 6; c++) a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; a.height = 24;
  if (sub) { const b = qrow(); b.getCell(1).value = sub; b.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: AUB } }; p2.mergeCells(q, 1, q, 6); }
}
function qheader(t) {
  q++; const a = qrow(); a.getCell(1).value = t;
  for (let c = 1; c <= 6; c++) { a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AUB } }; a.getCell(c).font = { name: "Calibri", bold: true, size: 10, color: { argb: WHITE } }; }
  a.getCell(1).alignment = { indent: 1 }; return q;
}
function qline(label, value, note, fmt, opts = {}) {
  const a = qrow(); a.getCell(1).value = label; a.getCell(1).font = { name: "Calibri", size: 10, bold: !!opts.bold, color: { argb: opts.bold ? NAVY : "FF333333" } }; a.getCell(1).alignment = { indent: 1 };
  const vc = a.getCell(2); vc.value = value; if (fmt) vc.numFmt = fmt; vc.font = { name: "Calibri", size: 10, bold: opts.bold !== false, color: { argb: opts.accent ? AUB : NAVY } }; vc.alignment = { horizontal: "right" };
  if (note) { const nc = a.getCell(3); p2.mergeCells(q, 3, q, 6); nc.value = note; nc.font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } }; }
  if (opts.fill) for (let c = 1; c <= 2; c++) a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };
  return q;
}
function qcf(label, items) {
  const start = q + 1;
  items.forEach((it, i) => {
    const a = qrow(); a.getCell(1).value = label + " — Year " + i; a.getCell(1).alignment = { indent: 1 }; a.getCell(1).font = { name: "Calibri", size: 10, color: { argb: "FF333333" } };
    const c = a.getCell(2); c.value = it; c.numFmt = money; c.font = { name: "Calibri", size: 10, color: { argb: NAVY } }; c.alignment = { horizontal: "right" };
  });
  return { start, end: q };
}

qband("Queenstown Harbor — Phase 2: Hospitality (Hotel + Restaurants)",
  "Hotel is a REQUIRED component (per Bob). VIVÂMEE-led resort; JAL leads capital formation.  ·  Dynamic.  ·  ILLUSTRATIVE.");

qheader("INPUTS  (edit these — everything below recalculates)");
const hKeys = qline("Hotel — keys", IN.hKeys, "boutique resort lodge", "#,##0");
const hCostKey = qline("Hotel — cost / key ($)", IN.hCostKey, "HVS 2025: select ~$223K, full-service ~$409K", money);
const hADR = qline("Hotel — ADR ($)", IN.hADR, "upscale Eastern Shore resort (weddings / golf)", money);
const hOcc = qline("Hotel — stabilized occupancy", IN.hOcc, "leisure / resort", pct);
const hMult = qline("Hotel — total-revenue multiple (x rooms)", IN.hMult, "+ F&B, banquets, spa", "0.00");
const hMgn = qline("Hotel — NOI margin", IN.hMgn, "stabilized", pct);
const rSF = qline("Restaurants — GLA (SF)", IN.rSF, "a suite of destination restaurants", "#,##0");
const rCostSF = qline("Restaurants — build cost / SF ($)", IN.rCostSF, "full-service ~$555/SF", money);
const rSalesSF = qline("Restaurants — sales / SF ($)", IN.rSalesSF, "upscale destination dining", money);
const rMgn = qline("Restaurants — NOI (% of sales)", IN.rMgn, "rent / operating contribution", pct);
const pLand = qline("Land — hospitality parcel ($)", IN.pLand, "contributed by Capital H6", money);
const pSoft = qline("Soft / FF&E / pre-opening ($)", IN.pSoft, "", money);
const pCap = qline("Exit cap rate (hold case)", IN.pCap, "hotel cap ~8% (2025)", pct);
const pLTC = qline("Construction loan (% of cost)", IN.pLTC, "", pct);
const pRate = qline("Loan interest rate", IN.pRate, "", pct);
const pGrow = qline("NOI growth / yr", IN.pGrow, "", pct);
const eqSplit = qline("Equity drawn — Year 0 (balance Yr 1)", IN.eqSplit, "construction equity timing", pct);
const rampYr2 = qline("First operating year ramp (hold)", IN.rampYr2, "Yr 2 opens partway through", pct);

qheader("HOTEL  (required component)");
const hCost = qline("Hotel development cost", { formula: `${Q(hKeys)}*${Q(hCostKey)}`, result: P2.hCost }, "keys × cost/key", money);
const hRoom = qline("Room revenue", { formula: `${Q(hKeys)}*365*${Q(hADR)}*${Q(hOcc)}`, result: P2.hRoom }, "keys × 365 × ADR × occ", money);
const hRev = qline("Total hotel revenue", { formula: `${Q(hRoom)}*${Q(hMult)}`, result: P2.hRev }, "× revenue multiple", money);
const hNOI = qline("Hotel NOI (stabilized)", { formula: `${Q(hRev)}*${Q(hMgn)}`, result: P2.hNOI }, "× NOI margin", money, { bold: true, accent: true });

qheader("RESTAURANTS  (a suite of nice restaurants)");
const rCost = qline("Restaurant development cost", { formula: `${Q(rSF)}*${Q(rCostSF)}`, result: P2.rCost }, "SF × cost/SF", money);
const rSales = qline("Restaurant sales", { formula: `${Q(rSF)}*${Q(rSalesSF)}`, result: P2.rSales }, "SF × sales/SF", money);
const rNOI = qline("Restaurant NOI (to owner)", { formula: `${Q(rSales)}*${Q(rMgn)}`, result: P2.rNOI }, "rent / operating contribution", money, { bold: true, accent: true });

qheader("TOTAL PROJECT  ->  STABILIZED");
const pCost = qline("Total project cost", { formula: `${Q(hCost)}+${Q(rCost)}+${Q(pLand)}+${Q(pSoft)}`, result: P2.pCost }, "hotel + restaurants + land + soft", money, { bold: true, accent: true, fill: true });
const pNOI = qline("Stabilized NOI", { formula: `${Q(hNOI)}+${Q(rNOI)}`, result: P2.pNOI }, "hotel + restaurant NOI", money, { bold: true, accent: true, fill: true });
const pYoC = qline("Yield on cost", { formula: `${Q(pNOI)}/${Q(pCost)}`, result: P2.pYoC }, "NOI / cost", pct, { bold: true });
const pLoan = qline("Construction loan", { formula: `${Q(pCost)}*${Q(pLTC)}`, result: P2.pLoan }, "% LTC", money);
const pEq = qline("Equity (land + cash)", { formula: `${Q(pCost)}-${Q(pLoan)}`, result: P2.pEq }, "cost − loan", money);
const stabLev = qline("Stabilized levered cash flow", { formula: `${Q(pNOI)}-${Q(pLoan)}*${Q(pRate)}`, result: P2.stabLevCF }, "NOI − loan interest", money, { bold: true });
const pCoC = qline("Cash-on-cash (stabilized, levered)", { formula: `${Q(stabLev)}/${Q(pEq)}`, result: P2.pCoC }, "stabilized levered CF / equity", pct);

qheader("EXIT  (hold case — sell at exit cap)");
const pXNOI = qline("Exit-year NOI (+3%/yr × 4)", { formula: `${Q(pNOI)}*(1+${Q(pGrow)})^4`, result: P2.pXNOI }, "stabilized × growth", money);
const pXVal = qline("Exit value (@ exit cap)", { formula: `${Q(pXNOI)}/${Q(pCap)}`, result: P2.pXVal }, "exit-year NOI / cap", money, { bold: true });
const pXEq = qline("Exit equity (value − loan)", { formula: `${Q(pXVal)}-${Q(pLoan)}`, result: P2.pXEq }, "net of loan", money);

qheader("HOLD-CASE EQUITY CASH FLOW  (Yr 0–7, levered $) & RETURNS");
const cf2items = [
  { formula: `-${Q(pEq)}*${Q(eqSplit)}`, result: P2.cf2[0] },
  { formula: `-${Q(pEq)}*(1-${Q(eqSplit)})`, result: P2.cf2[1] },
  { formula: `${Q(stabLev)}*${Q(rampYr2)}`, result: P2.cf2[2] },
  { formula: `${Q(stabLev)}`, result: P2.cf2[3] },
  { formula: `${Q(stabLev)}*(1+${Q(pGrow)})`, result: P2.cf2[4] },
  { formula: `${Q(stabLev)}*(1+${Q(pGrow)})^2`, result: P2.cf2[5] },
  { formula: `${Q(stabLev)}*(1+${Q(pGrow)})^3`, result: P2.cf2[6] },
  { formula: `${Q(stabLev)}*(1+${Q(pGrow)})^4+${Q(pXEq)}`, result: P2.cf2[7] },
];
const cf2blk = qcf("Equity CF", cf2items);
qline("Equity multiple (7-yr hold)", { formula: `SUM(B${cf2blk.start + 2}:B${cf2blk.end})/-(B${cf2blk.start}+B${cf2blk.start + 1})`, result: P2.em2 }, "distributions / equity", mult, { bold: true, accent: true });
qline("Project IRR (levered, 7-yr hold)", { formula: `IRR(B${cf2blk.start}:B${cf2blk.end})`, result: P2.irr2 }, "income / hold — develops to ~cost", pct, { bold: true, accent: true });

qheader("BUILD-TO-CORE CASE  (recap / sell at stabilization — path to mid–high 20s IRR)");
const bcUp = qline("NOI uplift — premium execution", IN.bcUp, "VIVÂMEE experiential: ADR ~$275, occ ~63%, strong F&B / events", pct);
const bcNOI = qline("Premium stabilized NOI", { formula: `${Q(pNOI)}*(1+${Q(bcUp)})`, result: P2.bcNOI }, "base NOI × (1 + uplift)", money, { bold: true, accent: true });
const bcCap = qline("Exit cap (build-to-core)", IN.bcCap, "trophy / experiential resort", pct);
const bcLTC = qline("Leverage (% of cost)", IN.bcLTC, "construction loan", pct);
const bcRamp = qline("First operating year ramp (b-t-c)", IN.bcRampYr2, "Yr 2 opens partway through", pct);
const bcYoC = qline("Yield on cost (premium)", { formula: `${Q(bcNOI)}/${Q(pCost)}`, result: P2.bcYoC }, "premium NOI / cost — vs 7.5% exit = the spread", pct, { bold: true });
const bcVal = qline("Recap / sale value (~Yr 4)", { formula: `${Q(bcNOI)}/${Q(bcCap)}`, result: P2.bcVal }, "premium NOI / exit cap", money, { bold: true });
const bcLoan = qline("Loan", { formula: `${Q(pCost)}*${Q(bcLTC)}`, result: P2.bcLoan }, "% of cost", money);
const bcEq = qline("Equity", { formula: `${Q(pCost)}-${Q(bcLoan)}`, result: P2.bcEq }, "cost − loan", money);
const bcStabLev = qline("Premium stabilized levered cash flow", { formula: `${Q(bcNOI)}-${Q(bcLoan)}*${Q(pRate)}`, result: P2.bcStabLevCF }, "premium NOI − loan interest", money);
const bcXEq = qline("Exit equity (value − loan)", { formula: `${Q(bcVal)}-${Q(bcLoan)}`, result: P2.bcXEq }, "at recap / sale", money);

qheader("BUILD-TO-CORE EQUITY CASH FLOW  (Yr 0–4, levered $) & RETURNS");
const bcfItems = [
  { formula: `-${Q(bcEq)}*${Q(eqSplit)}`, result: P2.bcf[0] },
  { formula: `-${Q(bcEq)}*(1-${Q(eqSplit)})`, result: P2.bcf[1] },
  { formula: `${Q(bcStabLev)}*${Q(bcRamp)}`, result: P2.bcf[2] },
  { formula: `${Q(bcStabLev)}`, result: P2.bcf[3] },
  { formula: `${Q(bcStabLev)}+${Q(bcXEq)}`, result: P2.bcf[4] },
];
const bcfblk = qcf("Equity CF", bcfItems);
qline("Equity multiple (build-to-core)", { formula: `SUM(B${bcfblk.start + 2}:B${bcfblk.end})/-(B${bcfblk.start}+B${bcfblk.start + 1})`, result: P2.bcEM }, "distributions / equity, ~4-yr", mult, { bold: true, accent: true });
qline("Project IRR (build-to-core, ~4-yr)", { formula: `IRR(B${bcfblk.start}:B${bcfblk.end})`, result: P2.bcIRR }, "recap / sell at stabilization", pct, { bold: true, accent: true });

qheader("BUILD-TO-CORE IRR SENSITIVITY  (stabilized NOI  ×  exit cap)");
const p2Up = [[0.10, "$4.2M NOI"], [0.22, "$4.65M NOI (base)"], [0.35, "$5.15M NOI"]];
const p2Caps = [0.07, 0.075, 0.08, 0.085];
const ph = qrow(); ph.getCell(1).value = "Stabilized NOI  /  exit cap"; ph.getCell(1).font = { bold: true, size: 9, color: { argb: WHITE } }; ph.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
p2Caps.forEach((cp, i) => { const c = ph.getCell(2 + i); c.value = cp; c.numFmt = "0.0%"; c.font = { bold: true, size: 10, color: { argb: WHITE } }; c.alignment = { horizontal: "center" }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
p2Up.forEach(([u, lbl]) => {
  const rr = qrow(); rr.getCell(1).value = lbl; rr.getCell(1).font = { bold: true, size: 10, color: { argb: NAVY } }; rr.getCell(1).alignment = { indent: 1 };
  p2Caps.forEach((cp, i) => { const v = p2IRR(u, cp); const c = rr.getCell(2 + i); c.value = v; c.numFmt = "0.0%"; const base = (u === IN.bcUp && cp === IN.bcCap); c.font = { size: 10, bold: base, color: { argb: base ? AUB : NAVY } }; c.alignment = { horizontal: "center" }; if (base) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } }; });
});
{
  const a = qrow(); p2.mergeCells(a.number, 1, a.number, 6);
  a.getCell(1).value = "Build-to-core project IRR at 65% LTC, recap at stabilization (~Yr 4). Base = $4.65M NOI / 7.5% cap (highlighted). Each cell re-runs the engine.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } };
}

q++;
{
  const a = qrow(); p2.mergeCells(q, 1, q, 6);
  a.getCell(1).value = "Illustrative. The hotel is a required component (per Bob); at Renault the plan is +100–200 rooms. Hospitality develops to ~cost (income/hold ~12% IRR) unless you build to core and recap at stabilization (~28%). VIVÂMEE operates; JAL leads capital formation. The resort also lifts the lot values. Subject to confirmation.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 8.5, color: { argb: "FF6B5A6B" } };
  a.getCell(1).alignment = { wrapText: true, vertical: "top" }; a.height = 50;
}

wb.xlsx.writeFile("JAL_Queenstown_Harbor_Model.xlsx")
  .then(() => console.log(
    "Wrote model  P1 (entitled): rev $" + (P1.rev / 1e6).toFixed(1) + "M  cost $" + (P1.totCost / 1e6).toFixed(1) +
    "M  profit $" + (P1.profit / 1e6).toFixed(1) + "M  (" + P1.roiCost.toFixed(1) + "x on cost, " + (P1.margin * 100).toFixed(0) + "% margin)  IRR ~" + (P1.projIRR * 100).toFixed(0) + "%" +
    "\n             50/50: $" + (P1.propShare / 1e6).toFixed(1) + "M property (Josh ~$" + (P1.joshShare / 1e6).toFixed(1) + "M) / $" + (P1.devShare / 1e6).toFixed(1) + "M dev co" +
    "\n             P1 netCF($M): [" + P1.netCF.map((v) => (v / 1e6).toFixed(1)).join(", ") + "]" +
    "\n             vs finished: profit $" + (P1.finProfit / 1e6).toFixed(1) + "M on $" + (P1.finCost / 1e6).toFixed(1) + "M cost (" + P1.finRoiCost.toFixed(1) + "x)" +
    "\n             P2: hold ~" + (P2.irr2 * 100).toFixed(1) + "% / b-t-c ~" + (P2.bcIRR * 100).toFixed(1) + "% (" + P2.bcEM.toFixed(2) + "x)"))
  .catch((e) => { console.error(e); process.exit(1); });
