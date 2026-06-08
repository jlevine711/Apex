/**
 * Queenstown Harbor — Finished-Lot Development Model (backup to the deck).
 *
 * Finished-lot basis: the JV funds the horizontal (roads, sewer, utilities)
 * and sells FINISHED lots to homebuilders; builders just build the homes.
 *
 * DYNAMIC & CALCULATED. There are no hardcoded results: a single INPUTS object
 * (`IN`) feeds two pure calc engines (`calcP1`, `calcP2`) that derive every
 * figure — totals, cash flows, IRRs, the promote split and the sensitivity
 * grids. In the workbook, the INPUTS block plus the cash-flow timing vectors
 * are the only typed numbers; every other cell is a live Excel formula that
 * references them (the cached `result` shown before recalc is the engine's own
 * computed value, so spreadsheet and engine always agree). Edit any input —
 * in JS or in the sheet — and everything recalculates. All figures illustrative.
 *
 *   npm run model   ->   JAL_Queenstown_Harbor_Model.xlsx
 */
const ExcelJS = require("exceljs");

// ---- IRR (bisection) — caches the value Excel shows before it recalculates --
function irr(cfs) {
  const npv = (r) => cfs.reduce((a, c, i) => a + c / Math.pow(1 + r, i), 0);
  let lo = -0.9, hi = 5;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (npv(mid) > 0) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

// ---- promote waterfall: split a project equity cash flow into LP / GP -------
// Pari-passu return of capital + pref, then IRR-hurdle promote tiers.
function waterfall(eqCF, lpFrac, pref, tiers) {
  const n = eqCF.length;
  const lp = Array(n).fill(0), gp = Array(n).fill(0);
  for (let t = 0; t < n; t++) if (eqCF[t] < 0) { lp[t] = eqCF[t] * lpFrac; gp[t] = eqCF[t] * (1 - lpFrac); }
  for (let t = 0; t < n; t++) {
    let dist = eqCF[t] > 0 ? eqCF[t] : 0;
    if (dist <= 0) continue;
    const N = 2000, step = dist / N;
    for (let k = 0; k < N; k++) {
      const cur = irr(lp.slice(0, t + 1));                 // LP IRR achieved so far
      let s = tiers[tiers.length - 1].lp;                  // top tier by default
      if (cur < pref - 1e-6) s = lpFrac;                   // still earning pref → pari-passu
      else for (const tr of tiers) if (cur < tr.irr - 1e-6) { s = tr.lp; break; }
      lp[t] += step * s; gp[t] += step * (1 - s);
    }
  }
  return { lp, gp };
}

// ============================================================================
// INPUTS — the single source of truth.  Edit here (or in the sheet) and
// everything below recalculates.  Phasing vectors run Yr 0..4 and sum to 1.
// ============================================================================
const IN = {
  // ---- Phase 1 — finished-lot development ----
  grossAc: 700, consAc: 198, golfAc: 302, critAc: 60,
  density: 2.0, lotPx: 150000, horizPerLot: 60000, soft: 2800000,
  mktPct: 0.03, contPct: 0.07, fin: 1400000, landBas: 4000000,
  ltc: 0.60, pref: 0.08, lpFrac: 0.90,
  tiers: [{ irr: 0.15, lp: 0.80 }, { irr: 0.20, lp: 0.70 }, { irr: Infinity, lp: 0.60 }],
  // cash-flow timing (% by year, Yr 0..4)
  revPhase:  [0, 0, 8 / 42, 17 / 42, 17 / 42],   // finished-lot sales, Yr 2–4
  costPhase: [0.22, 0.38, 0.36, 0.03, 0.01],     // land/soft/horizontal front-loaded
  drawPhase: [0.607, 0.393, 0, 0, 0],            // equity drawn Yr 0–1
  distPhase: [0, 0, 0.088, 0.307, 0.605],        // equity distributions Yr 2–4

  // ---- Phase 2 — hospitality (hotel + restaurants) ----
  hKeys: 100, hCostKey: 325000, hADR: 250, hOcc: 0.62, hMult: 1.6, hMgn: 0.31,
  rSF: 16000, rCostSF: 550, rSalesSF: 700, rMgn: 0.09,
  pLand: 2000000, pSoft: 3700000, pCap: 0.08, pLTC: 0.60, pRate: 0.07, pGrow: 0.03,
  eqSplit: 0.5,        // equity drawn Yr 0 (balance Yr 1)
  rampYr2: 0.5,        // first operating year (hold case) runs at half
  // build-to-core
  bcUp: 0.22, bcCap: 0.075, bcLTC: 0.65, bcRampYr2: 0.4,
};

// ---- Phase 1 engine: derive everything from the inputs ----------------------
function calcP1(over = {}) {
  const i = { ...IN, ...over };
  const netAc = i.grossAc - i.consAc - i.golfAc - i.critAc;
  const lots = netAc * i.density;
  const rev = lots * i.lotPx;
  const uLand = i.landBas;
  const uHoriz = lots * i.horizPerLot;
  const uSoft = i.soft;
  const uMkt = rev * i.mktPct;
  const uCont = uHoriz * i.contPct;
  const uFin = i.fin;
  const totCost = uLand + uHoriz + uSoft + uMkt + uCont + uFin;
  const loan = totCost * i.ltc;
  const landEq = i.landBas;
  const cashEq = totCost - loan - landEq;
  const totEquity = landEq + cashEq;          // = totCost − loan
  const profit = rev - totCost;
  const dist = totEquity + profit;            // distributable to equity
  const em = dist / totEquity;
  // annual cash flows, derived from totals × phasing
  const revArr = i.revPhase.map((p) => rev * p);
  const costArr = i.costPhase.map((p) => -totCost * p);
  const unlevCF = revArr.map((v, t) => v + costArr[t]);
  const drawArr = i.drawPhase.map((p) => -totEquity * p);
  const distArr = i.distPhase.map((p) => dist * p);
  const levCF = drawArr.map((v, t) => v + distArr[t]);
  const unlevIRR = irr(unlevCF), levIRR = irr(levCF);
  const wf = waterfall(levCF, i.lpFrac, i.pref, i.tiers);
  const lpIRR = irr(wf.lp), gpIRR = irr(wf.gp), gpProfit = wf.gp.reduce((a, b) => a + b, 0);
  return { netAc, lots, rev, uLand, uHoriz, uSoft, uMkt, uCont, uFin, totCost, loan,
    landEq, cashEq, totEquity, profit, dist, em, revArr, costArr, unlevCF,
    drawArr, distArr, levCF, unlevIRR, levIRR, wf, lpIRR, gpIRR, gpProfit };
}

// ---- Phase 2 engine ---------------------------------------------------------
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
  const stabLevCF = pNOI - pLoan * i.pRate;       // stabilized levered cash flow
  const pCoC = stabLevCF / pEq;
  const pXNOI = pNOI * Math.pow(1 + i.pGrow, 4);
  const pXVal = pXNOI / i.pCap;
  const pXEq = pXVal - pLoan;
  // 7-year hold equity cash flow (Yr 0..7)
  const cf2 = [
    -pEq * i.eqSplit, -pEq * (1 - i.eqSplit),
    stabLevCF * i.rampYr2,
    stabLevCF,
    stabLevCF * Math.pow(1 + i.pGrow, 1),
    stabLevCF * Math.pow(1 + i.pGrow, 2),
    stabLevCF * Math.pow(1 + i.pGrow, 3),
    stabLevCF * Math.pow(1 + i.pGrow, 4) + pXEq,
  ];
  const irr2 = irr(cf2), em2 = cf2.slice(2).reduce((a, b) => a + b, 0) / -(cf2[0] + cf2[1]);
  // build-to-core (recap / sell at stabilization, ~Yr 4)
  const bcNOI = pNOI * (1 + i.bcUp);
  const bcVal = bcNOI / i.bcCap;
  const bcLoan = pCost * i.bcLTC;
  const bcEq = pCost - bcLoan;
  const bcStabLevCF = bcNOI - bcLoan * i.pRate;
  const bcYoC = bcNOI / pCost;
  const bcXEq = bcVal - bcLoan;
  const bcf = [
    -bcEq * i.eqSplit, -bcEq * (1 - i.eqSplit),
    bcStabLevCF * i.bcRampYr2,
    bcStabLevCF,
    bcStabLevCF + bcXEq,
  ];
  const bcIRR = irr(bcf), bcEM = bcf.slice(2).reduce((a, b) => a + b, 0) / -(bcf[0] + bcf[1]);
  return { hCost, hRoom, hRev, hNOI, rCost, rSales, rNOI, pCost, pNOI, pYoC, pLoan,
    pEq, stabLevCF, pCoC, pXNOI, pXVal, pXEq, cf2, irr2, em2,
    bcNOI, bcVal, bcLoan, bcEq, bcStabLevCF, bcYoC, bcXEq, bcf, bcIRR, bcEM };
}

// sensitivity helpers reuse the engines (no duplicated constants)
const p1LPIRR = (lotPx, ltc) => calcP1({ lotPx, ltc }).lpIRR;
const p2IRR = (bcUp, bcCap, bcLTC = IN.bcLTC) => calcP2({ bcUp, bcCap, bcLTC }).bcIRR;

const P1 = calcP1();
const P2 = calcP2();

// ============================================================================
// WORKBOOK
// ============================================================================
const NAVY = "FF0B163C", AUB = "FF3A243A", CREAM = "FFF4F2ED", WHITE = "FFFFFFFF";
const wb = new ExcelJS.Workbook();
wb.creator = "JAL Strategies";
wb.title = "Queenstown Harbor — Finished-Lot Development Model";
wb.calcProperties.fullCalcOnLoad = true; // recalc every formula when opened
const ws = wb.addWorksheet("Finished-Lot Model", {
  views: [{ showGridLines: false }],
  properties: { defaultRowHeight: 16 },
});
ws.columns = [{ width: 40 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }];

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
  r++; // spacer
  const a = row();
  a.getCell(1).value = text;
  for (let c = 1; c <= 6; c++) {
    a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AUB } };
    a.getCell(c).font = { name: "Calibri", bold: true, size: 10, color: { argb: WHITE } };
  }
  a.getCell(1).alignment = { indent: 1 };
  return r;
}
// label / value(or formula) / note ; returns the row index for referencing
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
// year-labelled header (Yr 0..4 across B:F)
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
// horizontal row of 5 cells (numbers or {formula,result}) across B:F
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
// build 5 formulas: scalarRef × phaseRow (optionally negated), with engine results
function phased(scalarRef, phaseRow, results, sign = 1) {
  return COLS.map((col, t) => ({
    formula: `${sign < 0 ? "-" : ""}${scalarRef}*${col}${phaseRow}`,
    result: results[t],
  }));
}
// build 5 formulas summing two rows column-by-column
function sumRows(rowA, rowB, results) {
  return COLS.map((col, t) => ({ formula: `${col}${rowA}+${col}${rowB}`, result: results[t] }));
}

band("Queenstown Harbor — Finished-Lot Development Model",
  "JV funds the horizontal (roads, sewer, utilities); homebuilders just build the homes.  ·  Dynamic — every cell below is a live formula off INPUTS.  ·  ILLUSTRATIVE.");

// ---- INPUTS ----
header("INPUTS  (edit these — everything below recalculates)");
const grossAc = line("Gross Queenstown land (acres)", IN.grossAc, "owner's figure — to confirm by survey", "#,##0");
const consAc  = line("Less: permanent conservation (acres)", IN.consAc, "recorded easement (qhgolf.com)", "#,##0");
const golfAc  = line("Less: golf, range, water & lodging (acres)", IN.golfAc, "two 18s + 9-ac range + clubhouse/cottages (est.)", "#,##0");
const critAc  = line("Less: Critical Area & wetlands (acres)", IN.critAc, "Chesapeake Bay 1,000-ft zone (est.)", "#,##0");
const density = line("Lot density (lots / acre)", IN.density, "blended; Queen Anne's Co. zoning to confirm", "0.0");
const lotPx   = line("Finished lot price ($ / lot)", IN.lotPx, "conservative — QAC lots avg ~$237K; new homes $600–705K", money);
const horiz   = line("Horizontal site dev cost ($ / lot)", IN.horizPerLot, "JV-funded; national benchmark $50–150K/acre", money);
const soft    = line("Soft costs ($)", IN.soft, "entitlement, civil, environmental, legal, mgmt", money);
const mktPct  = line("Marketing & brokerage (% of revenue)", IN.mktPct, "lot sales", pct);
const contPct = line("Contingency (% of hard cost)", IN.contPct, "on horizontal", pct);
const fin     = line("Financing / interest carry ($)", IN.fin, "development loan interest", money);
const landBas = line("Land basis — contributed ($)", IN.landBas, "allocated from $25M H6 purchase (~$29K/ac); appraisal to set", money);
const ltc     = line("Development loan (% of total cost)", IN.ltc, "land-development / A&D financing (60% LTC)", pct);
const pref    = line("Preferred return", IN.pref, "to equity, before promote", pct);

const B = (i) => `B${i}`;

// ---- LAND -> LOTS ----
header("DEVELOPABLE LAND  →  LOTS");
const netAc = line("Net developable land (acres)",
  { formula: `${B(grossAc)}-${B(consAc)}-${B(golfAc)}-${B(critAc)}`, result: P1.netAc },
  "gross − conservation − golf − Critical Area", "#,##0", { bold: true });
const lots = line("Finished lots",
  { formula: `${B(netAc)}*${B(density)}`, result: P1.lots },
  "net developable acres × density", "#,##0", { bold: true, accent: true });

// ---- REVENUE ----
header("REVENUE");
const rev = line("Gross lot revenue",
  { formula: `${B(lots)}*${B(lotPx)}`, result: P1.rev },
  "finished lots × finished lot price", money, { bold: true, accent: true });

// ---- USES ----
header("USES OF CAPITAL  (development budget)");
const uLand = line("Land contribution", { formula: `${B(landBas)}`, result: P1.uLand }, "contributed by Capital H6", money);
const uHoriz = line("Horizontal site development", { formula: `${B(lots)}*${B(horiz)}`, result: P1.uHoriz }, "lots × $/lot (roads, sewer, utilities)", money);
const uSoft = line("Soft costs", { formula: `${B(soft)}`, result: P1.uSoft }, "entitlement, civil, env, legal", money);
const uMkt = line("Marketing & brokerage", { formula: `${B(rev)}*${B(mktPct)}`, result: P1.uMkt }, "% of revenue", money);
const uCont = line("Contingency", { formula: `${B(uHoriz)}*${B(contPct)}`, result: P1.uCont }, "% of hard cost", money);
const uFin = line("Financing / interest carry", { formula: `${B(fin)}`, result: P1.uFin }, "loan interest", money);
const totCost = line("TOTAL PROJECT COST", { formula: `SUM(${B(uLand)}:${B(uFin)})`, result: P1.totCost }, "sum of uses", money, { bold: true, accent: true, fill: true });

// ---- SOURCES ----
header("SOURCES OF CAPITAL  (capital stack)");
const sLoan = line("Development loan (~60% LTC)", { formula: `${B(totCost)}*${B(ltc)}`, result: P1.loan }, "land-development financing", money);
const sLand = line("Land equity (contributed)", { formula: `${B(landBas)}`, result: P1.landEq }, "contributed land", money);
const sCash = line("Cash equity", { formula: `${B(totCost)}-${B(sLoan)}-${B(sLand)}`, result: P1.cashEq }, "balance — co-invest + LP raise", money);
const totCap = line("TOTAL CAPITAL", { formula: `SUM(${B(sLoan)}:${B(sCash)})`, result: P1.totCost }, "= total project cost", money, { bold: true, accent: true, fill: true });

// ---- RETURNS ----
header("RETURNS");
const profit = line("Net development profit", { formula: `${B(rev)}-${B(totCost)}`, result: P1.profit }, "revenue − total cost", money, { bold: true, accent: true });
line("Profit margin (on revenue)", { formula: `${B(profit)}/${B(rev)}`, result: P1.profit / P1.rev }, "", pct);
line("Profit on cost", { formula: `${B(profit)}/${B(totCost)}`, result: P1.profit / P1.totCost }, "", pct);
const equity = line("Total equity (land + cash)", { formula: `${B(sLand)}+${B(sCash)}`, result: P1.totEquity }, "land contribution + cash equity", money);
const distRow = line("Distributable to equity (capital + profit)", { formula: `${B(equity)}+${B(profit)}`, result: P1.dist }, "returned to equity over the hold", money);
line("Equity multiple", { formula: `${B(distRow)}/${B(equity)}`, result: P1.em }, "distributable / equity, levered", mult, { bold: true, accent: true });

// ---- CASH-FLOW TIMING (phasing inputs) ----
header("CASH-FLOW TIMING  (% by year — edit to re-phase; each row sums to 100%)");
yrHead();
const revPhaseRow  = hrow("Revenue phasing (lot sales)", IN.revPhase, pct);
const costPhaseRow = hrow("Cost phasing (land/horizontal/soft)", IN.costPhase, pct);
const drawPhaseRow = hrow("Equity-draw phasing", IN.drawPhase, pct);
const distPhaseRow = hrow("Equity-distribution phasing", IN.distPhase, pct);

// ---- UNLEVERED PROJECT CASH FLOW ----
header("UNLEVERED PROJECT CASH FLOW  (revenue × phasing − cost × phasing)");
yrHead();
const p1RevRow  = hrow("Lot revenue", phased(`$B$${rev}`, revPhaseRow, P1.revArr), money);
const p1CostRow = hrow("Total cost", phased(`$B$${totCost}`, costPhaseRow, P1.costArr, -1), money);
const p1NetRow  = hrow("Net project cash flow", sumRows(p1RevRow, p1CostRow, P1.unlevCF), money, { bold: true, accent: true });
line("Project IRR (unlevered)", { formula: `IRR(B${p1NetRow}:F${p1NetRow})`, result: P1.unlevIRR },
  "lots sell Yr 2–4; land & horizontal up front", pct, { bold: true, accent: true });

// ---- LEVERED EQUITY CASH FLOW + PROMOTE WATERFALL ----
header(`LEVERED EQUITY CASH FLOW  (${(IN.ltc * 100).toFixed(0)}% LTC development loan)`);
yrHead();
const p1DrawRow = hrow("Equity drawn", phased(`$B$${equity}`, drawPhaseRow, P1.drawArr, -1), money);
const p1DistRow = hrow("Equity distributions", phased(`$B$${distRow}`, distPhaseRow, P1.distArr), money);
const p1LevRow  = hrow("Project equity cash flow", sumRows(p1DrawRow, p1DistRow, P1.levCF), money, { bold: true });
line("Project equity IRR (levered)", { formula: `IRR(B${p1LevRow}:F${p1LevRow})`, result: P1.levIRR },
  `equity ~$${(P1.totEquity / 1e6).toFixed(1)}M; unlevered project IRR ~${(P1.unlevIRR * 100).toFixed(0)}%`, pct, { bold: true, accent: true });

header("PROMOTE WATERFALL  (LP 90% / GP 10% · 8% pref · 80/20 to 15% · 70/30 to 20% · 60/40 above)");
yrHead();
const lpRow = hrow("LP cash flow (after promote)", P1.wf.lp.map(Math.round), money);
const gpRow = hrow("GP cash flow (co-invest + promote)", P1.wf.gp.map(Math.round), money);
line("LP IRR (after promote)", { formula: `IRR(B${lpRow}:F${lpRow})`, result: P1.lpIRR }, "what the equity investors earn", pct, { bold: true, accent: true });
line("GP IRR (co-invest + carried interest)", { formula: `IRR(B${gpRow}:F${gpRow})`, result: P1.gpIRR }, "sponsor group — Accountable Equity / Capital H6 / Bob / JAL + partners", pct, { bold: true });
line("GP net profit (co-invest + promote)", { formula: `SUM(B${gpRow}:F${gpRow})`, result: Math.round(P1.gpProfit) }, "GP group total — JAL earns a share", money, { bold: true });
{
  const a = row(); ws.mergeCells(r, 1, r, 6);
  a.getCell(1).value = "LP / GP split is solved by the model's promote engine (an IRR-hurdle waterfall isn't a closed-form cell formula); the IRR / profit cells above are live over those rows.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } };
}

// ---- LP IRR SENSITIVITY ----
header("LP IRR SENSITIVITY  (finished lot price  ×  LTC)");
const sPrices = [130000, 150000, 170000];
const sLTC = [0.55, 0.60, 0.65, 0.70];
const shr = row(); shr.getCell(1).value = "Lot price  /  LTC"; shr.getCell(1).font = { bold: true, size: 9, color: { argb: WHITE } }; shr.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
sLTC.forEach((l, i) => { const c = shr.getCell(2 + i); c.value = l; c.numFmt = "0%"; c.font = { bold: true, size: 10, color: { argb: WHITE } }; c.alignment = { horizontal: "center" }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
sPrices.forEach((p) => {
  const rr = row(); rr.getCell(1).value = "$" + (p / 1000) + "K / lot"; rr.getCell(1).font = { bold: true, size: 10, color: { argb: NAVY } }; rr.getCell(1).alignment = { indent: 1 };
  sLTC.forEach((l, i) => { const v = p1LPIRR(p, l); const c = rr.getCell(2 + i); c.value = v; c.numFmt = "0.0%"; const base = (p === IN.lotPx && l === IN.ltc); c.font = { size: 10, bold: base, color: { argb: base ? AUB : NAVY } }; c.alignment = { horizontal: "center" }; if (base) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } }; });
});
{
  const a = row(); ws.mergeCells(r, 1, r, 6);
  a.getCell(1).value = "LP IRR after the 8% pref + tiered promote. Base = $150K lot / 60% LTC (highlighted). Each cell re-runs the full waterfall engine.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } };
}

// ---- footer note ----
r++;
{
  const a = row(); ws.mergeCells(r, 1, r, 6);
  a.getCell(1).value = "Illustrative and for discussion only. Figures compiled from public sources and independent research; subject to change and confirmation. Finished-lot model: the JV captures the homebuilder's development margin but takes horizontal execution risk and more capital than selling entitled (paper) lots.";
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
// vertical cash-flow block: one row per year, value (or {formula,result}) in col B
function qcf(label, items) {
  const start = q + 1;
  items.forEach((it, i) => {
    const a = qrow(); a.getCell(1).value = label + " — Year " + i; a.getCell(1).alignment = { indent: 1 }; a.getCell(1).font = { name: "Calibri", size: 10, color: { argb: "FF333333" } };
    const c = a.getCell(2); c.value = it; c.numFmt = money; c.font = { name: "Calibri", size: 10, color: { argb: NAVY } }; c.alignment = { horizontal: "right" };
  });
  return { start, end: q };
}

qband("Queenstown Harbor — Phase 2: Hospitality (Hotel + Restaurants)",
  "VIVÂMEE-led resort; JAL leads capital formation.  ·  Dynamic — every cell below is a live formula off INPUTS.  ·  ILLUSTRATIVE.");

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

qheader("HOTEL");
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
  a.getCell(1).value = "Illustrative. Hospitality is an income / hold play: it develops to roughly cost, so the return is durable cash flow + appreciation (lower IRR than the Phase 1 land, longer hold) — unless you build to core and recap at stabilization. VIVÂMEE operates; JAL leads capital formation. The resort also lifts Phase 1 lot values. Subject to confirmation.";
  a.getCell(1).font = { name: "Calibri", italic: true, size: 8.5, color: { argb: "FF6B5A6B" } };
  a.getCell(1).alignment = { wrapText: true, vertical: "top" }; a.height = 50;
}

wb.xlsx.writeFile("JAL_Queenstown_Harbor_Model.xlsx")
  .then(() => console.log(
    "Wrote model  P1: unlev ~" + (P1.unlevIRR * 100).toFixed(1) + "% / lev ~" + (P1.levIRR * 100).toFixed(1) +
    "% / LP ~" + (P1.lpIRR * 100).toFixed(1) + "% / GP ~" + (P1.gpIRR * 100).toFixed(1) + "% / EM " + P1.em.toFixed(2) + "x" +
    "\n             P1 unlevCF($M): [" + P1.unlevCF.map((v) => (v / 1e6).toFixed(1)).join(", ") + "]" +
    "  levCF($M): [" + P1.levCF.map((v) => (v / 1e6).toFixed(1)).join(", ") + "]" +
    "\n             P2: hold ~" + (P2.irr2 * 100).toFixed(1) + "% (" + P2.em2.toFixed(2) + "x) / b-t-c ~" + (P2.bcIRR * 100).toFixed(1) + "% (" + P2.bcEM.toFixed(2) + "x)" +
    "\n             P2 b-t-c CF($M): [" + P2.bcf.map((v) => (v / 1e6).toFixed(1)).join(", ") + "]  recap $" + (P2.bcVal / 1e6).toFixed(1) + "M" +
    "\n             P2 sens NOI$4.65M: 7%=" + (p2IRR(.22, .07) * 100).toFixed(0) + " 7.5%=" + (p2IRR(.22, .075) * 100).toFixed(0) + " 8.5%=" + (p2IRR(.22, .085) * 100).toFixed(0)))
  .catch((e) => { console.error(e); process.exit(1); });
