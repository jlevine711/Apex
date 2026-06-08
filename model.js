/**
 * Queenstown Harbor — Finished-Lot Development Model (backup to the deck).
 *
 * Finished-lot basis: the JV funds the horizontal (roads, sewer, utilities)
 * and sells FINISHED lots to homebuilders; builders just build the homes.
 * Every output cell is a live formula referencing the INPUTS block, so the
 * model shows exactly how each number is derived. All figures illustrative.
 *
 *   npm run model   ->   JAL_Queenstown_Harbor_Model.xlsx
 */
const ExcelJS = require("exceljs");

// ---- IRR (bisection) so we can cache a result Excel shows before recalc ----
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

const NAVY = "FF0B163C", AUB = "FF3A243A", CREAM = "FFF4F2ED", WHITE = "FFFFFFFF", MAUVE = "FFC4B8C4";

const wb = new ExcelJS.Workbook();
wb.creator = "JAL Strategies";
wb.title = "Queenstown Harbor — Finished-Lot Development Model";
const ws = wb.addWorksheet("Finished-Lot Model", {
  views: [{ showGridLines: false }],
  properties: { defaultRowHeight: 16 },
});
ws.columns = [
  { width: 40 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 12 },
];

let r = 0;
const row = () => ws.getRow(++r);
const money = '$#,##0';
const money1 = '$#,##0,,"M"';
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
  a.getCell(1).font = { name: "Calibri", size: 10, bold: !!opts.bold, color: { argb: opts.bold ? NAVY.replace("FF","FF") : "FF333333" } };
  a.getCell(1).alignment = { indent: 1 };
  const vc = a.getCell(2);
  vc.value = value;
  if (fmt) vc.numFmt = fmt;
  vc.font = { name: "Calibri", size: 10, bold: opts.bold !== false, color: { argb: opts.accent ? AUB : "FF0B163C" } };
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

band("Queenstown Harbor — Finished-Lot Development Model",
  "JV funds the horizontal (roads, sewer, utilities); homebuilders just build the homes.  ·  ILLUSTRATIVE — subject to survey, easement, entitlement & capital terms.");

// ---- INPUTS ----
header("INPUTS  (edit these — everything below recalculates)");
const grossAc = line("Gross Queenstown land (acres)", 700, "owner's figure — to confirm by survey", "#,##0");
const consAc  = line("Less: permanent conservation (acres)", 198, "recorded easement (qhgolf.com)", "#,##0");
const golfAc  = line("Less: golf, range, water & lodging (acres)", 302, "two 18s + 9-ac range + clubhouse/cottages (est.)", "#,##0");
const critAc  = line("Less: Critical Area & wetlands (acres)", 60, "Chesapeake Bay 1,000-ft zone (est.)", "#,##0");
const density = line("Lot density (lots / acre)", 2.0, "blended; Queen Anne's Co. zoning to confirm", "0.0");
const lotPx   = line("Finished lot price ($ / lot)", 150000, "conservative — QAC lots avg ~$237K; new homes $600–705K", money);
const horiz   = line("Horizontal site dev cost ($ / lot)", 60000, "JV-funded; national benchmark $50–150K/acre", money);
const soft    = line("Soft costs ($)", 2800000, "entitlement, civil, environmental, legal, mgmt", money);
const mktPct  = line("Marketing & brokerage (% of revenue)", 0.03, "lot sales", pct);
const contPct = line("Contingency (% of hard cost)", 0.07, "on horizontal", pct);
const fin     = line("Financing / interest carry ($)", 1400000, "development loan interest", money);
const landBas = line("Land basis — contributed ($)", 4000000, "allocated from $25M H6 purchase (~$29K/ac); appraisal to set", money);
const ltc     = line("Development loan (% of total cost)", 0.70, "land-development / A&D financing (70% LTC)", pct);
const pref    = line("Preferred return", 0.08, "to equity, before promote", pct);

const B = (i) => `B${i}`;

// ---- LAND -> LOTS ----
header("DEVELOPABLE LAND  →  LOTS");
const netAc = line("Net developable land (acres)",
  { formula: `${B(grossAc)}-${B(consAc)}-${B(golfAc)}-${B(critAc)}`, result: 140 },
  "gross − conservation − golf − Critical Area", "#,##0", { bold: true });
const lots = line("Finished lots",
  { formula: `${B(netAc)}*${B(density)}`, result: 280 },
  "net developable acres × density", "#,##0", { bold: true, accent: true });

// ---- REVENUE ----
header("REVENUE");
const rev = line("Gross lot revenue",
  { formula: `${B(lots)}*${B(lotPx)}`, result: 42000000 },
  "finished lots × finished lot price", money, { bold: true, accent: true });

// ---- USES ----
header("USES OF CAPITAL  (development budget)");
const uLand = line("Land contribution", { formula: `${B(landBas)}`, result: 4000000 }, "contributed by Capital H6", money);
const uHoriz = line("Horizontal site development", { formula: `${B(lots)}*${B(horiz)}`, result: 16800000 }, "lots × $/lot (roads, sewer, utilities)", money);
const uSoft = line("Soft costs", { formula: `${B(soft)}`, result: 2800000 }, "entitlement, civil, env, legal", money);
const uMkt = line("Marketing & brokerage", { formula: `${B(rev)}*${B(mktPct)}`, result: 1260000 }, "% of revenue", money);
const uCont = line("Contingency", { formula: `${B(uHoriz)}*${B(contPct)}`, result: 1176000 }, "% of hard cost", money);
const uFin = line("Financing / interest carry", { formula: `${B(fin)}`, result: 1400000 }, "loan interest", money);
const totCost = line("TOTAL PROJECT COST", { formula: `SUM(${B(uLand)}:${B(uFin)})`, result: 27436000 }, "sum of uses", money, { bold: true, accent: true, fill: true });

// ---- SOURCES ----
header("SOURCES OF CAPITAL  (capital stack)");
const sLoan = line("Development loan (~60% LTC)", { formula: `${B(totCost)}*${B(ltc)}`, result: 16461600 }, "land-development financing", money);
const sLand = line("Land equity (H6)", { formula: `${B(landBas)}`, result: 4000000 }, "contributed land", money);
const sCash = line("Cash equity (JAL + LP)", { formula: `${B(totCost)}-${B(sLoan)}-${B(sLand)}`, result: 6974400 }, "balance — co-invest + LP raise", money);
const totCap = line("TOTAL CAPITAL", { formula: `SUM(${B(sLoan)}:${B(sCash)})`, result: 27436000 }, "= total project cost", money, { bold: true, accent: true, fill: true });

// ---- RETURNS ----
header("RETURNS");
const profit = line("Net development profit", { formula: `${B(rev)}-${B(totCost)}`, result: 14564000 }, "revenue − total cost", money, { bold: true, accent: true });
line("Profit margin (on revenue)", { formula: `${B(profit)}/${B(rev)}`, result: 0.35 }, "", pct);
line("Profit on cost", { formula: `${B(profit)}/${B(totCost)}`, result: 0.537 }, "", pct);
const equity = line("Total equity (land + cash)", { formula: `${B(sLand)}+${B(sCash)}`, result: 10974400 }, "land contribution + cash equity", money);
line("Equity multiple", { formula: `(${B(equity)}+${B(profit)})/${B(equity)}`, result: 2.34 }, "(equity + profit) / equity, levered", mult, { bold: true, accent: true });

// ---- CASH FLOW / IRR ----
const cfHeaderRow = header("PHASED PROJECT CASH FLOW  (unlevered, $ — for IRR)");
const yr = row();
yr.getCell(1).value = "Year";
["Yr 0", "Yr 1", "Yr 2", "Yr 3", "Yr 4"].forEach((y, i) => {
  const c = yr.getCell(2 + i); c.value = y; c.font = { bold: true, size: 9, color: { argb: "FF6B5A6B" } }; c.alignment = { horizontal: "right" };
});
const cfs = [-8000000, -10000000, 2500000, 18000000, 12000000];
const cfRow = row();
cfRow.getCell(1).value = "Net project cash flow";
cfRow.getCell(1).font = { size: 10, color: { argb: "FF333333" } };
cfRow.getCell(1).alignment = { indent: 1 };
cfs.forEach((v, i) => { const c = cfRow.getCell(2 + i); c.value = v; c.numFmt = money; c.font = { size: 10, color: { argb: "FF0B163C" } }; c.alignment = { horizontal: "right" }; });
const irrVal = irr(cfs);
const irrRow = row();
irrRow.getCell(1).value = "Project IRR (unlevered)";
irrRow.getCell(1).font = { bold: true, size: 10, color: { argb: "FF0B163C" } };
irrRow.getCell(1).alignment = { indent: 1 };
const irrCell = irrRow.getCell(2);
irrCell.value = { formula: `IRR(B${cfRow.number}:F${cfRow.number})`, result: irrVal };
irrCell.numFmt = pct;
irrCell.font = { bold: true, size: 10, color: { argb: AUB } };
irrCell.alignment = { horizontal: "right" };
const cfnote = irrRow.getCell(3); ws.mergeCells(irrRow.number, 3, irrRow.number, 6);
cfnote.value = "entitlement/land Yr 0–1, horizontal Yr 1–2, finished-lot sales Yr 2–4"; cfnote.font = { italic: true, size: 9, color: { argb: "FF6B5A6B" } };

// ---- LEVERED EQUITY + PROMOTE WATERFALL ----
header("LEVERED EQUITY RETURN  (70% LTC development loan)");
const lyr = row(); lyr.getCell(1).value = "Year"; ["Yr 0", "Yr 1", "Yr 2", "Yr 3", "Yr 4"].forEach((y, i) => { const c = lyr.getCell(2 + i); c.value = y; c.font = { bold: true, size: 9, color: { argb: "FF6B5A6B" } }; c.alignment = { horizontal: "right" }; });
const pcf = [-5000000, -3231000, 2000000, 7000000, 13795000]; // equity in (Yr 0–1) → distributions (Yr 2–4)
const pcfRow = row(); pcfRow.getCell(1).value = "Project equity cash flow"; pcfRow.getCell(1).font = { size: 10, color: { argb: "FF333333" } }; pcfRow.getCell(1).alignment = { indent: 1 };
pcf.forEach((v, i) => { const c = pcfRow.getCell(2 + i); c.value = v; c.numFmt = money; c.font = { size: 10, color: { argb: "FF0B163C" } }; c.alignment = { horizontal: "right" }; });
const lirr = irr(pcf);
const lirrRow = row(); lirrRow.getCell(1).value = "Project equity IRR (levered)"; lirrRow.getCell(1).font = { bold: true, size: 10, color: { argb: "FF0B163C" } }; lirrRow.getCell(1).alignment = { indent: 1 };
const lc = lirrRow.getCell(2); lc.value = { formula: `IRR(B${pcfRow.number}:F${pcfRow.number})`, result: lirr }; lc.numFmt = pct; lc.font = { bold: true, size: 10, color: { argb: AUB } }; lc.alignment = { horizontal: "right" };
const lnn = lirrRow.getCell(3); ws.mergeCells(lirrRow.number, 3, lirrRow.number, 6); lnn.value = "equity ~$8.25M (70% LTC); lot sales Yr 2–4. Unlevered project IRR ~24%."; lnn.font = { italic: true, size: 9, color: { argb: "FF6B5A6B" } };

header("PROMOTE WATERFALL  (LP 90% / GP 10% · 8% pref · 80/20 to 15% · 70/30 to 20% · 60/40 above)");
const wf = waterfall(pcf, 0.90, 0.08, [{ irr: 0.15, lp: 0.80 }, { irr: 0.20, lp: 0.70 }, { irr: Infinity, lp: 0.60 }]);
const lpIRR = irr(wf.lp), gpIRR = irr(wf.gp);
const lpYr = row(); lpYr.getCell(1).value = "LP cash flow (after promote)"; lpYr.getCell(1).font = { size: 10, color: { argb: "FF333333" } }; lpYr.getCell(1).alignment = { indent: 1 };
wf.lp.forEach((v, i) => { const c = lpYr.getCell(2 + i); c.value = Math.round(v); c.numFmt = money; c.font = { size: 10, color: { argb: "FF0B163C" } }; c.alignment = { horizontal: "right" }; });
const gpYr = row(); gpYr.getCell(1).value = "GP cash flow (co-invest + promote)"; gpYr.getCell(1).font = { size: 10, color: { argb: "FF333333" } }; gpYr.getCell(1).alignment = { indent: 1 };
wf.gp.forEach((v, i) => { const c = gpYr.getCell(2 + i); c.value = Math.round(v); c.numFmt = money; c.font = { size: 10, color: { argb: "FF0B163C" } }; c.alignment = { horizontal: "right" }; });
line("LP IRR (after promote)", { formula: `IRR(B${lpYr.number}:F${lpYr.number})`, result: lpIRR }, "what the equity investors earn", pct, { bold: true, accent: true });
line("GP IRR (co-invest + carried interest)", { formula: `IRR(B${gpYr.number}:F${gpYr.number})`, result: gpIRR }, "JAL", pct, { bold: true });
line("GP net profit (co-invest + promote)", { formula: `SUM(B${gpYr.number}:F${gpYr.number})`, result: Math.round(wf.gp.reduce((a, b) => a + b, 0)) }, "JAL's share of the upside", money, { bold: true });

// ---- footer note ----
r++;
const note = row();
ws.mergeCells(r, 1, r, 6);
note.getCell(1).value = "Illustrative and for discussion only. Figures compiled from public sources and independent research; subject to change and confirmation in further conversations. Finished-lot model: the JV captures the homebuilder's development margin but takes horizontal execution risk and more capital than selling entitled (paper) lots.";
note.getCell(1).font = { italic: true, size: 8.5, color: { argb: "FF6B5A6B" } };
note.getCell(1).alignment = { wrapText: true, vertical: "top" };
note.height = 42;

// =================== PHASE 2 — HOSPITALITY (HOTEL + RESTAURANTS) ===================
// Income / hold play, VIVAMEE-led, JAL as capital partner. Source of truth for
// the deck's Phase 2 slide — live formulas off the INPUTS block.
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
  const a = qrow(); a.getCell(1).value = label; a.getCell(1).font = { name: "Calibri", size: 10, bold: !!opts.bold, color: { argb: opts.bold ? "FF0B163C" : "FF333333" } }; a.getCell(1).alignment = { indent: 1 };
  const vc = a.getCell(2); vc.value = value; if (fmt) vc.numFmt = fmt; vc.font = { name: "Calibri", size: 10, bold: opts.bold !== false, color: { argb: opts.accent ? AUB : "FF0B163C" } }; vc.alignment = { horizontal: "right" };
  if (note) { const nc = a.getCell(3); p2.mergeCells(q, 3, q, 6); nc.value = note; nc.font = { name: "Calibri", italic: true, size: 9, color: { argb: "FF6B5A6B" } }; }
  if (opts.fill) for (let c = 1; c <= 2; c++) a.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };
  return q;
}

qband("Queenstown Harbor — Phase 2: Hospitality (Hotel + Restaurants)",
  "VIVAMEE-led resort; JAL as capital partner. Income / hold play — develops to ~cost; returns come from cash flow + appreciation.  ·  ILLUSTRATIVE.");

qheader("INPUTS  (edit these — everything below recalculates)");
const hKeys = qline("Hotel — keys", 100, "boutique resort lodge", "#,##0");
const hCostKey = qline("Hotel — cost / key ($)", 325000, "HVS 2025: select ~$223K, full-service ~$409K", money);
const hADR = qline("Hotel — ADR ($)", 250, "upscale Eastern Shore resort (weddings / golf)", money);
const hOcc = qline("Hotel — stabilized occupancy", 0.62, "leisure / resort", pct);
const hMult = qline("Hotel — total-revenue multiple (x rooms)", 1.6, "+ F&B, banquets, spa", "0.00");
const hMgn = qline("Hotel — NOI margin", 0.31, "stabilized", pct);
const rSF = qline("Restaurants — GLA (SF)", 16000, "a suite of destination restaurants", "#,##0");
const rCostSF = qline("Restaurants — build cost / SF ($)", 550, "full-service ~$555/SF", money);
const rSalesSF = qline("Restaurants — sales / SF ($)", 700, "upscale destination dining", money);
const rMgn = qline("Restaurants — NOI (% of sales)", 0.09, "rent / operating contribution", pct);
const pLand = qline("Land — hospitality parcel ($)", 2000000, "contributed by Capital H6", money);
const pSoft = qline("Soft / FF&E / pre-opening ($)", 3700000, "", money);
const pCap = qline("Exit cap rate", 0.08, "hotel cap ~8% (2025)", pct);
const pLTC = qline("Construction loan (% of cost)", 0.60, "", pct);
const pRate = qline("Loan interest rate", 0.07, "", pct);
const pGrow = qline("NOI growth / yr", 0.03, "", pct);

qheader("HOTEL");
const hCost = qline("Hotel development cost", { formula: `${Q(hKeys)}*${Q(hCostKey)}`, result: 32500000 }, "keys x cost/key", money);
const hRoom = qline("Room revenue", { formula: `${Q(hKeys)}*365*${Q(hADR)}*${Q(hOcc)}`, result: 5657500 }, "keys x 365 x ADR x occ", money);
const hRev = qline("Total hotel revenue", { formula: `${Q(hRoom)}*${Q(hMult)}`, result: 9052000 }, "x revenue multiple", money);
const hNOI = qline("Hotel NOI (stabilized)", { formula: `${Q(hRev)}*${Q(hMgn)}`, result: 2806120 }, "x NOI margin", money, { bold: true, accent: true });

qheader("RESTAURANTS  (a suite of nice restaurants)");
const rCost = qline("Restaurant development cost", { formula: `${Q(rSF)}*${Q(rCostSF)}`, result: 8800000 }, "SF x cost/SF", money);
const rSales = qline("Restaurant sales", { formula: `${Q(rSF)}*${Q(rSalesSF)}`, result: 11200000 }, "SF x sales/SF", money);
const rNOI = qline("Restaurant NOI (to owner)", { formula: `${Q(rSales)}*${Q(rMgn)}`, result: 1008000 }, "rent / operating contribution", money, { bold: true, accent: true });

qheader("TOTAL PROJECT  ->  STABILIZED");
const pCost = qline("Total project cost", { formula: `${Q(hCost)}+${Q(rCost)}+${Q(pLand)}+${Q(pSoft)}`, result: 47000000 }, "hotel + restaurants + land + soft", money, { bold: true, accent: true, fill: true });
const pNOI = qline("Stabilized NOI", { formula: `${Q(hNOI)}+${Q(rNOI)}`, result: 3814120 }, "hotel + restaurant NOI", money, { bold: true, accent: true, fill: true });
const pYoC = qline("Yield on cost", { formula: `${Q(pNOI)}/${Q(pCost)}`, result: 0.0811 }, "NOI / cost", pct, { bold: true });
const pLoan = qline("Construction loan", { formula: `${Q(pCost)}*${Q(pLTC)}`, result: 28200000 }, "% LTC", money);
const pEq = qline("Equity (land + cash)", { formula: `${Q(pCost)}-${Q(pLoan)}`, result: 18800000 }, "cost - loan", money);
const pCoC = qline("Cash-on-cash (stabilized, levered)", { formula: `(${Q(pNOI)}-${Q(pLoan)}*${Q(pRate)})/${Q(pEq)}`, result: 0.098 }, "(NOI - interest) / equity", pct);

qheader("EXIT  &  RETURNS  (7-year hold)");
const pXNOI = qline("Year-7 NOI (3%/yr growth)", { formula: `${Q(pNOI)}*(1+${Q(pGrow)})^4`, result: 4292800 }, "stabilized x growth", money);
const pXVal = qline("Exit value (@ exit cap)", { formula: `${Q(pXNOI)}/${Q(pCap)}`, result: 53660000 }, "Year-7 NOI / cap", money, { bold: true });
const pXEq = qline("Exit equity (value - loan)", { formula: `${Q(pXVal)}-${Q(pLoan)}`, result: 25460000 }, "net of loan", money);

qheader("EQUITY CASH FLOW (levered, $) & RETURNS");
const cf2 = [-9400000, -9400000, 900000, 1840000, 1900000, 1950000, 2010000, 27530000];
const cf2Start = q + 1;
cf2.forEach((v, i) => {
  const a = qrow(); a.getCell(1).value = "Year " + i; a.getCell(1).alignment = { indent: 1 }; a.getCell(1).font = { name: "Calibri", size: 10, color: { argb: "FF333333" } };
  const c = a.getCell(2); c.value = v; c.numFmt = money; c.font = { name: "Calibri", size: 10, color: { argb: "FF0B163C" } }; c.alignment = { horizontal: "right" };
});
const cf2End = q;
const irr2 = irr(cf2);
const em2 = cf2.slice(2).reduce((a, b) => a + b, 0) / -(cf2[0] + cf2[1]);
qline("Equity multiple (7-yr)", { formula: `SUM(B${cf2Start + 2}:B${cf2End})/-(B${cf2Start}+B${cf2Start + 1})`, result: em2 }, "distributions / equity", mult, { bold: true, accent: true });
qline("Project IRR (levered, 7-yr)", { formula: `IRR(B${cf2Start}:B${cf2End})`, result: irr2 }, "equity cash flows Yr 0–7", pct, { bold: true, accent: true });

qheader("BUILD-TO-CORE CASE  (path to mid–high 20s IRR — recap / sell at stabilization)");
const bcUp = qline("NOI uplift — premium execution", 0.22, "VIVAMEE experiential: ADR ~$275, occ ~63%, strong F&B / events", pct);
const bcNOI = qline("Premium stabilized NOI", { formula: `${Q(pNOI)}*(1+${Q(bcUp)})`, result: 4653226 }, "base NOI x (1 + uplift)", money, { bold: true, accent: true });
const bcCap = qline("Exit cap (build-to-core)", 0.075, "trophy / experiential resort", pct);
const bcLTC = qline("Leverage (% of cost)", 0.65, "construction loan", pct);
const bcYoC = qline("Yield on cost (premium)", { formula: `${Q(bcNOI)}/${Q(pCost)}`, result: 0.099 }, "premium NOI / cost — vs 7.5% exit = the spread", pct, { bold: true });
const bcVal = qline("Recap / sale value (~Yr 4)", { formula: `${Q(bcNOI)}/${Q(bcCap)}`, result: 62043015 }, "premium NOI / exit cap", money, { bold: true });
const bcLoan = qline("Loan", { formula: `${Q(pCost)}*${Q(bcLTC)}`, result: 30550000 }, "% of cost", money);
const bcEq = qline("Equity", { formula: `${Q(pCost)}-${Q(bcLoan)}`, result: 16450000 }, "cost - loan", money);
const bcXEq = qline("Exit equity (value - loan)", { formula: `${Q(bcVal)}-${Q(bcLoan)}`, result: 31493015 }, "at recap / sale", money);
qheader("BUILD-TO-CORE EQUITY CASH FLOW (Yr 0–4, $) & RETURNS");
const bcf = [-8225000, -8225000, 1000000, 2514726, 34007741];
const bcfStart = q + 1;
bcf.forEach((v, i) => {
  const a = qrow(); a.getCell(1).value = "Year " + i; a.getCell(1).alignment = { indent: 1 }; a.getCell(1).font = { name: "Calibri", size: 10, color: { argb: "FF333333" } };
  const c = a.getCell(2); c.value = v; c.numFmt = money; c.font = { name: "Calibri", size: 10, color: { argb: "FF0B163C" } }; c.alignment = { horizontal: "right" };
});
const bcfEnd = q;
const bcIRR = irr(bcf);
const bcEM = bcf.slice(2).reduce((a, b) => a + b, 0) / -(bcf[0] + bcf[1]);
qline("Equity multiple (build-to-core)", { formula: `SUM(B${bcfStart + 2}:B${bcfEnd})/-(B${bcfStart}+B${bcfStart + 1})`, result: bcEM }, "distributions / equity, ~4-yr", mult, { bold: true, accent: true });
qline("Project IRR (build-to-core, ~4-yr)", { formula: `IRR(B${bcfStart}:B${bcfEnd})`, result: bcIRR }, "recap / sell at stabilization", pct, { bold: true, accent: true });

q++;
const n2 = qrow(); p2.mergeCells(q, 1, q, 6);
n2.getCell(1).value = "Illustrative. Hospitality is an income / hold play: it develops to roughly cost, so the return is durable cash flow + appreciation (lower IRR than the Phase 1 land, longer hold). VIVAMEE operates; JAL is the capital partner (debt + equity placement fees, co-invest, promote). The resort also lifts Phase 1 lot values. Subject to confirmation.";
n2.getCell(1).font = { name: "Calibri", italic: true, size: 8.5, color: { argb: "FF6B5A6B" } };
n2.getCell(1).alignment = { wrapText: true, vertical: "top" }; n2.height = 50;

wb.xlsx.writeFile("JAL_Queenstown_Harbor_Model.xlsx")
  .then(() => console.log("Wrote model  (P1 unlev ~" + (irrVal * 100).toFixed(1) + "% / levered ~" + (lirr * 100).toFixed(1) + "% / LP ~" + (lpIRR * 100).toFixed(1) + "% / GP ~" + (gpIRR * 100).toFixed(1) + "%; P2 build-to-core ~" + (bcIRR * 100).toFixed(1) + "%)"))
  .catch((e) => { console.error(e); process.exit(1); });
