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
const ltc     = line("Development loan (% of total cost)", 0.60, "land-development / A&D financing", pct);
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

// ---- footer note ----
r++;
const note = row();
ws.mergeCells(r, 1, r, 6);
note.getCell(1).value = "Illustrative and for discussion only. Figures compiled from public sources and independent research; subject to change and confirmation in further conversations. Finished-lot model: the JV captures the homebuilder's development margin but takes horizontal execution risk and more capital than selling entitled (paper) lots.";
note.getCell(1).font = { italic: true, size: 8.5, color: { argb: "FF6B5A6B" } };
note.getCell(1).alignment = { wrapText: true, vertical: "top" };
note.height = 42;

wb.xlsx.writeFile("JAL_Queenstown_Harbor_Model.xlsx")
  .then(() => console.log("Wrote JAL_Queenstown_Harbor_Model.xlsx  (project IRR ~", (irrVal * 100).toFixed(1) + "%)"))
  .catch((e) => { console.error(e); process.exit(1); });
