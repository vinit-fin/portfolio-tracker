/**
 * ============================================================
 *  INDIAN EQUITY PORTFOLIO TRACKER — Google Apps Script
 *  Author  : [vinit kumar]
 *  GitHub  : [vinit-fin]
 *  Version : 2.0  |  Live prices via GOOGLEFINANCE()
 * ============================================================
 *
 *  HOW TO RUN:
 *  1. Go to https://script.google.com  → New Project
 *  2. Paste this entire file → Save
 *  3. Select function  buildPortfolio  → Click ▶ Run
 *  4. Approve permissions once → Done!
 *
 *  UPDATE PRICES: Prices refresh automatically every 5 min
 *  via GOOGLEFINANCE(). Force refresh: Ctrl+Shift+F9
 * ============================================================
 */

/* ── CONFIGURATION — Edit only this block ── */
const CONFIG = {
  spreadsheetName : "Indian Equity Portfolio Tracker",
  currency        : "₹",
  refreshMinutes  : 5,      // trigger interval for auto-refresh
  stocks: [
    // [ Stock Name,                    NSE Ticker,   Sector,       Qty,  Buy Price ]
    [ "Infosys Ltd",                  "INFY",       "IT",         10,   1750  ],
    [ "Tech Mahindra Ltd",            "TECHM",      "IT",         15,   1400  ],
    [ "KPIT Technologies Ltd",        "KPITTECH",   "IT",         20,   1650  ],
    [ "Axis Bank Ltd",                "AXISBANK",   "Banking",    25,   1100  ],
    [ "HDFC Bank Ltd",                "HDFCBANK",   "Banking",    20,   1700  ],
    [ "RBL Bank Ltd",                 "RBLBANK",    "Banking",   100,    290  ],
    [ "Glenmark Pharmaceuticals Ltd", "GLENMARK",   "Pharma",     30,    860  ],
    [ "Cipla Ltd",                    "CIPLA",      "Pharma",     15,   1380  ],
    [ "Sun Pharmaceutical Ind. Ltd",  "SUNPHARMA",  "Pharma",     10,   1700  ],
    [ "Reliance Industries Ltd",      "RELIANCE",   "Conglomerate", 5,  2850  ],
  ]
};

/* ── SECTOR COLORS (hex without #) ── */
const COLORS = {
  IT          : { bg: "#DEEAF1", fg: "#1A3A5C" },
  Banking     : { bg: "#E2EFDA", fg: "#1A3A1A" },
  Pharma      : { bg: "#FFF2CC", fg: "#5C4A00" },
  Conglomerate: { bg: "#FCE4D6", fg: "#5C2000" },
  headerDark  : "#1F3864",
  headerMid   : "#2E75B6",
  white       : "#FFFFFF",
  rowAlt      : "#F9F9F9",
  green       : "#1E7B4B",
  red         : "#C00000",
};

/* ════════════════════════════════════════════
   MAIN ENTRY POINT
════════════════════════════════════════════ */
function buildPortfolio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
           || SpreadsheetApp.create(CONFIG.spreadsheetName);

  // Remove existing sheets to rebuild cleanly
  ["Portfolio Tracker","Dashboard","README"].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  buildTrackerSheet(ss);
  buildDashboardSheet(ss);
  buildReadmeSheet(ss);
  setupAutoRefresh();

  // Activate Portfolio Tracker
  ss.setActiveSheet(ss.getSheetByName("Portfolio Tracker"));
  SpreadsheetApp.getUi().alert(
    "✅ Portfolio Tracker built!\n\nLive prices load via GOOGLEFINANCE() — allow ~5 sec.\nRefreshes every " + CONFIG.refreshMinutes + " min automatically."
  );
}

/* ════════════════════════════════════════════
   SHEET 1 — PORTFOLIO TRACKER
════════════════════════════════════════════ */
function buildTrackerSheet(ss) {
  const ws = ss.insertSheet("Portfolio Tracker", 0);
  ws.setHiddenGridlines(true);

  // ── Row heights
  ws.setRowHeight(1, 38);
  ws.setRowHeight(2, 18);
  ws.setRowHeight(3, 38);
  for (let i = 4; i <= 13; i++) ws.setRowHeight(i, 24);
  ws.setRowHeight(14, 26);

  // ── Column widths
  const colWidths = [40, 260, 120, 90, 80, 110, 110, 140, 140, 130, 90, 90];
  colWidths.forEach((w, i) => ws.setColumnWidth(i + 1, w));

  // ── Title
  const titleRange = ws.getRange("A1:L1");
  titleRange.merge()
    .setValue("📊  INDIAN EQUITY PORTFOLIO TRACKER  — Live Prices via GOOGLEFINANCE()")
    .setBackground(COLORS.headerDark)
    .setFontColor(COLORS.white)
    .setFontSize(13)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ── Subtitle
  ws.getRange("A2:F2").merge()
    .setValue("🔵 Blue = Inputs you can edit  |  ⚫ Black = Auto-calculated  |  🟢 Green = Live via GOOGLEFINANCE()")
    .setFontSize(8).setFontColor("#595959").setFontStyle("italic")
    .setVerticalAlignment("middle");
  ws.getRange("G2:L2").merge()
    .setValue("📡 Prices auto-refresh every " + CONFIG.refreshMinutes + " min  |  Force refresh: Ctrl+Shift+F9")
    .setFontSize(8).setFontColor("#595959").setFontStyle("italic")
    .setHorizontalAlignment("right").setVerticalAlignment("middle");

  // ── Column headers
  const headers = [
    "#","Stock Name","Sector","Ticker (NSE)",
    "Qty","Buy Price (₹)","Live Price (₹)",
    "Invested (₹)","Current Value (₹)","P&L (₹)","P&L (%)","52W High (₹)"
  ];
  const hdrRange = ws.getRange(3, 1, 1, headers.length);
  hdrRange.setValues([headers])
    .setBackground(COLORS.headerMid)
    .setFontColor(COLORS.white)
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  setBorderAll(hdrRange);

  // ── Stock rows
  CONFIG.stocks.forEach((stock, idx) => {
    const [name, ticker, sector, qty, buyPrice] = stock;
    const r   = idx + 4;
    const col = COLORS[sector] || { bg: "#F9F9F9", fg: "#000000" };
    const ns  = `NSE:${ticker}`;   // e.g. "NSE:INFY"

    // Col A — serial
    setCellValue(ws, r, 1, idx + 1, col.bg, col.fg, "center");

    // Col B — stock name
    setCellValue(ws, r, 2, name, col.bg, col.fg, "left");
    ws.getRange(r, 2).setFontWeight("bold");

    // Col C — sector
    setCellValue(ws, r, 3, sector, col.bg, col.fg, "center");

    // Col D — ticker (blue = editable label)
    setCellValue(ws, r, 4, ticker, col.bg, "#0000FF", "center");

    // Col E — qty (blue = user input)
    setCellValue(ws, r, 5, qty, col.bg, "#0000FF", "right");
    ws.getRange(r, 5).setNumberFormat("#,##0");

    // Col F — buy price (blue = user input)
    setCellValue(ws, r, 6, buyPrice, col.bg, "#0000FF", "right");
    ws.getRange(r, 6).setNumberFormat("₹#,##0.00");

    // Col G — LIVE price via GOOGLEFINANCE (green = live)
    setCellFormula(ws, r, 7,
      `=IFERROR(GOOGLEFINANCE("${ns}","price"),"⏳")`,
      col.bg, COLORS.green, "right");
    ws.getRange(r, 7).setNumberFormat("₹#,##0.00");

    // Col H — Invested value = Qty × Buy Price
    setCellFormula(ws, r, 8, `=E${r}*F${r}`, col.bg, "#000000", "right");
    ws.getRange(r, 8).setNumberFormat("₹#,##0.00");

    // Col I — Current value = Qty × Live Price
    setCellFormula(ws, r, 9,
      `=IFERROR(E${r}*G${r},"⏳")`,
      col.bg, "#000000", "right");
    ws.getRange(r, 9).setNumberFormat("₹#,##0.00");

    // Col J — P&L ₹ = Current - Invested
    setCellFormula(ws, r, 10,
      `=IFERROR(I${r}-H${r},"⏳")`,
      col.bg, "#000000", "right");
    ws.getRange(r, 10).setNumberFormat("₹#,##0.00");

    // Col K — P&L %
    setCellFormula(ws, r, 11,
      `=IFERROR((I${r}-H${r})/H${r},"⏳")`,
      col.bg, "#000000", "right");
    ws.getRange(r, 11).setNumberFormat("0.00%");

    // Col L — 52-week high via GOOGLEFINANCE (green = live)
    setCellFormula(ws, r, 12,
      `=IFERROR(GOOGLEFINANCE("${ns}","high52"),"⏳")`,
      col.bg, COLORS.green, "right");
    ws.getRange(r, 12).setNumberFormat("₹#,##0.00");
  });

  // ── Conditional formatting: P&L % green/red
  const pnlPctRange = ws.getRange("K4:K13");
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor(COLORS.green).setRanges([pnlPctRange]).build();
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor(COLORS.red).setRanges([pnlPctRange]).build();
  ws.setConditionalFormatRules([greenRule, redRule]);

  // Same for P&L ₹
  const pnlRange = ws.getRange("J4:J13");
  const gRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor(COLORS.green).setRanges([pnlRange]).build();
  const rRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor(COLORS.red).setRanges([pnlRange]).build();
  ws.setConditionalFormatRules(
    ws.getConditionalFormatRules().concat([gRule2, rRule2])
  );

  // ── Totals row (row 14)
  const totCells = [
    [8, "=SUM(H4:H13)","₹#,##0.00"],
    [9, "=IFERROR(SUM(I4:I13),\"⏳\")","₹#,##0.00"],
    [10,"=IFERROR(I14-H14,\"⏳\")","₹#,##0.00"],
    [11,"=IFERROR((I14-H14)/H14,\"⏳\")","0.00%"],
  ];
  ws.getRange("A14:G14").merge()
    .setValue("PORTFOLIO TOTAL")
    .setBackground(COLORS.headerDark).setFontColor(COLORS.white)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  setBorderAll(ws.getRange("A14:G14"));

  totCells.forEach(([col, formula, fmt]) => {
    const c = ws.getRange(14, col);
    c.setFormula(formula)
      .setBackground(COLORS.headerDark).setFontColor(COLORS.white)
      .setFontWeight("bold").setFontSize(10)
      .setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setNumberFormat(fmt);
    setBorderAll(c);
  });

  // ── Freeze header rows
  ws.setFrozenRows(3);
  ws.setFrozenColumns(2);
}

/* ════════════════════════════════════════════
   SHEET 2 — DASHBOARD
════════════════════════════════════════════ */
function buildDashboardSheet(ss) {
  const ws = ss.insertSheet("Dashboard", 1);
  ws.setHiddenGridlines(true);

  // Col widths
  [40,160,110,120,120,120,90,100].forEach((w,i) => ws.setColumnWidth(i+1, w));
  ws.setRowHeight(1, 38);

  // Title
  ws.getRange("A1:H1").merge()
    .setValue("📈  LIVE PORTFOLIO DASHBOARD")
    .setBackground(COLORS.headerDark).setFontColor(COLORS.white)
    .setFontSize(13).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── KPI Cards (row 3-4)
  ws.setRowHeight(2, 10);
  ws.setRowHeight(3, 22);
  ws.setRowHeight(4, 32);

  const kpis = [
    { label:"Total Invested",   formula:"='Portfolio Tracker'!H14",                                             fmt:"₹#,##0.00",  color:"2E75B6" },
    { label:"Current Value",    formula:"='Portfolio Tracker'!I14",                                             fmt:"₹#,##0.00",  color:"1E7B4B" },
    { label:"Total P&L (₹)",   formula:"='Portfolio Tracker'!J14",                                             fmt:"₹#,##0.00",  color:"C00000" },
    { label:"Total P&L (%)",   formula:"=IFERROR(('Portfolio Tracker'!I14-'Portfolio Tracker'!H14)/'Portfolio Tracker'!H14,0)", fmt:"0.00%", color:"7030A0" },
  ];

  kpis.forEach((kpi, i) => {
    const col1 = i * 2 + 1;
    const col2 = col1 + 1;
    const bg = "#" + kpi.color;
    const l1 = columnLetter(col1);
    const l2 = columnLetter(col2);

    ws.getRange(`${l1}3:${l2}3`).merge()
      .setValue(kpi.label)
      .setBackground(bg).setFontColor(COLORS.white)
      .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
    setBorderAll(ws.getRange(`${l1}3:${l2}3`));

    ws.getRange(`${l1}4:${l2}4`).merge()
      .setFormula(kpi.formula)
      .setBackground("#F9F9F9").setFontColor(bg)
      .setFontWeight("bold").setFontSize(14)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setNumberFormat(kpi.fmt);
    setBorderAll(ws.getRange(`${l1}4:${l2}4`));
  });

  // ── Allocation Table (row 6+)
  ws.setRowHeight(5, 14);
  ws.setRowHeight(6, 22);
  ws.setRowHeight(7, 22);

  ws.getRange("A6").setValue("📊 Stock-wise Allocation")
    .setFontWeight("bold").setFontColor(COLORS.headerDark).setFontSize(11);

  const dHeaders = ["#","Stock","Sector","Invested (₹)","Live Value (₹)","P&L (₹)","P&L (%)","Allocation %"];
  ws.getRange(7, 1, 1, 8).setValues([dHeaders])
    .setBackground(COLORS.headerMid).setFontColor(COLORS.white)
    .setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  setBorderAll(ws.getRange(7, 1, 1, 8));

  CONFIG.stocks.forEach((stock, idx) => {
    const [name,,sector] = stock;
    const r   = idx + 8;
    const ptr = idx + 4;   // row in Portfolio Tracker
    const col = COLORS[sector] || { bg: "#F9F9F9" };
    ws.setRowHeight(r, 22);

    const rowData = [
      [idx+1, null],
      [`='Portfolio Tracker'!B${ptr}`, null],
      [`='Portfolio Tracker'!C${ptr}`, null],
      [`='Portfolio Tracker'!H${ptr}`, "₹#,##0.00"],
      [`='Portfolio Tracker'!I${ptr}`, "₹#,##0.00"],
      [`='Portfolio Tracker'!J${ptr}`, "₹#,##0.00"],
      [`='Portfolio Tracker'!K${ptr}`, "0.00%"],
      [`=IFERROR(D${r}/SUM($D$8:$D$17),0)`, "0.00%"],
    ];

    rowData.forEach(([val, fmt], ci) => {
      const cell = ws.getRange(r, ci + 1);
      if (typeof val === "string" && val.startsWith("=")) {
        cell.setFormula(val);
      } else {
        cell.setValue(val);
      }
      cell.setBackground(col.bg)
        .setFontColor("#1F3864").setFontSize(10)
        .setHorizontalAlignment(ci < 3 ? "center" : "right")
        .setVerticalAlignment("middle");
      if (fmt) cell.setNumberFormat(fmt);
      setBorderAll(cell);
    });
  });

  // Totals
  ws.setRowHeight(18, 24);
  ws.getRange("A18:C18").merge()
    .setValue("TOTAL").setBackground(COLORS.headerDark).setFontColor(COLORS.white)
    .setFontWeight("bold").setHorizontalAlignment("center");
  setBorderAll(ws.getRange("A18:C18"));
  [[4,"₹#,##0.00"],[5,"₹#,##0.00"],[6,"₹#,##0.00"],[7,"0.00%"],[8,"0.00%"]].forEach(([col,fmt])=>{
    const cl = columnLetter(col);
    const formula = col===7 ? `=IFERROR((E18-D18)/D18,0)` : `=SUM(${cl}8:${cl}17)`;
    const cell = ws.getRange(18, col);
    cell.setFormula(formula).setBackground(COLORS.headerDark).setFontColor(COLORS.white)
      .setFontWeight("bold").setHorizontalAlignment("right").setNumberFormat(fmt);
    setBorderAll(cell);
  });

  // ── Conditional formatting on Dashboard P&L
  const cf1 = ws.getRange("F8:F17");
  const cf2 = ws.getRange("G8:G17");
  [cf1, cf2].forEach(r => {
    ws.setConditionalFormatRules(ws.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(COLORS.green).setRanges([r]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(COLORS.red).setRanges([r]).build(),
    ]));
  });

  // ── Embedded Pie Chart
  const chartData  = ws.getRange("B8:B17");   // labels
  const chartVals  = ws.getRange("D8:D17");   // values

  const chart = ws.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(chartData)
    .addRange(chartVals)
    .setOption("title", "Portfolio Allocation by Invested Value")
    .setOption("pieHole", 0.4)           // donut style
    .setOption("legend", { position: "right", textStyle: { fontSize: 10 } })
    .setOption("chartArea", { left: 20, top: 40, width: "65%", height: "80%" })
    .setOption("backgroundColor", "#FAFAFA")
    .setPosition(20, 1, 0, 20)          // row 20, col A
    .setNumRows(10)
    .setNumColumns(2)
    .build();
  ws.insertChart(chart);

  // ── Sector Summary (col F, rows 20+)
  ws.setRowHeight(20, 20);
  ws.getRange("F20").setValue("🏭 Sector Summary")
    .setFontWeight("bold").setFontColor(COLORS.headerDark).setFontSize(11);

  const secHeaders = ["Sector","# Stocks","Invested (₹)","Allocation %"];
  ws.getRange(21, 6, 1, 4).setValues([secHeaders])
    .setBackground(COLORS.headerMid).setFontColor(COLORS.white)
    .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  setBorderAll(ws.getRange(21, 6, 1, 4));

  const sectors = [
    { name:"IT",           cnt:3, rows:[8,9,10]    },
    { name:"Banking",      cnt:3, rows:[11,12,13]  },
    { name:"Pharma",       cnt:3, rows:[14,15,16]  },
    { name:"Conglomerate", cnt:1, rows:[17]         },
  ];
  sectors.forEach((s, si) => {
    const r = si + 22;
    const col = COLORS[s.name] || { bg:"#F9F9F9" };
    const sumF = s.rows.map(x => `D${x}`).join("+");
    ws.setRowHeight(r, 22);
    setCellValue(ws, r, 6, s.name, col.bg, "#1F3864", "center");
    setCellValue(ws, r, 7, s.cnt,  col.bg, "#1F3864", "center");
    setCellFormula(ws, r, 8, `=${sumF}`, col.bg, "#000000", "right");
    ws.getRange(r, 8).setNumberFormat("₹#,##0.00");
    setCellFormula(ws, r, 9, `=IFERROR(H${r}/SUM($H$22:$H$25),0)`, col.bg, "#000000", "right");
    ws.getRange(r, 9).setNumberFormat("0.00%");
    [6,7,8,9].forEach(c => setBorderAll(ws.getRange(r, c)));
  });

  ws.setFrozenRows(1);
}

/* ════════════════════════════════════════════
   SHEET 3 — README
════════════════════════════════════════════ */
function buildReadmeSheet(ss) {
  const ws = ss.insertSheet("README", 2);
  ws.setHiddenGridlines(true);
  ws.setColumnWidth(1, 200);
  ws.setColumnWidth(2, 480);

  ws.getRange("A1:B1").merge()
    .setValue("📋  HOW TO USE — Indian Equity Portfolio Tracker v2.0")
    .setBackground(COLORS.headerDark).setFontColor(COLORS.white)
    .setFontSize(12).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 32);

  const lines = [
    ["SECTION_HEADER", "⚡ DAILY WORKFLOW (Prices update automatically!)"],
    ["Step 1", "Open the Portfolio Tracker sheet"],
    ["Step 2", "GOOGLEFINANCE() auto-fetches live NSE prices in Col G — no manual entry needed"],
    ["Step 3", "Check Dashboard for P&L summary and pie chart"],
    ["Force Refresh", "Press Ctrl+Shift+F9 to force recalculate all GOOGLEFINANCE formulas"],
    ["", ""],
    ["SECTION_HEADER", "🖍️ COLOR GUIDE"],
    ["🔵 Blue text",   "Hardcoded inputs (Qty, Buy Price) — edit these to match your actual positions"],
    ["⚫ Black text",  "Auto-calculated formulas — do NOT edit"],
    ["🟢 Green text",  "Live data from GOOGLEFINANCE() — auto-updates"],
    ["🟩 Light Green", "Banking sector rows"],
    ["🔷 Light Blue",  "IT sector rows"],
    ["🟨 Light Yellow","Pharma sector rows"],
    ["🟧 Light Orange","Conglomerate rows"],
    ["", ""],
    ["SECTION_HEADER", "📡 LIVE DATA — GOOGLEFINANCE() FORMULAS USED"],
    ["Current Price",  `=GOOGLEFINANCE("NSE:INFY","price")`],
    ["52-Week High",   `=GOOGLEFINANCE("NSE:INFY","high52")`],
    ["P&L is Green",   "When Current Value > Invested Value"],
    ["P&L is Red",     "When Current Value < Invested Value"],
    ["", ""],
    ["SECTION_HEADER", "🔧 HOW TO ADD A NEW STOCK"],
    ["Step 1", "Add a new row in Portfolio Tracker (copy an existing row)"],
    ["Step 2", "Update: Stock Name, NSE Ticker in Col D, Qty, Buy Price"],
    ["Step 3", "Change the GOOGLEFINANCE ticker in Col G to 'NSE:YOURTICKER'"],
    ["Step 4", "Update Dashboard formulas to include the new row"],
    ["", ""],
    ["SECTION_HEADER", "📚 DATA SOURCES"],
    ["NSE India",      "https://www.nseindia.com — Official NSE data"],
    ["Google Finance", "https://www.google.com/finance — Powers GOOGLEFINANCE()"],
    ["Zerodha Varsity", "https://zerodha.com/varsity — Learn stock market concepts"],
    ["Moneycontrol",   "https://www.moneycontrol.com — News & analysis"],
    ["", ""],
    ["SECTION_HEADER", "⚠️ DISCLAIMER"],
    ["Purpose",      "This tracker is for educational & personal tracking only"],
    ["Not Advice",   "This is NOT financial/investment advice"],
    ["Data Accuracy","GOOGLEFINANCE may have 15-min delay for some data points"],
    ["Verification", "Always verify prices on official NSE/BSE before trading"],
  ];

  let row = 2;
  lines.forEach(([key, val]) => {
    ws.setRowHeight(row, 20);
    if (key === "") { row++; return; }
    if (key === "SECTION_HEADER") {
      ws.getRange(row, 1, 1, 2).merge()
        .setValue(val).setBackground(COLORS.headerMid).setFontColor(COLORS.white)
        .setFontWeight("bold").setFontSize(10)
        .setHorizontalAlignment("left").setVerticalAlignment("middle")
        .setIndent(1);
      setBorderAll(ws.getRange(row, 1, 1, 2));
    } else {
      const kc = ws.getRange(row, 1);
      kc.setValue(key).setFontWeight("bold").setBackground("#F2F2F2")
        .setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      setBorderAll(kc);
      const vc = ws.getRange(row, 2);
      vc.setValue(val).setBackground(COLORS.white)
        .setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      setBorderAll(vc);
    }
    row++;
  });
}

/* ════════════════════════════════════════════
   AUTO-REFRESH TRIGGER
════════════════════════════════════════════ */
function setupAutoRefresh() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "forceRefresh") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("forceRefresh")
    .timeBased().everyMinutes(CONFIG.refreshMinutes).create();
}

function forceRefresh() {
  // Touching a cell forces GOOGLEFINANCE to re-evaluate
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Portfolio Tracker");
  if (!ws) return;
  const stamp = ws.getRange("M1");
  stamp.setValue("Last refresh: " + new Date().toLocaleTimeString("en-IN"));
}

/* ════════════════════════════════════════════
   UTILITY HELPERS
════════════════════════════════════════════ */
function setCellValue(ws, row, col, value, bg, fg, align) {
  const c = ws.getRange(row, col);
  c.setValue(value)
    .setBackground(bg || COLORS.white)
    .setFontColor(fg || "#000000")
    .setFontSize(10)
    .setHorizontalAlignment(align || "left")
    .setVerticalAlignment("middle");
  setBorderAll(c);
  return c;
}

function setCellFormula(ws, row, col, formula, bg, fg, align) {
  const c = ws.getRange(row, col);
  c.setFormula(formula)
    .setBackground(bg || COLORS.white)
    .setFontColor(fg || "#000000")
    .setFontSize(10)
    .setHorizontalAlignment(align || "right")
    .setVerticalAlignment("middle");
  setBorderAll(c);
  return c;
}

function setBorderAll(range) {
  range.setBorder(true, true, true, true, true, true,
    "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
}

function columnLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
