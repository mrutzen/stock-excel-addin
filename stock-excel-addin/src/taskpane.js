/* global Office, Excel */

// ─── Office.js bootstrap ────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  const btn = document.getElementById("pull-btn");
  const input = document.getElementById("ticker-input");

  btn.addEventListener("click", pullStockData);

  // Allow Enter key to trigger pull
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") pullStockData();
  });

  // Force uppercase as user types
  input.addEventListener("input", (e) => {
    const pos = e.target.selectionStart;
    e.target.value = e.target.value.toUpperCase().replace(/[^A-Z0-9.\-^]/g, "");
    e.target.setSelectionRange(pos, pos);
  });
});

// ─── Helpers ─────────────────────────────────────────────────────────────────

function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = `status ${type}`;
}

/** Get the .raw value from a Yahoo Finance field object, or null. */
function raw(field) {
  if (field === null || field === undefined) return null;
  if (typeof field === "object" && "raw" in field) return field.raw;
  return field;
}

function fmtCurrency(val, symbol = "$") {
  const n = raw(val);
  if (n === null || n === undefined) return "N/A";
  return symbol + parseFloat(n).toFixed(2);
}

function fmtNum(val, decimals = 2) {
  const n = raw(val);
  if (n === null || n === undefined) return "N/A";
  return parseFloat(n).toFixed(decimals);
}

function fmtPct(val) {
  const n = raw(val);
  if (n === null || n === undefined) return "N/A";
  return (parseFloat(n) * 100).toFixed(2) + "%";
}

function fmtLarge(val) {
  const n = raw(val);
  if (n === null || n === undefined) return "N/A";
  const abs = Math.abs(n);
  if (abs >= 1e12) return (n / 1e12).toFixed(2) + "T";
  if (abs >= 1e9) return (n / 1e9).toFixed(2) + "B";
  if (abs >= 1e6) return (n / 1e6).toFixed(2) + "M";
  if (abs >= 1e3) return (n / 1e3).toFixed(2) + "K";
  return parseFloat(n).toFixed(2);
}

function fmtDate(unixTs) {
  return new Date(unixTs * 1000).toISOString().slice(0, 10);
}

function fmtChangeStr(changeRaw, changePctRaw) {
  const c = raw(changeRaw);
  const p = raw(changePctRaw);
  if (c === null) return "N/A";
  const sign = c >= 0 ? "+" : "";
  const pStr = p !== null ? ` (${sign}${(p * 100).toFixed(2)}%)` : "";
  return `${sign}${c.toFixed(2)}${pStr}`;
}

// ─── Main pull function ───────────────────────────────────────────────────────

async function pullStockData() {
  const ticker = document.getElementById("ticker-input").value.trim();
  const range = document.getElementById("range-select").value;

  if (!ticker) {
    showStatus("Please enter a ticker symbol.", "error");
    return;
  }

  const btn = document.getElementById("pull-btn");
  btn.disabled = true;
  btn.textContent = "Loading…";
  showStatus(`Fetching data for ${ticker}…`, "info");

  try {
    const [chartRes, quoteRes] = await Promise.all([
      fetch(`/api/stock/${ticker}?range=${range}`),
      fetch(`/api/quote/${ticker}`),
    ]);

    // Parse JSON regardless of status so we can read error messages
    const chartData = await chartRes.json();
    const quoteData = await quoteRes.json();

    if (!chartRes.ok) throw new Error(chartData.error || `Chart fetch failed (${chartRes.status})`);
    if (!quoteRes.ok) throw new Error(quoteData.error || `Quote fetch failed (${quoteRes.status})`);

    await writeToExcel(ticker, range, chartData, quoteData);
    showStatus(`\u2713 Data for ${ticker} loaded into the sheet.`, "success");
  } catch (err) {
    showStatus(`Error: ${err.message}`, "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "\u21D3\u00A0 Pull Stock Data";
  }
}

// ─── Write data to the active Excel sheet ────────────────────────────────────

async function writeToExcel(ticker, range, chartData, quoteData) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Clear previous data
    try {
      sheet.getUsedRange().clear();
    } catch (_) {
      // Sheet may be empty
    }

    // ── Parse Yahoo Finance payloads ──────────────────────────────

    const qs = quoteData.quoteSummary?.result?.[0];
    if (!qs) throw new Error("No quote data returned. Check the ticker symbol.");

    const priceObj = qs.price || {};
    const sumObj   = qs.summaryDetail || {};
    const finObj   = qs.financialData || {};
    const statObj  = qs.defaultKeyStatistics || {};
    const profObj  = qs.assetProfile || {};

    const chartResult = chartData.chart?.result?.[0];
    const timestamps  = chartResult?.timestamp || [];
    const ohlcv       = chartResult?.indicators?.quote?.[0] || {};

    // ── Build cell value arrays ───────────────────────────────────

    const companyName = raw(priceObj.longName) || raw(priceObj.shortName) || ticker;
    const currency    = raw(priceObj.currency) || "USD";
    const exchange    = raw(priceObj.exchangeName) || "";
    const sector      = raw(profObj.sector) || "";
    const industry    = raw(profObj.industry) || "";

    const currentPrice   = raw(priceObj.regularMarketPrice);
    const priceChange    = raw(priceObj.regularMarketChange);
    const priceChangePct = raw(priceObj.regularMarketChangePercent);
    const prevClose      = raw(priceObj.regularMarketPreviousClose);
    const openPrice      = raw(priceObj.regularMarketOpen);
    const dayHigh        = raw(priceObj.regularMarketDayHigh);
    const dayLow         = raw(priceObj.regularMarketDayLow);
    const volume         = raw(priceObj.regularMarketVolume);
    const marketCap      = raw(priceObj.marketCap);

    // ── Section A: Title ──────────────────────────────────────────
    const titleRows = [
      ["Stock Data:", companyName, "", "Ticker:", ticker],
      ["Exchange:", exchange, "", "Currency:", currency],
      ["Sector:", sector, "", "Industry:", industry],
      ["Pulled at:", new Date().toLocaleString(), "", "", ""],
      ["", "", "", "", ""],
    ];

    // ── Section B: Current Quote ──────────────────────────────────
    const quoteRows = [
      ["CURRENT QUOTE", "", "", "", ""],
      ["Current Price",   currentPrice !== null ? currentPrice : "N/A",   "", "Prev Close", prevClose !== null ? prevClose : "N/A"],
      ["Change",          fmtChangeStr(priceChange, priceChangePct),       "", "Open",       openPrice !== null ? openPrice : "N/A"],
      ["Day High",        dayHigh !== null ? dayHigh : "N/A",              "", "Day Low",    dayLow !== null ? dayLow : "N/A"],
      ["Volume",          volume !== null ? volume : "N/A",                "", "Market Cap", fmtLarge(marketCap)],
      ["", "", "", "", ""],
    ];

    // ── Section C: Fundamentals ───────────────────────────────────
    const fundamentalRows = [
      ["FUNDAMENTALS", "", "", "", ""],
      ["P/E Ratio (TTM)",    fmtNum(sumObj.trailingPE),         "", "Forward P/E",     fmtNum(sumObj.forwardPE)],
      ["EPS (TTM)",          fmtCurrency(statObj.trailingEps),  "", "EPS (Forward)",   fmtCurrency(statObj.forwardEps)],
      ["Revenue (TTM)",      fmtLarge(finObj.totalRevenue),     "", "Gross Profit",    fmtLarge(finObj.grossProfits)],
      ["Profit Margin",      fmtPct(finObj.profitMargins),      "", "Operating Margin",fmtPct(finObj.operatingMargins)],
      ["Return on Equity",   fmtPct(finObj.returnOnEquity),     "", "Return on Assets",fmtPct(finObj.returnOnAssets)],
      ["Debt / Equity",      fmtNum(finObj.debtToEquity),       "", "Current Ratio",   fmtNum(finObj.currentRatio)],
      ["Free Cash Flow",     fmtLarge(finObj.freeCashflow),     "", "Op. Cash Flow",   fmtLarge(finObj.operatingCashflow)],
      ["", "", "", "", ""],
    ];

    // ── Section D: Valuation & dividends ─────────────────────────
    const valuationRows = [
      ["VALUATION & DIVIDENDS", "", "", "", ""],
      ["52-Week High",    fmtCurrency(sumObj.fiftyTwoWeekHigh),  "", "52-Week Low",   fmtCurrency(sumObj.fiftyTwoWeekLow)],
      ["50-Day MA",       fmtCurrency(priceObj.fiftyDayAverage), "", "200-Day MA",    fmtCurrency(priceObj.twoHundredDayAverage)],
      ["Beta",            fmtNum(sumObj.beta),                   "", "Short Ratio",   fmtNum(statObj.shortRatio)],
      ["Shares Out.",     fmtLarge(statObj.sharesOutstanding),   "", "Float Shares",  fmtLarge(statObj.floatShares)],
      ["Dividend Rate",   fmtCurrency(sumObj.dividendRate),      "", "Dividend Yield",fmtPct(sumObj.dividendYield)],
      ["Ex-Div Date",     raw(sumObj.exDividendDate) ? fmtDate(raw(sumObj.exDividendDate)) : "N/A",
                                                                 "", "Payout Ratio",  fmtPct(sumObj.payoutRatio)],
      ["Book Value/Share",fmtCurrency(statObj.bookValue),        "", "Price/Book",    fmtNum(statObj.priceToBook)],
      ["", "", "", "", ""],
    ];

    // ── Section E: Historical OHLCV ───────────────────────────────
    const histHeaderRows = [
      [`HISTORICAL PRICES (${range})`, "", "", "", ""],
      ["Date", "Open", "High", "Low", "Close", "Volume"],
    ];

    const histDataRows = timestamps.map((ts, i) => {
      const safeVal = (arr) => (arr && arr[i] !== null && arr[i] !== undefined ? parseFloat(arr[i].toFixed(4)) : "N/A");
      return [
        fmtDate(ts),
        safeVal(ohlcv.open),
        safeVal(ohlcv.high),
        safeVal(ohlcv.low),
        safeVal(ohlcv.close),
        ohlcv.volume?.[i] ?? "N/A",
      ];
    });

    // ── Assemble all rows ─────────────────────────────────────────
    const allRows = [
      ...titleRows,
      ...quoteRows,
      ...fundamentalRows,
      ...valuationRows,
      ...histHeaderRows,
      ...histDataRows,
    ];

    const totalRows = allRows.length;
    const totalCols = 6; // A–F

    // Write values in one API call (much faster than row-by-row)
    // Pad each row to exactly 6 columns
    const padded = allRows.map((row) => {
      const r = [...row];
      while (r.length < totalCols) r.push("");
      return r.slice(0, totalCols);
    });

    const dataRange = sheet.getRangeByIndexes(0, 0, totalRows, totalCols);
    dataRange.values = padded;

    // ── Formatting ────────────────────────────────────────────────

    // Title row: dark blue header spanning A1:F1
    const titleRange = sheet.getRangeByIndexes(0, 0, 1, totalCols);
    titleRange.format.font.bold = true;
    titleRange.format.font.size = 13;
    titleRange.format.font.color = "#ffffff";
    titleRange.format.fill.color = "#0078d4";

    // Metadata rows 2-4 (light blue tint)
    const metaRange = sheet.getRangeByIndexes(1, 0, 3, totalCols);
    metaRange.format.fill.color = "#ddeeff";

    // Helper: style a section header row
    function styleSection(rowIndex) {
      const r = sheet.getRangeByIndexes(rowIndex, 0, 1, totalCols);
      r.format.font.bold = true;
      r.format.font.color = "#ffffff";
      r.format.fill.color = "#106ebe";
    }

    // Calculate section header positions
    const quoteHeaderRow      = titleRows.length;
    const fundamentalHeaderRow = quoteHeaderRow + quoteRows.length;
    const valuationHeaderRow  = fundamentalHeaderRow + fundamentalRows.length;
    const histSectionRow      = valuationHeaderRow + valuationRows.length;
    const histColHeaderRow    = histSectionRow + 1;

    styleSection(quoteHeaderRow);
    styleSection(fundamentalHeaderRow);
    styleSection(valuationHeaderRow);
    styleSection(histSectionRow);

    // Historical column headers
    const colHdrRange = sheet.getRangeByIndexes(histColHeaderRow, 0, 1, totalCols);
    colHdrRange.format.font.bold = true;
    colHdrRange.format.fill.color = "#f3f2f1";

    // Label columns (A and D) bold in detail sections
    const labelCols = [0, 3];
    const detailStart = titleRows.length + 1; // skip section header
    const detailEnd   = histSectionRow - 1;
    for (const col of labelCols) {
      const r = sheet.getRangeByIndexes(detailStart, col, detailEnd - detailStart, 1);
      r.format.font.bold = true;
    }

    // Auto-fit all columns
    sheet.getUsedRange().format.autofitColumns();

    // Freeze the top row
    sheet.freezePanes.freezeRows(1);

    await context.sync();
  });
}
