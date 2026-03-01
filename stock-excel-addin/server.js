const express = require("express");
const axios = require("axios");
const path = require("path");
const fs = require("fs");
const https = require("https");
const os = require("os");

const app = express();
const PORT = 3000;

// Serve static files from /src
app.use(express.static(path.join(__dirname, "src")));

// ─── Yahoo Finance: historical chart data ────────────────────────────────────
app.get("/api/stock/:ticker", async (req, res) => {
  const { ticker } = req.params;
  const range = req.query.range || "1mo";
  const interval = range === "1d" ? "5m" : "1d";

  try {
    const url =
      `https://query1.finance.yahoo.com/v8/finance/chart/` +
      `${encodeURIComponent(ticker)}?interval=${interval}&range=${range}`;

    const response = await axios.get(url, {
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
          "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        Accept: "application/json",
      },
      timeout: 10000,
    });
    res.json(response.data);
  } catch (err) {
    const status = err.response?.status || 500;
    const message =
      status === 404
        ? `Ticker "${ticker}" not found.`
        : err.message;
    res.status(status).json({ error: message });
  }
});

// ─── Yahoo Finance: fundamentals / quote summary ─────────────────────────────
app.get("/api/quote/:ticker", async (req, res) => {
  const { ticker } = req.params;
  const modules =
    "summaryDetail,price,financialData,defaultKeyStatistics,assetProfile";

  try {
    const url =
      `https://query1.finance.yahoo.com/v10/finance/quoteSummary/` +
      `${encodeURIComponent(ticker)}?modules=${modules}`;

    const response = await axios.get(url, {
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
          "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        Accept: "application/json",
      },
      timeout: 10000,
    });
    res.json(response.data);
  } catch (err) {
    const status = err.response?.status || 500;
    const message =
      status === 404
        ? `Ticker "${ticker}" not found.`
        : err.message;
    res.status(status).json({ error: message });
  }
});

// ─── Start server (HTTPS with dev certs, or HTTP fallback) ───────────────────
function tryHttps() {
  const certBase = path.join(os.homedir(), ".office-addin-dev-certs");
  const keyPath = path.join(certBase, "localhost.key");
  const certPath = path.join(certBase, "localhost.crt");

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) return false;

  try {
    const options = {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    };
    https.createServer(options, app).listen(PORT, () => {
      console.log(`\n✓ HTTPS server running at https://localhost:${PORT}`);
      console.log(`  Sideload manifest.xml in Excel to use the add-in.\n`);
    });
    return true;
  } catch {
    return false;
  }
}

if (!tryHttps()) {
  app.listen(PORT, () => {
    console.log(`\n⚠  HTTP server running at http://localhost:${PORT}`);
    console.log(`   Excel requires HTTPS. Run: npm run setup  (installs dev certs)\n`);
  });
}
