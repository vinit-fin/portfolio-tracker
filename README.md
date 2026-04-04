📊 Indian Equity Portfolio Tracker

A **Google Apps Script** project that builds a fully-formatted Google Sheets portfolio tracker with **live NSE prices** via `GOOGLEFINANCE()` — no APIs, no paid tools, no manual updates.

Built as part of my finance portfolio to demonstrate real-world use of Google Workspace automation and equity market concepts.

---

## 🚀 Features

- **Live prices** auto-fetched from NSE via `GOOGLEFINANCE()`
- **3 sheets** — Portfolio Tracker, Dashboard, README guide
- **Auto-refresh** every 5 minutes via a time-based trigger
- **Color-coded by sector** — IT, Banking, Pharma, Conglomerate
- **P&L tracking** — both ₹ and % with green/red conditional formatting
- **52-week high** displayed alongside each stock
- **Donut pie chart** for portfolio allocation by invested value
- **Sector summary table** with allocation percentages
- **KPI cards** on the Dashboard — Total Invested, Current Value, P&L ₹, P&L %

---

## 📸 Preview

| Portfolio Tracker | Dashboard |
|:-:|:-:|
| Live prices, P&L per stock, color-coded rows | KPI cards, allocation chart, sector summary |

> *Screenshots to be added after first run*
<img width="1893" height="536" alt="image" src="https://github.com/user-attachments/assets/1d206e12-3558-4af8-90ff-10a0dcad6668" />

---

## 🗂️ Project Structure

```
indian-equity-portfolio-tracker/
├── portfolio_tracker.gs   ← Main script (paste into Apps Script editor)
├── README.md
└── LICENSE
```

---

## ⚡ How to Run

**Step 1** — Open [script.google.com](https://script.google.com) and create a New Project

**Step 2** — Delete the default `myFunction()` code, then paste the full contents of `portfolio_tracker.gs`

**Step 3** — Save the file (`Ctrl + S`), then select the `buildPortfolio` function from the dropdown

**Step 4** — Click ▶ **Run** and approve the permissions popup

**Step 5** — Switch to your Google Sheet — everything is built and live!

> 💡 **Force a price refresh anytime:** `Ctrl + Shift + F9` inside the Sheet

---

## 🖍️ Color Legend

| Color | Meaning |
|-------|---------|
| 🔵 Blue text | User-editable inputs (Qty, Buy Price) |
| ⚫ Black text | Auto-calculated — do not edit |
| 🟢 Green text | Live data from `GOOGLEFINANCE()` |
| 🟩 Light green rows | Banking sector |
| 🔷 Light blue rows | IT sector |
| 🟨 Light yellow rows | Pharma sector |
| 🟧 Light orange rows | Conglomerate sector |

---

## 📁 Sheet Overview

### Sheet 1 — Portfolio Tracker
| Column | Contents |
|--------|----------|
| A | Serial # |
| B | Stock Name |
| C | Sector |
| D | NSE Ticker |
| E | Quantity *(editable)* |
| F | Buy Price *(editable)* |
| G | Live Price via `GOOGLEFINANCE()` |
| H | Invested Value = E × F |
| I | Current Value = E × G |
| J | P&L in ₹ |
| K | P&L in % |
| L | 52-Week High via `GOOGLEFINANCE()` |

### Sheet 2 — Dashboard
- 4 KPI cards: Total Invested, Current Value, P&L ₹, P&L %
- Full stock-wise breakdown table with allocation %
- Donut pie chart for visual allocation
- Sector summary (IT, Banking, Pharma, Conglomerate)

### Sheet 3 — README
- In-sheet how-to guide for non-technical users

---

## 🛠️ Customising Your Holdings

To change stocks, edit the `PORTFOLIO.stocks` array at the top of the script:

```javascript
stocks: [
  // [ Stock Name, NSE Ticker, Sector, Quantity, Buy Price ]
  [ "Infosys Ltd", "INFY", "IT", 10, 1750 ],
  // add your own rows here...
]
```

To add a new sector, add a color entry to the `COLORS` object:

```javascript
const COLORS = {
  // existing sectors...
  FMCG: { bg: "#EAD1DC", fg: "#4A0028" },
};
```

---

## 📡 GOOGLEFINANCE Formulas Used

```
=GOOGLEFINANCE("NSE:INFY", "price")     ← Live price
=GOOGLEFINANCE("NSE:INFY", "high52")    ← 52-week high
```

> ⚠️ `GOOGLEFINANCE()` may carry a **15-minute delay** for some data points. Always verify prices on [NSE India](https://www.nseindia.com) or [BSE](https://www.bseindia.com) before making any trading decisions.

---

## 📚 Resources

- [NSE India](https://www.nseindia.com) — Official exchange data
- [Zerodha Varsity](https://zerodha.com/varsity) — Stock market learning
- [Moneycontrol](https://www.moneycontrol.com) — News & analysis
- [Google Apps Script Docs](https://developers.google.com/apps-script)

---

## ⚠️ Disclaimer

This project is built for **educational and personal tracking purposes only**.  
It is **not financial or investment advice**.  
Always consult a SEBI-registered advisor before making investment decisions.

---

## 👤 Author

**Vinit** — BCA Student, Amity University Online  
Interests: Stock markets, F&O trading, Financial analysis  
GitHub: [vinit-fin](https://github.com/vinit-fin)

---
