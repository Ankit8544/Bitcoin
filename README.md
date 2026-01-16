<div align="center">
    <img src="Img/Picture1.png" width="420"/>
    <p>
        <strong>Enterprise-Grade Real-Time Excel‑Based Bitcoin Market Intelligence, Risk & MIS Platform</strong>
    </p>
</div>
<artifact identifier="btc-mis-documentation" type="text/markdown" title="Bitcoin Real-Time MIS & Risk Analysis - Complete Documentation">
<div align="center">

![Bitcoin](https://img.shields.io/badge/Bitcoin-BTC-orange?style=for-the-badge\&logo=bitcoin)
![Binance](https://img.shields.io/badge/Binance-API-yellow?style=for-the-badge\&logo=binance)
![Binance](https://img.shields.io/badge/Binance-WebSocket-yellow?style=for-the-badge\&logo=binance)
![XlOil](https://img.shields.io/badge/XlOil-Data%20Stream-blue?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.11+-blue?style=for-the-badge\&logo=python\&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-Analytics-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-red?style=for-the-badge\&logo=microsoft)

*Comprehensive Risk Analytics | Live Market Data | Institutional-Quality MIS Reporting | Advanced Visualization*

[Features](#-key-features) • [Installation](#-installation) • [Documentation](#-documentation) • [Architecture](#-system-architecture)

</div>
---

## 📌 Executive Summary

This project is an **enterprise‑grade, Excel‑native Bitcoin Market Intelligence, Risk & MIS platform** designed to deliver **real‑time market visibility**, **portfolio‑level risk analytics**, and **institutional reporting** without requiring external BI tools.

The system integrates **Binance REST + WebSocket APIs**, **xlOil streaming functions**, **Python async engines**, and **advanced Excel modeling** to create a **live, auditable, and extensible Bitcoin analytics stack** suitable for:

* Portfolio tracking
* Risk monitoring
* Trade behavior analysis
* Management Information System (MIS) reporting

---

## 🎯 Project Objectives

* 🔴 Real‑time Bitcoin market monitoring inside Excel
* 📊 Professional‑grade portfolio & risk reporting
* 🧮 Quantitative risk metrics (drawdown, volatility, Sharpe, CAGR)
* 🧠 Trade‑level and microstructure insights
* 🏦 Institutional‑style MIS dashboards
* 🔐 Strong data governance & fault‑tolerant architecture

---

## 🚀 Key Features

### 📡 Live Market Intelligence

* 24‑Hour Rolling Ticker (price, volume, volatility)
* Multi‑timeframe OHLC (1m → 1d)
* Aggregate trade flow analytics
* All‑market coin scanner

### 💼 Portfolio & Asset Management

* Secure Bitcoin asset entry
* Real‑time valuation & P&L
* Holding‑period analytics
* Risk‑adjusted performance metrics

### ⚠️ Risk & Behavior Analytics

* Max Drawdown & recovery analysis
* Annualized volatility
* Sharpe Ratio
* Win/Loss behavior & distribution

### 📑 MIS & Reporting

* Executive Overview Dashboard
* Dedicated Portfolio Report
* Trade & Market Microstructure Report
* Data‑driven insights & alerts

---

## 🧱 System Architecture

```
Binance API (REST + WebSocket)
        ↓
Python Async Engines (xlOil)
        ↓
Excel Streaming Sheets (Raw Data)
        ↓
Data Transformation Layer
        ↓
Risk Models & Metrics Engine
        ↓
Dashboards & MIS Reports
```

---

## 🗂️ Data Architecture & Implementation

### 1️⃣ 24‑Hour Rolling Ticker Sheet (`24h Ticker`)

**Source:** Binance WebSocket `@ticker`

| Field               | xlOil Formula                      | Description               |
| ------------------- | ---------------------------------- | ------------------------- |
| Event Time          | `=TickerStream("BTCUSDT","E")`     | Event timestamp (IST)     |
| Symbol              | `=TickerStream("BTCUSDT","s")`     | Trading pair              |
| Price Change        | `=TickerStream("BTCUSDT","p")`     | Absolute 24h price change |
| Price Change %      | `=TickerStream("BTCUSDT","P")/100` | Normalized percentage     |
| Weighted Avg Price  | `=TickerStream("BTCUSDT","w")`     | VWAP                      |
| Last Price          | `=TickerStream("BTCUSDT","c")`     | Latest traded price       |
| Last Quantity       | `=TickerStream("BTCUSDT","Q")`     | Last trade size           |
| Open Price          | `=TickerStream("BTCUSDT","o")`     | 24h open                  |
| High / Low          | `h / l`                            | Intraday range            |
| Base / Quote Volume | `v / q`                            | Liquidity metrics         |
| Trade Count         | `n`                                | Market activity           |

**Usage:**

* Intraday volatility monitoring
* Market regime classification
* Executive price snapshot

---

### 2️⃣ OHLC Market Data (Multi‑Timeframe)

**Source:** Binance REST + WebSocket klines

| Sheet            | Formula                             | Purpose                 |
| ---------------- | ----------------------------------- | ----------------------- |
| `1m`             | `=KlineStream("BTCUSDT","1m",61)`   | Microstructure analysis |
| `15m`            | `=KlineStream("BTCUSDT","15m",500)` | Short‑term trends       |
| `1h`             | `=KlineStream("BTCUSDT","1h",500)`  | Swing structure         |
| `4h`             | `=KlineStream("BTCUSDT","4h",300)`  | Market regimes          |
| `Holding Period` | `=KlineStream("BTCUSDT","1d",Days)` | Portfolio analytics     |
| `1d`             | Dynamic OFFSET logic                | Rolling daily history   |

**Captured Metrics:**

* OHLC prices
* Volume & quote volume
* Number of trades
* Taker buy/sell pressure

---

### 3️⃣ Aggregate Trade Streams (`AT_*`)

**Source:** Binance `@aggTrade`

| Sheet  | Formula                                 | Window               |
| ------ | --------------------------------------- | -------------------- |
| AT_1m  | `=AggTradeStreamWindow("BTCUSDT",1)`    | Order flow           |
| AT_5m  | `=AggTradeStreamWindow("BTCUSDT",5)`    | Momentum             |
| AT_15m | `=AggTradeStreamWindow("BTCUSDT",15)`   | Intraday behavior    |
| AT_1h  | `=AggTradeStreamWindow("BTCUSDT",60)`   | Market participation |
| AT_4h  | `=AggTradeStreamWindow("BTCUSDT",240)`  | Institutional flow   |
| AT_1d  | `=AggTradeStreamWindow("BTCUSDT",1440)` | Daily structure      |

**Captured Fields:**

* Trade time (IST)
* Price & quantity
* AggTrade IDs
* Buyer/Seller aggressor flag

---

### 4️⃣ All‑Market Scanner (`All Coins`)

**Formula:** `=AllCoinsTickerStream()`

**Purpose:**

* Cross‑market comparison
* Correlation screening
* Market heatmap generation

---

### 5️⃣ Comparative Asset Analysis

**Sheet:** `Comparing Asset`

```
=KlineStream($F$1,"1h",500)
```

Used for:

* BTC vs Altcoin correlation
* Risk diversification analysis
* Relative strength modeling

---

### 6️⃣ Portfolio & Asset Data (`Assets`)

**User‑Entered Fields:**

* Quantity (BTC)
* Buy Date
* Buy Price
* Invested Amount

**Derived Metrics:**

* Current Value
* Absolute & % P&L
* Holding Days
* CAGR‑style returns

---

## ⚠️ Risk & Performance Metrics Engine

| Metric         | Description               |
| -------------- | ------------------------- |
| Max Drawdown   | Peak‑to‑trough loss       |
| Volatility     | Annualized std. deviation |
| Sharpe Ratio   | Risk‑adjusted return      |
| CAGR           | Annualized performance    |
| Win Rate       | % profitable days         |
| Best/Worst Day | Tail risk analysis        |

---

## 🛡️ Data Governance & Reliability

### 🔐 Data Integrity

* Immutable raw data sheets
* Clear separation: Raw → Model → Report

### 🔄 Fault Tolerance

* WebSocket auto‑reconnect
* REST backfill on reconnect
* Last‑snapshot freeze (no Excel errors)

### 🕒 Time Standardization

* UTC → IST conversion at ingestion
* Consistent timestamps across sheets

### 📜 Auditability

* Transparent Excel formulas
* Deterministic calculations
* Reproducible metrics

---

## 📊 Dashboards & Reports

### 🧭 Overview Dashboard

* Market snapshot
* Portfolio value
* Risk indicators

### 💼 Portfolio Report

* Holdings summary
* Risk & performance metrics
* Drawdown visualization

### 🔁 Trade Report

* Aggregate trade behavior
* Buy/Sell pressure
* Volume clusters

### 🧠 Insights Report

* Market regime classification
* Volatility alerts
* Risk concentration signals

---

## 🛣️ Roadmap (Planned Enhancements)

* 🔮 Predictive volatility models
* 📈 Value‑at‑Risk (VaR / CVaR)
* 🤖 Signal‑based trade analytics
* ☁️ Cloud backup & versioning
* 🧾 Export‑ready institutional reports

---

## ✅ Conclusion

This project establishes a **professional‑grade, Excel‑native Bitcoin analytics platform** that bridges the gap between **retail dashboards** and **institutional risk systems**, delivering **live data, quantitative rigor, and executive‑ready MIS reporting** — all within a transparent and governed Excel environment.

---

🧠 *Designed for analysts. Built for decision‑makers. Engineered for real‑time intelligence.*
