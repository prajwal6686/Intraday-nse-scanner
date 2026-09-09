# 📈 Intraday NSE Momentum Scanner

A Python-based quantitative screening tool for identifying and ranking potential intraday trading setups across a selected universe of liquid NSE-listed stocks.

The application combines **price action, VWAP, EMA trends, ATR-based volatility, relative volume, opening-range levels and risk-based position sizing** to generate a ranked trade plan with entry, stop-loss, targets and risk/reward metrics.

🔗 **Live Application:** https://intraday-nse-scanner.streamlit.app/

---

## 🚀 Overview

The scanner is designed to reduce the manual work involved in screening a large number of NSE stocks for intraday momentum opportunities.

Instead of manually checking individual stocks, the system:

1. Downloads market data
2. Filters for liquid and sufficiently volatile stocks
3. Calculates technical and market-structure indicators
4. Scores potential BUY and SELL setups
5. Ranks the strongest candidates
6. Calculates entry, stop-loss and profit targets
7. Calculates position size based on predefined risk
8. Exports the resulting trade plan to Excel or CSV

The project is intended as a **research and decision-support tool**, not an automated trading system.

---

## ⚙️ Key Features

### Stock Screening

* Screens a predefined universe of liquid NSE stocks
* Minimum price and turnover filters
* ATR-based volatility filtering
* Sector classification
* Historical trend analysis

### Technical Indicators

The scanner incorporates:

* **VWAP**
* **20-day EMA**
* **50-day EMA**
* **ATR (Average True Range)**
* **Relative Volume (RVOL)**
* Opening Range High / Low
* Gap percentage
* Intraday range utilization
* Price relative to VWAP
* 52-week high distance

### Momentum Scoring

Stocks are evaluated using a scoring framework based on:

* Price relative to VWAP
* Opening-range breakout/breakdown
* EMA trend alignment
* Relative volume
* Gap behaviour
* Intraday range exhaustion

The strongest setups are ranked and assigned a confidence rating.

### Risk Management

The system does not simply generate a stock signal.

It also calculates:

* Entry price
* Stop-loss
* Target 1
* Target 2
* Position quantity
* Capital requirement
* Margin requirement
* Rupee risk per trade
* Expected reward
* Risk/reward ratio

Position sizing is linked to the user's selected **capital and risk-per-trade percentage**.

### Export

Results can be downloaded as:

* **Excel workbook**
* **CSV file**

The Excel output contains multiple sheets covering rankings, trade plans, reasons, summary information and underlying data.

---

## 🧠 Methodology

### 1. Liquidity & Volatility Filter

Stocks are first filtered using minimum price, average turnover and ATR-based volatility conditions.

This reduces the likelihood of producing setups in stocks with insufficient liquidity or unsuitable volatility.

### 2. Intraday Data

The system uses intraday OHLCV data and converts timestamps to **Asia/Kolkata (IST)** for session analysis.

The scanner evaluates price and volume behaviour up to a configurable signal cutoff time.

### 3. VWAP & Opening Range

The scanner evaluates whether price is trading above or below VWAP and whether it has broken the opening-range high or low.

A combination such as:

**Price > VWAP + Opening Range High Breakout**

contributes strongly toward a bullish score.

Similarly:

**Price < VWAP + Opening Range Low Breakdown**

supports a bearish score.

### 4. Trend Confirmation

20-day and 50-day EMA relationships are used as additional trend filters.

This helps distinguish momentum aligned with the broader price trend from weaker setups.

### 5. Relative Volume

Current intraday volume is compared with historical intraday volume at comparable points in previous sessions.

Higher RVOL indicates stronger-than-normal participation.

### 6. Trade Construction

Once a direction is selected, the system constructs a trade plan consisting of:

**Entry → Stop Loss → Target 1 → Target 2**

Target levels are derived using predefined risk multiples.

### 7. Position Sizing

Position size is calculated from the maximum rupee risk permitted by the selected risk percentage.

Conceptually:

`Risk Budget = Capital × Risk %`

`Position Size = Risk Budget ÷ Risk Per Share`

The resulting quantity is also constrained by available buying power/margin assumptions.

---

## 🖥️ Application Inputs

Users can configure:

| Parameter        | Description                                 |
| ---------------- | ------------------------------------------- |
| Capital          | Trading capital used for position sizing    |
| Leverage         | Assumed MIS leverage                        |
| Risk per Trade   | Maximum capital risk allocated per setup    |
| Number of Stocks | Number of ranked setups returned            |
| Signal Cutoff    | Time up to which intraday data is evaluated |
| Allow Shorts     | Enables/disables SELL setups                |

---

## 📊 Example Output

Each ranked setup can contain:

| Metric     | Description                            |
| ---------- | -------------------------------------- |
| Rank       | Relative ranking among screened stocks |
| Symbol     | NSE ticker                             |
| Direction  | BUY / SELL                             |
| Score      | Quantitative setup score               |
| Confidence | Confidence rating                      |
| Entry      | Proposed entry level                   |
| Stop       | Calculated stop-loss                   |
| Target 1   | First profit objective                 |
| Target 2   | Second profit objective                |
| Quantity   | Risk-adjusted position size            |
| Margin     | Estimated margin requirement           |
| Risk       | Estimated rupee risk                   |
| Reward     | Estimated reward                       |
| R:R        | Risk/reward ratio                      |
| RVOL       | Relative volume                        |
| ATR %      | Volatility measure                     |
| Gap %      | Opening gap                            |
| Trend      | Trend classification                   |

---

## 🏗️ Project Structure

```text
Intraday-nse-scanner/
│
├── app.py              # Streamlit user interface
├── scanner_core.py     # Scanning, scoring, risk and export engine
├── requirements.txt    # Python dependencies
└── README.md           # Project documentation
```

### Architecture

```text
Market Data
     ↓
Data Cleaning & Session Handling
     ↓
Liquidity / Volatility Filters
     ↓
Technical Indicators
     ↓
Momentum & Trend Scoring
     ↓
BUY / SELL Classification
     ↓
Entry / Stop / Target Construction
     ↓
Risk-Based Position Sizing
     ↓
Ranked Trade Plan
     ↓
Excel / CSV Export
```

---

## 🛠️ Technology Stack

* **Python**
* **Streamlit**
* **Pandas**
* **NumPy**
* **yfinance**
* **OpenPyXL**

---

## ▶️ Run Locally

Clone the repository:

```bash
git clone https://github.com/prajwal6686/Intraday-nse-scanner.git
cd Intraday-nse-scanner
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the Streamlit application:

```bash
streamlit run app.py
```

The application will then open in your browser.

---

## 🌐 Live Demo

**Try the application:**

https://intraday-nse-scanner.streamlit.app/

---

## ⚠️ Disclaimer

This project is intended for **research, educational and analytical purposes only**.

The scanner generates quantitative trade setups based on historical and intraday market data. It does not guarantee profitability or future performance.

The output should not be interpreted as investment advice or a recommendation to buy or sell securities.

---

## 🔮 Future Improvements

Potential extensions include:

* Backtesting framework
* Historical signal performance analysis
* Walk-forward testing
* Strategy parameter optimization
* Market-regime detection
* Sector-relative momentum
* More robust data providers
* Real-time market data integration
* Signal persistence and trade journaling
* Performance analytics and win-rate statistics
* Automated alerts
* Portfolio-level risk management

---

## 👤 Author

**Prajwal Aughade**

Built as a personal quantitative finance project exploring the intersection of:

**Financial Markets × Quantitative Analysis × Python**

GitHub: https://github.com/prajwal6686

---

⭐ If you find the project useful, consider starring the repository.
