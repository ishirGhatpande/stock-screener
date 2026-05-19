# config.py
# Central configuration for the portfolio optimizer.

# --- Universe of stocks ---
# 15 tickers spanning 5 sectors for diversification
TICKERS = [
    # Technology
    "AAPL", "MSFT", "NVDA",
    # Financials
    "JPM", "BAC", "GS",
    # Energy
    "XOM", "CVX", "SLB",
    # Healthcare
    "JNJ", "UNH", "PFE",
    # Consumer (staples + discretionary)
    "PG", "KO", "AMZN",
]

# --- Historical data window ---
# Number of calendar years of daily adjusted-close prices to fetch
YEARS_OF_DATA = 3

# --- Monte Carlo simulation ---
# Number of random weight combinations to generate
NUM_PORTFOLIOS = 10_000

# --- Risk-free rate ---
# Approximate current annualised 10-year US Treasury yield (as a decimal)
RISK_FREE_RATE = 0.043

# --- Output file names ---
CHART_FILENAME = "efficient_frontier.png"
EXCEL_FILENAME = "portfolio_results.xlsx"
