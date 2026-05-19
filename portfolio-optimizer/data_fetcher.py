# data_fetcher.py
# Fetches historical price data and derives return statistics.

import datetime
import warnings

import numpy as np
import pandas as pd
import yfinance as yf

from config import TICKERS, YEARS_OF_DATA

TRADING_DAYS = 252  # standard annualisation factor


def _date_range():
    """Return (start, end) as ISO-format strings covering YEARS_OF_DATA."""
    end = datetime.date.today()
    start = end.replace(year=end.year - YEARS_OF_DATA)
    return start.isoformat(), end.isoformat()


def fetch_prices(tickers: list[str] = TICKERS) -> pd.DataFrame:
    """
    Download daily adjusted closing prices for every ticker.

    Returns a DataFrame with dates as the index and tickers as columns.
    Tickers with insufficient data are dropped with a warning.
    """
    start, end = _date_range()
    print(f"Fetching price data from {start} to {end} …")

    # auto_adjust=True returns the split/dividend-adjusted close in "Close"
    raw = yf.download(
        tickers,
        start=start,
        end=end,
        auto_adjust=True,
        progress=False,
    )["Close"]

    # yfinance may return a MultiIndex if only one ticker is requested
    if isinstance(raw, pd.Series):
        raw = raw.to_frame()

    # Drop columns that are entirely NaN (delistings, bad tickers, etc.)
    prices = raw.dropna(axis=1, how="all")

    dropped = set(tickers) - set(prices.columns)
    if dropped:
        warnings.warn(f"Dropped tickers with no data: {sorted(dropped)}")

    # Forward-fill then drop any remaining NaN rows
    prices = prices.ffill().dropna()

    print(f"Loaded {len(prices)} trading days for {len(prices.columns)} tickers.")
    return prices


def compute_statistics(prices: pd.DataFrame) -> dict:
    """
    Derive return-related statistics from a price DataFrame.

    Returns a dict with:
        daily_returns      – DataFrame of daily log returns
        expected_returns   – Series of annualised mean returns per ticker
        volatilities       – Series of annualised std-dev per ticker
        cov_matrix         – DataFrame of annualised covariance matrix
    """
    # Log returns are additive over time and normally distributed
    daily_returns = np.log(prices / prices.shift(1)).dropna()

    expected_returns = daily_returns.mean() * TRADING_DAYS
    volatilities = daily_returns.std() * np.sqrt(TRADING_DAYS)
    cov_matrix = daily_returns.cov() * TRADING_DAYS

    return {
        "daily_returns": daily_returns,
        "expected_returns": expected_returns,
        "volatilities": volatilities,
        "cov_matrix": cov_matrix,
    }
