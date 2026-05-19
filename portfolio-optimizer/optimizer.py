# optimizer.py
# Monte Carlo simulation of the efficient frontier.

import numpy as np
import pandas as pd

from config import NUM_PORTFOLIOS, RISK_FREE_RATE


def _portfolio_performance(
    weights: np.ndarray,
    expected_returns: pd.Series,
    cov_matrix: pd.DataFrame,
) -> tuple[float, float, float]:
    """Return (annualised_return, annualised_volatility, sharpe_ratio) for a weight vector."""
    ret = float(np.dot(weights, expected_returns))
    vol = float(np.sqrt(weights @ cov_matrix.values @ weights))
    sharpe = (ret - RISK_FREE_RATE) / vol if vol > 0 else 0.0
    return ret, vol, sharpe


def simulate_portfolios(
    expected_returns: pd.Series,
    cov_matrix: pd.DataFrame,
    num_portfolios: int = NUM_PORTFOLIOS,
    seed: int = 42,
) -> pd.DataFrame:
    """
    Generate `num_portfolios` random long-only portfolios via Monte Carlo.

    Each portfolio's weights are drawn from a uniform Dirichlet distribution
    so that they sum to 1 and are all non-negative.

    Returns a DataFrame with columns:
        Return, Volatility, Sharpe  – portfolio-level metrics
        <ticker> …                  – weight for each asset
    """
    rng = np.random.default_rng(seed)
    n_assets = len(expected_returns)
    tickers = expected_returns.index.tolist()

    records = []
    for _ in range(num_portfolios):
        # Dirichlet sample ensures weights >= 0 and sum == 1
        weights = rng.dirichlet(np.ones(n_assets))
        ret, vol, sharpe = _portfolio_performance(weights, expected_returns, cov_matrix)
        records.append([ret, vol, sharpe, *weights])

    columns = ["Return", "Volatility", "Sharpe"] + tickers
    return pd.DataFrame(records, columns=columns)


def find_optimal_portfolios(simulation: pd.DataFrame) -> dict:
    """
    Identify the two key portfolios on the efficient frontier.

    Returns a dict with keys:
        max_sharpe  – row (Series) of the highest-Sharpe portfolio
        min_vol     – row (Series) of the lowest-volatility portfolio
    """
    max_sharpe_idx = simulation["Sharpe"].idxmax()
    min_vol_idx = simulation["Volatility"].idxmin()

    return {
        "max_sharpe": simulation.loc[max_sharpe_idx],
        "min_vol": simulation.loc[min_vol_idx],
    }
