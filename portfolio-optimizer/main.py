# main.py
# Entry point – runs the full portfolio optimisation pipeline.

import sys

from config import TICKERS
from data_fetcher import fetch_prices, compute_statistics
from optimizer import simulate_portfolios, find_optimal_portfolios
from visualizer import plot_efficient_frontier
from excel_output import write_excel


def main() -> None:
    print("=" * 60)
    print("  Modern Portfolio Theory — Efficient Frontier Optimizer")
    print("=" * 60)

    # 1. Download price data and compute return statistics
    prices = fetch_prices(TICKERS)
    stats = compute_statistics(prices)

    expected_returns = stats["expected_returns"]
    volatilities     = stats["volatilities"]
    cov_matrix       = stats["cov_matrix"]
    daily_returns    = stats["daily_returns"]

    # Use only tickers that survived the data fetch
    live_tickers = prices.columns.tolist()

    # 2. Monte Carlo simulation of random portfolios
    print(f"\nSimulating portfolios …")
    simulation = simulate_portfolios(expected_returns, cov_matrix)
    print(f"Generated {len(simulation):,} portfolios.")

    # 3. Identify optimal portfolios
    optimal = find_optimal_portfolios(simulation)

    ms = optimal["max_sharpe"]
    mv = optimal["min_vol"]

    print("\n--- Max Sharpe Portfolio ---")
    print(f"  Return     : {ms['Return']*100:.2f}%")
    print(f"  Volatility : {ms['Volatility']*100:.2f}%")
    print(f"  Sharpe     : {ms['Sharpe']:.4f}")

    print("\n--- Min Volatility Portfolio ---")
    print(f"  Return     : {mv['Return']*100:.2f}%")
    print(f"  Volatility : {mv['Volatility']*100:.2f}%")
    print(f"  Sharpe     : {mv['Sharpe']:.4f}")

    # 4. Save the efficient frontier chart
    print("\nGenerating chart …")
    plot_efficient_frontier(simulation, optimal)

    # 5. Export results to Excel
    print("Exporting to Excel …")
    write_excel(
        optimal=optimal,
        simulation=simulation,
        expected_returns=expected_returns,
        volatilities=volatilities,
        daily_returns=daily_returns,
        tickers=live_tickers,
    )

    print("\nDone!")


if __name__ == "__main__":
    sys.exit(main())
