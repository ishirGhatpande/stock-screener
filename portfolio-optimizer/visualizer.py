# visualizer.py
# Plots the efficient frontier and saves it to disk.

import matplotlib.pyplot as plt
import pandas as pd

from config import CHART_FILENAME


def plot_efficient_frontier(
    simulation: pd.DataFrame,
    optimal: dict,
    filename: str = CHART_FILENAME,
) -> None:
    """
    Scatter-plot all simulated portfolios coloured by Sharpe ratio.

    Special markers highlight the max-Sharpe and min-volatility portfolios.
    The chart is saved to `filename`.
    """
    fig, ax = plt.subplots(figsize=(12, 8))

    # --- Background scatter: all simulated portfolios ---
    sc = ax.scatter(
        simulation["Volatility"] * 100,
        simulation["Return"] * 100,
        c=simulation["Sharpe"],
        cmap="viridis",
        alpha=0.5,
        s=8,
        linewidths=0,
    )
    cbar = fig.colorbar(sc, ax=ax, pad=0.02)
    cbar.set_label("Sharpe Ratio", fontsize=11)

    # --- Max Sharpe portfolio ---
    ms = optimal["max_sharpe"]
    ax.scatter(
        ms["Volatility"] * 100,
        ms["Return"] * 100,
        marker="*",
        color="gold",
        edgecolors="black",
        linewidths=0.8,
        s=400,
        zorder=5,
        label=f"Max Sharpe  ({ms['Sharpe']:.2f})  —  "
              f"Ret {ms['Return']*100:.1f}%  Vol {ms['Volatility']*100:.1f}%",
    )

    # --- Min volatility portfolio ---
    mv = optimal["min_vol"]
    ax.scatter(
        mv["Volatility"] * 100,
        mv["Return"] * 100,
        marker="D",
        color="tomato",
        edgecolors="black",
        linewidths=0.8,
        s=200,
        zorder=5,
        label=f"Min Volatility  ({mv['Volatility']*100:.1f}%)  —  "
              f"Ret {mv['Return']*100:.1f}%  Sharpe {mv['Sharpe']:.2f}",
    )

    # --- Formatting ---
    ax.set_title("Efficient Frontier — Modern Portfolio Theory", fontsize=15, pad=14)
    ax.set_xlabel("Annualised Volatility (%)", fontsize=12)
    ax.set_ylabel("Annualised Expected Return (%)", fontsize=12)
    ax.legend(fontsize=10, loc="upper left")
    ax.grid(True, linestyle="--", alpha=0.4)

    # Tight layout then save
    plt.tight_layout()
    plt.savefig(filename, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Chart saved -> {filename}")
