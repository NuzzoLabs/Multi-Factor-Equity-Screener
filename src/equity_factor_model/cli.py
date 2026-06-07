"""Command-line entry point for the equity factor model."""

from __future__ import annotations

import argparse
import os
from pathlib import Path

from .model import AlphaVantageClient


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the Alpha Vantage fundamental equity factor model."
    )
    parser.add_argument(
        "--tickers",
        nargs="+",
        required=True,
        help="Ticker symbols to analyze, e.g. AAPL MSFT JOBY.",
    )
    parser.add_argument(
        "--api-key",
        default=os.getenv("ALPHAVANTAGE_API_KEY"),
        help="Alpha Vantage API key. Defaults to ALPHAVANTAGE_API_KEY environment variable.",
    )
    parser.add_argument(
        "--output",
        default="results/stock_factors_analysis.xlsx",
        help="Excel output path.",
    )
    return parser.parse_args()


def run_factor_model(tickers: list[str], api_key: str, output: str = "results/stock_factors_analysis.xlsx"):
    """Run the factor model for a list of tickers and export an Excel report."""
    client = AlphaVantageClient(api_key=api_key)
    stock_data = client.Collect_Data(tickers)

    bal_sheets = [stock_data["Balance_Sheet"][i]["annualReports"] for i, _ in enumerate(stock_data["Balance_Sheet"])]
    income_statements = [stock_data["Income_Statement"][i]["annualReports"] for i, _ in enumerate(stock_data["Income_Statement"])]
    cash_flow = [stock_data["Cash_Flow"][i]["annualReports"] for i, _ in enumerate(stock_data["Cash_Flow"])]
    overview = [stock_data["Info"][i] for i, _ in enumerate(stock_data["Info"])]
    insiders = [stock_data["Insider_Activity"][i] for i, _ in enumerate(stock_data["Insider_Activity"])]

    stock_size = [len(bal_sheets[i]) for i in range(len(tickers))]

    tax_rate = client.compute_TaxRate(stock_size, income_statements)
    ebit = client.compute_EBIT(stock_size, income_statements)
    nopat = client.compute_NOPAT(ebit, tax_rate)
    invested_capital = client.compute_InvestedCapital(stock_size, bal_sheets)
    da = client.compute_DepreciationandAmortization(stock_size, income_statements)
    delta_nwc = client.compute_Del_NWC(stock_size, bal_sheets)
    net_capex = client.compute_NetCapEx(stock_size, cash_flow)
    fcf = client.compute_FCF(nopat, da, delta_nwc, net_capex)
    revenue = client.compute_Revenue(stock_size, income_statements)
    roic = client.compute_ROIC(stock_size, bal_sheets, nopat, invested_capital)

    liquidity = client.compute_liquidity_factor(bal_sheets, income_statements)
    growth = client.compute_growth_factor(tickers, bal_sheets, cash_flow, income_statements)
    cap_eff = client.compute_capital_efficiency_factor(nopat, invested_capital, fcf, revenue, roic)
    coverage = client.compute_coverage_factor(stock_size, overview, bal_sheets, income_statements, fcf, ebit)
    value = client.compute_value_factor(stock_size, overview, bal_sheets, ebit)
    risk = client.compute_risk_factor(stock_size, insiders, overview, bal_sheets, revenue)

    factors = client.Factor_Model(liquidity, growth, cap_eff, coverage, value, risk, tickers)
    macro_weights, factors_adjusted = client.compute_Macro_Weight(factors)

    Path(output).parent.mkdir(parents=True, exist_ok=True)
    client.export_stock_factors_to_excel(factors, factors_adjusted, output_file=output)
    return factors, macro_weights, factors_adjusted


def main() -> None:
    args = parse_args()
    if not args.api_key:
        raise SystemExit("Missing API key. Set ALPHAVANTAGE_API_KEY or pass --api-key.")
    factors, macro_weights, factors_adjusted = run_factor_model(args.tickers, args.api_key, args.output)
    print("Factors:")
    print(factors)
    print("\nMacro weights:")
    print(macro_weights)
    print("\nMacro-adjusted factors:")
    print(factors_adjusted)


if __name__ == "__main__":
    main()
