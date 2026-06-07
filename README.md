## Equity Factor Model
Automated equity screening framework built on fundamental data from the Alpha Vantage API.
This project evaluates companies using a multi-factor framework inspired by quality-first investing principles. The model emphasizes capital efficiency, balance sheet strength, free cash flow generation, valuation, risk, and macroeconomic context.

## Features
- Pulls company fundamentals from Alpha Vantage
- Computes standardized 0-100 factor scores
- Evaluates liquidity, growth, capital efficiency, coverage, value, and risk
- Applies macroeconomic weighting adjustments
- Exports normal, macro-adjusted, and delta factor tables to Excel
- Includes radar/spider chart visualization for factor comparison

## Planned Functionality
- Backtesting engine
- Portfolio risk analysis
- Sector concentration analysis

## Factor Categories
- Liquidity: current ratio, quick ratio, working capital per revenue
- Growth: reinvestment quality, revenue CAGR, free cash flow CAGR
- Capital Efficiency: ROIC, ROIIC, free cash flow margin
- Coverage: debt coverage, interest coverage, dividend coverage
- Value: price/book, price/earnings, EV/EBIT
- Risk: insider activity, dilution, revenue volatility, debt capitalization

## Installation
```bash
pip install -r requirements.txt
pip install -e .
```

## API Key Setup
This project uses Alpha Vantage data. Do not hard-code API keys in source code.
Set your key as an environment variable:
```bash
set ALPHAVANTAGE_API_KEY=your_api_key_here
```
On macOS/Linux:
```bash
export ALPHAVANTAGE_API_KEY=your_api_key_here
```
## Example Usage
```bash
python -m equity_factor_model.cli --tickers AAPL MSFT JOBY --output results/stock_factors.xlsx
```
## Repository Structure
```text
equity-factor-model/
├── src/
│   └── equity_factor_model/
│       ├── __init__.py
│       ├── model.py
│       └── cli.py
├── tests/
│   └── test_import.py
├── docs/
│   └── original_research_script_sanitized.py
├── README.md
├── requirements.txt
├── pyproject.toml
└── .gitignore
```
## Research Code Notice
This repository was converted from a personal investment research script into a reusable Python package structure. The methodology and outputs should be independently validated before use in investment decisions.
This is not financial advice. The model is intended for research, education, and analytical workflow development.

## Status
Research code under active refactoring and validation.
