"""Equity Factor Model.

Tools for pulling Alpha Vantage fundamental data, computing standardized equity factor scores, applying macroeconomic factor weights, and exporting analysis reports.
"""

from .model import AlphaVantageClient

__version__ = "0.1.0"
__all__ = ["AlphaVantageClient"]
