# Portfolio Analysis & Dashboard

A comprehensive suite for portfolio management, performance tracking, and risk analysis. This project combines a dynamic web-based dashboard with advanced portfolio optimization scripts.

## Project Overview

The project is designed to help investors analyze their stock portfolios by parsing transaction histories, fetching real-time market data, and calculating key financial metrics.

### Key Components

- **Web Dashboard (`app.py`)**: A Dash-based interface for visualizing portfolio allocation, performance vs. benchmark (CAC 40), risk metrics (VaR, Sharpe, etc.), and drawdowns. It includes a position calculator for risk-based sizing.
- **Core Engine (`engine.py`)**: The data processing layer that handles transaction parsing, price fetching (via `yfinance`), and P&L calculations.
- **Portfolio Modelling (`Modelling_Portfolio.py`)**: A standalone script for portfolio optimization using the Efficient Frontier (via `PyPortfolioOpt`). It interacts with a local Excel workbook.
- **Financial Utilities (`Portfolio_functions.py`)**: Reusable functions for calculating drawdowns, return ratios, and fetching company fundamental data.
- **Data Source**: Excel files (.xlsx or .xlsm) containing transaction history.

## Building and Running

### Prerequisites

You will need Python 3.8+ and the following libraries:

```bash
pip install dash dash-bootstrap-components pandas numpy plotly yfinance scipy PyPortfolioOpt xlwings requests matplotlib
```

### Running the Dashboard

To start the interactive web dashboard:

```bash
python app.py
```
Then open your browser at `http://127.0.0.1:8050`. You can upload your `.xlsx` transaction file directly through the UI.

### Running Portfolio Optimization

The standalone modelling script requires a local Excel file (`Portfolio_Le_M.xlsm`):

```bash
python Modelling_Portfolio.py
```
*Note: Set the `PORTFOLIO_XLSM` environment variable to your workbook path, or place `Portfolio_Le_M.xlsm` in the working directory.*

## Development Conventions

- **Architecture**: The project follows a clear separation between the UI layer (`app.py`), the data logic layer (`engine.py`), and specialized analytics (`Modelling_Portfolio.py`).
- **Data Ingestion**: Transactions are parsed from Excel files with specific columns (Date, Ticker, Type, Quantite, Cash_Flow).
- **Styling**: The dashboard uses `dash-bootstrap-components` and custom CSS in `assets/style.css`.
- **Market Data**: Real-time and historical price data is fetched primarily using `yfinance`.
- **API Integration**: Fundamental data is fetched from the Financial Modeling Prep API. Keys are read from environment variables (`FMP_API_KEY`, `FRED_API_KEY`) — never hardcoded. See `.env.example`.

## Key Files

- `app.py`: Main entry point for the web application.
- `engine.py`: Logic for computing current positions and P&L.
- `Modelling_Portfolio.py`: Script for portfolio optimization and historical benchmark analysis.
- `Portfolio_functions.py`: Core financial metrics and data fetching utilities.
- `assets/style.css`: Custom styling for the Dash dashboard.
- `Portfolio_Le_M.xlsm`: Excel template/source for portfolio data.
- `base_transactions_propre.xlsx`: Example cleaned transaction data.
