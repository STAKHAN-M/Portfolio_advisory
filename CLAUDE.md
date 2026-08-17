# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Project

**Install dependencies:**
```bash
pip install dash dash-bootstrap-components pandas numpy plotly yfinance scipy PyPortfolioOpt xlwings requests matplotlib
```

**Start the dashboard:**
```bash
python app.py
# Access at http://127.0.0.1:8050
```

**Run portfolio optimization (standalone, requires local Excel):**
```bash
python Modelling_Portfolio.py
# Excel path: set env var PORTFOLIO_XLSM, or place Portfolio_Le_M.xlsm in the working directory
```

There are no automated tests and no linting configuration.

## Architecture

The codebase has a strict layered separation:

```
app.py  →  engine.py  →  yfinance / Excel
              ↑
Portfolio_functions.py  (used by Modelling_Portfolio.py only)
```

**`engine.py`** is the sole data and analytics layer for the dashboard. It exposes:
- `build_positions_table(df)` — parses raw transactions and fetches live prices, returns a positions DataFrame + summary dict. Summary keys: `total_value`, `total_cost`, `total_pnl`, `pnl_pct`, `total_dividends`, `total_invested`, `total_return_pct`, `cash_balance`
- `compute_analytics(tickers, weights, start, df=None)` — downloads monthly historical data. When `df` is provided, computes performance via **TWR** (Time-Weighted Return) from actual transaction history instead of applying current weights retroactively. Returns performance series, per-ticker metrics, Sharpe, Sortino, Max DD, VaR/CVaR, Beta, Alpha vs CAC 40 (`^FCHI`)
- `_compute_twr(df, price_data)` — internal helper. Reconstructs monthly portfolio holdings from transactions, computes sub-period returns, and chains them to eliminate cash-flow distortion
- `compute_optimal_portfolios(tickers, start)` — runs SLSQP optimization for MSR, GMV, and Equal Weight portfolios

**`app.py`** is pure UI — all Dash callbacks, layout, and formatting. It imports only `engine`. State is passed between callbacks via two `dcc.Store` components:
- `store-data` — raw transactions JSON + Valorisation sheet JSON (uploaded file)
- `store-analytics` — all computed analytics, serialized as JSON

The file upload triggers `store-data`; analytics are computed on demand via a separate callback into `store-analytics`.

**`Portfolio_functions.py`** is a separate utility library used exclusively by `Modelling_Portfolio.py`. It duplicates some analytics logic from `engine.py` (annualization, Sharpe, drawdown) but adds `information_portfolio()` which calls the Financial Modeling Prep API for fundamentals.

**`Modelling_Portfolio.py`** is a standalone, non-interactive script that reads directly from a local `.xlsm` file via `xlwings`, runs PyPortfolioOpt's Efficient Frontier, and plots results. It is not connected to the Dash app.

## Dashboard Layout — Tabs

The dashboard has 4 main areas:

**KPI Row** (always visible after data load) — 5 cards:
1. Capital Versé (`total_invested`)
2. Montant Investi (`total_cost`)
3. Valeur Actuelle (`total_value`)
4. Cash Disponible (`cash_balance`)
5. Plus-Values Latentes (`total_pnl` + `pnl_pct` badge)

**Tab 1 — Vue d'ensemble:**
- Positions table (sortable/filterable) + Répartition pie chart
- Performance vs Benchmark (small, base 100) + Métriques de Risque panel
- Comparaison de Portefeuilles Optimaux + Drawdown chart
- Simulateur d'Achat & Pondération (PRU calculator, 3 simulations, weighting vs current portfolio value)

**Tab 2 — Analyse par Titre:**
- Table: Rendement Mensuel/Annuel, Volatilité Mensuelle/Annuelle, Rendement/Volatilité per ticker + benchmark. Color-coded: R/V ≥ 0.7 → green, < 0.2 → red.

**Tab 3 — Performance Historique:**
- Large performance chart (460px) vs CAC 40, base 100
- Monthly history table: Valeur PF (€), Rend. Mensuel PF (%), Perf. Cumulée PF (%), Rend. Mensuel CAC 40 (%), Perf. Cumulée CAC 40 (%), Écart Mensuel (%)
- Source label indicates whether data comes from Valorisation sheet or TWR fallback

## Data Format

The input file is `base_transactions_propre.xlsx` with two sheets:

**Sheet `BDD`** (primary source — always required):
Columns: `Date`, `Type`, `Ticker`, `Description_Operation`, `Quantite`, `Prix_Unitaire`, `Montant_Total`, `Cash_Flow`
Transaction types: `Achat`, `Vente`, `Dividende`, `Versement`

**Sheet `Valorisation`** (optional — used for performance chart):
Columns: `Date`, `Valeur_Portefeuille`, `Cash`, `Total_Versements`, `Plus_Moins_Value`
Monthly snapshots. If present and populated, used directly for the performance chart. If absent or empty, the dashboard falls back to TWR computed from BDD.

The upload callback reads `sheet_name="BDD"` explicitly and attempts `sheet_name="Valorisation"` with a try/except fallback to an empty DataFrame.

## Key Constants (engine.py)

- `BENCHMARK = "^FCHI"` — CAC 40 index used for all benchmark comparisons
- `RF = 0.033` — annual risk-free rate used in Sharpe/Sortino/Alpha calculations

## Position Calculator Logic

- Capital reference = `total_value` (current market value), **not** `total_cost`
- Simulates up to 3 additional buy tranches per ticker
- Outputs: new PRU, new quantity, simulated market value, new portfolio weight

## API Keys

Keys are read from environment variables only — none are hardcoded. `Portfolio_functions.py` reads `FMP_API_KEY` (Financial Modeling Prep, `information_portfolio()`), and `macro_engine.py` reads `FRED_API_KEY`. The dashboard (`app.py` + `engine.py`) does **not** use FMP — it relies solely on `yfinance` for market data. See `.env.example`.
