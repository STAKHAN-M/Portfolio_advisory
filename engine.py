"""
Portfolio Engine - Computes positions, P&L, and risk analytics
from a transaction history (Excel file).
"""

import warnings
import numpy as np
import pandas as pd
import yfinance as yf
from scipy.stats import norm
from scipy.optimize import minimize

warnings.filterwarnings("ignore")

BENCHMARK = "^FCHI"
RF = 0.033  # Risk-free rate (annual)

# ── Ticker categories ──────────────────────────────────────────────────────────

CAC40_TICKERS = {
    "AI.PA", "AIR.PA", "ACA.PA", "AM.PA", "BN.PA", "BNP.PA", "CA.PA",
    "CAP.PA", "CS.PA", "DG.PA", "DSY.PA", "EN.PA", "ENGI.PA", "ERF.PA",
    "GLE.PA", "HO.PA", "KER.PA", "LR.PA", "MC.PA", "ML.PA", "OR.PA",
    "ORA.PA", "PUB.PA", "RI.PA", "RMS.PA", "RNO.PA", "SAF.PA", "SAN.PA",
    "SGO.PA", "STMPA.PA", "SU.PA", "TEP.PA", "TTE.PA", "URW.AS",
    "VIE.PA", "VIV.PA", "WLN.PA",
}

ETF_TICKERS = {
    "PE500.PA", "P500.PA", "BNKE.PA", "CACC.PA", "GRE.PA", "PINR.PA",
    "PASI.PA", "EUDV.PA", "AWAT.PA", "PAEEM.PA", "PANX.PA", "PFAM.PA",
    "CW8.PA", "WPEA.PA", "PSP5.PA",
}


def get_category(ticker: str) -> str:
    if ticker in CAC40_TICKERS:
        return "CAC 40"
    if ticker in ETF_TICKERS:
        return "Index"
    return "Action"


def _yf_close(tickers, **kwargs) -> pd.DataFrame:
    """
    Download closing prices compatible with both old yfinance (Price, Ticker)
    and new yfinance >=0.2.50 (Ticker, Price) MultiIndex column formats.
    Always returns a DataFrame with tickers as columns.
    """
    if isinstance(tickers, str):
        tickers = [tickers]
    raw = yf.download(tickers, progress=False, auto_adjust=True, **kwargs)
    if raw.empty:
        return pd.DataFrame()
    if isinstance(raw.columns, pd.MultiIndex):
        lvl0 = raw.columns.get_level_values(0).unique()
        lvl1 = raw.columns.get_level_values(1).unique()
        if "Close" in lvl0:
            data = raw["Close"]          # old format: (Price, Ticker)
        elif "Close" in lvl1:
            data = raw.xs("Close", level=1, axis=1)  # new format: (Ticker, Price)
        else:
            data = raw
    else:
        data = raw
    if isinstance(data, pd.Series):
        data = data.to_frame(name=tickers[0])
    if isinstance(data.index, pd.DatetimeIndex) and data.index.tz is not None:
        data.index = data.index.tz_convert(None)
    return data


# ─── Transaction Parsing ──────────────────────────────────────────────────────

def parse_transactions(df: pd.DataFrame) -> dict:
    """Parse transaction DataFrame to derive current open positions."""
    positions = {}

    for _, row in df.iterrows():
        ticker = row["Ticker"]
        if pd.isna(ticker):
            continue

        if ticker not in positions:
            positions[ticker] = {"qty": 0.0, "total_cost": 0.0, "dividends": 0.0}

        p = positions[ticker]
        t = str(row["Type"]).strip().capitalize()
        if t not in ["Achat", "Vente", "Dividende", "Versement"]:
            continue


        if t == "Achat":
            qty = float(row["Quantite"])
            cost = abs(float(row["Cash_Flow"]))
            p["qty"] += qty
            p["total_cost"] += cost

        elif t == "Vente":
            qty = float(row["Quantite"])
            if p["qty"] > 0:
                avg = p["total_cost"] / p["qty"]
                p["total_cost"] = max(0.0, p["total_cost"] - avg * qty)
                p["qty"] = max(0.0, p["qty"] - qty)

        elif t == "Dividende":
            p["dividends"] += float(row["Cash_Flow"])

    return {t: p for t, p in positions.items() if p["qty"] > 0.001}


def compute_realized_and_dividends(df: pd.DataFrame) -> tuple:
    """
    Calcule, par titre :
      - le P&L réalisé (méthode du coût moyen) sur les ventes,
      - les dividendes perçus.
    Retourne (closed, dividends) :
      closed = liste de dicts pour les positions CLÔTURÉES (qté finale ≈ 0, au moins 1 vente)
      dividends = liste de dicts {ticker, dividends} triée décroissante.
    """
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values("Date")

    state = {}
    for _, row in df.iterrows():
        t = row["Ticker"]
        if pd.isna(t):
            continue
        typ = str(row["Type"]).strip().capitalize()
        s = state.setdefault(t, {"qty": 0.0, "cost": 0.0, "realized": 0.0,
                                 "dividends": 0.0, "n_sells": 0,
                                 "buy_cost": 0.0, "proceeds": 0.0})
        if typ == "Achat":
            qty = float(row["Quantite"]); cost = abs(float(row["Cash_Flow"]))
            s["qty"] += qty; s["cost"] += cost; s["buy_cost"] += cost
        elif typ == "Vente":
            qty = float(row["Quantite"]); proceeds = abs(float(row["Cash_Flow"]))
            if s["qty"] > 0:
                avg = s["cost"] / s["qty"]
                sold = min(qty, s["qty"])
                cost_sold = avg * sold
                s["realized"] += proceeds - cost_sold
                s["cost"] = max(0.0, s["cost"] - cost_sold)
                s["qty"] = max(0.0, s["qty"] - qty)
            s["proceeds"] += proceeds
            s["n_sells"] += 1
        elif typ == "Dividende":
            s["dividends"] += float(row["Cash_Flow"])

    closed = []
    for t, s in state.items():
        if s["n_sells"] > 0 and s["qty"] <= 0.001:
            closed.append({
                "ticker": t,
                "realized": round(s["realized"], 2),
                "proceeds": round(s["proceeds"], 2),
                "buy_cost": round(s["buy_cost"], 2),
                "dividends": round(s["dividends"], 2),
                "total": round(s["realized"] + s["dividends"], 2),
            })
    closed.sort(key=lambda x: x["realized"], reverse=True)

    dividends = [{"ticker": t, "dividends": round(s["dividends"], 2)}
                 for t, s in state.items() if s["dividends"] > 0.001]
    dividends.sort(key=lambda x: x["dividends"], reverse=True)

    return closed, dividends


# ─── Price Fetching ───────────────────────────────────────────────────────────

def fetch_current_prices(tickers: list) -> pd.Series:
    """Fetch latest closing prices for a list of tickers."""
    if not tickers:
        return pd.Series(dtype=float)
    try:
        data = _yf_close(tickers, period="5d", interval="1d")
        return data.ffill().iloc[-1]
    except Exception:
        return pd.Series(dtype=float)


# ─── Portfolio Summary ────────────────────────────────────────────────────────

def build_positions_table(df: pd.DataFrame) -> tuple:
    """
    Returns (positions_df, summary_dict) from raw transactions DataFrame.
    summary_dict keys: total_value, total_cost, total_pnl, pnl_pct,
                       total_dividends, total_invested, total_return_pct, cash_balance
    """
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"])

    positions = parse_transactions(df)
    if not positions:
        return pd.DataFrame(), {}

    tickers = list(positions.keys())
    current_prices = fetch_current_prices(tickers)

    rows = []
    for ticker, pos in positions.items():
        qty = pos["qty"]
        total_cost = pos["total_cost"]
        avg_price = total_cost / qty if qty > 0 else 0.0

        curr_price = float(current_prices.get(ticker, avg_price))
        market_value = curr_price * qty
        pnl = market_value - total_cost
        pnl_pct = (pnl / total_cost * 100) if total_cost > 0 else 0.0
        dividends = pos["dividends"]
        total_return_eur = pnl + dividends
        total_return_pct = (total_return_eur / total_cost * 100) if total_cost > 0 else 0.0

        rows.append({
            "Ticker": ticker,
            "Qté": round(qty, 2),
            "PRU (€)": round(avg_price, 2),
            "Cours (€)": round(curr_price, 2),
            "Valeur (€)": round(market_value, 2),
            "P&L (€)": round(pnl, 2),
            "P&L (%)": round(pnl_pct, 2),
            "Dividendes (€)": round(dividends, 2),
            "Retour Total (%)": round(total_return_pct, 2),
        })

    positions_df = pd.DataFrame(rows)
    total_value = positions_df["Valeur (€)"].sum()
    positions_df["Poids (%)"] = (positions_df["Valeur (€)"] / total_value * 100).round(2)
    positions_df["Catégorie"] = positions_df["Ticker"].map(get_category)
    positions_df = positions_df[[
        "Ticker", "Catégorie", "Qté", "PRU (€)", "Cours (€)", "Valeur (€)",
        "Poids (%)", "P&L (€)", "P&L (%)", "Dividendes (€)", "Retour Total (%)"
    ]].sort_values("Valeur (€)", ascending=False).reset_index(drop=True)

    # --- GLOBAL PERFORMANCE CALCULATION ---
    total_invested = float(df[df["Type"].str.strip().str.capitalize() == "Versement"]["Cash_Flow"].sum())
    total_cost_current = sum(p["total_cost"] for p in positions.values())

    # All-time dividends — includes dividends from fully sold positions
    total_dividends_all = float(
        df[df["Type"].str.strip().str.capitalize() == "Dividende"]["Cash_Flow"].sum()
    )

    cash_balance = float(df["Cash_Flow"].sum())

    # Performance totale = (valeur marché + cash) - capital versé
    # Inclut gains réalisés + non réalisés + dividendes
    total_pnl = (total_value + cash_balance) - total_invested

    # +/- Values latentes = valeur marché - prix de revient positions ouvertes (comparable broker)
    unrealized_pnl = total_value - total_cost_current
    unrealized_pnl_pct = (unrealized_pnl / total_cost_current * 100) if total_cost_current > 0 else 0.0

    total_portfolio_value = total_value + cash_balance

    # ── TRI / rendement money-weighted (XIRR) ────────────────────────────────
    # Flux investisseur = versements (sorties, négatifs) + valeur actuelle (entrée finale)
    vers = df[df["Type"].str.strip().str.capitalize() == "Versement"]
    cfs = [(d, -float(cf)) for d, cf in zip(vers["Date"], vers["Cash_Flow"])
           if pd.notna(cf) and float(cf) != 0]
    money_weighted = None
    if cfs and total_portfolio_value > 0:
        cfs.append((pd.Timestamp.today().normalize(), float(total_portfolio_value)))
        try:
            mwr = _xirr(cfs)
            money_weighted = round(mwr * 100, 2) if mwr is not None else None
        except Exception as e:
            print(f"[engine] XIRR error: {e}")

    summary = {
        "total_portfolio_value": round(total_portfolio_value, 2),
        "total_value": round(total_value, 2),
        "total_cost": round(total_cost_current, 2),
        "total_pnl": round(total_pnl, 2),
        "pnl_pct": round((total_pnl / total_invested * 100) if total_invested > 0 else 0, 2),
        "total_dividends": round(total_dividends_all, 2),
        "total_invested": round(total_invested, 2),
        "cash_balance": round(cash_balance, 2),
        "unrealized_pnl": round(unrealized_pnl, 2),
        "unrealized_pnl_pct": round(unrealized_pnl_pct, 2),
        "money_weighted_return": money_weighted,
    }

    return positions_df, summary


# ─── Money-Weighted Return (TRI / XIRR) ────────────────────────────────────────

def _xirr(cashflows, lo=-0.9999, hi=10.0):
    """TRI annualisé (XIRR) par bissection. cashflows = [(date, montant), ...]
    Convention : versements négatifs (sorties investisseur), valeur finale positive."""
    cashflows = [(pd.Timestamp(d), float(a)) for d, a in cashflows if pd.notna(d)]
    if len(cashflows) < 2:
        return None
    cashflows.sort(key=lambda x: x[0])
    t0 = cashflows[0][0]
    years = [(d - t0).days / 365.0 for d, _ in cashflows]
    amts  = [a for _, a in cashflows]

    def npv(r):
        return sum(a / (1.0 + r) ** y for a, y in zip(amts, years))

    flo, fhi = npv(lo), npv(hi)
    if flo * fhi > 0:                 # pas de changement de signe → pas de solution
        return None
    for _ in range(200):
        mid = (lo + hi) / 2.0
        fmid = npv(mid)
        if abs(fmid) < 1e-7:
            return mid
        if flo * fmid < 0:
            hi, fhi = mid, fmid
        else:
            lo, flo = mid, fmid
    return (lo + hi) / 2.0


# ─── Time-Weighted Return ─────────────────────────────────────────────────────

def _compute_twr(df: pd.DataFrame, price_data: pd.DataFrame):
    """
    Compute Time-Weighted Return from transaction history.

    For each monthly sub-period [t, t+1]:
      - Holdings are fixed at whatever was held at the END of month t
        (i.e. after all transactions dated ≤ t)
      - Sub-period return = value(holdings, t+1) / value(holdings, t) - 1

    Chaining these sub-period returns gives the TWR:
    it eliminates the distortion from cash deposits / withdrawals
    and only counts the months where the portfolio was actually invested.
    """
    df_txn = (
        df[df["Type"].isin(["Achat", "Vente"])]
        .dropna(subset=["Ticker"])
        .sort_values("Date")
        .reset_index(drop=True)
    )
    if df_txn.empty:
        return None, None

    # Forward-fill prices once (handles gaps / newly-listed tickers)
    prices = price_data.ffill()
    monthly_dates = prices.index.sort_values()

    # ── Build holdings snapshot at each month-end ──────────────────────────
    holdings: dict = {}
    txn_idx = 0
    n_txn = len(df_txn)
    holdings_at: dict = {}   # date -> {ticker: qty}

    for date in monthly_dates:
        while txn_idx < n_txn and df_txn.loc[txn_idx, "Date"] <= date:
            row = df_txn.loc[txn_idx]
            t = row["Ticker"]
            holdings.setdefault(t, 0.0)
            if row["Type"] == "Achat":
                holdings[t] += float(row["Quantite"])
            elif row["Type"] == "Vente":
                holdings[t] = max(0.0, holdings[t] - float(row["Quantite"]))
            txn_idx += 1
        holdings_at[date] = {t: q for t, q in holdings.items() if q > 0.001}

    # ── Value a holdings dict at a given date ──────────────────────────────
    def _val(h: dict, date) -> float:
        total = 0.0
        for t, qty in h.items():
            if t in prices.columns:
                p = prices.at[date, t]
                if not np.isnan(p):
                    total += qty * float(p)
        return total

    # ── Chain sub-period returns ───────────────────────────────────────────
    first_invest = df_txn["Date"].min()
    sub_rets: list = []
    dates_out: list = []
    monthly_list = list(monthly_dates)

    for i in range(len(monthly_list) - 1):
        d0, d1 = monthly_list[i], monthly_list[i + 1]
        if d1 <= first_invest:          # portfolio not yet started
            continue
        h = holdings_at.get(d0, {})
        if not h:
            continue
        v0 = _val(h, d0)
        v1 = _val(h, d1)
        if v0 > 0:
            sub_rets.append(v1 / v0 - 1)
            dates_out.append(d1)

    if not sub_rets:
        return None, None

    pf_ret = pd.Series(sub_rets, index=pd.DatetimeIndex(dates_out, name="Date"))
    perf_pf = (100 * (1 + pf_ret).cumprod()).round(2)
    return pf_ret, perf_pf


# ─── Risk Analytics ───────────────────────────────────────────────────────────

def _annualize_ret(r: pd.Series, freq: int = 12) -> float:
    """Rendement annualisé GÉOMÉTRIQUE (CAGR réalisé), sans biais de volatilité."""
    r = r.dropna()
    if len(r) == 0:
        return 0.0
    cum = float((1 + r).prod())
    if cum <= 0:
        return -1.0
    return float(cum ** (freq / len(r)) - 1)


def _annualize_vol(r: pd.Series, freq: int = 12) -> float:
    return float(r.std() * np.sqrt(freq))


def _sharpe(r: pd.Series, rf: float = RF, freq: int = 12) -> float:
    ret = _annualize_ret(r, freq)
    vol = _annualize_vol(r, freq)
    return float((ret - rf) / vol) if vol > 0 else 0.0


def _sortino(r: pd.Series, rf: float = RF, freq: int = 12) -> float:
    ret = _annualize_ret(r, freq)
    neg = r[r < 0]
    if len(neg) < 2:
        return 0.0
    down_vol = neg.std() * np.sqrt(freq)
    return float((ret - rf) / down_vol) if down_vol > 0 else 0.0


def _max_drawdown(r: pd.Series) -> float:
    wealth = (1 + r).cumprod()
    peak = wealth.cummax()
    dd = (wealth - peak) / peak
    return float(dd.min())


def _var_cvar(r: pd.Series, level: float = 5.0) -> tuple:
    var_hist = float(-np.percentile(r, level))
    z = norm.ppf(level / 100)
    var_gauss = float(-(r.mean() + z * r.std(ddof=0)))
    below = r[r <= -var_hist]
    cvar = float(-below.mean()) if len(below) > 0 else var_hist
    return var_hist, var_gauss, cvar


def compute_analytics(tickers: list, weights: dict, start: str = "2019-01-01",
                       df: pd.DataFrame = None) -> dict:
    """
    Fetch historical monthly data and compute:
    - Portfolio & benchmark performance index (base 100)  ← TWR when df provided
    - Risk/return metrics: Sharpe, Sortino, Max DD, VaR, CVaR, Beta, Alpha
    """
    try:
        # All tickers ever held (needed for TWR price download)
        hist_tickers = list(tickers)
        if df is not None:
            ever_held = (
                df[df["Type"].isin(["Achat", "Vente"])]["Ticker"]
                .dropna().unique().tolist()
            )
            hist_tickers = list(set(tickers + ever_held))

        all_dl = hist_tickers + [BENCHMARK]
        data = _yf_close(all_dl, start=start, interval="1mo")

        # Current tickers only — used for per-ticker metrics & weights
        available = [t for t in tickers if t in data.columns]
        if not available:
            return {}

        stock_data = data[available].copy()
        bench_data = data[BENCHMARK].copy() if BENCHMARK in data.columns else None

        # Aligned returns (intersection of all current tickers + benchmark)
        # Used for portfolio-level metrics: Sharpe, Beta, Alpha, drawdown
        stock_ret_aligned = stock_data.pct_change().dropna()
        if bench_data is not None:
            bench_ret = bench_data.pct_change().dropna()
            idx = stock_ret_aligned.index.intersection(bench_ret.index)
            stock_ret_aligned = stock_ret_aligned.loc[idx]
            bench_ret = bench_ret.loc[idx]
        else:
            bench_ret = None

        # Per-ticker returns computed independently (each ticker's own full history)
        stock_ret = stock_data.pct_change()

        # Weights vector — always defined
        w = np.array([weights.get(t, 0.0) for t in available])
        w = w / w.sum() if w.sum() > 0 else np.repeat(1.0 / len(available), len(available))

        # ── Portfolio return series ────────────────────────────────────────
        pf_ret, perf_pf = None, None
        if df is not None:
            try:
                hist_available = [t for t in hist_tickers if t in data.columns]
                pf_ret, perf_pf = _compute_twr(df, data[hist_available])
            except Exception as e:
                print(f"[engine] TWR error: {e}")

        if pf_ret is None:
            # Fallback: current weights applied to aligned history
            pf_ret = stock_ret_aligned.dot(w)
            perf_pf = (100 * (1 + pf_ret).cumprod()).round(2)

        var_h, var_g, cvar = _var_cvar(pf_ret)

        metrics = {
            "ann_ret": round(_annualize_ret(pf_ret) * 100, 2),
            "ann_vol": round(_annualize_vol(pf_ret) * 100, 2),
            "sharpe": round(_sharpe(pf_ret), 3),
            "sortino": round(_sortino(pf_ret), 3),
            "max_dd": round(_max_drawdown(pf_ret) * 100, 2),
            "var_hist": round(var_h * 100, 2),
            "var_gauss": round(var_g * 100, 2),
            "cvar": round(cvar * 100, 2),
        }

        # Per-ticker metrics — each ticker uses its own full history (no cross-ticker alignment)
        per_ticker = {}
        for t in available:
            r = stock_ret[t].dropna()   # independent history per ticker
            if len(r) < 2:
                continue
            m_ret = float(r.mean() * 100)
            y_ret = float(_annualize_ret(r) * 100)
            m_vol = float(r.std() * 100)
            y_vol = float(r.std() * np.sqrt(12) * 100)
            per_ticker[t] = {
                "monthly_ret": round(m_ret, 2),
                "yearly_ret":  round(y_ret, 2),
                "monthly_vol": round(m_vol, 2),
                "yearly_vol":  round(y_vol, 2),
                "ret_vol":     round(y_ret / y_vol, 2) if y_vol > 0 else 0.0,
            }

        result = {
            "available_tickers": available,
            "weights": dict(zip(available, w)),
            "pf_ret": pf_ret,
            "perf_pf": perf_pf,
            "metrics": metrics,
            "stock_ret": stock_ret_aligned,
            "per_ticker": per_ticker,
        }

        if bench_ret is not None:
            bench_ann_ret = _annualize_ret(bench_ret)
            bench_ann_vol = _annualize_vol(bench_ret)
            bench_sharpe = float((bench_ann_ret - RF) / bench_ann_vol) if bench_ann_vol > 0 else 0.0

            # Align pf_ret and bench_ret on common dates for covariance (TWR may differ)
            common_idx = pf_ret.index.intersection(bench_ret.index)
            pf_aligned = pf_ret.loc[common_idx]
            bench_aligned = bench_ret.loc[common_idx]

            if len(common_idx) >= 2:
                cov = np.cov(pf_aligned.values, bench_aligned.values)
                beta = float(cov[0, 1] / cov[1, 1]) if cov[1, 1] > 0 else 0.0
            else:
                beta = 0.0
            alpha = float(_annualize_ret(pf_ret) - (RF + beta * (bench_ann_ret - RF)))

            # Benchmark perf index aligned to portfolio start date
            bench_pf_start = bench_ret.loc[bench_ret.index >= pf_ret.index.min()]
            perf_bench = (100 * (1 + bench_pf_start).cumprod()).round(2)

            metrics["beta"] = round(beta, 3)
            metrics["alpha"] = round(alpha * 100, 2)

            b_m_ret = float(bench_ret.mean() * 100)
            b_y_ret = float(_annualize_ret(bench_ret) * 100)
            b_m_vol = float(bench_ret.std() * 100)
            b_y_vol = float(bench_ret.std() * np.sqrt(12) * 100)
            per_ticker[BENCHMARK] = {
                "monthly_ret": round(b_m_ret, 2),
                "yearly_ret":  round(b_y_ret, 2),
                "monthly_vol": round(b_m_vol, 2),
                "yearly_vol":  round(b_y_vol, 2),
                "ret_vol":     round(b_y_ret / b_y_vol, 2) if b_y_vol > 0 else 0.0,
            }

            result["bench_ret"] = bench_ret
            result["perf_bench"] = perf_bench
            result["bench_metrics"] = {
                "ann_ret": round(bench_ann_ret * 100, 2),
                "ann_vol": round(bench_ann_vol * 100, 2),
                "sharpe": round(bench_sharpe, 3),
            }

        return result

    except Exception as e:
        print(f"[engine] Analytics error: {e}")
        return {}


# ─── Portfolio Optimization ───────────────────────────────────────────────────

def _portfolio_return(w, mu):
    return float(w @ mu)


def _portfolio_vol(w, cov):
    return float(np.sqrt(w @ cov @ w))


def _neg_sharpe(w, mu, cov, rf=RF):
    r = _portfolio_return(w, mu)
    v = _portfolio_vol(w, cov)
    return -(r - rf) / v if v > 0 else 0.0


def compute_optimal_portfolios(tickers: list, start: str = "2019-01-01") -> dict:
    """
    Compute MSR, GMV, and Equal Weight portfolio metrics.
    Returns dict with weights, return, vol, sharpe for each.
    """
    try:
        data = _yf_close(tickers, start=start, interval="1mo")
        available = [t for t in tickers if t in data.columns]
        data = data[available]

        ret = data.pct_change().dropna()
        mu = np.array([(1 + ret[t].mean()) ** 12 - 1 for t in available])
        cov = ret.cov().values * 12

        n = len(available)
        bounds = [(0.0, 1.0)] * n
        constraints = [{"type": "eq", "fun": lambda w: np.sum(w) - 1}]

        # Max Sharpe Ratio
        res_msr = minimize(
            _neg_sharpe, np.repeat(1/n, n), args=(mu, cov),
            method="SLSQP", bounds=bounds, constraints=constraints,
            options={"disp": False}
        )
        w_msr = res_msr.x if res_msr.success else np.repeat(1/n, n)

        # Global Minimum Variance
        res_gmv = minimize(
            lambda w: _portfolio_vol(w, cov), np.repeat(1/n, n),
            method="SLSQP", bounds=bounds, constraints=constraints,
            options={"disp": False}
        )
        w_gmv = res_gmv.x if res_gmv.success else np.repeat(1/n, n)

        # Equal Weight
        w_ew = np.repeat(1.0 / n, n)

        def portfolio_summary(w):
            r = _portfolio_return(w, mu)
            v = _portfolio_vol(w, cov)
            return {
                "weights": dict(zip(available, w.round(4).tolist())),
                "return": round(r * 100, 2),
                "vol": round(v * 100, 2),
                "sharpe": round((r - RF) / v if v > 0 else 0, 3),
            }

        return {
            "tickers": available,
            "msr": portfolio_summary(w_msr),
            "gmv": portfolio_summary(w_gmv),
            "equal": portfolio_summary(w_ew),
        }

    except Exception as e:
        print(f"[engine] Optimization error: {e}")
        return {}
