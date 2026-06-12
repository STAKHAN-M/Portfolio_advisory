"""
macro_engine.py — Macroeconomic data fetcher
Sources:
  - FRED API      (US, France, Allemagne) — nécessite FRED_API_KEY
  - Eurostat API  (chômage mensuel FR/DE) — sans clé
  - World Bank API (tous les pays, sans clé) — données annuelles / fallback
Cache: .macro_cache.json, TTL 24h
Historique: depuis 2000 (2 cycles complets)
"""

import os
import json
import time
import requests
import pandas as pd
import numpy as np
from pathlib import Path

# ─── Config ───────────────────────────────────────────────────────────────────

CACHE_PATH    = Path(__file__).parent / ".macro_cache.json"
CACHE_TTL     = 24 * 3600
FRED_URL      = "https://api.stlouisfed.org/fred/series/observations"
WB_URL        = "https://api.worldbank.org/v2/country/{iso2}/indicator/{ind}"
EUROSTAT_URL  = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/UNE_RT_M"
HISTORY_START = "2000-01-01"

# ─── FRED — séries par pays ───────────────────────────────────────────────────

FRED_SERIES = {
    "US": {
        "gdp_qq":    "A191RL1Q225SBEA",  # PIB réel QoQ annualisé (%)
        "unemp":     "UNRATE",            # Taux de chômage mensuel
        "cpi":       "CPIAUCSL",          # Indice CPI → calcul YoY
        "yield10":   "DGS10",             # Taux 10 ans Treasury
        "rate_short":"FEDFUNDS",          # Fed Funds Rate (taux court)
    },
    "DE": {
        "gdp_qq":    "NAEXKP01DEQ657S",  # PIB Allemagne QoQ (OCDE via FRED)
        "unemp":     "LMUNRRTTDEM156S",   # Chômage harmonisé mensuel (fallback FRED)
        "cpi":       "CP0000DEM086NEST",  # HICP Allemagne (Eurostat via FRED) → YoY
        "yield10":   "IRLTLT01DEM156N",   # Bund 10 ans
        "rate_short":"ECBDFR",            # ECB Deposit Facility Rate
        "unemp_eurostat": "DE",           # Source principale chômage
    },
    "FR": {
        "gdp_qq":    "NAEXKP01FRQ657S",  # PIB France QoQ
        "unemp":     "LRUNTTTTFRQ156S",   # Chômage France trimestriel (fallback)
        "cpi":       "CP0000FRM086NEST",  # HICP France (Eurostat via FRED) → YoY
        "yield10":   "IRLTLT01FRM156N",   # OAT 10 ans
        "rate_short":"ECBDFR",            # ECB Deposit Facility Rate
        "unemp_eurostat": "FR",           # Source principale chômage
    },
}

EUROSTAT_UNEMP_PRIORITY = {"FR", "DE"}
EUROSTAT_HICP_PRIORITY  = {"FR", "DE"}
EUROSTAT_HICP_URL = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/prc_hicp_manr"

# ─── World Bank — indicateurs ─────────────────────────────────────────────────

WB_INDICATORS = {
    "gdp_growth": "NY.GDP.MKTP.KD.ZG",
    "unemp":      "SL.UEM.TOTL.ZS",
    "inflation":  "FP.CPI.TOTL.ZG",
}

# Taux (FRED) pour les pays couverts par la World Bank (CN/IN)
WB_FRED_RATES = {
    # 10Y chinois indisponible sur FRED -> None
    "CN": {"rate_short": "IR3TIB01CNM156N", "yield10": None},
    "IN": {"rate_short": "INDIR3TIB01STM", "yield10": "INDIRLTLT01STM"},
}

# Compléments FRED plus frais que l'annuel World Bank (séries = taux YoY déjà calculés)
WB_FRED_SUPPLEMENT = {
    "CN": {"inflation": "CPALTT01CNM659N"},                          # IPC YoY mensuel
    "IN": {"inflation": "CPALTT01INM659N", "gdp_yoy": "INDGDPRQPSMEI"},  # IPC + PIB YoY trim.
}

# ─── Métadonnées pays ─────────────────────────────────────────────────────────

COUNTRY_META = {
    "US": {"label": "États-Unis", "flag": "🇺🇸", "wb": "US"},
    "DE": {"label": "Allemagne",  "flag": "🇩🇪", "wb": "DE"},
    "FR": {"label": "France",     "flag": "🇫🇷", "wb": "FR"},
    "CN": {"label": "Chine",      "flag": "🇨🇳", "wb": "CN"},
    "IN": {"label": "Inde",       "flag": "🇮🇳", "wb": "IN"},
}

CYCLE_PHASES = {
    "expansion":      {"label": "Expansion",      "color": "#00c896", "icon": "↗"},
    "surchauffe":     {"label": "Surchauffe",      "color": "#f59e0b", "icon": "🔥"},
    "ralentissement": {"label": "Ralentissement",  "color": "#f04f6a", "icon": "↘"},
    "recession":      {"label": "Récession",       "color": "#dc2626", "icon": "↙"},
    "reprise":        {"label": "Reprise",         "color": "#60a5fa", "icon": "↗"},
    "inconnu":        {"label": "Indéterminé",     "color": "#64748b", "icon": "?"},
}

PHASE_BG_COLORS = {
    "expansion":      "rgba(0,200,150,0.13)",
    "surchauffe":     "rgba(245,158,11,0.13)",
    "ralentissement": "rgba(240,79,106,0.10)",
    "recession":      "rgba(220,38,38,0.15)",
    "reprise":        "rgba(96,165,250,0.13)",
    "inconnu":        "rgba(0,0,0,0)",
}

# ─── Cycle économique : séquence & secteurs ────────────────────────────────────

# Ordre canonique de l'horloge des cycles (pour anticiper la phase suivante)
NEXT_PHASE = {
    "reprise":        "expansion",
    "expansion":      "surchauffe",
    "surchauffe":     "ralentissement",
    "ralentissement": "recession",
    "recession":      "reprise",
    "inconnu":        "inconnu",
}

# GICS (libellé yfinance anglais) → secteur français
SECTOR_FR = {
    "Technology":             "Technologie",
    "Information Technology":  "Technologie",
    "Financial Services":     "Finance",
    "Financials":             "Finance",
    "Consumer Cyclical":      "Consommation discrétionnaire",
    "Consumer Discretionary": "Consommation discrétionnaire",
    "Consumer Defensive":     "Consommation de base",
    "Consumer Staples":       "Consommation de base",
    "Industrials":            "Industrie",
    "Basic Materials":        "Matériaux",
    "Materials":              "Matériaux",
    "Energy":                 "Énergie",
    "Healthcare":             "Santé",
    "Health Care":            "Santé",
    "Utilities":              "Services publics",
    "Real Estate":            "Immobilier",
    "Communication Services": "Communication",
}

# Secteurs favorisés par phase (modèle "investment clock")
PHASE_SECTORS = {
    "reprise":        ["Consommation discrétionnaire", "Technologie", "Finance",
                       "Industrie", "Immobilier"],
    "expansion":      ["Technologie", "Industrie", "Matériaux", "Communication"],
    "surchauffe":     ["Énergie", "Matériaux"],
    "ralentissement": ["Santé", "Consommation de base", "Services publics", "Énergie"],
    "recession":      ["Consommation de base", "Services publics", "Santé"],
    "inconnu":        [],
}

# ─── Cache ────────────────────────────────────────────────────────────────────

def _load_cache() -> dict | None:
    try:
        if CACHE_PATH.exists():
            raw = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
            if time.time() - raw.get("_ts", 0) < CACHE_TTL:
                return raw
    except Exception:
        pass
    return None


def _save_cache(data: dict) -> None:
    try:
        data["_ts"] = time.time()
        CACHE_PATH.write_text(json.dumps(data, default=str), encoding="utf-8")
    except Exception:
        pass


# ─── Eurostat helper ──────────────────────────────────────────────────────────

def _fetch_eurostat_unemployment(geo: str) -> pd.Series:
    """Taux de chômage mensuel harmonisé (ILO) depuis Eurostat."""
    try:
        resp = requests.get(
            EUROSTAT_URL,
            params={"geo": geo, "sex": "T", "age": "TOTAL",
                    "unit": "PC_ACT", "s_adj": "SA", "lang": "en"},
            timeout=15,
        )
        if resp.status_code != 200:
            return pd.Series(dtype=float)
        data   = resp.json()
        t_idx  = data["dimension"]["time"]["category"]["index"]
        values = data.get("value", {})
        rows   = {
            pd.Timestamp(k + "-01"): float(values[str(v)])
            for k, v in t_idx.items()
            if str(v) in values and values[str(v)] is not None
        }
        s = pd.Series(rows, dtype=float).sort_index()
        return s[s.index >= HISTORY_START]
    except Exception as e:
        print(f"[macro] Eurostat unemployment {geo} error: {e}")
        return pd.Series(dtype=float)


def _fetch_eurostat_hicp(geo: str) -> pd.Series:
    """Inflation HICP YoY mensuelle (taux de variation annuel) depuis Eurostat."""
    try:
        resp = requests.get(
            EUROSTAT_HICP_URL,
            params={"geo": geo, "coicop": "CP00", "unit": "RCH_A", "lang": "en"},
            timeout=15,
        )
        if resp.status_code != 200:
            return pd.Series(dtype=float)
        data   = resp.json()
        t_idx  = data["dimension"]["time"]["category"]["index"]
        values = data.get("value", {})
        rows   = {
            pd.Timestamp(k + "-01"): float(values[str(v)])
            for k, v in t_idx.items()
            if str(v) in values and values[str(v)] is not None
        }
        s = pd.Series(rows, dtype=float).sort_index()
        return s[s.index >= HISTORY_START]
    except Exception as e:
        print(f"[macro] Eurostat HICP {geo} error: {e}")
        return pd.Series(dtype=float)


# ─── FRED helpers ─────────────────────────────────────────────────────────────

def _fred_fetch(series_id: str, api_key: str, limit: int = 100000) -> pd.Series:
    """Retourne une pd.Series (index DatetimeIndex, valeurs float) ou Series vide."""
    try:
        resp = requests.get(
            FRED_URL,
            params={
                "series_id":        series_id,
                "api_key":          api_key,
                "file_type":        "json",
                "sort_order":       "asc",
                "limit":            limit,
                "observation_start": HISTORY_START,
            },
            timeout=10,
        )
        if resp.status_code != 200:
            return pd.Series(dtype=float)
        obs  = resp.json().get("observations", [])
        rows = {
            pd.Timestamp(o["date"]): float(o["value"])
            for o in obs
            if o["value"] not in (".", "")
        }
        return pd.Series(rows, dtype=float).sort_index()
    except Exception as e:
        print(f"[macro] FRED {series_id} error: {e}")
        return pd.Series(dtype=float)


def _yoy_from_index(s: pd.Series) -> pd.Series:
    """Calcule le YoY (%) à partir d'un indice de prix mensuel."""
    s   = s.resample("MS").last().dropna()
    yoy = s.pct_change(12) * 100
    return yoy.dropna()


def _latest(s: pd.Series, n: int = 1):
    """Dernière(s) valeur(s) non-NaN."""
    clean = s.dropna()
    if clean.empty:
        return None if n == 1 else []
    if n == 1:
        return round(float(clean.iloc[-1]), 2)
    return [round(float(v), 2) for v in clean.iloc[-n:]]


def _trend(s: pd.Series, periods: int = 3) -> float | None:
    """Variation entre la dernière valeur et la moyenne des `periods` précédents."""
    clean = s.dropna()
    if len(clean) < periods + 1:
        return None
    current = float(clean.iloc[-1])
    prior   = float(clean.iloc[-periods - 1:-1].mean())
    return round(current - prior, 3)


def _series_to_records(s: pd.Series) -> list:
    """Sérialise une pd.Series en liste [[date_str, value], ...]."""
    clean = s.dropna()
    return [[str(idx.date()), round(float(v), 3)] for idx, v in clean.items()]


# ─── Phase history ────────────────────────────────────────────────────────────

def detect_cycle(gdp_trend, unemp_trend, inflation, gdp_level=None) -> str:
    if gdp_trend is None or unemp_trend is None:
        return "inconnu"
    accelerating = gdp_trend >= 0
    hiring       = unemp_trend <= 0
    hot_infl     = (inflation or 0) > 3.5
    # Une vraie récession exige une croissance faible/négative en niveau,
    # pas seulement une décélération (ex. Inde à +6 % ne doit pas être "récession").
    contracting  = gdp_level is not None and gdp_level <= 0.0

    if accelerating and hiring:
        return "surchauffe" if hot_infl else "expansion"
    if accelerating and not hiring:
        return "reprise"
    if not accelerating and not hiring:
        return "recession" if contracting else "ralentissement"
    return "ralentissement"


def _phases_to_spans(phase_s: pd.Series) -> list:
    """Regroupe une série de phases en plages [{start, end, phase}]."""
    if phase_s.empty:
        return []
    spans = []
    dates  = list(phase_s.index)
    vals   = list(phase_s.values)
    cur    = vals[0]
    start  = dates[0]
    for i in range(1, len(dates)):
        if vals[i] != cur:
            spans.append({"start": str(start.date()), "end": str(dates[i].date()), "phase": cur})
            cur   = vals[i]
            start = dates[i]
    spans.append({"start": str(start.date()), "end": str(dates[-1].date()), "phase": cur})
    return spans


def compute_phase_history(gdp_s: pd.Series, unemp_s: pd.Series,
                           infl_s: pd.Series) -> list:
    """Calcule les phases du cycle sur toute la période historique."""
    if gdp_s.empty or unemp_s.empty:
        return []
    try:
        gdp_m   = gdp_s.resample("MS").first().ffill()
        unemp_m = unemp_s.resample("MS").last().ffill()
        infl_m  = infl_s.resample("MS").last().ffill() if not infl_s.empty else pd.Series(dtype=float)

        start = max(gdp_m.index.min(), unemp_m.index.min())
        end   = min(gdp_m.index.max(), unemp_m.index.max())
        dates = pd.date_range(start, end, freq="MS")
        if len(dates) < 8:
            return []

        gdp_m   = gdp_m.reindex(dates).ffill()
        unemp_m = unemp_m.reindex(dates).ffill()
        if not infl_m.empty:
            infl_m = infl_m.reindex(dates).ffill().fillna(2.0)
        else:
            infl_m = pd.Series(2.0, index=dates)

        window = 4
        phases = {}
        for i, date in enumerate(dates):
            if i < window:
                phases[date] = "inconnu"
                continue
            g_trend = float(gdp_m.iloc[i])   - float(gdp_m.iloc[i - window:i].mean())
            u_trend = float(unemp_m.iloc[i]) - float(unemp_m.iloc[max(0, i - 3)])
            infl    = float(infl_m.iloc[i])
            phases[date] = detect_cycle(g_trend, u_trend, infl,
                                        gdp_level=float(gdp_m.iloc[i]))

        return _phases_to_spans(pd.Series(phases))
    except Exception as e:
        print(f"[macro] compute_phase_history error: {e}")
        return []


# ─── Fetch FRED country ───────────────────────────────────────────────────────

def _fetch_fred_country(code: str, api_key: str) -> dict:
    series = FRED_SERIES[code]
    result = {}

    # ── PIB trimestriel ──────────────────────────────────────────────────────
    gdp_s = _fred_fetch(series["gdp_qq"], api_key)
    result["gdp_qq_latest"]  = _latest(gdp_s)
    result["gdp_qq_prev"]    = _latest(gdp_s.dropna().iloc[:-1]) if len(gdp_s.dropna()) >= 2 else None
    result["gdp_qq_hist"]    = _latest(gdp_s, n=8)
    result["gdp_qq_date"]    = str(gdp_s.dropna().index[-1].date()) if not gdp_s.empty else None
    result["gdp_trend"]      = _trend(gdp_s, 2)
    result["gdp_qq_series"]  = _series_to_records(gdp_s)          # SÉRIE COMPLÈTE

    # YoY approximatif (moyenne mobile 4 trimestres)
    gdp_yoy = gdp_s.rolling(4).mean() if not gdp_s.empty else gdp_s
    result["gdp_yoy_latest"] = _latest(gdp_yoy)

    # ── Chômage ──────────────────────────────────────────────────────────────
    eurostat_geo = series.get("unemp_eurostat")
    if code in EUROSTAT_UNEMP_PRIORITY and eurostat_geo:
        unemp_s   = _fetch_eurostat_unemployment(eurostat_geo)
        unemp_src = "Eurostat"
    else:
        unemp_s   = pd.Series(dtype=float)
        unemp_src = "FRED"
    if unemp_s.empty:
        unemp_s   = _fred_fetch(series["unemp"], api_key)
        unemp_src = "FRED"

    result["unemp_latest"]   = _latest(unemp_s)
    result["unemp_hist"]     = _latest(unemp_s, n=12)
    result["unemp_date"]     = str(unemp_s.dropna().index[-1].date()) if not unemp_s.empty else None
    result["unemp_trend"]    = _trend(unemp_s, 3)
    result["unemp_source"]   = unemp_src
    result["unemp_series"]   = _series_to_records(unemp_s)        # SÉRIE COMPLÈTE

    # ── Inflation YoY ────────────────────────────────────────────────────────
    # FR/DE : indice HICP Eurostat-via-FRED (frais, ~M-2) ; repli sur l'API Eurostat.
    cpi_s    = _fred_fetch(series["cpi"], api_key)
    infl_s   = _yoy_from_index(cpi_s)
    infl_src = "FRED (HICP)" if code in EUROSTAT_HICP_PRIORITY else "FRED"
    eurostat_geo = series.get("unemp_eurostat")
    if infl_s.empty and code in EUROSTAT_HICP_PRIORITY and eurostat_geo:
        infl_s   = _fetch_eurostat_hicp(eurostat_geo)
        infl_src = "Eurostat HICP"
    result["inflation_latest"]  = _latest(infl_s)
    result["inflation_hist"]    = _latest(infl_s, n=12)
    result["inflation_date"]    = str(infl_s.dropna().index[-1].date()) if not infl_s.empty else None
    result["inflation_trend"]   = _trend(infl_s, 3)
    result["inflation_source"]  = infl_src
    result["inflation_series"]  = _series_to_records(infl_s)      # SÉRIE COMPLÈTE

    # ── Taux 10 ans ──────────────────────────────────────────────────────────
    y10_s  = _fred_fetch(series["yield10"], api_key)
    # Resample mensuel pour lisser les données journalières (US DGS10)
    y10_m  = y10_s.resample("MS").mean() if not y10_s.empty else y10_s
    result["yield10_latest"]    = _latest(y10_m)
    result["yield10_hist"]      = _latest(y10_m, n=12)
    result["yield10_date"]      = str(y10_m.dropna().index[-1].date()) if not y10_m.empty else None
    result["yield10_trend"]     = _trend(y10_m, 3)
    result["yield10_series"]    = _series_to_records(y10_m)       # SÉRIE COMPLÈTE

    # ── Taux court (politique monétaire) ─────────────────────────────────────
    rs_id  = series.get("rate_short")
    if rs_id:
        rs_s  = _fred_fetch(rs_id, api_key)
        rs_m  = rs_s.resample("MS").mean() if not rs_s.empty else rs_s
        result["rate_short_latest"] = _latest(rs_m)
        result["rate_short_date"]   = str(rs_m.dropna().index[-1].date()) if not rs_m.empty else None
        result["rate_short_trend"]  = _trend(rs_m, 3)
        result["rate_short_series"] = _series_to_records(rs_m)    # SÉRIE COMPLÈTE
    else:
        result["rate_short_latest"] = None
        result["rate_short_date"]   = None
        result["rate_short_trend"]  = None
        result["rate_short_series"] = []

    # ── Phase history ─────────────────────────────────────────────────────────
    result["phase_spans"] = compute_phase_history(gdp_s, unemp_s, infl_s)

    # ── Cycle actuel ─────────────────────────────────────────────────────────
    result["cycle"]  = detect_cycle(result["gdp_trend"], result["unemp_trend"],
                                    result["inflation_latest"],
                                    gdp_level=result["gdp_qq_latest"])
    result["source"] = "FRED"
    return result


# ─── World Bank ───────────────────────────────────────────────────────────────

def _fetch_wb_country(iso2: str, code: str = None, fred_api_key: str = None) -> dict:
    result = {}
    try:
        def _wb(ind, mrv=25):
            url  = WB_URL.format(iso2=iso2, ind=ind)
            r    = requests.get(url, params={"format": "json", "mrv": mrv,
                                             "per_page": mrv}, timeout=25)
            data = r.json()
            if not isinstance(data, list) or len(data) < 2:
                return pd.Series(dtype=float)
            rows = {
                pd.Timestamp(f"{d['date']}-01-01"): float(d["value"])
                for d in data[1]
                if d.get("value") is not None
            }
            return pd.Series(rows, dtype=float).sort_index()

        gdp_s   = _wb(WB_INDICATORS["gdp_growth"])
        unemp_s = _wb(WB_INDICATORS["unemp"])
        infl_s  = _wb(WB_INDICATORS["inflation"])

        result["gdp_qq_latest"]     = None
        result["gdp_yoy_latest"]    = _latest(gdp_s)
        result["gdp_qq_hist"]       = _latest(gdp_s, n=8)
        result["gdp_qq_date"]       = str(gdp_s.dropna().index[-1].date()) if not gdp_s.empty else None
        result["gdp_trend"]         = _trend(gdp_s, 2)
        result["gdp_qq_series"]     = _series_to_records(gdp_s)

        result["unemp_latest"]      = _latest(unemp_s)
        result["unemp_hist"]        = _latest(unemp_s, n=8)
        result["unemp_date"]        = str(unemp_s.dropna().index[-1].date()) if not unemp_s.empty else None
        result["unemp_trend"]       = _trend(unemp_s, 2)
        result["unemp_source"]      = "World Bank (annuel)"
        result["unemp_series"]      = _series_to_records(unemp_s)

        result["inflation_latest"]  = _latest(infl_s)
        result["inflation_hist"]    = _latest(infl_s, n=8)
        result["inflation_date"]    = str(infl_s.dropna().index[-1].date()) if not infl_s.empty else None
        result["inflation_trend"]   = _trend(infl_s, 2)
        result["inflation_source"]  = "World Bank (annuel)"
        result["inflation_series"]  = _series_to_records(infl_s)
        result["gdp_source"]        = "World Bank (annuel)"

        # ── Compléments FRED plus frais (inflation mensuelle, PIB trimestriel) ────
        supp = WB_FRED_SUPPLEMENT.get(code or "", {})
        if supp and fred_api_key:
            infl_id = supp.get("inflation")
            if infl_id:
                infl_m = _fred_fetch(infl_id, fred_api_key)  # déjà en % YoY
                if not infl_m.dropna().empty:
                    infl_s = infl_m
                    result["inflation_latest"] = _latest(infl_m)
                    result["inflation_hist"]   = _latest(infl_m, n=12)
                    result["inflation_date"]   = str(infl_m.dropna().index[-1].date())
                    result["inflation_trend"]  = _trend(infl_m, 3)
                    result["inflation_source"] = "FRED (mensuel)"
                    result["inflation_series"] = _series_to_records(infl_m)
            gdp_id = supp.get("gdp_yoy")
            if gdp_id:
                gdp_q = _fred_fetch(gdp_id, fred_api_key)    # PIB YoY trimestriel
                if not gdp_q.dropna().empty:
                    gdp_s = gdp_q
                    result["gdp_yoy_latest"] = _latest(gdp_q)
                    result["gdp_qq_hist"]    = _latest(gdp_q, n=8)
                    result["gdp_qq_date"]    = str(gdp_q.dropna().index[-1].date())
                    result["gdp_trend"]      = _trend(gdp_q, 2)
                    result["gdp_qq_series"]  = _series_to_records(gdp_q)
                    result["gdp_source"]     = "FRED (trimestriel)"

        # Valeurs par défaut (taux indisponibles)
        result["yield10_latest"]    = None
        result["yield10_hist"]      = []
        result["yield10_date"]      = None
        result["yield10_trend"]     = None
        result["yield10_series"]    = []
        result["rate_short_latest"] = None
        result["rate_short_date"]   = None
        result["rate_short_trend"]  = None
        result["rate_short_series"] = []

        # ── Taux via FRED (CN/IN) ────────────────────────────────────────────
        rate_cfg = WB_FRED_RATES.get(code or "", {})
        if rate_cfg and fred_api_key:
            y10_id = rate_cfg.get("yield10")
            if y10_id:
                y10_s = _fred_fetch(y10_id, fred_api_key)
                y10_m = y10_s.resample("MS").mean() if not y10_s.empty else y10_s
                if not y10_m.dropna().empty:
                    result["yield10_latest"] = _latest(y10_m)
                    result["yield10_hist"]   = _latest(y10_m, n=12)
                    result["yield10_date"]   = str(y10_m.dropna().index[-1].date())
                    result["yield10_trend"]  = _trend(y10_m, 3)
                    result["yield10_series"] = _series_to_records(y10_m)
            rs_id = rate_cfg.get("rate_short")
            if rs_id:
                rs_s = _fred_fetch(rs_id, fred_api_key)
                rs_m = rs_s.resample("MS").mean() if not rs_s.empty else rs_s
                if not rs_m.dropna().empty:
                    result["rate_short_latest"] = _latest(rs_m)
                    result["rate_short_date"]   = str(rs_m.dropna().index[-1].date())
                    result["rate_short_trend"]  = _trend(rs_m, 3)
                    result["rate_short_series"] = _series_to_records(rs_m)

        result["phase_spans"] = compute_phase_history(gdp_s, unemp_s, infl_s)
        result["cycle"]       = detect_cycle(result["gdp_trend"], result["unemp_trend"],
                                             result["inflation_latest"],
                                             gdp_level=result["gdp_yoy_latest"])
        result["source"] = "World Bank (annuel)"
    except Exception as e:
        print(f"[macro] World Bank {iso2} error: {e}")
    return result


# ─── Entry point ──────────────────────────────────────────────────────────────

def get_macro_data(fred_api_key: str = None) -> dict:
    """Retourne un dict {code_pays: {indicateurs…}, '_fetched_at': str}."""
    cached = _load_cache()
    if cached:
        cached.pop("_ts", None)
        return cached

    if not fred_api_key:
        fred_api_key = os.environ.get("FRED_API_KEY", "")

    result = {}
    for code, meta in COUNTRY_META.items():
        try:
            if code in FRED_SERIES and fred_api_key:
                data = _fetch_fred_country(code, fred_api_key)
            else:
                data = _fetch_wb_country(meta["wb"], code=code, fred_api_key=fred_api_key)
            result[code] = data
        except Exception as e:
            print(f"[macro] {code} fetch failed: {e}")
            result[code] = {"source": "erreur", "cycle": "inconnu",
                            "gdp_qq_series": [], "unemp_series": [],
                            "inflation_series": [], "yield10_series": [],
                            "rate_short_series": [], "phase_spans": []}

    result["_fetched_at"] = time.strftime("%d/%m/%Y %H:%M")
    _save_cache(result)
    result.pop("_ts", None)
    return result


# ─── Spread de taux (signal de récession) ──────────────────────────────────────

def compute_yield_spread(yield10_records: list, rate_short_records: list) -> dict:
    """Spread 10 ans − taux court. Inversion (négatif) = signal avancé de récession."""
    if not yield10_records or not rate_short_records:
        return {"series": [], "latest": None, "signal": "n/a", "label": "Données indisponibles"}
    s10 = pd.Series({pd.Timestamp(d): v for d, v in yield10_records}).sort_index()
    srs = pd.Series({pd.Timestamp(d): v for d, v in rate_short_records}).sort_index()
    df  = pd.concat([s10.rename("y10"), srs.rename("rs")], axis=1).dropna()
    if df.empty:
        return {"series": [], "latest": None, "signal": "n/a", "label": "Données indisponibles"}
    spread = (df["y10"] - df["rs"]).round(3)
    latest = float(spread.iloc[-1])
    if latest < 0:
        signal, label = "inverse", "Courbe inversée — signal de récession"
    elif latest < 0.5:
        signal, label = "plat", "Courbe aplatie — vigilance"
    else:
        signal, label = "normal", "Courbe normale"
    return {
        "series": [[str(idx.date()), round(float(v), 3)] for idx, v in spread.items()],
        "latest": round(latest, 2),
        "signal": signal,
        "label": label,
    }


# ─── Secteurs des titres (yfinance) ─────────────────────────────────────────────

_SECTOR_CACHE = Path(__file__).parent / ".sector_cache.json"
_SECTOR_TTL   = 30 * 24 * 3600  # 30 jours

def get_ticker_sectors(tickers: list) -> dict:
    """Retourne {ticker: secteur_fr} via yfinance, avec cache 30 j. ETF/indices → None."""
    tickers = [t for t in dict.fromkeys(tickers) if t]
    cache = {}
    try:
        if _SECTOR_CACHE.exists():
            raw = json.loads(_SECTOR_CACHE.read_text(encoding="utf-8"))
            if time.time() - raw.get("_ts", 0) < _SECTOR_TTL:
                cache = raw.get("data", {})
    except Exception:
        cache = {}

    missing = [t for t in tickers if t not in cache]
    if missing:
        try:
            import yfinance as yf
            for t in missing:
                sec_fr = None
                try:
                    info = yf.Ticker(t).info or {}
                    sec_en = info.get("sector")
                    if sec_en:
                        sec_fr = SECTOR_FR.get(sec_en, sec_en)
                except Exception:
                    sec_fr = None
                cache[t] = sec_fr
            _SECTOR_CACHE.write_text(
                json.dumps({"_ts": time.time(), "data": cache}, default=str),
                encoding="utf-8")
        except Exception as e:
            print(f"[macro] sector fetch error: {e}")

    return {t: cache.get(t) for t in tickers}


# ─── FX & matières premières (yfinance) ─────────────────────────────────────────

_MARKET_CACHE = Path(__file__).parent / ".market_cache.json"
_MARKET_TTL   = 6 * 3600  # 6 h

MARKET_TICKERS = {
    "EURUSD": {"yf": "EURUSD=X", "label": "EUR / USD",        "unit": "",   "icon": "💱"},
    "GOLD":   {"yf": "GC=F",     "label": "Or (once)",         "unit": " $", "icon": "🥇"},
    "BRENT":  {"yf": "BZ=F",     "label": "Pétrole Brent",     "unit": " $", "icon": "🛢️"},
}

def fetch_market_data() -> dict:
    """Niveaux mensuels EUR/USD, Or, Brent (yfinance) + variation YoY. Cache 6 h."""
    try:
        if _MARKET_CACHE.exists():
            raw = json.loads(_MARKET_CACHE.read_text(encoding="utf-8"))
            if time.time() - raw.get("_ts", 0) < _MARKET_TTL:
                return raw.get("data", {})
    except Exception:
        pass

    out = {}
    try:
        import yfinance as yf
        for key, meta in MARKET_TICKERS.items():
            try:
                df = yf.download(meta["yf"], period="6y", interval="1mo",
                                 progress=False, auto_adjust=True)
                if df is None or df.empty:
                    out[key] = {"latest": None, "yoy": None, "series": []}
                    continue
                close = df["Close"]
                if hasattr(close, "columns"):
                    close = close.iloc[:, 0]
                close = close.dropna()
                s = pd.Series(close.values, index=pd.to_datetime(close.index))
                latest = float(s.iloc[-1])
                yoy = None
                if len(s) > 12:
                    prior = float(s.iloc[-13])
                    if prior:
                        yoy = round((latest / prior - 1) * 100, 1)
                out[key] = {
                    "latest": round(latest, 4 if key == "EURUSD" else 2),
                    "yoy": yoy,
                    "series": [[str(idx.date()), round(float(v), 4)] for idx, v in s.items()],
                }
            except Exception as e:
                print(f"[macro] market {key} error: {e}")
                out[key] = {"latest": None, "yoy": None, "series": []}
        _MARKET_CACHE.write_text(json.dumps({"_ts": time.time(), "data": out}, default=str),
                                 encoding="utf-8")
    except Exception as e:
        print(f"[macro] market fetch error: {e}")
    return out
