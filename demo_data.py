"""
Portefeuille de démonstration.

Génère un historique de transactions fictif mais cohérent, au format attendu par
`engine.build_positions_table` (feuille BDD). Les tickers sont réels afin que les
cours actuels soient récupérés par yfinance : les plus-values latentes affichées
sont donc de vraies variations de marché appliquées à des achats fictifs.

Aucune donnée personnelle n'est impliquée ; ce module ne lit aucun fichier.
"""

import math

import pandas as pd

# Prix d'achat/vente plausibles aux dates indiquées (EUR/USD selon la place).
# Format : (date, type, ticker, libellé, quantité, prix unitaire)
_TRADES = [
    # ── 2023 : constitution du socle ─────────────────────────────────────────
    ("2023-01-16", "Achat", "CW8.PA",  "ETF MSCI World",        8,  380.00),
    ("2023-02-13", "Achat", "MC.PA",   "LVMH",                  2,  760.00),
    ("2023-03-20", "Achat", "TTE.PA",  "TotalEnergies",        40,   55.00),
    ("2023-04-17", "Achat", "ORA.PA",  "Orange",              280,   11.20),
    ("2023-05-15", "Achat", "SAN.PA",  "Sanofi",               22,   92.00),
    ("2023-06-19", "Achat", "DG.PA",   "Vinci",                18,  103.00),
    ("2023-09-18", "Achat", "AIR.PA",  "Airbus",               14,  126.00),
    ("2023-11-20", "Achat", "CW8.PA",  "ETF MSCI World",        6,  400.00),

    # ── 2024 : renforcement + diversification US ─────────────────────────────
    ("2024-01-15", "Achat", "AAPL",    "Apple",                12,  182.00),
    ("2024-02-19", "Achat", "OR.PA",   "L'Oreal",               5,  425.00),
    ("2024-03-18", "Achat", "MSFT",    "Microsoft",             6,  405.00),
    ("2024-05-20", "Achat", "CW8.PA",  "ETF MSCI World",        6,  455.00),
    ("2024-06-17", "Achat", "BNP.PA",  "BNP Paribas",          40,   58.00),
    ("2024-09-16", "Achat", "ASML.AS", "ASML",                  2,  700.00),
    ("2024-10-21", "Achat", "AIR.PA",  "Airbus",                8,  148.00),

    # ── 2025 : arbitrages (2 positions soldées) ──────────────────────────────
    ("2025-02-17", "Vente", "ORA.PA",  "Arbitrage Orange",    280,   10.45),
    ("2025-02-24", "Achat", "CW8.PA",  "ETF MSCI World",        6,  505.00),
    ("2025-04-14", "Achat", "TTE.PA",  "TotalEnergies",        20,   53.50),
    ("2025-06-16", "Vente", "DG.PA",   "Prise de benefice",    18,  118.00),
    ("2025-06-23", "Achat", "MSFT",    "Microsoft",             4,  455.00),
    ("2025-09-15", "Achat", "SAN.PA",  "Sanofi",               10,   98.00),
    ("2025-11-17", "Achat", "CW8.PA",  "ETF MSCI World",        5,  545.00),

    # ── 2026 : poursuite du plan d'investissement ────────────────────────────
    ("2026-01-19", "Achat", "BNP.PA",  "BNP Paribas",          25,   74.00),
    ("2026-02-16", "Achat", "AAPL",    "Apple",                 6,  242.00),
    ("2026-04-20", "Achat", "CW8.PA",  "ETF MSCI World",        5,  575.00),
    ("2026-06-15", "Achat", "AIR.PA",  "Airbus",                6,  191.00),
]

# Dividendes encaissés : (date, ticker, libellé, montant)
_DIVIDENDS = [
    ("2023-05-25", "TTE.PA",  "Acompte TotalEnergies",  30.00),
    ("2023-06-08", "ORA.PA",  "Dividende Orange",      196.00),
    ("2023-07-20", "SAN.PA",  "Dividende Sanofi",       81.40),
    ("2023-10-05", "TTE.PA",  "Acompte TotalEnergies",  31.20),
    ("2024-04-25", "DG.PA",   "Dividende Vinci",        82.80),
    ("2024-05-23", "MC.PA",   "Dividende LVMH",         26.00),
    ("2024-06-06", "ORA.PA",  "Dividende Orange",      196.00),
    ("2024-07-18", "SAN.PA",  "Dividende Sanofi",       82.50),
    ("2024-10-03", "TTE.PA",  "Acompte TotalEnergies",  32.00),
    ("2024-11-14", "BNP.PA",  "Acompte BNP Paribas",    62.00),
    ("2025-04-24", "DG.PA",   "Dividende Vinci",        84.60),
    ("2025-05-22", "MC.PA",   "Dividende LVMH",         26.00),
    ("2025-06-05", "BNP.PA",  "Dividende BNP Paribas",  92.00),
    ("2025-07-17", "SAN.PA",  "Dividende Sanofi",       114.40),
    ("2025-10-02", "TTE.PA",  "Acompte TotalEnergies",  50.40),
    ("2026-05-21", "MC.PA",   "Dividende LVMH",         28.00),
    ("2026-06-04", "BNP.PA",  "Dividende BNP Paribas",  143.00),
    ("2026-07-16", "SAN.PA",  "Dividende Sanofi",      118.40),
]

_COLUMNS = ["Date", "Type", "Ticker", "Description_Operation",
            "Quantite", "Prix_Unitaire", "Montant_Total", "Cash_Flow"]


def _row(date, typ, ticker, label, qty, price, total, cash):
    return {"Date": date, "Type": typ, "Ticker": ticker,
            "Description_Operation": label, "Quantite": qty,
            "Prix_Unitaire": price, "Montant_Total": round(total, 2),
            "Cash_Flow": round(cash, 2)}


def build_demo_transactions() -> pd.DataFrame:
    """
    Construit la feuille BDD de démonstration.

    Les versements ne sont pas codés en dur : ils sont insérés automatiquement
    avant chaque achat que la trésorerie ne couvre pas, arrondis au multiple de
    500 € supérieur. Le solde de cash reste ainsi positif à tout instant et la
    chronologie des flux est cohérente (indispensable pour le TRI/XIRR).
    """
    events = []
    for date, typ, ticker, label, qty, price in _TRADES:
        events.append((date, "trade", (typ, ticker, label, qty, price)))
    for date, ticker, label, amount in _DIVIDENDS:
        events.append((date, "div", (ticker, label, amount)))
    events.sort(key=lambda e: (e[0], 0 if e[1] == "div" else 1))

    rows, cash = [], 0.0
    for date, kind, payload in events:
        if kind == "div":
            ticker, label, amount = payload
            cash += amount
            rows.append(_row(date, "Dividende", ticker, label, 0, 0, amount, amount))
            continue

        typ, ticker, label, qty, price = payload
        total = qty * price

        if typ == "Achat":
            # Versement d'appoint si la trésorerie est insuffisante
            if cash < total:
                need = math.ceil((total - cash) / 500.0) * 500.0
                vers_date = (pd.Timestamp(date) - pd.Timedelta(days=3)).strftime("%Y-%m-%d")
                cash += need
                rows.append(_row(vers_date, "Versement", "", "Virement mensuel",
                                 0, 0, need, need))
            cash -= total
            rows.append(_row(date, "Achat", ticker, label, qty, price, total, -total))
        else:  # Vente
            cash += total
            rows.append(_row(date, "Vente", ticker, label, qty, price, total, total))

    df = pd.DataFrame(rows, columns=_COLUMNS)
    return df.sort_values("Date").reset_index(drop=True)


def build_demo_payload() -> dict:
    """
    Retourne le payload prêt pour `store-data`, au même format que celui produit
    par l'upload d'un fichier Excel. La feuille Valorisation est laissée vide :
    la courbe de performance est alors calculée en TWR depuis les transactions.
    """
    df = build_demo_transactions()
    return {
        "df": df.to_json(orient="records"),
        "val_df": "[]",
        "filename": "Portefeuille de démonstration",
        "is_demo": True,
    }
