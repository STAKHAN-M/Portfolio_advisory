"""
Portfolio SaaS Dashboard
Run: python app.py
Access: http://127.0.0.1:8050
"""

import base64
import io
import json
import os
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from dash import (
    Dash, dcc, html, dash_table, Input, Output, State,
    callback_context, no_update, ALL
)
import dash_bootstrap_components as dbc

import engine
import pdf_report
import macro_engine

# ─── Chargement clé FRED depuis .env ─────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # python-dotenv non installé : lecture manuelle du .env
    _env_path = Path(__file__).parent / ".env"
    if _env_path.exists():
        for _line in _env_path.read_text().splitlines():
            if "=" in _line and not _line.strip().startswith("#"):
                _k, _v = _line.strip().split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# ─── App Init ─────────────────────────────────────────────────────────────────

app = Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@600;700&display=swap",
    ],
    suppress_callback_exceptions=True,
    title="Portfolio Dashboard",
)
server = app.server

# ─── Color Palette ────────────────────────────────────────────────────────────

PALETTE = [
    "#c4a24a", "#00c896", "#f04f6a", "#1a2d4a", "#0d1e35",
    "#1e3a5f", "#2a4f7c", "#5b7490", "#8ba8c4", "#dde6f0",
]

def hex_rgba(hex_color, alpha):
    """Convertit '#rrggbb' (+ alpha 0-1) en 'rgba(r,g,b,a)' pour Plotly."""
    h = hex_color.lstrip("#")[:6]
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


CHART_LAYOUT = dict(
    font_family="Inter",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=16, r=16, t=30, b=16),
    hoverlabel=dict(
        bgcolor="#0f1b2d",
        bordercolor="rgba(255,255,255,0.12)",
        font=dict(family="Inter", size=12, color="#e8eef6"),
    ),
    legend=dict(
        orientation="h",
        yanchor="bottom", y=1.02,
        xanchor="left", x=0,
        font=dict(size=11, color="#5b7490"),
    ),
    xaxis=dict(
        gridcolor="rgba(255,255,255,0.04)", gridwidth=1,
        showline=False, zeroline=False,
        tickfont=dict(size=10, color="#5b7490"),
    ),
    yaxis=dict(
        gridcolor="rgba(255,255,255,0.04)", gridwidth=1,
        showline=False, zeroline=False,
        tickfont=dict(size=10, color="#5b7490"),
    ),
)

# ─── Helpers ──────────────────────────────────────────────────────────────────

def fmt_eur(v: float) -> str:
    if v is None:
        return "—"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:,.2f} €".replace(",", " ")


def fmt_pct(v: float) -> str:
    if v is None:
        return "—"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.2f} %"


def badge(value: float, fmt_fn, neutral: bool = False) -> html.Span:
    cls = "badge-neutral" if neutral else ("badge-positive" if value >= 0 else "badge-negative")
    return html.Span(fmt_fn(value), className=f"kpi-badge {cls}")


# ─── Table Style ──────────────────────────────────────────────────────────────

TABLE_STYLE_CELL = {
    "fontFamily": "Inter, sans-serif",
    "fontSize": "12px",
    "padding": "9px 14px",
    "backgroundColor": "#080e1a",
    "color": "#8ba8c4",
    "border": "none",
    "borderBottom": "1px solid rgba(255,255,255,0.04)",
    "textOverflow": "ellipsis",
    "whiteSpace": "normal",
}

TABLE_STYLE_HEADER = {
    "backgroundColor": "#04080f",
    "fontWeight": "700",
    "fontSize": "9px",
    "color": "#c4a24a",
    "textTransform": "uppercase",
    "letterSpacing": "1.5px",
    "borderBottom": "1px solid rgba(196,162,74,0.2)",
    "padding": "10px 14px",
    "border": "none",
}


# ─── Layout ───────────────────────────────────────────────────────────────────

def build_layout():
    return html.Div([

        # — Data Store —
        dcc.Store(id="store-data"),
        dcc.Store(id="store-analytics"),
        dcc.Store(id="store-macro"),
        dcc.Store(id="macro-country-store", data="US"),
        dcc.Store(id="new-ticker-store"),
        dcc.Download(id="download-pdf"),

        # — Macro indicator modal —
        dbc.Modal([
            dbc.ModalHeader(dbc.ModalTitle(id="macro-modal-title")),
            dbc.ModalBody(dcc.Graph(id="macro-modal-chart",
                                   config={"displayModeBar": True},
                                   style={"height": "380px"})),
            dbc.ModalFooter(
                dbc.Button("Fermer", id="macro-modal-close",
                           className="export-btn", n_clicks=0)
            ),
        ], id="macro-modal", size="xl", is_open=False),

        # — Navbar —
        html.Div([
            html.Div([
                html.Div([
                    html.Div("◈", style={"fontSize": "22px", "color": "#c4a24a", "lineHeight": "1"}),
                    html.Div([
                        html.Div("PORTFOLIO DASHBOARD ANALYSIS", className="navbar-title"),
                        html.Div("Analyse & Suivi de Portefeuille", className="navbar-subtitle"),
                    ]),
                ], className="navbar-brand"),
            ]),

            html.Div([
                html.Span(id="last-update-text", className="last-update"),
                html.Button("⬇  Export PDF", id="export-pdf-btn", className="export-btn", n_clicks=0),
                dcc.Upload(
                    id="upload-data",
                    children=html.Div(["↑  Importer fichier .xlsx"], className="upload-btn"),
                    accept=".xlsx",
                    multiple=False,
                ),
            ], className="navbar-right"),
        ], className="navbar"),

        # — Main Content —
        html.Div([

            # — Loading indicator —
            dcc.Loading(
                id="global-loading",
                type="circle",
                color="#c4a24a",
                children=html.Div(id="loading-trigger", style={"height": "0px"}),
            ),

            # — Empty State (shown before data load) —
            html.Div(
                id="empty-state",
                className="empty-state",
                children=[
                    html.Div("📊", className="empty-state-icon"),
                    html.Div("Aucun portefeuille chargé", className="empty-state-title"),
                    html.Div(
                        "Importez votre fichier de transactions Excel (.xlsx) "
                        "en cliquant sur le bouton ci-dessus pour initialiser le dashboard.",
                        className="empty-state-text"
                    ),
                ],
            ),

            # — Dashboard Body (hidden until data loaded) —
            html.Div(id="dashboard-body", style={"display": "none"}, children=[

                # KPI Row
                html.Div(id="kpi-row", className="kpi-row"),

                # Tabs
                dcc.Tabs(id="main-tabs", value="tab-overview", className="main-tabs", children=[

                dcc.Tab(label="Vue d'ensemble", value="tab-overview", className="main-tab", selected_className="main-tab--selected", children=[

                # Row 1: Positions table + Allocation pie
                html.Div([
                    # Positions Table
                    html.Div([
                        html.Div([
                            html.Span("Positions Actuelles", className="card-title"),
                            html.Span(id="positions-count",
                                      style={"fontSize": "11px", "color": "#94a3b8"}),
                        ], className="card-header"),
                        html.Div(
                            id="positions-table-container",
                            className="card-body",
                            style={"padding": "0"},
                        ),
                    ], className="card"),

                    # Right column: two charts stacked
                    html.Div([

                        # Allocation by ticker
                        html.Div([
                            html.Div([
                                html.Span("Répartition par Titre", className="card-title"),
                            ], className="card-header"),
                            html.Div([
                                dcc.Graph(id="pie-allocation", config={"displayModeBar": False},
                                          style={"height": "240px"}),
                            ], className="card-body"),
                        ], className="card"),

                        # Allocation by category
                        html.Div([
                            html.Div([
                                html.Span("Allocation par Catégorie", className="card-title"),
                            ], className="card-header"),
                            html.Div([
                                dcc.Graph(id="pie-category", config={"displayModeBar": False},
                                          style={"height": "200px"}),
                            ], className="card-body"),
                        ], className="card"),

                    ], style={"display": "flex", "flexDirection": "column", "gap": "0"}),

                ], className="grid-70-30"),

                # Row 2: Performance chart + Risk metrics
                html.Div([
                    # Performance Chart
                    html.Div([
                        html.Div([
                            html.Span("Performance vs Benchmark (CAC 40)", className="card-title"),
                            html.Span("Base 100 — données mensuelles",
                                      style={"fontSize": "10px", "color": "#94a3b8"}),
                        ], className="card-header"),
                        html.Div([
                            dcc.Graph(id="chart-performance", config={"displayModeBar": True},
                                      style={"height": "310px"}),
                        ], className="card-body"),
                    ], className="card"),

                    # Risk Metrics
                    html.Div([
                        html.Div([
                            html.Span("Métriques de Risque", className="card-title"),
                        ], className="card-header"),
                        html.Div(id="risk-metrics-panel", className="card-body"),
                    ], className="card"),

                ], className="grid-60-40"),

                # Row 3: Portfolio Optimization + Drawdown
                html.Div([
                    # Optimisation Table
                    html.Div([
                        html.Div([
                            html.Span("Comparaison de Portefeuilles Optimaux", className="card-title"),
                            html.Span("Taux sans risque: 3.3%",
                                      style={"fontSize": "10px", "color": "#94a3b8"}),
                        ], className="card-header"),
                        html.Div(id="optimization-table-container", className="card-body",
                                 style={"padding": "0"}),
                    ], className="card"),

                    # Drawdown chart
                    html.Div([
                        html.Div([
                            html.Span("Drawdown du Portefeuille", className="card-title"),
                        ], className="card-header"),
                        html.Div([
                            dcc.Graph(id="chart-drawdown", config={"displayModeBar": False},
                                      style={"height": "230px"}),
                        ], className="card-body"),
                    ], className="card"),

                ], className="grid-50-50"),

                # Row 3b: Attribution de performance
                html.Div([
                    html.Div([
                        html.Span("Contribution à la Performance par Titre", className="card-title"),
                        html.Span("Plus/moins-values latentes + dividendes (positions ouvertes)",
                                  style={"fontSize": "10px", "color": "#94a3b8"}),
                    ], className="card-header"),
                    html.Div(id="attribution-container", className="card-body"),
                ], className="card"),

                # Row 4: Position Calculator
                html.Div([
                    html.Div([
                        html.Div([
                            html.Span("Simulateur d'Achat & Pondération", className="card-title"),
                            html.Span("Analysez l'impact de nouveaux achats sur votre PRU et votre allocation",
                                      style={"fontSize": "10px", "color": "#94a3b8"}),
                        ], className="card-header"),
                        html.Div([
                            html.Div([
                                # Selection & Fixed Capital
                                html.Div([
                                    html.Label("Sélectionner un Actif", className="input-label"),
                                    dcc.Dropdown(id="calc-ticker-select", className="dash-dropdown"),
                                ], className="input-group"),
                                html.Div([
                                    html.Label("Valeur Actuelle du Portefeuille (€)", className="input-label"),
                                    dcc.Input(id="calc-capital-fixed", type="number", disabled=True,
                                              className="calc-input", style={"opacity": "0.5"}),
                                ], className="input-group"),
                            ], className="calculator-grid", style={"gridTemplateColumns": "1fr 1fr", "marginBottom": "20px"}),

                            html.Div([
                                html.Div([
                                    html.Div("Position Actuelle", style={"fontWeight": "600", "fontSize": "12px", "marginBottom": "8px"}),
                                    html.Div(id="calc-current-pos-info", className="calc-info-box"),
                                ], style={"flex": "1"}),

                                html.Div([
                                    html.Div("Simulations d'Achats", style={"fontWeight": "600", "fontSize": "12px", "marginBottom": "8px"}),
                                    html.Div([
                                        html.Div([
                                            dcc.Input(id=f"calc-new-price-{i}", type="number", placeholder="Prix d'achat €", className="calc-input-small"),
                                            dcc.Input(id=f"calc-new-qty-{i}", type="number", placeholder="Quantité", className="calc-input-small"),
                                        ], className="calc-sim-row") for i in range(1, 4)
                                    ], id="calc-simulations-container"),
                                ], style={"flex": "2", "marginLeft": "20px"}),
                            ], style={"display": "flex", "marginBottom": "20px"}),

                            html.Div(id="calc-results", className="calc-results"),

                            # ── Simulateur Nouvel Actif ────────────────────────────────
                            html.Hr(style={"borderColor": "rgba(255,255,255,0.06)", "margin": "28px 0 20px"}),

                            html.Div("Simuler un Actif Non Détenu", style={
                                "fontWeight": "600", "fontSize": "12px", "marginBottom": "16px", "color": "#c4a24a",
                            }),

                            html.Div([
                                html.Div([
                                    html.Label("Ticker (ex: MC.PA, AAPL, BNP.PA)", className="input-label"),
                                    dcc.Input(id="new-ticker-input", type="text",
                                              placeholder="Ex: MC.PA", debounce=True,
                                              className="calc-input"),
                                ], className="input-group"),
                                html.Div([
                                    html.Label("Cours Actuel", className="input-label"),
                                    html.Div(id="new-ticker-price-display", className="calc-info-box",
                                             style={"minHeight": "36px", "padding": "8px 12px",
                                                    "fontSize": "13px", "display": "flex",
                                                    "alignItems": "center"}),
                                ], className="input-group"),
                            ], className="calculator-grid", style={"gridTemplateColumns": "1fr 1fr", "marginBottom": "14px"}),

                            html.Div([
                                html.Div([
                                    html.Label("Prix d'Achat Souhaité (€)", className="input-label"),
                                    dcc.Input(id="new-ticker-buy-price", type="number",
                                              placeholder="Prix €", className="calc-input"),
                                ], className="input-group"),
                                html.Div([
                                    html.Label("Quantité", className="input-label"),
                                    dcc.Input(id="new-ticker-qty", type="number",
                                              placeholder="Quantité", className="calc-input"),
                                ], className="input-group"),
                            ], className="calculator-grid", style={"gridTemplateColumns": "1fr 1fr", "marginBottom": "16px"}),

                            html.Div(id="new-ticker-results", className="calc-results"),

                        ], className="card-body"),
                    ], className="card"),
                ], style={"marginBottom": "14px"}),

                ]),  # end tab-overview

                # ── Tab 2: Analyse par Titre ──────────────────────────────
                dcc.Tab(label="Analyse par Titre", value="tab-tickers", className="main-tab", selected_className="main-tab--selected", children=[
                    html.Div([
                        html.Div([
                            html.Div([
                                html.Span("Rentabilité & Risque par Titre", className="card-title"),
                                html.Span("Données mensuelles historiques depuis 2019 — benchmark CAC 40 inclus",
                                          style={"fontSize": "10px", "color": "#94a3b8"}),
                            ], className="card-header"),
                            html.Div(id="ticker-analysis-table", className="card-body", style={"padding": "0"}),
                        ], className="card"),
                    ], style={"marginTop": "14px"}),
                ]),  # end tab-tickers

                # ── Tab 3: Performance Historique ─────────────────────────
                dcc.Tab(label="Performance Historique", value="tab-performance", className="main-tab", selected_className="main-tab--selected", children=[
                    html.Div([

                        # Large performance chart
                        html.Div([
                            html.Div([
                                html.Span("Performance vs Benchmark (CAC 40)", className="card-title"),
                                html.Span(id="perf-source-label",
                                          style={"fontSize": "10px", "color": "#94a3b8"}),
                            ], className="card-header"),
                            html.Div([
                                dcc.Graph(id="chart-performance-large",
                                          config={"displayModeBar": True},
                                          style={"height": "460px"}),
                            ], className="card-body"),
                        ], className="card"),

                        # Historical performance table
                        html.Div([
                            html.Div([
                                html.Span("Historique des Performances Mensuelles", className="card-title"),
                            ], className="card-header"),
                            html.Div(id="perf-history-table", className="card-body", style={"padding": "0"}),
                        ], className="card"),

                    ], style={"marginTop": "14px"}),
                ]),  # end tab-performance

                # ── Tab 4: Macroéconomie ──────────────────────────────────
                dcc.Tab(label="Macroéconomie", value="tab-macro", className="main-tab", selected_className="main-tab--selected", children=[
                    html.Div([

                        # ── Barre clé FRED + bouton refresh ──────────────────
                        html.Div([
                            html.Div([
                                html.Span("Données Macroéconomiques", className="card-title"),
                                html.Span("FRED (US/FR/DE) · World Bank (CN/IN) · Eurostat",
                                          style={"fontSize": "10px", "color": "#94a3b8"}),
                            ], className="card-header", style={"flex": "1"}),
                            html.Div([
                                dcc.Input(
                                    id="fred-api-key-input",
                                    type="password",
                                    placeholder="Clé API FRED (laisser vide si configurée sur le serveur)",
                                    value="",
                                    debounce=True,
                                    className="calc-input",
                                    style={"width": "300px", "marginRight": "10px"},
                                ),
                                html.Button("↻  Actualiser", id="macro-refresh-btn",
                                            className="export-btn", n_clicks=0),
                            ], style={"display": "flex", "alignItems": "center", "padding": "12px 16px"}),
                        ], className="card", style={"display": "flex", "alignItems": "center",
                                                    "justifyContent": "space-between"}),

                        dcc.Loading(id="macro-loading", type="circle", color="#c4a24a",
                                    children=html.Div(id="macro-loading-trigger", style={"height": "0px"})),

                        # ── Strip cycle (tous pays) ───────────────────────────
                        html.Div(id="macro-cycle-strip"),

                        # ── Sélecteur pays ────────────────────────────────────
                        html.Div([
                            html.Div([
                                html.Span("Analyse détaillée par pays", className="card-title"),
                                html.Span("Historique depuis 2000 — cliquez sur un indicateur pour voir l'historique complet",
                                          style={"fontSize": "10px", "color": "#94a3b8"}),
                            ], className="card-header"),
                            html.Div([
                                html.Button(
                                    f"{meta['flag']} {meta['label']}",
                                    id=f"mac-btn-{code}",
                                    n_clicks=0,
                                    style={
                                        "padding": "8px 16px", "borderRadius": "20px",
                                        "border": "1px solid rgba(196,162,74,0.3)",
                                        "backgroundColor": "rgba(196,162,74,0.08)",
                                        "color": "#c4a24a", "cursor": "pointer",
                                        "fontFamily": "Inter", "fontSize": "12px",
                                        "fontWeight": "600", "marginRight": "8px",
                                    }
                                )
                                for code, meta in macro_engine.COUNTRY_META.items()
                            ], style={"padding": "12px 16px", "display": "flex", "flexWrap": "wrap"}),
                        ], className="card"),

                        # ── Graphiques cycle + taux ───────────────────────────
                        html.Div([
                            html.Div([
                                html.Div([
                                    html.Span(id="macro-gdp-title", className="card-title"),
                                    html.Span("PIB QoQ (%) · phases détectées en fond",
                                              style={"fontSize": "10px", "color": "#94a3b8"}),
                                ], className="card-header"),
                                dcc.Graph(id="macro-gdp-chart",
                                          config={"displayModeBar": True},
                                          style={"height": "320px"}),
                            ], className="card"),
                            html.Div([
                                html.Div([
                                    html.Span(id="macro-rate-title", className="card-title"),
                                    html.Span("Taux 10 ans · taux directeur (court)",
                                              style={"fontSize": "10px", "color": "#94a3b8"}),
                                ], className="card-header"),
                                dcc.Graph(id="macro-rate-chart",
                                          config={"displayModeBar": True},
                                          style={"height": "320px"}),
                            ], className="card"),
                        ], className="grid-50-50"),

                        # ── Tuiles indicateurs cliquables ─────────────────────
                        html.Div(id="macro-indicator-tiles"),

                        # ── Spread de taux (signal récession) ─────────────────
                        html.Div(id="macro-spread-section"),

                        # ── Rotation sectorielle ──────────────────────────────
                        html.Div(id="macro-rotation-section"),

                        # ── Portefeuille & cycle (actuel + à venir) ───────────
                        html.Div(id="macro-portfolio-cycle"),

                        # ── FX & matières premières ───────────────────────────
                        html.Div(id="macro-markets-section"),

                    ], style={"marginTop": "14px"}),
                ]),  # end tab-macro

                ]),  # end dcc.Tabs

            ]),  # end dashboard-body

        ], className="main-content"),
    ])


app.layout = build_layout()


# ─── Callbacks ────────────────────────────────────────────────────────────────

# 1. Parse uploaded file → store data
@app.callback(
    Output("store-data", "data"),
    Output("last-update-text", "children"),
    Output("loading-trigger", "children"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True,
)
def process_upload(contents, filename):
    if contents is None:
        return no_update, no_update, no_update

    _, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    df = pd.read_excel(io.BytesIO(decoded), sheet_name="BDD")
    df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m-%d")

    try:
        val_df = pd.read_excel(io.BytesIO(decoded), sheet_name="Valorisation")
        val_df["Date"] = pd.to_datetime(val_df["Date"]).dt.strftime("%Y-%m-%d")
        val_json = val_df.to_json(orient="records")
    except Exception:
        val_json = "[]"

    payload = {
        "df": df.to_json(orient="records"),
        "val_df": val_json,
        "filename": filename,
    }
    ts = f"Mis à jour : {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    return json.dumps(payload), ts, ""


# 2. Compute analytics from stored data
@app.callback(
    Output("store-analytics", "data"),
    Input("store-data", "data"),
    prevent_initial_call=True,
)
def compute_analytics_callback(raw_data):
    if not raw_data:
        return None

    try:
        data = json.loads(raw_data)
        df = pd.read_json(io.StringIO(data["df"]), orient="records")
        df["Date"] = pd.to_datetime(df["Date"])

        # ── Valorisation sheet (optional) ─────────────────────────────────────
        val_json = data.get("val_df", "[]")
        val_df = pd.read_json(io.StringIO(val_json), orient="records") if val_json and val_json != "[]" else pd.DataFrame()
        use_valorisation = (
            not val_df.empty
            and "Valeur_Portefeuille" in val_df.columns
            and len(val_df) > 1
        )

        positions_df, summary = engine.build_positions_table(df)
        if positions_df.empty:
            return None

        tickers = positions_df["Ticker"].tolist()
        weights_dict = dict(zip(
            positions_df["Ticker"].tolist(),
            (positions_df["Poids (%)"] / 100).tolist()
        ))

        analytics = engine.compute_analytics(tickers, weights_dict, start="2019-01-01", df=df)
        optimization = engine.compute_optimal_portfolios(tickers, start="2019-01-01")
        bench_ret_series = analytics.get("bench_ret")

        # ── Portfolio performance series ───────────────────────────────────────
        if use_valorisation:
            val_df["Date"] = pd.to_datetime(val_df["Date"])
            val_sorted = val_df.sort_values("Date").reset_index(drop=True)
            val_series = val_sorted.set_index("Date")["Valeur_Portefeuille"]
            # Normalize val_series index to 1st of month for robust alignment
            val_series.index = val_series.index.to_period('M').to_timestamp()
            pf_ret = val_series.pct_change().dropna()
            # Normalize pf_ret index to 1st of month
            pf_ret.index = pf_ret.index.to_period('M').to_timestamp()
            perf_pf = (100 * (1 + pf_ret).cumprod()).round(2)
            perf_source = "valorisation"
        else:
            # Fallback: TWR computed from BDD transactions
            pf_ret = analytics.get("pf_ret")
            if pf_ret is not None:
                # Normalize pf_ret index to 1st of month
                pf_ret.index = pf_ret.index.to_period('M').to_timestamp()
            perf_pf = analytics.get("perf_pf")
            if perf_pf is not None:
                # Normalize perf_pf index to 1st of month
                perf_pf.index = perf_pf.index.to_period('M').to_timestamp()
            val_series = None
            perf_source = "twr"

        # ── Align benchmark to portfolio start ────────────────────────────────
        bench_raw = analytics.get("perf_bench")
        if bench_raw is not None and perf_pf is not None and not perf_pf.empty:
            # Normalize bench_raw to 1st of month
            bench_raw.index = bench_raw.index.to_period('M').to_timestamp()
            perf_bench = bench_raw[bench_raw.index >= perf_pf.index.min()]
        else:
            perf_bench = bench_raw

        # ── Build historical performance table ────────────────────────────────
        perf_history_json = None
        if pf_ret is not None and not (hasattr(pf_ret, "empty") and pf_ret.empty):
            # Create robust lookups for benchmark data (normalized to 1st of month)
            bench_lookup_ret = None
            bench_lookup_cum = None
            if bench_ret_series is not None:
                # Normalize benchmark returns to 1st of month
                bench_lookup_ret = bench_ret_series.copy()
                bench_lookup_ret.index = bench_lookup_ret.index.to_period('M').to_timestamp()

                # Create cumulative benchmark lookup (base 1, normalized, ffilled)
                bench_lookup_cum = (1 + bench_lookup_ret).cumprod()
                bench_lookup_cum.index = bench_lookup_cum.index.to_period('M').to_timestamp()
                bench_lookup_cum = bench_lookup_cum.reindex(pf_ret.index, method='ffill')

            rows = []
            cum_pf = 1.0
            cum_bench = 1.0
            for date, r_pf in pf_ret.items():
                cum_pf *= (1 + r_pf)

                # Normalize current date for lookup
                norm_date = date.to_period('M').to_timestamp()

                r_bench = None
                if bench_lookup_ret is not None and norm_date in bench_lookup_ret.index:
                    val = bench_lookup_ret.loc[norm_date]
                    if not pd.isna(val):
                        r_bench = float(val)

                if r_bench is not None:
                    cum_bench *= (1 + r_bench)

                # Use cumulative lookup for the benchmark to ensure it stays continuous
                # even if a monthly return is missing in the lookup.
                # If we have a valid r_bench, we use the calculated cum_bench.
                # Otherwise, we try to fall back to the benchmark's own cumulative series.
                current_cum_bench = cum_bench
                if r_bench is None and bench_lookup_cum is not None and norm_date in bench_lookup_cum.index:
                    current_cum_bench = bench_lookup_cum.loc[norm_date]

                row = {
                    "Date": date.strftime("%b %Y") if hasattr(date, "strftime") else str(date)[:7],
                    "Rend. Mensuel PF (%)": round(r_pf * 100, 2),
                    "Perf. Cumulée PF (%)": round((cum_pf - 1) * 100, 2),
                    "Rend. Mensuel CAC 40 (%)": round(r_bench * 100, 2) if r_bench is not None else None,
                    "Perf. Cumulée CAC 40 (%)": round((current_cum_bench - 1) * 100, 2) if r_bench is not None else None,
                    "Écart Mensuel (%)": round((r_pf - r_bench) * 100, 2) if r_bench is not None else None,
                }
                if use_valorisation and val_series is not None and norm_date in val_series.index:
                    row["Valeur PF (€)"] = round(float(val_series.loc[norm_date]), 2)
                rows.append(row)
            # Most recent first
            rows.reverse()
            perf_history_json = json.dumps(rows)

        # Serialize DataFrames
        # Use json.loads(to_json()) so NaN→null and unicode column names survive the
        # outer json.dumps/loads round-trip without escaping issues.
        pos_records = json.loads(positions_df.to_json(orient="records"))

        def _ser(s):
            if s is None:
                return None
            df_s = s.reset_index()
            df_s.columns = ["date", "val"]  # normalize regardless of index/series name
            return df_s.to_json(date_format="iso")

        payload = {
            "summary": summary,
            "positions": pos_records,
            "metrics": analytics.get("metrics", {}),
            "bench_metrics": analytics.get("bench_metrics", {}),
            "per_ticker": analytics.get("per_ticker", {}),
            "optimization": optimization,
            "perf_source": perf_source,
            "perf_pf": _ser(perf_pf),
            "perf_bench": _ser(perf_bench),
            "pf_ret": pf_ret.reset_index().rename(columns={"Date": "date", 0: "ret"}).to_json() if pf_ret is not None else None,
            "perf_history": perf_history_json,
        }

        return json.dumps(payload)

    except Exception as e:
        import traceback
        print(f"[app] compute_analytics_callback error: {e}")
        traceback.print_exc()
        return None


# 3. Show/hide dashboard body
@app.callback(
    Output("dashboard-body", "style"),
    Output("empty-state", "style"),
    Input("store-analytics", "data"),
)
def toggle_dashboard(data):
    if data:
        return {"display": "block"}, {"display": "none"}
    return {"display": "none"}, {"display": "flex", "flexDirection": "column", "alignItems": "center",
                                   "justifyContent": "center", "padding": "80px 20px", "color": "#94a3b8",
                                   "textAlign": "center"}


# 4. KPI Row
@app.callback(
    Output("kpi-row", "children"),
    Input("store-analytics", "data"),
)
def update_kpis(data):
    if not data:
        return []

    d = json.loads(data)
    s = d.get("summary", {})

    total_portfolio = s.get("total_portfolio_value", 0)
    valeur_titres   = s.get("total_value", 0)
    cash            = s.get("cash_balance", 0)
    dividendes      = s.get("total_dividends", 0)
    unrealized      = s.get("unrealized_pnl", 0)
    unrealized_pct  = s.get("unrealized_pnl_pct", 0)
    capital_verse   = s.get("total_invested", 0)
    total_pnl       = s.get("total_pnl", 0)
    pnl_pct         = s.get("pnl_pct", 0)

    cards = [
        ("Valorisation Totale", fmt_eur(total_portfolio),
         html.Span("Titres + Cash", className="kpi-badge badge-neutral")),
        ("Valeur des Titres", fmt_eur(valeur_titres),
         html.Span("Au cours du marché", className="kpi-badge badge-neutral")),
        ("Cash Disponible", fmt_eur(cash),
         html.Span("Solde espèces", className="kpi-badge badge-neutral")),
        ("Dividendes Perçus", fmt_eur(dividendes),
         html.Span("Depuis le début", className="kpi-badge badge-neutral")),
        ("+/- Values Latentes", fmt_eur(unrealized),
         badge(unrealized_pct, fmt_pct)),
        ("Capital Versé", fmt_eur(capital_verse),
         html.Span("Dépôts cumulés", className="kpi-badge badge-neutral")),
        ("Performance Totale", fmt_eur(total_pnl),
         badge(pnl_pct, fmt_pct)),
    ]

    return [
        html.Div([
            html.Div(label, className="kpi-label"),
            html.Div(value, className="kpi-value"),
            sub,
        ], className="kpi-card")
        for label, value, sub in cards
    ]


# 5. Positions Table
@app.callback(
    Output("positions-table-container", "children"),
    Output("positions-count", "children"),
    Input("store-analytics", "data"),
)
def update_positions_table(data):
    if not data:
        return [], ""

    d = json.loads(data)
    positions_df = pd.DataFrame(d["positions"])

    def row_style(row):
        pnl = row.get("P&L (%)", 0)
        if pnl > 5:
            return {"backgroundColor": "#f0fdf4"}
        elif pnl < -5:
            return {"backgroundColor": "#fff5f5"}
        return {}

    table = dash_table.DataTable(
        data=positions_df.to_dict("records"),
        columns=[{"name": c, "id": c} for c in positions_df.columns],
        style_cell=TABLE_STYLE_CELL,
        style_header=TABLE_STYLE_HEADER,
        style_table={"overflowX": "auto", "minWidth": "100%"},
        style_data_conditional=[
            # Catégorie badges
            {
                "if": {"filter_query": '{Catégorie} = "CAC 40"', "column_id": "Catégorie"},
                "backgroundColor": "rgba(196,162,74,0.15)", "color": "#c4a24a",
                "fontWeight": "600", "borderRadius": "4px",
            },
            {
                "if": {"filter_query": '{Catégorie} = "Index"', "column_id": "Catégorie"},
                "backgroundColor": "rgba(0,200,150,0.12)", "color": "#00c896",
                "fontWeight": "600", "borderRadius": "4px",
            },
            {
                "if": {"filter_query": '{Catégorie} = "Action"', "column_id": "Catégorie"},
                "backgroundColor": "rgba(96,165,250,0.12)", "color": "#3d8fd1",
                "fontWeight": "600", "borderRadius": "4px",
            },
            # P&L
            {
                "if": {"filter_query": "{P&L (%)} > 0", "column_id": "P&L (%)"},
                "color": "#00c896", "fontWeight": "600",
            },
            {
                "if": {"filter_query": "{P&L (%)} < 0", "column_id": "P&L (%)"},
                "color": "#f04f6a", "fontWeight": "600",
            },
            {
                "if": {"filter_query": "{P&L (€)} > 0", "column_id": "P&L (€)"},
                "color": "#00c896",
            },
            {
                "if": {"filter_query": "{P&L (€)} < 0", "column_id": "P&L (€)"},
                "color": "#f04f6a",
            },
            {
                "if": {"filter_query": "{Retour Total (%)} > 0", "column_id": "Retour Total (%)"},
                "color": "#00c896", "fontWeight": "600",
            },
            {
                "if": {"filter_query": "{Retour Total (%)} < 0", "column_id": "Retour Total (%)"},
                "color": "#f04f6a", "fontWeight": "600",
            },
            {"if": {"row_index": "odd"}, "backgroundColor": "rgba(255,255,255,0.02)"},
        ],
        page_size=15,
        sort_action="native",
        filter_action="native",
        style_filter={"backgroundColor": "#1E293B", "color": "#E2E8F0", "fontSize": "11px"},
    )

    count = f"{len(positions_df)} positions"
    return table, count


# 6. Pie Chart
@app.callback(
    Output("pie-allocation", "figure"),
    Input("store-analytics", "data"),
)
def update_pie(data):
    if not data:
        return go.Figure()

    d = json.loads(data)
    positions_df = pd.DataFrame(d["positions"])
    cash = d.get("summary", {}).get("cash_balance", 0)

    labels = list(positions_df["Ticker"])
    values = list(positions_df["Valeur (€)"])
    colors = list(PALETTE[:len(positions_df)])

    if cash and cash > 0:
        labels.append("Cash")
        values.append(cash)
        colors.append("#64748b")

    fig = go.Figure(go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        textinfo="label+percent",
        textfont=dict(size=10, family="Inter", color="#ffffff"),
        marker=dict(
            colors=colors,
            line=dict(color="#ffffff", width=2),
        ),
        hovertemplate="<b>%{label}</b><br>%{value:.2f} €<br>%{percent}<extra></extra>",
    ))

    total = sum(values)
    fig.update_layout(
        **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
        showlegend=True,
        legend=dict(
            orientation="v", x=1.0, y=0.5,
            font=dict(size=10), xanchor="left",
        ),
        margin=dict(l=10, r=120, t=20, b=10),
        annotations=[dict(
            text=f"<b>{total:,.0f} €</b>".replace(",", " "),
            x=0.5, y=0.5, font_size=13, showarrow=False,
            font=dict(family="Inter", color="#E8EFF8"),
        )],
    )
    return fig


# 6b. Category Pie Chart
@app.callback(
    Output("pie-category", "figure"),
    Input("store-analytics", "data"),
)
def update_category_pie(data):
    if not data:
        return go.Figure()

    d = json.loads(data)
    positions_df = pd.DataFrame(d["positions"])

    if "Catégorie" not in positions_df.columns:
        return go.Figure()

    cat_data = (
        positions_df.groupby("Catégorie")["Valeur (€)"]
        .sum()
        .reset_index()
        .sort_values("Valeur (€)", ascending=False)
    )

    cash = d.get("summary", {}).get("cash_balance", 0)
    if cash and cash > 0:
        cash_row = pd.DataFrame([{"Catégorie": "Cash", "Valeur (€)": cash}])
        cat_data = pd.concat([cat_data, cash_row], ignore_index=True)

    CAT_COLORS = {"CAC 40": "#c4a24a", "Index": "#00c896", "Action": "#3d8fd1", "Cash": "#64748b"}
    colors = [CAT_COLORS.get(c, "#5b7490") for c in cat_data["Catégorie"]]

    fig = go.Figure(go.Pie(
        labels=cat_data["Catégorie"],
        values=cat_data["Valeur (€)"],
        hole=0.52,
        textinfo="label+percent",
        textfont=dict(size=11, family="Inter", color="#ffffff"),
        marker=dict(colors=colors, line=dict(color="#07101f", width=3)),
        hovertemplate="<b>%{label}</b><br>%{value:,.2f} €<br>%{percent}<extra></extra>",
    ))

    fig.update_layout(
        **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
        showlegend=True,
        legend=dict(
            orientation="h", x=0.5, y=-0.08,
            xanchor="center", font=dict(size=11, color="#5b7490"),
        ),
        margin=dict(l=10, r=10, t=10, b=30),
    )
    return fig


# 7. Performance Chart
@app.callback(
    Output("chart-performance", "figure"),
    Input("store-analytics", "data"),
)
def update_performance_chart(data):
    if not data:
        return go.Figure()

    d = json.loads(data)
    perf_pf_raw = d.get("perf_pf")
    perf_bench_raw = d.get("perf_bench")

    if not perf_pf_raw:
        return go.Figure()

    perf_pf = pd.read_json(io.StringIO(perf_pf_raw))
    fig = go.Figure()

    # Portfolio line
    fig.add_trace(go.Scatter(
        x=perf_pf.iloc[:, 0],
        y=perf_pf.iloc[:, 1],
        name="Portefeuille Actuel",
        line=dict(color="#c4a24a", width=2),
        hovertemplate="%{y:.1f}<extra>Portefeuille</extra>",
    ))

    # Benchmark line
    if perf_bench_raw:
        perf_bench = pd.read_json(io.StringIO(perf_bench_raw))
        fig.add_trace(go.Scatter(
            x=perf_bench.iloc[:, 0],
            y=perf_bench.iloc[:, 1],
            name="CAC 40 (^FCHI)",
            line=dict(color="#f59e0b", width=2, dash="dash"),
            hovertemplate="%{y:.1f}<extra>CAC 40</extra>",
        ))

    fig.add_hline(y=100, line_color="rgba(255,255,255,0.15)", line_width=1, line_dash="dot")

    fig.update_layout(**CHART_LAYOUT)
    fig.update_layout(
        hovermode="x unified",
        yaxis_tickformat=".0f",
    )
    return fig


# 8. Risk Metrics Panel
@app.callback(
    Output("risk-metrics-panel", "children"),
    Input("store-analytics", "data"),
)
def update_risk_panel(data):
    if not data:
        return []

    d = json.loads(data)
    m = d.get("metrics", {})
    bm = d.get("bench_metrics", {})
    summ = d.get("summary", {})
    mwr = summ.get("money_weighted_return")

    def metric(label, value, sub="", positive_good=True):
        v_str = str(value) if value is not None else "—"
        color = "#8ba8c4"
        try:
            v_num = float(value)
            if positive_good:
                color = "#00c896" if v_num > 0 else ("#f04f6a" if v_num < 0 else "#8ba8c4")
            else:
                color = "#f04f6a" if v_num > 0 else ("#00c896" if v_num < 0 else "#8ba8c4")
        except (TypeError, ValueError):
            pass

        return html.Div([
            html.Div(label, className="metric-label"),
            html.Div(v_str, className="metric-value", style={"color": color}),
            html.Div(sub, className="metric-sub") if sub else None,
        ], className="metric-item")

    items = [
        metric("Rendement Annualisé (TWR)", f"{m.get('ann_ret', '—')} %", "Time-weighted, géométrique"),
        metric("TRI (pondéré flux)", f"{mwr} %" if mwr is not None else "—", "Money-weighted, inclut cash & timing"),
        metric("Volatilité Annuelle", f"{m.get('ann_vol', '—')} %", "Portefeuille actuel", positive_good=False),
        metric("Sharpe Ratio", m.get("sharpe", "—"), f"Bench: {bm.get('sharpe', '—')}"),
        metric("Sortino Ratio", m.get("sortino", "—"), "Uniquement baisse"),
        metric("Max Drawdown", f"{m.get('max_dd', '—')} %", "Pire perte consécutive", positive_good=False),
        metric("Beta", m.get("beta", "—"), "vs CAC 40"),
        metric("Alpha (Jensen)", f"{m.get('alpha', '—')} %", "Surperformance ajustée"),
        metric("VaR Historique (5%)", f"{m.get('var_hist', '—')} %", "Perte mensuelle maximale", positive_good=False),
        metric("VaR Gaussien (5%)", f"{m.get('var_gauss', '—')} %", "", positive_good=False),
        metric("CVaR (5%)", f"{m.get('cvar', '—')} %", "Expected Shortfall", positive_good=False),
    ]
    items = [i for i in items if i]

    bench_section = []
    if bm:
        bench_section = [
            html.Div(
                "Benchmark — CAC 40",
                style={"fontSize": "10px", "fontWeight": "600", "color": "#94a3b8",
                       "textTransform": "uppercase", "letterSpacing": "0.6px",
                       "marginTop": "12px", "marginBottom": "6px"},
            ),
            html.Div([
                metric("Rendement Annualisé", f"{bm.get('ann_ret', '—')} %"),
                metric("Volatilité Annuelle", f"{bm.get('ann_vol', '—')} %", positive_good=False),
            ], className="metrics-grid"),
        ]

    return [
        html.Div(items, className="metrics-grid"),
        *bench_section,
    ]


# 8b. Attribution de performance par titre
@app.callback(
    Output("attribution-container", "children"),
    Input("store-analytics", "data"),
)
def update_attribution(data):
    if not data:
        return []
    try:
        d = json.loads(data)
        positions = d.get("positions", [])
        if not positions:
            return []

        rows = []
        for p in positions:
            t = p.get("Ticker")
            pnl = float(p.get("P&L (€)", 0) or 0)
            div = float(p.get("Dividendes (€)", 0) or 0)
            rows.append({"ticker": t, "contrib": pnl + div, "pnl": pnl, "div": div})
        rows.sort(key=lambda r: r["contrib"])

        total_contrib = sum(r["contrib"] for r in rows)
        tickers = [r["ticker"] for r in rows]
        contribs = [round(r["contrib"], 2) for r in rows]
        colors = ["#00c896" if c >= 0 else "#f04f6a" for c in contribs]
        texts = [fmt_eur(c) for c in contribs]

        fig = go.Figure(go.Bar(
            x=contribs, y=tickers, orientation="h",
            marker=dict(color=colors),
            text=texts, textposition="auto",
            customdata=[[fmt_eur(r["pnl"]), fmt_eur(r["div"])] for r in rows],
            hovertemplate="<b>%{y}</b><br>Contribution : %{x:.2f} €"
                          "<br>Plus-value : %{customdata[0]}"
                          "<br>Dividendes : %{customdata[1]}<extra></extra>",
        ))
        fig.add_vline(x=0, line_color="rgba(255,255,255,0.3)", line_width=1)
        fig.update_layout(
            **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
            height=max(260, 26 * len(rows) + 60),
            margin=dict(l=8, r=16, t=8, b=24),
            xaxis_ticksuffix=" €", showlegend=False,
        )

        best = rows[-1] if rows else None
        worst = rows[0] if rows else None

        def _call(lbl, row, color):
            if not row:
                return None
            return html.Div([
                html.Div(lbl, style={"fontSize": "10px", "color": "#64748b",
                                     "textTransform": "uppercase", "letterSpacing": "0.8px"}),
                html.Div(row["ticker"], style={"fontSize": "16px", "fontWeight": "700",
                                               "color": "#e8eef6"}),
                html.Div(fmt_eur(row["contrib"]), style={"fontSize": "18px",
                         "fontWeight": "700", "color": color}),
            ], style={"flex": "1", "minWidth": "120px", "padding": "10px 14px",
                      "borderRadius": "8px", "backgroundColor": hex_rgba(color, 0.10),
                      "border": f"1px solid {hex_rgba(color, 0.3)}"})

        return html.Div([
            html.Div([
                _call("Meilleur contributeur", best, "#00c896"),
                _call("Pire contributeur", worst, "#f04f6a"),
                html.Div([
                    html.Div("Total (positions ouvertes)",
                             style={"fontSize": "10px", "color": "#64748b",
                                    "textTransform": "uppercase", "letterSpacing": "0.8px"}),
                    html.Div(fmt_eur(total_contrib), style={"fontSize": "18px",
                             "fontWeight": "700",
                             "color": "#00c896" if total_contrib >= 0 else "#f04f6a",
                             "marginTop": "16px"}),
                ], style={"flex": "1", "minWidth": "120px", "padding": "10px 14px",
                          "borderRadius": "8px", "border": "1px solid rgba(255,255,255,0.06)"}),
            ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap",
                      "padding": "0 0 12px"}),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ])
    except Exception:
        import traceback; traceback.print_exc()
        return html.Div("Erreur de rendu de l'attribution.",
                        style={"color": "#f04f6a", "padding": "16px"})


# 9. Drawdown Chart
@app.callback(
    Output("chart-drawdown", "figure"),
    Input("store-analytics", "data"),
)
def update_drawdown_chart(data):
    if not data:
        return go.Figure()

    d = json.loads(data)
    pf_ret_raw = d.get("pf_ret")
    if not pf_ret_raw:
        return go.Figure()

    pf_ret = pd.read_json(io.StringIO(pf_ret_raw))
    dates = pf_ret.iloc[:, 0]
    rets = pf_ret.iloc[:, 1]

    wealth = (1 + rets).cumprod()
    peak = wealth.cummax()
    drawdown = ((wealth - peak) / peak * 100).round(2)

    fig = go.Figure(go.Scatter(
        x=dates, y=drawdown,
        fill="tozeroy",
        fillcolor="rgba(220, 38, 38, 0.12)",
        line=dict(color="#dc2626", width=1.5),
        hovertemplate="%{y:.2f}%<extra></extra>",
        name="Drawdown",
    ))

    fig.add_hline(y=0, line_color="rgba(255,255,255,0.15)", line_width=1)

    layout = {**CHART_LAYOUT, "yaxis": {**CHART_LAYOUT["yaxis"], "ticksuffix": "%"}}
    fig.update_layout(**layout)
    fig.update_layout(showlegend=False, margin=dict(l=16, r=16, t=16, b=16))
    return fig


# 10. Optimization Table
@app.callback(
    Output("optimization-table-container", "children"),
    Input("store-analytics", "data"),
)
def update_optimization_table(data):
    if not data:
        return []

    d = json.loads(data)
    opti = d.get("optimization", {})
    if not opti:
        return html.Div("Données insuffisantes pour l'optimisation.",
                        style={"padding": "16px", "color": "#94a3b8", "fontSize": "12px"})

    m = d.get("metrics", {})
    s = d.get("summary", {})

    # Build comparison rows
    rows = [
        {
            "Stratégie": "Portefeuille Actuel",
            "Rendement Ann. (%)": m.get("ann_ret", "—"),
            "Volatilité (%)": m.get("ann_vol", "—"),
            "Sharpe": m.get("sharpe", "—"),
        },
        {
            "Stratégie": "Max Sharpe Ratio",
            "Rendement Ann. (%)": opti.get("msr", {}).get("return", "—"),
            "Volatilité (%)": opti.get("msr", {}).get("vol", "—"),
            "Sharpe": opti.get("msr", {}).get("sharpe", "—"),
        },
        {
            "Stratégie": "Global Min. Variance",
            "Rendement Ann. (%)": opti.get("gmv", {}).get("return", "—"),
            "Volatilité (%)": opti.get("gmv", {}).get("vol", "—"),
            "Sharpe": opti.get("gmv", {}).get("sharpe", "—"),
        },
        {
            "Stratégie": "Equal Weight",
            "Rendement Ann. (%)": opti.get("equal", {}).get("return", "—"),
            "Volatilité (%)": opti.get("equal", {}).get("vol", "—"),
            "Sharpe": opti.get("equal", {}).get("sharpe", "—"),
        },
    ]

    table = dash_table.DataTable(
        data=rows,
        columns=[{"name": c, "id": c} for c in rows[0].keys()],
        style_cell=TABLE_STYLE_CELL,
        style_header=TABLE_STYLE_HEADER,
        style_table={"overflowX": "auto"},
        style_data_conditional=[
            {"if": {"row_index": 0}, "fontWeight": "600", "backgroundColor": "rgba(197,160,89,0.08)", "color": "#E8EFF8"},
            {"if": {"row_index": "odd"}, "backgroundColor": "rgba(255,255,255,0.02)"},
        ],
    )

    # Weight breakdown charts for MSR & GMV
    tickers = opti.get("tickers", [])
    msr_weights = opti.get("msr", {}).get("weights", {})
    gmv_weights = opti.get("gmv", {}).get("weights", {})

    def weight_fig(weights_dict, title, color):
        active_w = {t: w for t, w in weights_dict.items() if w > 0.005}
        if not active_w:
            return go.Figure()
        fig = go.Figure(go.Bar(
            x=list(active_w.keys()),
            y=[v * 100 for v in active_w.values()],
            marker_color=color,
            text=[f"{v*100:.1f}%" for v in active_w.values()],
            textposition="outside",
            textfont=dict(size=9),
            hovertemplate="%{x}: %{y:.1f}%<extra></extra>",
        ))
        fig.update_layout(
            **{k: v for k, v in CHART_LAYOUT.items() if k != "margin"},
            showlegend=False,
            yaxis_ticksuffix="%",
            title=dict(text=title, font=dict(size=11, color="#8ba8c4")),
            margin=dict(l=16, r=16, t=36, b=30),
        )
        return fig

    charts = html.Div([
        dcc.Graph(
            figure=weight_fig(msr_weights, "Poids MSR (Max Sharpe)", "#c4a24a"),
            config={"displayModeBar": False},
            style={"height": "200px"},
        ),
        dcc.Graph(
            figure=weight_fig(gmv_weights, "Poids GMV (Min Variance)", "#00c896"),
            config={"displayModeBar": False},
            style={"height": "200px"},
        ),
    ], style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "8px",
              "padding": "12px 18px", "paddingTop": "0"})

    return [table, charts]


# 11. Position Calculator Initialization
@app.callback(
    [Output("calc-ticker-select", "options"),
     Output("calc-capital-fixed", "value")],
    Input("store-analytics", "data")
)
def init_calculator(data_json):
    if not data_json:
        return [], None
    
    try:
        data = json.loads(data_json)
        df_pos = pd.DataFrame(data["positions"])
        
        if df_pos.empty:
            return [], None
        
        tickers = df_pos["Ticker"].unique().tolist()
        total_capital = data["summary"].get("total_portfolio_value", 0)

        options = [{"label": t, "value": t} for t in tickers]
        return options, total_capital
    except Exception as e:
        print(f"DEBUG Error initializing calculator: {e}")
        return [], None


# 11b. Position Calculator logic
@app.callback(
    [Output("calc-current-pos-info", "children"),
     Output("calc-results", "children")],
    [Input("calc-ticker-select", "value"),
     Input("calc-capital-fixed", "value"),
     Input("calc-new-price-1", "value"), Input("calc-new-qty-1", "value"),
     Input("calc-new-price-2", "value"), Input("calc-new-qty-2", "value"),
     Input("calc-new-price-3", "value"), Input("calc-new-qty-3", "value")],
    State("store-analytics", "data")
)
def update_calculator(ticker, capital, p1, q1, p2, q2, p3, q3, data_json):
    if not ticker or not data_json:
        return "Sélectionnez un actif...", []

    try:
        data = json.loads(data_json)
        df_pos = pd.DataFrame(data["positions"])

        # Filter for the selected ticker
        pos = df_pos[df_pos["Ticker"] == ticker]
        if pos.empty:
            return "Actif non trouvé", []
            
        pos = pos.iloc[0]
        # Match the exact column names from engine.py: "PRU (€)", "Qté"
        p_pru = float(pos.get("PRU (€)", 0))
        p_qty = float(pos.get("Qté", 0))
        p_cours = float(pos.get("Cours (€)", p_pru))

        current_info = html.Div([
            html.Div(f"PRU : {p_pru:,.2f} €"),
            html.Div(f"Cours actuel : {p_cours:,.2f} €"),
            html.Div(f"Quantité : {p_qty:,.2f}"),
            html.Div(f"Valeur de Marché : {p_cours * p_qty:,.2f} €", style={"fontWeight": "600", "marginTop": "4px"}),
        ])

        # Simulation logic
        sims = [(p1, q1), (p2, q2), (p3, q3)]
        new_total_cost = p_pru * p_qty
        new_total_qty = p_qty

        # Default: value at current market price; updated to last sim price if simulations entered
        simulated_price = p_cours

        for price, qty in sims:
            if price is not None and qty is not None:
                new_total_cost += float(price) * float(qty)
                new_total_qty += float(qty)
                simulated_price = float(price)

        new_pru = new_total_cost / new_total_qty if new_total_qty > 0 else 0
        new_market_value = new_total_qty * simulated_price
        
        # weighting against the total initial cost of the portfolio
        weighting = (new_market_value / capital * 100) if capital and capital > 0 else 0

        def result_card(label, value, color="#E8EFF8"):
            return html.Div([
                html.Div(label, className="calc-result-label"),
                html.Div(value, className="calc-result-value", style={"color": color}),
            ], className="calc-result-item")

        results = [
            result_card("Nouveau PRU", f"{new_pru:,.2f} €".replace(",", " ")),
            result_card("Nouvelle Qté", f"{new_total_qty:,.1f}"),
            result_card("Valeur Simulée", f"{new_market_value:,.2f} €".replace(",", " "), "#c4a24a"),
            result_card("Nouvelle Pondé.", f"{weighting:.2f} %", "#d4b87a"),
        ]

        return current_info, results
    except Exception as e:
        print(f"DEBUG Error updating calculator: {e}")
        return "Erreur lors du calcul", []


# 12. Performance Historique — grand graphique + table
@app.callback(
    Output("chart-performance-large", "figure"),
    Output("perf-source-label", "children"),
    Output("perf-history-table", "children"),
    Input("store-analytics", "data"),
)
def update_performance_page(data):
    empty_fig = go.Figure()
    if not data:
        return empty_fig, "", []

    d = json.loads(data)
    perf_pf_raw  = d.get("perf_pf")
    perf_bench_raw = d.get("perf_bench")
    perf_source  = d.get("perf_source", "twr")
    perf_history_raw = d.get("perf_history")

    # ── Large chart ───────────────────────────────────────────────────────
    fig = go.Figure()
    if perf_pf_raw:
        perf_pf = pd.read_json(io.StringIO(perf_pf_raw))
        fig.add_trace(go.Scatter(
            x=perf_pf.iloc[:, 0],
            y=perf_pf.iloc[:, 1],
            name="Portefeuille",
            line=dict(color="#c4a24a", width=2.5),
            hovertemplate="%{y:.1f}<extra>Portefeuille</extra>",
        ))

    if perf_bench_raw:
        perf_bench = pd.read_json(io.StringIO(perf_bench_raw))
        fig.add_trace(go.Scatter(
            x=perf_bench.iloc[:, 0],
            y=perf_bench.iloc[:, 1],
            name="CAC 40 (^FCHI)",
            line=dict(color="#f59e0b", width=2, dash="dash"),
            hovertemplate="%{y:.1f}<extra>CAC 40</extra>",
        ))

    fig.add_hline(y=100, line_color="rgba(255,255,255,0.15)", line_width=1, line_dash="dot")
    fig.update_layout(
        **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
        hovermode="x unified",
        yaxis_tickformat=".0f",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=12)),
        margin=dict(l=16, r=16, t=40, b=16),
    )

    source_label = (
        "Source : onglet Valorisation (valeurs réelles)"
        if perf_source == "valorisation"
        else "Source : calculé depuis les transactions BDD (TWR)"
    )

    # ── History table ─────────────────────────────────────────────────────
    if not perf_history_raw:
        return fig, source_label, html.Div("Données insuffisantes.", style={"padding": "16px", "color": "#94a3b8", "fontSize": "12px"})

    rows = json.loads(perf_history_raw)
    if not rows:
        return fig, source_label, []

    cols_order = ["Date"]
    if "Valeur PF (€)" in rows[0]:
        cols_order.append("Valeur PF (€)")
    cols_order += [
        "Rend. Mensuel PF (%)", "Perf. Cumulée PF (%)",
        "Rend. Mensuel CAC 40 (%)", "Perf. Cumulée CAC 40 (%)",
        "Écart Mensuel (%)",
    ]
    # Filter to columns actually present
    cols_order = [c for c in cols_order if c in rows[0]]

    style_cond = [
        {"if": {"filter_query": "{Rend. Mensuel PF (%)} > 0", "column_id": "Rend. Mensuel PF (%)"},
         "color": "#00c896", "fontWeight": "600"},
        {"if": {"filter_query": "{Rend. Mensuel PF (%)} < 0", "column_id": "Rend. Mensuel PF (%)"},
         "color": "#f04f6a", "fontWeight": "600"},
        {"if": {"filter_query": "{Perf. Cumulée PF (%)} > 0", "column_id": "Perf. Cumulée PF (%)"},
         "color": "#00c896", "fontWeight": "600"},
        {"if": {"filter_query": "{Perf. Cumulée PF (%)} < 0", "column_id": "Perf. Cumulée PF (%)"},
         "color": "#f04f6a", "fontWeight": "600"},
        {"if": {"filter_query": "{Écart Mensuel (%)} > 0", "column_id": "Écart Mensuel (%)"},
         "color": "#00c896"},
        {"if": {"filter_query": "{Écart Mensuel (%)} < 0", "column_id": "Écart Mensuel (%)"},
         "color": "#f04f6a"},
        {"if": {"row_index": "odd"}, "backgroundColor": "rgba(255,255,255,0.02)"},
    ]

    table = dash_table.DataTable(
        data=rows,
        columns=[{"name": c, "id": c, "type": "numeric", "format": {"specifier": ".2f"}}
                 if c != "Date" else {"name": c, "id": c}
                 for c in cols_order],
        style_cell=TABLE_STYLE_CELL,
        style_header=TABLE_STYLE_HEADER,
        style_table={"overflowX": "auto", "minWidth": "100%"},
        style_data_conditional=style_cond,
        style_cell_conditional=[
            {"if": {"column_id": "Date"}, "fontWeight": "600", "width": "90px"},
        ] + [{"if": {"column_id": c}, "textAlign": "right"} for c in cols_order if c != "Date"],
        sort_action="native",
        page_size=24,
    )

    return fig, source_label, table


# 13. Ticker Analysis Table
@app.callback(
    Output("ticker-analysis-table", "children"),
    Input("store-analytics", "data"),
)
def update_ticker_analysis(data):
    if not data:
        return []

    d = json.loads(data)
    per_ticker = d.get("per_ticker", {})
    if not per_ticker:
        return html.Div("Données insuffisantes.", style={"padding": "16px", "color": "#94a3b8", "fontSize": "12px"})

    # Build rows — benchmark last
    rows = []
    for ticker, m in per_ticker.items():
        if ticker != engine.BENCHMARK:
            rows.append({
                "Titre": ticker,
                "Rendement Mensuel (%)": m["monthly_ret"],
                "Rendement Annuel (%)": m["yearly_ret"],
                "Volatilité Mensuelle (%)": m["monthly_vol"],
                "Volatilité Annuelle (%)": m["yearly_vol"],
                "Rendement / Volatilité": m["ret_vol"],
            })

    # Sort by yearly return descending
    rows.sort(key=lambda r: r["Rendement Annuel (%)"], reverse=True)

    # Append benchmark at the end
    bm = per_ticker.get(engine.BENCHMARK)
    if bm:
        rows.append({
            "Titre": engine.BENCHMARK,
            "Rendement Mensuel (%)": bm["monthly_ret"],
            "Rendement Annuel (%)": bm["yearly_ret"],
            "Volatilité Mensuelle (%)": bm["monthly_vol"],
            "Volatilité Annuelle (%)": bm["yearly_vol"],
            "Rendement / Volatilité": bm["ret_vol"],
        })

    cols = ["Titre", "Rendement Mensuel (%)", "Rendement Annuel (%)",
            "Volatilité Mensuelle (%)", "Volatilité Annuelle (%)", "Rendement / Volatilité"]

    # Conditional formatting
    style_cond = [
        # Benchmark row styling
        {
            "if": {"filter_query": '{Titre} = "' + engine.BENCHMARK + '"'},
            "backgroundColor": "rgba(197,160,89,0.06)",
            "fontWeight": "600",
            "borderTop": "1px solid rgba(197,160,89,0.22)",
        },
        # Rendement Annuel positive
        {
            "if": {"filter_query": "{Rendement Annuel (%)} > 0", "column_id": "Rendement Annuel (%)"},
            "color": "#00c896", "fontWeight": "600",
        },
        {
            "if": {"filter_query": "{Rendement Annuel (%)} < 0", "column_id": "Rendement Annuel (%)"},
            "color": "#f04f6a", "fontWeight": "600",
        },
        # Rendement / Volatilité — green if >= 0.7
        {
            "if": {"filter_query": "{Rendement / Volatilité} >= 0.7", "column_id": "Rendement / Volatilité"},
            "backgroundColor": "rgba(0,200,150,0.12)", "color": "#00c896", "fontWeight": "700",
        },
        # Rendement / Volatilité — orange if 0.2 <= x < 0.7
        {
            "if": {"filter_query": "{Rendement / Volatilité} >= 0.2 && {Rendement / Volatilité} < 0.7",
                   "column_id": "Rendement / Volatilité"},
            "color": "#d4b87a",
        },
        # Rendement / Volatilité — red if < 0.2
        {
            "if": {"filter_query": "{Rendement / Volatilité} < 0.2", "column_id": "Rendement / Volatilité"},
            "backgroundColor": "rgba(240,82,82,0.12)", "color": "#f04f6a", "fontWeight": "700",
        },
        {"if": {"row_index": "odd"}, "backgroundColor": "rgba(255,255,255,0.02)"},
    ]

    table = dash_table.DataTable(
        data=rows,
        columns=[{"name": c, "id": c, "type": "numeric", "format": {"specifier": ".2f"}}
                 if c != "Titre" else {"name": c, "id": c}
                 for c in cols],
        style_cell=TABLE_STYLE_CELL,
        style_header=TABLE_STYLE_HEADER,
        style_table={"overflowX": "auto", "minWidth": "100%"},
        style_data_conditional=style_cond,
        style_cell_conditional=(
            [{"if": {"column_id": "Titre"}, "fontWeight": "600", "color": "#E8EFF8", "width": "110px"}] +
            [{"if": {"column_id": c}, "textAlign": "right"} for c in cols[1:]]
        ),
        sort_action="native",
        page_size=30,
    )

    return table


# 14. Fetch current price for free-ticker simulator
@app.callback(
    Output("new-ticker-store", "data"),
    Output("new-ticker-price-display", "children"),
    Input("new-ticker-input", "value"),
    prevent_initial_call=True,
)
def fetch_new_ticker_price(ticker):
    if not ticker or len(ticker.strip()) < 1:
        return None, "—"
    t = ticker.strip().upper()
    try:
        prices = engine.fetch_current_prices([t])
        price = prices.get(t)
        if price is not None and not pd.isna(price):
            price = float(price)
            return price, html.Span(f"{price:,.2f} €", style={"color": "#E8EFF8", "fontWeight": "600"})
        return None, html.Span("Ticker introuvable", style={"color": "#f04f6a", "fontSize": "11px"})
    except Exception:
        return None, html.Span("Erreur de récupération", style={"color": "#f04f6a", "fontSize": "11px"})


# 15. Free-ticker simulator results
@app.callback(
    Output("new-ticker-results", "children"),
    Input("new-ticker-buy-price", "value"),
    Input("new-ticker-qty", "value"),
    State("new-ticker-input", "value"),
    State("new-ticker-store", "data"),
    State("store-analytics", "data"),
    prevent_initial_call=True,
)
def update_new_ticker_simulator(buy_price, qty, ticker, current_price, data_json):
    if buy_price is None or qty is None or not data_json:
        return []

    try:
        data = json.loads(data_json)
        total_portfolio = data["summary"].get("total_portfolio_value", 0)

        buy_price = float(buy_price)
        qty = float(qty)
        position_value = buy_price * qty
        new_total = total_portfolio + position_value
        weighting = (position_value / new_total * 100) if new_total > 0 else 0

        def result_card(label, value, color="#E8EFF8"):
            return html.Div([
                html.Div(label, className="calc-result-label"),
                html.Div(value, className="calc-result-value", style={"color": color}),
            ], className="calc-result-item")

        cards = [
            result_card("PRU", f"{buy_price:,.2f} €".replace(",", " ")),
            result_card("Quantité", f"{qty:,.1f}"),
            result_card("Valeur Position", f"{position_value:,.2f} €".replace(",", " "), "#c4a24a"),
            result_card("Pondération Résultante", f"{weighting:.2f} %", "#d4b87a"),
        ]

        if current_price is not None:
            gap_pct = (current_price - buy_price) / buy_price * 100
            sign = "+" if gap_pct >= 0 else ""
            color = "#00c896" if gap_pct >= 0 else "#f04f6a"
            label = "Décote vs Cours Actuel" if gap_pct < 0 else "Surcote vs Cours Actuel"
            cards.append(result_card(label, f"{sign}{gap_pct:.2f} %", color))

        return cards

    except Exception as e:
        print(f"DEBUG Error new ticker simulator: {e}")
        return []


# ─── Macro: fetch & store ─────────────────────────────────────────────────────

@app.callback(
    Output("store-macro", "data"),
    Output("macro-loading-trigger", "children"),
    Input("macro-refresh-btn", "n_clicks"),
    State("fred-api-key-input", "value"),
    prevent_initial_call=True,
)
def fetch_macro(n_clicks, fred_key):
    if not n_clicks:
        return no_update, no_update
    try:
        # Supprimer le cache pour forcer un refresh complet
        macro_engine.CACHE_PATH.unlink(missing_ok=True)
        data = macro_engine.get_macro_data(fred_api_key=fred_key or "")
        return json.dumps(data, default=str), ""
    except Exception as e:
        import traceback; traceback.print_exc()
        return None, ""


# ─── Macro: cycle strip (vue d'ensemble) ──────────────────────────────────────

@app.callback(
    Output("macro-cycle-strip", "children"),
    Input("store-macro", "data"),
)
def render_macro_cycle_strip(data_json):
    if not data_json:
        return html.Div([
            html.Div("📡", style={"fontSize": "40px", "marginBottom": "10px"}),
            html.Div("Entrez votre clé API FRED et cliquez sur Actualiser",
                     style={"color": "#94a3b8", "fontSize": "13px"}),
            html.A("→ Obtenir une clé gratuite",
                   href="https://fred.stlouisfed.org/docs/api/api_key.html",
                   target="_blank",
                   style={"color": "#c4a24a", "fontSize": "12px", "display": "block", "marginTop": "6px"}),
        ], style={"textAlign": "center", "padding": "40px 20px"})

    d      = json.loads(data_json)
    PHASES = macro_engine.CYCLE_PHASES
    META   = macro_engine.COUNTRY_META

    chips = []
    for code in ["US", "DE", "FR", "CN", "IN"]:
        cdata = d.get(code, {})
        phase = cdata.get("cycle", "inconnu")
        ph    = PHASES.get(phase, PHASES["inconnu"])
        m     = META[code]
        chips.append(html.Div([
            html.Span(m["flag"], style={"fontSize": "20px"}),
            html.Div(m["label"], style={"fontSize": "11px", "color": "#8ba8c4", "marginTop": "3px"}),
            html.Div(f"{ph['icon']} {ph['label']}",
                     style={"fontSize": "11px", "fontWeight": "700",
                            "color": ph["color"], "marginTop": "4px"}),
        ], style={
            "display": "flex", "flexDirection": "column", "alignItems": "center",
            "padding": "12px 24px", "borderRadius": "8px",
            "border": f"1px solid {ph['color']}33",
            "backgroundColor": f"{ph['color']}0d", "minWidth": "120px",
        }))

    return html.Div([
        html.Div([
            html.Span("Positionnement dans le Cycle", className="card-title"),
            html.Span(f"Mis à jour : {d.get('_fetched_at', '')}",
                      style={"fontSize": "10px", "color": "#94a3b8"}),
        ], className="card-header"),
        html.Div(chips, style={"display": "flex", "gap": "12px",
                               "flexWrap": "wrap", "padding": "16px"}),
    ], className="card")


# ─── Macro: sélection pays ────────────────────────────────────────────────────

@app.callback(
    Output("macro-country-store", "data"),
    [Input(f"mac-btn-{code}", "n_clicks") for code in ["US", "DE", "FR", "CN", "IN"]],
    prevent_initial_call=True,
)
def select_macro_country(*_):
    ctx = callback_context
    if not ctx.triggered:
        return no_update
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]
    return btn_id.replace("mac-btn-", "")


# ─── Macro: graphiques + tuiles (pays sélectionné) ───────────────────────────

@app.callback(
    Output("macro-gdp-chart",        "figure"),
    Output("macro-rate-chart",       "figure"),
    Output("macro-gdp-title",        "children"),
    Output("macro-rate-title",       "children"),
    Output("macro-indicator-tiles",  "children"),
    Input("store-macro",             "data"),
    Input("macro-country-store",     "data"),
)
def render_macro_detail(data_json, country):
    empty = go.Figure()
    empty.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    if not data_json or not country:
        return empty, empty, "—", "—", []
    try:
        return _render_macro_detail_impl(data_json, country)
    except Exception:
        import traceback; traceback.print_exc()
        return (empty, empty, "—", "—",
                html.Div("Erreur de rendu des données macro — cliquez sur Actualiser.",
                         style={"color": "#f04f6a", "padding": "16px"}))


def _render_macro_detail_impl(data_json, country):
    empty = go.Figure()
    empty.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    d     = json.loads(data_json)
    cdata = d.get(country, {})
    META  = macro_engine.COUNTRY_META
    PHASES = macro_engine.CYCLE_PHASES
    m     = META.get(country, {})
    flag  = m.get("flag", "")
    label = m.get("label", country)

    def _records_to_xy(records):
        if not records:
            return [], []
        xs = [r[0] for r in records]
        ys = [r[1] for r in records]
        return xs, ys

    # ── GDP chart avec phases en fond ────────────────────────────────────────
    gdp_fig = go.Figure()
    gdp_series = cdata.get("gdp_qq_series") or cdata.get("gdp_yoy_series") or []
    xs_g, ys_g = _records_to_xy(gdp_series)
    phase_spans = cdata.get("phase_spans", [])

    # Fond coloré par phase (sans étiquette — la légende des phases est sous le graphique)
    for span in phase_spans:
        ph    = span["phase"]
        color = macro_engine.PHASE_BG_COLORS.get(ph, "rgba(0,0,0,0)")
        if color == "rgba(0,0,0,0)":
            continue
        gdp_fig.add_vrect(
            x0=span["start"], x1=span["end"],
            fillcolor=color, opacity=1, layer="below", line_width=0,
        )

    if xs_g:
        gdp_fig.add_trace(go.Scatter(
            x=xs_g, y=ys_g,
            name="PIB QoQ (%)" if cdata.get("gdp_qq_latest") else "PIB YoY (%)",
            line=dict(color="#c4a24a", width=2.5),
            hovertemplate="%{x|%b %Y} : <b>%{y:.2f}%</b><extra></extra>",
        ))
        gdp_fig.add_hline(y=0, line_dash="dot",
                          line_color="rgba(255,255,255,0.25)", line_width=1)

    # Légende des phases (carrés de couleur) pour les phases présentes
    seen = []
    for span in phase_spans:
        ph = span["phase"]
        if ph not in seen and ph in macro_engine.CYCLE_PHASES and ph != "inconnu":
            seen.append(ph)
    for ph in seen:
        info = PHASES.get(ph, {})
        gdp_fig.add_trace(go.Scatter(
            x=[None], y=[None], mode="markers",
            marker=dict(size=10, symbol="square", color=info.get("color", "#64748b")),
            name=info.get("label", ph), hoverinfo="skip", showlegend=True,
        ))

    gdp_fig.update_layout(
        **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
        hovermode="x unified",
        yaxis_ticksuffix="%",
        legend=dict(orientation="h", y=1.08, x=0),
        margin=dict(l=16, r=16, t=40, b=16),
    )

    # ── Rate chart : taux 10 ans + taux court ────────────────────────────────
    rate_fig = go.Figure()
    y10_series = cdata.get("yield10_series", [])
    rs_series  = cdata.get("rate_short_series", [])
    xs_10, ys_10 = _records_to_xy(y10_series)
    xs_rs, ys_rs = _records_to_xy(rs_series)

    if xs_rs:
        rate_fig.add_trace(go.Scatter(
            x=xs_rs, y=ys_rs,
            name="Taux court (politique monétaire)",
            line=dict(color="#60a5fa", width=2, dash="dash"),
            hovertemplate="%{x|%b %Y} : <b>%{y:.2f}%</b><extra>Taux court</extra>",
        ))
    if xs_10:
        rate_fig.add_trace(go.Scatter(
            x=xs_10, y=ys_10,
            name="Taux 10 ans",
            line=dict(color="#c4a24a", width=2.5),
            fill="tonexty" if xs_rs else "none",
            fillcolor="rgba(196,162,74,0.08)",
            hovertemplate="%{x|%b %Y} : <b>%{y:.2f}%</b><extra>10 ans</extra>",
        ))
    if not xs_10 and not xs_rs:
        rate_fig.add_annotation(text="Données non disponibles",
                                xref="paper", yref="paper", x=0.5, y=0.5,
                                showarrow=False, font=dict(color="#64748b", size=13))

    rate_fig.update_layout(
        **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
        hovermode="x unified",
        yaxis_ticksuffix="%",
        legend=dict(orientation="h", y=1.08, x=0),
        margin=dict(l=16, r=16, t=40, b=16),
    )

    # ── Tuiles indicateurs cliquables ────────────────────────────────────────
    def _fmt(v, suf="", dec=2):
        if v is None: return "N/D"
        try:    return f"{float(v):.{dec}f}{suf}"
        except: return "N/D"

    def _arrow(t):
        if t is None: return ""
        return " ▲" if t > 0.05 else (" ▼" if t < -0.05 else " ▶")

    def _color_gdp(v):
        if v is None: return "#8ba8c4"
        return "#00c896" if v >= 1 else ("#f59e0b" if v >= 0 else "#f04f6a")

    def _color_unemp(t):
        if t is None: return "#8ba8c4"
        return "#00c896" if t < -0.05 else ("#f04f6a" if t > 0.05 else "#8ba8c4")

    def _color_infl(v):
        if v is None: return "#8ba8c4"
        return "#f04f6a" if v > 4 else ("#f59e0b" if v > 2 else "#00c896")

    def _mini_spark(records, color="#c4a24a"):
        if not records or len(records) < 3:
            return html.Div(style={"height": "44px"})
        ys = [r[1] for r in records[-24:]]
        xs = list(range(len(ys)))
        fig = go.Figure(go.Scatter(
            x=xs, y=ys, mode="lines",
            line=dict(color=color, width=1.5),
            fill="tozeroy", fillcolor=hex_rgba(color, 0.13),
        ))
        fig.update_layout(
            height=44, margin=dict(l=0, r=0, t=0, b=0),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(visible=False), yaxis=dict(visible=False), showlegend=False,
        )
        return dcc.Graph(figure=fig, config={"displayModeBar": False},
                         style={"height": "44px", "margin": "6px -12px 0"})

    gdp_val    = cdata.get("gdp_qq_latest") or cdata.get("gdp_yoy_latest")
    gdp_lbl    = "PIB QoQ" if cdata.get("gdp_qq_latest") else "PIB YoY"
    gdp_trend  = cdata.get("gdp_trend")
    unemp_val  = cdata.get("unemp_latest")
    u_trend    = cdata.get("unemp_trend")
    infl_val   = cdata.get("inflation_latest")
    i_trend    = cdata.get("inflation_trend")
    y10_val    = cdata.get("yield10_latest")
    y10_trend  = cdata.get("yield10_trend")
    rs_val     = cdata.get("rate_short_latest")
    rs_trend   = cdata.get("rate_short_trend")

    def _fmt_date(ds):
        if not ds:
            return ""
        try:
            return "à jour : " + pd.Timestamp(ds).strftime("%m/%Y")
        except Exception:
            return ""

    TILES = [
        (gdp_lbl,     f"{gdp_lbl}{_arrow(gdp_trend)}",   _fmt(gdp_val, "%"),  _color_gdp(gdp_val),   "gdp",        cdata.get("gdp_qq_series") or [], "#c4a24a", cdata.get("gdp_qq_date")),
        ("Chômage",   f"Chômage{_arrow(u_trend)}",        _fmt(unemp_val, "%"),_color_unemp(u_trend), "unemp",      cdata.get("unemp_series", []),     "#60a5fa", cdata.get("unemp_date")),
        ("Inflation", f"Inflation YoY{_arrow(i_trend)}",  _fmt(infl_val, "%"), _color_infl(infl_val), "inflation",  cdata.get("inflation_series", []), "#f59e0b", cdata.get("inflation_date")),
        ("Taux 10Y",  f"Taux 10 ans{_arrow(y10_trend)}",  _fmt(y10_val, "%"),  "#c4a24a",             "yield10",    cdata.get("yield10_series", []),   "#c4a24a", cdata.get("yield10_date")),
        ("Taux court",f"Taux court{_arrow(rs_trend)}",    _fmt(rs_val, "%"),   "#60a5fa",             "rate_short", cdata.get("rate_short_series", []),"#60a5fa", cdata.get("rate_short_date")),
    ]

    tiles = []
    for _, display_lbl, val_str, color, metric, hist, spark_color, tdate in TILES:
        tile = html.Div([
            html.Div(display_lbl, style={"fontSize": "10px", "color": "#64748b",
                                          "textTransform": "uppercase", "letterSpacing": "0.8px"}),
            html.Div(val_str, style={"fontSize": "26px", "fontWeight": "700",
                                      "color": color, "margin": "6px 0 2px"}),
            html.Div(_fmt_date(tdate), style={"fontSize": "9px", "color": "#475569"}),
            _mini_spark(hist, spark_color),
            html.Button(
                "Voir l'historique complet →",
                id={"type": "mac-ind-btn", "country": country, "metric": metric},
                n_clicks=0,
                style={
                    "background": "none", "border": "none", "color": "#c4a24a",
                    "fontSize": "10px", "cursor": "pointer", "padding": "6px 0 0",
                    "fontFamily": "Inter", "letterSpacing": "0.3px",
                }
            ),
        ], style={
            "flex": "1", "minWidth": "180px",
            "backgroundColor": "#07101f",
            "border": "1px solid rgba(255,255,255,0.06)",
            "borderRadius": "10px", "padding": "16px 16px 12px",
        })
        tiles.append(tile)

    phase    = cdata.get("cycle", "inconnu")
    ph_info  = PHASES.get(phase, PHASES["inconnu"])
    src_lbl  = cdata.get("source", "")

    tiles_section = html.Div([
        html.Div([
            html.Div([
                html.Span(f"{flag} {label}", className="card-title"),
                html.Span(f"{ph_info['icon']} {ph_info['label']}",
                          style={"fontSize": "12px", "fontWeight": "700",
                                 "color": ph_info["color"],
                                 "backgroundColor": f"{ph_info['color']}1a",
                                 "padding": "3px 10px", "borderRadius": "12px"}),
            ], style={"display": "flex", "alignItems": "center", "gap": "12px"}),
            html.Span(f"Source : {src_lbl}",
                      style={"fontSize": "10px", "color": "#334155"}),
        ], className="card-header", style={"justifyContent": "space-between"}),
        html.Div(tiles, style={"display": "flex", "gap": "12px",
                               "flexWrap": "wrap", "padding": "16px"}),
    ], className="card")

    gdp_title  = f"Croissance PIB — {flag} {label}"
    rate_title = f"Courbe de Taux — {flag} {label}"
    return gdp_fig, rate_fig, gdp_title, rate_title, tiles_section


# ─── Macro: rotation sectorielle ─────────────────────────────────────────────

@app.callback(
    Output("macro-rotation-section", "children"),
    Input("store-macro", "data"),
)
def render_macro_rotation(data_json):
    if not data_json:
        return []
    try:
        return _render_macro_rotation_impl(data_json)
    except Exception:
        import traceback; traceback.print_exc()
        return html.Div("Erreur de rendu de la rotation sectorielle.",
                        style={"color": "#f04f6a", "padding": "16px"})


def _render_macro_rotation_impl(data_json):
    d = json.loads(data_json)

    SECTORS = {
        "reprise":        ("#60a5fa",  "Finance, Technologie, Conso. Discrétionnaire, Immobilier"),
        "expansion":      ("#00c896",  "Industrie, Matériaux, Services de Communication"),
        "surchauffe":     ("#f59e0b",  "Énergie, Matières Premières, Pétrole/Gaz"),
        "ralentissement": ("#f04f6a",  "Santé, Conso. de Base, Cash"),
        "recession":      ("#dc2626",  "Obligations, Services Publics (Utilities), Conso. de Base"),
    }
    PHASES = macro_engine.CYCLE_PHASES
    META   = macro_engine.COUNTRY_META

    # ── Horloge des cycles (Investment Clock) ────────────────────────────────
    # Axe X = dynamique de croissance (← ralentit | accélère →)
    # Axe Y = pression inflationniste (↑ haute | basse ↓)
    QUADRANTS = [
        # (x0, x1, y0, y1, phase_key, secteurs court)
        ( 0,  1,  0,  1, "expansion",      "Industrie · Matériaux · Communication"),
        ( 0,  1,  0, -1, "reprise",        "Actions · Finance · Tech · Immobilier"),
        ( 0, -1,  0,  1, "ralentissement", "Santé · Conso. base · Cash"),
        ( 0, -1,  0, -1, "recession",      "Obligations · Utilities · Conso. base"),
    ]
    clock = go.Figure()
    for x0, x1, y0, y1, phase, secs in QUADRANTS:
        ph    = PHASES.get(phase, PHASES["inconnu"])
        color = ph["color"]
        clock.add_shape(type="rect", x0=x0, y0=y0, x1=x1, y1=y1,
                        fillcolor=hex_rgba(color, 0.12),
                        line=dict(color=hex_rgba(color, 0.33), width=1),
                        layer="below")
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        clock.add_annotation(x=cx, y=cy + 0.12 * (1 if cy >= 0 else -1),
                             text=f"<b>{ph['icon']} {ph['label']}</b>",
                             showarrow=False, font=dict(size=13, color=color))
        clock.add_annotation(x=cx, y=cy - 0.14 * (1 if cy >= 0 else -1),
                             text=secs, showarrow=False,
                             font=dict(size=9, color="#8ba8c4"),
                             align="center")

    # Axes médians
    clock.add_hline(y=0, line_color="rgba(255,255,255,0.25)", line_width=1)
    clock.add_vline(x=0, line_color="rgba(255,255,255,0.25)", line_width=1)

    def _clamp(v, lo=-0.9, hi=0.9):
        return max(lo, min(hi, v))

    # Positionner chaque pays selon (tendance croissance, inflation vs cible 2 %)
    px, py, ptext, pcolor = [], [], [], []
    for code, meta in META.items():
        cdata = d.get(code, {})
        g_tr  = cdata.get("gdp_trend")
        infl  = cdata.get("inflation_latest")
        phase = cdata.get("cycle", "inconnu")
        if g_tr is None or infl is None:
            continue
        x = _clamp(g_tr * 1.5)
        y = _clamp((infl - 2.0) * 0.4)
        px.append(x); py.append(y)
        ptext.append(f"{meta.get('flag','')} {code}")
        pcolor.append(PHASES.get(phase, PHASES["inconnu"])["color"])

    if px:
        clock.add_trace(go.Scatter(
            x=px, y=py, mode="markers+text",
            text=ptext, textposition="top center",
            textfont=dict(size=12, color="#ffffff", family="Inter"),
            marker=dict(size=16, color=pcolor, line=dict(color="#ffffff", width=1.5)),
            hovertemplate="%{text}<extra></extra>",
        ))

    clock.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        height=420, showlegend=False,
        margin=dict(l=10, r=10, t=30, b=30),
        xaxis=dict(range=[-1.05, 1.05], zeroline=False, showgrid=False,
                   showticklabels=False,
                   title=dict(text="←  Croissance ralentit       Croissance accélère  →",
                              font=dict(size=10, color="#64748b"))),
        yaxis=dict(range=[-1.05, 1.05], zeroline=False, showgrid=False,
                   showticklabels=False, scaleanchor="x", scaleratio=1,
                   title=dict(text="←  Inflation basse        Inflation haute  →",
                              font=dict(size=10, color="#64748b"))),
    )

    # ── Cartes secteurs par pays (5 pays) ────────────────────────────────────
    def _phase_card(code):
        meta  = META.get(code, {})
        lbl   = f"{meta.get('flag','')} {meta.get('label', code)}"
        phase = d.get(code, {}).get("cycle", "inconnu")
        ph    = PHASES.get(phase, PHASES["inconnu"])
        if phase not in SECTORS:
            return html.Div([
                html.Div(lbl, style={"fontSize": "11px", "color": "#94a3b8",
                                     "fontWeight": "600"}),
                html.Div("Cycle indéterminé", style={"color": "#64748b",
                                                      "marginTop": "6px", "fontSize": "12px"}),
            ], style={"flex": "1", "minWidth": "150px", "padding": "14px",
                      "borderRadius": "8px", "border": "1px solid rgba(255,255,255,0.06)"})
        color, secs = SECTORS[phase]
        return html.Div([
            html.Div(lbl, style={"fontSize": "11px", "color": "#94a3b8", "fontWeight": "600"}),
            html.Div(f"{ph['icon']} {ph['label']}",
                     style={"fontSize": "14px", "fontWeight": "700",
                            "color": color, "margin": "4px 0 8px"}),
            html.Div("Secteurs favorisés :", style={"fontSize": "10px", "color": "#64748b",
                                                     "marginBottom": "4px"}),
            html.Div(secs, style={"fontSize": "12px", "color": "#8ba8c4", "lineHeight": "1.5"}),
        ], style={
            "flex": "1", "minWidth": "150px", "padding": "14px 16px", "borderRadius": "8px",
            "backgroundColor": f"{color}0d", "border": f"1px solid {color}33",
        })

    return html.Div([
        html.Div([
            html.Span("Horloge des Cycles & Rotation Sectorielle", className="card-title"),
            html.Span("Position des pays selon croissance × inflation — indicatif",
                      style={"fontSize": "10px", "color": "#94a3b8"}),
        ], className="card-header"),
        dcc.Graph(figure=clock, config={"displayModeBar": False},
                  style={"padding": "8px 16px 0"}),
        html.Div([_phase_card(c) for c in META.keys()],
                 style={"display": "flex", "gap": "12px", "padding": "16px",
                        "flexWrap": "wrap"}),
    ], className="card")


# ─── Macro: spread de taux (signal de récession) ──────────────────────────────

@app.callback(
    Output("macro-spread-section", "children"),
    Input("store-macro", "data"),
    Input("macro-country-store", "data"),
)
def render_macro_spread(data_json, country):
    if not data_json or not country:
        return []
    try:
        d = json.loads(data_json)
        cdata = d.get(country, {})
        sp = macro_engine.compute_yield_spread(
            cdata.get("yield10_series", []), cdata.get("rate_short_series", []))
        meta = macro_engine.COUNTRY_META.get(country, {})
        if not sp["series"]:
            return html.Div([
                html.Div([html.Span("Spread de Taux (10 ans − court)", className="card-title")],
                         className="card-header"),
                html.Div(f"Données de taux indisponibles pour {meta.get('label', country)}.",
                         style={"padding": "16px", "color": "#64748b", "fontSize": "12px"}),
            ], className="card")

        SIG_COLORS = {"inverse": "#dc2626", "plat": "#f59e0b",
                      "normal": "#00c896", "n/a": "#64748b"}
        color = SIG_COLORS.get(sp["signal"], "#64748b")
        xs = [r[0] for r in sp["series"]]
        ys = [r[1] for r in sp["series"]]

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines", line=dict(color=color, width=2),
            fill="tozeroy", fillcolor=hex_rgba(color, 0.10),
            hovertemplate="%{x|%b %Y} : <b>%{y:.2f} pts</b><extra></extra>",
        ))
        fig.add_hline(y=0, line_dash="dot", line_color="rgba(255,255,255,0.35)", line_width=1)
        fig.update_layout(
            **{k: v for k, v in CHART_LAYOUT.items() if k not in ("margin", "legend")},
            hovermode="x unified", yaxis_ticksuffix=" pts",
            margin=dict(l=16, r=16, t=16, b=16), height=240, showlegend=False,
        )
        return html.Div([
            html.Div([
                html.Span(f"Spread de Taux — {meta.get('flag','')} {meta.get('label', country)}",
                          className="card-title"),
                html.Span(f"{sp['label']}",
                          style={"fontSize": "11px", "fontWeight": "700", "color": color,
                                 "backgroundColor": hex_rgba(color, 0.12),
                                 "padding": "3px 10px", "borderRadius": "12px"}),
            ], className="card-header", style={"justifyContent": "space-between"}),
            html.Div([
                html.Div([
                    html.Div("Spread actuel (10 ans − court)",
                             style={"fontSize": "10px", "color": "#64748b",
                                    "textTransform": "uppercase", "letterSpacing": "0.8px"}),
                    html.Div(f"{sp['latest']:+.2f} pts",
                             style={"fontSize": "30px", "fontWeight": "700", "color": color}),
                    html.Div("Un spread négatif (courbe inversée) précède historiquement "
                             "les récessions de 6 à 18 mois.",
                             style={"fontSize": "11px", "color": "#8ba8c4",
                                    "marginTop": "6px", "maxWidth": "260px", "lineHeight": "1.5"}),
                ], style={"minWidth": "240px", "padding": "8px 4px"}),
                html.Div(dcc.Graph(figure=fig, config={"displayModeBar": False}),
                         style={"flex": "1", "minWidth": "320px"}),
            ], style={"display": "flex", "gap": "16px", "padding": "12px 16px",
                      "flexWrap": "wrap", "alignItems": "center"}),
        ], className="card")
    except Exception:
        import traceback; traceback.print_exc()
        return html.Div("Erreur de rendu du spread.", style={"color": "#f04f6a", "padding": "16px"})


# ─── Macro: portefeuille & cycle (actuel + à venir) ───────────────────────────

@app.callback(
    Output("macro-portfolio-cycle", "children"),
    Input("store-analytics", "data"),
    Input("store-macro", "data"),
)
def render_portfolio_cycle(analytics_json, macro_json):
    if not analytics_json or not macro_json:
        return []
    try:
        a = json.loads(analytics_json)
        d = json.loads(macro_json)
        positions = a.get("positions", [])
        if not positions:
            return []

        REF = "FR"  # portefeuille majoritairement français/européen
        cur_phase = d.get(REF, {}).get("cycle", "inconnu")
        nxt_phase = macro_engine.NEXT_PHASE.get(cur_phase, "inconnu")
        PHASES = macro_engine.CYCLE_PHASES
        cur_secs = set(macro_engine.PHASE_SECTORS.get(cur_phase, []))
        nxt_secs = set(macro_engine.PHASE_SECTORS.get(nxt_phase, []))

        tickers = [p.get("Ticker") for p in positions if p.get("Ticker")]
        sectors = macro_engine.get_ticker_sectors(tickers)

        total_val = sum(float(p.get("Valeur (€)", 0) or 0) for p in positions)
        groups = {"now": [], "next": [], "neutre": [], "diversifie": []}
        val_now = val_next = 0.0
        for p in positions:
            t = p.get("Ticker"); val = float(p.get("Valeur (€)", 0) or 0)
            sec = sectors.get(t)
            row = {"ticker": t, "sec": sec, "val": val}
            if sec is None:
                groups["diversifie"].append(row)
            elif sec in cur_secs:
                groups["now"].append(row); val_now += val
            elif sec in nxt_secs:
                groups["next"].append(row); val_next += val
            else:
                groups["neutre"].append(row)

        pct_now  = (val_now / total_val * 100) if total_val else 0
        pct_next = (val_next / total_val * 100) if total_val else 0
        ph_cur = PHASES.get(cur_phase, PHASES["inconnu"])
        ph_nxt = PHASES.get(nxt_phase, PHASES["inconnu"])

        def _chip(row, color):
            label = f"{row['ticker']}"
            sub = row['sec'] or "ETF / Indice"
            return html.Div([
                html.Span(label, style={"fontWeight": "700", "color": "#e8eef6", "fontSize": "13px"}),
                html.Span(sub, style={"fontSize": "10px", "color": "#8ba8c4", "marginLeft": "6px"}),
                html.Span(fmt_eur(row['val']),
                          style={"fontSize": "11px", "color": "#94a3b8", "marginLeft": "8px"}),
            ], style={"padding": "6px 10px", "borderRadius": "8px",
                      "backgroundColor": hex_rgba(color, 0.10),
                      "border": f"1px solid {hex_rgba(color, 0.3)}",
                      "display": "inline-flex", "alignItems": "center",
                      "gap": "2px", "margin": "3px"})

        def _bucket(title, subtitle, rows, color):
            if not rows:
                body = html.Div("Aucun titre", style={"color": "#64748b", "fontSize": "12px",
                                                        "padding": "4px"})
            else:
                rows = sorted(rows, key=lambda r: -r["val"])
                body = html.Div([_chip(r, color) for r in rows],
                                style={"display": "flex", "flexWrap": "wrap"})
            return html.Div([
                html.Div([
                    html.Span(title, style={"fontSize": "12px", "fontWeight": "700", "color": color}),
                    html.Span(subtitle, style={"fontSize": "10px", "color": "#64748b",
                                               "marginLeft": "8px"}),
                ], style={"marginBottom": "8px"}),
                body,
            ], style={"flex": "1", "minWidth": "260px", "padding": "12px",
                      "borderRadius": "10px", "border": "1px solid rgba(255,255,255,0.06)",
                      "backgroundColor": "#07101f"})

        secs_cur = ", ".join(macro_engine.PHASE_SECTORS.get(cur_phase, [])) or "—"
        secs_nxt = ", ".join(macro_engine.PHASE_SECTORS.get(nxt_phase, [])) or "—"

        return html.Div([
            html.Div([
                html.Span("Portefeuille & Cycle Économique", className="card-title"),
                html.Span("Référence : 🇫🇷 France / zone euro — indicatif, non un conseil",
                          style={"fontSize": "10px", "color": "#94a3b8"}),
            ], className="card-header"),
            html.Div([
                # Bandeau phase actuelle → phase à venir
                html.Div([
                    html.Div([
                        html.Div("PHASE ACTUELLE", style={"fontSize": "9px", "color": "#64748b",
                                  "letterSpacing": "1px"}),
                        html.Div(f"{ph_cur['icon']} {ph_cur['label']}",
                                 style={"fontSize": "18px", "fontWeight": "700",
                                        "color": ph_cur["color"]}),
                        html.Div(f"Secteurs porteurs : {secs_cur}",
                                 style={"fontSize": "10px", "color": "#8ba8c4", "marginTop": "4px",
                                        "maxWidth": "300px", "lineHeight": "1.4"}),
                    ]),
                    html.Div("→", style={"fontSize": "26px", "color": "#475569",
                                          "alignSelf": "center"}),
                    html.Div([
                        html.Div("PHASE À VENIR (anticipée)", style={"fontSize": "9px",
                                  "color": "#64748b", "letterSpacing": "1px"}),
                        html.Div(f"{ph_nxt['icon']} {ph_nxt['label']}",
                                 style={"fontSize": "18px", "fontWeight": "700",
                                        "color": ph_nxt["color"]}),
                        html.Div(f"Secteurs à anticiper : {secs_nxt}",
                                 style={"fontSize": "10px", "color": "#8ba8c4", "marginTop": "4px",
                                        "maxWidth": "300px", "lineHeight": "1.4"}),
                    ]),
                ], style={"display": "flex", "gap": "20px", "padding": "8px 4px 16px",
                          "flexWrap": "wrap"}),
                # KPIs d'alignement
                html.Div([
                    html.Div([
                        html.Div(f"{pct_now:.0f}%", style={"fontSize": "26px", "fontWeight": "700",
                                  "color": ph_cur["color"]}),
                        html.Div("du portefeuille aligné avec la phase actuelle",
                                 style={"fontSize": "11px", "color": "#8ba8c4"}),
                    ], style={"flex": "1", "minWidth": "200px"}),
                    html.Div([
                        html.Div(f"{pct_next:.0f}%", style={"fontSize": "26px", "fontWeight": "700",
                                  "color": ph_nxt["color"]}),
                        html.Div("positionné pour la phase à venir",
                                 style={"fontSize": "11px", "color": "#8ba8c4"}),
                    ], style={"flex": "1", "minWidth": "200px"}),
                ], style={"display": "flex", "gap": "16px", "padding": "0 4px 16px"}),
                # Buckets de titres
                html.Div([
                    _bucket("✅ Alignés (phase actuelle)", f"{ph_cur['label']}",
                            groups["now"], ph_cur["color"]),
                    _bucket("⏭️ À anticiper (phase à venir)", f"{ph_nxt['label']}",
                            groups["next"], ph_nxt["color"]),
                    _bucket("➖ Neutres / contra-cycliques", "Hors secteurs porteurs",
                            groups["neutre"], "#64748b"),
                    _bucket("🧺 Diversifiés", "ETF / Indices",
                            groups["diversifie"], "#3d8fd1"),
                ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap"}),
            ], style={"padding": "16px"}),
        ], className="card")
    except Exception:
        import traceback; traceback.print_exc()
        return html.Div("Erreur de rendu du lien portefeuille/cycle.",
                        style={"color": "#f04f6a", "padding": "16px"})


# ─── Macro: FX & matières premières ────────────────────────────────────────────

@app.callback(
    Output("macro-markets-section", "children"),
    Input("store-macro", "data"),
)
def render_markets(macro_json):
    if not macro_json:
        return []
    try:
        mk = macro_engine.fetch_market_data()
        if not mk:
            return []

        def _spark(records, color):
            if not records or len(records) < 3:
                return html.Div(style={"height": "50px"})
            ys = [r[1] for r in records]
            fig = go.Figure(go.Scatter(y=ys, x=list(range(len(ys))), mode="lines",
                                       line=dict(color=color, width=1.5),
                                       fill="tozeroy", fillcolor=hex_rgba(color, 0.12)))
            fig.update_layout(height=50, margin=dict(l=0, r=0, t=0, b=0),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              xaxis=dict(visible=False), yaxis=dict(visible=False),
                              showlegend=False)
            return dcc.Graph(figure=fig, config={"displayModeBar": False},
                             style={"height": "50px", "margin": "6px -8px 0"})

        cards = []
        for key, meta in macro_engine.MARKET_TICKERS.items():
            v = mk.get(key, {})
            latest = v.get("latest"); yoy = v.get("yoy")
            if latest is None:
                continue
            yoy_color = "#00c896" if (yoy or 0) >= 0 else "#f04f6a"
            unit = meta["unit"]
            val_str = (f"{latest:,.4f}" if key == "EURUSD" else f"{latest:,.2f}") + unit
            cards.append(html.Div([
                html.Div([
                    html.Span(meta["icon"], style={"fontSize": "18px"}),
                    html.Span(meta["label"], style={"fontSize": "11px", "color": "#94a3b8",
                              "marginLeft": "6px", "fontWeight": "600"}),
                ]),
                html.Div(val_str, style={"fontSize": "24px", "fontWeight": "700",
                          "color": "#e8eef6", "margin": "6px 0 0"}),
                html.Div(f"{yoy:+.1f}% sur 1 an" if yoy is not None else "",
                         style={"fontSize": "11px", "color": yoy_color}),
                _spark(v.get("series", []), yoy_color),
            ], style={"flex": "1", "minWidth": "200px", "padding": "16px",
                      "borderRadius": "10px", "backgroundColor": "#07101f",
                      "border": "1px solid rgba(255,255,255,0.06)"}))

        return html.Div([
            html.Div([
                html.Span("Devises & Matières Premières", className="card-title"),
                html.Span("EUR/USD · Or · Brent — niveaux mensuels (yfinance)",
                          style={"fontSize": "10px", "color": "#94a3b8"}),
            ], className="card-header"),
            html.Div(cards, style={"display": "flex", "gap": "12px", "padding": "16px",
                                   "flexWrap": "wrap"}),
        ], className="card")
    except Exception:
        import traceback; traceback.print_exc()
        return html.Div("Erreur de rendu des marchés.",
                        style={"color": "#f04f6a", "padding": "16px"})


# ─── Macro: modal indicateur (historique complet) ─────────────────────────────

@app.callback(
    Output("macro-modal",       "is_open"),
    Output("macro-modal-title", "children"),
    Output("macro-modal-chart", "figure"),
    Input({"type": "mac-ind-btn", "country": ALL, "metric": ALL}, "n_clicks"),
    Input("macro-modal-close",  "n_clicks"),
    State("store-macro",        "data"),
    prevent_initial_call=True,
)
def toggle_indicator_modal(tile_clicks, close_clicks, data_json):
    ctx = callback_context
    if not ctx.triggered or not data_json:
        return False, "", go.Figure()

    trigger = ctx.triggered[0]["prop_id"]
    if "macro-modal-close" in trigger:
        return False, "", go.Figure()

    # Ignorer le déclenchement au montage des tuiles (n_clicks = 0/None)
    if not ctx.triggered[0].get("value"):
        return no_update, no_update, no_update

    # Identifier le bouton cliqué
    try:
        btn_info = json.loads(trigger.split(".")[0])
    except Exception:
        return False, "", go.Figure()

    country  = btn_info.get("country", "US")
    metric   = btn_info.get("metric", "gdp")
    d        = json.loads(data_json)
    cdata    = d.get(country, {})
    META     = macro_engine.COUNTRY_META
    m        = META.get(country, {})

    SERIES_MAP = {
        "gdp":       ("gdp_qq_series",    "PIB QoQ (%)",         "#c4a24a", True),
        "unemp":     ("unemp_series",     "Taux de Chômage (%)", "#60a5fa", False),
        "inflation": ("inflation_series", "Inflation YoY (%)",   "#f59e0b", True),
        "yield10":   ("yield10_series",   "Taux 10 ans (%)",     "#c4a24a", True),
        "rate_short":("rate_short_series","Taux court — politique monétaire (%)", "#60a5fa", True),
    }
    series_key, y_label, color, zero_ref = SERIES_MAP.get(metric, ("gdp_qq_series", "", "#c4a24a", False))
    records = cdata.get(series_key, [])

    title = f"{y_label} — {m.get('flag', '')} {m.get('label', country)} (depuis 2000)"

    fig = go.Figure()
    if records:
        xs = [r[0] for r in records]
        ys = [r[1] for r in records]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, name=y_label,
            line=dict(color=color, width=2),
            fill="tozeroy", fillcolor=hex_rgba(color, 0.08),
            hovertemplate="%{x|%b %Y} : <b>%{y:.2f}%</b><extra></extra>",
        ))
        if zero_ref:
            fig.add_hline(y=0, line_dash="dot",
                          line_color="rgba(255,255,255,0.25)", line_width=1)

        # Annoter min/max
        y_arr  = np.array(ys)
        idx_mx = int(np.argmax(y_arr))
        idx_mn = int(np.argmin(y_arr))
        for idx, symbol, pos in [(idx_mx, "▲", "top center"), (idx_mn, "▼", "bottom center")]:
            fig.add_annotation(
                x=xs[idx], y=ys[idx],
                text=f"{ys[idx]:.2f}%",
                showarrow=True, arrowhead=2, arrowsize=0.8,
                arrowcolor=color, font=dict(size=10, color=color),
                ax=0, ay=-30 if symbol == "▲" else 30,
            )
    else:
        fig.add_annotation(text="Données non disponibles", xref="paper", yref="paper",
                           x=0.5, y=0.5, showarrow=False,
                           font=dict(color="#64748b", size=14))

    fig.update_layout(
        **{k: v for k, v in CHART_LAYOUT.items()
           if k not in ("margin", "paper_bgcolor", "plot_bgcolor")},
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis_ticksuffix="%",
        hovermode="x unified",
        paper_bgcolor="#07101f",
        plot_bgcolor="#07101f",
    )
    return True, title, fig


# ─── PDF Export ───────────────────────────────────────────────────────────────

@app.callback(
    Output("download-pdf", "data"),
    Input("export-pdf-btn", "n_clicks"),
    State("store-data", "data"),
    State("store-analytics", "data"),
    prevent_initial_call=True,
)
def export_pdf(n_clicks, store_data, store_analytics):
    if not n_clicks or not store_analytics:
        return no_update

    try:
        d = json.loads(store_analytics)

        summary      = d.get("summary", {})
        positions    = d.get("positions", [])
        metrics      = d.get("metrics", {})
        per_ticker   = d.get("per_ticker", {})

        pdf_bytes = pdf_report.build_pdf(summary, positions, metrics, per_ticker)
        date_str  = datetime.now().strftime("%Y%m%d_%H%M")
        return dcc.send_bytes(pdf_bytes, f"portfolio_report_{date_str}.pdf")

    except Exception as e:
        import traceback
        traceback.print_exc()
        return no_update


# ─── Run ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import socket
    try:
        local_ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        local_ip = "introuvable"
    print("=" * 48)
    print("  Portfolio Dashboard — en ligne")
    print("=" * 48)
    print(f"  PC         : http://127.0.0.1:8050")
    print(f"  Tablette   : http://{local_ip}:8050")
    print("=" * 48)
    app.run(debug=False, port=8050, host="0.0.0.0")
