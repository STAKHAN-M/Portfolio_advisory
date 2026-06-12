import sys; sys.path.insert(0, '.')
import pdf_report

summary = {
    'total_portfolio_value': 12500, 'total_value': 11800, 'cash_balance': 700,
    'total_invested': 9000, 'total_dividends': 350, 'unrealized_pnl': 2800,
    'unrealized_pnl_pct': 31.1, 'total_return_pct': 38.9
}
positions = [
    {'Ticker': 'MC.PA',   'Description': 'LVMH',          'Categorie': 'CAC 40', 'Quantite': 5,  'PRU': 700, 'Cours': 762, 'Valeur': 3810, 'PnL': 310,  'PnL_pct': 8.86, 'Poids (%)': 32.3, 'Catégorie': 'CAC 40', 'Quantité': 5, 'PRU (€)': 700, 'Cours (€)': 762, 'Valeur (€)': 3810, 'P&L (€)': 310,  'P&L (%)': 8.86},
    {'Ticker': 'CW8.PA',  'Description': 'Amundi MSCI W', 'Catégorie': 'Index',  'Quantité': 12, 'PRU (€)': 380, 'Cours (€)': 412, 'Valeur (€)': 4944, 'P&L (€)': 384,  'P&L (%)': 8.42, 'Poids (%)': 41.9},
    {'Ticker': 'BNKE.PA', 'Description': 'BNP ETF',       'Catégorie': 'Index',  'Quantité': 20, 'PRU (€)': 15,  'Cours (€)': 17,  'Valeur (€)': 340,  'P&L (€)': 40,   'P&L (%)': 13.3, 'Poids (%)': 12.1},
    {'Ticker': 'TTE.PA',  'Description': 'TotalEnergies', 'Catégorie': 'CAC 40', 'Quantité': 10, 'PRU (€)': 56,  'Cours (€)': 52,  'Valeur (€)': 520,  'P&L (€)': -40,  'P&L (%)': -7.1, 'Poids (%)': 8.7},
    {'Ticker': 'OR.PA',   'Description': 'LOreal',        'Catégorie': 'CAC 40', 'Quantité': 3,  'PRU (€)': 390, 'Cours (€)': 402, 'Valeur (€)': 1206, 'P&L (€)': 36,   'P&L (%)': 3.1,  'Poids (%)': 5.0},
]
metrics = {
    'ann_ret': 18.4, 'ann_vol': 14.2, 'sharpe': 1.18, 'sortino': 1.52,
    'max_dd': -11.3, 'var_hist': 4.1, 'var_gauss': 3.8, 'cvar': 6.2,
    'beta': 0.87, 'alpha': 3.2
}
per_ticker = {p['Ticker']: {'monthly_ret': 1.2, 'yearly_ret': 15.0, 'monthly_vol': 3.8, 'yearly_vol': 13.2, 'ret_vol': 1.14} for p in positions}

pdf_bytes = pdf_report.build_pdf(summary, positions, metrics, per_ticker)
with open('test_report.pdf', 'wb') as f:
    f.write(pdf_bytes)
print("OK", len(pdf_bytes), "bytes")
