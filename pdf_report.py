"""
PDF Report Generator — Portfolio Dashboard Analysis
Uses reportlab to produce a clean, professional financial report.
"""

import io
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, HRFlowable, PageTemplate,
    Paragraph, Spacer, Table, TableStyle, KeepTogether,
)
from reportlab.graphics.shapes import Drawing, String, Circle
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF

# ─── Color Palette ────────────────────────────────────────────────────────────

C_BG_DARK  = colors.HexColor("#0d1929")
C_BG_MID   = colors.HexColor("#111e30")
C_BG_LIGHT = colors.HexColor("#f5f7fa")
C_GOLD     = colors.HexColor("#c4a24a")
C_GOLD_L   = colors.HexColor("#d4b870")
C_GREEN    = colors.HexColor("#00a87a")
C_RED      = colors.HexColor("#d63b54")
C_TEXT     = colors.HexColor("#1a2535")
C_TEXT_SEC = colors.HexColor("#4a6080")
C_TEXT_LT  = colors.HexColor("#dde6f0")
C_BORDER   = colors.HexColor("#d8e2ef")
C_ROW_ALT  = colors.HexColor("#eef2f7")
C_WHITE    = colors.white

PAGE_W, PAGE_H = landscape(A4)
MARGIN = 16 * mm

# Pie chart color palettes
PIE_PALETTE = [
    colors.HexColor("#c4a24a"), colors.HexColor("#3d8fd1"), colors.HexColor("#00a87a"),
    colors.HexColor("#d63b54"), colors.HexColor("#7b5ea7"), colors.HexColor("#e07b39"),
    colors.HexColor("#2e9cca"), colors.HexColor("#54b88e"), colors.HexColor("#e05c8a"),
    colors.HexColor("#8ab4d4"), colors.HexColor("#a8c580"), colors.HexColor("#f0c96e"),
]
CAT_PALETTE = {
    "CAC 40": colors.HexColor("#c4a24a"),
    "Index":  colors.HexColor("#3d8fd1"),
    "Action": colors.HexColor("#00a87a"),
}


# ─── Styles ───────────────────────────────────────────────────────────────────

def _styles():
    s = getSampleStyleSheet()
    base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=C_TEXT)

    return {
        "title": ParagraphStyle("title", **{**base,
            "fontName": "Helvetica-Bold", "fontSize": 18,
            "textColor": C_WHITE, "leading": 22, "spaceAfter": 2}),
        "subtitle": ParagraphStyle("subtitle", **{**base,
            "fontSize": 8.5, "textColor": C_GOLD,
            "letterSpacing": 2, "spaceAfter": 0}),
        "section": ParagraphStyle("section", **{**base,
            "fontName": "Helvetica-Bold", "fontSize": 8,
            "textColor": C_GOLD, "letterSpacing": 1.5,
            "spaceBefore": 10, "spaceAfter": 5}),
        "body": ParagraphStyle("body", **base),
        "small": ParagraphStyle("small", **{**base, "fontSize": 7.5, "textColor": C_TEXT_SEC}),
        "footer": ParagraphStyle("footer", **{**base,
            "fontSize": 7, "textColor": C_TEXT_SEC, "alignment": TA_CENTER}),
        "kpi_label": ParagraphStyle("kpi_label", **{**base,
            "fontName": "Helvetica-Bold", "fontSize": 7,
            "textColor": C_TEXT_SEC, "letterSpacing": 1, "alignment": TA_CENTER}),
        "kpi_value": ParagraphStyle("kpi_value", **{**base,
            "fontName": "Helvetica-Bold", "fontSize": 13,
            "textColor": C_TEXT, "alignment": TA_CENTER, "leading": 16}),
        "kpi_badge": ParagraphStyle("kpi_badge", **{**base,
            "fontSize": 8, "alignment": TA_CENTER}),
        "th": ParagraphStyle("th", **{**base,
            "fontName": "Helvetica-Bold", "fontSize": 7,
            "textColor": C_GOLD, "letterSpacing": 0.8, "alignment": TA_CENTER}),
        "td": ParagraphStyle("td", **{**base,
            "fontSize": 8, "alignment": TA_CENTER}),
        "td_left": ParagraphStyle("td_left", **{**base,
            "fontSize": 8, "alignment": TA_LEFT}),
        "td_num": ParagraphStyle("td_num", **{**base,
            "fontName": "Helvetica", "fontSize": 8, "alignment": TA_RIGHT}),
    }


# ─── Page Template ────────────────────────────────────────────────────────────

def _header_footer(canvas, doc):
    canvas.saveState()
    w, h = PAGE_W, PAGE_H

    # Header bar
    canvas.setFillColor(C_BG_DARK)
    canvas.rect(0, h - 22*mm, w, 22*mm, fill=1, stroke=0)

    # Gold accent line under header
    canvas.setFillColor(C_GOLD)
    canvas.rect(0, h - 22*mm, w, 0.6, fill=1, stroke=0)

    # Title
    canvas.setFillColor(C_WHITE)
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawString(MARGIN, h - 13*mm, "PORTFOLIO DASHBOARD ANALYSIS")

    # Subtitle
    canvas.setFillColor(C_GOLD)
    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(MARGIN, h - 17.5*mm, "RAPPORT DE PORTEFEUILLE")

    # Date (right side)
    date_str = datetime.now().strftime("%d %B %Y  —  %H:%M")
    canvas.setFillColor(C_TEXT_LT)
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(w - MARGIN, h - 13*mm, date_str)

    # Page number
    canvas.setFillColor(C_TEXT_SEC)
    canvas.setFont("Helvetica", 7)
    canvas.drawRightString(w - MARGIN, h - 18*mm, f"Page {doc.page}")

    # Footer line
    canvas.setStrokeColor(C_BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN, 10*mm, w - MARGIN, 10*mm)

    canvas.setFillColor(C_TEXT_SEC)
    canvas.setFont("Helvetica", 6.5)
    canvas.drawCentredString(w / 2, 7*mm,
        "Document généré automatiquement — Les données proviennent de Yahoo Finance (cours de clôture de la veille). Aucun conseil en investissement.")

    canvas.restoreState()


# ─── Helper: colored text ─────────────────────────────────────────────────────

def _signed_color(value: float, fmt: str, invert: bool = False) -> str:
    """Return HTML-colored string for a signed numeric value."""
    if value is None:
        return "—"
    pos_color = C_RED.hexval() if invert else C_GREEN.hexval()
    neg_color = C_GREEN.hexval() if invert else C_RED.hexval()
    color = pos_color if value >= 0 else neg_color
    sign = "+" if value > 0 else ""
    return f'<font color="{color}">{sign}{fmt.format(value)}</font>'


# ─── KPI Block ────────────────────────────────────────────────────────────────

def _kpi_table(summary: dict, styles: dict) -> Table:
    kpis = [
        ("VALORISATION TOTALE",  f"{summary.get('total_portfolio_value', 0):,.0f} €", None),
        ("VALEUR DES TITRES",    f"{summary.get('total_value', 0):,.0f} €", None),
        ("CASH DISPONIBLE",      f"{summary.get('cash_balance', 0):,.0f} €", None),
        ("CAPITAL VERSÉ",        f"{summary.get('total_invested', 0):,.0f} €", None),
        ("DIVIDENDES PERÇUS",    f"{summary.get('total_dividends', 0):,.0f} €", None),
        ("+/- VALUES LATENTES",  summary.get('unrealized_pnl', 0), "pnl"),
        ("PERFORMANCE TOTALE",   summary.get('total_return_pct', 0), "pct"),
    ]

    col_w = (PAGE_W - 2 * MARGIN) / len(kpis)
    labels, values, badges = [], [], []

    for label, val, typ in kpis:
        labels.append(Paragraph(label, styles["kpi_label"]))
        if typ == "pnl":
            pnl = val or 0
            pct = summary.get('unrealized_pnl_pct', 0) or 0
            values.append(Paragraph(_signed_color(pnl, "{:,.0f} €"), styles["kpi_value"]))
            badges.append(Paragraph(_signed_color(pct, "{:.2f} %"), styles["kpi_badge"]))
        elif typ == "pct":
            pct = val or 0
            values.append(Paragraph(_signed_color(pct, "{:.2f} %"), styles["kpi_value"]))
            badges.append(Paragraph("", styles["kpi_badge"]))
        else:
            values.append(Paragraph(str(val), styles["kpi_value"]))
            badges.append(Paragraph("", styles["kpi_badge"]))

    data = [labels, values, badges]
    col_widths = [col_w] * len(kpis)

    t = Table(data, colWidths=col_widths, rowHeights=[9*mm, 11*mm, 7*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), C_BG_LIGHT),
        ("BACKGROUND",   (0, 0), (-1, 0),  colors.HexColor("#e8eef6")),
        ("LINEAFTER",    (0, 0), (-2, -1), 0.5, C_BORDER),
        ("LINEBEFORE",   (0, 0), (0, -1),  0,   C_BORDER),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("ROUNDEDCORNERS", [3]),
        ("BOX",          (0, 0), (-1, -1), 0.5, C_BORDER),
        ("LINEBELOW",    (0, 0), (-1, 0),  0.5, C_BORDER),
    ]))
    return t


# ─── Positions Table ──────────────────────────────────────────────────────────

def _positions_table(positions: list, styles: dict) -> Table:
    headers = ["TICKER", "DESCRIPTION", "CATÉGORIE", "QTÉ", "PRU (€)", "COURS (€)",
               "VALEUR (€)", "P&L (€)", "P&L (%)", "POIDS (%)"]

    th = lambda s: Paragraph(s, styles["th"])
    td = lambda s: Paragraph(str(s), styles["td"])
    tl = lambda s: Paragraph(str(s), styles["td_left"])
    tr = lambda s: Paragraph(str(s), styles["td_num"])

    rows = [[th(h) for h in headers]]

    for pos in positions:
        pnl_eur = pos.get("P&L (€)", 0) or 0
        pnl_pct = pos.get("P&L (%)", 0) or 0
        cat = pos.get("Catégorie", "")
        rows.append([
            td(pos.get("Ticker", "")),
            tl(pos.get("Description", "")),
            td(cat),
            tr(f"{pos.get('Quantité', 0):,.0f}"),
            tr(f"{pos.get('PRU (€)', 0):,.2f}"),
            tr(f"{pos.get('Cours (€)', 0):,.2f}"),
            tr(f"{pos.get('Valeur (€)', 0):,.0f}"),
            Paragraph(_signed_color(pnl_eur, "{:,.0f}"), styles["td_num"]),
            Paragraph(_signed_color(pnl_pct, "{:.2f} %"), styles["td_num"]),
            tr(f"{pos.get('Poids (%)', 0):.1f}"),
        ])

    avail_w = PAGE_W - 2 * MARGIN
    col_widths = [
        avail_w * 0.07,  # ticker
        avail_w * 0.18,  # description
        avail_w * 0.07,  # cat
        avail_w * 0.05,  # qty
        avail_w * 0.07,  # pru
        avail_w * 0.07,  # cours
        avail_w * 0.09,  # valeur
        avail_w * 0.09,  # pnl eur
        avail_w * 0.08,  # pnl pct
        avail_w * 0.07,  # poids
    ]

    row_heights = [8*mm] + [7*mm] * (len(rows) - 1)

    style = TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  C_BG_DARK),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  C_GOLD),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, C_GOLD),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_BORDER),
        ("INNERGRID",     (0, 1), (-1, -1), 0.3, C_BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ])

    for i in range(1, len(rows)):
        if i % 2 == 0:
            style.add("BACKGROUND", (0, i), (-1, i), C_ROW_ALT)

    t = Table(rows, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)
    t.setStyle(style)
    return t


# ─── Risk Metrics Table ───────────────────────────────────────────────────────

def _metrics_table(analytics: dict, styles: dict) -> Table:
    th = lambda s: Paragraph(s, styles["th"])

    # Values from engine.py: ratios are decimals, % values already ×100
    def fv(val, is_pct=False, is_ratio=False):
        if val is None:
            return "—"
        if is_pct:
            return f"{val:.2f} %"
        if is_ratio:
            return f"{val:.3f}"
        return f"{val:.3f}"

    m = analytics
    rows_data = [
        ("RENDEMENT ANNUEL",   fv(m.get("ann_ret"),   is_pct=True),
         "VaR 95% (HIST.)",    fv(m.get("var_hist"),  is_pct=True)),
        ("VOLATILITÉ ANNUELLE",fv(m.get("ann_vol"),   is_pct=True),
         "CVaR 95% (HIST.)",   fv(m.get("cvar"),      is_pct=True)),
        ("SHARPE RATIO",       fv(m.get("sharpe"),    is_ratio=True),
         "VaR 95% (NORM.)",    fv(m.get("var_gauss"), is_pct=True)),
        ("SORTINO RATIO",      fv(m.get("sortino"),   is_ratio=True),
         "BETA (CAC 40)",      fv(m.get("beta"),      is_ratio=True)),
        ("MAX DRAWDOWN",       fv(m.get("max_dd"),    is_pct=True),
         "ALPHA (CAC 40)",     fv(m.get("alpha"),     is_pct=True)),
    ]

    header = [th("MÉTRIQUE"), th("VALEUR"), Paragraph("", styles["th"]),
              th("MÉTRIQUE"), th("VALEUR")]
    rows = [header]
    for lk, lv, rk, rv in rows_data:
        rows.append([
            Paragraph(lk, styles["td_left"]),
            Paragraph(lv, styles["td_num"]),
            Paragraph("", styles["td"]),
            Paragraph(rk, styles["td_left"]),
            Paragraph(rv, styles["td_num"]),
        ])

    avail_w = PAGE_W - 2 * MARGIN
    col_widths = [avail_w*0.22, avail_w*0.13, avail_w*0.03, avail_w*0.22, avail_w*0.13]
    # total = 0.73, pad with spacer col at end implicitly — make it sum to total
    # let's just use these 5 cols within a smaller frame by reducing widths
    col_widths = [80*mm, 40*mm, 8*mm, 80*mm, 40*mm]

    row_heights = [8*mm] + [7*mm] * (len(rows) - 1)

    t = Table(rows, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  C_BG_DARK),
        ("BACKGROUND",    (2, 0), (2, -1),  colors.HexColor("#f0f4f9")),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, C_GOLD),
        ("BOX",           (0, 0), (1, -1),  0.5, C_BORDER),
        ("BOX",           (3, 0), (4, -1),  0.5, C_BORDER),
        ("INNERGRID",     (0, 1), (1, -1),  0.3, C_BORDER),
        ("INNERGRID",     (3, 1), (4, -1),  0.3, C_BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        *[("BACKGROUND", (0, i), (1, i), C_ROW_ALT) for i in range(1, len(rows)) if i % 2 == 0],
        *[("BACKGROUND", (3, i), (4, i), C_ROW_ALT) for i in range(1, len(rows)) if i % 2 == 0],
    ]))
    return t


# ─── Per-Ticker Table ─────────────────────────────────────────────────────────

def _ticker_table(ticker_metrics: dict, styles: dict) -> Table:
    th = lambda s: Paragraph(s, styles["th"])

    headers = ["TICKER", "REND. MENS. (%)", "REND. ANNUEL (%)",
               "VOL. MENS. (%)", "VOL. ANNUELLE (%)", "REND/VOL"]
    rows = [[th(h) for h in headers]]

    avail_w = PAGE_W - 2 * MARGIN
    col_widths = [avail_w*0.12, avail_w*0.14, avail_w*0.14,
                  avail_w*0.14, avail_w*0.14, avail_w*0.12]

    for ticker, m in ticker_metrics.items():
        def fv(key):
            v = m.get(key)
            return f"{v:.2f}" if v is not None else "—"

        rows.append([
            Paragraph(ticker, styles["td"]),
            Paragraph(fv("monthly_ret"), styles["td_num"]),
            Paragraph(fv("yearly_ret"),  styles["td_num"]),
            Paragraph(fv("monthly_vol"), styles["td_num"]),
            Paragraph(fv("yearly_vol"),  styles["td_num"]),
            Paragraph(fv("ret_vol"),     styles["td_num"]),
        ])

    row_heights = [8*mm] + [7*mm] * (len(rows) - 1)

    t = Table(rows, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  C_BG_DARK),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, C_GOLD),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_BORDER),
        ("INNERGRID",     (0, 1), (-1, -1), 0.3, C_BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        *[("BACKGROUND", (0, i), (-1, i), C_ROW_ALT) for i in range(1, len(rows)) if i % 2 == 0],
    ]))
    return t


# ─── Pie Charts ──────────────────────────────────────────────────────────────

def _make_pie(labels, values, slice_colors, title, size=130):
    """Build a single reportlab Pie Drawing with legend."""
    d = Drawing(size + 80, size + 30)
    radius = size * 0.38

    # Title
    d.add(String(
        (size + 80) / 2, size + 18, title,
        fontName="Helvetica-Bold", fontSize=8,
        fillColor=C_TEXT_SEC, textAnchor="middle"
    ))

    pie = Pie()
    pie.x = 10
    pie.y = 10
    pie.width  = size * 0.76
    pie.height = size * 0.76
    pie.data   = values
    pie.labels = [""] * len(labels)  # labels drawn manually in legend

    pie.sideLabels         = False
    pie.simpleLabels       = True
    pie.startAngle         = 90
    pie.direction          = "clockwise"

    for i, c in enumerate(slice_colors):
        pie.slices[i].fillColor       = c
        pie.slices[i].strokeColor     = C_WHITE
        pie.slices[i].strokeWidth     = 0.8
        pie.slices[i].labelRadius     = 0
        pie.slices[i].popout          = 0

    d.add(pie)

    # Legend (right side)
    legend_x = size * 0.76 + 18
    for i, (lbl, val) in enumerate(zip(labels, values)):
        y = size * 0.76 - i * 14
        if y < 0:
            break
        d.add(Circle(legend_x, y + 3, 4, fillColor=slice_colors[i], strokeWidth=0))
        d.add(String(
            legend_x + 9, y,
            f"{lbl}  {val:.1f}%",
            fontName="Helvetica", fontSize=6.5,
            fillColor=C_TEXT, textAnchor="start"
        ))

    return d


def _pie_charts_row(positions: list, styles: dict) -> Table:
    """Returns a 1×2 table containing the two pie charts side by side."""
    avail_w = PAGE_W - 2 * MARGIN
    cell_w  = avail_w / 2

    # ── Pie 1 : Répartition par Titre ─────────────────────────────────────
    ticker_data = {}
    for pos in positions:
        t   = pos.get("Ticker", "?")
        pct = pos.get("Poids (%)", 0) or 0
        if pct > 0:
            ticker_data[t] = ticker_data.get(t, 0) + pct

    labels1  = list(ticker_data.keys())
    values1  = [ticker_data[k] for k in labels1]
    colors1  = [PIE_PALETTE[i % len(PIE_PALETTE)] for i in range(len(labels1))]
    pie1     = _make_pie(labels1, values1, colors1, "RÉPARTITION PAR TITRE", size=140)

    # ── Pie 2 : Allocation par Catégorie ──────────────────────────────────
    cat_data = {}
    for pos in positions:
        cat = pos.get("Catégorie", "Autre")
        pct = pos.get("Poids (%)", 0) or 0
        if pct > 0:
            cat_data[cat] = cat_data.get(cat, 0) + pct

    labels2  = list(cat_data.keys())
    values2  = [cat_data[k] for k in labels2]
    colors2  = [CAT_PALETTE.get(c, colors.HexColor("#8ba8c4")) for c in labels2]
    pie2     = _make_pie(labels2, values2, colors2, "ALLOCATION PAR CATÉGORIE", size=140)

    title1 = Paragraph("RÉPARTITION PAR TITRE", styles["section"])
    title2 = Paragraph("ALLOCATION PAR CATÉGORIE", styles["section"])

    data = [[pie1, pie2]]
    t = Table(data, colWidths=[cell_w, cell_w], rowHeights=[170])
    t.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("BACKGROUND",   (0, 0), (-1, -1), C_BG_LIGHT),
        ("BOX",          (0, 0), (-1, -1), 0.5, C_BORDER),
        ("LINEAFTER",    (0, 0), (0, -1),  0.5, C_BORDER),
    ]))
    return t


# ─── Main Build Function ──────────────────────────────────────────────────────

def build_pdf(summary: dict, positions: list, analytics: dict, ticker_metrics: dict) -> bytes:
    buf = io.BytesIO()
    styles = _styles()

    doc = BaseDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=26 * mm,
        bottomMargin=15 * mm,
    )

    frame = Frame(
        MARGIN, 15*mm,
        PAGE_W - 2*MARGIN, PAGE_H - 26*mm - 15*mm,
        leftPadding=0, rightPadding=0,
        topPadding=0, bottomPadding=0,
    )
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=_header_footer)])

    story = []
    S = styles

    # ── KPIs ──
    story.append(Paragraph("SYNTHÈSE DU PORTEFEUILLE", S["section"]))
    story.append(_kpi_table(summary, S))
    story.append(Spacer(1, 5*mm))

    # ── Positions ──
    story.append(Paragraph("POSITIONS ACTUELLES", S["section"]))
    story.append(_positions_table(positions, S))
    story.append(Spacer(1, 5*mm))

    # ── Pie Charts ──
    if positions:
        story.append(Paragraph("ALLOCATION DU PORTEFEUILLE", S["section"]))
        story.append(_pie_charts_row(positions, S))
        story.append(Spacer(1, 5*mm))

    # ── Métriques de risque ──
    if analytics:
        story.append(KeepTogether([
            Paragraph("MÉTRIQUES DE RISQUE", S["section"]),
            _metrics_table(analytics, S),
        ]))
        story.append(Spacer(1, 5*mm))

    # ── Par titre ──
    if ticker_metrics:
        story.append(KeepTogether([
            Paragraph("ANALYSE PAR TITRE", S["section"]),
            _ticker_table(ticker_metrics, S),
        ]))

    doc.build(story)
    buf.seek(0)
    return buf.read()
