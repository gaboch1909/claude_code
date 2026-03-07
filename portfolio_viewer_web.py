"""
Gaboch Portfolio Viewer — Web App (Streamlit)

Run:  streamlit run portfolio_viewer_web.py

Reads Gaboch_portfolio.xlsx (row-pair format) and displays active
stock transactions in a browser-based UI with theme customization.
"""
from __future__ import annotations

import contextlib
import io
import re
from datetime import datetime

import openpyxl
import streamlit as st
import yfinance as yf

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
SHEET_NAME    = "Gaboch_portfolio"
DEFAULT_EXCEL = r"C:\Gaboch_Portfolio\Gaboch_portfolio.xlsx"

CANADIAN_EXCHANGES = {"XTSE", "XTSX", "XCNQ", "TSX", "TSXV"}

THEME_PRESETS: dict[str, dict] = {
    "Dark Navy": {
        "bg": "#0d1b2a", "card_bg": "#1b2a3b", "accent": "#1565c0",
        "text": "#e8edf2", "profit": "#4caf50", "loss": "#ef5350",
        "header": "#64b5f6", "subhead": "#90caf9",
        "border": "#1e3a5f", "font_size": 15,
    },
    "Steel Blue": {
        "bg": "#0a1929", "card_bg": "#102a43", "accent": "#0288d1",
        "text": "#dce8f0", "profit": "#26a69a", "loss": "#ef5350",
        "header": "#4fc3f7", "subhead": "#81d4fa",
        "border": "#0d47a1", "font_size": 15,
    },
    "Forest Green": {
        "bg": "#0a1a0a", "card_bg": "#0d2b0d", "accent": "#388e3c",
        "text": "#e0ebe0", "profit": "#66bb6a", "loss": "#ef5350",
        "header": "#81c784", "subhead": "#a5d6a7",
        "border": "#1b5e20", "font_size": 15,
    },
    "Deep Purple": {
        "bg": "#1a0533", "card_bg": "#220a40", "accent": "#7b1fa2",
        "text": "#ede0f5", "profit": "#ba68c8", "loss": "#ef5350",
        "header": "#ce93d8", "subhead": "#e1bee7",
        "border": "#4a148c", "font_size": 15,
    },
    "Light Mode": {
        "bg": "#f4f6f9", "card_bg": "#ffffff", "accent": "#1565c0",
        "text": "#1a1a2e", "profit": "#2e7d32", "loss": "#c62828",
        "header": "#1565c0", "subhead": "#1976d2",
        "border": "#cbd5e1", "font_size": 15,
    },
}

DEFAULTS = THEME_PRESETS["Dark Navy"]

# ─────────────────────────────────────────────────────────────────────────────
# Data helpers
# ─────────────────────────────────────────────────────────────────────────────

def is_canadian(*fields) -> bool:
    """Return True if any field contains a Canadian exchange code, .TO/.V/.CN
    suffix, or an account name that mentions CAD."""
    for val in fields:
        s = str(val or "").strip().upper()
        for code in CANADIAN_EXCHANGES:
            if f"({code}:" in s or f"({code} " in s:
                return True
        if s.endswith(".TO") or s.endswith(".V") or s.endswith(".CN"):
            return True
        if " CAD " in f" {s} ":
            return True
    return False


def extract_company_name(ticker_full: str) -> str:
    m = re.match(r'^(.+?)\s*\(', str(ticker_full).strip())
    return m.group(1).strip() if m else str(ticker_full).strip()


def fmt_currency(value, canadian: bool, show_sign: bool = False) -> str:
    sym = "CAD$" if canadian else "USD$"
    if value is None:
        return "N/A"
    try:
        v = float(value)
        if show_sign:
            return f"+{sym} {v:,.2f}" if v >= 0 else f"-{sym} {abs(v):,.2f}"
        return f"{sym} {v:,.2f}"
    except (TypeError, ValueError):
        return f"{sym} {value}"


def fmt_pct(value) -> str:
    if value is None:
        return "N/A"
    try:
        v = float(value)
        return f"{'+' if v >= 0 else ''}{v:.2f}%"
    except (TypeError, ValueError):
        return str(value)


def fmt_shares(value) -> str:
    if value is None:
        return "N/A"
    try:
        v = float(value)
        return f"{int(v):,}" if v == int(v) else f"{v:,.4f}"
    except (TypeError, ValueError):
        return str(value)


def fmt_date(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, str) and not value.strip():
        return "—"
    if isinstance(value, datetime):
        return value.strftime(f"%A, %B {value.day}, %Y")
    return str(value)


# ── Yahoo Finance helpers ─────────────────────────────────────────────────────

_TICKER_RE = re.compile(
    r'^[A-Z0-9]{1,6}(\.(TO|V|CN))?$'
    r'|^[A-Z0-9]{1,5}-(USD|CAD|GBP|EUR|BTC)$'
)


def extract_yf_ticker(ticker_full: str) -> str:
    """Extract a yfinance-compatible symbol from 'NAME (EXCHANGE:TICKER)'."""
    s = str(ticker_full).strip()
    if _TICKER_RE.match(s.upper()):
        return s
    m = re.search(r'\(([^:)]+):([^)]+)\)', s)
    if not m:
        return ""
    exchange = m.group(1).strip().upper()
    symbol   = m.group(2).strip()
    if exchange in {"XTSE", "TSX"}:
        return symbol + ".TO"
    if exchange in {"XTSX", "TSXV"}:
        return symbol + ".V"
    if exchange in {"XCNQ"}:
        return symbol + ".CN"
    return symbol


@st.cache_data(ttl=60)
def fetch_yf_info(yf_ticker: str) -> dict:
    """Return company name, live price, and currency from Yahoo Finance (cached 60s)."""
    if not yf_ticker:
        return {}
    try:
        with contextlib.redirect_stderr(io.StringIO()), \
             contextlib.redirect_stdout(io.StringIO()):
            info = yf.Ticker(yf_ticker).info
        return {
            "company_name":  info.get("longName") or info.get("shortName") or "",
            "current_price": (info.get("currentPrice")
                              or info.get("regularMarketPrice")
                              or info.get("previousClose")),
            "is_canadian":   info.get("currency", "USD").upper() == "CAD",
        }
    except Exception:
        return {"company_name": "", "current_price": None, "is_canadian": False}


def _parse_rows(rows: list) -> dict:
    """Core row-pair parser shared by both load functions."""
    portfolio: dict = {}
    i = 0
    while i + 1 < len(rows):
        r1 = rows[i]      # purchase / even Excel row
        r2 = rows[i + 1]  # current  / odd  Excel row
        i += 2

        if all(c is None for c in r1):
            continue
        if r1[0] != 1:
            continue

        # Columns C and D carry the full ticker name; use whichever is not an error
        c2 = str(r1[2] or "")
        d2 = str(r1[3] or "")
        c3 = str(r2[2] or "")
        d3 = str(r2[3] or "")
        ticker_full = next((v for v in [c2, d2] if v and not v.startswith("#")), c2)
        ticker      = next((v for v in [c3, d3, c2, d2] if v and not v.startswith("#")), "").strip()
        if not ticker:
            continue

        txn = {
            "account":        r1[1],   # B
            "current_price":  r1[4],   # E
            "purchase_price": r1[5],   # F (even row = purchase price)
            "shares":         r1[6],   # G
            "subtotal":       r1[7],   # H
            "action":         r1[8],   # I
            "profits_usd":    r1[9],   # J
            "profits_pct":    r1[10],  # K
            "purchase_date":  r2[11],  # L — dates live in the odd (simple-ticker) row
            "selling_date":   r2[12],  # M
        }

        if ticker not in portfolio:
            account_name = str(r1[1] or "")
            is_cad = is_canadian(c2, d2, c3, d3, account_name)
            yf_tk  = extract_yf_ticker(ticker_full)
            # If extraction failed but stock is Canadian, build from first word + .TO
            if not yf_tk and is_cad:
                base = ticker.split()[0] if ticker else ""
                if base and _TICKER_RE.match((base + ".TO").upper()):
                    yf_tk = base + ".TO"
            portfolio[ticker] = {
                "company_name": extract_company_name(ticker_full),
                "ticker_full":  ticker_full,
                "yf_ticker":    yf_tk,
                "is_canadian":  is_cad,
                "transactions": [],
            }
        portfolio[ticker]["transactions"].append(txn)

    return portfolio


@st.cache_data(ttl=60)
def load_portfolio(filepath: str) -> dict:
    """Load from a local file path (laptop). Cached 60 s."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    return _parse_rows(list(wb[SHEET_NAME].iter_rows(min_row=2, values_only=True)))


@st.cache_data(ttl=60)
def load_portfolio_bytes(data: bytes) -> dict:
    """Load from uploaded bytes (Streamlit Cloud). Cached 60 s."""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    return _parse_rows(list(wb[SHEET_NAME].iter_rows(min_row=2, values_only=True)))


# ─────────────────────────────────────────────────────────────────────────────
# Theme / CSS
# ─────────────────────────────────────────────────────────────────────────────

def _init_theme() -> None:
    if "theme" not in st.session_state:
        st.session_state.theme = dict(DEFAULTS)


def _get_theme() -> dict:
    return st.session_state.theme


def _inject_css(t: dict) -> None:
    fs = t["font_size"]
    st.markdown(f"""
    <style>
    /* App background */
    .stApp {{ background-color: {t['bg']} !important; }}
    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background-color: {t['card_bg']} !important;
    }}
    section[data-testid="stSidebar"] * {{ color: {t['text']} !important; }}
    /* Report container */
    .rpt-container {{
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: {fs}px;
        color: {t['text']};
        max-width: 860px;
    }}
    .rpt-ticker {{
        color: {t['header']};
        font-size: {fs + 14}px;
        font-weight: 700;
        letter-spacing: 2px;
        margin-bottom: 2px;
    }}
    .rpt-company {{
        color: {t['subhead']};
        font-size: {fs + 2}px;
        font-weight: 600;
        margin-bottom: 10px;
    }}
    .rpt-divider {{
        border: none;
        border-top: 2px solid {t['accent']};
        margin: 10px 0 18px 0;
    }}
    .rpt-thin-divider {{
        border: none;
        border-top: 1px solid {t['border']};
        margin: 4px 0 10px 0;
    }}
    .rpt-total-box {{
        margin: 0 0 22px 0;
        padding: 18px 24px;
        border: 2px solid {t['accent']};
        border-radius: 10px;
        background: {t['card_bg']};
        display: flex;
        align-items: center;
        gap: 16px;
    }}
    .rpt-total-lbl {{
        color: {t['text']};
        font-size: {fs + 1}px;
        font-weight: 700;
    }}
    .rpt-total-profit {{
        color: {t['profit']};
        font-size: {fs + 5}px;
        font-weight: 800;
    }}
    .rpt-total-loss {{
        color: {t['loss']};
        font-size: {fs + 5}px;
        font-weight: 800;
    }}
    .rpt-txn-hdr {{
        color: {t['subhead']};
        font-size: {fs + 1}px;
        font-weight: 700;
        margin: 18px 0 8px 0;
        padding-bottom: 4px;
        border-bottom: 1px solid {t['border']};
    }}
    .rpt-row {{
        display: flex;
        padding: 5px 0;
        border-bottom: 1px solid rgba(255,255,255,0.06);
        align-items: center;
    }}
    .rpt-lbl {{
        color: {t['text']};
        width: 220px;
        opacity: 0.8;
        font-size: {fs}px;
    }}
    .rpt-val {{
        color: {t['text']};
        font-weight: 600;
        font-size: {fs}px;
    }}
    .rpt-profit {{
        color: {t['profit']};
        font-weight: 700;
        font-size: {fs}px;
    }}
    .rpt-loss {{
        color: {t['loss']};
        font-weight: 700;
        font-size: {fs}px;
    }}
    /* Streamlit widget overrides */
    div[data-testid="stSelectbox"] label,
    div[data-testid="stTextInput"] label {{
        color: {t['subhead']} !important;
        font-weight: 600;
    }}
    div[data-testid="metric-container"] {{
        background: {t['card_bg']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 10px 14px;
    }}
    div[data-testid="metric-container"] label {{
        color: {t['subhead']} !important;
    }}
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {{
        color: {t['header']} !important;
        font-weight: 700;
    }}
    </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Report HTML builder
# ─────────────────────────────────────────────────────────────────────────────

def _build_report_html(portfolio: dict, ticker: str, live_price, yf_company: str) -> str:
    data         = portfolio[ticker]
    transactions = data["transactions"]
    is_cad       = data["is_canadian"]

    # Strip Excel formula errors from company name
    excel_company = data["company_name"]
    if not excel_company or str(excel_company).strip().startswith("#"):
        excel_company = ""
    company = yf_company or excel_company or ticker

    def field_row(label: str, value: str, css_class: str = "rpt-val") -> str:
        return (
            f'<div class="rpt-row">'
            f'<span class="rpt-lbl">{label}</span>'
            f'<span class="{css_class}">{value}</span>'
            f'</div>'
        )

    # Pre-calculate total P&L
    total_profit = 0.0
    pnl_per_txn  = []
    for txn in transactions:
        try:
            cur  = live_price if live_price is not None else float(txn["current_price"])
            pur  = float(txn["purchase_price"])
            shs  = float(txn["shares"])
            pf   = (cur - pur) * shs
            pp   = ((cur - pur) / pur * 100) if pur != 0 else 0.0
            total_profit += pf
        except (TypeError, ValueError):
            pf, pp, cur = None, None, None
        pnl_per_txn.append((pf, pp, cur))

    html = ['<div class="rpt-container">']
    html.append(f'<div class="rpt-ticker">{ticker}</div>')
    if company:
        html.append(f'<div class="rpt-company">{company}</div>')
    html.append('<hr class="rpt-divider">')

    # Total box at top
    total_class = "rpt-total-profit" if total_profit >= 0 else "rpt-total-loss"
    total_str   = fmt_currency(total_profit, is_cad, show_sign=True)
    html.append(
        f'<div class="rpt-total-box">'
        f'<span class="rpt-total-lbl">TOTAL PROFIT / LOSS:</span>'
        f'<span class="{total_class}">{total_str}</span>'
        f'</div>'
    )

    for idx, (txn, (p_float, p_pct, cur_used)) in enumerate(
        zip(transactions, pnl_per_txn), 1
    ):
        acct     = txn.get("account") or ""
        acct_str = (
            f" &nbsp;·&nbsp; <span style='opacity:0.7;font-weight:400'>{acct}</span>"
            if acct else ""
        )
        html.append(f'<div class="rpt-txn-hdr">Transaction #{idx}{acct_str}</div>')

        cur_display = cur_used if cur_used is not None else txn["current_price"]
        html.append(field_row("Current Price:",    fmt_currency(cur_display,          is_cad)))
        html.append(field_row("Purchase Price:",   fmt_currency(txn["purchase_price"], is_cad)))
        html.append(field_row("Amount of Shares:", fmt_shares(txn["shares"])))

        pnl_class = "rpt-profit" if (p_float or 0) >= 0 else "rpt-loss"
        pnl_str   = fmt_currency(p_float, is_cad, show_sign=True)
        pct_str   = fmt_pct(p_pct)
        html.append(field_row(
            "Profit / Loss:",
            f"{pnl_str} &nbsp;<span style='opacity:0.75'>({pct_str})</span>",
            pnl_class,
        ))

        html.append(field_row("Purchase Date:", fmt_date(txn["purchase_date"])))
        html.append(field_row("Selling Date:",  fmt_date(txn["selling_date"])))

    html.append('</div>')
    return "\n".join(html)


# ─────────────────────────────────────────────────────────────────────────────
# Streamlit App
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    st.set_page_config(
        page_title="Gaboch Portfolio",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    _init_theme()
    theme = _get_theme()
    _inject_css(theme)

    # ── Determine data source: local file (laptop) or upload (cloud) ─────────
    import os as _os
    excel_path   = st.session_state.get("excel_path", DEFAULT_EXCEL)
    local_exists = _os.path.exists(excel_path)
    load_error   = None
    portfolio: dict = {}

    if local_exists:
        try:
            portfolio = load_portfolio(excel_path)
        except Exception as exc:
            load_error = str(exc)
    else:
        # Cloud deployment — use bytes stored in session_state
        if "excel_bytes" in st.session_state:
            try:
                portfolio = load_portfolio_bytes(st.session_state["excel_bytes"])
            except Exception as exc:
                load_error = str(exc)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(
            f"<h2 style='color:{theme['header']};margin-bottom:0;padding-bottom:0'>"
            f"📊 Gaboch Portfolio</h2>",
            unsafe_allow_html=True,
        )
        st.markdown("---")

        if local_exists:
            # Laptop: show path input
            new_path = st.text_input("📁 Excel File Path:", value=excel_path)
            if new_path != excel_path:
                st.session_state.excel_path = new_path
                load_portfolio.clear()
                st.rerun()
        else:
            # Cloud: show file uploader
            st.markdown(f"<span style='color:{theme['subhead']};font-weight:600'>"
                        f"📁 Upload your Excel file:</span>", unsafe_allow_html=True)
            uploaded = st.file_uploader(
                "Gaboch_portfolio.xlsx", type=["xlsx", "xlsm"],
                label_visibility="collapsed",
            )
            if uploaded is not None:
                st.session_state["excel_bytes"] = uploaded.read()
                load_portfolio_bytes.clear()
                st.rerun()

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                load_portfolio.clear()
                load_portfolio_bytes.clear()
                fetch_yf_info.clear()
                st.rerun()

        if load_error:
            st.error(f"⚠️ Load error:\n{load_error}")

        st.markdown("---")

        # Ticker selector
        tickers = sorted(portfolio.keys()) if portfolio else []
        if tickers:
            selected_ticker: str | None = st.selectbox(
                "📈 Select Ticker:", tickers,
                help="Only active (not sold) tickers are listed.",
            )
        else:
            selected_ticker = None
            if not load_error:
                st.info("No active tickers found.")

        st.markdown("---")

        # ── Customize Layout ──────────────────────────────────────────────────
        with st.expander("⚙️  Customize Layout", expanded=False):
            preset_name = st.selectbox(
                "Theme Preset:", list(THEME_PRESETS.keys()), label_visibility="visible",
            )
            if st.button("✔  Apply Preset", use_container_width=True):
                st.session_state.theme = dict(THEME_PRESETS[preset_name])
                st.rerun()

            st.markdown("**Custom Colors:**")
            c1, c2 = st.columns(2)
            with c1:
                profit_c = st.color_picker("Profit",     theme["profit"],  key="cp_profit")
                header_c = st.color_picker("Header",     theme["header"],  key="cp_header")
                accent_c = st.color_picker("Accent",     theme["accent"],  key="cp_accent")
            with c2:
                loss_c   = st.color_picker("Loss",       theme["loss"],    key="cp_loss")
                subhd_c  = st.color_picker("Sub-header", theme["subhead"], key="cp_sub")
                text_c   = st.color_picker("Text",       theme["text"],    key="cp_text")

            font_size = st.slider(
                "Font Size (px)", min_value=12, max_value=22,
                value=theme["font_size"], step=1,
            )

            bc1, bc2 = st.columns(2)
            with bc1:
                if st.button("✔  Apply Colors", use_container_width=True):
                    st.session_state.theme.update({
                        "profit": profit_c, "loss": loss_c,
                        "header": header_c, "subhead": subhd_c,
                        "accent": accent_c, "text": text_c,
                        "font_size": font_size,
                    })
                    st.rerun()
            with bc2:
                if st.button("↺  Defaults", use_container_width=True):
                    st.session_state.theme = dict(DEFAULTS)
                    st.rerun()

    # ── Main content ──────────────────────────────────────────────────────────
    if load_error:
        st.error(f"**Failed to load portfolio file:** `{excel_path}`")
        st.code(load_error)
        st.info("Update the Excel file path in the sidebar.")
        return

    if not selected_ticker or not portfolio:
        st.markdown(
            f"<div style='color:{theme['subhead']};font-size:18px;margin-top:40px;text-align:center'>"
            f"📈 Select a ticker from the sidebar to view transactions."
            f"</div>",
            unsafe_allow_html=True,
        )
        return

    data = portfolio[selected_ticker]

    # ── Live Yahoo Finance data ───────────────────────────────────────────────
    yf_ticker = data.get("yf_ticker", "") or ""
    if not yf_ticker:
        base = selected_ticker.split()[0] if " " in selected_ticker else selected_ticker
        if data.get("is_canadian") and not base.upper().endswith(".TO"):
            yf_ticker = base + ".TO"
        else:
            yf_ticker = base

    yf_data    = fetch_yf_info(yf_ticker)
    yf_company = yf_data.get("company_name", "")
    live_price = yf_data.get("current_price")
    is_cad     = data["is_canadian"] or yf_data.get("is_canadian", False)

    txns = data["transactions"]
    sym  = "CAD$" if is_cad else "USD$"

    # Company name with #VALUE! filter
    excel_company = data["company_name"]
    if not excel_company or str(excel_company).strip().startswith("#"):
        excel_company = ""
    company = yf_company or excel_company or selected_ticker

    # ── Summary metrics ───────────────────────────────────────────────────────
    total_pnl = 0.0
    for txn in txns:
        try:
            cur = live_price if live_price is not None else float(txn["current_price"])
            pur = float(txn["purchase_price"])
            shs = float(txn["shares"])
            total_pnl += (cur - pur) * shs
        except (TypeError, ValueError):
            pass

    total_str = fmt_currency(total_pnl, is_cad, show_sign=True)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ticker",       selected_ticker)
    col2.metric("Company",      company[:28] + "…" if len(company) > 28 else company)
    col3.metric("Transactions", len(txns))
    col4.metric("Currency",     sym)

    st.markdown("<div style='margin-top:6px'></div>", unsafe_allow_html=True)

    # ── Transaction report ────────────────────────────────────────────────────
    html = _build_report_html(portfolio, selected_ticker, live_price, yf_company)
    st.markdown(html, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
