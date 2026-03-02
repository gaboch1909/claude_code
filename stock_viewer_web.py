import streamlit as st
import pandas as pd
import requests
import io
import math
import base64
import time
import yfinance as yf
from datetime import datetime

st.set_page_config(
    page_title="Stock Viewer",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Field definitions (matches desktop app) ───────────────────────
FIELD_GROUPS = {
    "Company Info":  ["Company Name", "Company Description"],
    "Market Data":   ["Mkt Cap/Net Assets (M)", "Price", "Prev Close",
                      "Change %", "Day High", "Day Low", "EPS (TTM)"],
    "Revenue":       ["Revenue TTM (M)", "Revenue 1Y (M)",
                      "Revenue 3Y (M)", "Revenue 5Y (M)"],
    "Net Income":    ["Net Income TTM (M)", "Net Income 1Y (M)",
                      "Net Income 3Y (M)", "Net Income 5Y (M)"],
    "Balance Sheet": ["Total Debt (M)", "Cash (M)"],
    "Other":         ["Last Updated"],
}

PRICE_FIELDS   = {"Price", "Prev Close", "Day High", "Day Low", "EPS (TTM)"}
MILLION_FIELDS = {
    "Mkt Cap/Net Assets (M)",
    "Revenue TTM (M)", "Revenue 1Y (M)", "Revenue 3Y (M)", "Revenue 5Y (M)",
    "Net Income TTM (M)", "Net Income 1Y (M)", "Net Income 3Y (M)", "Net Income 5Y (M)",
    "Total Debt (M)", "Cash (M)",
}

# Top metric cards (shown as big cards at the top of the report)
TOP_METRICS = ["Price", "Change %", "Mkt Cap/Net Assets (M)", "Day High", "Day Low"]

# All fields in display order
ALL_FIELDS = [f for fields in FIELD_GROUPS.values() for f in fields]


# ── Theme ─────────────────────────────────────────────────────────
THEME_DEFAULTS = {
    "accent":  "#1F4E79",
    "pos":     "#2e7d32",
    "neg":     "#c62828",
    "na":      "#c62828",
    "font_sz": 16,
}

THEME_PRESETS = {
    "Dark Navy (default)": {"accent": "#1F4E79", "pos": "#2e7d32", "neg": "#c62828", "na": "#c62828", "font_sz": 16},
    "Steel Blue":          {"accent": "#1565c0", "pos": "#2e7d32", "neg": "#b71c1c", "na": "#b71c1c", "font_sz": 16},
    "Forest Green":        {"accent": "#2e7d32", "pos": "#1565c0", "neg": "#c62828", "na": "#c62828", "font_sz": 16},
    "Deep Purple":         {"accent": "#4a148c", "pos": "#00695c", "neg": "#c62828", "na": "#c62828", "font_sz": 16},
    "Charcoal":            {"accent": "#37474f", "pos": "#388e3c", "neg": "#d32f2f", "na": "#d32f2f", "font_sz": 16},
}

# Maps theme keys → color picker widget keys
_CP_KEYS = {"accent": "cp_accent", "pos": "cp_pos", "neg": "cp_neg", "na": "cp_na"}


def _init_theme():
    """Seed session state with defaults on first run."""
    for k, v in THEME_DEFAULTS.items():
        if f"theme_{k}" not in st.session_state:
            st.session_state[f"theme_{k}"] = v
        if k in _CP_KEYS and _CP_KEYS[k] not in st.session_state:
            st.session_state[_CP_KEYS[k]] = v
    if "sl_font" not in st.session_state:
        st.session_state["sl_font"] = THEME_DEFAULTS["font_sz"]


def _t(key):
    return st.session_state.get(f"theme_{key}", THEME_DEFAULTS[key])


def _apply_preset(preset_dict):
    """Apply a preset or defaults dict, updating both theme_ and widget keys."""
    for k, v in preset_dict.items():
        st.session_state[f"theme_{k}"] = v
        if k in _CP_KEYS:
            st.session_state[_CP_KEYS[k]] = v
        if k == "font_sz":
            st.session_state["sl_font"] = v


# ── Helpers ───────────────────────────────────────────────────────
def is_na(raw):
    if raw is None:
        return True
    if isinstance(raw, float) and math.isnan(raw):
        return True
    return str(raw).strip().lower() in ("", "n/a", "none", "nan", "nat")


def fmt_value(field, raw):
    """Return (display_string, css_class). css_class is 'na', 'pos', 'neg', or ''."""
    if is_na(raw):
        return "N/A", "na"
    try:
        v = float(raw)
        if field == "Change %":
            s = f"{'+' if v >= 0 else ''}{v:.2f}%"
            return s, "pos" if v >= 0 else "neg"
        if field in MILLION_FIELDS:
            return f"${v:,.2f} M", ""
        if field in PRICE_FIELDS:
            return f"${v:,.2f}", ""
    except (ValueError, TypeError):
        pass
    return str(raw), ""


# ── yfinance fetch ────────────────────────────────────────────────
def _safe_m(v):
    try:
        f = float(v)
        return "N/A" if math.isnan(f) else round(f / 1e6, 2)
    except Exception:
        return "N/A"


def _safe_v(v):
    try:
        f = float(v)
        return "N/A" if math.isnan(f) else round(f, 4)
    except Exception:
        return "N/A"


def _get_annual(fin, row, idx):
    try:
        if fin is not None and not fin.empty and row in fin.index and fin.shape[1] > idx:
            return _safe_m(fin.iloc[:, idx][row])
    except Exception:
        pass
    return "N/A"


def _build_desc(info):
    d = info.get("longBusinessSummary")
    if d:
        return d
    qt, parts = info.get("quoteType", ""), []
    if qt in ("ETF", "MUTUALFUND"):
        parts.append(f"{info.get('longName', '')} is an {qt}.")
        if info.get("category"):
            parts.append(f"Category: {info['category']}.")
        if info.get("fundFamily"):
            parts.append(f"Fund family: {info['fundFamily']}.")
    return " ".join(parts) if parts else "N/A"


def fetch_ticker(symbol, retries=3, base_delay=5):
    """Fetch live data for a ticker from yfinance. Returns (data_dict, error_str)."""
    for attempt in range(retries):
        try:
            t = yf.Ticker(symbol)
            info = t.info
            if not info or (info.get("currentPrice") is None
                            and info.get("regularMarketPrice") is None
                            and info.get("previousClose") is None):
                return None, f"No data found for '{symbol}'. Check the ticker symbol."

            price = _safe_v(info.get("currentPrice") or info.get("regularMarketPrice"))
            prev  = _safe_v(info.get("previousClose") or info.get("regularMarketPreviousClose"))
            try:
                chg = round(((float(price) - float(prev)) / float(prev)) * 100, 2) \
                      if price != "N/A" and prev != "N/A" else "N/A"
            except Exception:
                chg = "N/A"

            try:
                fin = t.financials
            except Exception:
                fin = None

            return {
                "Stock Ticker":           symbol.upper(),
                "Company Name":           info.get("longName", "N/A"),
                "Company Description":    _build_desc(info),
                "Mkt Cap/Net Assets (M)": _safe_m(info.get("marketCap") or info.get("totalAssets")),
                "Price":                  price,
                "Prev Close":             prev,
                "Change %":               chg,
                "Day High":               _safe_v(info.get("dayHigh") or info.get("regularMarketDayHigh")),
                "Day Low":                _safe_v(info.get("dayLow") or info.get("regularMarketDayLow")),
                "EPS (TTM)":              _safe_v(info.get("trailingEps")),
                "Revenue TTM (M)":        _safe_m(info.get("totalRevenue")),
                "Revenue 1Y (M)":         _get_annual(fin, "Total Revenue", 0),
                "Revenue 3Y (M)":         _get_annual(fin, "Total Revenue", 2),
                "Revenue 5Y (M)":         _get_annual(fin, "Total Revenue", 4),
                "Net Income TTM (M)":     _safe_m(info.get("netIncomeToCommon")),
                "Net Income 1Y (M)":      _get_annual(fin, "Net Income", 0),
                "Net Income 3Y (M)":      _get_annual(fin, "Net Income", 2),
                "Net Income 5Y (M)":      _get_annual(fin, "Net Income", 4),
                "Total Debt (M)":         _safe_m(info.get("totalDebt")),
                "Cash (M)":               _safe_m(info.get("totalCash")),
                "Last Updated":           datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }, None

        except Exception as e:
            err_str = str(e)
            if any(k in err_str.lower() for k in ("too many requests", "rate limit", "429")):
                time.sleep(base_delay * (2 ** attempt))
            else:
                return None, f"Error: {e}"
    return None, "Rate limited. Wait a moment and try again."


# ── GitHub write-back ─────────────────────────────────────────────
def push_df_to_github(df):
    """Save the DataFrame as stocks.xlsx and push it back to GitHub."""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        owner = st.secrets["GITHUB_OWNER"]
        repo  = st.secrets["GITHUB_REPO"]
    except KeyError as e:
        return False, f"Missing secret: {e}"

    url = f"https://api.github.com/repos/{owner}/{repo}/contents/stocks.xlsx"
    headers = {
        "Authorization":        f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    r = requests.get(url, headers=headers, timeout=15)
    if r.status_code != 200:
        return False, f"GitHub read failed ({r.status_code}): {r.text[:200]}"
    sha = r.json().get("sha", "")

    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    content = base64.b64encode(buf.getvalue()).decode()

    data = {"message": "update stocks", "content": content, "sha": sha}
    r = requests.put(url, headers=headers, json=data, timeout=30)
    if r.status_code in (200, 201):
        return True, None
    if r.status_code == 403:
        return False, (
            "**403 Permission Denied** — your GitHub token is read-only.\n\n"
            "Fix: Go to **GitHub → Settings → Developer settings → "
            "Personal access tokens → Tokens (classic)**, create a new token "
            "with the **`repo`** scope, then update **GITHUB_TOKEN** in "
            "Streamlit → Manage app → Settings → Secrets."
        )
    return False, f"GitHub push failed ({r.status_code}): {r.text[:200]}"


# ── Data loading ──────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    """Fetch stocks.xlsx from the GitHub repo using stored secrets."""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        owner = st.secrets["GITHUB_OWNER"]
        repo  = st.secrets["GITHUB_REPO"]
    except KeyError as e:
        return None, f"Missing Streamlit secret: {e}. Add it in App Settings → Secrets."

    url = f"https://api.github.com/repos/{owner}/{repo}/contents/stocks.xlsx"
    headers = {
        "Authorization":        f"Bearer {token}",
        "Accept":               "application/vnd.github.raw+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    try:
        r = requests.get(url, headers=headers, timeout=20)
        if r.status_code == 200:
            return pd.read_excel(io.BytesIO(r.content)), None
        return None, f"GitHub API error {r.status_code}: {r.text[:300]}"
    except Exception as e:
        return None, f"Connection error: {e}"


# ── App ───────────────────────────────────────────────────────────
def main():
    _init_theme()

    # Read theme values — these were written by widgets on the previous rerun
    accent  = _t("accent")
    pos_col = _t("pos")
    neg_col = _t("neg")
    na_col  = _t("na")
    font_sz = _t("font_sz")

    # Inject CSS — uses current theme values
    st.markdown(f"""
    <style>
    .na   {{ color: {na_col}  !important; font-style: italic; }}
    .pos  {{ color: {pos_col} !important; font-weight: 700; }}
    .neg  {{ color: {neg_col} !important; font-weight: 700; }}
    .section-hdr {{
        background: {accent} !important;
        color: #ffffff !important;
        padding: 6px 14px;
        border-radius: 5px;
        font-size: {font_sz - 1}px;
        font-weight: 700;
        margin: 16px 0 8px 0;
        letter-spacing: 0.04em;
    }}
    div[data-testid="metric-container"] {{ padding: 8px 10px !important; }}
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────
    with st.sidebar:
        st.title("📈 Stock Viewer")

        if st.button("🔄 Refresh Data", use_container_width=True,
                     help="Pull the latest stocks.xlsx from GitHub"):
            load_data.clear()
            st.rerun()
        st.caption("Data auto-refreshes every 5 min. Press Refresh for immediate update.")
        st.divider()

        with st.spinner("Loading data..."):
            df, load_err = load_data()

        if load_err:
            st.error(load_err)
            st.stop()
        if df is None or df.empty:
            st.warning("stocks.xlsx is empty or missing from the repository.")
            st.stop()

        tickers = sorted(df["Stock Ticker"].dropna().astype(str).unique())
        selected = st.selectbox("Ticker", tickers)

        st.divider()

        # ── Add New Ticker ─────────────────────────────────────────
        with st.expander("➕ Add New Ticker"):
            new_sym = st.text_input(
                "Ticker symbol (e.g. AAPL, TD.TO)", key="new_sym"
            ).strip().upper()
            if st.button("Fetch & Add", use_container_width=True, key="btn_add"):
                if not new_sym:
                    st.warning("Enter a ticker symbol.")
                else:
                    with st.spinner(f"Fetching {new_sym}…"):
                        data, fetch_err = fetch_ticker(new_sym)
                    if fetch_err:
                        st.error(fetch_err)
                    else:
                        with st.spinner("Saving to GitHub…"):
                            updated = df[df["Stock Ticker"] != new_sym].copy()
                            updated = pd.concat(
                                [updated, pd.DataFrame([data])], ignore_index=True
                            ).sort_values("Stock Ticker").reset_index(drop=True)
                            ok, push_err = push_df_to_github(updated)
                        if ok:
                            st.success(f"✓ {new_sym} added!")
                            load_data.clear()
                            st.rerun()
                        else:
                            st.error(push_err)

        # ── Customize Layout ───────────────────────────────────────
        with st.expander("🎨 Customize Layout"):
            preset_names = list(THEME_PRESETS.keys())
            preset = st.selectbox("Theme preset", preset_names, key="preset_sel")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Apply", use_container_width=True, key="btn_preset"):
                    _apply_preset(THEME_PRESETS[preset])
                    st.rerun()
            with col2:
                if st.button("↺ Defaults", use_container_width=True, key="btn_reset_theme"):
                    _apply_preset(THEME_DEFAULTS)
                    st.rerun()

            st.markdown("**Colors**")

            new_accent = st.color_picker("Section headers", _t("accent"), key="cp_accent")
            if new_accent != _t("accent"):
                st.session_state["theme_accent"] = new_accent
                st.rerun()

            new_pos = st.color_picker("Positive values", _t("pos"), key="cp_pos")
            if new_pos != _t("pos"):
                st.session_state["theme_pos"] = new_pos
                st.rerun()

            new_neg = st.color_picker("Negative values", _t("neg"), key="cp_neg")
            if new_neg != _t("neg"):
                st.session_state["theme_neg"] = new_neg
                st.rerun()

            new_na = st.color_picker("N/A values", _t("na"), key="cp_na")
            if new_na != _t("na"):
                st.session_state["theme_na"] = new_na
                st.rerun()

            new_sz = st.slider("Font size", 12, 22, _t("font_sz"), key="sl_font")
            if new_sz != _t("font_sz"):
                st.session_state["theme_font_sz"] = new_sz
                st.rerun()

        st.divider()

        # ── Fields to display ──────────────────────────────────────
        st.markdown("**Fields to display**")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Select All", use_container_width=True, key="btn_sel_all"):
                for f in ALL_FIELDS:
                    st.session_state[f"chk_{f}"] = True
        with c2:
            if st.button("Clear All", use_container_width=True, key="btn_clr_all"):
                for f in ALL_FIELDS:
                    st.session_state[f"chk_{f}"] = False

        chosen = []
        for group, fields in FIELD_GROUPS.items():
            with st.expander(group, expanded=True):
                for field in fields:
                    if st.checkbox(field, value=True, key=f"chk_{field}"):
                        chosen.append(field)

    # ── Main panel ─────────────────────────────────────────────────
    row_df = df[df["Stock Ticker"] == selected]
    if row_df.empty:
        st.warning(f"No data found for {selected}.")
        return
    row = row_df.iloc[0]

    company = str(row.get("Company Name", ""))
    st.markdown(f"## {selected}")
    if company and company.lower() not in ("n/a", "nan", "none"):
        st.markdown(f"*{company}*")

    if not chosen:
        st.info("Select at least one field from the sidebar.")
        return

    st.divider()

    # ── Metric cards (Price block) ─────────────────────────────────
    metric_show = [f for f in TOP_METRICS if f in chosen]
    if metric_show:
        cols = st.columns(len(metric_show))
        for col, field in zip(cols, metric_show):
            raw  = row.get(field)
            disp, css = fmt_value(field, raw)
            with col:
                if field == "Change %" and css in ("pos", "neg"):
                    try:
                        v = float(raw)
                        st.metric(label=field, value=disp,
                                  delta=f"{v:+.2f}%", delta_color="normal")
                    except Exception:
                        st.metric(label=field, value=disp)
                else:
                    st.metric(label=field, value=disp)
        st.divider()

    # ── Report sections ────────────────────────────────────────────
    skip = set(TOP_METRICS)
    for group, fields in FIELD_GROUPS.items():
        section = [f for f in fields if f in chosen and f not in skip]
        if not section:
            continue

        st.markdown(f'<div class="section-hdr">{group.upper()}</div>',
                    unsafe_allow_html=True)

        for i, field in enumerate(section):
            raw  = row.get(field)
            disp, css = fmt_value(field, raw)

            if field == "Company Description":
                st.caption("Company Description")
                if css == "na":
                    st.markdown('<span class="na">N/A</span>',
                                unsafe_allow_html=True)
                else:
                    st.write(str(raw))
                continue

            if field == "Company Name":
                # Already shown as the page header — skip duplicate row
                continue

            c1, c2 = st.columns([2, 3])
            with c1:
                st.markdown(
                    f'<span style="color:#6a85a0;font-size:{font_sz}px">{field}</span>',
                    unsafe_allow_html=True,
                )
            with c2:
                if css:
                    st.markdown(
                        f'<span class="{css}" style="font-size:{font_sz}px">{disp}</span>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<span style="font-size:{font_sz}px;font-weight:600">{disp}</span>',
                        unsafe_allow_html=True,
                    )

        st.markdown("")  # spacer between sections


if __name__ == "__main__":
    main()
