import streamlit as st
import pandas as pd
import requests
import io
import math

st.set_page_config(
    page_title="Stock Viewer",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Field definitions ─────────────────────────────────────────────────────────
FIELD_GROUPS = {
    "Overview":      ["Company Name", "Company Description", "Mkt Cap/Net Assets (M)", "Last Updated"],
    "Price":         ["Price", "Prev Close", "Change %", "Day High", "Day Low"],
    "Earnings":      ["EPS (TTM)"],
    "Revenue":       ["Revenue TTM (M)", "Revenue 1Y (M)", "Revenue 3Y (M)", "Revenue 5Y (M)"],
    "Net Income":    ["Net Income TTM (M)", "Net Income 1Y (M)", "Net Income 3Y (M)", "Net Income 5Y (M)"],
    "Balance Sheet": ["Total Debt (M)", "Cash (M)"],
}

PRICE_FIELDS   = {"Price", "Prev Close", "Day High", "Day Low"}
MILLION_FIELDS = {
    "Mkt Cap/Net Assets (M)",
    "Revenue TTM (M)", "Revenue 1Y (M)", "Revenue 3Y (M)", "Revenue 5Y (M)",
    "Net Income TTM (M)", "Net Income 1Y (M)", "Net Income 3Y (M)", "Net Income 5Y (M)",
    "Total Debt (M)", "Cash (M)",
}

# Shown as metric cards at the top of the report
TOP_METRICS = ["Price", "Change %", "Mkt Cap/Net Assets (M)", "Day High", "Day Low"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def is_na(raw):
    if raw is None:
        return True
    if isinstance(raw, float) and math.isnan(raw):
        return True
    return str(raw).strip().lower() in ("", "n/a", "none", "nan", "nat")


def fmt_value(field, raw):
    """Return (display_string, css_class).  css_class is 'na', 'pos', 'neg', or ''."""
    if is_na(raw):
        return "N/A", "na"
    try:
        v = float(raw)
        if field == "Change %":
            s = f"{'+' if v >= 0 else ''}{v:.2f}%"
            return s, "pos" if v >= 0 else "neg"
        if field in MILLION_FIELDS:
            return f"${v:,.2f} M", ""
        if field in PRICE_FIELDS or field == "EPS (TTM)":
            return f"${v:,.2f}", ""
    except (ValueError, TypeError):
        pass
    return str(raw), ""


# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    """Fetch stocks.xlsx from the private GitHub repo using stored secrets."""
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


# ── App ───────────────────────────────────────────────────────────────────────
def main():
    st.markdown("""
    <style>
    /* Coloured value spans */
    .na  { color: #c0392b; font-style: italic; }
    .pos { color: #27ae60; font-weight: 700; }
    .neg { color: #c0392b; font-weight: 700; }
    /* Section header bar */
    .section-hdr {
        background: #1F4E79; color: white;
        padding: 5px 12px; border-radius: 5px;
        font-size: 0.92em; font-weight: 700;
        margin: 14px 0 6px 0;
        letter-spacing: 0.04em;
    }
    /* Tighten up metric cards on mobile */
    div[data-testid="metric-container"] { padding: 8px 10px !important; }
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("📈 Stock Viewer")

        if st.button("🔄 Refresh Data", use_container_width=True,
                     help="Pull the latest stocks.xlsx from GitHub"):
            load_data.clear()
            st.rerun()
        st.caption("Data auto-refreshes every 5 min. Press Refresh for immediate update.")
        st.divider()

        with st.spinner("Loading data..."):
            df, err = load_data()

        if err:
            st.error(err)
            st.stop()
        if df is None or df.empty:
            st.warning("stocks.xlsx is empty or missing from the repository.")
            st.stop()

        tickers = sorted(df["Stock Ticker"].dropna().astype(str).unique())
        selected = st.selectbox("Ticker", tickers)

        st.markdown("**Fields to display**")
        chosen = []
        for group, fields in FIELD_GROUPS.items():
            with st.expander(group, expanded=True):
                for field in fields:
                    if st.checkbox(field, value=True, key=f"chk_{field}"):
                        chosen.append(field)

    # ── Main panel ─────────────────────────────────────────────────────────────
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

    # ── Metric cards (Price block) ─────────────────────────────────────────────
    metric_show = [f for f in TOP_METRICS if f in chosen]
    if metric_show:
        cols = st.columns(len(metric_show))
        for col, field in zip(cols, metric_show):
            raw  = row.get(field)
            disp, css = fmt_value(field, raw)
            with col:
                if field == "Change %" and css in ("pos", "neg"):
                    v = float(raw)
                    st.metric(label=field, value=disp,
                              delta=f"{v:+.2f}%", delta_color="normal")
                else:
                    st.metric(label=field, value=disp)
        st.divider()

    # ── Remaining sections ─────────────────────────────────────────────────────
    skip = set(TOP_METRICS)
    for group, fields in FIELD_GROUPS.items():
        section = [f for f in fields if f in chosen and f not in skip]
        if not section:
            continue

        st.markdown(f'<div class="section-hdr">{group}</div>', unsafe_allow_html=True)

        for field in section:
            raw  = row.get(field)
            disp, css = fmt_value(field, raw)

            if field == "Company Description":
                st.caption(field)
                if css == "na":
                    st.markdown('<span class="na">N/A</span>', unsafe_allow_html=True)
                else:
                    st.write(str(raw))
                continue

            c1, c2 = st.columns([2, 3])
            with c1:
                st.caption(field)
            with c2:
                if css:
                    st.markdown(f'<span class="{css}">{disp}</span>', unsafe_allow_html=True)
                else:
                    st.write(disp)

        st.markdown("")  # spacer between sections


if __name__ == "__main__":
    main()
