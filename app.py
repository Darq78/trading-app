"""
Trading App Scoring — DCF / DE Intrinsic Value Calculator
Streamlit app calculating intrinsic value using the exact logic from Calculator_Intrinsic_Value.xlsx.

DCF model (20-year projection, 3 growth phases, Beta-based discount rate).
"""

from __future__ import annotations

import streamlit as st
import pandas as pd
import yfinance as yf

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

TICKERS = [
    "AAPL", "MSFT", "GOOGL", "AMZN", "NVDA", "META", "TSLA", "AVGO",
    "COST", "NFLX", "AMD", "ADBE", "INTC", "QCOM", "TXN", "ASML",
    "AMAT", "MU", "LRCX", "KLAC",
]

# Default growth rates (from Excel: B24, B30, B31)
DEFAULT_EPS_5Y = 0.12       # EPS next 5Y — overridden per-stock when available
DEFAULT_EPS_6_10Y = 0.15    # EPS next 6-10Y
DEFAULT_EPS_11_20Y = 0.0418 # EPS next 11-20Y (long-term GDP-like)

# ---------------------------------------------------------------------------
# DCF Engine — exact replica of Excel logic
# ---------------------------------------------------------------------------

def beta_to_discount_rate(beta: float) -> float:
    """
    Excel formula from B34:
    Beta ≤ 0.8  → 5%
    0.8 < β ≤ 1.05  → 6%
    1.05 < β ≤ 1.15 → 6.5%
    1.15 < β ≤ 1.25 → 7%
    1.25 < β ≤ 1.35 → 7.5%
    1.35 < β ≤ 1.45 → 7.7%
    1.45 < β ≤ 1.55 → 8%
    β > 1.55         → 8.2%
    """
    if beta <= 0.80:
        return 0.050
    if beta <= 1.05:
        return 0.060
    if beta <= 1.15:
        return 0.065
    if beta <= 1.25:
        return 0.070
    if beta <= 1.35:
        return 0.075
    if beta <= 1.45:
        return 0.077
    if beta <= 1.55:
        return 0.080
    return 0.082


def fetch_stock_data(ticker: str) -> dict | None:
    """Fetch all financial inputs from Yahoo Finance for a single ticker."""
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        cashflow = stock.cashflow

        if cashflow is None or cashflow.empty:
            return None

        # --- Operating Cash Flow (TTM) ---
        op_cf = None
        for label in ["Operating Cash Flow", "Total Cash From Operating Activities"]:
            if label in cashflow.index:
                op_cf = cashflow.loc[label].dropna().iloc[0]
                break
        if op_cf is None:
            return None

        # --- Capital Expenditures (average of up to 3 most recent years) ---
        capex_values = []
        for label in ["Capital Expenditure", "Capital Expenditures"]:
            if label in cashflow.index:
                capex_values = cashflow.loc[label].dropna().tolist()[:3]
                break
        avg_capex = sum(capex_values) / len(capex_values) if capex_values else 0
        # CapEx is typically negative in Yahoo Finance; we need its absolute value
        avg_capex = abs(avg_capex)

        # --- Free Cash Flow base = Operating CF - avg(CapEx) ---
        # This matches Excel B16: =B15-B11 where B11=AVERAGE(B12:B14)
        base_cf = float(op_cf) - avg_capex

        # --- Other inputs ---
        current_price = info.get("currentPrice") or info.get("regularMarketPrice")
        shares_outstanding = info.get("sharesOutstanding")
        beta = info.get("beta")
        total_debt = info.get("totalDebt", 0) or 0
        cash = info.get("totalCash", 0) or 0
        company_name = info.get("shortName", ticker)

        # EPS growth next 5Y — use analyst consensus from growth_estimates.
        # Priority: growth_estimates['+1y'] > earnings_estimate['+1y'] > info > default
        # Cap to 3-25% range for a realistic 5-year forward projection.
        EPS_MIN, EPS_MAX = 0.03, 0.25
        raw_eps = None
        eps_source = "default"

        # 1) Analyst consensus: growth_estimates '+1y' stockTrend
        try:
            ge = stock.growth_estimates
            if ge is not None and not ge.empty and "+1y" in ge.index:
                val = ge.loc["+1y", "stockTrend"]
                if pd.notna(val):
                    raw_eps = float(val)
                    eps_source = "growth_estimates +1y"
        except Exception:
            pass

        # 2) Fallback: earnings_estimate '+1y' growth
        if raw_eps is None:
            try:
                ee = stock.earnings_estimate
                if ee is not None and not ee.empty and "+1y" in ee.index:
                    val = ee.loc["+1y", "growth"]
                    if pd.notna(val):
                        raw_eps = float(val)
                        eps_source = "earnings_estimate +1y"
            except Exception:
                pass

        # 3) Fallback: info earningsGrowth / revenueGrowth
        if raw_eps is None:
            raw_eps = info.get("earningsGrowth") or info.get("revenueGrowth")
            if raw_eps is not None:
                raw_eps = float(raw_eps)
                eps_source = "info earningsGrowth"

        # Clamp to reasonable range or use default
        if raw_eps is not None and raw_eps > 0:
            eps_5y = max(EPS_MIN, min(raw_eps, EPS_MAX))
            eps_capped = raw_eps > EPS_MAX or raw_eps < EPS_MIN
        else:
            eps_5y = DEFAULT_EPS_5Y
            eps_capped = False

        # Check if operating cash flow is growing (compare last 2 years)
        growing_cf = True
        for label in ["Operating Cash Flow", "Total Cash From Operating Activities"]:
            if label in cashflow.index:
                cf_series = cashflow.loc[label].dropna()
                if len(cf_series) >= 2:
                    growing_cf = float(cf_series.iloc[0]) > float(cf_series.iloc[1])
                break

        # For DE model (not growing CF), use net income instead
        net_income = None
        if not growing_cf:
            bs = stock.income_stmt
            if bs is not None and not bs.empty:
                for label in ["Net Income", "Net Income Common Stockholders"]:
                    if label in bs.index:
                        net_income = float(bs.loc[label].dropna().iloc[0])
                        break

        if current_price is None or shares_outstanding is None or shares_outstanding == 0:
            return None
        if beta is None:
            beta = 1.0  # fallback

        return {
            "ticker": ticker,
            "name": company_name,
            "current_price": float(current_price),
            "op_cf": float(op_cf),
            "avg_capex": avg_capex,
            "base_cf": base_cf,
            "net_income": net_income,
            "growing_cf": growing_cf,
            "shares_outstanding": float(shares_outstanding),
            "total_debt": float(total_debt),
            "cash": float(cash),
            "beta": float(beta),
            "raw_eps": float(raw_eps) if raw_eps is not None else None,
            "eps_source": eps_source,
            "eps_capped": eps_capped,
            "eps_5y": float(eps_5y) if eps_5y else DEFAULT_EPS_5Y,
        }
    except Exception:
        return None


def calculate_intrinsic_value(
    data: dict,
    eps_5y_override: float | None,
    eps_6_10y: float,
    eps_11_20y: float,
) -> dict | None:
    """
    Calculate intrinsic value per share — exact Excel logic:

    1. Determine valuation method: DCF (growing CF) or DE (not growing CF)
    2. Base value = FCF (for DCF) or Net Income (for DE)
    3. Project cash flows for 20 years in 3 phases:
       - Years 1-5:   grow at EPS next 5Y rate
       - Years 6-10:  grow at EPS next 6-10Y rate
       - Years 11-20: grow at EPS next 11-20Y rate
    4. Discount factors: 1/(1+r), compounding each year
    5. Discounted CF = projected CF × discount factor for each year
    6. Sum all 20 discounted CFs
    7. Value per share = sum / shares outstanding
    8. Intrinsic Value = Value per share - Debt per share + Cash per share
    9. Under/Overpriced = (Market Price - Intrinsic Value) / Intrinsic Value
    """
    try:
        method = "DCF" if data["growing_cf"] else "DE"

        # Base cash flow
        if method == "DCF":
            base = data["base_cf"]
        else:
            base = data["net_income"]
            if base is None or base <= 0:
                return None

        if base <= 0:
            return None

        eps_5y = eps_5y_override if eps_5y_override is not None else data["eps_5y"]
        discount_rate = beta_to_discount_rate(data["beta"])
        shares = data["shares_outstanding"]

        # --- Step 3: Project Cash Flows (20 years, 3 phases) ---
        projected = []
        prev = base
        for year in range(1, 21):
            if year <= 5:
                growth = eps_5y
            elif year <= 10:
                growth = eps_6_10y
            else:
                growth = eps_11_20y
            prev = prev * (1 + growth)
            projected.append(prev)

        # --- Step 4: Discount Factors ---
        discount_factors = []
        df = 1.0
        for year in range(1, 21):
            df = df / (1 + discount_rate)
            discount_factors.append(df)

        # --- Step 5-6: Discounted Cash Flows ---
        discounted = [p * d for p, d in zip(projected, discount_factors)]
        total_dcf = sum(discounted)

        # --- Step 7-8: Intrinsic Value per Share ---
        value_per_share = total_dcf / shares
        debt_per_share = data["total_debt"] / shares
        cash_per_share = data["cash"] / shares
        intrinsic_value = value_per_share - debt_per_share + cash_per_share

        # --- Step 9: Under/Overpriced ---
        # Skip if intrinsic value is non-positive or absurdly low vs market price
        # (e.g. price > 3x IV means model inputs are unreliable for this stock)
        if intrinsic_value <= 0:
            return None
        if data["current_price"] > intrinsic_value * 3:
            return None
        under_over = (data["current_price"] - intrinsic_value) / intrinsic_value
        diff_pct = -under_over * 100  # positive = undervalued

        return {
            "ticker": data["ticker"],
            "name": data["name"],
            "method": method,
            "current_price": data["current_price"],
            "intrinsic_value": round(intrinsic_value, 2),
            "diff_pct": round(diff_pct, 1),
            "beta": data["beta"],
            "discount_rate": discount_rate,
            "eps_5y": eps_5y,
            "base_cf": base,
            "total_dcf": round(total_dcf, 0),
            "value_per_share": round(value_per_share, 2),
            "debt_per_share": round(debt_per_share, 2),
            "cash_per_share": round(cash_per_share, 2),
            "total_debt": data["total_debt"],
            "cash": data["cash"],
            "shares_outstanding": data["shares_outstanding"],
        }
    except Exception:
        return None


def compute_all(
    tickers: list[str],
    eps_6_10y: float,
    eps_11_20y: float,
    progress_bar=None,
) -> tuple[pd.DataFrame, list[str], list[dict]]:
    """Fetch data and calculate intrinsic value for all tickers.
    Returns (results_df, list_of_skipped_tickers, raw_data_for_all_fetched)."""
    results = []
    skipped = []
    all_raw = []
    for i, ticker in enumerate(tickers):
        if progress_bar:
            progress_bar.progress((i + 1) / len(tickers), text=f"Pobieram: {ticker}...")
        data = fetch_stock_data(ticker)
        if data is None:
            skipped.append(ticker)
            continue
        all_raw.append(data)
        result = calculate_intrinsic_value(data, None, eps_6_10y, eps_11_20y)
        if result is None:
            skipped.append(ticker)
            continue
        results.append(result)

    if not results:
        return pd.DataFrame(), skipped, all_raw

    df = pd.DataFrame(results)
    df = df.sort_values("diff_pct", ascending=False).reset_index(drop=True)
    return df, skipped, all_raw


# ---------------------------------------------------------------------------
# Data quality checks for verification tab
# ---------------------------------------------------------------------------

def assess_quality(row: dict) -> tuple[str, list[str]]:
    """Return (status_emoji, list_of_reasons) for a single stock's raw data.

    Returns:
        ("green" | "yellow" | "red",  [reason strings])
    """
    issues_yellow = []
    issues_red = []

    ocf_m = row["op_cf"] / 1e6
    capex_m = row["avg_capex"] / 1e6
    fcf_m = row["base_cf"] / 1e6
    shares_m = row["shares_outstanding"] / 1e6
    raw_eps = row["raw_eps"]

    # RED checks — data is probably wrong
    if shares_m < 1:
        issues_red.append(f"Akcje < 1M ({shares_m:.2f}M)")
    if ocf_m == 0:
        issues_red.append("OCF = 0")
    if ocf_m > 1_000_000:
        issues_red.append(f"OCF > 1 000 000M ({ocf_m:,.0f}M) — prawdopodobnie bledne jednostki")
    if row["total_debt"] / 1e6 > 1_000_000:
        issues_red.append(f"Dlug > 1 000 000M — prawdopodobnie bledne jednostki")
    if row["cash"] / 1e6 > 1_000_000:
        issues_red.append(f"Cash > 1 000 000M — prawdopodobnie bledne jednostki")

    # YELLOW checks — something is suspicious
    eps_capped = row.get("eps_capped", False)
    eps_source = row.get("eps_source", "default")
    if raw_eps is not None and raw_eps > 0.25:
        issues_yellow.append(f"EPS raw = {raw_eps:.0%} → ograniczono do 25% ({eps_source})")
    elif raw_eps is not None and 0 < raw_eps < 0.03:
        issues_yellow.append(f"EPS raw = {raw_eps:.1%} → podniesiono do 3% ({eps_source})")
    if raw_eps is not None and raw_eps <= 0:
        issues_yellow.append(f"EPS raw = {raw_eps:.1%} (ujemny) → domyslne 12% ({eps_source})")
    if raw_eps is None:
        issues_yellow.append("Brak EPS growth z Yahoo → domyslne 12%")
    if eps_source == "default":
        issues_yellow.append("Uzyto domyslny EPS 12% (brak danych analitycznych)")
    if fcf_m < 0:
        issues_yellow.append(f"FCF ujemny ({fcf_m:,.0f}M)")
    if not row["growing_cf"]:
        issues_yellow.append("OCF nie rosnie — metoda DE zamiast DCF")
    if row["beta"] > 2.0:
        issues_yellow.append(f"Beta bardzo wysoka ({row['beta']:.2f})")

    if issues_red:
        return "red", issues_red + issues_yellow
    if issues_yellow:
        return "yellow", issues_yellow
    return "green", []


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def color_row(row):
    """Green for undervalued (positive diff), red for overvalued (negative diff)."""
    if row["Roznica %"] > 0:
        return ["background-color: #d4edda; color: #155724"] * len(row)
    elif row["Roznica %"] < 0:
        return ["background-color: #f8d7da; color: #721c24"] * len(row)
    return [""] * len(row)


def render_valuation_tab(df, skipped):
    """Render the main valuation results tab."""
    # --- Summary metrics ---
    undervalued = len(df[df["diff_pct"] > 0])
    overvalued = len(df[df["diff_pct"] < 0])
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Przeanalizowane", len(df))
    col2.metric("Niedowartosciowane", undervalued)
    col3.metric("Przewartosciowane", overvalued)
    col4.metric("DCF / DE", f'{len(df[df["method"]=="DCF"])} / {len(df[df["method"]=="DE"])}')

    if skipped:
        st.info(
            f"Pominiete spolki (brak danych lub ujemny FCF/zysk netto): "
            f"**{', '.join(skipped)}**"
        )

    # --- Main table ---
    display_df = df[[
        "ticker", "name", "method", "current_price",
        "intrinsic_value", "diff_pct",
    ]].copy()
    display_df.columns = [
        "Ticker", "Spolka", "Metoda", "Cena ($)",
        "Wartosc wewnetrzna ($)", "Roznica %",
    ]

    styled = (
        display_df.style
        .apply(color_row, axis=1)
        .format({
            "Cena ($)": "${:,.2f}",
            "Wartosc wewnetrzna ($)": "${:,.2f}",
            "Roznica %": "{:+.1f}%",
        })
    )

    st.dataframe(styled, use_container_width=True, hide_index=True, height=740)

    # --- Expandable: DCF breakdown ---
    with st.expander("Szczegoly kalkulacji"):
        detail_df = df[[
            "ticker", "method", "beta", "discount_rate", "eps_5y",
            "base_cf", "total_dcf", "value_per_share",
            "debt_per_share", "cash_per_share", "intrinsic_value",
        ]].copy()
        detail_df.columns = [
            "Ticker", "Metoda", "Beta", "Stopa dysk.", "EPS 5Y",
            "Baza CF ($)", "Suma DCF ($)", "Wartosc/akcje ($)",
            "Dlug/akcje ($)", "Gotowka/akcje ($)", "Wart. wewn. ($)",
        ]
        st.dataframe(
            detail_df.style.format({
                "Beta": "{:.2f}",
                "Stopa dysk.": "{:.1%}",
                "EPS 5Y": "{:.1%}",
                "Baza CF ($)": "${:,.0f}",
                "Suma DCF ($)": "${:,.0f}",
                "Wartosc/akcje ($)": "${:,.2f}",
                "Dlug/akcje ($)": "${:,.2f}",
                "Gotowka/akcje ($)": "${:,.2f}",
                "Wart. wewn. ($)": "${:,.2f}",
            }),
            use_container_width=True,
            hide_index=True,
        )


def render_verification_tab(all_raw):
    """Render the data verification / quality control tab."""
    if not all_raw:
        st.warning("Brak danych do weryfikacji. Odswierz dane na zakladce Wycena.")
        return

    # Build verification table and quality assessments
    rows = []
    for d in all_raw:
        status, reasons = assess_quality(d)
        method = "DCF" if d["growing_cf"] else "DE"
        emoji = {"green": "🟢", "yellow": "🟡", "red": "🔴"}[status]
        rows.append({
            "Status": emoji,
            "Ticker": d["ticker"],
            "OCF TTM ($M)": round(d["op_cf"] / 1e6, 1),
            "CapEx sr. 3 lata ($M)": round(d["avg_capex"] / 1e6, 1),
            "FCF ($M)": round(d["base_cf"] / 1e6, 1),
            "Total Debt ($M)": round(d["total_debt"] / 1e6, 1),
            "Cash ($M)": round(d["cash"] / 1e6, 1),
            "Akcje (M)": round(d["shares_outstanding"] / 1e6, 1),
            "Beta": round(d["beta"], 2),
            "Stopa dysk. (%)": round(beta_to_discount_rate(d["beta"]) * 100, 1),
            "EPS Growth 5Y (%)": round(d["eps_5y"] * 100, 1),
            "EPS Yahoo raw (%)": round(d["raw_eps"] * 100, 1) if d["raw_eps"] is not None else None,
            "Zrodlo EPS": d.get("eps_source", "default"),
            "Metoda": method,
            "_status": status,
            "_reasons": reasons,
        })

    vdf = pd.DataFrame(rows)

    # --- Summary counts ---
    n_green = sum(1 for r in rows if r["_status"] == "green")
    n_yellow = sum(1 for r in rows if r["_status"] == "yellow")
    n_red = sum(1 for r in rows if r["_status"] == "red")

    st.markdown(
        f"**Kontrola jakosci danych:** "
        f"🟢 {n_green} OK  &nbsp;&nbsp; "
        f"🟡 {n_yellow} watpliwe  &nbsp;&nbsp; "
        f"🔴 {n_red} prawdopodobnie bledne  &nbsp;&nbsp; "
        f"(lacznie {len(rows)} spolek)"
    )

    # --- Main verification table ---
    display_cols = [
        "Status", "Ticker", "OCF TTM ($M)", "CapEx sr. 3 lata ($M)",
        "FCF ($M)", "Total Debt ($M)", "Cash ($M)", "Akcje (M)",
        "Beta", "Stopa dysk. (%)", "EPS Growth 5Y (%)", "EPS Yahoo raw (%)",
        "Zrodlo EPS", "Metoda",
    ]

    def color_status_row(row):
        status_char = row["Status"]
        if status_char == "🔴":
            return ["background-color: #f8d7da; color: #721c24"] * len(row)
        elif status_char == "🟡":
            return ["background-color: #fff3cd; color: #856404"] * len(row)
        elif status_char == "🟢":
            return ["background-color: #d4edda; color: #155724"] * len(row)
        return [""] * len(row)

    styled_v = (
        vdf[display_cols].style
        .apply(color_status_row, axis=1)
        .format({
            "OCF TTM ($M)": "{:,.1f}",
            "CapEx sr. 3 lata ($M)": "{:,.1f}",
            "FCF ($M)": "{:,.1f}",
            "Total Debt ($M)": "{:,.1f}",
            "Cash ($M)": "{:,.1f}",
            "Akcje (M)": "{:,.1f}",
            "Beta": "{:.2f}",
            "Stopa dysk. (%)": "{:.1f}%",
            "EPS Growth 5Y (%)": "{:.1f}%",
            "EPS Yahoo raw (%)": "{:.1f}%",
        }, na_rep="brak")
    )

    st.dataframe(styled_v, use_container_width=True, hide_index=True, height=740)

    # --- Detailed issues per stock ---
    flagged = [r for r in rows if r["_reasons"]]
    if flagged:
        with st.expander(f"Szczegoly problemow ({len(flagged)} spolek)"):
            for r in flagged:
                emoji = r["Status"]
                reasons_str = " | ".join(r["_reasons"])
                st.markdown(f"{emoji} **{r['Ticker']}**: {reasons_str}")


def main():
    st.set_page_config(
        page_title="DCF Intrinsic Value Calculator",
        page_icon="📊",
        layout="wide",
    )

    st.title("DCF Intrinsic Value Calculator")
    st.caption(
        "Wycena wartosci wewnetrznej spolek metoda zdyskontowanych przeplywow pienieznych (DCF) / "
        "zdyskontowanych zyskow (DE) — logika z Calculator_Intrinsic_Value.xlsx"
    )

    # --- Sidebar: growth assumptions ---
    with st.sidebar:
        st.header("Zalozenia wzrostu")
        st.markdown(
            "Stopa dyskontowa jest wyznaczana automatycznie na podstawie **Beta** "
            "(tabela progowa z Excela)."
        )

        st.subheader("Fazy wzrostu (20 lat)")

        st.markdown("**Lata 1-5:** stopa EPS next 5Y pobierana z Yahoo Finance per spolka")

        eps_6_10y = st.slider(
            "Lata 6-10: EPS growth (%)", 1.0, 30.0, DEFAULT_EPS_6_10Y * 100, 0.5,
            help="Domyslnie 15% — z Excela B30",
        ) / 100

        eps_11_20y = st.slider(
            "Lata 11-20: EPS growth (%)", 1.0, 15.0, DEFAULT_EPS_11_20Y * 100, 0.1,
            help="Domyslnie 4.18% — z Excela B31 (zblizona do dlugoterminowego wzrostu PKB)",
        ) / 100

        st.divider()
        st.markdown("**Tabela stop dyskontowych (Beta -> r):**")
        beta_table = pd.DataFrame({
            "Beta": ["<= 0.80", "0.81-1.05", "1.06-1.15", "1.16-1.25",
                      "1.26-1.35", "1.36-1.45", "1.46-1.55", "> 1.55"],
            "r": ["5.0%", "6.0%", "6.5%", "7.0%", "7.5%", "7.7%", "8.0%", "8.2%"],
        })
        st.dataframe(beta_table, hide_index=True, use_container_width=True)

        st.divider()
        st.markdown("**Metoda wyceny:**")
        st.markdown(
            "- **DCF** — gdy Operating Cash Flow rosnie (baza = FCF)\n"
            "- **DE** — gdy nie rosnie (baza = Net Income)"
        )

    # --- Main area ---
    if st.button("Odswiez dane", type="primary", use_container_width=True):
        st.session_state.pop("df_results", None)
        st.session_state.pop("params", None)
        st.session_state.pop("all_raw", None)
        st.session_state.pop("skipped", None)

    current_params = (eps_6_10y, eps_11_20y)

    if st.session_state.get("params") != current_params:
        st.session_state.pop("df_results", None)

    if "df_results" not in st.session_state:
        progress = st.progress(0, text="Pobieram dane z Yahoo Finance...")
        df, skipped, all_raw = compute_all(TICKERS, eps_6_10y, eps_11_20y, progress_bar=progress)
        progress.empty()
        st.session_state["df_results"] = df
        st.session_state["skipped"] = skipped
        st.session_state["all_raw"] = all_raw
        st.session_state["params"] = current_params

    df = st.session_state.get("df_results", pd.DataFrame())
    skipped = st.session_state.get("skipped", [])
    all_raw = st.session_state.get("all_raw", [])

    # --- Tabs ---
    tab_valuation, tab_verification = st.tabs(["Wycena DCF", "Weryfikacja danych"])

    with tab_valuation:
        if df.empty:
            st.warning("Brak danych do wyswietlenia. Sprobuj odswiez.")
        else:
            render_valuation_tab(df, skipped)

    with tab_verification:
        render_verification_tab(all_raw)

    st.divider()
    st.caption(
        "Dane: Yahoo Finance | Metoda: DCF/DE (20-letnia projekcja, 3 fazy wzrostu) | "
        "Logika: Calculator_Intrinsic_Value.xlsx | Nie stanowi porady inwestycyjnej"
    )


if __name__ == "__main__":
    main()
