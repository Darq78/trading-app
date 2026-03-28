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

        # EPS growth next 5Y from Yahoo Finance.
        # earningsGrowth is YoY historical growth which can be extreme (e.g. 756% for MU).
        # We cap it to a reasonable range for a 5-year forward projection.
        raw_eps = info.get("earningsGrowth") or info.get("revenueGrowth")
        if raw_eps is not None and 0 < raw_eps <= 0.30:
            eps_5y = raw_eps
        else:
            eps_5y = DEFAULT_EPS_5Y

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
            "base_cf": base_cf,
            "net_income": net_income,
            "growing_cf": growing_cf,
            "shares_outstanding": float(shares_outstanding),
            "total_debt": float(total_debt),
            "cash": float(cash),
            "beta": float(beta),
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
) -> tuple[pd.DataFrame, list[str]]:
    """Fetch data and calculate intrinsic value for all tickers.
    Returns (results_df, list_of_skipped_tickers)."""
    results = []
    skipped = []
    for i, ticker in enumerate(tickers):
        if progress_bar:
            progress_bar.progress((i + 1) / len(tickers), text=f"Pobieram: {ticker}...")
        data = fetch_stock_data(ticker)
        if data is None:
            skipped.append(ticker)
            continue
        result = calculate_intrinsic_value(data, None, eps_6_10y, eps_11_20y)
        if result is None:
            skipped.append(ticker)
            continue
        results.append(result)

    if not results:
        return pd.DataFrame(), skipped

    df = pd.DataFrame(results)
    df = df.sort_values("diff_pct", ascending=False).reset_index(drop=True)
    return df, skipped


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def color_row(row):
    """Green for undervalued (positive diff), red for overvalued (negative diff)."""
    if row["Różnica %"] > 0:
        return ["background-color: #d4edda; color: #155724"] * len(row)
    elif row["Różnica %"] < 0:
        return ["background-color: #f8d7da; color: #721c24"] * len(row)
    return [""] * len(row)


def main():
    st.set_page_config(
        page_title="DCF Intrinsic Value Calculator",
        page_icon="📊",
        layout="wide",
    )

    st.title("DCF Intrinsic Value Calculator")
    st.caption(
        "Wycena wartości wewnętrznej spółek metodą zdyskontowanych przepływów pieniężnych (DCF) / "
        "zdyskontowanych zysków (DE) — logika z Calculator_Intrinsic_Value.xlsx"
    )

    # --- Sidebar: growth assumptions ---
    with st.sidebar:
        st.header("Założenia wzrostu")
        st.markdown(
            "Stopa dyskontowa jest wyznaczana automatycznie na podstawie **Beta** "
            "(tabela progowa z Excela)."
        )

        st.subheader("Fazy wzrostu (20 lat)")

        st.markdown("**Lata 1-5:** stopa EPS next 5Y pobierana z Yahoo Finance per spółka")

        eps_6_10y = st.slider(
            "Lata 6-10: EPS growth (%)", 1.0, 30.0, DEFAULT_EPS_6_10Y * 100, 0.5,
            help="Domyślnie 15% — z Excela B30",
        ) / 100

        eps_11_20y = st.slider(
            "Lata 11-20: EPS growth (%)", 1.0, 15.0, DEFAULT_EPS_11_20Y * 100, 0.1,
            help="Domyślnie 4.18% — z Excela B31 (zbliżona do długoterminowego wzrostu PKB)",
        ) / 100

        st.divider()
        st.markdown("**Tabela stóp dyskontowych (Beta → r):**")
        beta_table = pd.DataFrame({
            "Beta": ["≤ 0.80", "0.81–1.05", "1.06–1.15", "1.16–1.25",
                      "1.26–1.35", "1.36–1.45", "1.46–1.55", "> 1.55"],
            "r": ["5.0%", "6.0%", "6.5%", "7.0%", "7.5%", "7.7%", "8.0%", "8.2%"],
        })
        st.dataframe(beta_table, hide_index=True, use_container_width=True)

        st.divider()
        st.markdown("**Metoda wyceny:**")
        st.markdown(
            "- **DCF** — gdy Operating Cash Flow rośnie (baza = FCF)\n"
            "- **DE** — gdy nie rośnie (baza = Net Income)"
        )

    # --- Main area ---
    if st.button("Odśwież dane", type="primary", use_container_width=True):
        st.session_state.pop("df_results", None)
        st.session_state.pop("params", None)

    current_params = (eps_6_10y, eps_11_20y)

    if st.session_state.get("params") != current_params:
        st.session_state.pop("df_results", None)

    if "df_results" not in st.session_state:
        progress = st.progress(0, text="Pobieram dane z Yahoo Finance...")
        df, skipped = compute_all(TICKERS, eps_6_10y, eps_11_20y, progress_bar=progress)
        progress.empty()
        st.session_state["df_results"] = df
        st.session_state["skipped"] = skipped
        st.session_state["params"] = current_params

    df = st.session_state.get("df_results", pd.DataFrame())
    skipped = st.session_state.get("skipped", [])

    if df.empty:
        st.warning("Brak danych do wyświetlenia. Spróbuj odświeżyć.")
        return

    # --- Summary metrics ---
    undervalued = len(df[df["diff_pct"] > 0])
    overvalued = len(df[df["diff_pct"] < 0])
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Przeanalizowane", len(df))
    col2.metric("Niedowartościowane", undervalued)
    col3.metric("Przewartościowane", overvalued)
    col4.metric("DCF / DE", f'{len(df[df["method"]=="DCF"])} / {len(df[df["method"]=="DE"])}')

    if skipped:
        st.info(
            f"Pominięte spółki (brak danych lub ujemny FCF/zysk netto): "
            f"**{', '.join(skipped)}**"
        )

    # --- Main table ---
    display_df = df[[
        "ticker", "name", "method", "current_price",
        "intrinsic_value", "diff_pct",
    ]].copy()
    display_df.columns = [
        "Ticker", "Spółka", "Metoda", "Cena ($)",
        "Wartość wewnętrzna ($)", "Różnica %",
    ]

    styled = (
        display_df.style
        .apply(color_row, axis=1)
        .format({
            "Cena ($)": "${:,.2f}",
            "Wartość wewnętrzna ($)": "${:,.2f}",
            "Różnica %": "{:+.1f}%",
        })
    )

    st.dataframe(styled, use_container_width=True, hide_index=True, height=740)

    # --- Expandable: DCF breakdown ---
    with st.expander("Szczegóły kalkulacji"):
        detail_df = df[[
            "ticker", "method", "beta", "discount_rate", "eps_5y",
            "base_cf", "total_dcf", "value_per_share",
            "debt_per_share", "cash_per_share", "intrinsic_value",
        ]].copy()
        detail_df.columns = [
            "Ticker", "Metoda", "Beta", "Stopa dysk.", "EPS 5Y",
            "Baza CF ($)", "Suma DCF ($)", "Wartość/akcję ($)",
            "Dług/akcję ($)", "Gotówka/akcję ($)", "Wart. wewn. ($)",
        ]
        st.dataframe(
            detail_df.style.format({
                "Beta": "{:.2f}",
                "Stopa dysk.": "{:.1%}",
                "EPS 5Y": "{:.1%}",
                "Baza CF ($)": "${:,.0f}",
                "Suma DCF ($)": "${:,.0f}",
                "Wartość/akcję ($)": "${:,.2f}",
                "Dług/akcję ($)": "${:,.2f}",
                "Gotówka/akcję ($)": "${:,.2f}",
                "Wart. wewn. ($)": "${:,.2f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.caption(
        "Dane: Yahoo Finance | Metoda: DCF/DE (20-letnia projekcja, 3 fazy wzrostu) | "
        "Logika: Calculator_Intrinsic_Value.xlsx | Nie stanowi porady inwestycyjnej"
    )


if __name__ == "__main__":
    main()
