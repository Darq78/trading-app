"""
Trading App Scoring — DCF Intrinsic Value Calculator
Streamlit app calculating intrinsic value using the exact logic from IV calc.xlsx Sheet4.

DCF model: 10-year projection, revenue-based growth Y1-5, 4% Y6-10,
3% terminal growth, Beta-based discount rate, FX conversion for foreign stocks.
"""

from __future__ import annotations

import streamlit as st
import pandas as pd
import yfinance as yf

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

TICKERS = [
    "AAPL", "ACN", "ADBE", "AMZN", "ASML", "AVGO", "AZO", "BABA",
    "BKNG", "CELH", "CNSWF", "CPRT", "CRM", "CRWD", "DPZ", "ELV",
    "FTNT", "GOOG", "GOOGL", "HCA", "HSY", "ICE", "LMT", "LOW",
    "LVMUY", "MA", "MEDP", "MELI", "META", "MSCI", "MSFT", "MSI",
    "NKE", "NOW", "NVDA", "PANW", "PEP", "PLTR", "SPGI", "TMO",
    "UNH", "V", "VEEV", "WM",
]

# DCF assumptions (from Sheet4)
GROWTH_6_10 = 0.04          # 4% FCF growth years 6-10
TERMINAL_GROWTH = 0.03      # 3% terminal growth rate
DEFAULT_GROWTH_5Y = 0.08    # 8% fallback when no analyst estimates

# ---------------------------------------------------------------------------
# DCF Engine — exact replica of IV calc.xlsx Sheet4
# ---------------------------------------------------------------------------

def beta_to_discount_rate(beta: float) -> float:
    """Discount rate from Beta (Sheet4 B22 formula)."""
    if beta <= 0.80: return 0.050
    if beta <= 1.05: return 0.060
    if beta <= 1.15: return 0.065
    if beta <= 1.25: return 0.070
    if beta <= 1.35: return 0.075
    if beta <= 1.45: return 0.077
    if beta <= 1.55: return 0.080
    return 0.082


@st.cache_data(ttl=7200, show_spinner=False)
def get_fx_rates() -> dict[str, float]:
    """Fetch FX rates for foreign currency conversion. Cached 2 hours."""
    rates = {"USD": 1.0}
    for pair, curr in [("EURUSD=X", "EUR"), ("CNYUSD=X", "CNY")]:
        try:
            t = yf.Ticker(pair)
            rate = t.info.get("regularMarketPrice") or t.info.get("previousClose")
            if rate:
                rates[curr] = float(rate)
        except Exception:
            pass
    if "EUR" not in rates: rates["EUR"] = 1.08
    if "CNY" not in rates: rates["CNY"] = 0.137
    return rates


def fetch_stock_data(ticker: str, fx_rates: dict[str, float]) -> dict | None:
    """Fetch all financial inputs from Yahoo Finance for a single ticker.

    Returns all raw data needed for DCF calculation, with FX conversion
    applied for stocks reporting in non-USD currencies.
    """
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        cashflow = stock.cashflow

        if cashflow is None or cashflow.empty:
            return None

        # Currency conversion factor
        fin_curr = info.get("financialCurrency", "USD")
        fx = fx_rates.get(fin_curr, 1.0)

        # --- Operating Cash Flow TTM ---
        ocf = None
        for label in ["Operating Cash Flow", "Total Cash From Operating Activities"]:
            if label in cashflow.index:
                ocf = float(cashflow.loc[label].dropna().iloc[0]) * fx
                break
        if ocf is None:
            return None

        # --- CapEx TTM (single most recent year, NOT averaged) ---
        capex = 0
        for label in ["Capital Expenditure", "Capital Expenditures"]:
            if label in cashflow.index:
                capex = abs(float(cashflow.loc[label].dropna().iloc[0])) * fx
                break

        # --- FCF = OCF - CapEx (Sheet4 B2: =B3-B4) ---
        fcf = ocf - capex

        # --- Revenue growth Y1-5 from analyst estimates ---
        # Sheet4: B5 = B7/B6 - 1 (Revenue estimate next yr / Revenue current yr - 1)
        growth_flag = "ok"
        rev_curr = rev_next = None
        growth_5y = None
        try:
            re = stock.revenue_estimate
            if re is not None and not re.empty:
                if "0y" in re.index and "+1y" in re.index:
                    rc = re.loc["0y", "avg"]
                    rn = re.loc["+1y", "avg"]
                    if pd.notna(rc) and pd.notna(rn) and rc > 0:
                        rev_curr = float(rc)
                        rev_next = float(rn)
                        growth_5y = float(rn / rc - 1)
        except Exception:
            pass

        if growth_5y is None or growth_5y <= -0.5 or growth_5y > 1.0:
            growth_5y = DEFAULT_GROWTH_5Y
            growth_flag = "default"

        # --- Other inputs ---
        current_price = info.get("currentPrice") or info.get("regularMarketPrice")
        shares = info.get("sharesOutstanding")
        beta = info.get("beta")
        total_debt = float(info.get("totalDebt", 0) or 0)
        cash = float(info.get("totalCash", 0) or 0)
        name = info.get("shortName", ticker)

        # Convert debt/cash for foreign stocks
        if fin_curr != "USD":
            total_debt *= fx
            cash *= fx

        if current_price is None or shares is None or shares == 0:
            return None
        if beta is None:
            beta = 1.0

        return {
            "ticker": ticker,
            "name": name,
            "current_price": float(current_price),
            "ocf": ocf,
            "capex": capex,
            "fcf": fcf,
            "rev_curr": rev_curr,
            "rev_next": rev_next,
            "growth_5y": growth_5y,
            "growth_flag": growth_flag,
            "shares": float(shares),
            "total_debt": total_debt,
            "cash": cash,
            "beta": float(beta),
            "fin_curr": fin_curr,
            "fx": fx,
        }
    except Exception:
        return None


def calculate_dcf(data: dict) -> dict | None:
    """Calculate intrinsic value per share — exact Sheet4 logic.

    Sheet4 formula chain:
    1. FCF = OCF - CapEx                              (B2 = B3 - B4)
    2. Growth Y1-5 = RevNext/RevCurr - 1              (B5 = B7/B6 - 1)
    3. Project FCF 10 years (compound)                 (B11:B20)
    4. Discount factors: 1/(1+r)^n                     (B24:B33)
    5. Discounted CFs = projected * discount           (B35:B44)
    6. Terminal Value = CF10*(1+gT)/(r-gT)             (B45)
    7. TV Present Value = TV/(1+r)^10                  (B46)
    8. Enterprise Value = sum(discounted CFs) + TV_PV  (B48 = B34 + B46)
    9. Equity Value = EV + Cash - Debt                 (B47 = B48 + B49 - B50)
    10. Intrinsic Value = Equity / Shares              (B52 = B47 / B51)
    """
    try:
        fcf = data["fcf"]
        if fcf <= 0:
            return None

        dr = beta_to_discount_rate(data["beta"])
        g5 = data["growth_5y"]
        denom = dr - TERMINAL_GROWTH
        if denom <= 0:
            return None

        # Step 3: Project FCF 10 years
        projected = []
        prev = fcf
        for yr in range(1, 11):
            g = g5 if yr <= 5 else GROWTH_6_10
            prev = prev * (1 + g)
            projected.append(prev)

        # Step 4: Discount factors
        disc_factors = []
        d = 1.0
        for yr in range(1, 11):
            d = d / (1 + dr)
            disc_factors.append(d)

        # Step 5: Discounted CFs
        discounted = [p * d for p, d in zip(projected, disc_factors)]
        sum_dcf = sum(discounted)

        # Step 6-7: Terminal Value
        tv = projected[9] * (1 + TERMINAL_GROWTH) / denom
        tv_pv = tv / (1 + dr) ** 10

        # Step 8: Enterprise Value (Sheet4 naming)
        ev = sum_dcf + tv_pv

        # Step 9: Equity Value = EV + Cash - Debt
        equity = ev + data["cash"] - data["total_debt"]

        # Step 10: Intrinsic Value per share
        iv = equity / data["shares"]

        if iv <= 0:
            return None

        over_under = (data["current_price"] - iv) / iv * 100

        return {
            "ticker": data["ticker"],
            "name": data["name"],
            "current_price": data["current_price"],
            "intrinsic_value": round(iv, 2),
            "over_under": round(over_under, 1),
            "beta": data["beta"],
            "discount_rate": dr,
            "growth_5y": g5,
            "fcf": fcf,
            "sum_dcf": sum_dcf,
            "tv": tv,
            "tv_pv": tv_pv,
            "ev": ev,
            "equity": equity,
            "total_debt": data["total_debt"],
            "cash": data["cash"],
            "shares": data["shares"],
            "ocf": data["ocf"],
            "capex": data["capex"],
            "growth_flag": data["growth_flag"],
            "rev_curr": data["rev_curr"],
            "rev_next": data["rev_next"],
            "fin_curr": data["fin_curr"],
            "fx": data["fx"],
        }
    except Exception:
        return None


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_all_raw(tickers: tuple[str, ...]) -> list[dict]:
    """Fetch raw data for all tickers. Cached for 1 hour."""
    fx_rates = get_fx_rates()
    all_raw = []
    for ticker in tickers:
        data = fetch_stock_data(ticker, fx_rates)
        if data is not None:
            all_raw.append(data)
    return all_raw


def compute_all(all_raw: list[dict]) -> tuple[pd.DataFrame, list[str]]:
    """Calculate DCF for all fetched tickers."""
    results = []
    skipped = []
    fetched = {d["ticker"] for d in all_raw}

    for t in TICKERS:
        if t not in fetched:
            skipped.append(t)

    for data in all_raw:
        result = calculate_dcf(data)
        if result is None:
            skipped.append(data["ticker"])
            continue
        results.append(result)

    if not results:
        return pd.DataFrame(), skipped

    df = pd.DataFrame(results)
    df = df.sort_values("over_under", ascending=True).reset_index(drop=True)
    return df, skipped


# ---------------------------------------------------------------------------
# Data quality checks
# ---------------------------------------------------------------------------

def assess_quality(row: dict) -> tuple[str, list[str]]:
    """Quality flag: green/yellow/red with reasons."""
    issues_yellow = []
    issues_red = []

    M = 1e6
    ocf_m = row["ocf"] / M
    fcf_m = row["fcf"] / M
    shares_m = row["shares"] / M

    # RED — data probably wrong
    if shares_m < 1:
        issues_red.append(f"Shares < 1M ({shares_m:.2f}M)")
    if ocf_m == 0:
        issues_red.append("OCF = 0")
    if abs(ocf_m) > 1_000_000:
        issues_red.append(f"OCF > 1,000,000M — likely wrong units")

    # YELLOW — suspicious
    if row["growth_flag"] == "default":
        issues_yellow.append("No analyst revenue estimates — using default 8%")
    if fcf_m < 0:
        issues_yellow.append(f"FCF negative ({fcf_m:,.0f}M)")
    if row["beta"] > 2.0:
        issues_yellow.append(f"High Beta ({row['beta']:.2f})")
    if row["fin_curr"] != "USD":
        issues_yellow.append(f"FX conversion: {row['fin_curr']} -> USD (rate: {row['fx']:.4f})")
    if row["growth_5y"] > 0.30:
        issues_yellow.append(f"Revenue growth {row['growth_5y']:.0%} > 30%")

    if issues_red:
        return "red", issues_red + issues_yellow
    if issues_yellow:
        return "yellow", issues_yellow
    return "green", []


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def color_valuation_row(row):
    """Green when IV > price (undervalued), red when IV < price (overvalued)."""
    if row["Over/Under (%)"] < 0:
        return ["background-color: #d4edda; color: #155724"] * len(row)
    elif row["Over/Under (%)"] > 0:
        return ["background-color: #f8d7da; color: #721c24"] * len(row)
    return [""] * len(row)


def render_valuation_tab(df, skipped):
    """Main valuation results tab."""
    undervalued = len(df[df["over_under"] < 0])
    overvalued = len(df[df["over_under"] > 0])
    c1, c2, c3 = st.columns(3)
    c1.metric("Przeanalizowane", len(df))
    c2.metric("Niedowartosciowane", undervalued)
    c3.metric("Przewartosciowane", overvalued)

    if skipped:
        st.info(f"Pominiete (brak danych / ujemny FCF): **{', '.join(skipped)}**")

    display_df = df[[
        "ticker", "name", "current_price", "intrinsic_value", "over_under",
    ]].copy()
    display_df.columns = [
        "Ticker", "Spolka", "Cena ($)", "Intrinsic Value ($)", "Over/Under (%)",
    ]

    styled = (
        display_df.style
        .apply(color_valuation_row, axis=1)
        .format({
            "Cena ($)": "${:,.2f}",
            "Intrinsic Value ($)": "${:,.2f}",
            "Over/Under (%)": "{:+.1f}%",
        })
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=740)

    with st.expander("Szczegoly kalkulacji DCF"):
        det = df[[
            "ticker", "beta", "discount_rate", "growth_5y",
            "fcf", "sum_dcf", "tv_pv", "ev", "equity", "intrinsic_value",
        ]].copy()
        det.columns = [
            "Ticker", "Beta", "Stopa dysk.", "Growth Y1-5",
            "FCF ($)", "Sum DCF ($)", "TV PV ($)", "EV ($)",
            "Equity ($)", "IV ($/share)",
        ]
        st.dataframe(
            det.style.format({
                "Beta": "{:.2f}", "Stopa dysk.": "{:.1%}", "Growth Y1-5": "{:.1%}",
                "FCF ($)": "${:,.0f}", "Sum DCF ($)": "${:,.0f}",
                "TV PV ($)": "${:,.0f}", "EV ($)": "${:,.0f}",
                "Equity ($)": "${:,.0f}", "IV ($/share)": "${:,.2f}",
            }),
            use_container_width=True, hide_index=True,
        )


def render_verification_tab(all_raw):
    """Data verification tab with quality flags."""
    if not all_raw:
        st.warning("Brak danych. Kliknij 'Odswiez dane'.")
        return

    rows = []
    for d in all_raw:
        status, reasons = assess_quality(d)
        emoji = {"green": "\U0001f7e2", "yellow": "\U0001f7e1", "red": "\U0001f534"}[status]
        M = 1e6
        B = 1e9
        rows.append({
            "Status": emoji,
            "Ticker": d["ticker"],
            "Waluta": d["fin_curr"],
            "FX": round(d["fx"], 4) if d["fin_curr"] != "USD" else None,
            "OCF TTM ($M)": round(d["ocf"] / M, 1),
            "CapEx TTM ($M)": round(d["capex"] / M, 1),
            "FCF ($M)": round(d["fcf"] / M, 1),
            "Rev Curr ($B)": round(d["rev_curr"] / B, 2) if d["rev_curr"] else None,
            "Rev Next ($B)": round(d["rev_next"] / B, 2) if d["rev_next"] else None,
            "Growth Y1-5": d["growth_5y"],
            "Growth src": d["growth_flag"],
            "Beta": round(d["beta"], 2),
            "Stopa dysk.": beta_to_discount_rate(d["beta"]),
            "Debt ($M)": round(d["total_debt"] / M, 1),
            "Cash ($M)": round(d["cash"] / M, 1),
            "Shares (M)": round(d["shares"] / M, 1),
            "_status": status,
            "_reasons": reasons,
        })

    vdf = pd.DataFrame(rows)
    n_g = sum(1 for r in rows if r["_status"] == "green")
    n_y = sum(1 for r in rows if r["_status"] == "yellow")
    n_r = sum(1 for r in rows if r["_status"] == "red")

    st.markdown(
        f"**Kontrola jakosci:** "
        f"\U0001f7e2 {n_g} OK &nbsp;&nbsp; "
        f"\U0001f7e1 {n_y} watpliwe &nbsp;&nbsp; "
        f"\U0001f534 {n_r} bledne &nbsp;&nbsp; "
        f"(lacznie {len(rows)})"
    )

    display_cols = [
        "Status", "Ticker", "Waluta", "FX",
        "OCF TTM ($M)", "CapEx TTM ($M)", "FCF ($M)",
        "Rev Curr ($B)", "Rev Next ($B)", "Growth Y1-5", "Growth src",
        "Beta", "Stopa dysk.", "Debt ($M)", "Cash ($M)", "Shares (M)",
    ]

    def color_status_row(row):
        s = row["Status"]
        if s == "\U0001f534":
            return ["background-color: #f8d7da; color: #721c24"] * len(row)
        elif s == "\U0001f7e1":
            return ["background-color: #fff3cd; color: #856404"] * len(row)
        elif s == "\U0001f7e2":
            return ["background-color: #d4edda; color: #155724"] * len(row)
        return [""] * len(row)

    styled = (
        vdf[display_cols].style
        .apply(color_status_row, axis=1)
        .format({
            "FX": "{:.4f}",
            "OCF TTM ($M)": "{:,.1f}", "CapEx TTM ($M)": "{:,.1f}",
            "FCF ($M)": "{:,.1f}", "Rev Curr ($B)": "{:,.2f}",
            "Rev Next ($B)": "{:,.2f}", "Growth Y1-5": "{:.1%}",
            "Beta": "{:.2f}", "Stopa dysk.": "{:.1%}",
            "Debt ($M)": "{:,.1f}", "Cash ($M)": "{:,.1f}",
            "Shares (M)": "{:,.1f}",
        }, na_rep="-")
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=740)

    flagged = [r for r in rows if r["_reasons"]]
    if flagged:
        with st.expander(f"Szczegoly problemow ({len(flagged)} spolek)"):
            for r in flagged:
                st.markdown(f"{r['Status']} **{r['Ticker']}**: {' | '.join(r['_reasons'])}")


def main():
    st.set_page_config(
        page_title="DCF Intrinsic Value Calculator",
        page_icon="📊",
        layout="wide",
    )

    st.title("DCF Intrinsic Value Calculator")
    st.caption(
        "Wycena 44 spolek metoda DCF — logika z IV calc.xlsx Sheet4 | "
        "10-letnia projekcja FCF, revenue-based growth, Terminal Value (Gordon Growth)"
    )

    # --- Sidebar ---
    with st.sidebar:
        st.header("Model DCF")
        st.markdown(
            "**Logika z IV calc.xlsx Sheet4:**\n"
            "- FCF = OCF TTM - CapEx TTM\n"
            "- Growth Y1-5 = Revenue Next / Revenue Curr - 1\n"
            "- Growth Y6-10 = 4%\n"
            "- Terminal Growth = 3%\n"
            "- Discount Rate = f(Beta)\n"
            "- TV = CF10 x (1+gT) / (r-gT)\n"
            "- Equity = EV + Cash - Debt\n"
            "- IV = Equity / Shares"
        )
        st.divider()
        st.markdown("**Stopy dyskontowe (Beta -> r):**")
        bt = pd.DataFrame({
            "Beta": ["<= 0.80", "0.81-1.05", "1.06-1.15", "1.16-1.25",
                      "1.26-1.35", "1.36-1.45", "1.46-1.55", "> 1.55"],
            "r": ["5.0%", "6.0%", "6.5%", "7.0%", "7.5%", "7.7%", "8.0%", "8.2%"],
        })
        st.dataframe(bt, hide_index=True, use_container_width=True)
        st.divider()
        st.caption(
            "Spolki zagraniczne (BABA, ASML, LVMUY) — dane finansowe "
            "przeliczane z CNY/EUR na USD wg aktualnego kursu."
        )

    # --- Main ---
    if st.button("Odswiez dane", type="primary", use_container_width=True):
        fetch_all_raw.clear()
        get_fx_rates.clear()
        st.session_state.pop("df_results", None)

    if "df_results" not in st.session_state:
        with st.spinner(f"Pobieram dane z Yahoo Finance dla {len(TICKERS)} spolek..."):
            all_raw = fetch_all_raw(tuple(TICKERS))
        df, skipped = compute_all(all_raw)
        st.session_state["df_results"] = df
        st.session_state["skipped"] = skipped
        st.session_state["all_raw"] = all_raw
    else:
        all_raw = st.session_state.get("all_raw", [])

    df = st.session_state.get("df_results", pd.DataFrame())
    skipped = st.session_state.get("skipped", [])

    tab1, tab2 = st.tabs(["Wycena DCF", "Weryfikacja danych"])

    with tab1:
        if df.empty:
            st.warning("Brak danych. Kliknij 'Odswiez dane'.")
        else:
            render_valuation_tab(df, skipped)

    with tab2:
        render_verification_tab(all_raw)

    st.divider()
    st.caption(
        "Dane: Yahoo Finance | Metoda: DCF 10Y + Terminal Value (IV calc.xlsx Sheet4) | "
        "Nie stanowi porady inwestycyjnej"
    )


if __name__ == "__main__":
    main()
