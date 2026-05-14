import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

APP_NAME = "Generic OEM Competitor Version"
PRODUCT_VERSION = "Product 1"
DEFAULT_DATA_PATH = Path(__file__).parent / "data" / "sample_oem_data.csv"
REQUIRED_COLUMNS = {
    "market",
    "period",
    "oem",
    "oem_grouping",
    "visitors",
    "contracts",
    "website_to_contract_conversion_rate",
}

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    .block-container { padding-top: 1.5rem; }
    .metric-card {
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 18px;
        background: white;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
        min-height: 120px;
    }
    .dark-card {
        border-radius: 22px;
        padding: 24px;
        background: #0f172a;
        color: white;
    }
    .small-label {
        color: #64748b;
        font-size: 0.82rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    .insight-card {
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 18px;
        background: #ffffff;
        min-height: 170px;
    }
    .good { color: #047857; font-weight: 700; }
    .bad { color: #be123c; font-weight: 700; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def generate_sample_data() -> pd.DataFrame:
    markets = ["United Kingdom", "France", "Germany", "Italy", "Spain"]
    periods = ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025", "Q1 2026"]
    oems = [
        ("Honda", "Asian OEMs", 1.00, 0.00),
        ("Volkswagen", "European OEMs", 1.48, -0.10),
        ("Hyundai", "Asian OEMs", 1.08, 0.18),
        ("Kia", "Asian OEMs", 1.00, 0.14),
        ("BMW", "Premium", 0.74, 0.40),
        ("Mercedes-Benz", "Premium", 0.70, 0.34),
        ("Nissan", "Asian OEMs", 0.92, -0.03),
        ("Renault", "European OEMs", 1.04, -0.18),
    ]
    market_base = {
        "United Kingdom": (440_000, 2.62),
        "France": (315_000, 2.08),
        "Germany": (390_000, 2.20),
        "Italy": (275_000, 2.78),
        "Spain": (255_000, 2.30),
    }
    period_lift = {
        "Q1 2025": -0.22,
        "Q2 2025": -0.10,
        "Q3 2025": 0.00,
        "Q4 2025": 0.12,
        "Q1 2026": 0.20,
    }
    rows = []
    for p_i, period in enumerate(periods):
        for m_i, market in enumerate(markets):
            base_visitors, base_conversion = market_base[market]
            for o_i, (oem, grouping, visitor_multiplier, conversion_shift) in enumerate(oems):
                noise = ((p_i + m_i + o_i) % 5 - 2) * 0.035
                visitors = int(base_visitors * visitor_multiplier * (1 + p_i * 0.025 + noise))
                conversion = round(max(1.1, base_conversion + period_lift[period] + conversion_shift + (((m_i + o_i) % 3) - 1) * 0.06), 2)
                contracts = int(round(visitors * conversion / 100))
                rows.append(
                    {
                        "market": market,
                        "period": period,
                        "oem": oem,
                        "oem_grouping": grouping,
                        "visitors": visitors,
                        "contracts": contracts,
                        "website_to_contract_conversion_rate": conversion,
                    }
                )
    return pd.DataFrame(rows)


def ensure_sample_data() -> None:
    DEFAULT_DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not DEFAULT_DATA_PATH.exists():
        generate_sample_data().to_csv(DEFAULT_DATA_PATH, index=False)


def load_data(uploaded_file) -> pd.DataFrame:
    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
    else:
        ensure_sample_data()
        df = pd.read_csv(DEFAULT_DATA_PATH)

    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        st.error(
            "The dataset is missing required columns: " + ", ".join(sorted(missing))
        )
        st.stop()

    df["visitors"] = pd.to_numeric(df["visitors"], errors="coerce").fillna(0)
    df["contracts"] = pd.to_numeric(df["contracts"], errors="coerce").fillna(0)
    df["website_to_contract_conversion_rate"] = pd.to_numeric(
        df["website_to_contract_conversion_rate"], errors="coerce"
    ).fillna(0)
    return df


def fmt_int(value: float) -> str:
    return f"{int(round(value)):,}"


def fmt_pct(value: float) -> str:
    return f"{value:.2f}%"


def period_sort_key(period: str):
    # Handles common labels like Q1 2026 while keeping unknown labels stable.
    parts = str(period).replace("-", " ").split()
    year = next((int(p) for p in parts if p.isdigit() and len(p) == 4), 0)
    quarter = next((int(p[1]) for p in parts if len(p) == 2 and p[0].upper() == "Q" and p[1].isdigit()), 0)
    return (year, quarter, str(period))


def get_periods(df: pd.DataFrame) -> list[str]:
    return sorted(df["period"].dropna().unique().tolist(), key=period_sort_key)


def get_markets(df: pd.DataFrame) -> list[str]:
    return sorted(df["market"].dropna().unique().tolist())


def get_oems(df: pd.DataFrame) -> list[str]:
    return sorted(df["oem"].dropna().unique().tolist())


def calculate_rankings(df: pd.DataFrame, period: str, market: str) -> pd.DataFrame:
    market_df = df[(df["period"] == period) & (df["market"] == market)].copy()
    market_df["rank"] = market_df["website_to_contract_conversion_rate"].rank(
        ascending=False, method="min"
    ).astype(int)
    return market_df.sort_values("rank")


def selected_oem_market_rows(df: pd.DataFrame, selected_oem: str, selected_period: str) -> pd.DataFrame:
    rows = df[(df["oem"] == selected_oem) & (df["period"] == selected_period)].copy()
    if rows.empty:
        return rows
    enriched = []
    periods = get_periods(df)
    previous_period = None
    if selected_period in periods and periods.index(selected_period) > 0:
        previous_period = periods[periods.index(selected_period) - 1]

    for _, row in rows.iterrows():
        ranking_df = calculate_rankings(df, selected_period, row["market"])
        rank = int(ranking_df.loc[ranking_df["oem"] == selected_oem, "rank"].iloc[0])
        market_avg = ranking_df["website_to_contract_conversion_rate"].mean()
        prior_conversion = np.nan
        if previous_period:
            prior_match = df[
                (df["oem"] == selected_oem)
                & (df["period"] == previous_period)
                & (df["market"] == row["market"])
            ]
            if not prior_match.empty:
                prior_conversion = prior_match["website_to_contract_conversion_rate"].iloc[0]
        row_dict = row.to_dict()
        row_dict["rank"] = rank
        row_dict["market_average_conversion"] = market_avg
        row_dict["benchmark_gap"] = row["website_to_contract_conversion_rate"] - market_avg
        row_dict["prior_conversion"] = prior_conversion
        row_dict["movement"] = (
            row["website_to_contract_conversion_rate"] - prior_conversion
            if not pd.isna(prior_conversion)
            else np.nan
        )
        enriched.append(row_dict)
    return pd.DataFrame(enriched)


def render_top_bar():
    st.sidebar.markdown(f"### {APP_NAME}")
    st.sidebar.caption(f"{PRODUCT_VERSION} · OEM-neutral competitor reporting")
    st.sidebar.divider()


def start_here_page(df: pd.DataFrame):
    markets = get_markets(df)
    periods = get_periods(df)
    groups = sorted(df["oem_grouping"].dropna().unique().tolist())

    st.caption(f"{PRODUCT_VERSION} · Start Here")
    st.title(APP_NAME)
    st.write(
        "Benchmark OEM performance across selected markets using one consistent competitor model. "
        "The primary metric is website to contract conversion rate."
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Markets included", len(markets))
    c2.metric("Time periods", len(periods))
    c3.metric("Primary metric", "Website → Contract")

    st.subheader("Report coverage")
    m1, m2, m3 = st.columns([1, 1, 1])
    with m1:
        st.markdown("**Markets included**")
        for market in markets:
            st.markdown(f"- {market}")
    with m2:
        st.markdown("**Time periods included**")
        for period in periods:
            st.markdown(f"- {period}")
    with m3:
        st.markdown("**OEM groupings included**")
        for group in groups:
            st.markdown(f"- {group}")

    st.info(
        "This app has no locked OEMs, no OEM-specific presets, and no privileged benchmark logic. "
        "Any OEM in the dataset can be selected and analysed using the same rules."
    )


def insight_report_page(df: pd.DataFrame):
    st.caption(f"{PRODUCT_VERSION} · Page 2")
    st.title("OEM Insight Report")

    oems = get_oems(df)
    periods = get_periods(df)

    st.subheader("Step-by-step setup")
    s1, s2, s3 = st.columns(3)
    with s1:
        st.markdown("**Step 1: Select the OEM you want to analyse.**")
        selected_oem = st.selectbox("OEM", oems, index=0, key="insight_oem")
    with s2:
        st.markdown("**Step 2: Select the time period you want to review.**")
        selected_period = st.selectbox("Time period", periods, index=len(periods) - 1, key="insight_period")
    with s3:
        st.markdown("**Step 3: Review the generated insight report.**")
        st.write("The report updates from the selected OEM, selected period, and included markets.")
        st.button("Generate insight report", type="primary")

    rows = selected_oem_market_rows(df, selected_oem, selected_period)
    if rows.empty:
        st.warning("No rows found for this OEM and time period.")
        return

    best = rows.loc[rows["website_to_contract_conversion_rate"].idxmax()]
    weakest = rows.loc[rows["website_to_contract_conversion_rate"].idxmin()]
    total_contracts = rows["contracts"].sum()
    total_visitors = rows["visitors"].sum()
    avg_conversion = rows["website_to_contract_conversion_rate"].mean()

    momentum_rows = rows.dropna(subset=["movement"])
    if not momentum_rows.empty:
        strongest_momentum = momentum_rows.loc[momentum_rows["movement"].idxmax()]
        momentum_sentence = (
            f"{strongest_momentum['market']} has the strongest movement versus the prior period "
            f"at {fmt_pct(strongest_momentum['movement'])}."
        )
    else:
        momentum_sentence = "No prior-period movement is available for the selected period."

    st.markdown("---")
    st.subheader("Executive takeaway")
    st.markdown(
        f"""
        <div class="dark-card">
            <p class="small-label">Headline view</p>
            <h3>{selected_oem} across the included markets</h3>
            <p>{selected_oem} shows strongest conversion performance in <b>{best['market']}</b> and weakest performance in <b>{weakest['market']}</b> for <b>{selected_period}</b>. The clearest opportunity is to lift website-to-contract conversion in {weakest['market']}, while protecting the stronger position in {best['market']}. {momentum_sentence}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Average conversion", fmt_pct(avg_conversion))
    k2.metric("Contracts", fmt_int(total_contracts))
    k3.metric("Visitors", fmt_int(total_visitors))
    k4.metric("Best market", best["market"])

    st.subheader("Market performance cards")
    market_cols = st.columns(min(5, len(rows)))
    for idx, (_, row) in enumerate(rows.sort_values("market").iterrows()):
        col = market_cols[idx % len(market_cols)]
        gap = row["benchmark_gap"]
        gap_class = "good" if gap >= 0 else "bad"
        with col:
            st.markdown(
                f"""
                <div class="insight-card">
                    <p class="small-label">{row['market']}</p>
                    <h3>{fmt_pct(row['website_to_contract_conversion_rate'])}</h3>
                    <p><b>Contracts:</b> {fmt_int(row['contracts'])}</p>
                    <p><b>Visitors:</b> {fmt_int(row['visitors'])}</p>
                    <p><b>Rank:</b> {int(row['rank'])} of {len(oems)}</p>
                    <p><b>Benchmark gap:</b> <span class="{gap_class}">{fmt_pct(gap)}</span></p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    chart_rows = rows.sort_values("market")
    st.subheader("Cross-market visualisations")
    v1, v2 = st.columns(2)
    with v1:
        fig = px.bar(
            chart_rows,
            x="market",
            y="website_to_contract_conversion_rate",
            title="Website to contract conversion rate by market",
            labels={
                "market": "Market",
                "website_to_contract_conversion_rate": "Conversion rate (%)",
            },
        )
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with v2:
        trend = (
            df[df["oem"] == selected_oem]
            .groupby("period", as_index=False)["website_to_contract_conversion_rate"]
            .mean()
        )
        trend["period_sort"] = trend["period"].apply(period_sort_key)
        trend = trend.sort_values("period_sort")
        fig = px.line(
            trend,
            x="period",
            y="website_to_contract_conversion_rate",
            markers=True,
            title="Performance movement over time",
            labels={
                "period": "Period",
                "website_to_contract_conversion_rate": "Avg conversion rate (%)",
            },
        )
        st.plotly_chart(fig, use_container_width=True)

    v3, v4 = st.columns(2)
    with v3:
        ranking_table = chart_rows[
            [
                "market",
                "rank",
                "website_to_contract_conversion_rate",
                "market_average_conversion",
                "benchmark_gap",
            ]
        ].rename(
            columns={
                "market": "Market",
                "rank": "Rank",
                "website_to_contract_conversion_rate": "Conversion %",
                "market_average_conversion": "Market avg %",
                "benchmark_gap": "Benchmark gap",
            }
        )
        st.markdown("**Market ranking view**")
        st.dataframe(ranking_table, hide_index=True, use_container_width=True)

    with v4:
        fig = px.scatter(
            chart_rows,
            x="visitors",
            y="website_to_contract_conversion_rate",
            size="contracts",
            hover_name="market",
            title="Opportunity matrix: visitor scale vs conversion",
            labels={
                "visitors": "Visitors",
                "website_to_contract_conversion_rate": "Conversion rate (%)",
                "contracts": "Contracts",
            },
        )
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Generated insight cards")
    insight_cards = [
        (
            "Strongest market",
            f"{best['market']} is the strongest conversion market at {fmt_pct(best['website_to_contract_conversion_rate'])}.",
        ),
        (
            "Weakest market",
            f"{weakest['market']} is the weakest conversion market at {fmt_pct(weakest['website_to_contract_conversion_rate'])}.",
        ),
        (
            "Conversion opportunity",
            f"Prioritise {weakest['market']} for funnel and journey diagnostics because it is the lowest-converting market.",
        ),
        (
            "Scale opportunity",
            f"{best['market']} is a useful playbook market because it combines stronger conversion with {fmt_int(best['visitors'])} visitors.",
        ),
        (
            "Competitive pressure",
            f"The weakest rank is {int(rows['rank'].max())} of {len(oems)}. Investigate the competitor set in that market.",
        ),
        ("Momentum", momentum_sentence),
    ]
    cols = st.columns(3)
    for index, (title, body) in enumerate(insight_cards):
        with cols[index % 3]:
            st.markdown(
                f"""
                <div class="insight-card">
                    <p class="small-label">Insight</p>
                    <h3>{title}</h3>
                    <p>{body}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )


def use_cases_page():
    st.caption(f"{PRODUCT_VERSION} · Generic workflows")
    st.title("Use Cases")
    st.write(
        "Use this report to compare OEM performance across markets, identify conversion opportunities, "
        "and understand where each brand is gaining or losing competitive advantage."
    )

    cols = st.columns(4)
    use_cases = [
        ("Prioritise markets", "Identify where a selected OEM is strongest, weakest, and most exposed."),
        ("Compare competitors", "Benchmark OEMs using consistent market, period, and conversion definitions."),
        ("Improve conversion", "Spot where journey efficiency is below the market benchmark."),
        ("Prepare exec updates", "Convert cross-market performance into a concise action narrative."),
    ]
    for col, (title, body) in zip(cols, use_cases):
        with col:
            st.markdown(
                f"""
                <div class="insight-card">
                    <p class="small-label">Use case</p>
                    <h3>{title}</h3>
                    <p>{body}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
    st.info("This version does not include product-specific organisational toggles, locked OEM cards, or brand-specific presets.")


def market_summary_page(df: pd.DataFrame):
    st.caption(f"{PRODUCT_VERSION} · Market view")
    st.title("Market Summary")

    markets = get_markets(df)
    periods = get_periods(df)
    c1, c2 = st.columns(2)
    with c1:
        market = st.selectbox("Market", markets, index=0, key="market_summary_market")
    with c2:
        period = st.selectbox("Time period", periods, index=len(periods) - 1, key="market_summary_period")

    rows = calculate_rankings(df, period, market)
    top = rows.iloc[0]
    avg_conversion = rows["website_to_contract_conversion_rate"].mean()

    k1, k2, k3 = st.columns(3)
    k1.metric("Top OEM", top["oem"])
    k2.metric("Market avg conversion", fmt_pct(avg_conversion))
    k3.metric("OEMs compared", len(rows))

    left, right = st.columns([1.2, 1])
    with left:
        fig = px.bar(
            rows,
            x="website_to_contract_conversion_rate",
            y="oem",
            orientation="h",
            title=f"OEM conversion performance — {market}, {period}",
            labels={
                "website_to_contract_conversion_rate": "Conversion rate (%)",
                "oem": "OEM",
            },
        )
        fig.update_yaxes(categoryorder="total ascending")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.markdown("**OEM ranking table**")
        table = rows[
            ["rank", "oem", "oem_grouping", "website_to_contract_conversion_rate", "contracts", "visitors"]
        ].rename(
            columns={
                "rank": "Rank",
                "oem": "OEM",
                "oem_grouping": "Grouping",
                "website_to_contract_conversion_rate": "Conversion %",
                "contracts": "Contracts",
                "visitors": "Visitors",
            }
        )
        st.dataframe(table, hide_index=True, use_container_width=True)


def scorecard_page(df: pd.DataFrame):
    st.caption(f"{PRODUCT_VERSION} · OEM comparison")
    st.title("Scorecard")

    oems = get_oems(df)
    periods = get_periods(df)
    c1, c2 = st.columns([1, 2])
    with c1:
        period = st.selectbox("Time period", periods, index=len(periods) - 1, key="scorecard_period")
    with c2:
        selected_oems = st.multiselect("OEMs to compare", oems, default=oems[:3])

    if not selected_oems:
        st.warning("Select at least one OEM to generate the scorecard.")
        return

    rows = df[(df["period"] == period) & (df["oem"].isin(selected_oems))]
    summary = (
        rows.groupby("oem", as_index=False)
        .agg(
            visitors=("visitors", "sum"),
            contracts=("contracts", "sum"),
            avg_conversion=("website_to_contract_conversion_rate", "mean"),
        )
        .sort_values("avg_conversion", ascending=False)
    )

    st.info("No OEM is locked. Select any combination of brands for a like-for-like scorecard.")

    fig = px.bar(
        summary,
        x="oem",
        y="avg_conversion",
        title="Cross-market average conversion by selected OEM",
        labels={"oem": "OEM", "avg_conversion": "Average conversion rate (%)"},
    )
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        summary.rename(
            columns={
                "oem": "OEM",
                "visitors": "Visitors",
                "contracts": "Contracts",
                "avg_conversion": "Average conversion %",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )


def data_dictionary_page(df: pd.DataFrame):
    st.caption(f"{PRODUCT_VERSION} · Data")
    st.title("Data Dictionary & Upload")
    st.write("Upload a CSV using the expected column names below, or use the bundled sample dataset.")
    st.markdown("**Required columns**")
    st.code("\n".join(sorted(REQUIRED_COLUMNS)), language="text")

    with st.expander("Preview current dataset", expanded=True):
        st.dataframe(df.head(50), use_container_width=True)

    buffer = io.StringIO()
    df.head(200).to_csv(buffer, index=False)
    st.download_button(
        "Download current sample/loaded data as CSV",
        buffer.getvalue(),
        file_name="oem_competitor_data_preview.csv",
        mime="text/csv",
    )


def main():
    render_top_bar()

    uploaded_file = st.sidebar.file_uploader("Upload OEM data CSV", type=["csv"])
    df = load_data(uploaded_file)

    pages = [
        "1. Start Here",
        "2. Insight Report",
        "3. Use Cases",
        "4. Market Summary",
        "5. Scorecard",
        "6. Data Dictionary & Upload",
    ]
    page = st.sidebar.radio("Navigation", pages)

    st.sidebar.divider()
    st.sidebar.success("Generic mode active: no locked OEMs or product-specific presets.")

    if page == "1. Start Here":
        start_here_page(df)
    elif page == "2. Insight Report":
        insight_report_page(df)
    elif page == "3. Use Cases":
        use_cases_page()
    elif page == "4. Market Summary":
        market_summary_page(df)
    elif page == "5. Scorecard":
        scorecard_page(df)
    else:
        data_dictionary_page(df)


if __name__ == "__main__":
    main()
