import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

APP_TITLE = "Carvana Credit Dashboard"
APP_SUBTITLE = (
    "Illustrative monitoring dashboard for Carvana (CVNA) credit, funding, and used-car "
    "market indicators. All figures below are synthetic sample data for demo purposes."
)
SEED = 7


st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon=":chart_with_upwards_trend:")


@st.cache_data
def load_abs_data(seed: int = SEED) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    dates = pd.date_range(start="2025-01-31", periods=15, freq="ME")
    trust_2024 = np.linspace(1.15, 1.75, len(dates)) + rng.normal(0, 0.04, len(dates))
    trust_2025 = np.linspace(0.95, 2.85, len(dates)) + rng.normal(0, 0.08, len(dates))
    return pd.DataFrame(
        {
            "Date": dates,
            "CRVNA 2024-1 Trust (%)": np.clip(trust_2024, 0, None),
            "CRVNA 2025-4 Trust (%)": np.clip(trust_2025, 0, None),
        }
    )


@st.cache_data
def load_bond_data(seed: int = SEED) -> pd.DataFrame:
    rng = np.random.default_rng(seed + 1)
    dates = pd.date_range(start="2025-01-03", periods=60, freq="W-FRI")
    ytm = np.linspace(7.4, 11.4, len(dates)) + rng.normal(0, 0.18, len(dates))
    stock = np.linspace(140, 314, len(dates)) + rng.normal(0, 9, len(dates))
    return pd.DataFrame(
        {
            "Date": dates,
            "CVNA 2030 Bond YTM (%)": np.clip(ytm, 0, None),
            "Stock Price ($)": np.clip(stock, 0, None),
        }
    )


@st.cache_data
def load_insider_data(seed: int = SEED) -> pd.DataFrame:
    rng = np.random.default_rng(seed + 2)
    dates = pd.date_range(start="2025-10-03", periods=20, freq="W-FRI")
    sells = rng.exponential(scale=15, size=len(dates))
    return pd.DataFrame(
        {
            "Date": dates,
            "Garcia Family Sales ($M)": sells,
        }
    )


@st.cache_data
def load_macro_data(seed: int = SEED) -> pd.DataFrame:
    rng = np.random.default_rng(seed + 3)
    dates = pd.date_range(start="2024-01-31", periods=27, freq="ME")
    manheim = np.linspace(210, 185, len(dates)) + rng.normal(0, 1.8, len(dates))
    return pd.DataFrame(
        {
            "Date": dates,
            "Manheim Index": manheim,
        }
    )


def render_signal_box(column, condition: bool, good_text: str, bad_text: str) -> None:
    with column:
        if condition:
            st.error(bad_text)
        else:
            st.success(good_text)


abs_df = load_abs_data()
bond_df = load_bond_data()
insider_df = load_insider_data()
macro_df = load_macro_data()

latest_delinquency = abs_df["CRVNA 2025-4 Trust (%)"].iloc[-1]
latest_ytm = bond_df["CVNA 2030 Bond YTM (%)"].iloc[-1]
recent_insider_sales = insider_df["Garcia Family Sales ($M)"].tail(4).sum()
latest_manheim = macro_df["Manheim Index"].iloc[-1]

st.title(APP_TITLE)
st.caption(APP_SUBTITLE)
st.warning(
    "This app uses deterministic synthetic data to demonstrate layout and monitoring logic. "
    "Do not use it as an investment or credit decision tool without wiring in real data sources."
)

metric_1, metric_2, metric_3, metric_4 = st.columns(4)
metric_1.metric("Latest 60+ day delinquency", f"{latest_delinquency:.2f}%", delta=f"{latest_delinquency - 2.5:+.2f} vs trigger")
metric_2.metric("2030 bond YTM", f"{latest_ytm:.2f}%", delta=f"{latest_ytm - 10.0:+.2f} vs trigger")
metric_3.metric("Last 4 weeks insider sales", f"${recent_insider_sales:.1f}M", delta=f"{recent_insider_sales - 50:+.1f} vs trigger")
metric_4.metric("Latest Manheim index", f"{latest_manheim:.1f}")

st.subheader("Trigger Summary")
trigger_1, trigger_2, trigger_3 = st.columns(3)
render_signal_box(
    trigger_1,
    latest_delinquency > 2.5,
    f"ABS performance remains below the 2.5% watch level at {latest_delinquency:.2f}%.",
    f"ABS performance breached the 2.5% watch level at {latest_delinquency:.2f}%.",
)
render_signal_box(
    trigger_2,
    latest_ytm > 10.0,
    f"Funding markets remain constructive with bond YTM at {latest_ytm:.2f}%.",
    f"Funding stress is elevated with bond YTM at {latest_ytm:.2f}%.",
)
with trigger_3:
    if recent_insider_sales > 50:
        st.warning(f"Insider selling is elevated: ${recent_insider_sales:.1f}M over the last 4 weeks.")
    else:
        st.info(f"Insider selling is moderate: ${recent_insider_sales:.1f}M over the last 4 weeks.")

st.markdown("---")

abs_tab, bond_tab, insider_tab, macro_tab = st.tabs(
    [
        "ABS performance",
        "Bond vs. equity",
        "Insider activity",
        "Used-car market",
    ]
)

with abs_tab:
    st.markdown("### 60+ day delinquency trend")
    st.write(
        "Track recent trust-level delinquencies as a leading stress signal for Carvana's "
        "asset-backed securitizations."
    )
    abs_fig = go.Figure()
    abs_fig.add_trace(
        go.Scatter(
            x=abs_df["Date"],
            y=abs_df["CRVNA 2024-1 Trust (%)"],
            mode="lines",
            name="CRVNA 2024-1",
            line=dict(color="#1f77b4", width=2),
        )
    )
    abs_fig.add_trace(
        go.Scatter(
            x=abs_df["Date"],
            y=abs_df["CRVNA 2025-4 Trust (%)"],
            mode="lines",
            name="CRVNA 2025-4",
            line=dict(color="#d62728", width=3),
        )
    )
    abs_fig.add_hline(
        y=2.0,
        line_dash="dash",
        line_color="#ff7f0e",
        annotation_text="Watch level: 2.0%",
    )
    abs_fig.update_layout(yaxis_title="Delinquency rate (%)", hovermode="x unified")
    st.plotly_chart(abs_fig, use_container_width=True)

with bond_tab:
    st.markdown("### Equity momentum vs. credit stress")
    st.write(
        "A widening bond yield alongside a strong stock can indicate that equity and credit "
        "markets disagree on the durability of the recovery."
    )
    bond_fig = go.Figure()
    bond_fig.add_trace(
        go.Scatter(
            x=bond_df["Date"],
            y=bond_df["Stock Price ($)"],
            mode="lines",
            name="CVNA stock",
            yaxis="y1",
            line=dict(color="#2ca02c", width=2.5),
        )
    )
    bond_fig.add_trace(
        go.Scatter(
            x=bond_df["Date"],
            y=bond_df["CVNA 2030 Bond YTM (%)"],
            mode="lines",
            name="2030 bond YTM",
            yaxis="y2",
            line=dict(color="#d62728", width=2.5),
        )
    )
    bond_fig.update_layout(
        yaxis=dict(title="Stock price (USD)"),
        yaxis2=dict(title="YTM (%)", overlaying="y", side="right"),
        hovermode="x unified",
    )
    st.plotly_chart(bond_fig, use_container_width=True)

with insider_tab:
    st.markdown("### Insider selling activity")
    st.write(
        "Insider sales are not a standalone signal, but heavy selling during optimistic market "
        "periods can matter for risk framing."
    )
    insider_fig = px.bar(
        insider_df,
        x="Date",
        y="Garcia Family Sales ($M)",
        title="Weekly Garcia family stock sales",
        color_discrete_sequence=["#ff7f0e"],
    )
    insider_fig.update_layout(showlegend=False)
    st.plotly_chart(insider_fig, use_container_width=True)

with macro_tab:
    st.markdown("### Used-car pricing backdrop")
    st.write(
        "The Manheim index is a useful backdrop for unit economics, inventory values, and future "
        "gross profit normalization."
    )
    macro_fig = px.line(
        macro_df,
        x="Date",
        y="Manheim Index",
        title="Manheim used-vehicle value index",
    )
    macro_fig.update_traces(line_color="#17a2b8", line_width=3)
    st.plotly_chart(macro_fig, use_container_width=True)

with st.sidebar:
    st.header("Dashboard Notes")
    st.write("This public version is a static demo intended for layout and deployment validation.")
    st.markdown("**Current stock price assumption:** $314.38")
    st.markdown("**Illustrative position size:** $75,000 (5% of portfolio)")
    st.markdown("**Risk triggers:**")
    st.markdown("- ABS 60+ day delinquency above 2.5%")
    st.markdown("- Bond YTM above 10.0%")
    st.markdown("- 4-week insider selling above $50M")

    st.markdown("---")
    st.subheader("Interpretation")
    if latest_delinquency > 2.5 and latest_ytm > 10.0:
        st.error(
            "Credit and funding indicators are both flashing red. In a real production setup, "
            "this would justify a deeper review of ABS performance, refinancing risk, and "
            "liquidity runway."
        )
    elif latest_delinquency > 2.5 or latest_ytm > 10.0:
        st.warning(
            "One major credit signal is under pressure. In a real production setup, this would "
            "justify tighter monitoring rather than a full thesis break on its own."
        )
    else:
        st.success(
            "The current synthetic scenario is constructive. In a real production setup, that would "
            "still need confirmation from live ABS, bond, and operating data."
        )

    with st.expander("How to productionize this app"):
        st.markdown(
            "- Replace synthetic loaders with real data pipelines.\n"
            "- Move credentials into Streamlit secrets.\n"
            "- Add data freshness timestamps and source links.\n"
            "- Add error handling for missing or delayed data."
        )
