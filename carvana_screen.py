import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

APP_TITLE = "Carvana Live Data Dashboard"
APP_SUBTITLE = (
    "Live monitoring dashboard for Carvana (CVNA) built on official SEC EDGAR filings "
    "and the U.S. Bureau of Labor Statistics used-car CPI series."
)
TICKER = "CVNA"
COMPANY_NAME = "Carvana Co."
SEC_BASE = "https://www.sec.gov"
SEC_DATA_BASE = "https://data.sec.gov"
BLS_SERIES_ID = "CUUR0000SETA02"
DEFAULT_SEC_USER_AGENT = "carvana-dashboard/1.0 contact@example.com"
REQUEST_TIMEOUT = 30

st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon=":bar_chart:")


def get_sec_user_agent() -> str:
    try:
        secret_value = st.secrets.get("SEC_USER_AGENT", "").strip()
    except Exception:
        secret_value = ""
    return secret_value or DEFAULT_SEC_USER_AGENT


def fetch_json(url: str, headers: dict[str, str] | None = None) -> Any:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.json()


def fetch_text(url: str, headers: dict[str, str] | None = None) -> str:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.text


def sec_headers() -> dict[str, str]:
    return {
        "User-Agent": get_sec_user_agent(),
        "Accept-Encoding": "gzip, deflate",
    }


@st.cache_data(ttl=60 * 60 * 12)
def get_company_cik(ticker: str) -> str:
    data = fetch_json(f"{SEC_BASE}/files/company_tickers.json", headers=sec_headers())
    for item in data.values():
        if item["ticker"].upper() == ticker.upper():
            return str(item["cik_str"]).zfill(10)
    raise ValueError(f"Could not find CIK for ticker {ticker}.")


@st.cache_data(ttl=60 * 60)
def get_company_submissions(cik: str) -> dict[str, Any]:
    return fetch_json(f"{SEC_DATA_BASE}/submissions/CIK{cik}.json", headers=sec_headers())


@st.cache_data(ttl=60 * 60 * 12)
def get_company_facts(cik: str) -> dict[str, Any]:
    return fetch_json(f"{SEC_DATA_BASE}/api/xbrl/companyfacts/CIK{cik}.json", headers=sec_headers())


@st.cache_data(ttl=60 * 60 * 12)
def get_used_car_cpi() -> pd.DataFrame:
    payload = fetch_json(f"https://api.bls.gov/publicAPI/v2/timeseries/data/{BLS_SERIES_ID}")
    rows = payload["Results"]["series"][0]["data"]
    records: list[dict[str, Any]] = []
    for row in rows:
        period = row["period"]
        if not period.startswith("M") or period == "M13":
            continue
        month = int(period[1:])
        date = pd.Timestamp(year=int(row["year"]), month=month, day=1) + pd.offsets.MonthEnd(0)
        records.append(
            {
                "Date": date,
                "Index": float(row["value"]),
                "Period": f"{row['periodName']} {row['year']}",
                "Latest": row.get("latest") == "true",
            }
        )
    return pd.DataFrame(records).sort_values("Date").reset_index(drop=True)


def metric_candidates() -> dict[str, list[str]]:
    return {
        "Revenue": [
            "RevenueFromContractWithCustomerExcludingAssessedTax",
            "SalesRevenueNet",
        ],
        "Gross Profit": ["GrossProfit"],
        "Operating Income": ["OperatingIncomeLoss"],
    }


def extract_quarterly_metric(company_facts: dict[str, Any], label: str) -> pd.DataFrame:
    us_gaap = company_facts.get("facts", {}).get("us-gaap", {})
    for tag in metric_candidates()[label]:
        fact = us_gaap.get(tag)
        if not fact:
            continue
        rows = fact.get("units", {}).get("USD", [])
        quarter_records: list[dict[str, Any]] = []
        for row in rows:
            filed = row.get("filed")
            start = row.get("start")
            end = row.get("end")
            value = row.get("val")
            form = row.get("form")
            if not all([filed, start, end, value is not None, form]):
                continue
            duration_days = (pd.Timestamp(end) - pd.Timestamp(start)).days
            is_quarter = form.startswith("10-Q") and 70 <= duration_days <= 105
            is_fiscal_q4 = form.startswith("10-K") and row.get("frame") and "Q4" in row["frame"]
            if not (is_quarter or is_fiscal_q4):
                continue
            quarter_records.append(
                {
                    "Date": pd.Timestamp(end),
                    "Filed": pd.Timestamp(filed),
                    "Metric": label,
                    "Value": float(value),
                    "Form": form,
                    "Fiscal Period": row.get("fp"),
                }
            )
        if quarter_records:
            frame = pd.DataFrame(quarter_records)
            frame = frame.sort_values(["Date", "Filed"]).drop_duplicates(subset=["Date"], keep="last")
            return frame.sort_values("Date").reset_index(drop=True)
    return pd.DataFrame(columns=["Date", "Filed", "Metric", "Value", "Form", "Fiscal Period"])


def build_fundamentals_frame(company_facts: dict[str, Any]) -> pd.DataFrame:
    frames = [extract_quarterly_metric(company_facts, label) for label in metric_candidates()]
    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        return pd.DataFrame(columns=["Date", "Filed", "Metric", "Value", "Form", "Fiscal Period"])
    return pd.concat(frames, ignore_index=True)


def format_billions(value: float) -> str:
    return f"${value / 1_000_000_000:.2f}B"


def build_recent_filings(submissions: dict[str, Any], limit: int = 12) -> pd.DataFrame:
    recent = submissions["filings"]["recent"]
    rows: list[dict[str, Any]] = []
    count = min(limit, len(recent["filingDate"]))
    for idx in range(count):
        accession = recent["accessionNumber"][idx]
        accession_no_dashes = accession.replace("-", "")
        rows.append(
            {
                "Filing Date": pd.Timestamp(recent["filingDate"][idx]),
                "Form": recent["form"][idx],
                "Accession": accession,
                "Document": recent["primaryDocument"][idx],
                "SEC URL": (
                    f"{SEC_BASE}/Archives/edgar/data/{int(submissions['cik'])}/"
                    f"{accession_no_dashes}/{recent['primaryDocument'][idx]}"
                ),
            }
        )
    return pd.DataFrame(rows)


def resolve_form4_xml_url(cik: str, accession: str) -> str | None:
    accession_no_dashes = accession.replace("-", "")
    index_url = f"{SEC_BASE}/Archives/edgar/data/{int(cik)}/{accession_no_dashes}/index.json"
    index_payload = fetch_json(index_url, headers=sec_headers())
    items = index_payload.get("directory", {}).get("item", [])
    xml_files = [
        item["name"]
        for item in items
        if item["name"].endswith(".xml") and "-index" not in item["name"] and not item["name"].startswith("xsl")
    ]
    if not xml_files:
        return None
    return f"{SEC_BASE}/Archives/edgar/data/{int(cik)}/{accession_no_dashes}/{xml_files[0]}"


def text_at(parent: ET.Element | None, path: str) -> str:
    if parent is None:
        return ""
    node = parent.find(path)
    if node is None or node.text is None:
        return ""
    return node.text.strip()


@st.cache_data(ttl=60 * 60)
def get_recent_form4_transactions(cik: str, submissions: dict[str, Any], max_forms: int = 8) -> pd.DataFrame:
    recent = submissions["filings"]["recent"]
    form_indices = [idx for idx, form in enumerate(recent["form"]) if form in {"4", "4/A"}][:max_forms]
    rows: list[dict[str, Any]] = []
    headers = sec_headers()

    for idx in form_indices:
        accession = recent["accessionNumber"][idx]
        filing_date = recent["filingDate"][idx]
        xml_url = resolve_form4_xml_url(cik, accession)
        if not xml_url:
            continue
        xml_text = fetch_text(xml_url, headers=headers)
        root = ET.fromstring(xml_text)
        owner_name = text_at(root, "./reportingOwner/reportingOwnerId/rptOwnerName")
        transactions = root.findall(".//nonDerivativeTransaction")

        for transaction in transactions:
            transaction_date = text_at(transaction, "./transactionDate/value") or filing_date
            code = text_at(transaction, "./transactionCoding/transactionCode")
            shares = text_at(transaction, "./transactionAmounts/transactionShares/value")
            price = text_at(transaction, "./transactionAmounts/transactionPricePerShare/value")
            action = text_at(transaction, "./transactionAmounts/transactionAcquiredDisposedCode/value")
            security = text_at(transaction, "./securityTitle/value")
            share_count = float(shares) if shares else 0.0
            share_price = float(price) if price else 0.0
            signed_shares = share_count if action == "A" else -share_count if action == "D" else 0.0
            rows.append(
                {
                    "Date": pd.Timestamp(transaction_date),
                    "Filing Date": pd.Timestamp(filing_date),
                    "Owner": owner_name or "Unknown owner",
                    "Code": code or "N/A",
                    "Action": "Acquire" if action == "A" else "Dispose" if action == "D" else "Other",
                    "Shares": share_count,
                    "Signed Shares": signed_shares,
                    "Price": share_price,
                    "Estimated Value": share_count * share_price if share_price > 0 else None,
                    "Security": security or "Unknown security",
                    "Accession": accession,
                    "XML URL": xml_url,
                }
            )

    if not rows:
        return pd.DataFrame(
            columns=[
                "Date",
                "Filing Date",
                "Owner",
                "Code",
                "Action",
                "Shares",
                "Signed Shares",
                "Price",
                "Estimated Value",
                "Security",
                "Accession",
                "XML URL",
            ]
        )

    return pd.DataFrame(rows).sort_values(["Date", "Filing Date"], ascending=[True, True]).reset_index(drop=True)


def load_dashboard_data() -> dict[str, Any]:
    cik = get_company_cik(TICKER)
    submissions = get_company_submissions(cik)
    company_facts = get_company_facts(cik)
    fundamentals = build_fundamentals_frame(company_facts)
    filings = build_recent_filings(submissions)
    insider_trades = get_recent_form4_transactions(cik, submissions)
    used_car_cpi = get_used_car_cpi()
    return {
        "cik": cik,
        "submissions": submissions,
        "fundamentals": fundamentals,
        "filings": filings,
        "insider_trades": insider_trades,
        "used_car_cpi": used_car_cpi,
    }


def latest_metric_value(fundamentals: pd.DataFrame, metric: str) -> float | None:
    metric_frame = fundamentals[fundamentals["Metric"] == metric].sort_values("Date")
    if metric_frame.empty:
        return None
    return float(metric_frame.iloc[-1]["Value"])


st.title(APP_TITLE)
st.caption(APP_SUBTITLE)

if get_sec_user_agent() == DEFAULT_SEC_USER_AGENT:
    st.info(
        "For production deployment, set `SEC_USER_AGENT` in `.streamlit/secrets.toml` or "
        "Streamlit Community Cloud secrets to a real contact string such as "
        "`your-app-name/1.0 your-email@example.com`."
    )

with st.spinner("Loading live SEC and BLS data..."):
    try:
        dashboard = load_dashboard_data()
        load_error = None
    except Exception as exc:
        dashboard = None
        load_error = exc

if load_error:
    st.error(f"Live data load failed: {load_error}")
    st.stop()

fundamentals_df = dashboard["fundamentals"]
filings_df = dashboard["filings"]
insider_df = dashboard["insider_trades"]
used_car_df = dashboard["used_car_cpi"]

latest_revenue = latest_metric_value(fundamentals_df, "Revenue")
latest_gross_profit = latest_metric_value(fundamentals_df, "Gross Profit")
latest_operating_income = latest_metric_value(fundamentals_df, "Operating Income")
recent_form4_count = int((filings_df["Form"] == "4").sum() + (filings_df["Form"] == "4/A").sum())
latest_cpi = float(used_car_df.iloc[-1]["Index"])
prior_cpi = float(used_car_df.iloc[-2]["Index"])

metric_1, metric_2, metric_3, metric_4 = st.columns(4)
metric_1.metric(
    "Latest quarterly revenue",
    format_billions(latest_revenue) if latest_revenue is not None else "N/A",
)
metric_2.metric(
    "Latest gross profit",
    format_billions(latest_gross_profit) if latest_gross_profit is not None else "N/A",
)
metric_3.metric(
    "Latest operating income",
    format_billions(latest_operating_income) if latest_operating_income is not None else "N/A",
)
metric_4.metric(
    "Used-car CPI",
    f"{latest_cpi:.3f}",
    delta=f"{latest_cpi - prior_cpi:+.3f} MoM",
)

st.subheader("Live Signal Summary")
signal_1, signal_2, signal_3 = st.columns(3)

with signal_1:
    if latest_operating_income is not None and latest_operating_income > 0:
        st.success(f"Operating income is positive at {format_billions(latest_operating_income)}.")
    else:
        value_text = format_billions(latest_operating_income) if latest_operating_income is not None else "N/A"
        st.warning(f"Operating income is negative or unavailable at {value_text}.")

with signal_2:
    if recent_form4_count >= 3:
        st.warning(f"Recent issuer filing history includes {recent_form4_count} Form 4 filings.")
    else:
        st.info(f"Recent issuer filing history includes {recent_form4_count} Form 4 filings.")

with signal_3:
    cpi_change = latest_cpi - prior_cpi
    if cpi_change > 0:
        st.success(f"Used-car CPI rose by {cpi_change:.3f} month over month.")
    else:
        st.error(f"Used-car CPI fell by {abs(cpi_change):.3f} month over month.")

st.markdown("---")

tab_1, tab_2, tab_3, tab_4 = st.tabs(
    [
        "SEC fundamentals",
        "Recent filings",
        "Insider activity",
        "Used-car market",
    ]
)

with tab_1:
    st.markdown("### Quarterly fundamentals from SEC company facts")
    st.write(
        "Revenue, gross profit, and operating income are pulled from SEC XBRL company facts for "
        "Carvana and filtered to quarterly periods."
    )
    fundamentals_chart = px.line(
        fundamentals_df,
        x="Date",
        y="Value",
        color="Metric",
        markers=True,
        title="Quarterly fundamentals",
    )
    fundamentals_chart.update_yaxes(title="USD")
    fundamentals_chart.update_layout(hovermode="x unified")
    st.plotly_chart(fundamentals_chart, use_container_width=True)

    display_fundamentals = fundamentals_df.copy()
    display_fundamentals["Value"] = display_fundamentals["Value"].map(format_billions)
    display_fundamentals["Date"] = display_fundamentals["Date"].dt.date
    display_fundamentals["Filed"] = display_fundamentals["Filed"].dt.date
    st.dataframe(
        display_fundamentals.sort_values(["Date", "Metric"], ascending=[False, True]),
        use_container_width=True,
        hide_index=True,
    )

with tab_2:
    st.markdown("### Recent SEC filing feed")
    st.write(
        "These filings come directly from Carvana's EDGAR submissions feed. The table includes the "
        "latest issuer filings and direct SEC document links."
    )
    filings_display = filings_df.copy()
    filings_display["Filing Date"] = filings_display["Filing Date"].dt.date
    st.dataframe(filings_display, use_container_width=True, hide_index=True)

    form_mix = filings_df["Form"].value_counts().reset_index()
    form_mix.columns = ["Form", "Count"]
    filings_bar = px.bar(form_mix, x="Form", y="Count", title="Recent filing mix")
    st.plotly_chart(filings_bar, use_container_width=True)

with tab_3:
    st.markdown("### Recent Form 4 transactions")
    st.write(
        "Form 4 filings are pulled from EDGAR and parsed from the raw XML ownership documents. "
        "Share counts are signed positive for acquisitions and negative for disposals."
    )
    if insider_df.empty:
        st.info("No recent Form 4 transactions were parsed from the issuer submissions feed.")
    else:
        insider_chart = px.bar(
            insider_df,
            x="Date",
            y="Signed Shares",
            color="Action",
            hover_data=["Owner", "Code", "Price", "Estimated Value"],
            title="Signed insider share activity",
        )
        insider_chart.update_yaxes(title="Signed shares")
        st.plotly_chart(insider_chart, use_container_width=True)

        insider_display = insider_df.copy()
        insider_display["Date"] = insider_display["Date"].dt.date
        insider_display["Filing Date"] = insider_display["Filing Date"].dt.date
        insider_display["Estimated Value"] = insider_display["Estimated Value"].map(
            lambda x: f"${x:,.0f}" if pd.notna(x) else "N/A"
        )
        insider_display["Price"] = insider_display["Price"].map(lambda x: f"${x:,.2f}" if x else "N/A")
        insider_display["Shares"] = insider_display["Shares"].map(lambda x: f"{x:,.0f}")
        insider_display["Signed Shares"] = insider_display["Signed Shares"].map(lambda x: f"{x:,.0f}")
        st.dataframe(
            insider_display[
                [
                    "Date",
                    "Owner",
                    "Action",
                    "Code",
                    "Shares",
                    "Signed Shares",
                    "Price",
                    "Estimated Value",
                    "XML URL",
                ]
            ].sort_values("Date", ascending=False),
            use_container_width=True,
            hide_index=True,
        )

with tab_4:
    st.markdown("### U.S. used cars and trucks CPI")
    st.write(
        "This macro backdrop uses the BLS CPI series `CUUR0000SETA02` for used cars and trucks. "
        "It is not the Manheim index, but it is an official public proxy for used-vehicle pricing."
    )
    used_car_chart = go.Figure()
    used_car_chart.add_trace(
        go.Scatter(
            x=used_car_df["Date"],
            y=used_car_df["Index"],
            mode="lines+markers",
            name="Used cars and trucks CPI",
            line=dict(color="#17a2b8", width=3),
        )
    )
    used_car_chart.update_layout(
        title="BLS used cars and trucks CPI",
        yaxis_title="Index",
        hovermode="x unified",
    )
    st.plotly_chart(used_car_chart, use_container_width=True)

    used_car_display = used_car_df.copy()
    used_car_display["Date"] = used_car_display["Date"].dt.date
    st.dataframe(used_car_display.sort_values("Date", ascending=False), use_container_width=True, hide_index=True)

with st.sidebar:
    st.header("Source Notes")
    st.markdown(f"**Company:** {COMPANY_NAME}")
    st.markdown(f"**Ticker:** {TICKER}")
    st.markdown(f"**CIK:** {dashboard['cik']}")
    st.markdown(f"**Data refresh time:** {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    st.markdown("---")
    st.markdown("**Live sources in this app:**")
    st.markdown("- SEC EDGAR `company_tickers.json`")
    st.markdown("- SEC EDGAR submissions JSON")
    st.markdown("- SEC XBRL company facts JSON")
    st.markdown("- SEC Form 4 raw XML filings")
    st.markdown("- BLS Public Data API v2")

    st.markdown("---")
    st.subheader("Interpretation")
    st.write(
        "This version is live-data driven, but it still uses simple heuristics. A production-grade "
        "credit dashboard would add bond market data, ABS trust-level performance, source timestamps, "
        "and stronger validation around filing parsing."
    )

    with st.expander("Deployment secrets"):
        st.code('SEC_USER_AGENT = "carvana-dashboard/1.0 your-email@example.com"', language="toml")
