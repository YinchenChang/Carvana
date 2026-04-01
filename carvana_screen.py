import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# ==========================================
# 1. 頁面基本設定與樣式
# ==========================================
st.set_page_config(page_title="CVNA 空頭部位監控儀表板", layout="wide", page_icon="📉")
st.title("📉 Carvana (CVNA) 核心風險監控儀表板")
st.markdown("本儀表板用於追蹤 Carvana 的放空催化劑，包含 ABS 違約率、債務流動性、內部人交易與總經數據。")

# ==========================================
# 2. 數據獲取模組 (實務上這裡替換為 API 串接或爬蟲)
# ==========================================
@st.cache_data
def load_abs_data():
    """模擬 Form 10-D 的 61天以上逾期率數據 (Vintage Analysis)"""
    dates = pd.date_range(start='2025-01-01', periods=15, freq='M')
    # 模擬 2024年信託 與 2025年新信託的違約率差異
    df = pd.DataFrame({
        'Date': dates,
        'CRVNA_2024_1_Trust (%)': np.linspace(1.2, 1.8, 15) + np.random.normal(0, 0.05, 15),
        'CRVNA_2025_4_Trust (%)': np.linspace(1.0, 2.8, 15) + np.random.normal(0, 0.1, 15) # 新貸款違約率陡升
    })
    return df

@st.cache_data
def load_bond_data():
    """模擬 FINRA TRACE 公司債 YTM 與 股價對比"""
    dates = pd.date_range(start='2025-01-01', periods=60, freq='W')
    df = pd.DataFrame({
        'Date': dates,
        'CVNA_2030_Bond_YTM (%)': np.linspace(7.5, 11.5, 60) + np.random.normal(0, 0.2, 60),
        'Stock_Price ($)': np.linspace(140, 314, 60) + np.random.normal(0, 10, 60)
    })
    return df

@st.cache_data
def load_insider_data():
    """模擬 Form 4 內部人套現數據"""
    dates = pd.date_range(start='2025-10-01', periods=20, freq='W')
    df = pd.DataFrame({
        'Date': dates,
        'Garcia_Family_Sells (Million $)': np.random.exponential(scale=15, size=20)
    })
    return df

@st.cache_data
def load_macro_data():
    """模擬曼海姆二手車指數"""
    dates = pd.date_range(start='2024-01-01', periods=27, freq='M')
    df = pd.DataFrame({
        'Date': dates,
        'Manheim_Index': np.linspace(210, 185, 27) + np.random.normal(0, 2, 27)
    })
    return df

# 載入數據
df_abs = load_abs_data()
df_bond = load_bond_data()
df_insider = load_insider_data()
df_macro = load_macro_data()

# ==========================================
# 3. 頂部狀態指示器 (警示系統)
# ==========================================
st.subheader("🚨 即時風險狀態 (Triggers)")
col1, col2, col3 = st.columns(3)

# 邏輯判斷：如果最新一期 2025-4 信託逾期率大於 2.5%，亮紅燈
latest_delinquency = df_abs['CRVNA_2025_4_Trust (%)'].iloc[-1]
with col1:
    if latest_delinquency > 2.5:
        st.error(f"🔴 ABS 違約警報: 2025-4 逾期率達 {latest_delinquency:.2f}%")
    else:
        st.success(f"🟢 ABS 違約正常: {latest_delinquency:.2f}%")

# 邏輯判斷：如果債券 YTM 大於 10%，亮紅燈
latest_ytm = df_bond['CVNA_2030_Bond_YTM (%)'].iloc[-1]
with col2:
    if latest_ytm > 10.0:
        st.error(f"🔴 流動性警報: 2030公司債 YTM 飆至 {latest_ytm:.2f}%")
    else:
        st.success(f"🟢 債券 YTM 穩定: {latest_ytm:.2f}%")

# 邏輯判斷：內部人近期總拋售量
recent_sells = df_insider['Garcia_Family_Sells (Million $)'].tail(4).sum()
with col3:
    if recent_sells > 50:
        st.warning(f"🟡 內部人套現: 近四周倒貨達 ${recent_sells:.1f}M")
    else:
        st.info(f"🔵 內部人交易平穩")

st.markdown("---")

# ==========================================
# 4. 視覺化圖表區塊 (使用 Tabs 分類模組)
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["1. ABS 貸款質量", "2. 債務與流動性", "3. 內部人治理", "4. 總經環境"])

with tab1:
    st.markdown("### 61天以上逾期率 (帳齡分析 Vintage Analysis)")
    st.markdown("追蹤新舊貸款池在相同月份下的違約惡化速度。**(數據來源: Form 10-D)**")
    fig_abs = go.Figure()
    fig_abs.add_trace(go.Scatter(x=df_abs['Date'], y=df_abs['CRVNA_2024_1_Trust (%)'], mode='lines', name='2024-1 信託 (舊)'))
    fig_abs.add_trace(go.Scatter(x=df_abs['Date'], y=df_abs['CRVNA_2025_4_Trust (%)'], mode='lines', name='2025-4 信託 (新)', line=dict(color='red', width=3)))
    fig_abs.add_hline(y=2.0, line_dash="dash", line_color="orange", annotation_text="歷史平均安全閥值 (2.0%)")
    fig_abs.update_layout(yaxis_title="61+ Days Delinquency Rate (%)", hovermode="x unified")
    st.plotly_chart(fig_abs, use_container_width=True)

with tab2:
    st.markdown("### 股價 vs. 公司債殖利率 (YTM)")
    st.markdown("尋找股債背離的甜蜜點。若股價創高但 YTM 飆升，代表債市正在反映龐大的違約與再融資風險。")
    fig_bond = go.Figure()
    fig_bond.add_trace(go.Scatter(x=df_bond['Date'], y=df_bond['Stock_Price ($)'], mode='lines', name='CVNA 股價', yaxis='y1'))
    fig_bond.add_trace(go.Scatter(x=df_bond['Date'], y=df_bond['CVNA_2030_Bond_YTM (%)'], mode='lines', name='2030 公司債 YTM', line=dict(color='red'), yaxis='y2'))
    fig_bond.update_layout(
        yaxis=dict(title="股價 (USD)"),
        yaxis2=dict(title="YTM (%)", overlaying='y', side='right'),
        hovermode="x unified"
    )
    st.plotly_chart(fig_bond, use_container_width=True)

with tab3:
    st.markdown("### Garcia 家族內部人拋售追蹤")
    st.markdown("觀察創辦人是否在極端高點加速套現。**(數據來源: Form 4)**")
    fig_insider = px.bar(df_insider, x='Date', y='Garcia_Family_Sells (Million $)', title="內部人每週套現金額 (百萬美元)")
    fig_insider.update_traces(marker_color='orange')
    st.plotly_chart(fig_insider, use_container_width=True)

with tab4:
    st.markdown("### 曼海姆二手車價值指數 (Manheim Index)")
    st.markdown("衡量庫存跌價損失風險與抵押品殘值。指數崩跌將直接侵蝕 Carvana 毛利。")
    fig_macro = px.line(df_macro, x='Date', y='Manheim_Index', title="美國二手車批發價格指數走勢")
    fig_macro.update_traces(line_color='teal')
    st.plotly_chart(fig_macro, use_container_width=True)

# ==========================================
# 5. 側邊欄：策略執行輔助
# ==========================================
with st.sidebar:
    st.header("⚙️ 策略執行參數")
    st.markdown(f"**目前評估股價:** $314.38")
    st.markdown(f"**風險預算上限:** $75,000 (5%)")
    st.markdown("---")
    st.subheader("建議行動方針")
    if latest_delinquency > 2.5 and latest_ytm > 10.0:
        st.error("💥 **強烈放空訊號觸發！**\n\n指標共振，建議立即執行【策略 A: 空頭賣權價差】建立核心部位。")
    else:
        st.info("👀 **觀望中...**\n\n核心風險指標尚未完全共振，請維持監控，避免盲目迎戰軋空情緒。")