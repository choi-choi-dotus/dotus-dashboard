import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
 
# ── 페이지 설정 ──────────────────────────────────────────
st.set_page_config(
    page_title="Dotus 매출 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)
 
# ── 글로벌 CSS ────────────────────────────────────────────
st.markdown("""
<style>
/* 사이드바 */
[data-testid="stSidebar"] {
    background-color: #0f172a;
}
[data-testid="stSidebar"] * {
    color: #e2e8f0 !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #94a3b8 !important;
    font-size: 0.78rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
/* 멀티셀렉트 태그 */
[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: #1e40af !important;
}
/* KPI 메트릭 */
[data-testid="stMetric"] {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px 24px;
}
[data-testid="stMetricValue"] {
    font-size: 1.7rem !important;
    font-weight: 700 !important;
    color: #0f172a !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.82rem !important;
    color: #64748b !important;
}
/* 메인 헤더 */
h1 { color: #0f172a !important; }
/* 구분선 */
hr { border-color: #e2e8f0 !important; }
</style>
""", unsafe_allow_html=True)
 
# ── 비밀번호 로그인 ───────────────────────────────────────
PASSWORD = "dotus2026"
 
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
 
    if st.session_state.authenticated:
        return True
 
    st.markdown("""
    <div style='display:flex; justify-content:center; margin-top:80px;'>
        <div style='text-align:center;'>
            <h1 style='font-size:2rem; font-weight:700; margin-bottom:4px;'>📊 Dotus</h1>
            <p style='color:#888; margin-bottom:32px;'>매출 대시보드</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        pw = st.text_input("비밀번호를 입력하세요", type="password", placeholder="비밀번호")
        if st.button("로그인", use_container_width=True, type="primary"):
            if pw == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("비밀번호가 틀렸습니다.")
    return False
 
if not check_password():
    st.stop()
 
# ── 데이터 로드 ───────────────────────────────────────────
@st.cache_data
def load_data():
    wb = openpyxl.load_workbook("data.xlsx", read_only=True, data_only=True)
 
    # 결제데이터
    ws1 = wb["시트1"]
    rows1 = list(ws1.iter_rows(values_only=True))
    headers1 = rows1[0][:16]
    data1 = [r[:16] for r in rows1[1:] if r[0] is not None]
    df_sales = pd.DataFrame(data1, columns=headers1)
    df_sales["date"] = pd.to_datetime(df_sales["date"])
    df_sales["year_month"] = df_sales["date"].dt.to_period("M").astype(str)
    df_sales["net_sales"] = df_sales["sales_amount"] - df_sales["return_sales_amount"]
    df_sales["net_qty"] = df_sales["quantity"] - df_sales["return_quantity"]
 
    # 입고데이터
    ws2 = wb["시트2"]
    rows2 = list(ws2.iter_rows(values_only=True))
    headers2 = rows2[0][:13]
    data2 = [r[:13] for r in rows2[1:] if r[0] is not None]
    df_stock = pd.DataFrame(data2, columns=headers2)
    df_stock["date"] = pd.to_datetime(df_stock["date"])
    df_stock["year_month"] = df_stock["date"].dt.to_period("M").astype(str)
 
    # 품목마스터
    ws3 = wb["품목마스터"]
    rows3 = list(ws3.iter_rows(values_only=True))
    headers3 = rows3[0][:8]
    data3 = [r[:8] for r in rows3[1:] if r[0] is not None]
    df_master = pd.DataFrame(data3, columns=headers3)
 
    return df_sales, df_stock, df_master
 
df_sales, df_stock, df_master = load_data()
 
# 소분류 컬럼 존재 여부 확인
HAS_SMALL = "small_category" in df_sales.columns
 
# ── 사이드바 필터 ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dotus 대시보드")
    st.markdown("---")
 
    # 기간 필터
    st.markdown("### 🗓 기간")
    months = sorted(df_sales["year_month"].unique())
    month_start = st.selectbox("시작 월", months, index=0)
    month_end = st.selectbox("종료 월", months, index=len(months)-1)
 
    st.markdown("---")
 
    # 채널 필터
    st.markdown("### 🏪 채널")
    all_clients = sorted(df_sales["client"].unique())
    selected_clients = st.multiselect(
        "채널 선택",
        options=all_clients,
        default=all_clients
    )
 
    st.markdown("---")
 
    # ── 카테고리 필터 (연동) ──────────────────────────────
    st.markdown("### 📦 카테고리")
 
    # 대분류
    all_large = sorted(df_sales["large_category"].dropna().unique())
    selected_large = st.multiselect("대분류", options=all_large, default=all_large)
 
    # 중분류 (선택된 대분류에 맞게 필터)
    mid_pool = df_sales[df_sales["large_category"].isin(selected_large)]["medium_category"].dropna().unique()
    all_medium = sorted(mid_pool)
    selected_medium = st.multiselect("중분류", options=all_medium, default=all_medium)
 
    # 소분류 (컬럼 존재 시)
    selected_small = None
    if HAS_SMALL:
        small_pool = df_sales[
            df_sales["large_category"].isin(selected_large) &
            df_sales["medium_category"].isin(selected_medium)
        ]["small_category"].dropna().unique()
        all_small = sorted(small_pool)
        selected_small = st.multiselect("소분류", options=all_small, default=all_small)
 
    st.markdown("---")
    if st.button("🚪 로그아웃", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
 
# ── 데이터 필터링 ─────────────────────────────────────────
filtered = df_sales[
    (df_sales["year_month"] >= month_start) &
    (df_sales["year_month"] <= month_end) &
    (df_sales["client"].isin(selected_clients)) &
    (df_sales["large_category"].isin(selected_large)) &
    (df_sales["medium_category"].isin(selected_medium))
]
if HAS_SMALL and selected_small:
    filtered = filtered[filtered["small_category"].isin(selected_small)]
 
# 전 기간 (증가율 계산용)
prev_months = sorted(df_sales["year_month"].unique())
if month_start in prev_months and prev_months.index(month_start) > 0:
    prev_start_idx = prev_months.index(month_start) - 1
    prev_end_idx   = prev_months.index(month_end) - 1 if prev_months.index(month_end) > 0 else 0
    prev_start = prev_months[max(0, prev_start_idx)]
    prev_end   = prev_months[max(0, prev_end_idx)]
    prev_filtered = df_sales[
        (df_sales["year_month"] >= prev_start) &
        (df_sales["year_month"] <= prev_end) &
        (df_sales["client"].isin(selected_clients)) &
        (df_sales["large_category"].isin(selected_large))
    ]
else:
    prev_filtered = pd.DataFrame()
 
# ── KPI 계산 ──────────────────────────────────────────────
total_sales   = filtered["net_sales"].sum()
total_qty     = filtered["net_qty"].sum()
total_orders  = filtered["order_no_1"].nunique()
avg_order_val = total_sales / total_orders if total_orders > 0 else 0
 
if not prev_filtered.empty:
    prev_sales  = prev_filtered["net_sales"].sum()
    prev_qty    = prev_filtered["net_qty"].sum()
    prev_orders = prev_filtered["order_no_1"].nunique()
    sales_delta = f"{((total_sales - prev_sales) / prev_sales * 100):+.1f}%" if prev_sales else "N/A"
    qty_delta   = f"{((total_qty - prev_qty) / prev_qty * 100):+.1f}%"       if prev_qty   else "N/A"
    orders_delta= f"{((total_orders - prev_orders) / prev_orders * 100):+.1f}%" if prev_orders else "N/A"
else:
    sales_delta = qty_delta = orders_delta = None
 
# ── 헤더 ──────────────────────────────────────────────────
st.markdown("# 📊 Dotus 매출 대시보드")
cat_info = f"대분류 {len(selected_large)}개"
if selected_medium: cat_info += f" · 중분류 {len(selected_medium)}개"
if HAS_SMALL and selected_small: cat_info += f" · 소분류 {len(selected_small)}개"
st.markdown(f"**조회 기간:** {month_start} ~ {month_end}　|　**채널:** {len(selected_clients)}개　|　**카테고리:** {cat_info}")
st.markdown("---")
 
# ── KPI 카드 ──────────────────────────────────────────────
k1, k2, k3, k4 = st.columns(4)
 
with k1:
    st.metric("💰 순매출", f"₩{total_sales/1e8:.1f}억", delta=sales_delta)
with k2:
    st.metric("📦 판매 수량", f"{total_qty:,}개", delta=qty_delta)
with k3:
    st.metric("🛒 주문 건수", f"{total_orders:,}건", delta=orders_delta)
with k4:
    st.metric("💳 평균 주문금액", f"₩{avg_order_val:,.0f}")
 
st.markdown("---")
 
# ── 월별 매출 추이 ────────────────────────────────────────
st.markdown("### 📈 월별 매출 추이")
 
monthly = filtered.groupby("year_month").agg(
    순매출=("net_sales", "sum"),
    수량=("net_qty", "sum"),
    주문수=("order_no_1", "nunique")
).reset_index()
 
monthly["MoM증가율"] = monthly["순매출"].pct_change() * 100
 
fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])
fig_monthly.add_trace(
    go.Bar(
        x=monthly["year_month"],
        y=monthly["순매출"],
        name="순매출",
        marker_color="#1D4ED8",
        text=[f"₩{v/1e8:.1f}억" for v in monthly["순매출"]],
        textposition="outside"
    ),
    secondary_y=False
)
fig_monthly.add_trace(
    go.Scatter(
        x=monthly["year_month"],
        y=monthly["MoM증가율"],
        name="MoM 증가율(%)",
        mode="lines+markers",
        line=dict(color="#F97316", width=2.5),
        marker=dict(size=8, color="#F97316")
    ),
    secondary_y=True
)
fig_monthly.update_layout(
    height=380,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    margin=dict(t=40, b=20)
)
fig_monthly.update_yaxes(title_text="순매출 (원)", secondary_y=False, tickformat=",")
fig_monthly.update_yaxes(title_text="MoM 증가율 (%)", secondary_y=True)
st.plotly_chart(fig_monthly, use_container_width=True)
 
# ── 채널별 + 카테고리별 ───────────────────────────────────
col_left, col_right = st.columns(2)
 
with col_left:
    st.markdown("### 🏪 채널별 매출")
    channel_df = filtered.groupby("client").agg(
        순매출=("net_sales", "sum"),
        수량=("net_qty", "sum")
    ).sort_values("순매출", ascending=True).reset_index()
 
    fig_ch = go.Figure(go.Bar(
        x=channel_df["순매출"],
        y=channel_df["client"],
        orientation="h",
        marker_color="#1D4ED8",
        text=[f"₩{v/1e8:.1f}억" for v in channel_df["순매출"]],
        textposition="outside"
    ))
    fig_ch.update_layout(
        height=400,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=20, b=20, l=20, r=80),
        xaxis=dict(tickformat=",")
    )
    st.plotly_chart(fig_ch, use_container_width=True)
 
with col_right:
    st.markdown("### 📦 상품 중분류별 매출")
    mid_df = filtered.groupby("medium_category").agg(
        순매출=("net_sales", "sum")
    ).sort_values("순매출", ascending=False).head(10).reset_index()
 
    COLORS = ["#1D4ED8", "#2563EB", "#3B82F6", "#0EA5E9", "#06B6D4",
              "#0891B2", "#0284C7", "#1E40AF", "#1E3A8A", "#172554"]
    fig_mid = px.pie(
        mid_df,
        values="순매출",
        names="medium_category",
        color_discrete_sequence=COLORS
    )
    fig_mid.update_traces(textposition="inside", textinfo="percent+label")
    fig_mid.update_layout(
        height=400,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=20, b=20),
        showlegend=False
    )
    st.plotly_chart(fig_mid, use_container_width=True)
 
# ── 채널별 월별 히트맵 ────────────────────────────────────
st.markdown("### 🗺 채널 × 월별 매출 히트맵")
 
pivot_df = filtered.groupby(["client", "year_month"])["net_sales"].sum().reset_index()
pivot_table = pivot_df.pivot(index="client", columns="year_month", values="net_sales").fillna(0)
 
fig_heat = go.Figure(go.Heatmap(
    z=pivot_table.values / 1e8,
    x=pivot_table.columns.tolist(),
    y=pivot_table.index.tolist(),
    colorscale="Blues",
    text=[[f"₩{v:.1f}억" for v in row] for row in pivot_table.values / 1e8],
    texttemplate="%{text}",
    colorbar=dict(title="억원")
))
fig_heat.update_layout(
    height=380,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    margin=dict(t=20, b=20)
)
st.plotly_chart(fig_heat, use_container_width=True)
 
# ── 입고 vs 매출 비교 ─────────────────────────────────────
st.markdown("---")
st.markdown("### 🚚 입고 vs 매출 비교 (월별)")
 
stock_monthly = df_stock.groupby("year_month")["Total Amount (KRW)"].sum().reset_index()
stock_monthly.columns = ["year_month", "입고금액"]
 
sales_monthly = filtered.groupby("year_month")["net_sales"].sum().reset_index()
sales_monthly.columns = ["year_month", "매출금액"]
 
compare = pd.merge(sales_monthly, stock_monthly, on="year_month", how="outer").fillna(0)
compare = compare[
    (compare["year_month"] >= month_start) &
    (compare["year_month"] <= month_end)
]
 
fig_compare = go.Figure()
fig_compare.add_trace(go.Bar(
    x=compare["year_month"], y=compare["매출금액"],
    name="매출", marker_color="#1D4ED8",
    text=[f"₩{v/1e8:.1f}억" for v in compare["매출금액"]],
    textposition="outside"
))
fig_compare.add_trace(go.Bar(
    x=compare["year_month"], y=compare["입고금액"],
    name="입고", marker_color="#F59E0B",
    text=[f"₩{v/1e8:.1f}억" for v in compare["입고금액"]],
    textposition="outside"
))
fig_compare.update_layout(
    barmode="group",
    height=360,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    margin=dict(t=40, b=20),
    yaxis=dict(tickformat=",")
)
st.plotly_chart(fig_compare, use_container_width=True)
 
# ── 상세 데이터 테이블 ────────────────────────────────────
st.markdown("---")
st.markdown("### 📋 월별 채널 상세")
 
detail = filtered.groupby(["year_month", "client"]).agg(
    순매출=("net_sales", "sum"),
    판매수량=("net_qty", "sum"),
    주문건수=("order_no_1", "nunique"),
    반품수량=("return_quantity", "sum"),
    반품금액=("return_sales_amount", "sum")
).reset_index()
detail["순매출"] = detail["순매출"].apply(lambda x: f"₩{x:,}")
detail["반품금액"] = detail["반품금액"].apply(lambda x: f"₩{x:,}")
detail.columns = ["월", "채널", "순매출", "판매수량", "주문건수", "반품수량", "반품금액"]
 
st.dataframe(detail, use_container_width=True, hide_index=True)
 
st.markdown("---")
st.caption("Dotus Dashboard · 데이터 기준: 제공된 엑셀 파일")
 
