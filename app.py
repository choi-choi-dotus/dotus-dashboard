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

# ── 다크 테마 CSS ────────────────────────────────────────
st.markdown("""
<style>
/* 전체 배경 */
.stApp {
    background-color: #0d1117;
}
/* 메인 콘텐츠 */
[data-testid="stAppViewContainer"] > .main {
    background-color: #0d1117;
}
/* 사이드바 */
[data-testid="stSidebar"] {
    background-color: #161b22;
    border-right: 1px solid #30363d;
}
[data-testid="stSidebar"] * {
    color: #c9d1d9 !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #8b949e !important;
    font-size: 0.75rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
[data-testid="stSidebar"] h2 {
    color: #f0f6fc !important;
    font-size: 1rem !important;
}
[data-testid="stSidebar"] h3 {
    color: #8b949e !important;
    font-size: 0.72rem !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
/* 멀티셀렉트 태그 */
[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: #1f6feb !important;
}
/* 일반 텍스트 */
p, span, label, div {
    color: #c9d1d9;
}
h1, h2, h3 {
    color: #f0f6fc !important;
}
/* KPI 카드 */
[data-testid="stMetric"] {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 20px 24px;
}
[data-testid="stMetricValue"] {
    font-size: 1.8rem !important;
    font-weight: 700 !important;
    color: #f0f6fc !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.82rem !important;
    color: #8b949e !important;
}
[data-testid="stMetricDelta"] {
    font-size: 0.85rem !important;
}
/* 구분선 */
hr {
    border-color: #30363d !important;
}
/* 데이터프레임 */
[data-testid="stDataFrame"] {
    border: 1px solid #30363d;
    border-radius: 8px;
}
/* 버튼 */
.stButton button {
    background-color: #21262d !important;
    color: #c9d1d9 !important;
    border: 1px solid #30363d !important;
}
.stButton button:hover {
    border-color: #58a6ff !important;
    color: #58a6ff !important;
}
/* 섹션 제목 */
.section-title {
    color: #f0f6fc;
    font-size: 1rem;
    font-weight: 600;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 1px solid #30363d;
}
</style>
""", unsafe_allow_html=True)

# ── 차트 기본 레이아웃 (다크) ────────────────────────────
DARK_LAYOUT = dict(
    plot_bgcolor="#0d1117",
    paper_bgcolor="#161b22",
    font=dict(color="#c9d1d9", size=12),
    xaxis=dict(gridcolor="#21262d", linecolor="#30363d", tickcolor="#8b949e"),
    yaxis=dict(gridcolor="#21262d", linecolor="#30363d", tickcolor="#8b949e"),
    margin=dict(t=40, b=20, l=10, r=10),
)

# ── 색상 팔레트 ──────────────────────────────────────────
C_BLUE   = "#58a6ff"
C_CYAN   = "#79c0ff"
C_GREEN  = "#3fb950"
C_ORANGE = "#f79520"
C_RED    = "#ff7b72"
C_PURPLE = "#d2a8ff"
PIE_COLORS = [C_BLUE, C_CYAN, C_GREEN, C_ORANGE, C_RED, C_PURPLE,
              "#ffa657", "#a5d6ff", "#7ee787", "#ff9492"]

# ── 비밀번호 로그인 ───────────────────────────────────────
PASSWORD = "dotus2026"

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True

    st.markdown("""
    <div style='display:flex;justify-content:center;margin-top:80px;'>
        <div style='text-align:center;'>
            <h1 style='font-size:2rem;font-weight:700;margin-bottom:4px;color:#f0f6fc;'>📊 Dotus</h1>
            <p style='color:#8b949e;margin-bottom:32px;'>매출 대시보드</p>
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

    ws1 = wb["시트1"]
    rows1 = list(ws1.iter_rows(values_only=True))
    headers1 = rows1[0][:16]
    data1 = [r[:16] for r in rows1[1:] if r[0] is not None]
    df_sales = pd.DataFrame(data1, columns=headers1)
    df_sales["date"] = pd.to_datetime(df_sales["date"])
    df_sales["year_month"] = df_sales["date"].dt.to_period("M").astype(str)
    df_sales["net_sales"] = df_sales["sales_amount"] - df_sales["return_sales_amount"]
    df_sales["net_qty"] = df_sales["quantity"] - df_sales["return_quantity"]

    ws2 = wb["시트2"]
    rows2 = list(ws2.iter_rows(values_only=True))
    headers2 = rows2[0][:13]
    data2 = [r[:13] for r in rows2[1:] if r[0] is not None]
    df_stock = pd.DataFrame(data2, columns=headers2)
    df_stock["date"] = pd.to_datetime(df_stock["date"])
    df_stock["year_month"] = df_stock["date"].dt.to_period("M").astype(str)

    ws3 = wb["품목마스터"]
    rows3 = list(ws3.iter_rows(values_only=True))
    headers3 = rows3[0][:8]
    data3 = [r[:8] for r in rows3[1:] if r[0] is not None]
    df_master = pd.DataFrame(data3, columns=headers3)

    return df_sales, df_stock, df_master

df_sales, df_stock, df_master = load_data()

HAS_SMALL = "small_category" in df_sales.columns

# ── 사이드바 ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dotus 대시보드")
    st.markdown("---")

    st.markdown("### 🗓 기간")
    months = sorted(df_sales["year_month"].unique())
    month_start = st.selectbox("시작 월", months, index=0)
    month_end = st.selectbox("종료 월", months, index=len(months)-1)

    st.markdown("---")
    st.markdown("### 🏪 채널")
    all_clients = sorted(df_sales["client"].unique())
    selected_clients = st.multiselect("채널 선택", options=all_clients, default=all_clients)

    st.markdown("---")
    st.markdown("### 📦 카테고리")

    all_large = sorted(df_sales["large_category"].dropna().unique())
    selected_large = st.multiselect("대분류", options=all_large, default=all_large)

    mid_pool = df_sales[df_sales["large_category"].isin(selected_large)]["medium_category"].dropna().unique()
    selected_medium = st.multiselect("중분류", options=sorted(mid_pool), default=sorted(mid_pool))

    selected_small = None
    if HAS_SMALL:
        small_pool = df_sales[
            df_sales["large_category"].isin(selected_large) &
            df_sales["medium_category"].isin(selected_medium)
        ]["small_category"].dropna().unique()
        selected_small = st.multiselect("소분류", options=sorted(small_pool), default=sorted(small_pool))

    st.markdown("---")
    if st.button("🚪 로그아웃", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# ── 필터링 ────────────────────────────────────────────────
filtered = df_sales[
    (df_sales["year_month"] >= month_start) &
    (df_sales["year_month"] <= month_end) &
    (df_sales["client"].isin(selected_clients)) &
    (df_sales["large_category"].isin(selected_large)) &
    (df_sales["medium_category"].isin(selected_medium))
]
if HAS_SMALL and selected_small:
    filtered = filtered[filtered["small_category"].isin(selected_small)]

prev_months = sorted(df_sales["year_month"].unique())
if month_start in prev_months and prev_months.index(month_start) > 0:
    prev_start = prev_months[max(0, prev_months.index(month_start) - 1)]
    prev_end   = prev_months[max(0, prev_months.index(month_end) - 1)]
    prev_filtered = df_sales[
        (df_sales["year_month"] >= prev_start) &
        (df_sales["year_month"] <= prev_end) &
        (df_sales["client"].isin(selected_clients)) &
        (df_sales["large_category"].isin(selected_large))
    ]
else:
    prev_filtered = pd.DataFrame()

# ── KPI ───────────────────────────────────────────────────
total_sales   = filtered["net_sales"].sum()
total_qty     = filtered["net_qty"].sum()
total_orders  = filtered["order_no_1"].nunique()
avg_order_val = total_sales / total_orders if total_orders > 0 else 0

if not prev_filtered.empty:
    prev_sales  = prev_filtered["net_sales"].sum()
    prev_qty    = prev_filtered["net_qty"].sum()
    prev_orders = prev_filtered["order_no_1"].nunique()
    sales_delta  = f"{((total_sales  - prev_sales)  / prev_sales  * 100):+.1f}%" if prev_sales  else "N/A"
    qty_delta    = f"{((total_qty    - prev_qty)    / prev_qty    * 100):+.1f}%" if prev_qty    else "N/A"
    orders_delta = f"{((total_orders - prev_orders) / prev_orders * 100):+.1f}%" if prev_orders else "N/A"
else:
    sales_delta = qty_delta = orders_delta = None

# ── 헤더 ──────────────────────────────────────────────────
st.markdown("# 📊 Dotus 매출 대시보드")
cat_info = f"대분류 {len(selected_large)}개"
if selected_medium: cat_info += f" · 중분류 {len(selected_medium)}개"
if HAS_SMALL and selected_small: cat_info += f" · 소분류 {len(selected_small)}개"
st.markdown(f"<span style='color:#8b949e;font-size:0.9rem;'>조회 기간: **{month_start} ~ {month_end}** &nbsp;|&nbsp; 채널: **{len(selected_clients)}개** &nbsp;|&nbsp; {cat_info}</span>", unsafe_allow_html=True)
st.markdown("---")

# ── KPI 카드 ──────────────────────────────────────────────
k1, k2, k3, k4 = st.columns(4)
with k1: st.metric("💰 순매출",      f"₩{total_sales/1e8:.1f}억",    delta=sales_delta)
with k2: st.metric("📦 판매 수량",   f"{total_qty:,}개",              delta=qty_delta)
with k3: st.metric("🛒 주문 건수",   f"{total_orders:,}건",           delta=orders_delta)
with k4: st.metric("💳 평균 주문금액", f"₩{avg_order_val:,.0f}")

st.markdown("---")

# ── 월별 매출 추이 (에어리어 + 라인) ─────────────────────
st.markdown("### 📈 월별 매출 추이")

monthly = filtered.groupby("year_month").agg(
    순매출=("net_sales", "sum"),
    수량=("net_qty", "sum"),
    주문수=("order_no_1", "nunique")
).reset_index()
monthly["MoM증가율"] = monthly["순매출"].pct_change() * 100

fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])

# 에어리어 (그라데이션 느낌)
fig_monthly.add_trace(go.Scatter(
    x=monthly["year_month"],
    y=monthly["순매출"],
    name="순매출",
    mode="lines+markers",
    fill="tozeroy",
    fillcolor="rgba(88,166,255,0.15)",
    line=dict(color=C_BLUE, width=3),
    marker=dict(size=7, color=C_BLUE),
    text=[f"₩{v/1e8:.1f}억" for v in monthly["순매출"]],
    hovertemplate="%{text}<extra></extra>"
), secondary_y=False)

# MoM 라인
fig_monthly.add_trace(go.Scatter(
    x=monthly["year_month"],
    y=monthly["MoM증가율"],
    name="MoM 증가율(%)",
    mode="lines+markers",
    line=dict(color=C_ORANGE, width=2, dash="dot"),
    marker=dict(size=6, color=C_ORANGE),
    hovertemplate="%{y:.1f}%<extra></extra>"
), secondary_y=True)

fig_monthly.update_layout(
    height=360,
    **DARK_LAYOUT,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#c9d1d9")),
)
fig_monthly.update_yaxes(title_text="순매출 (원)", secondary_y=False, tickformat=",", gridcolor="#21262d", linecolor="#30363d")
fig_monthly.update_yaxes(title_text="MoM 증가율 (%)", secondary_y=True, gridcolor="#21262d", linecolor="#30363d")
st.plotly_chart(fig_monthly, use_container_width=True)

# ── 채널별 + 중분류 파이 ──────────────────────────────────
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("### 🏪 채널별 매출")
    channel_df = filtered.groupby("client").agg(
        순매출=("net_sales", "sum")
    ).sort_values("순매출", ascending=True).reset_index()

    fig_ch = go.Figure(go.Bar(
        x=channel_df["순매출"],
        y=channel_df["client"],
        orientation="h",
        marker=dict(
            color=channel_df["순매출"],
            colorscale=[[0, "#1f3a5f"], [1, C_BLUE]],
            showscale=False
        ),
        text=[f"₩{v/1e8:.1f}억" for v in channel_df["순매출"]],
        textposition="outside",
        textfont=dict(color="#c9d1d9")
    ))
    fig_ch.update_layout(
        height=400,
        **DARK_LAYOUT,
        xaxis=dict(tickformat=",", gridcolor="#21262d", linecolor="#30363d"),
        yaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
        margin=dict(t=20, b=20, l=10, r=80)
    )
    st.plotly_chart(fig_ch, use_container_width=True)

with col_right:
    st.markdown("### 📦 상품 중분류별 매출")
    mid_df = filtered.groupby("medium_category").agg(
        순매출=("net_sales", "sum")
    ).sort_values("순매출", ascending=False).head(10).reset_index()

    fig_mid = go.Figure(go.Pie(
        values=mid_df["순매출"],
        labels=mid_df["medium_category"],
        marker=dict(colors=PIE_COLORS, line=dict(color="#0d1117", width=2)),
        textposition="inside",
        textinfo="percent+label",
        textfont=dict(color="#0d1117", size=11),
        hole=0.35
    ))
    fig_mid.update_layout(
        height=400,
        plot_bgcolor="#0d1117",
        paper_bgcolor="#161b22",
        font=dict(color="#c9d1d9"),
        margin=dict(t=20, b=20, l=10, r=10),
        showlegend=False
    )
    st.plotly_chart(fig_mid, use_container_width=True)

# ── 채널별 월별 매출 추이 (멀티라인) ─────────────────────
st.markdown("### 🌊 채널별 월별 매출 흐름")

channel_monthly = filtered.groupby(["year_month", "client"])["net_sales"].sum().reset_index()
top_clients = channel_df.sort_values("순매출", ascending=False).head(6)["client"].tolist()

fig_multi = go.Figure()
line_colors = [C_BLUE, C_CYAN, C_GREEN, C_ORANGE, C_RED, C_PURPLE]

for i, client in enumerate(top_clients):
    df_c = channel_monthly[channel_monthly["client"] == client]
    color = line_colors[i % len(line_colors)]
    fig_multi.add_trace(go.Scatter(
        x=df_c["year_month"],
        y=df_c["net_sales"],
        name=client,
        mode="lines+markers",
        fill="tozeroy",
        fillcolor=f"rgba({int(color[1:3],16)},{int(color[3:5],16)},{int(color[5:7],16)},0.07)",
        line=dict(color=color, width=2.5),
        marker=dict(size=6, color=color),
        hovertemplate=f"{client}: ₩%{{y:,.0f}}<extra></extra>"
    ))

fig_multi.update_layout(
    height=380,
    **DARK_LAYOUT,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#c9d1d9")),
    yaxis=dict(tickformat=",", gridcolor="#21262d", linecolor="#30363d"),
    xaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
)
st.plotly_chart(fig_multi, use_container_width=True)

# ── 히트맵 ────────────────────────────────────────────────
st.markdown("### 🗺 채널 × 월별 매출 히트맵")

pivot_df = filtered.groupby(["client", "year_month"])["net_sales"].sum().reset_index()
pivot_table = pivot_df.pivot(index="client", columns="year_month", values="net_sales").fillna(0)

fig_heat = go.Figure(go.Heatmap(
    z=pivot_table.values / 1e8,
    x=pivot_table.columns.tolist(),
    y=pivot_table.index.tolist(),
    colorscale=[[0, "#0d1117"], [0.5, "#1f6feb"], [1, "#58a6ff"]],
    text=[[f"₩{v:.1f}억" for v in row] for row in pivot_table.values / 1e8],
    texttemplate="%{text}",
    textfont=dict(color="#f0f6fc", size=11),
    colorbar=dict(title="억원", tickfont=dict(color="#c9d1d9"), titlefont=dict(color="#c9d1d9"))
))
fig_heat.update_layout(
    height=380,
    plot_bgcolor="#0d1117",
    paper_bgcolor="#161b22",
    font=dict(color="#c9d1d9"),
    margin=dict(t=20, b=20, l=10, r=10),
    xaxis=dict(tickfont=dict(color="#8b949e")),
    yaxis=dict(tickfont=dict(color="#8b949e"))
)
st.plotly_chart(fig_heat, use_container_width=True)

# ── 입고 vs 매출 ──────────────────────────────────────────
st.markdown("---")
st.markdown("### 🚚 입고 vs 매출 비교 (월별)")

stock_monthly = df_stock.groupby("year_month")["Total Amount (KRW)"].sum().reset_index()
stock_monthly.columns = ["year_month", "입고금액"]

sales_monthly = filtered.groupby("year_month")["net_sales"].sum().reset_index()
sales_monthly.columns = ["year_month", "매출금액"]

compare = pd.merge(sales_monthly, stock_monthly, on="year_month", how="outer").fillna(0)
compare = compare[(compare["year_month"] >= month_start) & (compare["year_month"] <= month_end)]

fig_compare = go.Figure()
fig_compare.add_trace(go.Scatter(
    x=compare["year_month"], y=compare["매출금액"],
    name="매출",
    mode="lines+markers",
    fill="tozeroy",
    fillcolor="rgba(88,166,255,0.12)",
    line=dict(color=C_BLUE, width=3),
    marker=dict(size=7),
    hovertemplate="매출: ₩%{y:,.0f}<extra></extra>"
))
fig_compare.add_trace(go.Scatter(
    x=compare["year_month"], y=compare["입고금액"],
    name="입고",
    mode="lines+markers",
    fill="tozeroy",
    fillcolor="rgba(247,149,32,0.10)",
    line=dict(color=C_ORANGE, width=3),
    marker=dict(size=7),
    hovertemplate="입고: ₩%{y:,.0f}<extra></extra>"
))
fig_compare.update_layout(
    height=340,
    **DARK_LAYOUT,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#c9d1d9")),
    yaxis=dict(tickformat=",", gridcolor="#21262d", linecolor="#30363d"),
    xaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
)
st.plotly_chart(fig_compare, use_container_width=True)

# ── 상세 테이블 ───────────────────────────────────────────
st.markdown("---")
st.markdown("### 📋 월별 채널 상세")

detail = filtered.groupby(["year_month", "client"]).agg(
    순매출=("net_sales", "sum"),
    판매수량=("net_qty", "sum"),
    주문건수=("order_no_1", "nunique"),
    반품수량=("return_quantity", "sum"),
    반품금액=("return_sales_amount", "sum")
).reset_index()
detail["순매출"]  = detail["순매출"].apply(lambda x: f"₩{x:,}")
detail["반품금액"] = detail["반품금액"].apply(lambda x: f"₩{x:,}")
detail.columns = ["월", "채널", "순매출", "판매수량", "주문건수", "반품수량", "반품금액"]

st.dataframe(detail, use_container_width=True, hide_index=True)

st.markdown("---")
st.caption("Dotus Dashboard · 데이터 기준: 제공된 엑셀 파일")
