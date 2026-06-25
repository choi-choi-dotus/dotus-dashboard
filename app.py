import streamlit as st
import pandas as pd
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

# ── Google Material CSS ──────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;600;700&family=Roboto:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Google Sans', 'Roboto', sans-serif;
}
.stApp {
    background-color: #F8F9FA;
}
[data-testid="stAppViewContainer"] > .main {
    background-color: #F8F9FA;
}

/* 사이드바 */
[data-testid="stSidebar"] {
    background-color: #FFFFFF;
    border-right: 1px solid #DADCE0;
}
[data-testid="stSidebar"] * { color: #202124 !important; }
[data-testid="stSidebar"] h2 {
    color: #202124 !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
}
[data-testid="stSidebar"] h3 {
    color: #5F6368 !important;
    font-size: 0.72rem !important;
    font-weight: 500 !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #5F6368 !important;
    font-size: 0.75rem !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}
[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: #E8F0FE !important;
    color: #1A73E8 !important;
}
[data-testid="stSidebar"] span[data-baseweb="tag"] span {
    color: #1A73E8 !important;
}

/* 일반 텍스트 */
p, span, label, div { color: #202124; }
h1 { color: #202124 !important; font-weight: 600 !important; }
h2, h3, h4 { color: #3C4043 !important; font-weight: 500 !important; }
hr { border-color: #DADCE0 !important; }

/* KPI 카드 */
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #DADCE0;
    border-radius: 8px;
    padding: 20px 24px;
    box-shadow: 0 1px 3px rgba(60,64,67,0.08);
}
[data-testid="stMetricValue"] {
    font-size: 1.75rem !important;
    font-weight: 600 !important;
    color: #202124 !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.8rem !important;
    color: #5F6368 !important;
    font-weight: 500 !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

/* 탭 */
[data-testid="stTabs"] button {
    color: #5F6368 !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #1A73E8 !important;
    border-bottom-color: #1A73E8 !important;
}

/* 버튼 */
.stButton button {
    background-color: #FFFFFF !important;
    color: #1A73E8 !important;
    border: 1px solid #DADCE0 !important;
    border-radius: 4px !important;
    font-weight: 500 !important;
}
.stButton button:hover {
    background-color: #E8F0FE !important;
    border-color: #1A73E8 !important;
}

/* 라디오 */
.stRadio label { color: #3C4043 !important; font-size: 0.88rem !important; }

/* 데이터프레임 */
[data-testid="stDataFrame"] {
    border: 1px solid #DADCE0;
    border-radius: 8px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)

# ── 색상 (모노톤 블루/그레이) ────────────────────────────
C_BLUE    = "#1A73E8"
C_BLUE2   = "#1557B0"
C_BLUE3   = "#4285F4"
C_BLUE4   = "#8AB4F8"
C_GRAY    = "#5F6368"
C_GRAY2   = "#9AA0A6"
CHANNEL_COLORS = [C_BLUE, C_BLUE2, C_BLUE3, C_BLUE4, C_GRAY, C_GRAY2]
PIE_COLORS = ["#1A73E8","#4285F4","#8AB4F8","#AECBFA","#E8F0FE",
              "#5F6368","#9AA0A6","#BDC1C6","#1557B0","#185ABC"]

def glay(**kwargs):
    """Google Material light layout"""
    base = dict(
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        font=dict(color="#3C4043", size=12),
        margin=dict(t=40, b=20, l=10, r=10),
    )
    base.update(kwargs)
    return base

AXIS = dict(gridcolor="#F1F3F4", linecolor="#DADCE0",
            tickcolor="#9AA0A6", tickfont=dict(color="#5F6368"))

# ── 비밀번호 ──────────────────────────────────────────────
PASSWORD = "dotus2026"

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True

    st.markdown("""
    <div style='display:flex;justify-content:center;margin-top:80px;'>
        <div style='text-align:center;'>
            <h1 style='font-size:2rem;font-weight:600;color:#202124;margin-bottom:4px;'>📊 Dotus</h1>
            <p style='color:#5F6368;margin-bottom:32px;font-size:0.95rem;'>매출 대시보드</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        pw = st.text_input("비밀번호", type="password", placeholder="비밀번호를 입력하세요")
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

# ── 공통 필터 ─────────────────────────────────────────────
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

# KPI
total_sales   = filtered["net_sales"].sum()
total_qty     = filtered["net_qty"].sum()
total_orders  = filtered["order_no_1"].nunique()
avg_order_val = total_sales / total_orders if total_orders > 0 else 0

if not prev_filtered.empty:
    ps = prev_filtered["net_sales"].sum()
    pq = prev_filtered["net_qty"].sum()
    po = prev_filtered["order_no_1"].nunique()
    sales_delta  = f"{((total_sales - ps) / ps * 100):+.1f}%"   if ps else "N/A"
    qty_delta    = f"{((total_qty   - pq) / pq * 100):+.1f}%"   if pq else "N/A"
    orders_delta = f"{((total_orders- po) / po * 100):+.1f}%"   if po else "N/A"
else:
    sales_delta = qty_delta = orders_delta = None

# ── 헤더 ──────────────────────────────────────────────────
st.markdown("# 📊 Dotus 매출 대시보드")
cat_info = f"대분류 {len(selected_large)}개"
if selected_medium: cat_info += f" · 중분류 {len(selected_medium)}개"
if HAS_SMALL and selected_small: cat_info += f" · 소분류 {len(selected_small)}개"
st.markdown(
    f"<span style='color:#5F6368;font-size:0.88rem;'>"
    f"조회 기간: <b>{month_start} ~ {month_end}</b> &nbsp;|&nbsp; "
    f"채널: <b>{len(selected_clients)}개</b> &nbsp;|&nbsp; {cat_info}</span>",
    unsafe_allow_html=True
)
st.markdown("---")

# ════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📈 매출 대시보드", "🔍 상세 데이터 조회"])

# ════════════════════════════════════════════════════════
# TAB 1
# ════════════════════════════════════════════════════════
with tab1:

    # KPI
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("순매출",       f"₩{total_sales/1e8:.1f}억",  delta=sales_delta)
    with k2: st.metric("판매 수량",    f"{total_qty:,}개",            delta=qty_delta)
    with k3: st.metric("주문 건수",    f"{total_orders:,}건",         delta=orders_delta)
    with k4: st.metric("평균 주문금액", f"₩{avg_order_val:,.0f}")
    st.markdown("---")

    # ── 월별 매출 추이 (금액/수량 토글) ──────────────────
    st.markdown("### 월별 매출 추이")
    sales_metric = st.radio("기준", ["금액", "수량"], horizontal=True, key="sales_metric")

    monthly = filtered.groupby("year_month").agg(
        순매출=("net_sales", "sum"),
        수량=("net_qty", "sum"),
    ).reset_index()
    monthly["MoM"] = (monthly["순매출" if sales_metric == "금액" else "수량"].pct_change() * 100)

    y_col    = "순매출" if sales_metric == "금액" else "수량"
    y_fmt    = [f"₩{v/1e8:.1f}억" for v in monthly[y_col]] if sales_metric == "금액" else [f"{v:,}개" for v in monthly[y_col]]
    y_title  = "순매출 (원)" if sales_metric == "금액" else "판매 수량 (개)"

    fig_sales = make_subplots(specs=[[{"secondary_y": True}]])
    fig_sales.add_trace(go.Scatter(
        x=monthly["year_month"], y=monthly[y_col],
        name=y_col, mode="lines+markers",
        line=dict(color=C_BLUE, width=3, shape="spline"),
        marker=dict(size=7, color=C_BLUE),
        fill="tozeroy", fillcolor="rgba(26,115,232,0.08)",
        text=y_fmt, hovertemplate="%{text}<extra></extra>"
    ), secondary_y=False)
    fig_sales.add_trace(go.Scatter(
        x=monthly["year_month"], y=monthly["MoM"],
        name="MoM 증가율(%)", mode="lines+markers",
        line=dict(color=C_GRAY, width=1.5, dash="dot", shape="spline"),
        marker=dict(size=5, color=C_GRAY),
        hovertemplate="%{y:.1f}%<extra></extra>"
    ), secondary_y=True)
    fig_sales.update_layout(
        height=360,
        **glay(legend=dict(orientation="h", yanchor="bottom", y=1.02))
    )
    fig_sales.update_xaxes(type="category", **AXIS)
    fig_sales.update_yaxes(title_text=y_title, secondary_y=False,
                           tickformat="," if sales_metric == "수량" else None, **AXIS)
    fig_sales.update_yaxes(title_text="MoM (%)", secondary_y=True, **AXIS)
    st.plotly_chart(fig_sales, use_container_width=True)

    # ── 채널별 + 중분류 파이 ──────────────────────────────
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("### 채널별 매출")
        channel_df = filtered.groupby("client").agg(
            순매출=("net_sales", "sum")
        ).sort_values("순매출", ascending=True).reset_index()

        fig_ch = go.Figure(go.Bar(
            x=channel_df["순매출"], y=channel_df["client"],
            orientation="h",
            marker=dict(
                color=channel_df["순매출"],
                colorscale=[[0, "#E8F0FE"], [1, C_BLUE]],
                showscale=False
            ),
            text=[f"₩{v/1e8:.1f}억" for v in channel_df["순매출"]],
            textposition="outside", textfont=dict(color="#3C4043")
        ))
        fig_ch.update_layout(height=400, **glay(margin=dict(t=20, b=20, l=10, r=80)))
        fig_ch.update_xaxes(tickformat=",", **AXIS)
        fig_ch.update_yaxes(**AXIS)
        st.plotly_chart(fig_ch, use_container_width=True)

    with col_right:
        st.markdown("### 상품 중분류별 매출")
        mid_df = filtered.groupby("medium_category").agg(
            순매출=("net_sales", "sum")
        ).sort_values("순매출", ascending=False).head(10).reset_index()

        fig_mid = go.Figure(go.Pie(
            values=mid_df["순매출"], labels=mid_df["medium_category"],
            marker=dict(colors=PIE_COLORS, line=dict(color="#FFFFFF", width=2)),
            textposition="inside", textinfo="percent+label",
            textfont=dict(size=11), hole=0.38
        ))
        fig_mid.update_layout(
            height=400,
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            font=dict(color="#3C4043"),
            margin=dict(t=20, b=20, l=10, r=10),
            showlegend=False
        )
        st.plotly_chart(fig_mid, use_container_width=True)

    # ── 채널별 월별 흐름 ──────────────────────────────────
    st.markdown("### 채널별 월별 매출 흐름")
    channel_monthly = filtered.groupby(["year_month", "client"])["net_sales"].sum().reset_index()
    top_clients = channel_df.sort_values("순매출", ascending=False).head(6)["client"].tolist()

    fig_multi = go.Figure()
    for i, client in enumerate(top_clients):
        df_c = channel_monthly[channel_monthly["client"] == client]
        color = CHANNEL_COLORS[i % len(CHANNEL_COLORS)]
        r, g, b = int(color[1:3],16), int(color[3:5],16), int(color[5:7],16)
        fig_multi.add_trace(go.Scatter(
            x=df_c["year_month"], y=df_c["net_sales"],
            name=client, mode="lines+markers",
            line=dict(color=color, width=2, shape="spline"),
            marker=dict(size=5, color=color),
            fill="tozeroy", fillcolor=f"rgba({r},{g},{b},0.05)",
            hovertemplate=f"{client}: ₩%{{y:,.0f}}<extra></extra>"
        ))
    fig_multi.update_layout(
        height=360,
        **glay(legend=dict(orientation="h", yanchor="bottom", y=1.02))
    )
    fig_multi.update_xaxes(type="category", **AXIS)
    fig_multi.update_yaxes(tickformat=",", **AXIS)
    st.plotly_chart(fig_multi, use_container_width=True)

    # ── 히트맵 ────────────────────────────────────────────
    st.markdown("### 채널 × 월별 매출 히트맵")
    pivot_df = filtered.groupby(["client","year_month"])["net_sales"].sum().reset_index()
    pivot_table = pivot_df.pivot(index="client", columns="year_month", values="net_sales").fillna(0)

    fig_heat = go.Figure(go.Heatmap(
        z=pivot_table.values/1e8,
        x=pivot_table.columns.tolist(), y=pivot_table.index.tolist(),
        colorscale=[[0,"#F8F9FA"],[0.5,"#8AB4F8"],[1,C_BLUE]],
        text=[[f"₩{v:.1f}억" for v in row] for row in pivot_table.values/1e8],
        texttemplate="%{text}", textfont=dict(color="#202124", size=11),
        colorbar=dict(title=dict(text="억원", font=dict(color="#5F6368")),
                      tickfont=dict(color="#5F6368"))
    ))
    fig_heat.update_layout(
        height=360, plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        font=dict(color="#3C4043"), margin=dict(t=20, b=20, l=10, r=10),
        xaxis=dict(tickfont=dict(color="#5F6368")),
        yaxis=dict(tickfont=dict(color="#5F6368"))
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("---")

    # ── 매출 vs 입고 (분리) ───────────────────────────────
    left_c, right_c = st.columns(2)

    with left_c:
        st.markdown("### 월별 매출")
        sales_view = st.radio("표시 기준", ["금액", "수량"], horizontal=True, key="sales_view2")
        sm = filtered.groupby("year_month").agg(
            순매출=("net_sales","sum"), 수량=("net_qty","sum")
        ).reset_index()
        sv_col = "순매출" if sales_view == "금액" else "수량"
        sv_fmt = [f"₩{v/1e8:.1f}억" for v in sm[sv_col]] if sales_view == "금액" else [f"{v:,}개" for v in sm[sv_col]]

        fig_sv = go.Figure(go.Scatter(
            x=sm["year_month"], y=sm[sv_col],
            mode="lines+markers",
            line=dict(color=C_BLUE, width=2.5, shape="spline"),
            marker=dict(size=7, color=C_BLUE),
            fill="tozeroy", fillcolor="rgba(26,115,232,0.08)",
            text=sv_fmt, hovertemplate="%{text}<extra></extra>"
        ))
        fig_sv.update_layout(height=300, **glay())
        fig_sv.update_xaxes(type="category", **AXIS)
        fig_sv.update_yaxes(tickformat=",", **AXIS)
        st.plotly_chart(fig_sv, use_container_width=True)

    with right_c:
        st.markdown("### 월별 입고 수량")
        stock_monthly = df_stock[
            (df_stock["year_month"] >= month_start) &
            (df_stock["year_month"] <= month_end)
        ].groupby("year_month")["quantity"].sum().reset_index()
        stock_monthly.columns = ["year_month", "입고수량"]

        fig_stock = go.Figure(go.Scatter(
            x=stock_monthly["year_month"], y=stock_monthly["입고수량"],
            mode="lines+markers",
            line=dict(color=C_GRAY, width=2.5, shape="spline"),
            marker=dict(size=7, color=C_GRAY),
            fill="tozeroy", fillcolor="rgba(95,99,104,0.08)",
            hovertemplate="입고: %{y:,}개<extra></extra>"
        ))
        fig_stock.update_layout(height=300, **glay())
        fig_stock.update_xaxes(type="category", **AXIS)
        fig_stock.update_yaxes(tickformat=",", **AXIS)
        st.plotly_chart(fig_stock, use_container_width=True)

# ════════════════════════════════════════════════════════
# TAB 2
# ════════════════════════════════════════════════════════
with tab2:
    st.markdown("### 🔍 상세 데이터 조회")
    st.markdown("<span style='color:#5F6368;font-size:0.85rem;'>날짜 범위와 상품명으로 원하는 데이터를 조회하세요.</span>", unsafe_allow_html=True)
    st.markdown("---")

    # 필터
    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        date_min = df_sales["date"].min().date()
        date_max = df_sales["date"].max().date()
        date_from = st.date_input("시작일", value=date_min, min_value=date_min, max_value=date_max)
    with f2:
        date_to = st.date_input("종료일", value=date_max, min_value=date_min, max_value=date_max)
    with f3:
        detail_clients = st.multiselect(
            "채널 선택",
            options=sorted(df_sales["client"].unique()),
            default=sorted(df_sales["client"].unique()),
            key="detail_clients"
        )
    keyword = st.text_input("상품명 검색", placeholder="예: 오딧  →  오딧 캐리어, 오딧백 등 모두 검색")

    st.markdown("")

    # 필터 적용
    detail_df = df_sales[
        (df_sales["date"].dt.date >= date_from) &
        (df_sales["date"].dt.date <= date_to) &
        (df_sales["client"].isin(detail_clients))
    ].copy()

    if keyword.strip():
        for kw in keyword.strip().split():
            detail_df = detail_df[detail_df["product_name"].str.contains(kw, case=False, na=False)]

    # KPI
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("조회 건수",    f"{len(detail_df):,}건")
    with c2: st.metric("순매출 합계",  f"₩{detail_df['net_sales'].sum()/1e8:.2f}억")
    with c3: st.metric("판매 수량",    f"{detail_df['net_qty'].sum():,}개")

    st.markdown("---")

    # 상품별 합산
    agg_df = detail_df.groupby("product_name").agg(
        순매출=("net_sales", "sum"),
        판매수량=("net_qty", "sum"),
        주문건수=("order_no_1", "nunique"),
        환불수량=("return_quantity", "sum"),
        환불금액=("return_sales_amount", "sum")
    ).reset_index().rename(columns={"product_name": "상품명"})

    sort_by = st.radio("정렬 기준", ["순매출", "판매수량"], horizontal=True, key="sort_by")
    agg_df = agg_df.sort_values(sort_by, ascending=False).reset_index(drop=True)
    agg_df.index += 1

    # ── 스파크라인 (상위 8개 상품 월별 흐름) ─────────────
    st.markdown("#### 상위 상품 월별 흐름")
    top8 = agg_df.head(8)["상품명"].tolist()
    spark_data = detail_df[detail_df["product_name"].isin(top8)]
    spark_monthly = spark_data.groupby(["product_name","year_month"])["net_sales"].sum().reset_index()

    if not spark_monthly.empty:
        fig_spark = make_subplots(
            rows=4, cols=2,
            subplot_titles=top8[:8],
            vertical_spacing=0.12,
            horizontal_spacing=0.08
        )
        for i, product in enumerate(top8):
            row, col = divmod(i, 2)
            df_p = spark_monthly[spark_monthly["product_name"] == product].sort_values("year_month")
            fig_spark.add_trace(go.Scatter(
                x=df_p["year_month"], y=df_p["net_sales"],
                mode="lines+markers",
                line=dict(color=C_BLUE, width=2, shape="spline"),
                marker=dict(size=5, color=C_BLUE),
                fill="tozeroy", fillcolor="rgba(26,115,232,0.08)",
                showlegend=False,
                hovertemplate="₩%{y:,}<extra></extra>"
            ), row=row+1, col=col+1)

        fig_spark.update_layout(
            height=520,
            **glay(margin=dict(t=50, b=20, l=10, r=10))
        )
        fig_spark.update_xaxes(type="category", tickfont=dict(color="#9AA0A6", size=9),
                               gridcolor="#F1F3F4", linecolor="#DADCE0")
        fig_spark.update_yaxes(tickformat=",", tickfont=dict(color="#9AA0A6", size=9),
                               gridcolor="#F1F3F4", linecolor="#DADCE0")
        for ann in fig_spark.layout.annotations:
            ann.font.size = 11
            ann.font.color = "#3C4043"
        st.plotly_chart(fig_spark, use_container_width=True)

    # ── 상품별 합산 테이블 ────────────────────────────────
    st.markdown("#### 상품별 기간 합산")
    disp = agg_df.copy()
    disp["순매출"]  = disp["순매출"].apply(lambda x: f"₩{x:,}")
    disp["환불금액"] = disp["환불금액"].apply(lambda x: f"₩{x:,}")
    st.dataframe(disp, use_container_width=True, hide_index=False, height=450)

st.markdown("---")
st.caption("Dotus Dashboard · 데이터 기준: 제공된 엑셀 파일")
