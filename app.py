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
.stApp { background-color: #0d1117; }
[data-testid="stAppViewContainer"] > .main { background-color: #0d1117; }
[data-testid="stSidebar"] {
    background-color: #161b22;
    border-right: 1px solid #30363d;
}
[data-testid="stSidebar"] * { color: #c9d1d9 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #8b949e !important;
    font-size: 0.75rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
[data-testid="stSidebar"] h2 { color: #f0f6fc !important; font-size: 1rem !important; }
[data-testid="stSidebar"] h3 {
    color: #8b949e !important;
    font-size: 0.72rem !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
[data-testid="stSidebar"] span[data-baseweb="tag"] { background-color: #1f6feb !important; }
p, span, label, div { color: #c9d1d9; }
h1, h2, h3 { color: #f0f6fc !important; }
[data-testid="stMetric"] {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 20px 24px;
}
[data-testid="stMetricValue"] { font-size: 1.8rem !important; font-weight: 700 !important; color: #f0f6fc !important; }
[data-testid="stMetricLabel"] { font-size: 0.82rem !important; color: #8b949e !important; }
hr { border-color: #30363d !important; }
.stButton button {
    background-color: #21262d !important;
    color: #c9d1d9 !important;
    border: 1px solid #30363d !important;
}
.stButton button:hover { border-color: #58a6ff !important; color: #58a6ff !important; }
/* 탭 스타일 */
[data-testid="stTabs"] button {
    color: #8b949e !important;
    font-size: 0.95rem !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #f0f6fc !important;
    border-bottom-color: #58a6ff !important;
}
</style>
""", unsafe_allow_html=True)

# ── 색상 팔레트 ──────────────────────────────────────────
C_BLUE   = "#58a6ff"
C_CYAN   = "#79c0ff"
C_GREEN  = "#3fb950"
C_ORANGE = "#f79520"
C_RED    = "#ff7b72"
C_PURPLE = "#d2a8ff"
PIE_COLORS = [C_BLUE, C_CYAN, C_GREEN, C_ORANGE, C_RED, C_PURPLE,
              "#ffa657", "#a5d6ff", "#7ee787", "#ff9492"]

def dark_layout(**kwargs):
    base = dict(
        plot_bgcolor="#0d1117",
        paper_bgcolor="#161b22",
        font=dict(color="#c9d1d9", size=12),
        margin=dict(t=40, b=20, l=10, r=10),
    )
    base.update(kwargs)
    return base

AXIS = dict(gridcolor="#21262d", linecolor="#30363d", tickcolor="#8b949e", tickfont=dict(color="#8b949e"))

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

# ── 공통 필터링 ───────────────────────────────────────────
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

# ════════════════════════════════════════════════════════
# 탭 구성
# ════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📈 매출 대시보드", "🔍 상세 데이터 조회"])

# ════════════════════════════════════════════════════════
# TAB 1 : 매출 대시보드
# ════════════════════════════════════════════════════════
with tab1:

    # KPI 카드
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("💰 순매출",       f"₩{total_sales/1e8:.1f}억",  delta=sales_delta)
    with k2: st.metric("📦 판매 수량",    f"{total_qty:,}개",            delta=qty_delta)
    with k3: st.metric("🛒 주문 건수",    f"{total_orders:,}건",         delta=orders_delta)
    with k4: st.metric("💳 평균 주문금액", f"₩{avg_order_val:,.0f}")

    st.markdown("---")

    # 월별 매출 추이 (에어리어)
    st.markdown("### 📈 월별 매출 추이")
    monthly = filtered.groupby("year_month").agg(
        순매출=("net_sales", "sum"),
        수량=("net_qty", "sum"),
        주문수=("order_no_1", "nunique")
    ).reset_index()
    monthly["MoM증가율"] = monthly["순매출"].pct_change() * 100

    fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])
    fig_monthly.add_trace(go.Scatter(
        x=monthly["year_month"], y=monthly["순매출"],
        name="순매출", mode="lines+markers",
        fill="tozeroy", fillcolor="rgba(88,166,255,0.15)",
        line=dict(color=C_BLUE, width=3), marker=dict(size=7, color=C_BLUE),
        text=[f"₩{v/1e8:.1f}억" for v in monthly["순매출"]],
        hovertemplate="%{text}<extra></extra>"
    ), secondary_y=False)
    fig_monthly.add_trace(go.Scatter(
        x=monthly["year_month"], y=monthly["MoM증가율"],
        name="MoM 증가율(%)", mode="lines+markers",
        line=dict(color=C_ORANGE, width=2, dash="dot"),
        marker=dict(size=6, color=C_ORANGE),
        hovertemplate="%{y:.1f}%<extra></extra>"
    ), secondary_y=True)
    fig_monthly.update_layout(
        height=360,
        **dark_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#c9d1d9")))
    )
    fig_monthly.update_xaxes(type="category", **AXIS)
    fig_monthly.update_yaxes(title_text="순매출 (원)", secondary_y=False, tickformat=",", **AXIS)
    fig_monthly.update_yaxes(title_text="MoM 증가율 (%)", secondary_y=True, **AXIS)
    st.plotly_chart(fig_monthly, use_container_width=True)

    # 채널별 + 중분류 파이
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("### 🏪 채널별 매출")
        channel_df = filtered.groupby("client").agg(
            순매출=("net_sales", "sum")
        ).sort_values("순매출", ascending=True).reset_index()

        fig_ch = go.Figure(go.Bar(
            x=channel_df["순매출"], y=channel_df["client"],
            orientation="h",
            marker=dict(color=channel_df["순매출"],
                        colorscale=[[0, "#1f3a5f"], [1, C_BLUE]], showscale=False),
            text=[f"₩{v/1e8:.1f}억" for v in channel_df["순매출"]],
            textposition="outside", textfont=dict(color="#c9d1d9")
        ))
        fig_ch.update_layout(
            height=400,
            **dark_layout(margin=dict(t=20, b=20, l=10, r=80))
        )
        fig_ch.update_xaxes(tickformat=",", **AXIS)
        fig_ch.update_yaxes(**AXIS)
        st.plotly_chart(fig_ch, use_container_width=True)

    with col_right:
        st.markdown("### 📦 상품 중분류별 매출")
        mid_df = filtered.groupby("medium_category").agg(
            순매출=("net_sales", "sum")
        ).sort_values("순매출", ascending=False).head(10).reset_index()

        fig_mid = go.Figure(go.Pie(
            values=mid_df["순매출"], labels=mid_df["medium_category"],
            marker=dict(colors=PIE_COLORS, line=dict(color="#0d1117", width=2)),
            textposition="inside", textinfo="percent+label",
            textfont=dict(color="#0d1117", size=11), hole=0.35
        ))
        fig_mid.update_layout(
            height=400,
            plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
            font=dict(color="#c9d1d9"),
            margin=dict(t=20, b=20, l=10, r=10), showlegend=False
        )
        st.plotly_chart(fig_mid, use_container_width=True)

    # 채널별 월별 흐름
    st.markdown("### 🌊 채널별 월별 매출 흐름")
    channel_monthly = filtered.groupby(["year_month", "client"])["net_sales"].sum().reset_index()
    top_clients = channel_df.sort_values("순매출", ascending=False).head(6)["client"].tolist()

    fig_multi = go.Figure()
    line_colors = [C_BLUE, C_CYAN, C_GREEN, C_ORANGE, C_RED, C_PURPLE]
    for i, client in enumerate(top_clients):
        df_c = channel_monthly[channel_monthly["client"] == client]
        color = line_colors[i % len(line_colors)]
        r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
        fig_multi.add_trace(go.Scatter(
            x=df_c["year_month"], y=df_c["net_sales"],
            name=client, mode="lines+markers",
            fill="tozeroy", fillcolor=f"rgba({r},{g},{b},0.07)",
            line=dict(color=color, width=2.5), marker=dict(size=6, color=color),
            hovertemplate=f"{client}: ₩%{{y:,.0f}}<extra></extra>"
        ))
    fig_multi.update_layout(
        height=380,
        **dark_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#c9d1d9")))
    )
    fig_multi.update_xaxes(type="category", **AXIS)
    fig_multi.update_yaxes(tickformat=",", **AXIS)
    st.plotly_chart(fig_multi, use_container_width=True)

    # 히트맵
    st.markdown("### 🗺 채널 × 월별 매출 히트맵")
    pivot_df = filtered.groupby(["client", "year_month"])["net_sales"].sum().reset_index()
    pivot_table = pivot_df.pivot(index="client", columns="year_month", values="net_sales").fillna(0)

    fig_heat = go.Figure(go.Heatmap(
        z=pivot_table.values / 1e8,
        x=pivot_table.columns.tolist(), y=pivot_table.index.tolist(),
        colorscale=[[0, "#0d1117"], [0.5, "#1f6feb"], [1, "#58a6ff"]],
        text=[[f"₩{v:.1f}억" for v in row] for row in pivot_table.values / 1e8],
        texttemplate="%{text}", textfont=dict(color="#f0f6fc", size=11),
        colorbar=dict(title=dict(text="억원", font=dict(color="#c9d1d9")), tickfont=dict(color="#c9d1d9"))
    ))
    fig_heat.update_layout(
        height=380,
        plot_bgcolor="#0d1117", paper_bgcolor="#161b22",
        font=dict(color="#c9d1d9"), margin=dict(t=20, b=20, l=10, r=10),
        xaxis=dict(tickfont=dict(color="#8b949e")),
        yaxis=dict(tickfont=dict(color="#8b949e"))
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # 입고 vs 매출
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
        x=compare["year_month"], y=compare["매출금액"], name="매출",
        mode="lines+markers", fill="tozeroy", fillcolor="rgba(88,166,255,0.12)",
        line=dict(color=C_BLUE, width=3), marker=dict(size=7),
        hovertemplate="매출: ₩%{y:,.0f}<extra></extra>"
    ))
    fig_compare.add_trace(go.Scatter(
        x=compare["year_month"], y=compare["입고금액"], name="입고",
        mode="lines+markers", fill="tozeroy", fillcolor="rgba(247,149,32,0.10)",
        line=dict(color=C_ORANGE, width=3), marker=dict(size=7),
        hovertemplate="입고: ₩%{y:,.0f}<extra></extra>"
    ))
    fig_compare.update_layout(
        height=340,
        **dark_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#c9d1d9")))
    )
    fig_compare.update_xaxes(**AXIS)
    fig_compare.update_yaxes(tickformat=",", **AXIS)
    st.plotly_chart(fig_compare, use_container_width=True)

# ════════════════════════════════════════════════════════
# TAB 2 : 상세 데이터 조회
# ════════════════════════════════════════════════════════
with tab2:
    st.markdown("### 🔍 상세 데이터 조회")
    st.markdown("<span style='color:#8b949e;font-size:0.85rem;'>일별 날짜 범위와 상품명으로 원하는 데이터를 조회하세요.</span>", unsafe_allow_html=True)
    st.markdown("---")

    # 필터 행 1: 날짜 + 채널
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

    # 필터 행 2: 상품명 검색
    keyword = st.text_input("상품명 검색", placeholder="예: 오딧  →  오딧 캐리어, 오딧백 등 모두 검색")

    st.markdown("")

    # 조회 필터 적용
    detail_df = df_sales[
        (df_sales["date"].dt.date >= date_from) &
        (df_sales["date"].dt.date <= date_to) &
        (df_sales["client"].isin(detail_clients))
    ].copy()

    if keyword.strip():
        keywords = keyword.strip().split()
        mask = pd.Series([True] * len(detail_df), index=detail_df.index)
        for kw in keywords:
            mask = mask & detail_df["product_name"].str.contains(kw, case=False, na=False)
        detail_df = detail_df[mask]

    # 결과 요약 KPI
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        st.metric("조회 건수", f"{len(detail_df):,}건")
    with col_s2:
        st.metric("순매출 합계", f"₩{detail_df['net_sales'].sum()/1e8:.2f}억")
    with col_s3:
        st.metric("판매 수량", f"{detail_df['net_qty'].sum():,}개")

    st.markdown("")

    # 원본 데이터 테이블
    st.markdown("#### 📋 원본 데이터")
    display_cols = ["date", "product_name", "client",
                    "large_category", "medium_category",
                    "net_sales", "net_qty", "return_quantity", "return_sales_amount"]
    if HAS_SMALL:
        display_cols.insert(5, "small_category")

    available_cols = [c for c in display_cols if c in detail_df.columns]
    show_df = detail_df[available_cols].copy()
    show_df = show_df.rename(columns={
        "date": "날짜", "product_name": "상품명", "client": "채널",
        "large_category": "대분류", "medium_category": "중분류", "small_category": "소분류",
        "net_sales": "순매출", "net_qty": "순수량",
        "return_quantity": "반품수량", "return_sales_amount": "반품금액"
    })
    show_df["날짜"] = show_df["날짜"].dt.strftime("%Y-%m-%d")

    st.dataframe(
        show_df.sort_values("날짜", ascending=False),
        use_container_width=True, hide_index=True, height=400
    )
    st.caption(f"총 {len(show_df):,}행 표시 중")

    st.markdown("---")

    # 상품별 기간 합산
    st.markdown("#### 📊 상품별 기간 합산")
    sort_col_label = st.radio("정렬 기준", ["순매출", "판매수량"], horizontal=True)
    sort_col = "순매출" if sort_col_label == "순매출" else "판매수량"

    agg_df = detail_df.groupby("product_name").agg(
        순매출=("net_sales", "sum"),
        판매수량=("net_qty", "sum"),
        주문건수=("order_no_1", "nunique"),
        반품수량=("return_quantity", "sum"),
        반품금액=("return_sales_amount", "sum")
    ).reset_index()
    agg_df = agg_df.rename(columns={"product_name": "상품명"})
    agg_df = agg_df.sort_values(sort_col, ascending=False).reset_index(drop=True)
    agg_df.index += 1  # 순위 1부터

    agg_display = agg_df.copy()
    agg_display["순매출"] = agg_display["순매출"].apply(lambda x: f"₩{x:,}")
    agg_display["반품금액"] = agg_display["반품금액"].apply(lambda x: f"₩{x:,}")

    st.dataframe(agg_display, use_container_width=True, height=400)

st.markdown("---")
st.caption("Dotus Dashboard · 데이터 기준: 제공된 엑셀 파일")
