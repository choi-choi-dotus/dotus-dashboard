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

# ── 색상 팔레트 ──────────────────────────────────────────
BG       = "#0d1b2a"   # 메인 배경
CARD     = "#132337"   # 카드 배경
CARD2    = "#1a2f45"   # 보조 카드
BORDER   = "#1e3a5f"   # 테두리
CYAN     = "#00BCD4"   # 포인트 (시안)
CYAN2    = "#4DD0E1"   # 연한 시안
GOLD     = "#FFC107"   # 골드 (강조선)
TEXT     = "#E0F7FA"   # 밝은 텍스트
TEXT2    = "#90A4AE"   # 보조 텍스트

CHART_COLORS = [CYAN, CYAN2, "#0097A7", "#00838F", "#006064", "#80DEEA"]
PIE_COLORS   = [CYAN, CYAN2, "#0097A7", "#00838F", "#006064",
                "#80DEEA", "#B2EBF2", "#E0F7FA", "#4DD0E1", "#26C6DA"]

def nl(**kwargs):
    """Navy layout base"""
    base = dict(
        plot_bgcolor=CARD,
        paper_bgcolor=CARD,
        font=dict(color=TEXT2, size=12),
        margin=dict(t=40, b=20, l=10, r=10),
    )
    base.update(kwargs)
    return base

AXIS = dict(
    gridcolor=CARD2, linecolor=BORDER,
    tickcolor=BORDER, tickfont=dict(color=TEXT2)
)

# ── CSS ──────────────────────────────────────────────────
st.markdown(f"""
<style>
.stApp {{
    background-color: {BG};
}}
[data-testid="stAppViewContainer"] > .main {{
    background-color: {BG};
}}
[data-testid="stSidebar"] {{
    background-color: #0a1628;
    border-right: 1px solid {BORDER};
}}
[data-testid="stSidebar"] * {{ color: {TEXT2} !important; }}
[data-testid="stSidebar"] h2 {{
    color: {TEXT} !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
}}
[data-testid="stSidebar"] h3 {{
    color: {TEXT2} !important;
    font-size: 0.72rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {{
    color: {TEXT2} !important;
    font-size: 0.75rem !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}}
[data-testid="stSidebar"] span[data-baseweb="tag"] {{
    background-color: #0e3a50 !important;
    color: {CYAN} !important;
}}
[data-testid="stSidebar"] span[data-baseweb="tag"] span {{
    color: {CYAN} !important;
}}
p, span, label, div {{ color: {TEXT2}; }}
h1, h2, h3, h4 {{ color: {TEXT} !important; }}
hr {{ border-color: {BORDER} !important; }}

[data-testid="stMetric"] {{
    background: {CARD};
    border: 1px solid {BORDER};
    border-top: 3px solid {CYAN};
    border-radius: 8px;
    padding: 20px 24px;
}}
[data-testid="stMetricValue"] {{
    font-size: 1.8rem !important;
    font-weight: 700 !important;
    color: {CYAN} !important;
}}
[data-testid="stMetricLabel"] {{
    font-size: 0.78rem !important;
    color: {TEXT2} !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}}
[data-testid="stMetricDelta"] svg {{ display: none; }}

[data-testid="stTabs"] button {{
    color: {TEXT2} !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    background: transparent !important;
}}
[data-testid="stTabs"] button[aria-selected="true"] {{
    color: {CYAN} !important;
    border-bottom-color: {CYAN} !important;
}}
[data-testid="stTabs"] {{
    border-bottom: 1px solid {BORDER};
}}

.stButton button {{
    background-color: {CARD} !important;
    color: {CYAN} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 4px !important;
}}
.stButton button:hover {{
    border-color: {CYAN} !important;
    background-color: #0e3a50 !important;
}}
.stRadio label {{ color: {TEXT2} !important; font-size: 0.88rem !important; }}
.stRadio [data-baseweb="radio"] div[data-checked="true"] {{
    background-color: {CYAN} !important;
    border-color: {CYAN} !important;
}}
[data-testid="stDataFrame"] {{
    border: 1px solid {BORDER};
    border-radius: 8px;
}}
</style>
""", unsafe_allow_html=True)

# ── 비밀번호 ──────────────────────────────────────────────
PASSWORD = "dotus2026"

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True

    st.markdown(f"""
    <div style='display:flex;justify-content:center;margin-top:80px;'>
        <div style='text-align:center;'>
            <h1 style='font-size:2.2rem;font-weight:700;color:{CYAN};margin-bottom:4px;'>📊 Dotus</h1>
            <p style='color:{TEXT2};margin-bottom:32px;font-size:0.95rem;'>매출 대시보드</p>
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
    df_sales["net_qty"]   = df_sales["quantity"] - df_sales["return_quantity"]

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
    month_end   = st.selectbox("종료 월", months, index=len(months)-1)

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
    prev_f = df_sales[
        (df_sales["year_month"] >= prev_start) &
        (df_sales["year_month"] <= prev_end) &
        (df_sales["client"].isin(selected_clients)) &
        (df_sales["large_category"].isin(selected_large))
    ]
else:
    prev_f = pd.DataFrame()

# KPI
total_sales  = filtered["net_sales"].sum()
total_qty    = filtered["net_qty"].sum()
total_orders = filtered["order_no_1"].nunique()
avg_order    = total_sales / total_orders if total_orders > 0 else 0

if not prev_f.empty:
    ps = prev_f["net_sales"].sum()
    pq = prev_f["net_qty"].sum()
    po = prev_f["order_no_1"].nunique()
    sales_delta  = f"{((total_sales - ps)/ps*100):+.1f}%" if ps else "N/A"
    qty_delta    = f"{((total_qty   - pq)/pq*100):+.1f}%" if pq else "N/A"
    orders_delta = f"{((total_orders- po)/po*100):+.1f}%" if po else "N/A"
else:
    sales_delta = qty_delta = orders_delta = None

# ── 헤더 ──────────────────────────────────────────────────
st.markdown("# 📊 Dotus 매출 대시보드")
cat_info = f"대분류 {len(selected_large)}개"
if selected_medium: cat_info += f" · 중분류 {len(selected_medium)}개"
if HAS_SMALL and selected_small: cat_info += f" · 소분류 {len(selected_small)}개"
st.markdown(
    f"<span style='color:{TEXT2};font-size:0.88rem;'>"
    f"조회 기간: <b style='color:{TEXT};'>{month_start} ~ {month_end}</b>"
    f" &nbsp;|&nbsp; 채널: <b style='color:{TEXT};'>{len(selected_clients)}개</b>"
    f" &nbsp;|&nbsp; {cat_info}</span>",
    unsafe_allow_html=True
)
st.markdown("---")

# ════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📈 매출 대시보드", "🔍 상세 데이터 조회"])

# ════════════════════════════════════════════════════════
# TAB 1
# ════════════════════════════════════════════════════════
with tab1:

    # ── KPI ──────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("순매출",       f"₩{total_sales/1e8:.1f}억",  delta=sales_delta)
    with k2: st.metric("판매 수량",    f"{total_qty:,}개",            delta=qty_delta)
    with k3: st.metric("주문 건수",    f"{total_orders:,}건",         delta=orders_delta)
    with k4: st.metric("평균 주문금액", f"₩{avg_order/1e4:.0f}만")
    st.markdown("---")

    # ── 월별 매출 추이 ────────────────────────────────────
    st.markdown("### 월별 매출 추이")
    sales_metric = st.radio("기준", ["금액 (억원)", "수량 (개)"], horizontal=True, key="sales_metric")

    monthly = filtered.groupby("year_month").agg(
        순매출=("net_sales", "sum"),
        수량=("net_qty", "sum"),
    ).reset_index()

    use_money = (sales_metric == "금액 (억원)")
    y_vals  = monthly["순매출"] / 1e8 if use_money else monthly["수량"]
    y_title = "억원" if use_money else "수량 (개)"
    y_hover = [f"₩{v:.1f}억" for v in y_vals] if use_money else [f"{v:,}개" for v in y_vals]
    monthly["MoM"] = y_vals.pct_change() * 100

    fig_sales = make_subplots(specs=[[{"secondary_y": True}]])
    fig_sales.add_trace(go.Scatter(
        x=monthly["year_month"], y=y_vals,
        name="매출", mode="lines+markers",
        line=dict(color=CYAN, width=3, shape="spline"),
        marker=dict(size=7, color=CYAN),
        fill="tozeroy", fillcolor="rgba(0,188,212,0.12)",
        text=y_hover, hovertemplate="%{text}<extra></extra>"
    ), secondary_y=False)
    fig_sales.add_trace(go.Scatter(
        x=monthly["year_month"], y=monthly["MoM"],
        name="MoM(%)", mode="lines+markers",
        line=dict(color=GOLD, width=2, dash="dot", shape="spline"),
        marker=dict(size=5, color=GOLD),
        hovertemplate="%{y:.1f}%<extra></extra>"
    ), secondary_y=True)
    fig_sales.update_layout(
        height=340,
        **nl(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color=TEXT2)))
    )
    fig_sales.update_xaxes(type="category", **AXIS)
    fig_sales.update_yaxes(title_text=y_title, secondary_y=False, **AXIS)
    fig_sales.update_yaxes(title_text="MoM (%)", secondary_y=True, **AXIS)
    st.plotly_chart(fig_sales, use_container_width=True)

    # ── 채널별 + 중분류 ───────────────────────────────────
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("### 채널별 매출")
        channel_df = filtered.groupby("client").agg(
            순매출=("net_sales", "sum")
        ).sort_values("순매출", ascending=True).reset_index()

        fig_ch = go.Figure(go.Bar(
            x=channel_df["순매출"] / 1e8,
            y=channel_df["client"],
            orientation="h",
            marker=dict(
                color=channel_df["순매출"],
                colorscale=[[0, CARD2], [1, CYAN]],
                showscale=False
            ),
            text=[f"₩{v/1e8:.1f}억" for v in channel_df["순매출"]],
            textposition="outside", textfont=dict(color=TEXT2)
        ))
        fig_ch.update_layout(height=380, **nl(margin=dict(t=20, b=20, l=10, r=80)))
        fig_ch.update_xaxes(title_text="억원", **AXIS)
        fig_ch.update_yaxes(**AXIS)
        st.plotly_chart(fig_ch, use_container_width=True)

    with col_right:
        st.markdown("### 상품 중분류별 매출")
        mid_df = filtered.groupby("medium_category").agg(
            순매출=("net_sales", "sum")
        ).sort_values("순매출", ascending=False).head(10).reset_index()

        fig_mid = go.Figure(go.Pie(
            values=mid_df["순매출"], labels=mid_df["medium_category"],
            marker=dict(colors=PIE_COLORS, line=dict(color=BG, width=2)),
            textposition="inside", textinfo="percent+label",
            textfont=dict(size=11, color=BG), hole=0.4
        ))
        fig_mid.update_layout(
            height=380,
            plot_bgcolor=CARD, paper_bgcolor=CARD,
            font=dict(color=TEXT2),
            margin=dict(t=20, b=20, l=10, r=10),
            showlegend=False
        )
        st.plotly_chart(fig_mid, use_container_width=True)

    # ── 채널별 월별 흐름 ──────────────────────────────────
    st.markdown("### 채널별 월별 매출 흐름")
    ch_monthly = filtered.groupby(["year_month","client"])["net_sales"].sum().reset_index()
    top6 = channel_df.sort_values("순매출", ascending=False).head(6)["client"].tolist()

    fig_multi = go.Figure()
    for i, client in enumerate(top6):
        df_c = ch_monthly[ch_monthly["client"] == client]
        color = CHART_COLORS[i % len(CHART_COLORS)]
        r,g,b = int(color[1:3],16), int(color[3:5],16), int(color[5:7],16)
        fig_multi.add_trace(go.Scatter(
            x=df_c["year_month"], y=df_c["net_sales"]/1e8,
            name=client, mode="lines+markers",
            line=dict(color=color, width=2, shape="spline"),
            marker=dict(size=5, color=color),
            fill="tozeroy", fillcolor=f"rgba({r},{g},{b},0.06)",
            hovertemplate=f"{client}: ₩%{{y:.1f}}억<extra></extra>"
        ))
    fig_multi.update_layout(
        height=340,
        **nl(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color=TEXT2)))
    )
    fig_multi.update_xaxes(type="category", **AXIS)
    fig_multi.update_yaxes(title_text="억원", **AXIS)
    st.plotly_chart(fig_multi, use_container_width=True)

    # ── 히트맵 ────────────────────────────────────────────
    st.markdown("### 채널 × 월별 매출 히트맵")
    pivot_df    = filtered.groupby(["client","year_month"])["net_sales"].sum().reset_index()
    pivot_table = pivot_df.pivot(index="client", columns="year_month", values="net_sales").fillna(0)

    fig_heat = go.Figure(go.Heatmap(
        z=pivot_table.values/1e8,
        x=pivot_table.columns.tolist(), y=pivot_table.index.tolist(),
        colorscale=[[0, CARD2],[0.5, "#00838F"],[1, CYAN]],
        text=[[f"₩{v:.1f}억" for v in row] for row in pivot_table.values/1e8],
        texttemplate="%{text}", textfont=dict(color=TEXT, size=11),
        colorbar=dict(
            title=dict(text="억원", font=dict(color=TEXT2)),
            tickfont=dict(color=TEXT2)
        )
    ))
    fig_heat.update_layout(
        height=340, plot_bgcolor=CARD, paper_bgcolor=CARD,
        font=dict(color=TEXT2), margin=dict(t=20, b=20, l=10, r=10),
        xaxis=dict(tickfont=dict(color=TEXT2)),
        yaxis=dict(tickfont=dict(color=TEXT2))
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("---")

    # ── 월별 매출 / 입고 수량 (나란히, 같은 높이) ─────────
    left_c, right_c = st.columns(2)
    CHART_H = 300

    with left_c:
        st.markdown("### 월별 매출")
        sv_opt = st.radio("표시 기준", ["금액 (억원)", "수량 (개)"], horizontal=True, key="sv2")
        sm = filtered.groupby("year_month").agg(
            순매출=("net_sales","sum"), 수량=("net_qty","sum")
        ).reset_index()
        use_m2 = (sv_opt == "금액 (억원)")
        sv_y   = sm["순매출"]/1e8 if use_m2 else sm["수량"]
        sv_fmt = [f"₩{v:.1f}억" for v in sv_y] if use_m2 else [f"{v:,}개" for v in sv_y]

        fig_sv = go.Figure(go.Scatter(
            x=sm["year_month"], y=sv_y,
            mode="lines+markers",
            line=dict(color=CYAN, width=2.5, shape="spline"),
            marker=dict(size=7, color=CYAN),
            fill="tozeroy", fillcolor="rgba(0,188,212,0.10)",
            text=sv_fmt, hovertemplate="%{text}<extra></extra>"
        ))
        fig_sv.update_layout(height=CHART_H, **nl())
        fig_sv.update_xaxes(type="category", **AXIS)
        fig_sv.update_yaxes(title_text="억원" if use_m2 else "수량", **AXIS)
        st.plotly_chart(fig_sv, use_container_width=True)

    with right_c:
        st.markdown("### 월별 입고 수량")
        st.markdown("&nbsp;", unsafe_allow_html=True)  # 높이 맞춤용 spacer
        stk = df_stock[
            (df_stock["year_month"] >= month_start) &
            (df_stock["year_month"] <= month_end)
        ].groupby("year_month")["quantity"].sum().reset_index()
        stk.columns = ["year_month","입고수량"]

        fig_stk = go.Figure(go.Scatter(
            x=stk["year_month"], y=stk["입고수량"],
            mode="lines+markers",
            line=dict(color=CYAN2, width=2.5, shape="spline"),
            marker=dict(size=7, color=CYAN2),
            fill="tozeroy", fillcolor="rgba(77,208,225,0.10)",
            hovertemplate="입고: %{y:,}개<extra></extra>"
        ))
        fig_stk.update_layout(height=CHART_H, **nl())
        fig_stk.update_xaxes(type="category", **AXIS)
        fig_stk.update_yaxes(title_text="수량 (개)", tickformat=",", **AXIS)
        st.plotly_chart(fig_stk, use_container_width=True)

# ════════════════════════════════════════════════════════
# TAB 2
# ════════════════════════════════════════════════════════
with tab2:
    st.markdown("### 🔍 상세 데이터 조회")
    st.markdown(f"<span style='color:{TEXT2};font-size:0.85rem;'>날짜 범위와 상품명으로 원하는 데이터를 조회하세요.</span>", unsafe_allow_html=True)
    st.markdown("---")

    # 필터
    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        date_min  = df_sales["date"].min().date()
        date_max  = df_sales["date"].max().date()
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
    with c1: st.metric("조회 건수",   f"{len(detail_df):,}건")
    with c2: st.metric("순매출 합계", f"₩{detail_df['net_sales'].sum()/1e8:.2f}억")
    with c3: st.metric("판매 수량",   f"{detail_df['net_qty'].sum():,}개")

    st.markdown("---")

    # ── 스파크라인 (조회 조건 전체 월별 흐름 1개) ─────────
    st.markdown("### 조회 기간 매출 흐름")
    spark_monthly = detail_df.groupby("year_month")["net_sales"].sum().reset_index()
    spark_monthly = spark_monthly.sort_values("year_month")

    if not spark_monthly.empty:
        fig_spark = go.Figure()
        fig_spark.add_trace(go.Scatter(
            x=spark_monthly["year_month"],
            y=spark_monthly["net_sales"] / 1e8,
            mode="lines+markers",
            line=dict(color=CYAN, width=3, shape="spline"),
            marker=dict(size=7, color=CYAN,
                        line=dict(color=BG, width=2)),
            fill="tozeroy",
            fillcolor="rgba(0,188,212,0.12)",
            text=[f"₩{v:.2f}억" for v in spark_monthly["net_sales"]/1e8],
            hovertemplate="%{x}: %{text}<extra></extra>"
        ))
        fig_spark.update_layout(height=220, **nl(margin=dict(t=20, b=20, l=10, r=10)))
        fig_spark.update_xaxes(type="category", **AXIS)
        fig_spark.update_yaxes(title_text="억원", **AXIS)
        st.plotly_chart(fig_spark, use_container_width=True)

    st.markdown("---")

    # ── 상품별 합산 테이블 ────────────────────────────────
    st.markdown("### 상품별 기간 합산")
    sort_by = st.radio("정렬 기준", ["순매출", "판매수량"], horizontal=True, key="sort_by")

    agg_df = detail_df.groupby("product_name").agg(
        순매출=("net_sales", "sum"),
        판매수량=("net_qty", "sum"),
        주문건수=("order_no_1", "nunique"),
        환불수량=("return_quantity", "sum"),
        환불금액=("return_sales_amount", "sum")
    ).reset_index().rename(columns={"product_name": "상품명"})

    agg_df = agg_df.sort_values(sort_by, ascending=False).reset_index(drop=True)
    agg_df.index += 1

    disp = agg_df.copy()
    disp["순매출"]  = disp["순매출"].apply(lambda x: f"₩{x:,}")
    disp["환불금액"] = disp["환불금액"].apply(lambda x: f"₩{x:,}")
    st.dataframe(disp, use_container_width=True, hide_index=False, height=480)

st.markdown("---")
st.caption("Dotus Dashboard · 데이터 기준: 제공된 엑셀 파일")
