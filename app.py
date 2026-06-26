import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from itertools import combinations
from datetime import date, timedelta

# ── 페이지 설정 ──────────────────────────────────────────
st.set_page_config(
    page_title="Dotus 매출 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 색상 ─────────────────────────────────────────────────
BG      = "#0d1b2a"
CARD    = "#132337"
CARD2   = "#1a2f45"
BORDER  = "#1e3a5f"
CYAN    = "#00BCD4"
CYAN2   = "#4DD0E1"
GOLD    = "#FFC107"
TEXT    = "#E0F7FA"
TEXT2   = "#90A4AE"

CHART_COLORS = [CYAN, "#FFC107", "#FF6B6B", "#AB47BC", "#66BB6A", "#FF9800"]
PIE_COLORS   = [CYAN, CYAN2, "#0097A7", "#00838F", "#006064",
                "#80DEEA", "#B2EBF2", "#E0F7FA", "#4DD0E1", "#26C6DA"]

def nl(**kwargs):
    base = dict(plot_bgcolor=CARD, paper_bgcolor=CARD,
                font=dict(color=TEXT2, size=12), margin=dict(t=40,b=20,l=10,r=10))
    base.update(kwargs)
    return base

AXIS = dict(gridcolor=CARD2, linecolor=BORDER, tickcolor=BORDER, tickfont=dict(color=TEXT2))

# ── CSS ──────────────────────────────────────────────────
st.markdown(f"""
<style>
.stApp {{ background-color: {BG}; }}
[data-testid="stAppViewContainer"] > .main {{ background-color: {BG}; }}
[data-testid="stSidebar"] {{
    background-color: #0a1628;
    border-right: 1px solid {BORDER};
}}
[data-testid="stSidebar"] * {{ color: {TEXT2} !important; }}
[data-testid="stSidebar"] h2 {{ color: {TEXT} !important; font-size:1rem !important; font-weight:600 !important; }}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {{
    padding: 10px 14px !important;
    border-radius: 8px !important;
    font-size: 0.92rem !important;
    font-weight: 500 !important;
    cursor: pointer !important;
    display: flex !important;
    align-items: center !important;
    color: {TEXT2} !important;
    margin-bottom: 4px !important;
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
[data-testid="stMetricValue"] {{ font-size:1.8rem !important; font-weight:700 !important; color:{CYAN} !important; }}
[data-testid="stMetricLabel"] {{ font-size:0.78rem !important; color:{TEXT2} !important; text-transform:uppercase; letter-spacing:0.05em; }}
span[data-baseweb="tag"] {{ background-color: #0e3a50 !important; color:{CYAN} !important; }}
span[data-baseweb="tag"] span {{ color:{CYAN} !important; }}
.stRadio label {{ color:{TEXT2} !important; font-size:0.88rem !important; }}
.stButton button {{
    background-color: {CARD} !important; color:{CYAN} !important;
    border: 1px solid {BORDER} !important; border-radius:4px !important;
}}
.stButton button:hover {{ border-color:{CYAN} !important; background-color:#0e3a50 !important; }}
[data-testid="stDataFrame"] {{ border:1px solid {BORDER}; border-radius:8px; }}
.stTextInput input {{
    background-color: {CARD} !important;
    border: 1px solid {BORDER} !important;
    color: {TEXT} !important;
    border-radius: 6px !important;
}}
.stSelectbox > div > div {{
    background-color: {CARD} !important;
    border: 1px solid {BORDER} !important;
    color: {TEXT} !important;
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
            <p style='color:{TEXT2};margin-bottom:32px;'>매출 대시보드</p>
        </div>
    </div>""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1,1.2,1])
    with c2:
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
    df = pd.DataFrame(data1, columns=headers1)
    df["date"]       = pd.to_datetime(df["date"])
    df["year_month"] = df["date"].dt.to_period("M").astype(str)
    df["net_sales"]  = df["sales_amount"] - df["return_sales_amount"]
    df["net_qty"]    = df["quantity"] - df["return_quantity"]

    ws2 = wb["시트2"]
    rows2 = list(ws2.iter_rows(values_only=True))
    headers2 = rows2[0][:13]
    data2 = [r[:13] for r in rows2[1:] if r[0] is not None]
    df_stock = pd.DataFrame(data2, columns=headers2)
    df_stock["date"]       = pd.to_datetime(df_stock["date"])
    df_stock["year_month"] = df_stock["date"].dt.to_period("M").astype(str)

    ws3 = wb["품목마스터"]
    rows3 = list(ws3.iter_rows(values_only=True))
    headers3 = rows3[0][:8]
    data3 = [r[:8] for r in rows3[1:] if r[0] is not None]
    df_master = pd.DataFrame(data3, columns=headers3)

    return df, df_stock, df_master

df_sales, df_stock, df_master = load_data()
HAS_SMALL = "small_category" in df_sales.columns

# ── 전체 선택 멀티셀렉트 헬퍼 ────────────────────────────
def multiselect_with_all(label, options, key, **kwargs):
    ALL = "전체"
    choices = [ALL] + list(options)
    selected = st.multiselect(label, options=choices, default=[ALL], key=key, **kwargs)
    if not selected or ALL in selected:
        return list(options)
    return selected

# ── 사이드바 메뉴 ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dotus")
    st.markdown("---")
    page = st.radio(
        "메뉴",
        ["📈 매출 대시보드", "🔍 상세 데이터 조회", "🧾 주문번호별 조회", "📦 재고소진일정"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    if st.button("🚪 로그아웃", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# ════════════════════════════════════════════════════════
# PAGE 1 : 매출 대시보드
# ════════════════════════════════════════════════════════
if page == "📈 매출 대시보드":

    st.markdown("# 📈 매출 대시보드")
    st.markdown("---")

    with st.expander("🔧 조회 조건", expanded=True):
        r1c1, r1c2, r1c3 = st.columns(3)
        with r1c1:
            months = sorted(df_sales["year_month"].unique())
            month_start = st.selectbox("시작 월", months, index=0)
            month_end   = st.selectbox("종료 월", months, index=len(months)-1)
        with r1c2:
            all_clients = sorted(df_sales["client"].unique())
            selected_clients = multiselect_with_all("채널", all_clients, key="p1_ch")
        with r1c3:
            all_large = sorted(df_sales["large_category"].dropna().unique())
            selected_large = multiselect_with_all("대분류", all_large, key="p1_lg")
            mid_pool = df_sales[df_sales["large_category"].isin(selected_large)]["medium_category"].dropna().unique()
            selected_medium = multiselect_with_all("중분류", sorted(mid_pool), key="p1_md")
            selected_small = None
            if HAS_SMALL:
                small_pool = df_sales[
                    df_sales["large_category"].isin(selected_large) &
                    df_sales["medium_category"].isin(selected_medium)
                ]["small_category"].dropna().unique()
                selected_small = multiselect_with_all("소분류", sorted(small_pool), key="p1_sm")

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
        ps = prev_months[max(0, prev_months.index(month_start)-1)]
        pe = prev_months[max(0, prev_months.index(month_end)-1)]
        prev_f = df_sales[
            (df_sales["year_month"] >= ps) & (df_sales["year_month"] <= pe) &
            (df_sales["client"].isin(selected_clients)) &
            (df_sales["large_category"].isin(selected_large))
        ]
    else:
        prev_f = pd.DataFrame()

    total_sales  = filtered["net_sales"].sum()
    total_qty    = filtered["net_qty"].sum()
    total_orders = filtered["order_no_1"].nunique()
    avg_order    = total_sales / total_orders if total_orders > 0 else 0

    if not prev_f.empty:
        pss = prev_f["net_sales"].sum(); pq = prev_f["net_qty"].sum(); po = prev_f["order_no_1"].nunique()
        sales_d  = f"{((total_sales -pss)/pss*100):+.1f}%" if pss else "N/A"
        qty_d    = f"{((total_qty  -pq )/pq *100):+.1f}%" if pq  else "N/A"
        orders_d = f"{((total_orders-po )/po *100):+.1f}%" if po  else "N/A"
    else:
        sales_d = qty_d = orders_d = None

    k1,k2,k3,k4 = st.columns(4)
    with k1: st.metric("결제금액",      f"₩{total_sales/1e8:.1f}억",  delta=sales_d)
    with k2: st.metric("판매 수량",     f"{total_qty:,}개",            delta=qty_d)
    with k3: st.metric("주문 건수",     f"{total_orders:,}건",         delta=orders_d)
    with k4: st.metric("평균 주문금액", f"₩{avg_order/1e4:.0f}만")
    st.markdown("---")

    st.markdown("### 월별 결제금액 추이")
    sales_metric = st.radio("기준", ["금액 (억원)", "수량 (개)"], horizontal=True, key="sm1")
    monthly = filtered.groupby("year_month").agg(결제금액=("net_sales","sum"), 수량=("net_qty","sum")).reset_index()
    use_m = (sales_metric == "금액 (억원)")
    y_v   = monthly["결제금액"]/1e8 if use_m else monthly["수량"]
    y_fmt = [f"₩{v:.1f}억" for v in y_v] if use_m else [f"{v:,}개" for v in y_v]
    monthly["MoM"] = y_v.pct_change() * 100

    fig_sales = make_subplots(specs=[[{"secondary_y": True}]])
    fig_sales.add_trace(go.Scatter(
        x=monthly["year_month"], y=y_v, name="결제금액",
        mode="lines+markers", line=dict(color=CYAN, width=3, shape="spline"),
        marker=dict(size=7, color=CYAN), fill="tozeroy", fillcolor="rgba(0,188,212,0.12)",
        text=y_fmt, hovertemplate="%{text}<extra></extra>"
    ), secondary_y=False)
    fig_sales.add_trace(go.Scatter(
        x=monthly["year_month"], y=monthly["MoM"], name="MoM(%)",
        mode="lines+markers", line=dict(color=GOLD, width=2, dash="dot", shape="spline"),
        marker=dict(size=5, color=GOLD), hovertemplate="%{y:.1f}%<extra></extra>"
    ), secondary_y=True)
    fig_sales.update_layout(height=340, **nl(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color=TEXT2))))
    fig_sales.update_xaxes(type="category", **AXIS)
    fig_sales.update_yaxes(title_text="억원" if use_m else "수량", secondary_y=False, **AXIS)
    fig_sales.update_yaxes(title_text="MoM (%)", secondary_y=True, **AXIS)
    st.plotly_chart(fig_sales, use_container_width=True)

    cl, cr = st.columns(2)
    with cl:
        st.markdown("### 채널별 결제금액")
        ch_df = filtered.groupby("client").agg(결제금액=("net_sales","sum")).sort_values("결제금액", ascending=True).reset_index()
        fig_ch = go.Figure(go.Bar(
            x=ch_df["결제금액"]/1e8, y=ch_df["client"], orientation="h",
            marker=dict(color=ch_df["결제금액"], colorscale=[[0,CARD2],[1,CYAN]], showscale=False),
            text=[f"₩{v/1e8:.1f}억" for v in ch_df["결제금액"]],
            textposition="outside", textfont=dict(color=TEXT2)
        ))
        fig_ch.update_layout(height=380, **nl(margin=dict(t=20,b=20,l=10,r=80)))
        fig_ch.update_xaxes(title_text="억원", **AXIS)
        fig_ch.update_yaxes(**AXIS)
        st.plotly_chart(fig_ch, use_container_width=True)

    with cr:
        st.markdown("### 상품 중분류별 결제금액")
        mid_df = filtered.groupby("medium_category").agg(결제금액=("net_sales","sum")).sort_values("결제금액", ascending=False).head(10).reset_index()
        fig_mid = go.Figure(go.Pie(
            values=mid_df["결제금액"], labels=mid_df["medium_category"],
            marker=dict(colors=PIE_COLORS, line=dict(color=BG, width=2)),
            textposition="inside", textinfo="percent+label",
            textfont=dict(size=11, color=BG), hole=0.4
        ))
        fig_mid.update_layout(height=380, plot_bgcolor=CARD, paper_bgcolor=CARD,
                              font=dict(color=TEXT2), margin=dict(t=20,b=20,l=10,r=10), showlegend=False)
        st.plotly_chart(fig_mid, use_container_width=True)

    st.markdown("### 채널별 월별 결제금액 흐름")
    ch_monthly = filtered.groupby(["year_month","client"])["net_sales"].sum().reset_index()
    top6 = ch_df.sort_values("결제금액", ascending=False).head(6)["client"].tolist()
    fig_multi = go.Figure()
    for i, client in enumerate(top6):
        df_c = ch_monthly[ch_monthly["client"]==client]
        color = CHART_COLORS[i % len(CHART_COLORS)]
        r,g,b = int(color[1:3],16), int(color[3:5],16), int(color[5:7],16)
        fig_multi.add_trace(go.Scatter(
            x=df_c["year_month"], y=df_c["net_sales"]/1e8, name=client,
            mode="lines+markers", line=dict(color=color, width=2, shape="spline"),
            marker=dict(size=5, color=color), fill="tozeroy", fillcolor=f"rgba({r},{g},{b},0.06)",
            hovertemplate=f"{client}: ₩%{{y:.1f}}억<extra></extra>"
        ))
    fig_multi.update_layout(height=340, **nl(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color=TEXT2))))
    fig_multi.update_xaxes(type="category", **AXIS)
    fig_multi.update_yaxes(title_text="억원", **AXIS)
    st.plotly_chart(fig_multi, use_container_width=True)

    st.markdown("### 채널 × 월별 결제금액 히트맵")
    piv = filtered.groupby(["client","year_month"])["net_sales"].sum().reset_index()
    piv_t = piv.pivot(index="client", columns="year_month", values="net_sales").fillna(0)
    fig_heat = go.Figure(go.Heatmap(
        z=piv_t.values/1e8, x=piv_t.columns.tolist(), y=piv_t.index.tolist(),
        colorscale=[[0,CARD2],[0.5,"#00838F"],[1,CYAN]],
        text=[[f"₩{v:.1f}억" for v in row] for row in piv_t.values/1e8],
        texttemplate="%{text}", textfont=dict(color=TEXT, size=11),
        colorbar=dict(title=dict(text="억원", font=dict(color=TEXT2)), tickfont=dict(color=TEXT2))
    ))
    fig_heat.update_layout(height=340, plot_bgcolor=CARD, paper_bgcolor=CARD,
                           font=dict(color=TEXT2), margin=dict(t=20,b=20,l=10,r=10),
                           xaxis=dict(tickfont=dict(color=TEXT2)), yaxis=dict(tickfont=dict(color=TEXT2)))
    st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("---")

    sv_opt = st.radio("매출 표시 기준", ["금액 (억원)", "수량 (개)"], horizontal=True, key="sv2")
    lc, rc = st.columns(2)
    CHART_H = 300

    with lc:
        st.markdown("### 월별 결제금액")
        sm2 = filtered.groupby("year_month").agg(결제금액=("net_sales","sum"), 수량=("net_qty","sum")).reset_index()
        use_m2 = (sv_opt == "금액 (억원)")
        sv_y  = sm2["결제금액"]/1e8 if use_m2 else sm2["수량"]
        sv_fmt= [f"₩{v:.1f}억" for v in sv_y] if use_m2 else [f"{v:,}개" for v in sv_y]
        fig_sv = go.Figure(go.Scatter(
            x=sm2["year_month"], y=sv_y, mode="lines+markers",
            line=dict(color=CYAN, width=2.5, shape="spline"), marker=dict(size=7, color=CYAN),
            fill="tozeroy", fillcolor="rgba(0,188,212,0.10)",
            text=sv_fmt, hovertemplate="%{text}<extra></extra>"
        ))
        fig_sv.update_layout(height=CHART_H, **nl())
        fig_sv.update_xaxes(type="category", **AXIS)
        fig_sv.update_yaxes(title_text="억원" if use_m2 else "수량", **AXIS)
        st.plotly_chart(fig_sv, use_container_width=True)

    with rc:
        st.markdown("### 월별 입고 수량")
        stk = df_stock[
            (df_stock["year_month"] >= month_start) &
            (df_stock["year_month"] <= month_end)
        ].groupby("year_month")["quantity"].sum().reset_index()
        stk.columns = ["year_month","입고수량"]
        fig_stk = go.Figure(go.Scatter(
            x=stk["year_month"], y=stk["입고수량"], mode="lines+markers",
            line=dict(color=CYAN2, width=2.5, shape="spline"), marker=dict(size=7, color=CYAN2),
            fill="tozeroy", fillcolor="rgba(77,208,225,0.10)",
            hovertemplate="입고: %{y:,}개<extra></extra>"
        ))
        fig_stk.update_layout(height=CHART_H, **nl())
        fig_stk.update_xaxes(type="category", **AXIS)
        fig_stk.update_yaxes(title_text="수량 (개)", tickformat=",", **AXIS)
        st.plotly_chart(fig_stk, use_container_width=True)

# ════════════════════════════════════════════════════════
# PAGE 2 : 상세 데이터 조회
# ════════════════════════════════════════════════════════
elif page == "🔍 상세 데이터 조회":

    st.markdown("# 🔍 상세 데이터 조회")
    st.markdown("---")

    with st.expander("🔧 조회 조건", expanded=True):
        d1, d2, d3 = st.columns([1,1,2])
        with d1:
            date_from = st.date_input("시작일", value=df_sales["date"].min().date(),
                                      min_value=df_sales["date"].min().date(),
                                      max_value=df_sales["date"].max().date())
        with d2:
            date_to = st.date_input("종료일", value=df_sales["date"].max().date(),
                                    min_value=df_sales["date"].min().date(),
                                    max_value=df_sales["date"].max().date())
        with d3:
            det_clients = multiselect_with_all("채널", sorted(df_sales["client"].unique()), key="det_ch")
        keyword = st.text_input("상품명 검색", placeholder="예: 오딧  →  오딧 캐리어, 오딧백 등 모두 검색")

    detail_df = df_sales[
        (df_sales["date"].dt.date >= date_from) &
        (df_sales["date"].dt.date <= date_to) &
        (df_sales["client"].isin(det_clients))
    ].copy()
    if keyword.strip():
        for kw in keyword.strip().split():
            detail_df = detail_df[detail_df["product_name"].str.contains(kw, case=False, na=False)]

    # 조회 일수
    days = (date_to - date_from).days + 1
    total_sales_sum = detail_df["net_sales"].sum()
    total_qty_sum   = detail_df["net_qty"].sum()
    avg_qty_day     = total_qty_sum / days
    avg_sales_day   = total_sales_sum / days

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: st.metric("조회 건수",      f"{len(detail_df):,}건")
    with c2: st.metric("결제금액 합계",  f"₩{total_sales_sum/1e8:.2f}억")
    with c3: st.metric("판매 수량",      f"{total_qty_sum:,}개")
    with c4: st.metric("일평균 출고량",  f"{avg_qty_day:.1f}개")
    with c5: st.metric("일평균 결제금액", f"₩{avg_sales_day/1e4:.0f}만")

    st.markdown("---")

    st.markdown("### 조회 기간 결제금액 흐름")
    spark = detail_df.groupby("year_month")["net_sales"].sum().reset_index().sort_values("year_month")
    if not spark.empty:
        fig_sp = go.Figure(go.Scatter(
            x=spark["year_month"], y=spark["net_sales"]/1e8,
            mode="lines+markers", line=dict(color=CYAN, width=3, shape="spline"),
            marker=dict(size=7, color=CYAN, line=dict(color=BG, width=2)),
            fill="tozeroy", fillcolor="rgba(0,188,212,0.12)",
            text=[f"₩{v:.2f}억" for v in spark["net_sales"]/1e8],
            hovertemplate="%{x}: %{text}<extra></extra>"
        ))
        fig_sp.update_layout(height=220, **nl(margin=dict(t=20,b=20,l=10,r=10)))
        fig_sp.update_xaxes(type="category", **AXIS)
        fig_sp.update_yaxes(title_text="억원", **AXIS)
        st.plotly_chart(fig_sp, use_container_width=True)

    st.markdown("---")

    st.markdown("### 상품별 기간 합산")
    sort_by = st.radio("정렬 기준", ["결제금액", "판매수량"], horizontal=True)

    agg = detail_df.groupby("product_name").agg(
        결제금액=("net_sales","sum"), 판매수량=("net_qty","sum"),
        주문건수=("order_no_1","nunique"), 환불수량=("return_quantity","sum"),
        환불금액=("return_sales_amount","sum")
    ).reset_index().rename(columns={"product_name":"상품명"})

    agg["평균출고량"]   = (agg["판매수량"] / days).round(1)
    agg["평균결제금액"] = (agg["결제금액"] / days).round(0).astype(int)

    agg = agg.sort_values(sort_by, ascending=False).reset_index(drop=True)
    agg.index += 1

    disp = agg.copy()
    disp["결제금액"]   = disp["결제금액"].apply(lambda x: f"₩{x:,}")
    disp["환불금액"]   = disp["환불금액"].apply(lambda x: f"₩{x:,}")
    disp["평균결제금액"] = disp["평균결제금액"].apply(lambda x: f"₩{x:,}")
    disp = disp[["상품명","결제금액","판매수량","평균출고량","평균결제금액","주문건수","환불수량","환불금액"]]
    st.dataframe(disp, use_container_width=True, hide_index=False, height=480)

# ════════════════════════════════════════════════════════
# PAGE 3 : 주문번호별 조회
# ════════════════════════════════════════════════════════
elif page == "🧾 주문번호별 조회":

    st.markdown("# 🧾 주문번호별 조회")
    st.markdown("---")

    with st.expander("🔧 조회 조건", expanded=True):
        o1, o2, o3, o4 = st.columns([1,1,2,2])
        with o1:
            o_from = st.date_input("시작일", value=df_sales["date"].min().date(),
                                   min_value=df_sales["date"].min().date(),
                                   max_value=df_sales["date"].max().date(), key="o_from")
        with o2:
            o_to = st.date_input("종료일", value=df_sales["date"].max().date(),
                                 min_value=df_sales["date"].min().date(),
                                 max_value=df_sales["date"].max().date(), key="o_to")
        with o3:
            o_clients = multiselect_with_all("채널", sorted(df_sales["client"].unique()), key="o_ch")
        with o4:
            order_kw = st.text_input("주문번호 검색", placeholder="주문번호 일부만 입력해도 검색 가능")

    ord_df = df_sales[
        (df_sales["date"].dt.date >= o_from) &
        (df_sales["date"].dt.date <= o_to) &
        (df_sales["client"].isin(o_clients))
    ].copy()
    if order_kw.strip():
        ord_df = ord_df[ord_df["order_no_1"].astype(str).str.contains(order_kw.strip(), case=False, na=False)]

    ord_df["order_no_1_str"] = ord_df["order_no_1"].astype(str).str.strip()
    unknown_mask = ord_df["order_no_1_str"].isin(["-", "", "nan", "None"])
    ord_known = ord_df[~unknown_mask].copy()
    ord_unknown = ord_df[unknown_mask].copy()

    cnt_per_order = ord_known.groupby("order_no_1_str").size().reset_index(name="item_count")
    single_orders = cnt_per_order[cnt_per_order["item_count"] == 1]["order_no_1_str"]
    multi_orders  = cnt_per_order[cnt_per_order["item_count"] >= 2]["order_no_1_str"]

    n_single  = len(single_orders)
    n_multi   = len(multi_orders)
    n_unknown = ord_unknown["order_no_1_str"].nunique()
    n_total   = n_single + n_multi + n_unknown

    k1,k2,k3,k4 = st.columns(4)
    with k1: st.metric("전체 주문", f"{n_total:,}건")
    with k2: st.metric("단포장",   f"{n_single:,}건  ({n_single/n_total*100:.1f}%)" if n_total else "0건")
    with k3: st.metric("합포장",   f"{n_multi:,}건  ({n_multi/n_total*100:.1f}%)"  if n_total else "0건")
    with k4: st.metric("미확인",   f"{n_unknown:,}건")

    st.markdown("---")

    cl_d, cr_d = st.columns([1, 2])

    with cl_d:
        st.markdown("### 포장 유형 비율")
        fig_donut = go.Figure(go.Pie(
            values=[n_single, n_multi, n_unknown],
            labels=["단포장", "합포장", "미확인"],
            marker=dict(colors=[CYAN, GOLD, TEXT2], line=dict(color=BG, width=2)),
            textinfo="percent+label", textfont=dict(size=13, color=BG),
            hole=0.5
        ))
        fig_donut.update_layout(
            height=320, plot_bgcolor=CARD, paper_bgcolor=CARD,
            font=dict(color=TEXT2), margin=dict(t=20,b=20,l=10,r=10),
            showlegend=False,
            annotations=[dict(text=f"{n_total:,}건", x=0.5, y=0.5,
                              font=dict(size=16, color=TEXT), showarrow=False)]
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with cr_d:
        st.markdown("### 월별 단포장 / 합포장 추이")
        ord_known2 = ord_known.copy()
        ord_known2["pack_type"] = ord_known2["order_no_1_str"].isin(multi_orders).map({True:"합포장", False:"단포장"})
        pack_monthly = ord_known2.groupby(["year_month","pack_type"])["order_no_1_str"].nunique().reset_index()
        pack_monthly.columns = ["year_month","pack_type","건수"]

        fig_pack = go.Figure()
        for pt, color in [("단포장", CYAN), ("합포장", GOLD)]:
            df_pt = pack_monthly[pack_monthly["pack_type"]==pt]
            fig_pack.add_trace(go.Scatter(
                x=df_pt["year_month"], y=df_pt["건수"], name=pt,
                mode="lines+markers", line=dict(color=color, width=2.5, shape="spline"),
                marker=dict(size=6, color=color),
                hovertemplate=f"{pt}: %{{y:,}}건<extra></extra>"
            ))
        fig_pack.update_layout(height=320, **nl(legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color=TEXT2))))
        fig_pack.update_xaxes(type="category", **AXIS)
        fig_pack.update_yaxes(tickformat=",", **AXIS)
        st.plotly_chart(fig_pack, use_container_width=True)

    st.markdown("---")

    st.markdown("### 합포장 상품 조합 분석")
    multi_df = ord_known[ord_known["order_no_1_str"].isin(multi_orders)].copy()

    @st.cache_data
    def calc_combos(df_input):
        combo_counts = {}
        for order_id, group in df_input.groupby("order_no_1_str"):
            prods = sorted(group["product_name"].dropna().unique().tolist())
            for combo in combinations(prods, 2):
                combo_counts[combo] = combo_counts.get(combo, 0) + 1
        if not combo_counts:
            return pd.DataFrame(columns=["상품 A","상품 B","합포 건수"])
        combo_df = pd.DataFrame(
            [(a, b, cnt) for (a,b), cnt in combo_counts.items()],
            columns=["상품 A","상품 B","합포 건수"]
        ).sort_values("합포 건수", ascending=False).reset_index(drop=True)
        combo_df.index += 1
        return combo_df

    combo_df = calc_combos(multi_df)

    tab_a, tab_b = st.tabs(["📋 전체 조합 순위", "🔍 특정 상품 기준 합포 순위"])

    with tab_a:
        st.markdown(f"<span style='color:{TEXT2};font-size:0.85rem;'>합포장 주문에서 함께 주문된 상품 조합 순위 (상위 100개)</span>", unsafe_allow_html=True)
        st.dataframe(combo_df.head(100), use_container_width=True, hide_index=False, height=450)

    with tab_b:
        if not multi_df.empty:
            all_products = sorted(multi_df["product_name"].dropna().unique().tolist())
            search_kw = st.text_input("상품명 검색", placeholder="예: 오딧 캐리어", key="prod_search")
            filtered_products = [p for p in all_products if search_kw.strip().lower() in p.lower()] if search_kw.strip() else all_products

            if filtered_products:
                sel_product = st.selectbox("상품 선택", options=filtered_products, key="prod_select")
                paired = combo_df[
                    (combo_df["상품 A"] == sel_product) | (combo_df["상품 B"] == sel_product)
                ].copy()
                paired["함께 합포된 상품"] = paired.apply(
                    lambda r: r["상품 B"] if r["상품 A"] == sel_product else r["상품 A"], axis=1
                )
                paired = paired[["함께 합포된 상품","합포 건수"]].reset_index(drop=True)
                paired.index += 1
                st.markdown(f"**{sel_product}** 와(과) 함께 합포된 상품 순위")
                st.dataframe(paired, use_container_width=True, hide_index=False, height=450)
            else:
                st.warning("검색 결과가 없습니다.")
        else:
            st.info("합포장 데이터가 없습니다.")

# ════════════════════════════════════════════════════════
# PAGE 4 : 재고소진일정
# ════════════════════════════════════════════════════════
elif page == "📦 재고소진일정":

    st.markdown("# 📦 재고소진일정")
    st.markdown("---")

    today = date.today()

    with st.expander("🔧 조회 조건", expanded=True):
        ec1, ec2 = st.columns([3, 2])
        with ec1:
            preset = st.radio(
                "일평균 계산 기간",
                ["이번달", "지난달", "최근 30일", "최근 90일", "직접 입력"],
                horizontal=True, key="inv_preset"
            )
        with ec2:
            inv_file = st.file_uploader("재고파일 업로드 (.xlsx)", type=["xlsx"], key="inv_upload")

        # 기간 계산
        if preset == "이번달":
            period_from = today.replace(day=1)
            period_to   = today
        elif preset == "지난달":
            first_this  = today.replace(day=1)
            period_to   = first_this - timedelta(days=1)
            period_from = period_to.replace(day=1)
        elif preset == "최근 30일":
            period_from = today - timedelta(days=30)
            period_to   = today
        elif preset == "최근 90일":
            period_from = today - timedelta(days=90)
            period_to   = today
        else:  # 직접 입력
            dc1, dc2 = st.columns(2)
            with dc1:
                period_from = st.date_input("시작일", value=today - timedelta(days=30),
                                            min_value=df_sales["date"].min().date(),
                                            max_value=today, key="inv_from")
            with dc2:
                period_to = st.date_input("종료일", value=today,
                                          min_value=df_sales["date"].min().date(),
                                          max_value=today, key="inv_to")

        st.caption(f"📅 일평균 계산 기간: {period_from} ~ {period_to}  ({(period_to - period_from).days + 1}일)")

    # ── 일평균결제수량 계산 ───────────────────────────────
    days_avg = (period_to - period_from).days + 1
    sales_period = df_sales[
        (df_sales["date"].dt.date >= period_from) &
        (df_sales["date"].dt.date <= period_to)
    ]
    avg_by_pid = sales_period.groupby("product_id").agg(
        판매수량합=("net_qty", "sum")
    ).reset_index()
    avg_by_pid["product_id"] = avg_by_pid["product_id"].astype(str).str.strip()
    avg_by_pid["일평균결제수량"] = (avg_by_pid["판매수량합"] / days_avg).round(2)

    # ── 재고파일 처리 ─────────────────────────────────────
    if inv_file is None:
        st.info("👆 재고파일(.xlsx)을 업로드하면 소진 일정이 표시됩니다.")
        st.markdown(f"""
        <div style='background:{CARD};border:1px dashed {BORDER};border-radius:8px;padding:20px;margin-top:12px;'>
            <p style='color:{TEXT2};margin:0;font-size:0.9rem;'>
            📋 업로드 파일 형식: <span style='color:{CYAN};'>상품명 / 품번 / 대분류 / 중분류 / 소분류 / 재고수량</span>
            </p>
        </div>""", unsafe_allow_html=True)
    else:
        # 파일 읽기
        inv_wb = openpyxl.load_workbook(inv_file, read_only=True, data_only=True)
        ws_inv = inv_wb.active
        inv_rows = list(ws_inv.iter_rows(values_only=True))
        inv_data = [r[:6] for r in inv_rows[1:] if r[0] is not None]
        inv_df = pd.DataFrame(inv_data, columns=["상품명","product_id","대분류","중분류","소분류","재고수량"])
        inv_df["product_id"] = inv_df["product_id"].astype(str).str.strip()
        inv_df["재고수량"] = pd.to_numeric(inv_df["재고수량"], errors="coerce").fillna(0).astype(int)

        # 매출 데이터와 조인
        result = inv_df.merge(avg_by_pid[["product_id","일평균결제수량"]], on="product_id", how="left")

        # 소진 계산
        def calc_days(row):
            if pd.isna(row["일평균결제수량"]) or row["일평균결제수량"] <= 0:
                return None
            return int(row["재고수량"] / row["일평균결제수량"])

        result["잔여일수"] = result.apply(calc_days, axis=1)
        result["소진예정일"] = result["잔여일수"].apply(
            lambda d: (today + timedelta(days=int(d))).strftime("%Y-%m-%d") if d is not None else "-"
        )
        result["일평균결제수량_표시"] = result["일평균결제수량"].apply(
            lambda v: f"{v:.1f}개/일" if pd.notna(v) else "결제수량 없음"
        )
        result["잔여일수_표시"] = result["잔여일수"].apply(
            lambda d: f"{d:,}일" if d is not None else "-"
        )

        # 상태 컬럼
        def get_status(row):
            if row["잔여일수"] is None:
                return "⬜ 결제수량 없음"
            d = row["잔여일수"]
            if d <= 7:
                return "🔴 긴급"
            elif d <= 30:
                return "🟡 주의"
            else:
                return "🟢 여유"

        result["상태"] = result.apply(get_status, axis=1)

        # 정렬: 긴급→주의→여유→결제수량없음, 각 그룹 내 잔여일수 오름차순
        status_order = {"🔴 긴급": 0, "🟡 주의": 1, "🟢 여유": 2, "⬜ 결제수량 없음": 3}
        result["_sort"] = result["상태"].map(status_order)
        result = result.sort_values(["_sort", "잔여일수"]).drop(columns="_sort").reset_index(drop=True)
        result.index += 1

        # ── 요약 스코어카드 ───────────────────────────────
        n_urgent  = (result["상태"] == "🔴 긴급").sum()
        n_caution = (result["상태"] == "🟡 주의").sum()
        n_safe    = (result["상태"] == "🟢 여유").sum()
        n_nodata  = (result["상태"] == "⬜ 결제수량 없음").sum()

        k1,k2,k3,k4 = st.columns(4)
        with k1: st.metric("🔴 긴급 (7일 이내)",   f"{n_urgent:,}개")
        with k2: st.metric("🟡 주의 (30일 이내)",   f"{n_caution:,}개")
        with k3: st.metric("🟢 여유 (30일 초과)",   f"{n_safe:,}개")
        with k4: st.metric("⬜ 결제수량 없음",       f"{n_nodata:,}개")

        st.markdown("---")

        # ── 결과 테이블 ───────────────────────────────────
        st.markdown(f"### 재고 소진 예정 현황  <span style='color:{TEXT2};font-size:0.85rem;'>전체 {len(result):,}개 품목</span>", unsafe_allow_html=True)

        disp = result[["상품명","product_id","대분류","중분류","소분류","재고수량","일평균결제수량_표시","잔여일수_표시","소진예정일","상태"]].copy()
        disp.columns = ["상품명","품번","대분류","중분류","소분류","재고수량","일평균결제수량","잔여일수","소진예정일","상태"]

        st.dataframe(disp, use_container_width=True, hide_index=False, height=600)

        # ── 긴급 품목 별도 강조 ───────────────────────────
        urgent_df = result[result["상태"] == "🔴 긴급"]
        if not urgent_df.empty:
            st.markdown("---")
            st.markdown(f"### 🔴 긴급 소진 임박 품목 ({n_urgent}개)")
            urg_disp = urgent_df[["상품명","product_id","재고수량","일평균결제수량_표시","잔여일수_표시","소진예정일"]].copy()
            urg_disp.columns = ["상품명","품번","재고수량","일평균결제수량","잔여일수","소진예정일"]
            urg_disp = urg_disp.reset_index(drop=True)
            urg_disp.index += 1
            st.dataframe(urg_disp, use_container_width=True, hide_index=False)

st.markdown("---")
st.caption("Dotus Dashboard · 데이터 기준: 제공된 엑셀 파일")
