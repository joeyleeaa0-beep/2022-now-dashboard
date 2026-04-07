import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
from io import BytesIO
import datetime
import os

st.set_page_config(page_title="新媒体年度数据看板", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #f8f9fc; }
    #MainMenu, footer { visibility: hidden; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; background: transparent; border-bottom: 1px solid #eef0f4; }
    .stTabs [data-baseweb="tab"] { background: transparent; color: #6b7280; font-weight: 500; border-radius: 6px 6px 0 0; padding: 8px 18px; }
    .stTabs [aria-selected="true"] { background: white !important; color: #111827 !important; border-bottom: 2px solid #4f46e5 !important; }
    [data-testid="stSidebar"] { background: white; border-right: 1px solid #eef0f4; }
    hr { border-color: #eef0f4; }
</style>
""", unsafe_allow_html=True)

APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = "Jf7osfcF7h2pzFt945XcUkCwned"
SHEET_ID = "0piesb"
YEARS = ["2022", "2023", "2024", "2025", "2026"]
MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
CITIES = ["深圳", "上海", "成都", "天津"]
COLORS = ["#4f46e5","#06b6d4","#10b981","#f59e0b","#ef4444","#8b5cf6"]
PLOTLY_LAYOUT = dict(
    paper_bgcolor="white", plot_bgcolor="white",
    font=dict(family="sans-serif", size=12, color="#374151"),
    margin=dict(l=16, r=16, t=48, b=16),
    legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=11)),
    xaxis=dict(showgrid=False, linecolor="#eef0f4"),
    yaxis=dict(gridcolor="#f3f4f6", linecolor="#eef0f4"),
)

@st.cache_resource
def setup_chinese_font():
    font_path = "/tmp/NotoSansSC.ttf"
    if not os.path.exists(font_path):
        for url in [
            "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/SubsetOTF/SC/NotoSansSC-Regular.otf",
            "https://github.com/googlefonts/noto-cjk/raw/main/Sans/SubsetOTF/SC/NotoSansSC-Regular.otf",
        ]:
            try:
                r = requests.get(url, timeout=30)
                if r.status_code == 200:
                    with open(font_path, "wb") as f:
                        f.write(r.content)
                    break
            except Exception:
                continue
    if os.path.exists(font_path):
        try:
            fm.fontManager.addfont(font_path)
            prop = fm.FontProperties(fname=font_path)
            return prop.get_name(), font_path
        except Exception:
            pass
    return None, None

CHINESE_FONT_NAME, CHINESE_FONT_PATH = setup_chinese_font()

@st.cache_data(ttl=60)
def get_token():
    res = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return res.json().get("tenant_access_token")

def extract_text(val):
    if val is None:
        return ""
    if isinstance(val, list):
        return "".join(item.get("text", "") if isinstance(item, dict) else str(item) for item in val)
    return str(val).strip()

def extract_value(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, list):
        return extract_text(val)
    s = str(val).strip()
    if "(" in s and ")" in s and any(c.isalpha() for c in s[:5]):
        return 0
    if "+" in s and s[0].isalpha():
        return 0
    return s

@st.cache_data(ttl=60)
def read_sheet():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}"
        f"/values/{SHEET_ID}!A1:AG2000?renderType=FORMATTED_VALUE"
    )
    res = requests.get(url, headers=headers).json()
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers_row = [
        extract_text(h) if isinstance(h, list) else (str(h).strip() if h else f"列{i}")
        for i, h in enumerate(values[0])
    ]
    rows = []
    for row in values[1:]:
        new_row = [extract_value(cell) for cell in row]
        while len(new_row) < len(headers_row):
            new_row.append("")
        rows.append(new_row[:len(headers_row)])
    return pd.DataFrame(rows, columns=headers_row)

def to_num(series):
    if isinstance(series, pd.Series):
        s = series.astype(str).str.replace(",", "").str.strip()
        return pd.to_numeric(s, errors="coerce").fillna(0)
    return 0

def make_chart(fig):
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def safe_agg(df, group_col, agg_dict):
    g = df.groupby(group_col).agg(**agg_dict).reset_index()
    if "总花费" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["客资成本"] = 0.0
        g.loc[mask, "客资成本"] = (g.loc[mask, "总花费"] / g.loc[mask, "客资总数"]).round(2)
    if "总花费" in g.columns and "收销总量" in g.columns:
        mask = g["收销总量"] > 0
        g["成交成本"] = 0.0
        g.loc[mask, "成交成本"] = (g.loc[mask, "总花费"] / g.loc[mask, "收销总量"]).round(2)
    if "到店总量" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["到店率%"] = 0.0
        g.loc[mask, "到店率%"] = (g.loc[mask, "到店总量"] / g.loc[mask, "客资总数"] * 100).round(2)
    if "收销总量" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["成交率%"] = 0.0
        g.loc[mask, "成交率%"] = (g.loc[mask, "收销总量"] / g.loc[mask, "客资总数"] * 100).round(2)
    return g

@st.cache_data(ttl=60)
def clean_df():
    df = read_sheet()
    if df.empty:
        return pd.DataFrame()

    # 重命名列
    rename_map = {
        "收销总量/台": "收销总量原始",
        "销售量/台": "销售量",
        "收购量/台": "收购量",
        "视频成交/台": "视频成交",
        "直播成交/台": "直播成交",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # 年份月份清洗
    if "年份" in df.columns:
        df["年份"] = df["年份"].astype(str).str.strip().str.replace(".0", "", regex=False)
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip()
        df["月份"] = df["月份"].apply(lambda x: x + "月" if x.isdigit() else x)

    # 数值列转换
    skip_cols = ["城市", "年份", "月份"]
    for col in df.columns:
        if col not in skip_cols:
            df[col] = to_num(df[col])

    # ── 公式列用子列加总重新计算 ──

    # 总花费
    花费列 = ["抖音账号花费", "信息花费", "微信花费", "小红书花费", "其他平台花费"]
    有效花费列 = [c for c in 花费列 if c in df.columns]
    if 有效花费列:
        df["总花费"] = df[有效花费列].sum(axis=1)

    # 客资总数
    客资列 = ["直播客资数", "视频客资数", "抖音号客资", "信息流客资数",
              "微信客资客资", "小红书客资客资", "b站客资", "快手客资"]
    有效客资列 = [c for c in 客资列 if c in df.columns]
    if 有效客资列:
        df["客资总数"] = df[有效客资列].sum(axis=1)

    # 到店总量
    到店列 = ["销售到店", "收购到店"]
    有效到店列 = [c for c in 到店列 if c in df.columns]
    if 有效到店列:
        df["到店总量"] = df[有效到店列].sum(axis=1)

    # 收销总量：优先用原始列（2022年手填数字），为0时用销售量+收购量，再为0用视频+直播成交
    if "收销总量原始" in df.columns:
        df["收销总量"] = to_num(df["收销总量原始"])
    else:
        df["收销总量"] = 0

    if "销售量" in df.columns and "收购量" in df.columns:
        mask = df["收销总量"] == 0
        df.loc[mask, "收销总量"] = df.loc[mask, "销售量"] + df.loc[mask, "收购量"]

    if "视频成交" in df.columns and "直播成交" in df.columns:
        mask = df["收销总量"] == 0
        df.loc[mask, "收销总量"] = df.loc[mask, "视频成交"] + df.loc[mask, "直播成交"]

    # 过滤有效城市和年份
    if "城市" in df.columns:
        df = df[df["城市"].isin(CITIES)].copy()
        df["城市"] = pd.Categorical(df["城市"], categories=CITIES, ordered=True)
        df = df.sort_values("城市")
    if "年份" in df.columns:
        df = df[df["年份"].isin(YEARS)].copy()

    return df

def apply_filter(df, cities, years, month):
    d = df.copy()
    if cities:
        d = d[d["城市"].isin(cities)]
    if years:
        d = d[d["年份"].isin(years)]
    if month != "全部月份" and "月份" in d.columns:
        d = d[d["月份"] == month]
    return d

def metric_html(label, value):
    return f"""<div style="background:white;border:1px solid #eef0f4;border-radius:12px;
        padding:20px 24px;box-shadow:0 1px 4px rgba(0,0,0,0.05);margin-bottom:8px;">
        <div style="color:#6b7280;font-size:13px;font-weight:500;margin-bottom:6px;">{label}</div>
        <div style="color:#111827;font-size:26px;font-weight:700;line-height:1.2;">{value}</div>
    </div>"""

# ── 加载数据 ──
with st.spinner("正在加载数据..."):
    try:
        df = clean_df()
        if df.empty:
            st.error("数据加载失败或表格为空")
            st.stop()
    except Exception as e:
        st.error(f"加载失败：{e}")
        st.stop()

# ── 侧边栏 ──
with st.sidebar:
    st.markdown("## 筛选条件")
    sel_cities = st.multiselect("城市（可多选）", CITIES, default=CITIES)
    sel_years = st.multiselect("年份（可多选）", YEARS, default=YEARS)
    sel_month = st.selectbox("月份", ["全部月份"] + MONTHS)
    st.divider()
    df_filtered = apply_filter(df, sel_cities, sel_years, sel_month)
    st.caption(f"当前数据：{len(df_filtered)} 条")

# ── 页面标题 ──
st.markdown(f"""
<div style="padding:8px 0 20px 0;border-bottom:1px solid #eef0f4;margin-bottom:24px;">
    <h2 style="margin:0;color:#111827;font-weight:700;">📊 新媒体年度数据看板</h2>
    <p style="margin:4px 0 0 0;color:#6b7280;font-size:14px;">
        城市：{'、'.join(sel_cities) if sel_cities else '未选择'} ·
        年份：{'、'.join(sel_years) if sel_years else '未选择'} ·
        {sel_month} · 数据每60秒自动更新
    </p>
</div>
""", unsafe_allow_html=True)

# ── 核心指标 ──
total_spend     = df_filtered["总花费"].sum() if "总花费" in df_filtered.columns else 0
total_keizi     = df_filtered["客资总数"].sum() if "客资总数" in df_filtered.columns else 0
total_daodian   = df_filtered["到店总量"].sum() if "到店总量" in df_filtered.columns else 0
total_chengjiao = df_filtered["收销总量"].sum() if "收销总量" in df_filtered.columns else 0
total_xiaoshou  = df_filtered["销售量"].sum() if "销售量" in df_filtered.columns else 0
total_shougou   = df_filtered["收购量"].sum() if "收购量" in df_filtered.columns else 0
keizi_cost      = total_spend / total_keizi if total_keizi > 0 else 0
chengjiao_cost  = total_spend / total_chengjiao if total_chengjiao > 0 else 0
daodian_rate    = total_daodian / total_keizi * 100 if total_keizi > 0 else 0
chengjiao_rate  = total_chengjiao / total_keizi * 100 if total_keizi > 0 else 0

c1,c2,c3,c4 = st.columns(4)
c1.markdown(metric_html("总花费", f"¥{total_spend:,.0f}"), unsafe_allow_html=True)
c2.markdown(metric_html("总客资量", f"{int(total_keizi):,}"), unsafe_allow_html=True)
c3.markdown(metric_html("到店总量", f"{int(total_daodian):,}"), unsafe_allow_html=True)
c4.markdown(metric_html("总成交量", f"{int(total_chengjiao):,}"), unsafe_allow_html=True)
c5,c6,c7,c8 = st.columns(4)
c5.markdown(metric_html("销售量", f"{int(total_xiaoshou):,}"), unsafe_allow_html=True)
c6.markdown(metric_html("收购量", f"{int(total_shougou):,}"), unsafe_allow_html=True)
c7.markdown(metric_html("客资成本", f"¥{keizi_cost:.2f}"), unsafe_allow_html=True)
c8.markdown(metric_html("成交成本", f"¥{chengjiao_cost:.2f}"), unsafe_allow_html=True)
c9,c10,_,_ = st.columns(4)
c9.markdown(metric_html("到店率", f"{daodian_rate:.2f}%"), unsafe_allow_html=True)
c10.markdown(metric_html("成交率", f"{chengjiao_rate:.2f}%"), unsafe_allow_html=True)

st.divider()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏙️ 分城市", "📅 年度对比", "📈 趋势分析", "💰 花费分析", "📋 数据明细"
])

with tab1:
    st.subheader("分城市经营对比")
    if not df_filtered.empty and "城市" in df_filtered.columns:
        cg = safe_agg(df_filtered, "城市", {
            "总花费": ("总花费","sum"),
            "客资总数": ("客资总数","sum"),
            "到店总量": ("到店总量","sum"),
            "收销总量": ("收销总量","sum"),
            "销售量": ("销售量","sum"),
            "收购量": ("收购量","sum"),
        })
        cg["城市"] = pd.Categorical(cg["城市"], categories=CITIES, ordered=True)
        cg = cg.sort_values("城市")
        st.dataframe(cg, use_container_width=True, hide_index=True)
        ca,cb = st.columns(2)
        with ca:
            fig = px.bar(cg,x="城市",y="客资总数",title="各城市客资量",color="城市",
                        color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with cb:
            fig = px.bar(cg,x="城市",y="收销总量",title="各城市成交量",color="城市",
                        color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
            st.plotly_chart(make_chart(fig),use_container_width=True)
        cc,cd = st.columns(2)
        with cc:
            fig = px.bar(cg,x="城市",y="到店总量",title="各城市到店量",color="城市",
                        color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with cd:
            fig = px.bar(cg,x="城市",y="到店率%",title="各城市到店率(%)",color="城市",
                        color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
            st.plotly_chart(make_chart(fig),use_container_width=True)
        ce,cf = st.columns(2)
        with ce:
            fig = px.bar(cg,x="城市",y="客资成本",title="各城市客资成本",color="城市",
                        color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with cf:
            fig = px.bar(cg,x="城市",y="成交成本",title="各城市成交成本",color="城市",
                        color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
            st.plotly_chart(make_chart(fig),use_container_width=True)

with tab2:
    st.subheader("年度数据对比")
    df_year = apply_filter(df, sel_cities, sel_years, "全部月份")
    if not df_year.empty and "年份" in df_year.columns:
        yg = safe_agg(df_year, "年份", {
            "总花费": ("总花费","sum"),
            "客资总数": ("客资总数","sum"),
            "到店总量": ("到店总量","sum"),
            "收销总量": ("收销总量","sum"),
        })
        yg["年份"] = pd.Categorical(yg["年份"], categories=YEARS, ordered=True)
        yg = yg.sort_values("年份")
        st.dataframe(yg, use_container_width=True, hide_index=True)
        ya,yb = st.columns(2)
        with ya:
            fig = px.bar(yg,x="年份",y="客资总数",title="各年度客资量",color="年份",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with yb:
            fig = px.bar(yg,x="年份",y="收销总量",title="各年度成交量",color="年份",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        yc,yd = st.columns(2)
        with yc:
            fig = px.bar(yg,x="年份",y="到店总量",title="各年度到店量",color="年份",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with yd:
            fig = px.bar(yg,x="年份",y="总花费",title="各年度总花费",color="年份",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        ye,yf = st.columns(2)
        with ye:
            fig = px.line(yg,x="年份",y="客资成本",title="各年度客资成本趋势",markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with yf:
            fig = px.line(yg,x="年份",y="到店率%",title="各年度到店率趋势",markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)

with tab3:
    st.subheader("月度趋势分析")
    df_trend = apply_filter(df, sel_cities, sel_years, "全部月份")
    if not df_trend.empty and "月份" in df_trend.columns:
        tm = safe_agg(df_trend, ["年份","月份"], {
            "客资总数": ("客资总数","sum"),
            "到店总量": ("到店总量","sum"),
            "收销总量": ("收销总量","sum"),
            "总花费": ("总花费","sum"),
        })
        tm["月份"] = pd.Categorical(tm["月份"], categories=MONTHS, ordered=True)
        tm = tm.sort_values(["年份","月份"])
        fig = px.line(tm,x="月份",y="客资总数",color="年份",title="各年度客资量月度趋势",
                      markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig),use_container_width=True)
        fig2 = px.line(tm,x="月份",y="收销总量",color="年份",title="各年度成交量月度趋势",
                       markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig2),use_container_width=True)
        fig3 = px.line(tm,x="月份",y="到店率%",color="年份",title="各年度到店率月度趋势",
                       markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig3),use_container_width=True)
        fig4 = px.line(tm,x="月份",y="客资成本",color="年份",title="各年度客资成本月度趋势",
                       markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig4),use_container_width=True)

with tab4:
    st.subheader("投放花费分析")
    spend_cols = {
        "抖音账号": "抖音账号花费",
        "信息流": "信息花费",
        "微信": "微信花费",
        "小红书": "小红书花费",
        "其他平台": "其他平台花费",
    }
    available_spend = {k: v for k, v in spend_cols.items() if v in df_filtered.columns}
    if available_spend:
        spend_data = {k: float(df_filtered[v].sum()) for k, v in available_spend.items()}
        spend_df = pd.DataFrame({"渠道": list(spend_data.keys()), "花费": list(spend_data.values())})
        spend_df = spend_df[spend_df["花费"] > 0]
        if not spend_df.empty:
            sa,sb = st.columns(2)
            with sa:
                fig = px.pie(spend_df,names="渠道",values="花费",title="各渠道花费占比",
                             color_discrete_sequence=COLORS)
                fig.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(make_chart(fig),use_container_width=True)
            with sb:
                fig = px.bar(spend_df,x="渠道",y="花费",title="各渠道花费对比",
                             color="渠道",color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)
    df_spend_trend = apply_filter(df, sel_cities, sel_years, "全部月份")
    if not df_spend_trend.empty and "年份" in df_spend_trend.columns:
        sg = df_spend_trend.groupby("年份").agg(总花费=("总花费","sum")).reset_index()
        sg["年份"] = pd.Categorical(sg["年份"], categories=YEARS, ordered=True)
        sg = sg.sort_values("年份")
        fig = px.line(sg,x="年份",y="总花费",title="各年度总花费趋势",
                      markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig),use_container_width=True)

with tab5:
    st.subheader("数据明细")
    st.dataframe(df_filtered.dropna(axis=1, how='all'), use_container_width=True)
    buf = BytesIO()
    df_filtered.to_excel(buf, index=False, engine='xlsxwriter')
    buf.seek(0)
    st.download_button(
        label="⬇️ 导出 Excel",
        data=buf,
        file_name=f"新媒体数据_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
