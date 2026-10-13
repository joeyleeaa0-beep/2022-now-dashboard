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
import threading
from pathlib import Path

from data_cache import (
    FeishuAPIError,
    MonthlyQuotaExceeded,
    load_snapshot,
    load_with_fallback,
    parse_feishu_response,
)

st.set_page_config(page_title="AKD数据看板", page_icon="📊", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@600;700&display=swap');
    :root {
        --lime: #dcef4e;
        --lime-soft: #ecf7a3;
        --ink: #1c1c1a;
        --ink-2: #6f6f68;
        --line: #e6e6e0;
        --bg: #ececea;
    }
    .stApp { background-color: var(--bg); }
    #MainMenu, footer { visibility: hidden; }
    .block-container { padding-top: 2.2rem; }

    /* 侧边栏 */
    [data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid var(--line); }
    [data-testid="stSidebar"] * { color: var(--ink); }

    /* Tabs：lime 胶囊选中态 */
    .stTabs [data-baseweb="tab-list"] { gap: 6px; background: transparent; border-bottom: 1px solid var(--line); }
    .stTabs [data-baseweb="tab"] {
        background: transparent; color: var(--ink-2); font-weight: 500;
        border-radius: 999px; padding: 6px 16px; font-size: 13px;
    }
    .stTabs [aria-selected="true"] {
        background: var(--lime) !important; color: var(--ink) !important;
        border-bottom: none !important; font-weight: 600;
    }
    .stTabs [data-baseweb="tab-highlight"] { display: none; }
    .stTabs [data-baseweb="tab-border"] { display: none; }

    /* 折叠区 / 表格容器白卡化 */
    [data-testid="stExpander"] {
        background: #ffffff; border: 1px solid var(--line) !important;
        border-radius: 14px !important; overflow: hidden;
    }
    [data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 12px; }
    div[data-testid="stDownloadButton"] button {
        background: #ffffff; color: var(--ink); border: 1px solid var(--line);
        border-radius: 999px; font-weight: 500; padding: 6px 18px;
    }
    div[data-testid="stDownloadButton"] button:hover { border-color: var(--lime); background: #fafce2; }
    hr { border-color: var(--line); }

    /* 通用小节标题：中文 + 英文小注 */
    .mini-label { font-size: 11px; letter-spacing: .14em; color: var(--ink-2);
        text-transform: uppercase; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── 配置 ──
APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = "KOu0s7jKqh81tJtEJIgcwcNXnYf"
SHEET_ID = "0piesb"
YEARS = ["2022", "2023", "2024", "2025", "2026"]
MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
CITIES = ["深圳", "上海", "成都", "天津"]
COLORS = ["#dcef4e","#a8bf2a","#7c941a","#4f6610","#3f3f3a","#b5b5ae","#8a8a82"]
PLOTLY_LAYOUT = dict(
    paper_bgcolor="white", plot_bgcolor="white",
    font=dict(family="sans-serif", size=12, color="#4a4a44"),
    margin=dict(l=16, r=16, t=48, b=16),
    legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=11)),
    xaxis=dict(showgrid=False, linecolor="#e6e6e0"),
    yaxis=dict(gridcolor="#f1f1ec", linecolor="#e6e6e0"),
)

# ── 年度预算配置（手动填写，单位：元）──
ANNUAL_BUDGET = {
    "2022": 0,
    "2023": 8000000,
    "2024": 10000000,
    "2025": 8500000,
    "2026": 5400000,
}

PASSWORDS = {
    "看板1：预算进度": "akdys",
    "看板2：新媒体数据": "akdxmt",
}

# 最近一次成功数据会同时保存在内存和本地文件中。页面重跑不会请求飞书，
# 每个浏览器会话只有在首次通过密码验证时才刷新一次。
DATA_CACHE_PATH = Path(__file__).with_name(".last_successful_data.pkl")

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

@st.cache_data(ttl=5400)
def get_token():
    res = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
        timeout=20,
    )
    payload = parse_feishu_response(res, "获取访问凭证")
    token = payload.get("tenant_access_token")
    if not token:
        raise FeishuAPIError("飞书没有返回 tenant_access_token")
    return token

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

def read_sheet():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}"
        f"/values/{SHEET_ID}!A1:AQ2000"
    )
    response = requests.get(
        url,
        headers=headers,
        params={"valueRenderOption": "FormattedValue"},
        timeout=30,
    )
    payload = parse_feishu_response(response, "读取电子表格")
    values = payload.get("data", {}).get("valueRange", {}).get("values", [])
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

def _to_excel_bytes(df, sheet_name="数据"):
    """导出为格式化 Excel：表头样式、自适应列宽、数字千分位。"""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#f3f4f6", "border": 1,
            "border_color": "#e5e7eb", "align": "center", "valign": "vcenter",
        })
        num_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1, "border_color": "#f3f4f6"})
        int_fmt = wb.add_format({"num_format": "#,##0", "border": 1, "border_color": "#f3f4f6"})
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, header_fmt)
            # 列宽：中文按 2 计，据表头与内容估算
            max_len = len(str(col_name)) * 2
            for v in df[col_name].head(200):
                max_len = max(max_len, len(str(v)))
            ws.set_column(col_idx, col_idx, min(max(max_len + 2, 9), 28))
            # 数字列加千分位格式
            if pd.api.types.is_integer_dtype(df[col_name]):
                ws.set_column(col_idx, col_idx, min(max(max_len + 2, 9), 28), int_fmt)
            elif pd.api.types.is_float_dtype(df[col_name]):
                ws.set_column(col_idx, col_idx, min(max(max_len + 2, 9), 28), num_fmt)
    buf.seek(0)
    return buf

def safe_agg(df, group_col, agg_dict):
    g = df.groupby(group_col).agg(**agg_dict).reset_index()
    if "总花费" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["客资成本"] = 0.0
        g.loc[mask, "客资成本"] = (g.loc[mask, "总花费"] / g.loc[mask, "客资总数"]).round(2)
    if "总花费" in g.columns and "销收成交总量" in g.columns:
        mask = g["销收成交总量"] > 0
        g["成交成本"] = 0.0
        g.loc[mask, "成交成本"] = (g.loc[mask, "总花费"] / g.loc[mask, "销收成交总量"]).round(2)
    if "到店总量" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["总到店率%"] = 0.0
        g.loc[mask, "总到店率%"] = (g.loc[mask, "到店总量"] / g.loc[mask, "客资总数"] * 100).round(2)
    if "销收成交总量" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["总成交率%"] = 0.0
        g.loc[mask, "总成交率%"] = (g.loc[mask, "销收成交总量"] / g.loc[mask, "客资总数"] * 100).round(2)
    if "销售到店" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["销售线索到店率%"] = 0.0
        g.loc[mask, "销售线索到店率%"] = (g.loc[mask, "销售到店"] / g.loc[mask, "客资总数"] * 100).round(2)
    if "收购到店" in g.columns and "客资总数" in g.columns:
        mask = g["客资总数"] > 0
        g["收购线索到店率%"] = 0.0
        g.loc[mask, "收购线索到店率%"] = (g.loc[mask, "收购到店"] / g.loc[mask, "客资总数"] * 100).round(2)
    if "销售成交" in g.columns and "销售到店" in g.columns:
        mask = g["销售到店"] > 0
        g["销售到店成交率%"] = 0.0
        g.loc[mask, "销售到店成交率%"] = (g.loc[mask, "销售成交"] / g.loc[mask, "销售到店"] * 100).round(2)
    if "收购成交" in g.columns and "收购到店" in g.columns:
        mask = g["收购到店"] > 0
        g["收购到店成交率%"] = 0.0
        g.loc[mask, "收购到店成交率%"] = (g.loc[mask, "收购成交"] / g.loc[mask, "收购到店"] * 100).round(2)
    if "销售成交" in g.columns and "销售客资" in g.columns:
        mask = g["销售客资"] > 0
        g["销售线索成交率%"] = 0.0
        g.loc[mask, "销售线索成交率%"] = (g.loc[mask, "销售成交"] / g.loc[mask, "销售客资"] * 100).round(2)
    if "收购成交" in g.columns and "收购客资" in g.columns:
        mask = g["收购客资"] > 0
        g["收购线索成交率%"] = 0.0
        g.loc[mask, "收购线索成交率%"] = (g.loc[mask, "收购成交"] / g.loc[mask, "收购客资"] * 100).round(2)
    return g

def clean_df():
    df = read_sheet()
    if df.empty:
        return pd.DataFrame()
    if "年份" in df.columns:
        df["年份"] = df["年份"].astype(str).str.strip().str.replace(".0", "", regex=False)
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip()
        df["月份"] = df["月份"].apply(lambda x: x + "月" if x.isdigit() else x)
    skip_cols = ["城市", "年份", "月份"]
    for col in df.columns:
        if col not in skip_cols:
            df[col] = to_num(df[col])
    if "城市" in df.columns:
        df = df[df["城市"].isin(CITIES)].copy()
        df["城市"] = pd.Categorical(df["城市"], categories=CITIES, ordered=True)
        df = df.sort_values("城市")
    if "年份" in df.columns:
        df = df[df["年份"].isin(YEARS)].copy()
    return df


@st.cache_resource
def get_shared_data_cache():
    dataframe, updated_at = load_snapshot(DATA_CACHE_PATH)
    return {
        "dataframe": dataframe,
        "updated_at": updated_at,
        "lock": threading.Lock(),
    }


def load_data_once_per_session():
    if "dashboard_data" in st.session_state:
        return (
            st.session_state.dashboard_data,
            st.session_state.dashboard_data_source,
            st.session_state.dashboard_data_updated_at,
            st.session_state.get("dashboard_data_error"),
        )

    shared_cache = get_shared_data_cache()
    with shared_cache["lock"]:
        result = load_with_fallback(
            clean_df,
            DATA_CACHE_PATH,
            shared_cache["dataframe"],
            shared_cache["updated_at"],
        )
        if result.source == "live":
            shared_cache["dataframe"] = result.dataframe.copy()
            shared_cache["updated_at"] = result.updated_at

    st.session_state.dashboard_data = result.dataframe
    st.session_state.dashboard_data_source = result.source
    st.session_state.dashboard_data_updated_at = result.updated_at
    st.session_state.dashboard_data_error = str(result.error) if result.error else None
    return (
        result.dataframe,
        result.source,
        result.updated_at,
        st.session_state.dashboard_data_error,
    )

def apply_filter(df, cities, years, months):
    d = df.copy()
    if cities:
        d = d[d["城市"].isin(cities)]
    if years:
        d = d[d["年份"].isin(years)]
    if months and "月份" in d.columns:
        d = d[d["月份"].isin(months)]
    return d

def metric_html(label, value, sub="", en=""):
    """设计风格：白卡 + 小标签（中/英文小注）+ 衬线大数字 + lime 进度点。"""
    en_html = f'<span style="font-size:10px;letter-spacing:.14em;color:#9a9a92;font-weight:600;text-transform:uppercase;">{en}</span>' if en else ""
    sub_html = (
        f'<div style="margin-top:10px;"><span style="background:#f4f9c8;color:#5c6b0e;'
        f'font-size:11px;font-weight:600;border-radius:999px;padding:3px 10px;">{sub}</span></div>'
        if sub else ""
    )
    return f"""<div style="background:#ffffff;border:1px solid #e6e6e0;border-radius:16px;
        padding:18px 22px;margin-bottom:10px;">
        <div style="display:flex;justify-content:space-between;align-items:center;">
            <div style="color:#1c1c1a;font-size:13px;font-weight:600;">{label}</div>{en_html}
        </div>
        <div style="color:#1c1c1a;font-family:Georgia,'Noto Serif SC',serif;
            font-size:30px;font-weight:700;line-height:1.25;margin-top:6px;">{value}</div>{sub_html}
    </div>"""

# ── 侧边栏：看板选择 + 密码 ──
with st.sidebar:
    st.markdown("## 选择看板")
    selected_board = st.selectbox("", list(PASSWORDS.keys()), label_visibility="collapsed")
    pwd_input = st.text_input("请输入访问密码", type="password")
    is_authenticated = pwd_input == PASSWORDS[selected_board]

    if pwd_input and not is_authenticated:
        st.error("密码错误")

    st.divider()

# ── 未登录提示 ──
if not is_authenticated:
    st.markdown("""
    <div style="text-align:center;padding:80px 0;">
        <div style="font-size:48px;margin-bottom:16px;">🔒</div>
        <div style="font-size:20px;font-weight:600;color:#111827;margin-bottom:8px;">请在左侧选择看板并输入密码</div>
        <div style="font-size:14px;color:#6b7280;">不同看板需要对应的访问密码</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── 每个已登录浏览器会话只加载一次数据 ──
with st.spinner("正在加载数据..."):
    try:
        df, data_source, data_updated_at, data_error = load_data_once_per_session()
    except MonthlyQuotaExceeded:
        st.error("飞书 API 月度额度已耗尽，且暂时没有可用的历史数据缓存。")
        st.stop()
    except Exception as e:
        st.error(f"数据加载失败：{e}")
        st.stop()

if data_source == "cache":
    updated_text = data_updated_at or "未知时间"
    st.warning(
        f"飞书暂时无法更新，当前展示最后一次成功数据（更新时间：{updated_text}）。"
    )

display_updated_at = (
    data_updated_at.replace("T", " ")[:19] if data_updated_at else "未知时间"
)


def latest_data_month(frame):
    """取当前数据中最新的「年-月」作为覆盖月份上界（如 2026-09）。"""
    if frame is None or frame.empty or "年份" not in frame.columns or "月份" not in frame.columns:
        return None
    try:
        ym = frame.dropna(subset=["年份", "月份"]).copy()
        ym["_y"] = ym["年份"].astype(str).str.replace("年", "", regex=False).str.strip()
        ym["_m"] = ym["月份"].astype(str).str.replace("月", "", regex=False).str.strip()
        ym = ym[ym["_y"].str.fullmatch(r"\d{4}") & ym["_m"].str.fullmatch(r"\d{1,2}")]
        if ym.empty:
            return None
        key = ym["_y"] + "-" + ym["_m"].str.zfill(2)
        return str(key.max())
    except Exception:
        return None


def freshness_badge(source, updated_text, covered_text=None):
    """顶部数据新鲜度徽章：LIVE 风格胶囊，区分实时 / 缓存，标注抓取时间与覆盖月份。"""
    if source == "live":
        dot, pill_bg, pill_fg, label = "#a4c400", "#f4f9c8", "#5c6b0e", "LIVE 实时"
    else:
        dot, pill_bg, pill_fg, label = "#d97706", "#fdf0d5", "#92400e", "CACHE 缓存"
    covered = (
        f'<span style="color:#c9c9c2;">｜</span>'
        f'<span style="color:#6f6f68;">数据覆盖月份：<b style="color:#1c1c1a;">截止 {covered_text}</b></span>'
        if covered_text else ""
    )
    return f"""<div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;
        background:#ffffff;border:1px solid #e6e6e0;border-radius:999px;
        padding:8px 18px;margin-bottom:16px;font-size:13px;">
        <span style="display:inline-flex;align-items:center;gap:7px;background:{pill_bg};
            color:{pill_fg};font-weight:700;font-size:11px;letter-spacing:.08em;
            border-radius:999px;padding:4px 12px;">
            <span style="width:7px;height:7px;border-radius:50%;background:{dot};display:inline-block;"></span>
            {label}
        </span>
        <span style="color:#6f6f68;">数据抓取时间：<b style="color:#1c1c1a;">{updated_text}</b></span>
        {covered}
    </div>"""


# ── 数据加载完成后再渲染筛选器 ──
with st.sidebar:
    if selected_board == "看板2：新媒体数据":
        st.markdown("## 筛选条件")

        # 快捷时间范围：一键设置年份 / 月份筛选
        st.caption("快捷范围")
        quick = st.radio(
            "快捷范围", ["全部", "本年度", "去年", "最近3月", "最近6月"],
            horizontal=True, label_visibility="collapsed", key="quick_range",
        )
        _cur_year = str(datetime.datetime.now().year)
        _cur_month_idx = datetime.datetime.now().month
        if quick == "本年度":
            default_years, default_months = [_cur_year], []
        elif quick == "去年":
            default_years, default_months = [str(int(_cur_year) - 1)], []
        elif quick == "最近3月":
            default_years = sorted({_cur_year} | {str(int(_cur_year) - 1)})
            _ms = [( _cur_month_idx - i - 1) % 12 + 1 for i in range(3)]
            default_months = [MONTHS[m - 1] for m in sorted(set(_ms))]
        elif quick == "最近6月":
            default_years = sorted({_cur_year} | {str(int(_cur_year) - 1)})
            _ms = [(_cur_month_idx - i - 1) % 12 + 1 for i in range(6)]
            default_months = [MONTHS[m - 1] for m in sorted(set(_ms))]
        else:
            default_years, default_months = YEARS, []
        _years_opt = [y for y in YEARS if y in default_years] or YEARS
        _months_opt = [m for m in MONTHS if m in default_months]

        sel_cities = st.multiselect("城市（可多选）", CITIES, default=CITIES)
        sel_years = st.multiselect("年份（可多选）", YEARS, default=_years_opt or YEARS)
        sel_months = st.multiselect("月份（可多选）", MONTHS, default=_months_opt)
        st.divider()
        df_filtered = apply_filter(df, sel_cities, sel_years, sel_months)
        st.caption(f"当前数据：{len(df_filtered)} 条")
    else:
        st.markdown("## 筛选条件")
        sel_year_budget = st.selectbox("年份", ["2026"], index=0)

# ══════════════════════════════════════
# 看板1：预算进度
# ══════════════════════════════════════
if selected_board == "看板1：预算进度":
    year = sel_year_budget
    budget = ANNUAL_BUDGET.get(year, 0)

    # 从数据库读取该年度各月花费
    df_year = df[df["年份"] == year].copy()
    month_spend = df_year.groupby("月份")["总花费"].sum()
    month_spend_ordered = {}
    for m in MONTHS:
        month_spend_ordered[m] = month_spend.get(m, 0)

    total_spent = sum(month_spend_ordered.values())

    # 自动获取截至月份（当前月份）
    now = datetime.datetime.now()
    current_month_idx = now.month  # 1-12
    current_month_str = MONTHS[current_month_idx - 1]

    # 已过月份数（截至上个月）
    passed_months = current_month_idx - 1
    remaining_months = 12 - passed_months
    time_progress = passed_months / 12 * 100

    budget_progress = (total_spent / budget * 100) if budget > 0 else 0
    remaining = budget - total_spent
    monthly_available = remaining / remaining_months if remaining_months > 0 else 0

    def wan(v):
        return f"{v/10000:.1f}万"

    st.markdown(f"""
    <div style="padding:4px 0 20px 0;margin-bottom:18px;">
        <div style="display:flex;align-items:center;gap:12px;">
            <span style="width:22px;height:22px;border-radius:7px;background:#dcef4e;display:inline-block;"></span>
            <h2 style="margin:0;color:#1c1c1a;font-weight:700;font-size:26px;letter-spacing:-.01em;">{year}年度 常规预算执行进度</h2>
            <span style="font-size:11px;letter-spacing:.14em;color:#9a9a92;font-weight:600;">BUDGET · PROGRESS</span>
        </div>
        <p style="margin:8px 0 0 34px;color:#6f6f68;font-size:13px;">
            统计区间：{year}年 1月 — 12月 ｜ 单位：万元 ｜ 数据截至：{MONTHS[passed_months-1] if passed_months > 0 else "暂无"}
        </p>
    </div>
    """, unsafe_allow_html=True)

    # 核心指标
    c1,c2,c3,c4 = st.columns(4)
    c1.markdown(metric_html("年度总预算", wan(budget), f"{year}年全年"), unsafe_allow_html=True)
    c2.markdown(metric_html(f"已花费（1—{MONTHS[passed_months-1] if passed_months>0 else '—'}）",
                            wan(total_spent), f"占预算 {budget_progress:.1f}%"), unsafe_allow_html=True)
    c3.markdown(metric_html("剩余预算", wan(remaining), f"占预算 {100-budget_progress:.1f}%"), unsafe_allow_html=True)
    c4.markdown(metric_html(f"月均可用（{MONTHS[passed_months]}—12月）",
                            wan(monthly_available), f"剩余{remaining_months}个月均摊"), unsafe_allow_html=True)

    st.divider()

    # 进度条卡片
    diff = budget_progress - time_progress
    if diff > 0:
        status_text = f"预算消耗高于时间进度 {abs(diff):.1f}%，注意控制节奏"
        status_color = "#ef4444"
    elif diff < 0:
        status_text = f"预算消耗低于时间进度 {abs(diff):.1f}%，执行节奏良好"
        status_color = "#10b981"
    else:
        status_text = "预算消耗与时间进度一致"
        status_color = "#6b7280"

    st.markdown(f"""
    <div style="background:white;border:1px solid #eef0f4;border-radius:12px;padding:24px;margin-bottom:16px;">
        <div style="font-weight:600;font-size:16px;color:#111827;margin-bottom:20px;">预算执行进度对比</div>
        <div style="margin-bottom:16px;">
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <span style="font-size:14px;color:#374151;">预算消耗</span>
                <span style="font-size:14px;font-weight:600;color:#374151;">{wan(total_spent)} / {wan(budget)} · {budget_progress:.1f}%</span>
            </div>
            <div style="background:#e5e7eb;border-radius:999px;height:12px;">
                <div style="background:#10b981;width:{min(budget_progress,100):.1f}%;height:12px;border-radius:999px;"></div>
            </div>
        </div>
        <div style="margin-bottom:20px;">
            <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                <span style="font-size:14px;color:#374151;">时间进度</span>
                <span style="font-size:14px;font-weight:600;color:#374151;">{passed_months}个月 / 12个月 · {time_progress:.1f}%</span>
            </div>
            <div style="background:#e5e7eb;border-radius:999px;height:12px;">
                <div style="background:#f59e0b;width:{min(time_progress,100):.1f}%;height:12px;border-radius:999px;"></div>
            </div>
        </div>
        <div style="font-size:13px;color:#6b7280;">
            🟢 预算消耗 {budget_progress:.1f}%　　🟠 时间进度 {time_progress:.1f}%　　
            <span style="color:{status_color};font-weight:500;">· {status_text}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 各月花费分布
    st.markdown("""
    <div style="background:white;border:1px solid #eef0f4;border-radius:12px;padding:24px;margin-bottom:16px;">
        <div style="font-weight:600;font-size:16px;color:#111827;margin-bottom:20px;">各月花费分布</div>
    """, unsafe_allow_html=True)

    avg_spent = total_spent / passed_months if passed_months > 0 else 0

    # 响应式月度网格：宽屏 6 列 × 2 行，窄屏自动收窄，避免 12 列挤压。
    _month_items = list(month_spend_ordered.items())
    for _r in range(2):
        _rc = st.columns(6)
        for _c in range(6):
            _idx = _r * 6 + _c
            if _idx >= len(_month_items):
                continue
            m, v = _month_items[_idx]
            month_idx = _idx + 1
            with _rc[_c]:
                if month_idx < current_month_idx:
                    st.markdown(f"""
                    <div style="background:#dcfce7;border:1px solid #86efac;border-radius:8px;
                        padding:10px 4px;text-align:center;min-height:76px;margin-bottom:6px;">
                        <div style="font-size:13px;font-weight:600;color:#166534;">{m}</div>
                        <div style="font-size:12px;color:#15803d;margin-top:4px;">{v/10000:.1f}万</div>
                    </div>
                    """, unsafe_allow_html=True)
                elif month_idx == current_month_idx:
                    st.markdown(f"""
                    <div style="background:white;border:2px solid #f59e0b;border-radius:8px;
                        padding:10px 4px;text-align:center;min-height:76px;margin-bottom:6px;">
                        <div style="font-size:13px;font-weight:600;color:#92400e;">{m}</div>
                        <div style="font-size:12px;color:#b45309;margin-top:4px;">—</div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;
                        padding:10px 4px;text-align:center;min-height:76px;margin-bottom:6px;">
                        <div style="font-size:13px;color:#9ca3af;">{m}</div>
                        <div style="font-size:12px;color:#d1d5db;margin-top:4px;">—</div>
                    </div>
                    """, unsafe_allow_html=True)

    st.markdown(f"""
    </div>
    <div style="font-size:13px;color:#6b7280;margin-top:8px;">
        🟢 已完成（1—{MONTHS[passed_months-1] if passed_months>0 else '—'}）　
        🟡 当前月（{current_month_str}）　
        ⬜ 待执行（{MONTHS[passed_months+1] if passed_months+1 < 12 else ''}—12月）<br>
        * 月均花费（1—{MONTHS[passed_months-1] if passed_months>0 else '—'}）：{wan(avg_spent)}　｜　
        月均可用（{MONTHS[passed_months]}—12月）：{wan(monthly_available)}
    </div>
    """, unsafe_allow_html=True)

    # 月度花费柱状图
    st.divider()
    chart_data = pd.DataFrame({
        "月份": list(month_spend_ordered.keys())[:passed_months],
        "花费": [v for v in list(month_spend_ordered.values())[:passed_months]]
    })
    if not chart_data.empty:
        chart_data["月份"] = pd.Categorical(chart_data["月份"], categories=MONTHS, ordered=True)
        fig = px.bar(chart_data, x="月份", y="花费", title=f"{year}年各月花费趋势",
                     color_discrete_sequence=["#10b981"])
        fig.add_hline(y=monthly_available, line_dash="dash", line_color="#f59e0b",
                      annotation_text=f"月均可用 {wan(monthly_available)}")
        st.plotly_chart(make_chart(fig), use_container_width=True)

# ══════════════════════════════════════
# 看板2：新媒体数据
# ══════════════════════════════════════
else:
    st.markdown(f"""
    <div style="padding:4px 0 16px 0;margin-bottom:14px;">
        <div style="display:flex;align-items:center;gap:12px;">
            <span style="width:22px;height:22px;border-radius:7px;background:#dcef4e;display:inline-block;"></span>
            <h2 style="margin:0;color:#1c1c1a;font-weight:700;font-size:26px;letter-spacing:-.01em;">新媒体年度数据看板</h2>
            <span style="font-size:11px;letter-spacing:.14em;color:#9a9a92;font-weight:600;">ANALYTICS · OVERVIEW</span>
        </div>
        <p style="margin:8px 0 0 34px;color:#6f6f68;font-size:13px;">
            城市：{'、'.join(sel_cities) if sel_cities else '未选择'} ·
            年份：{'、'.join(sel_years) if sel_years else '未选择'} ·
            {'、'.join(sel_months) if sel_months else '全部月份'} ·
            每次打开看板时更新一次
        </p>
    </div>
    """, unsafe_allow_html=True)

    # 数据新鲜度徽章
    _covered_month = latest_data_month(df_filtered)
    st.markdown(
        freshness_badge(data_source, display_updated_at, _covered_month),
        unsafe_allow_html=True,
    )

    total_spend     = df_filtered["总花费"].sum() if "总花费" in df_filtered.columns else 0
    total_keizi     = df_filtered["客资总数"].sum() if "客资总数" in df_filtered.columns else 0
    total_daodian   = df_filtered["到店总量"].sum() if "到店总量" in df_filtered.columns else 0
    total_chengjiao = df_filtered["销收成交总量"].sum() if "销收成交总量" in df_filtered.columns else 0
    total_xiaoshou  = df_filtered["销售成交"].sum() if "销售成交" in df_filtered.columns else 0
    total_shougou   = df_filtered["收购成交"].sum() if "收购成交" in df_filtered.columns else 0
    keizi_cost      = total_spend / total_keizi if total_keizi > 0 else 0
    chengjiao_cost  = total_spend / total_chengjiao if total_chengjiao > 0 else 0
    daodian_rate    = total_daodian / total_keizi * 100 if total_keizi > 0 else 0
    chengjiao_rate  = total_chengjiao / total_keizi * 100 if total_keizi > 0 else 0

    daodian_cost    = total_spend / total_daodian if total_daodian > 0 else 0

    # ── 指标卡三排布局：花费 / 到店 / 成交 ──
    r1a,r1b,r1c = st.columns(3)
    r1a.markdown(metric_html("总花费", f"¥{total_spend:,.0f}"), unsafe_allow_html=True)
    r1b.markdown(metric_html("总客资量", f"{int(total_keizi):,}"), unsafe_allow_html=True)
    r1c.markdown(metric_html("客资成本", f"¥{keizi_cost:.2f}"), unsafe_allow_html=True)

    r2a,r2b,r2c = st.columns(3)
    r2a.markdown(metric_html("到店总量", f"{int(total_daodian):,}"), unsafe_allow_html=True)
    r2b.markdown(metric_html("到店成本", f"¥{daodian_cost:.2f}"), unsafe_allow_html=True)
    r2c.markdown(metric_html("客资到店率", f"{daodian_rate:.2f}%"), unsafe_allow_html=True)

    r3a,r3b,r3c,r3d = st.columns(4)
    r3a.markdown(metric_html("总成交量", f"{int(total_chengjiao):,}"), unsafe_allow_html=True)
    r3b.markdown(metric_html("销售成交", f"{int(total_xiaoshou):,}"), unsafe_allow_html=True)
    r3c.markdown(metric_html("收购成交", f"{int(total_shougou):,}"), unsafe_allow_html=True)
    r3d.markdown(metric_html("客资成交率", f"{chengjiao_rate:.2f}%"), unsafe_allow_html=True)

    st.divider()

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "分城市", "年度对比", "趋势分析", "分渠道", "数据明细"
    ])

    with tab1:
        st.subheader("分城市经营对比")
        st.markdown("""
        <div style="background:#fffbeb;border:1px solid #f59e0b;border-radius:8px;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#92400e;">
            ⚠️ <b>数据说明：</b><br>
            1. 2022年、2023年原始数据缺失及表格混乱，部分数据可能不准确，仅作参考；<br>
            2. 2024、2025、2026年数据基本正确，如发现差异请反馈媒介调整。
        </div>
        """, unsafe_allow_html=True)
        if not df_filtered.empty and "城市" in df_filtered.columns:
            cg = safe_agg(df_filtered, "城市", {
                "总花费": ("总花费","sum"),
                "客资总数": ("客资总数","sum"),
                "销售客资": ("销售客资","sum"),
                "收购客资": ("收购客资","sum"),
                "到店总量": ("到店总量","sum"),
                "销售到店": ("销售到店","sum"),
                "收购到店": ("收购到店","sum"),
                "销收成交总量": ("销收成交总量","sum"),
                "销售成交": ("销售成交","sum"),
                "收购成交": ("收购成交","sum"),
            })
            cg["城市"] = pd.Categorical(cg["城市"], categories=CITIES, ordered=True)
            cg = cg.sort_values("城市")
            for col in ["总花费", "客资成本", "成交成本"]:
                if col in cg.columns:
                    cg[col] = cg[col].round(2)
            st.dataframe(cg, use_container_width=True, hide_index=True)
            st.download_button(
                label="⬇️ 导出分城市汇总 Excel",
                data=_to_excel_bytes(cg.assign(城市=cg["城市"].astype(str)), sheet_name="分城市"),
                file_name=f"分城市汇总_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_city",
            )
            ca,cb = st.columns(2)
            with ca:
                fig = px.bar(cg,x="城市",y="客资总数",title="各城市客资量",color="城市",
                            color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
                st.plotly_chart(make_chart(fig),use_container_width=True)
            with cb:
                fig = px.bar(cg,x="城市",y="销收成交总量",title="各城市成交量",color="城市",
                            color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
                st.plotly_chart(make_chart(fig),use_container_width=True)
            cc,cd = st.columns(2)
            with cc:
                fig = px.bar(cg,x="城市",y="到店总量",title="各城市到店量",color="城市",
                            color_discrete_sequence=COLORS,category_orders={"城市":CITIES})
                st.plotly_chart(make_chart(fig),use_container_width=True)
            with cd:
                fig = px.bar(cg,x="城市",y="总到店率%",title="各城市总到店率(%)",color="城市",
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
        st.info("💡 年度对比显示全年数据，不受月份筛选影响")
        df_year = apply_filter(df, sel_cities, sel_years, [])
        if not df_year.empty and "年份" in df_year.columns:
            yg = safe_agg(df_year, "年份", {
                "总花费": ("总花费","sum"),
                "客资总数": ("客资总数","sum"),
                "销售客资": ("销售客资","sum"),
                "收购客资": ("收购客资","sum"),
                "到店总量": ("到店总量","sum"),
                "销售到店": ("销售到店","sum"),
                "收购到店": ("收购到店","sum"),
                "销收成交总量": ("销收成交总量","sum"),
                "销售成交": ("销售成交","sum"),
                "收购成交": ("收购成交","sum"),
            })
            yg["年份"] = pd.Categorical(yg["年份"], categories=YEARS, ordered=True)
            yg = yg.sort_values("年份")
            for col in ["总花费", "客资成本", "成交成本"]:
                if col in yg.columns:
                    yg[col] = yg[col].round(2)
            st.dataframe(yg, use_container_width=True, hide_index=True)
            st.download_button(
                label="⬇️ 导出年度对比 Excel",
                data=_to_excel_bytes(yg.assign(年份=yg["年份"].astype(str)), sheet_name="年度对比"),
                file_name=f"年度对比_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_year",
            )
            ya,yb = st.columns(2)
            with ya:
                fig = px.bar(yg,x="年份",y="客资总数",title="各年度客资量",color="年份",color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)
            with yb:
                fig = px.bar(yg,x="年份",y="销收成交总量",title="各年度成交量",color="年份",color_discrete_sequence=COLORS)
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
                fig = px.line(yg,x="年份",y="总到店率%",title="各年度到店率趋势",markers=True,color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)
            fig = px.line(yg,x="年份",y="总花费",title="各年度总花费趋势",
                          markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)

    with tab3:
        st.subheader("月度趋势分析")
        st.info("💡 趋势分析显示全年月度走势，不受月份筛选影响")
        df_trend = apply_filter(df, sel_cities, sel_years, [])
        if not df_trend.empty and "月份" in df_trend.columns:
            tm = safe_agg(df_trend, ["年份","月份"], {
                "客资总数": ("客资总数","sum"),
                "到店总量": ("到店总量","sum"),
                "销收成交总量": ("销收成交总量","sum"),
                "总花费": ("总花费","sum"),
                "销售到店": ("销售到店","sum"),
                "收购到店": ("收购到店","sum"),
                "销售成交": ("销售成交","sum"),
                "收购成交": ("收购成交","sum"),
                "销售客资": ("销售客资","sum"),
                "收购客资": ("收购客资","sum"),
            })
            tm["月份"] = pd.Categorical(tm["月份"], categories=MONTHS, ordered=True)
            tm = tm.sort_values(["年份","月份"])
            fig1 = px.line(tm,x="月份",y="总花费",color="年份",title="各年度花费月度趋势",
                           markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig1),use_container_width=True)
            fig2 = px.line(tm,x="月份",y="客资总数",color="年份",title="各年度客资量月度趋势",
                           markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig2),use_container_width=True)
            fig3 = px.line(tm,x="月份",y="客资成本",color="年份",title="各年度客资成本月度趋势",
                           markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig3),use_container_width=True)
            fig4 = px.line(tm,x="月份",y="总到店率%",color="年份",title="各年度到店率月度趋势",
                           markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig4),use_container_width=True)
            fig5 = px.line(tm,x="月份",y="销收成交总量",color="年份",title="各年度成交量月度趋势",
                           markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig5),use_container_width=True)

    with tab4:
        st.subheader("分渠道综合对比")
        channel_map = {
            "抖音号": {"花费": "抖音号花费", "客资": "抖音号客资", "成交": "抖音号成交", "到店": "抖音号到店"},
            "信息流": {"花费": "信息流花费", "客资": "信息流客资", "成交": "信息流成交", "到店": "信息流到店"},
            "微信":   {"花费": "微信花费",   "客资": "微信客资",   "成交": "微信成交",   "到店": "微信到店"},
            "小红书": {"花费": "小红书花费", "客资": "小红书客资", "成交": "小红书成交", "到店": "小红书到店"},
            "B站":    {"花费": "B站花费",    "客资": "b站客资",    "成交": "b站成交",    "到店": "b站到店"},
            "快手":   {"花费": "快手花费",   "客资": "快手客资",   "成交": "快手成交",   "到店": "快手到店"},
            "之家":   {"花费": "之家花费",   "客资": "之家客资",   "成交": "之家成交",   "到店": "之家到店"},
        }

        def find_col(df_, name):
            # 大小写不敏感地匹配列名（源表中 B站 等列大小写历史上不统一）
            lowered = {c.lower(): c for c in df_.columns}
            return lowered.get(name.lower())
        rows = []
        for ch, cols in channel_map.items():
            spend_col = find_col(df_filtered, cols["花费"])
            keizi_col = find_col(df_filtered, cols["客资"])
            deal_col = find_col(df_filtered, cols["成交"])
            daodian_col = find_col(df_filtered, cols.get("到店", ""))
            花费 = float(df_filtered[spend_col].sum()) if spend_col else 0
            客资 = float(df_filtered[keizi_col].sum()) if keizi_col else 0
            成交 = float(df_filtered[deal_col].sum()) if deal_col else 0
            到店 = float(df_filtered[daodian_col].sum()) if daodian_col else 0
            客资成本 = round(花费 / 客资, 2) if 客资 > 0 else 0
            成交成本 = round(花费 / 成交, 2) if 成交 > 0 else 0
            成交率 = round(成交 / 客资 * 100, 2) if 客资 > 0 else 0
            if 花费 > 0 or 客资 > 0:
                rows.append({
                    "渠道": ch, "花费": round(花费, 2), "客资": int(客资),
                    "到店量": int(到店), "客资成本": 客资成本, "成交": int(成交),
                    "成交成本": 成交成本, "成交率%": 成交率,
                })
        if rows:
            ch_df = pd.DataFrame(rows).sort_values("花费", ascending=False)
            st.dataframe(ch_df, use_container_width=True, hide_index=True)
            ca,cb = st.columns(2)
            with ca:
                fig = px.bar(ch_df,x="渠道",y="花费",title="各渠道花费",
                             color="渠道",color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)
            with cb:
                fig = px.bar(ch_df,x="渠道",y="客资",title="各渠道客资量",
                             color="渠道",color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)
            cc,cd = st.columns(2)
            with cc:
                fig = px.bar(ch_df,x="渠道",y="客资成本",title="各渠道客资成本",
                             color="渠道",color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)
            with cd:
                fig = px.bar(ch_df,x="渠道",y="成交成本",title="各渠道成交成本",
                             color="渠道",color_discrete_sequence=COLORS)
                st.plotly_chart(make_chart(fig),use_container_width=True)

            # ── 渠道多指标雷达图（各指标归一化到 0-100，便于同轴对比）──
            # 归一化口径：以所有渠道中的最大值为 100，避免量纲差异（花费上万元 vs 成交个位数）。
            radar_metrics = ["花费", "客资", "到店量", "成交量"]
            radar_available = [m for m in radar_metrics if m in ch_df.columns]
            if len(radar_available) >= 3:
                radar_df = ch_df[["渠道"] + radar_available].copy()
                for m in radar_available:
                    vmax = radar_df[m].max()
                    radar_df[m] = (radar_df[m] / vmax * 100).round(1) if vmax > 0 else 0
                rd, re_ = st.columns(2)
                with rd:
                    fig = px.line_polar(
                        radar_df.melt(id_vars="渠道", var_name="指标", value_name="归一值"),
                        r="归一值", theta="指标", color="渠道", line_close=True,
                        title="各渠道多指标雷达对比（归一化 0–100）",
                    )
                    fig.update_traces(fill="toself", opacity=0.35)
                    st.plotly_chart(make_chart(fig), use_container_width=True)
                with re_:
                    # 同时给出原始值热力表，避免归一化丢失绝对值
                    show_df = ch_df[["渠道"] + radar_available].copy()
                    st.caption("雷达图原始数值（未归一化）")
                    st.dataframe(show_df, use_container_width=True, hide_index=True)

        # ── 分渠道 × 分地区 × 分月明细 ──
        # 说明：受「城市 / 年份 / 月份」筛选控制，按所选范围展示各渠道分月数据；
        # 比率类字段在分母为 0（无客资 / 无成交）时显示「—」，以区分“真实为 0”与“无法计算”。
        st.divider()
        st.subheader("分渠道 × 分地区 × 分月明细")
        if sel_months:
            st.info(f"💡 本表按当前筛选显示（城市 / 年份 / 月份：{'、'.join(sel_months)}）")
        else:
            st.info("💡 本表按当前筛选显示；未选月份时展示各渠道分月完整走势")

        df_ch_month = apply_filter(df, sel_cities, sel_years, sel_months)
        ch_rows = []
        if not df_ch_month.empty and "城市" in df_ch_month.columns and "月份" in df_ch_month.columns:
            for ch, cols in channel_map.items():
                spend_col = find_col(df_ch_month, cols["花费"])
                keizi_col = find_col(df_ch_month, cols["客资"])
                deal_col = find_col(df_ch_month, cols["成交"])
                daodian_col = find_col(df_ch_month, cols["到店"])
                if not spend_col or not keizi_col:
                    continue
                for city in sel_cities:
                    city_df = df_ch_month[df_ch_month["城市"] == city]
                    if city_df.empty:
                        continue
                    for m in MONTHS:
                        month_df = city_df[city_df["月份"] == m]
                        if month_df.empty:
                            continue
                        花费 = float(month_df[spend_col].sum()) if spend_col else 0
                        客资 = float(month_df[keizi_col].sum()) if keizi_col else 0
                        成交 = float(month_df[deal_col].sum()) if deal_col else 0
                        到店 = float(month_df[daodian_col].sum()) if daodian_col else 0
                        # 四项全为 0 的空行跳过（保留任一非零值的行）
                        if 花费 == 0 and 客资 == 0 and 成交 == 0 and 到店 == 0:
                            continue
                        ch_rows.append({
                            "地区": city,
                            "渠道": ch,
                            "月份": m,
                            "花费": round(花费, 2),
                            "客资量": int(客资),
                            "客资成本": round(花费 / 客资, 2) if 客资 > 0 else None,
                            "到店量": int(到店),
                            "到店率%": round(到店 / 客资 * 100, 2) if 客资 > 0 else None,
                            "成交量": int(成交),
                            "客资成交率%": round(成交 / 客资 * 100, 2) if 客资 > 0 else None,
                            "成交成本": round(花费 / 成交, 2) if 成交 > 0 else None,
                        })
        if ch_rows:
            chm_df = pd.DataFrame(ch_rows)
            chm_df["月份"] = pd.Categorical(chm_df["月份"], categories=MONTHS, ordered=True)
            chm_df["地区"] = pd.Categorical(chm_df["地区"], categories=CITIES, ordered=True)
            chm_df = chm_df.sort_values(["地区", "渠道", "月份"]).reset_index(drop=True)
            # 比率列整列转为格式化字符串：空值显示「—」，避免混合类型导致渲染/导出失败
            for _c in ["客资成本", "成交成本"]:
                chm_df[_c] = chm_df[_c].map(lambda v: f"{v:,.2f}" if pd.notna(v) else "—")
            for _c in ["到店率%", "客资成交率%"]:
                chm_df[_c] = chm_df[_c].map(lambda v: f"{v:.2f}%" if pd.notna(v) else "—")
            st.dataframe(chm_df, use_container_width=True, hide_index=True)
            st.download_button(
                label="⬇️ 导出分渠道分月明细 Excel",
                data=_to_excel_bytes(chm_df),
                file_name=f"分渠道分月明细_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("当前筛选条件下暂无分渠道分月数据（请检查城市 / 年份筛选，或源表是否包含渠道列）。")

    with tab5:
        st.subheader("数据明细")
        st.caption(f"当前筛选范围共 {len(df_filtered)} 条记录")
        _detail = df_filtered.dropna(axis=1, how='all')
        st.dataframe(_detail, use_container_width=True)
        st.download_button(
            label="⬇️ 导出当前视图 Excel",
            data=_to_excel_bytes(_detail, sheet_name="数据明细"),
            file_name=f"新媒体数据_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
