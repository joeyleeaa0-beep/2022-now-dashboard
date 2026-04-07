import streamlit as st
import pandas as pd
import requests

APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = "Jf7osfcF7h2pzFt945XcUkCwned"
SHEET_ID = "0piesb"
CITIES = ["深圳", "上海", "成都", "天津"]
YEARS = ["2022", "2023", "2024", "2025", "2026"]

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
    if isinstance(val, list):
        return extract_text(val)
    if isinstance(val, (int, float)):
        return val
    return str(val).strip()

@st.cache_data(ttl=60)
def read_sheet():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AG2000?renderType=FORMATTED_VALUE"
    res = requests.get(url, headers=headers).json()
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers_row = [extract_text(h) if h else f"列{i}" for i, h in enumerate(values[0])]
    rows = []
    for row in values[1:]:
        new_row = [extract_value(cell) for cell in row]
        while len(new_row) < len(headers_row):
            new_row.append("")
        rows.append(new_row[:len(headers_row)])
    return pd.DataFrame(rows, columns=headers_row)

df = read_sheet()

st.subheader("总行数")
st.write(len(df))

st.subheader("年份列的唯一值")
st.write(df["年份"].unique().tolist() if "年份" in df.columns else "没有年份列")

st.subheader("城市列的唯一值")
st.write(df["城市"].unique().tolist() if "城市" in df.columns else "没有城市列")

st.subheader("总花费列前10个值")
st.write(df["总花费"].head(10).tolist() if "总花费" in df.columns else "没有总花费列")

st.subheader("过滤后（深圳+2023年）的总花费合计")
d = df[df["城市"]=="深圳"]
d = d[d["年份"].astype(str).str.strip().str.replace(".0","",regex=False)=="2023"]
st.write(f"行数：{len(d)}")
st.write(f"总花费合计：{d['总花费'].tolist()[:5] if '总花费' in d.columns else '无'}")
