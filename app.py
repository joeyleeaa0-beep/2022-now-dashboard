import streamlit as st
import pandas as pd
import requests

APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = "Jf7osfcF7h2pzFt945XcUkCwned"
SHEET_ID = "0piesb"

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
    if s.startswith("SUM(") or s.startswith("IFERROR(") or s.startswith("="):
        return 0
    return s

@st.cache_data(ttl=60)
def read_raw():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AG2000?renderType=FORMATTED_VALUE"
    res = requests.get(url, headers=headers).json()
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values:
        return pd.DataFrame()
    headers_row = [extract_text(h) if isinstance(h, list) else (str(h).strip() if h else f"列{i}") for i, h in enumerate(values[0])]
    rows = []
    for row in values[1:]:
        new_row = [extract_value(cell) for cell in row]
        while len(new_row) < len(headers_row):
            new_row.append("")
        rows.append(new_row[:len(headers_row)])
    return pd.DataFrame(rows, columns=headers_row)

df = read_raw()

st.subheader("上海2023年2月的数据（一行）")
row = df[(df["城市"]=="上海") & (df["年份"].astype(str).str.contains("2023"))]
if not row.empty:
    st.write(row.iloc[0].to_dict())
else:
    st.write("没找到上海2023年数据")
