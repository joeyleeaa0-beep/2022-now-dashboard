import streamlit as st
import pandas as pd
import requests
import plotly.express as px
from io import BytesIO
import datetime
import os

st.set_page_config(page_title="新媒体年度数据看板", page_icon="📊", layout="wide")

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

@st.cache_data(ttl=60)
def read_sheet():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AG2000?renderType=FORMATTED_VALUE"
    res = requests.get(url, headers=headers).json()
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers_row = [str(h).strip() if h else f"列{i}" for i, h in enumerate(values[0])]
    df = pd.DataFrame(values[1:], columns=headers_row)
    return df

df = read_sheet()

st.subheader("🔍 调试：实际读到的列名")
st.write(list(df.columns))

st.subheader("数据预览（前5行）")
st.dataframe(df.head())
