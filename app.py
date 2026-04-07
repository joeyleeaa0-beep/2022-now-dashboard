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

@st.cache_data(ttl=60)
def read_raw():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AG10?renderType=FORMATTED_VALUE"
    res = requests.get(url, headers=headers).json()
    return res.get("data", {}).get("valueRange", {}).get("values", [])

values = read_raw()

st.subheader("第1行（表头原始数据）")
st.write(values[0])

st.subheader("第2行（第一条数据原始值）")
st.write(values[1])

st.subheader("第3行（第二条数据原始值）")
st.write(values[2])
