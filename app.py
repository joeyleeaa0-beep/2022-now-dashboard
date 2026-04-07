import streamlit as st
import requests

APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = "KOu0s7jKqh81tJtEJIgcwcNXnYf"
SHEET_ID = "0piesb"

@st.cache_data(ttl=60)
def get_token():
    res = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return res.json().get("tenant_access_token")

token = get_token()
headers = {"Authorization": f"Bearer {token}"}
url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!J1:J5?renderType=FORMATTED_VALUE"
res = requests.get(url, headers=headers).json()
values = res.get("data", {}).get("valueRange", {}).get("values", [])

st.subheader("J列前5行原始数据")
for i, row in enumerate(values):
    st.write(f"第{i+1}行: {row}")
