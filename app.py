import streamlit as st
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

token = get_token()
headers = {"Authorization": f"Bearer {token}"}

# 只取前3行测试
url = (
    f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}"
    f"/values/{SHEET_ID}!A1:AG3"
    f"?valueRenderOption=FormattedValue&dateTimeRenderOption=FormattedString"
)
res = requests.get(url, headers=headers)

st.subheader("HTTP状态码")
st.write(res.status_code)

st.subheader("原始返回内容（前2000字）")
st.text(res.text[:2000])
