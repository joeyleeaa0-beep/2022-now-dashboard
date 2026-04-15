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

def extract_text(val):
    if val is None:
        return ""
    if isinstance(val, list):
        return "".join(item.get("text", "") if isinstance(item, dict) else str(item) for item in val)
    return str(val).strip()

token = get_token()
headers = {"Authorization": f"Bearer {token}"}
url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AQ2?renderType=FORMATTED_VALUE"
res = requests.get(url, headers=headers).json()
values = res.get("data", {}).get("valueRange", {}).get("values", [])

st.subheader("实际读到的所有列名")
for i, h in enumerate(values[0]):
    name = extract_text(h) if isinstance(h, list) else str(h)
    st.write(f"列{i} ({chr(65+i) if i<26 else 'A'+chr(65+i-26)}): {name}")

st.subheader("第2行数据（AJ列以后）")
if len(values) > 1:
    for i, v in enumerate(values[1]):
        if i >= 35:
            name = extract_text(values[0][i]) if isinstance(values[0][i], list) else str(values[0][i])
            st.write(f"列{i} {name}: {v}")
