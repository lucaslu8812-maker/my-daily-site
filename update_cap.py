import pandas as pd
import requests

url = "https://www.twse.com.tw/exchangeReport/BWIBBU_d?response=json"

res = requests.get(url)

try:
    data = res.json()
except:
    raise Exception("❌ API 不是 JSON")

if not data.get("data") or not data.get("fields"):
    raise Exception("❌ 沒抓到股本資料")

df = pd.DataFrame(data["data"], columns=data["fields"])

# 找股本欄
cap_col = None
for col in df.columns:
    if "股本" in col or "資本" in col:
        cap_col = col
        break

if cap_col is None:
    raise Exception(f"❌ 找不到股本欄位: {df.columns}")

df["股本"] = (
    df[cap_col]
    .astype(str)
    .str.replace(",", "")
    .replace("", "0")
    .astype(float)
)

df["發行股數"] = df["股本"] * 10_000_000

df = df[["證券代號", "發行股數"]]

df.to_csv("cap.csv", index=False, encoding="utf-8-sig")

print("✅ cap.csv 產生成功")
