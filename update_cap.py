import pandas as pd
import requests
from io import StringIO

url = "https://quality.data.gov.tw/dq_download_csv.php?nid=18419"

# ===== 正確抓資料 =====
res = requests.get(url)

if res.status_code != 200:
    raise Exception("❌ 下載失敗")

# 👉 用 StringIO 轉成 CSV
df = pd.read_csv(StringIO(res.text))

# ===== 找欄位 =====
code_col = None
cap_col = None

for col in df.columns:
    if "代號" in col:
        code_col = col
    if "股本" in col:
        cap_col = col

if code_col is None or cap_col is None:
    raise Exception(f"❌ 找不到欄位: {df.columns}")

# ===== 整理 =====
df = df[[code_col, cap_col]].copy()
df.columns = ["證券代號", "股本"]

# ===== 清洗 =====
df["證券代號"] = df["證券代號"].astype(str).str.strip()

df["股本"] = (
    df["股本"]
    .astype(str)
    .str.replace(",", "")
    .replace("", "0")
    .astype(float)
)

# 只留4碼股票
df = df[df["證券代號"].str.match(r"^\d{4}$")]

# ===== 輸出 =====
df.to_csv("cap.csv", index=False, encoding="utf-8-sig")

print(f"✅ cap.csv 更新完成，共 {len(df)} 筆")
