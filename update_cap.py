import pandas as pd

url = "https://mopsfin.twse.com.tw/opendata/t187ap03_L.csv"

df = pd.read_csv(url, encoding="utf-8")

# 找股數欄（最重要）
target_col = None
for col in df.columns:
    if "股數" in col:
        target_col = col
        break

if target_col is None:
    raise Exception(f"❌ 找不到股數欄位: {df.columns}")

df = df[["公司代號", target_col]]
df.columns = ["證券代號", "發行股數"]

# 轉數字
df["發行股數"] = (
    df["發行股數"]
    .astype(str)
    .str.replace(",", "")
    .astype(float)
)

df.to_csv("cap.csv", index=False, encoding="utf-8")

print("✅ cap.csv 更新完成")
