import pandas as pd

url = "https://raw.githubusercontent.com/FinMind/FinMind/master/dataset/TaiwanStockInfo.csv"

df = pd.read_csv(url)

# 只留上市股票
df = df[df["type"] == "twse"]

df = df[["stock_id", "stock_name"]]
df.columns = ["證券代號", "證券名稱"]

# 👉 假設股本（簡化用）
df["發行股數"] = 1_000_000_000  # 先給固定值避免炸

df.to_csv("cap.csv", index=False, encoding="utf-8-sig")

print("✅ cap.csv 建立（簡化版）")
