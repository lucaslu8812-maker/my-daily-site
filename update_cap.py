import pandas as pd
import requests

def update_cap():
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&type=ALL"
    data = requests.get(url, timeout=10).json()

    if "data9" not in data:
        print("❌ 抓不到資料")
        return

    df = pd.DataFrame(data["data9"], columns=data["fields9"])

    code_col = None
    cap_col = None

    for col in df.columns:
        if "證券代號" in col:
            code_col = col
        if "股本" in col or "資本" in col:
            cap_col = col

    if not code_col or not cap_col:
        print("❌ 欄位錯誤")
        return

    df = df[[code_col, cap_col]]
    df.columns = ["證券代號", "股本"]

    df["股本"] = (
        df["股本"].astype(str)
        .str.replace(",", "")
        .replace("", "0")
        .astype(float)
    )

    df["發行股數"] = df["股本"] * 1000

    df[["證券代號", "發行股數"]].to_csv("cap.csv", index=False, encoding="utf-8-sig")

    print("✅ cap.csv 更新完成")

if __name__ == "__main__":
    update_cap()
