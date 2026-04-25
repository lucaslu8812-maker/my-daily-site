import pandas as pd
import requests
import os

url = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"

try:
    res = requests.get(url, timeout=10)

    # ⭐ 檢查 HTTP 狀態
    if res.status_code != 200:
        raise Exception(f"HTTP錯誤: {res.status_code}")

    # ⭐ 檢查內容（避免空內容）
    if not res.text.strip():
        raise Exception("API回傳空資料")

    try:
        data = res.json()
    except:
        raise Exception("JSON解析失敗（可能被擋或API掛掉）")

    df = pd.DataFrame(data)

    # ===== 找欄位 =====
    code_col = None
    cap_col = None

    for col in df.columns:
        if "公司代號" in col:
            code_col = col
        if "實收資本額" in col:
            cap_col = col

    if code_col is None or cap_col is None:
        raise Exception(f"找不到欄位: {df.columns}")

    # ===== 整理 =====
    df = df[[code_col, cap_col]].copy()
    df.columns = ["證券代號", "股本"]

    df["證券代號"] = df["證券代號"].astype(str).str.zfill(4)

    df["股本"] = (
        df["股本"]
        .astype(str)
        .str.replace(",", "")
        .replace("", "0")
        .astype(float)
    )

    df["股本"] = df["股本"] / 100000000

    df = df[df["證券代號"].str.match(r"^\d{4}$")]

    df.to_csv("cap.csv", index=False, encoding="utf-8-sig")

    print(f"✅ cap.csv 更新完成，共 {len(df)} 筆")

except Exception as e:
    print(f"⚠️ cap API 失敗: {e}")

    # ⭐ 關鍵：如果已有 cap.csv → 不覆蓋
    if os.path.exists("cap.csv"):
        print("📁 保留舊的 cap.csv")
    else:
        print("❌ 沒有 cap.csv 可用")
