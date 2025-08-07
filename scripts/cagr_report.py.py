import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta


tickers = {
    "2330": "2330.TW",
    "2539": "2539.TW",
    "0050": "0050.TW",
    "加權指數": "^TWII",
    "S&P500": "^GSPC",
    "NASDAQ": "^IXIC"
}

periods = [5, 10, 15, 20]
end_date = datetime.today()

results = {p: {} for p in periods}
date_ranges = {}

def calc_cagr(start_price, end_price, years):
    return (end_price / start_price) ** (1 / years) - 1

for name, ticker in tickers.items():
    # 先抓滿足最長期間的資料，這裡先拉最大跨度(例如20年再多抓點)
    start_year = end_date.year - max(periods) - 1  # 多抓1年保險
    start_date = end_date.replace(year=start_year)

    data = yf.download(ticker, start=start_date, end=end_date, progress=False, auto_adjust=True)
    if data.empty:
        date_ranges[name] = (None, None)
        for p in periods:
            results[p][name] = None
        continue

    date_ranges[name] = (data.index[0].strftime('%Y-%m-%d'), data.index[-1].strftime('%Y-%m-%d'))

    # 判斷資料是否夠20年（用實際資料起始日比對）
    earliest_date = data.index[0]

    for period in periods:
        period_start_date = end_date - pd.DateOffset(years=period)
        # 如果資料起始日比需要的起始日還晚，表示資料不足該期
        if earliest_date > period_start_date:
            results[period][name] = None
            continue

        # 篩選該區間資料
        period_data = data.loc[data.index >= period_start_date]
        if period_data.empty:
            results[period][name] = None
            continue

        start_price = period_data['Close'].iloc[0]
        end_price = period_data['Close'].iloc[-1]

        results[period][name] = round(calc_cagr(start_price, end_price, period), 4)

# 之後照你原本寫的格式輸出即可
df = pd.DataFrame(results).T

def safe_format(v):
    import numpy as np
    if isinstance(v, (pd.Series, np.ndarray)):
        if len(v) == 0:
            return "N/A"
        # pandas Series 用 iloc，numpy ndarray 用 [0]
        v = v.iloc[0] if isinstance(v, pd.Series) else v[0]
    if v is None or pd.isna(v):
        return "N/A"
    try:
        return f"{float(v)*100:.2f}%"
    except:
        return "N/A"


with open("annualized_return_real.md", "w", encoding="utf-8") as f:
    f.write("# 真實年化報酬率 (CAGR) 換算表\n\n")
    f.write("單位：年化報酬率 (百分比，四捨五入至小數點後兩位)\n\n")

    f.write("## 資料範圍\n\n")
    f.write("| 資產 | 資料起始日 | 資料結束日 |\n")
    f.write("| --- | --- | --- |\n")
    for name, (start, end) in date_ranges.items():
        start = start if start is not None else "N/A"
        end = end if end is not None else "N/A"
        f.write(f"| {name} | {start} | {end} |\n")

    f.write("\n---\n")

    headers = ["期間"] + list(df.columns)
    f.write("| " + " | ".join(headers) + " |\n")
    f.write("|" + " --- |" * len(headers) + "\n")

    for row in df.itertuples():
        idx = row[0]
        row_str = [safe_format(v) for v in row[1:]]
        f.write(f"| {idx}年 | " + " | ".join(row_str) + " |\n")

    f.write("\n---\n")
    f.write("### 說明\n")
    f.write("- 表中年化報酬率 (CAGR) 根據指定期間的收盤價計算。\n")
    f.write("- CAGR 計算公式：\n\n")
    f.write("  ```\n")
    f.write("  CAGR = (終值 / 初值)^(1/年數) - 1\n")
    f.write("  ```\n\n")
    f.write("- 資料來源：Yahoo Finance。\n")
