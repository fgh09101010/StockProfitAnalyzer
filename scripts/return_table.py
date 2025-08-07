import pandas as pd

# 建立一個 DataFrame 儲存結果
data = []

# 計算每種年化報酬率在每一年的累積報酬
for i in range(1, 11):  # 年化報酬率 1% ~ 10%
    row = []
    for j in range(1, 41):  # 第 1 年 ~ 第 40 年
        cumulative_return = (1 + i * 0.01) ** j - 1
        row.append(round(cumulative_return, 4))  # 四捨五入至小數點後 4 位
    data.append(row)

# 建立 DataFrame，接著轉置
df = pd.DataFrame(data, index=[f"{i}%" for i in range(1, 11)], columns=[f"第{j}年" for j in range(1, 41)])
df_T = df.T  # 轉置：row → 年數，columns → 年化報酬率


with open("報酬率累積.txt", "w", encoding="utf-8") as f:
    f.write("# 年化報酬率累積表（轉置：年數為列，報酬率為欄）\n\n")
    f.write("單位：累積報酬率（四捨五入至小數點後四位，含複利）\n\n")

    # 表頭
    headers = ["年數"] + list(df_T.columns)
    f.write("| " + " | ".join(headers) + " |\n")
    f.write("|" + " --- |" * len(headers) + "\n")

    # 表格內容
    for idx, row in df_T.iterrows():
        row_str = [f"{v:.4f}" for v in row]
        f.write(f"| {idx} | " + " | ".join(row_str) + " |\n")

