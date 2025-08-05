import pandas as pd
import json
import os
import datetime
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import requests
from io import BytesIO
import pytz
from matplotlib.font_manager import FontProperties
# Google Drive 分享連結的檔案 ID
file_id = "1mVM7IlhmSqe85cghnIWuDfbXmpOEn4Qx"
download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

# 載入 Excel 檔案
response = requests.get(download_url)
excel_data = BytesIO(response.content)

df = pd.read_excel(excel_data, engine='openpyxl')

max_date = pd.to_datetime(df["資料日期"]).max()
# 轉換為中文格式
data_date_str = max_date.strftime("%Y-%m-%d %H:%M:%S")

tz = pytz.timezone('Asia/Taipei')  # UTC+8 時區
run_time_str = datetime.datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')
print(f"資料日期: {data_date_str}\n程式執行時間: {run_time_str}")
# 去除不要的欄位
drop_cols = ['試算價', '試算損益']
for col in drop_cols:
    if col in df.columns:
        df = df.drop(columns=[col])

# 去除關鍵欄位空值
df = df.dropna(subset=['損益率', '商品名稱', '項次'])

# 損益率：去除百分比符號，轉 float，乘 100（變成百分比數字）
df['損益率'] = df['損益率'].astype(str).str.replace('%', '', regex=False).str.strip()
df = df[df['損益率'] != '']
df['損益率'] = df['損益率'].astype(float) * 100
df['損益率'] = df['損益率'].round(2)

# 整理金額欄位格式（字串帶逗號與「元」）
money_cols = ['投資成本', '帳面收入', '損益', '市值']
for col in money_cols:
    if col in df.columns:
        df[col] = df[col].apply(lambda x: f"{int(x):,} 元" if pd.notna(x) else "")

# 數值型欄位，方便繪圖用
df['投資成本_數值'] = df['投資成本'].str.replace(' 元', '').str.replace(',', '').astype(float)
df['市值_數值'] = df['市值'].str.replace(' 元', '').str.replace(',', '').astype(float)

# 產生圖表用資料
labels = df['商品名稱'].astype(str).tolist()
profit_rates = df['損益率'].tolist()
investment_costs = df['投資成本_數值'].tolist()
market_values = df['市值_數值'].tolist()
shares = df['股數'].tolist()

def profit_category(pct):
    if pct >= 20:
        return "大幅獲利 >= 20%"
    elif 10 <= pct < 20:
        return "中度獲利 10~20%"
    elif 0 <= pct < 10:
        return "小幅獲利 0~10%"
    elif -10 <= pct < 0:
        return "小幅虧損 -10~0%"
    elif -20 <= pct < -10:
        return "中度虧損 -20~-10%"
    else:
        return "重度虧損 < -20%"


df['損益區間'] = df['損益率'].apply(profit_category)
cost_by_category = df.groupby('損益區間')['投資成本_數值'].sum().to_dict()

# JSON 序列化
labels_json = json.dumps(labels)
profit_rates_json = json.dumps(profit_rates)
investment_costs_json = json.dumps(investment_costs)
market_values_json = json.dumps(market_values)
shares_json = json.dumps(shares)
cost_by_category_json = json.dumps(cost_by_category)

# 計算總和資訊
total_investment = int(df['投資成本_數值'].sum())
total_market_value = int(df['市值_數值'].sum())
total_profit = int(df['損益'].str.replace(' 元', '').str.replace(',', '').astype(float).sum())
total_profit_rate = round(total_profit / total_investment * 100, 2) if total_investment != 0 else 0

if os.name == "nt":
    # Windows 環境
    font_path = r"C:\Windows\Fonts\msyh.ttc"  # Microsoft YaHei 字型檔案
else:
    # GitHub Actions / Linux 環境
    font_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"

# 透過 FontProperties 載入
font_prop = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()
font_size = 14  # 統一字體大小

# 損益率長條圖
plt.figure(figsize=(16, 12))
bars = plt.bar(labels, profit_rates, color=['red' if x >= 0 else 'green' for x in profit_rates])
plt.title(f"{data_date_str} 投資損益率（共 {len(labels)} 檔）", fontproperties=font_prop, fontsize=font_size * 1.5)
plt.ylabel("損益率 (%)", fontproperties=font_prop, fontsize=font_size)
plt.xticks(rotation=45, ha='right', fontproperties=font_prop, fontsize=font_size)
plt.yticks(fontsize=font_size)
plt.axhline(0, color='black', linewidth=0.8)

for bar, rate in zip(bars, profit_rates):
    va = 'bottom' if rate >= 0 else 'top'
    plt.text(
        bar.get_x() + bar.get_width() / 2, 
        bar.get_height(), 
        f"{rate:.1f}%", 
        ha='center', 
        va=va, 
        fontproperties=font_prop, 
        fontsize=font_size-4
    )

plt.tight_layout()
plt.savefig("docs/profit_rate_bar.png")
plt.close()

# 損益區間圓餅圖
plt.figure(figsize=(16, 12))
labels_pie = list(cost_by_category.keys())
sizes_pie = list(cost_by_category.values())

plt.pie(
    sizes_pie, 
    labels=labels_pie, 
    autopct='%1.1f%%', 
    startangle=140,
    textprops={'fontproperties': font_prop, 'fontsize': font_size}
)
plt.title("投資成本佔比（依損益區間分類）", fontproperties=font_prop, fontsize=font_size * 3)
plt.axis('equal')
plt.tight_layout()
plt.savefig("docs/profit_category_pie.png")
plt.close()

# 市值占比圓環圖（小於 5% 合併為「其他」）
total_value = sum(market_values)
threshold_pct = 5  # 小於 5% 合併為其他

# 重新整理 labels 和 values，把小於 threshold 的合併為 "其他"
new_labels = []
new_values = []
other_value = 0

for label, value in zip(labels, market_values):
    pct = value / total_value * 100
    if pct < threshold_pct:
        other_value += value
    else:
        new_labels.append(label)
        new_values.append(value)

# 若有其他，加入圖例與數值
if other_value > 0:
    new_labels.append("其他")
    new_values.append(other_value)


# 畫圖
plt.figure(figsize=(16, 12))
wedges, texts, autotexts = plt.pie(
    new_values,
    labels=None,
    autopct=lambda pct: f'{pct:.1f}%' if pct > 0 else '',
    startangle=140,
    colors=plt.cm.tab20.colors[:len(new_labels)],
    textprops={'fontproperties': font_prop, 'fontsize': font_size * 1.5},
    pctdistance=0.75
)

# 中心圓形遮罩（圓環）
centre_circle = plt.Circle((0, 0), 0.60, fc='white')
fig = plt.gcf()
fig.gca().add_artist(centre_circle)

legend_font = FontProperties(fname=font_path, size=font_size*1.3)  # 或你想要的倍數

legend = plt.legend(
    wedges,
    new_labels,
    title="商品名稱",
    loc="center left",
    bbox_to_anchor=(0.9, 0.5),
    prop=legend_font
)
legend.get_title().set_fontsize(font_size*1.3)  # 圖例標題大小也可放大

# 圖標題
plt.title("市值占比（依商品分類）", fontproperties=font_prop, fontsize=font_size * 3)
plt.axis('equal')

# 儲存圖檔，不使用 tight
plt.savefig("docs/market_value_doughnut.png", dpi=300)
plt.close()


with open("docs/investment_report.md", "w", encoding="utf-8") as f:
    f.write(f"# 投資損益報告\n\n")
    f.write(f"📅 資料日期：{data_date_str}　🕒 產生時間：{run_time_str}\n\n")
    f.write(f"## 總覽\n")
    f.write(f"- 💰 總投資金額：{total_investment:,} 元\n")
    f.write(f"- 📈 市值總額：{total_market_value:,} 元\n")
    f.write(f"- 🧮 總損益：{total_profit:,} 元\n")
    f.write(f"- 📊 報酬率：{total_profit_rate:.2f}%\n\n")

    f.write("## 圖表\n")
    f.write("### 損益率長條圖\n")
    f.write("![損益率](profit_rate_bar.png)\n\n")
    f.write("### 損益區間圓餅圖\n")
    f.write("![損益區間](profit_category_pie.png)\n\n")
    f.write("### 市值占比圓環圖\n")
    f.write("![市值占比](market_value_doughnut.png)\n\n")

    f.write("## 各股明細\n\n")
    f.write("| 商品名稱 | 股數 | 成本價 | 投資成本 | 帳面收入 | 損益 | 損益率 | 現價 | 市值 |\n")
    f.write("|----------|------|--------|------------|------------|--------|----------|--------|------------|\n")

    for _, row in df.iterrows():
        gain_str = row['損益'].replace(',', '').replace('元', '').strip()
        gain_rate_str = str(row['損益率']).replace('%', '').strip()

        gain = float(gain_str)
        gain_rate = float(gain_rate_str)

        color = "green" if gain >= 0 else "red"
        gain_colored = f"<span style='color:{color}'>{gain:,.0f} 元</span>"
        gain_rate_colored = f"<span style='color:{color}'>{gain_rate:.2f}%</span>"

        f.write(
            f"| {row['商品名稱']} | {row['股數']} | {row['成本價']} | {row['投資成本']} | "
            f"{row['帳面收入']} | {gain_colored} | {gain_rate_colored} | {row['現價']} | {row['市值']} |\n"
        )

# 產生 HTML
html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8" />
<title>投資損益分析報告</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />

<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC&display=swap" rel="stylesheet" />
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css" />

<style>
  body {{
    font-family: 'Noto Sans TC', sans-serif;
    background: linear-gradient(135deg, #f0f4ff, #d9e4ff);
    color: #1a237e;
    margin: 20px;
    font-size: 18px;
  }}
  h1 {{
    text-align: center;
    margin-bottom: 20px;
    font-size: 3rem;
  }}
  .cards {{
    display: flex;
    gap: 20px;
    justify-content: center;
    flex-wrap: wrap;
    margin-bottom: 40px;
  }}
  .card {{
    background: white;
    box-shadow: 0 4px 12px rgba(26,35,126,.15);
    border-radius: 12px;
    padding: 25px 40px;
    min-width: 200px;
    text-align: center;
  }}
  .card h2 {{
    margin: 0 0 10px;
    font-size: 3rem;
    color: #0d47a1;
  }}
  .card p {{
    margin: 0;
    font-size: 1.3rem;
    color: #3949ab;
    font-weight: 600;
  }}
  
  #charts {{
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 40px;
    margin-bottom: 40px;
    min-height: 520px;
  }}
  canvas {{
  background: white;
  border-radius: 12px;
  box-shadow: 0 4px 12px rgba(26,35,126,.1);
  padding: 20px;
  width: 100% !important;   /* 寬度改成100%，隨父容器寬度縮放 */
  height: auto !important;  /* 高度自動調整或用比例 */

  aspect-ratio: 16 / 9;     /* 維持寬高比 */
  }}
  table.dataTable {{
    border-collapse: collapse !important;
    background: white;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(26,35,126,.15);
    overflow: hidden;
  }}
  table.dataTable thead th {{
    background-color: #283593 !important;
    color: white !important;
    font-weight: 700;
    font-size: 1rem;
  }}
  table.dataTable tbody tr:hover {{
    background-color: #e3eafc !important;
  }}
  table.dataTable tbody td {{
    text-align: center;
    font-size: 1rem;
  }}
  .negative {{
    color: #e53935;
    font-weight: 700;
  }}
  .positive {{
    color: #43a047;
    font-weight: 700;
  }}
  footer {{
    text-align: center;
    margin-top: 60px;
    color: #666;
    font-size: 14px;
  }}
</style>

</head>
<body>

<h1>投資損益分析報告</h1>
<p style="text-align:center; color:#666; font-size:14px; margin-top:-10px; margin-bottom:30px;">
  資料日期：{data_date_str}
</p>
<p style="text-align:center; color:#666; font-size:14px; margin-top:-10px; margin-bottom:30px;">
  程式執行時間：{run_time_str}
</p>
<div class="cards">
  <div class="card">
    <h2>{total_investment:,} 元</h2>
    <p>總投資成本</p>
  </div>
  <div class="card">
    <h2>{total_market_value:,} 元</h2>
    <p>總市值</p>
  </div>
  <div class="card">
    <h2>{total_profit:,} 元</h2>
    <p>帳面損益</p>
  </div>
  <div class="card">
    <h2>{total_profit_rate}%</h2>
    <p>整體損益率</p>
  </div>
</div>

<div id="charts">
  <canvas id="barChart"></canvas>
  <canvas id="bubbleChart"></canvas>
  <canvas id="pieChart"></canvas>
</div>

<table id="detailTable" class="display" style="width:100%">
<thead>
  <tr>
    <th>項次</th>
    <th>商品名稱</th>
    <th>類別</th>
    <th>股數</th>
    <th>成本價</th>
    <th>投資成本</th>
    <th>帳面收入</th>
    <th>損益</th>
    <th>損益率</th>
    <th>現價</th>
    <th>市值</th>
    <th>幣別</th>
  </tr>
</thead>
<tbody>
"""

for _, row in df.iterrows():
    profit_rate_class = "positive" if row['損益率'] >= 0 else "negative"
    html += f"""
    <tr>
      <td>{row['項次']}</td>
      <td>{row['商品名稱']}</td>
      <td>{row['類別']}</td>
      <td>{row['股數']}</td>
      <td>{row['成本價']}</td>
      <td>{row['投資成本']}</td>
      <td>{row['帳面收入']}</td>
      <td>{row['損益']}</td>
      <td class="{profit_rate_class}">{row['損益率']}%</td>
      <td>{row['現價']}</td>
      <td>{row['市值']}</td>
      <td>{row['幣別']}</td>
    </tr>
    """

html += """
</tbody>
</table>
<div style="max-width:600px;margin:40px auto;text-align:center;">
  <h2 style="margin-bottom:10px;">💰 投資金額分配模擬</h2>
  <input type="number" id="allocationInput" placeholder="請輸入總投資金額（元）" style="font-size:1.2rem;padding:10px;width:80%;max-width:400px;">
  <button id="allocateBtn" style="font-size:1.2rem;padding:10px 20px;margin-top:10px;">分配金額</button>
  <div id="allocationResult" style="margin-top:30px;text-align:left;"></div>
</div>

<footer>
  <p>報告由 ChatGPT 根據使用者提供資料自動生成</p>
</footer>

<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<script>
const labels = """ + labels_json + """;
const investmentCosts = """ + investment_costs_json + """;
const marketValues = """ + market_values_json + """;
const profitRates = """ + profit_rates_json + """;
const shares = """ + shares_json + """;
const costByCategory = """ + cost_by_category_json + """;

// Bar Chart: 市值與投資成本比較
new Chart(document.getElementById('barChart').getContext('2d'), {
  type: 'bar',
  data: {
    labels: labels,
    datasets: [
      {
        label: '投資成本',
        data: investmentCosts,
        backgroundColor: 'rgba(26, 35, 126, 0.7)'
      },
      {
        label: '市值',
        data: marketValues,
        backgroundColor: 'rgba(67, 160, 71, 0.7)'
      }
    ]
  },
  options: {
    responsive: true,

    scales: {
      y: {
        beginAtZero: true,
        ticks: {
          font: { size: 22 },
          callback: value => value.toLocaleString() + ' 元'
        },
        title: {
          display: true,
          text: '金額 (元)',
          font: { size: 24 }
        }
      },
      x: {
        ticks: {
          maxRotation: 90,
          minRotation: 45,
          autoSkip: false,
          font: { size: 20 }
        }
      }
    },
    plugins: {
      legend: { labels: { font: { size: 20 }, padding: 20 } },
      tooltip: {
        bodyFont: { size: 20 },
        titleFont: { size: 22 },
        padding: 16,
        callbacks: {
          label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString() + ' 元'
        }
      }
    }
  }
});

// Bubble Chart: 投入成本與損益率關係，股數作為氣泡大小
new Chart(document.getElementById('bubbleChart').getContext('2d'), {
  type: 'bubble',
  data: {
    labels: labels,
    datasets: [{
      label: '投入成本 vs 損益率',
      data: labels.map((label, i) => {
        return {
          x: investmentCosts[i],
          y: profitRates[i],
          r: Math.sqrt(Number(shares[i]) || 1) * 2
        };
      }),
      backgroundColor: 'rgba(30, 136, 229, 0.7)'
    }]
  },
  options: {
    responsive: true,

    scales: {
      x: {
        beginAtZero: true,
        title: {
          display: true,
          text: '投入成本 (元)',
          font: { size: 24 }
        },
        ticks: {
          font: { size: 20 },
          callback: val => val.toLocaleString()
        }
      },
      y: {
        beginAtZero: false,
        title: {
          display: true,
          text: '損益率 (%)',
          font: { size: 24 }
        },
        ticks: {
          font: { size: 20 },
          callback: val => val + '%'
        }
      }
    },
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx =>
            labels[ctx.dataIndex] + ': 投入成本 ' + ctx.parsed.x.toLocaleString() + ' 元, 損益率 ' + ctx.parsed.y.toFixed(2) + '%, 股數 ' + shares[ctx.dataIndex]
        },
        bodyFont: { size: 20 },
        titleFont: { size: 22 },
        padding: 16
      }
    }
  }
});

// Pie Chart: 各損益區間投入成本比例
new Chart(document.getElementById('pieChart').getContext('2d'), {
  type: 'pie',
  data: {
    labels: Object.keys(costByCategory),
    datasets: [{
      data: Object.values(costByCategory),
      backgroundColor: ['#1e88e5', '#42a5f5', '#90caf9', '#c5cae9']
    }]
  },
  options: {
    responsive: true,

    plugins: {
      legend: {
        position: 'right',
        labels: {
          boxWidth: 20,
          padding: 15,
          font: { size: 20 }
        }
      },
      tooltip: {
        bodyFont: { size: 20 },
        titleFont: { size: 22 },
        padding: 16
      }
    }
  }
});

$(document).ready(function() {
  $('#detailTable').DataTable({
    pageLength: 10,
    lengthMenu: [5,10,20,50],
    language: {
      search: "搜尋：",
      lengthMenu: "顯示 _MENU_ 筆",
      info: "顯示第 _START_ 筆到 _END_ 筆，共 _TOTAL_ 筆",
      paginate: {
        first: "第一頁",
        last: "最後一頁",
        next: "下一頁",
        previous: "上一頁"
      },
      zeroRecords: "找不到符合的資料"
    },
    columnDefs: [
      { targets: [3,4,5,6,7,9,10], className: 'dt-center' }
    ]
  });
});

document.getElementById('allocateBtn').addEventListener('click', () => {
  const totalAmount = parseFloat(document.getElementById('allocationInput').value);
  const resultBox = document.getElementById('allocationResult');
  resultBox.innerHTML = '';

  if (isNaN(totalAmount) || totalAmount <= 0) {
    resultBox.innerHTML = '<p style="color:red;">請輸入有效的投資金額</p>';
    return;
  }

  // 篩選正報酬的項目
  const positiveItems = [];
  let sumProfitRate = 0;
  for (let i = 0; i < labels.length; i++) {
    const rate = profitRates[i];
    if (rate > 0) {
      positiveItems.push({ name: labels[i], rate });
      sumProfitRate += rate;
    }
  }

  if (positiveItems.length === 0) {
    resultBox.innerHTML = '<p style="color:red;">目前沒有報酬率為正的項目</p>';
    return;
  }

  // 按比例分配金額
  const allocations = positiveItems.map(item => {
    const allocated = totalAmount * (item.rate / sumProfitRate);
    return { ...item, allocated };
  });

  // 顯示結果
  let html = '<table style="width:100%;border-collapse:collapse;margin-top:10px;">';
  html += '<tr style="background-color:#3f51b5;color:white;"><th style="padding:10px;">商品</th><th style="padding:10px;">損益率</th><th style="padding:10px;">分配金額</th></tr>';
  allocations.forEach(item => {
    html += `<tr style="text-align:center;"><td style="padding:8px;">${item.name}</td><td style="padding:8px;">${item.rate.toFixed(2)}%</td><td style="padding:8px;">${Math.round(item.allocated).toLocaleString()} 元</td></tr>`;
  });
  html += '</table>';
  resultBox.innerHTML = html;
});
</script>

</body>
</html>
"""

# 寫出 HTML 檔案
with open("docs/投資損益分析報告.html", "w", encoding="utf-8") as f:
    f.write(html)

print("已生成 投資損益分析報告.html")
