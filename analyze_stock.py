import pandas as pd
import json
import os
import datetime

# 讀取 Excel
file_path = "未實現損益試算.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

timestamp = os.path.getmtime(file_path)

# 轉換成 datetime 物件
data_date = datetime.datetime.fromtimestamp(timestamp)

# 格式化成中文日期字串，例如：2025年07月24日
data_date_str = data_date.strftime('%Y年%m月%d日')

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

# 損益區間分類函數，用於圓餅圖
def profit_category(pct):
    if pct >= 20:
        return "大幅獲利 ≥20%"
    elif pct >= 0:
        return "獲利 0~20%"
    elif pct >= -10:
        return "小幅虧損 -10%~0"
    else:
        return "重度虧損 < -10%"

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
</script>

</body>
</html>
"""

# 寫出 HTML 檔案
with open("投資損益分析報告.html", "w", encoding="utf-8") as f:
    f.write(html)

print("已生成 投資損益分析報告.html")
