import subprocess
import time
import csv
from datetime import datetime
import re
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 支援中文字體（微軟正黑體）
plt.rcParams['axes.unicode_minus'] = False  # 正確顯示負號

# --- 設定 ---
target = "8.8.8.8"
log_time = 60  # 紀錄時間（秒），可自行調整

# --- 檔名 ---
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
csv_file = f"ping_log_{timestamp}.csv"
png_file = f"ping_plot_{timestamp}.png"

# --- 寫入 CSV ---
with open(csv_file, mode="w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["時間", "延遲(ms)", "狀態"])

# --- 開始紀錄 ---
print(f"⏱️ 開始紀錄 Ping（{log_time} 秒）")
print(f"📄 CSV：{csv_file}")

times = []
latencies = []

for i in range(log_time):
    result = subprocess.run(["ping", target, "-n", "1"], capture_output=True, text=True, encoding="cp950")
    output = result.stdout

    match = re.search(r"平均 = (\d+)ms", output)
    now = datetime.now().strftime('%H:%M:%S')

    if match:
        latency = int(match.group(1))
        status = "正常" if latency < 80 else "高延遲⚠️"
        print(f"[{now}] {latency}ms - {status}")
    else:
        latency = None
        status = "無法解析⚠️"
        print(f"[{now}] ⚠️ Ping 失敗")

    # 存檔
    with open(csv_file, mode="a", newline="") as file:
        writer = csv.writer(file)
        writer.writerow([now, latency if latency else "N/A", status])

    # 儲存資料供畫圖用
    times.append(now)
    latencies.append(latency if latency else 0)

    time.sleep(1)

# --- 繪圖 ---
plt.figure(figsize=(12, 6))
plt.plot(times, latencies, marker='o', linestyle='-', label='Ping (ms)')
plt.axhline(80, color='red', linestyle='--', label='高延遲門檻 80ms')
plt.xticks(rotation=45)
plt.xlabel("時間")
plt.ylabel("延遲 (ms)")
plt.title("Ping 延遲紀錄")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(png_file)
print(f"📈 圖表儲存為：{png_file}")
