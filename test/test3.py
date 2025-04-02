import mss
import pytesseract
from PIL import Image
import time
import matplotlib.pyplot as plt
import csv
from datetime import datetime

# 指定 tesseract 執行檔路徑
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# 擷取畫面右上角的 Ping 數字和幀數顯示區域
bbox = {'top': 50, 'left': 2200, 'width': 400, 'height': 50}  # 根據解析度微調

# 記錄時間和數據
times = []
fps_values = []
ping_values = []

# 輸出 CSV 文件名稱
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
csv_file = f"fps_ping_log_{timestamp}.csv"

# 寫入 CSV 標題
with open(csv_file, mode="w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["時間", "幀數", "延遲 (ms)"])

print(f"開始擷取 幀數 和 延遲 ...（按 Ctrl+C 停止）")

try:
    while True:
        with mss.mss() as sct:
            img = sct.grab(bbox)
            img_pil = Image.frombytes('RGB', img.size, img.rgb)

            # 使用繁體中文語言包，並只抓取數字和毫秒
            text = pytesseract.image_to_string(img_pil, lang='chi_tra', config='--psm 6 -c tessedit_char_whitelist=0123456789ms')

            now = datetime.now().strftime("%H:%M:%S")

            # 解析文本，分別抓取幀數和延遲（毫秒）
            parts = text.strip().split()
            if len(parts) >= 2:
                fps = parts[0]  # 幀數
                ping = parts[1].replace("毫秒", "")  # 延遲（毫秒）
                
                if fps.isdigit() and ping.isdigit():
                    print(f"[{now}] 幀數: {fps}，延遲: {ping} ms")
                    
                    # 記錄數據
                    times.append(now)
                    fps_values.append(int(fps))
                    ping_values.append(int(ping))

                    # 存到 CSV
                    with open(csv_file, mode="a", newline="") as file:
                        writer = csv.writer(file)
                        writer.writerow([now, fps, ping])

        time.sleep(1)

except KeyboardInterrupt:
    print("\n結束擷取")

