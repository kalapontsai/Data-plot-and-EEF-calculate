
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import tkinter as tk
from tkinter import filedialog

# 強制只用微軟正黑體，避免 fallback 到 DejaVu Sans
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
matplotlib.rcParams['axes.unicode_minus'] = False

def select_csv_file():
    """開啟檔案對話框選擇CSV檔案"""
    root = tk.Tk()
    root.withdraw()  # 隱藏主視窗
    
    file_path = filedialog.askopenfilename(
        title="選擇CSV檔案",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    
    root.destroy()
    return file_path

# 1. 選擇並讀取 gx20紀錄csv檔
input_file = select_csv_file()

if not input_file:
    print("未選擇檔案，程式結束")
    exit()

df = pd.read_csv(input_file)

# 2. 欄位 'Date' + 'Time' 為 X 軸
df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])

# 3. 設定閾值，計算 P(W) 數值為 'on' 或 'off' 屬性
THRESHOLD = 25  # 可依需求調整
df['status'] = df['P(W)'].apply(lambda x: 'on' if x >= THRESHOLD else 'off')

# 4. 計算每段連續 'on' 或 'off' 的時間長度（分鐘）
period_lengths = [0] * len(df)
current_status = df['status'].iloc[0]
current_start = 0
for i in range(1, len(df)):
	if df['status'].iloc[i] != current_status:
		start_time = df['datetime'].iloc[current_start]
		end_time = df['datetime'].iloc[i-1]
		minutes = int((end_time - start_time).total_seconds() // 60) + 1
		for j in range(current_start, i):
			period_lengths[j] = minutes
		current_status = df['status'].iloc[i]
		current_start = i
# 處理最後一段
start_time = df['datetime'].iloc[current_start]
end_time = df['datetime'].iloc[len(df)-1]
minutes = int((end_time - start_time).total_seconds() // 60) + 1
for j in range(current_start, len(df)):
	period_lengths[j] = minutes


# 5. 新增 'period_on' 與 'period_off' 欄位，並用 interpolate 讓線條平滑連接
import numpy as np
df['period_on'] = [p if s == 'on' else np.nan for p, s in zip(period_lengths, df['status'])]
df['period_off'] = [p if s == 'off' else np.nan for p, s in zip(period_lengths, df['status'])]
df['period_on_interp'] = df['period_on'].interpolate()
df['period_off_interp'] = df['period_off'].interpolate()

# 6. 計算WP(Wh)每60分鐘的差值
# 先建立一個空的Series來存放差值
wp_diff = pd.Series(index=df.index, dtype=float)

# 每60筆資料計算一次差值，並對該區間內所有點使用相同的差值
for i in range(0, len(df), 60):
    if i + 60 < len(df):
        diff_value = df['WP(Wh)'].iloc[i+59] - df['WP(Wh)'].iloc[i]
        wp_diff[i:i+60] = diff_value

# 將結果存入DataFrame，並用前值填補最後不足60筆的資料
df['WP_diff'] = wp_diff.fillna(0)

# 7. 繪製三個子圖的折線圖
fig, axs = plt.subplots(3, 1, figsize=(12, 10), sharex=True)

# 第1個子圖：原始資料除了 'Date', 'Time', 'P(W)' 欄位的折線圖
for col in df.columns:
	if col not in ['Date', 'Time', 'U(V)', 'I(A)', 'P(W)', 'WP(Wh)', 'datetime', 'status', 'period', 'period_on', 'period_off', 'period_on_interp', 'period_off_interp', 'WP_diff']:
		axs[0].plot(df['datetime'], df[col], label=col)
axs[0].set_ylabel('Temperature')
axs[0].legend()
axs[0].set_title(input_file.split('.')[0])  # 使用檔名（不含副檔名）作為標題

# 第2個子圖：'period_on' 與 'period_off' 兩條線（平滑連接）
axs[1].plot(df['datetime'], df['period_on_interp'], label='on', color='red')
axs[1].plot(df['datetime'], df['period_off_interp'], label='off', color='blue')
axs[1].set_ylabel('minutes')
axs[1].set_title('on/off 時間長度')
axs[1].legend()

# 計算WP(Wh)每60分鐘的差值
# 先建立一個空的Series來存放差值
wp_diff = pd.Series(index=df.index, dtype=float)

# 每60筆資料計算一次差值，並對該區間內所有點使用相同的差值
for i in range(0, len(df), 60):
    if i + 60 < len(df):
        diff_value = df['WP(Wh)'].iloc[i+59] - df['WP(Wh)'].iloc[i]
        wp_diff[i:i+60] = diff_value

# 將結果存入DataFrame，並用前值填補最後不足60筆的資料
df['WP_diff'] = wp_diff.fillna(0)  # 直接將NA值填補為0

# 第3個子圖：'P(W)' 欄位和 WP(Wh) 差值
ax1 = axs[2]  # 主要Y軸用於P(W)
ax1.plot(df['datetime'], df['P(W)'], label='P(W)', color='green')
ax1.set_ylabel('P(W)')
ax1.set_title('P(W) & Wh')

# 共用P(W)的Y軸來繪製WP(Wh)差值
ax1.plot(df['datetime'], df['WP_diff'], label='Wh', color='orange', alpha=0.7)
ax1.legend()

plt.xlabel('DateTime')
plt.tight_layout()
plt.show()
