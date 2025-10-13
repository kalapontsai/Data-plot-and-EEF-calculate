# csv資料輸出圖表
# 1.開啟對話框,選取csv檔案
# 2.讀取csv檔案,並將資料存入list
# 3.將資料繪製成圖表
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import os

# 強制只用微軟正黑體，避免 fallback 到 DejaVu Sans
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
matplotlib.rcParams['axes.unicode_minus'] = False

def select_and_plot_csv():
    # 建立 Tkinter 主視窗（隱藏）
    root = tk.Tk()
    root.withdraw()

    # 開啟檔案選取對話框
    file_path = filedialog.askopenfilename(
        title="選取 CSV 檔案",
        filetypes=[("CSV files", "*.csv")]
    )
    if not file_path:
        print("未選取檔案")
        return

    # 讀取 CSV 檔案
    df = pd.read_csv(file_path)
    print("資料預覽：")
    print(df.head())

    # 將資料存入 list
    data_list = df.values.tolist()

    # 以第1、2欄合併為DateTime作為X軸
    if df.shape[1] >= 3:
        # 假設第1欄為日期，第2欄為時間，第3欄之後為Y值
        date_col = df.iloc[:, 0].astype(str)
        time_col = df.iloc[:, 1].astype(str)
        dt_str = date_col + ' ' + time_col
        x = pd.to_datetime(dt_str, errors='coerce')
        
        # 設定X軸範圍：第一筆資料時間 + 3小時
        start_time = x.iloc[0]  # 第一筆資料時間
        end_time = start_time + pd.Timedelta(hours=3)  # 加3小時
        
        # 排除不需要顯示的欄位
        exclude_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)']
        columns_to_plot = [col for col in df.columns[2:] if col not in exclude_columns]
        
        plt.figure(figsize=(10, 6))
        linewidth = 1.0  # 線寬參數，可自行調整
        for col in columns_to_plot:
            y = df[col]
            plt.plot(x, y, label=col, linewidth=linewidth)
            #plt.plot(x, y, marker='o', label=col, linewidth=linewidth)
        
        # 設定X軸範圍
        plt.xlim(start_time, end_time)
        
        # 設定X軸刻度為30分鐘間隔
        from matplotlib.dates import MinuteLocator, DateFormatter
        plt.gca().xaxis.set_major_locator(MinuteLocator(interval=30))
        plt.gca().xaxis.set_major_formatter(DateFormatter('%H:%M'))
        
        # 設定Y軸特定刻度線為水平線條 (10和-5)
        ax = plt.gca()
        # 取得X軸範圍
        x_min, x_max = ax.get_xlim()
        
        # 在Y=10處畫水平虛線
        ax.axhline(y=10, xmin=0, xmax=1, color='red', linewidth=2, linestyle='--', alpha=0.7)
        # 在Y=-5處畫水平虛線
        ax.axhline(y=-5, xmin=0, xmax=1, color='red', linewidth=2, linestyle='--', alpha=0.7)
        
        # 加粗對應的刻度標籤
        for tick in ax.yaxis.get_major_ticks():
            if tick.get_loc() == 10 or tick.get_loc() == -5:
                tick.label1.set_fontweight('bold')  # 加粗刻度標籤
                tick.label1.set_fontsize(12)  # 增大字體
        
        plt.xlabel('日期時間')
        plt.ylabel('溫度')

        csv_filename = os.path.basename(file_path)
        plt.title(f"{csv_filename} (X軸: DateTime, 範圍: 3小時, 刻度: 30分鐘)")
        plt.legend()
        plt.grid(True)
        plt.gcf().autofmt_xdate()
        plt.show()
    else:
        print("資料欄位不足，需至少三欄：日期、時間、數值")

if __name__ == "__main__":
    select_and_plot_csv()