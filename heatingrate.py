# 匯入所需套件
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import tkinter as tk
from tkinter import filedialog
import numpy as np

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

def plot_temperature_data():
    """繪製溫度資料的三個子圖"""
    # 1. 開啟檔案對話框選擇CSV檔案
    file_path = select_csv_file()
    
    if not file_path:
        print("未選擇檔案，程式結束")
        return
    
    # 讀取CSV檔案
    df = pd.read_csv(file_path)
    
    # 合併日期和時間欄位
    if 'Date' in df.columns and 'Time' in df.columns:
        df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
    else:
        # 如果沒有Date和Time欄位，使用前兩欄
        df['datetime'] = pd.to_datetime(df.iloc[:, 0] + ' ' + df.iloc[:, 1])
    
    # 設定X軸為日期時間
    x = df['datetime']
    
    # 建立三個子圖
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 15))
    #fig.suptitle('溫度與功率分析圖表', fontsize=16)
    
    # 第一個子圖：顯示除了"U(V)","I(A)","P(W)","WP(Wh)"欄位的溫度曲線
    excluded_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)', 'Date', 'Time', 'datetime']
    temperature_columns = [col for col in df.columns if col not in excluded_columns]
    
    for col in temperature_columns:
        if df[col].dtype in ['float64', 'int64']:  # 只繪製數值欄位
            ax1.plot(x, df[col], label=col, linewidth=1.5)
    
    ax1.set_title('溫度曲線圖')
    ax1.set_ylabel('溫度 (°C)')
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    ax1.set_xticklabels([])  # 隱藏X軸座標標籤
    
    # 第二個子圖：顯示欄名第一個字元為"*"的溫度，每5分鐘的溫度變化曲線
    star_columns = [col for col in df.columns if col.startswith('*')]
    
    if star_columns:
        # 計算每5分鐘的溫度變化
        for col in star_columns:
            if df[col].dtype in ['float64', 'int64']:
                # 計算溫度變化（當前值減去前一個值）
                temp_change = df[col].diff()
                ax2.plot(x, temp_change, label=f'{col} 溫度變化', linewidth=1.5)
        
        ax2.set_title('溫度變化速率（每5分鐘）')
        ax2.set_ylabel('溫度變化 (°C)')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        ax2.set_xticklabels([])  # 隱藏X軸座標標籤
    else:
        ax2.text(0.5, 0.5, '沒有找到以*開頭的欄位', 
                transform=ax2.transAxes, ha='center', va='center', fontsize=12)
        ax2.set_title('溫度變化速率')
        ax2.set_xticklabels([])  # 隱藏X軸座標標籤
    
    # 第三個子圖：顯示"P(W)"的功率數值
    if 'P(W)' in df.columns:
        ax3.plot(x, df['P(W)'], label='功率 P(W)', color='red', linewidth=2)
        ax3.set_title('功率數值圖')
        ax3.set_ylabel('功率 (W)')
        ax3.legend()
        ax3.grid(True, alpha=0.3)
    else:
        ax3.text(0.5, 0.5, '沒有找到P(W)欄位', 
                transform=ax3.transAxes, ha='center', va='center', fontsize=12)
        ax3.set_title('功率數值圖')
    
    # 設定所有子圖共用X軸
    ax3.set_xlabel('日期時間')
    
    # 調整子圖間距
    plt.tight_layout()
    
    # 顯示圖表
    plt.show()

if __name__ == "__main__":
    plot_temperature_data()
