import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
import os

# 設定中文字體和圖形樣式
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'DejaVu Sans', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False
matplotlib.rcParams['font.family'] = 'sans-serif'

# 嘗試使用 seaborn 樣式，如果失敗則使用預設樣式
try:
    plt.style.use('seaborn-v0_8-whitegrid')
except:
    try:
        plt.style.use('seaborn-whitegrid')
    except:
        # 如果 seaborn 不可用，使用 matplotlib 預設樣式
        pass

def select_csv_file():
    """開啟對話視窗,選取csv檔案"""
    root = tk.Tk()
    root.withdraw()  # 隱藏主視窗
    
    file_path = filedialog.askopenfilename(
        title="選擇CSV檔案",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    
    root.destroy()
    return file_path

def read_csv_data(file_path):
    """讀取csv檔案並處理數據"""
    try:
        # 讀取csv檔案
        df = pd.read_csv(file_path)
        
        # 讀取csv檔案的第1欄為Date,第2欄為Time,其餘欄位視為資料
        if len(df.columns) < 2:
            raise ValueError("CSV檔案至少需要包含Date和Time兩欄")
        
        # 將Date與Time合併為datetime
        df['DateTime'] = pd.to_datetime(df.iloc[:, 0].astype(str) + ' ' + df.iloc[:, 1].astype(str))
        
        # 獲取數據欄位（除了Date, Time和DateTime）
        data_columns = [col for col in df.columns[2:] if col != 'DateTime']
        
        return df, data_columns
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return None, None

def select_columns(data_columns):
    """讓使用者選擇要繪製的欄位"""
    print(f"\n可用的數據欄位:")
    for i, col in enumerate(data_columns, 1):
        print(f"{i}: {col}")
    
    print(f"\n請選擇要繪製的欄位:")
    print("格式說明:")
    print("- 單一欄位: 1")
    print("- 多個欄位: 1,3,5")
    print("- 範圍選擇: 1-3 (代表 1,2,3)")
    print("- 混合格式: 1,3-5,7")
    print("- 全部欄位: all 或 直接按 Enter")
    
    while True:
        user_input = input("請輸入選擇: ").strip()
        
        if not user_input or user_input.lower() == 'all':
            return data_columns
        
        try:
            selected_indices = parse_column_selection(user_input, len(data_columns))
            if selected_indices:
                selected_columns = [data_columns[i-1] for i in selected_indices]
                print(f"已選擇欄位: {', '.join(selected_columns)}")
                return selected_columns
            else:
                print("未選擇任何有效欄位，請重新輸入")
        except ValueError as e:
            print(f"輸入格式錯誤: {e}")
            print("請重新輸入")

def parse_column_selection(user_input, max_columns):
    """解析使用者輸入的欄位選擇"""
    selected_indices = set()
    
    # 分割逗號
    parts = [part.strip() for part in user_input.split(',')]
    
    for part in parts:
        if '-' in part:
            # 處理範圍格式 (如 1-3)
            range_parts = part.split('-')
            if len(range_parts) != 2:
                raise ValueError(f"範圍格式錯誤: {part}")
            
            try:
                start = int(range_parts[0].strip())
                end = int(range_parts[1].strip())
                
                if start < 1 or end < 1:
                    raise ValueError("欄位編號必須大於 0")
                if start > max_columns or end > max_columns:
                    raise ValueError(f"欄位編號不能超過 {max_columns}")
                if start > end:
                    raise ValueError("範圍起始值不能大於結束值")
                
                # 添加範圍內的所有索引
                for i in range(start, end + 1):
                    selected_indices.add(i)
                    
            except ValueError as e:
                raise ValueError(f"範圍格式錯誤: {part} - {e}")
        else:
            # 處理單一編號
            try:
                index = int(part)
                if index < 1 or index > max_columns:
                    raise ValueError(f"欄位編號必須在 1 到 {max_columns} 之間")
                selected_indices.add(index)
            except ValueError as e:
                raise ValueError(f"編號格式錯誤: {part} - {e}")
    
    return sorted(list(selected_indices))

def plot_line_chart(df, data_columns, file_name):
    """繪製折線圖"""
    # 確保中文字體設定
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'DejaVu Sans', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    # 設定背景顏色
    fig.patch.set_facecolor('#f8f9fa')
    ax.set_facecolor('#ffffff')
    
    # 定義顏色調色板
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f']
    
    for i, col in enumerate(data_columns):
        color = colors[i % len(colors)]
        ax.plot(df['DateTime'], df[col], label=col, linewidth=2.5, color=color, alpha=0.8)
    
    # 美化圖表
    ax.set_xlabel('時間', fontsize=12, fontweight='bold', color='#333333')
    ax.set_ylabel('數值', fontsize=12, fontweight='bold', color='#333333')
    ax.set_title(f'{file_name} - 數據折線圖', fontsize=16, fontweight='bold', color='#2c3e50', pad=20)
    
    # 美化圖例
    legend = ax.legend(loc='upper right', frameon=True, fancybox=True, shadow=True)
    legend.get_frame().set_facecolor('#f8f9fa')
    legend.get_frame().set_alpha(0.9)
    
    # 美化網格
    ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5, color='#cccccc')
    ax.set_axisbelow(True)
    
    # 美化座標軸
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cccccc')
    ax.spines['bottom'].set_color('#cccccc')
    
    plt.xticks(rotation=45, fontsize=10)
    plt.yticks(fontsize=10)
    plt.tight_layout()
    plt.show()

def plot_histogram(df, data_columns, file_name, center_line=None):
    """繪製常態分佈圖（直方圖）"""
    # 確保中文字體設定
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'DejaVu Sans', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
    
    fig, axes = plt.subplots(len(data_columns), 1, figsize=(12, 5*len(data_columns)))
    if len(data_columns) == 1:
        axes = [axes]
    
    # 設定整體背景
    fig.patch.set_facecolor('#f8f9fa')
    
    # 定義顏色調色板
    colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#e67e22']
    
    for i, col in enumerate(data_columns):
        # 設定子圖背景
        axes[i].set_facecolor('#ffffff')
        
        data = df[col].dropna()
        color = colors[i % len(colors)]
        
        # 繪製直方圖（以百分比顯示）
        n, bins, patches = axes[i].hist(data, bins=50, alpha=0.7, 
                                       weights=np.ones(len(data))/len(data)*100, 
                                       label=f'{col} 分佈', 
                                       color=color, edgecolor='white', linewidth=0.5)
        
        # 計算平均值和標準差
        mean_val = data.mean()
        std_val = data.std()
        
        # 繪製常態分佈曲線（調整為百分比）
        x = np.linspace(data.min(), data.max(), 100)
        normal_curve = (1/(std_val * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((x - mean_val) / std_val) ** 2)
        # 將密度轉換為百分比
        bin_width = (data.max() - data.min()) / 50
        normal_curve_percent = normal_curve * bin_width * 100
        axes[i].plot(x, normal_curve_percent, 'r-', linewidth=3, label='常態分佈曲線', alpha=0.8)
        
        # 繪製中心線
        if center_line is None:
            center_line = mean_val
        
        axes[i].axvline(center_line, color='#27ae60', linestyle='--', linewidth=3, 
                       label=f'中心線 ({center_line:.2f})', alpha=0.8)
        axes[i].axvline(mean_val, color='#f39c12', linestyle=':', linewidth=3, 
                       label=f'平均值 ({mean_val:.2f})', alpha=0.8)
        
        # 美化標籤和標題
        axes[i].set_xlabel('數值', fontsize=12, fontweight='bold', color='#333333')
        axes[i].set_ylabel('百分比 (%)', fontsize=12, fontweight='bold', color='#333333')
        axes[i].set_title(f'{file_name} - {col} 常態分佈圖', fontsize=14, fontweight='bold', color='#2c3e50', pad=15)
        
        # 美化圖例
        legend = axes[i].legend(loc='upper right', frameon=True, fancybox=True, shadow=True)
        legend.get_frame().set_facecolor('#f8f9fa')
        legend.get_frame().set_alpha(0.9)
        
        # 美化網格
        axes[i].grid(True, alpha=0.3, linestyle='-', linewidth=0.5, color='#cccccc')
        axes[i].set_axisbelow(True)
        
        # 美化座標軸
        axes[i].spines['top'].set_visible(False)
        axes[i].spines['right'].set_visible(False)
        axes[i].spines['left'].set_color('#cccccc')
        axes[i].spines['bottom'].set_color('#cccccc')
        
        # 美化刻度
        axes[i].tick_params(axis='both', which='major', labelsize=10, colors='#333333')
    
    plt.tight_layout()
    plt.show()

def main():
    """主程式"""
    print("=== 數據繪圖程式 ===")
    
    # 開啟對話視窗,選取csv檔案
    file_path = select_csv_file()
    
    if not file_path:
        print("未選擇檔案，程式結束")
        return
    
    print(f"已選擇檔案: {os.path.basename(file_path)}")
    
    # 取得檔案名稱（不含副檔名）
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # 讀取csv檔案
    df, data_columns = read_csv_data(file_path)
    
    if df is None:
        print("讀取檔案失敗，程式結束")
        return
    
    print(f"成功讀取數據，共 {len(df)} 筆記錄")
    print(f"數據欄位: {', '.join(data_columns)}")
    
    # 讓使用者選擇要繪製的欄位
    selected_columns = select_columns(data_columns)
    if not selected_columns:
        print("未選擇任何欄位，程式結束")
        return
    
    # 在終端視窗詢問要繪製的圖形型態
    while True:
        print("\n請選擇要繪製的圖形型態:")
        print("1: 折線圖")
        print("2: 常態分佈圖")
        
        choice = input("請輸入選擇 (1 或 2): ").strip()
        
        if choice == '1':
            # 如果選擇1:折線圖,則繪製折線圖
            print("正在繪製折線圖...")
            plot_line_chart(df, selected_columns, file_name)
            break
            
        elif choice == '2':
            # 如果選擇2:常態分佈圖,繼續詢問中心線的位置
            print("\n常態分佈圖選項:")
            print("1: 使用平均值作為中心線")
            print("2: 自訂中心線位置")
            
            center_choice = input("請選擇中心線設定 (1 或 2): ").strip()
            
            center_line = None
            if center_choice == '2':
                try:
                    center_line = float(input("請輸入中心線位置: "))
                except ValueError:
                    print("輸入無效，將使用平均值作為中心線")
                    center_line = None
            
            print("正在繪製常態分佈圖...")
            plot_histogram(df, selected_columns, file_name, center_line)
            break
            
        else:
            print("無效的選擇，請重新輸入")

if __name__ == "__main__":
    main()
