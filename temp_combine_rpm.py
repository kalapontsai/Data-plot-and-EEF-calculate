#新增python程式
# -*- coding: utf-8 -*-
#開啟溫度檔案,依照前兩欄Date與Time,將第三欄以後的欄位視為Temperature資料讀出
#再開啟轉速檔案,依照前兩欄Date與Time,將第三欄RPM資料讀出
#將兩個檔案的Date與Time欄位比對,若相同,則將Temperature與RPM資料合併輸出至新檔案
#輸出檔案欄位: Date, Time, Temperature, RPM
#輸出檔案名稱: 溫度檔案名稱+"RPM".csv

import pandas as pd
import numpy as np
import os
from tkinter import Tk, filedialog

# 隱藏主視窗
Tk().withdraw()

# 選取溫度檔案
temp_file = filedialog.askopenfilename(title="選取溫度檔案", filetypes=[("CSV files", "*.csv")])
if not temp_file:
    print("未選取溫度檔案")
    exit()

# 選取轉速檔案
rpm_file = filedialog.askopenfilename(title="選取轉速檔案", filetypes=[("CSV files", "*.csv")])
if not rpm_file:
    print("未選取轉速檔案")
    exit()

# 讀取溫度檔案
# 前兩欄為Date, Time，第三欄以後為Temperature
temp_df = pd.read_csv(temp_file)
# 統一 datetime 格式
temp_df['DateTime'] = temp_df['Date'].astype(str) + ' ' + temp_df['Time'].astype(str)
temp_df['DateTime'] = pd.to_datetime(temp_df['DateTime'], format='%Y/%m/%d %H:%M:%S', errors='coerce')
temp_df = temp_df.dropna(subset=['DateTime'])

# 讀取轉速檔案
# 前兩欄為Date, Time，第三欄為RPM
rpm_df = pd.read_csv(rpm_file)
# 統一 datetime 格式，支援多種分隔符
def parse_rpm_datetime(row):
    dt_str = str(row['date']) + ' ' + str(row['time'])
    for fmt in ('%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M:%S'):
        try:
            return pd.to_datetime(dt_str, format=fmt)
        except Exception:
            continue
    return pd.NaT
rpm_df['DateTime'] = [parse_rpm_datetime(row) for _, row in rpm_df.iterrows()]
rpm_df = rpm_df.dropna(subset=['DateTime'])

# 合併資料
rpm_df = rpm_df.sort_values('DateTime')
temp_df = temp_df.sort_values('DateTime')
# 以 temp_df 的每一筆 datetime，尋找 rpm_df "之前"最近的 datetime，取出 read 欄位
merged = pd.merge_asof(
    temp_df,
    rpm_df[['DateTime','read']],
    on='DateTime',
    direction='backward',
    allow_exact_matches=True
)
merged = merged.rename(columns={'read': 'RPM'})
merged['RPM'] = merged['RPM'] * 0.01

# 拆分Date與Time
merged['Date'] = merged['DateTime'].dt.date.astype(str)
merged['Time'] = merged['DateTime'].dt.time.astype(str)

# 輸出
output_cols = [col for col in temp_df.columns if col != 'DateTime'] + ['RPM']
output_file = os.path.splitext(os.path.basename(temp_file))[0] + '-RPM.csv'
output_dir = os.path.dirname(temp_file)
output_path = os.path.join(output_dir, output_file)
merged.to_csv(output_path, columns=output_cols, index=False, encoding='utf-8-sig')
print(f'已輸出: {output_path}')

