#data log plot and EEF simple calculate for MS access 97 data base raw data
#mdb file need to export to excel 97 format (*.xls)
# fork from plot_gui_2_4_1.py
# pip install xlrd>=2.0.0.0 dedicated for office 97 *.xls, can not used for *.xlsx
#-------------------------------------------------------------------------------
#Rev.0.1 initial
#Rev.0.1.0.1 delete '日期' from xls and csv file once import to df
#Rev.0.2.0.0 26/2/9 更正原數據筆數間隔非10秒整,能耗計算需先計算1min平均,再累積WH值
#-------------------------------------------------------------------------------
# 
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator
from matplotlib import rcParams
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
import sys  # 引入 sys 模組以便退出程式
import os

def resource_path(relative_path):
    """ 獲取資源的絕對路徑 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(getattr(sys, '_MEIPASS'), relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

class EnergyCalculator:
    def __init__(self):
        pass
    def current_ef_thresholds(self,energy_allowance,fridge_type):
        if fridge_type == 5:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.72,1)
            threshold_lv2 = round(energy_allowance * 1.54,1)
            threshold_lv3 = round(energy_allowance * 1.36,1)
            threshold_lv4 = round(energy_allowance * 1.18,1)
        else:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.6,1)
            threshold_lv2 = round(energy_allowance * 1.45,1)
            threshold_lv3 = round(energy_allowance * 1.3,1)
            threshold_lv4 = round(energy_allowance * 1.15,1)
        return[ threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4 ]

    def future_ef_thresholds(self,energy_allowance,fridge_type):
        if fridge_type == 5:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.294,1)
            threshold_lv2 = round(energy_allowance * 1.221,1)
            threshold_lv3 = round(energy_allowance * 1.147,1)
            threshold_lv4 = round(energy_allowance * 1.074,1)
        else:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.308,1)
            threshold_lv2 = round(energy_allowance * 1.231,1)
            threshold_lv3 = round(energy_allowance * 1.154,1)
            threshold_lv4 = round(energy_allowance * 1.077,1)
        return[ threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4 ]


    def calculate(self, VF, VR, daily_consumption, fridge_temp, freezer_temp, fan_type):
        """
        計算冰箱能耗相關指標
        
        參數:
            VR: 冷藏室容積(L)
            VF: 冷凍室容積(L)
            daily_consumption: 日耗電量(kWh/日)
            fridge_temp: 冷藏室溫度(°C), 預設3.0
            freezer_temp: 冷凍室溫度(°C), 預設-18.0
        
        返回:
            包含所有計算結果的字典
        """
        results = {}
        # 1. 計算標準K值 (溫度係數)
        #Rev.2.4.1 K值分為標準K和實測K
        standard_K = 1.78
        measured_K = self.calculate_K_value(freezer_temp, fridge_temp)
        #print(f"標準K值: {standard_K}, 實測K值: {measured_K}")
        
        # 2. 計算等效內容積,有效內容積
        #Rev.2.4.1 等效內容積計算分為標準公式用標準K值,實測公式用實測K值
        equivalent_volume = self.calculate_equivalent_volume(VR, VF, standard_K)
        measured_equivalent_volume = self.calculate_equivalent_volume(VR, VF, measured_K)
        effective_volume = VR + VF
        
        #Rev.2.4 修正2027容許耗用能源基準計算公式 型式改用有效內容積來判定
        # 3. 確定冰箱型式
        fridge_type = self.determine_fridge_type(effective_volume, VR, VF, fan_type)
        #print(f"冰箱型式: {fridge_type}")
        
        # 4. 計算容許耗用能源基準 (每月)
        energy_allowance = self.calculate_energy_allowance(equivalent_volume, fridge_type)
        
        # 5. 計算2027容許耗用能源基準
        future_energy_allowance = self.calculate_future_energy_allowance(equivalent_volume, fridge_type)
        
        # 6. 計算耗電量基準 (每月)
        benchmark_consumption = self.calculate_benchmark_consumption(equivalent_volume, energy_allowance)
        
        # 7. 計算2027耗電量基準
        future_benchmark_consumption = self.calculate_future_benchmark_consumption(equivalent_volume, future_energy_allowance)
        
        # 8. 計算實測月耗電量
        monthly_consumption = int(daily_consumption * 30)
        
        # 9. 計算EF值 (能效因子)
        #Rev.2.4.1 實測用實測等效內容積
        ef_value = round(measured_equivalent_volume / monthly_consumption,1)
        
        # 9.1 計算現有效率基準百分比和等級
        current_ef_thresholds = self.current_ef_thresholds(energy_allowance, fridge_type)

        # 10. 計算現有效率基準百分比和等級
        current_percent, current_grade = self.calculate_current_efficiency(ef_value, current_ef_thresholds)
        
        # 10.1 計算2027新效率基準百分比和等級
        future_ef_thresholds = self.future_ef_thresholds(future_energy_allowance, fridge_type)

        # 11. 計算2027新效率基準百分比和等級
        future_percent, future_grade = self.calculate_future_efficiency(ef_value, future_ef_thresholds)
        
        # 整理所有結果
        results.update({
            '冷凍室溫度(°C)': freezer_temp,
            '冷藏室溫度(°C)': fridge_temp,
            '實測K值': measured_K,
            'VF(L)': int(VF),
            'VR(L)': int(VR),
            '等效內容積(L)': equivalent_volume,
            '實測等效內容積(L)': measured_equivalent_volume,
            '冰箱型式': fridge_type,
            '2018容許耗用能源基準(L/kWh/月)': energy_allowance,
            '2027容許耗用能源基準(L/kWh/月)': future_energy_allowance,
            '2018耗電量基準(kWh/月)': benchmark_consumption,
            '2027耗電量基準(kWh/月)': future_benchmark_consumption,
            '實測月耗電量(kWh/月)': monthly_consumption,
            '實測EF值': ef_value,
            '2018效率基準百分比(%)': current_percent,
            '2018效率等級': current_grade,
            '2027新效率基準百分比(%)': future_percent,
            '2027新效率等級': future_grade
        })
        
        return results
    
    def calculate_K_value(self, freezer_temp, fridge_temp):
        """計算K值 (溫度係數)"""
        # 根據公式 K = (30 - 冷凍庫溫度) / (30 - 冷藏庫溫度)
        print(f"冷凍庫溫度: {freezer_temp}, 冷藏庫溫度: {fridge_temp}")        
        return round((30 - freezer_temp) / (30 - fridge_temp), 2)
    
    def calculate_equivalent_volume(self, VR, VF, K):
        """計算等效內容積"""
        return int(VR + (K * VF))
    
    def determine_fridge_type(self, equivalent_volume, VR, VF, fan_type):
        """確定冰箱型式"""
        if VF == 0:  # 只有冷藏室
            return 5
        elif equivalent_volume < 400 and fan_type == 1:
            return 1  # 假設是風冷式(實際應根據具體設計)
        elif equivalent_volume >= 400 and fan_type == 1:
            return 2
        elif equivalent_volume < 400 and fan_type == 0:
            return 3
        else:
            return 4  # 假設是風冷式(實際應根據具體設計)
    
    def calculate_energy_allowance(self, equivalent_volume, fridge_type):
        """計算容許耗用能源基準"""
        # 根據公式，ROUND(IFS(fridge_type=1,equivalent_volume/(0.037*equivalent_volume+24.3),fridge_type=2,equivalent_volume/(0.031*M4+21),fridge_type=3,equivalent_volume/(0.033*equivalent_volume+19.7),fridge_type=4,equivalent_volume/(0.029*equivalent_volume+17),fridge_type=5,equivalent_volume/(0.033*equivalent_volume+15.8)),1)
        if fridge_type == 1:
            return round( equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round( equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round( equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round( equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round( equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)
    
    def calculate_future_energy_allowance(self, equivalent_volume, fridge_type):
        """計算2027年容許耗用能源基準"""
        # 公式:=ROUND(IFS(F4=1,1.3*M4/(0.037*M4+24.3),F4=2,1.3*M4/(0.031*M4+21),F4=3,1.3*M4/(0.033*M4+19.7),F4=4,1.3*M4/(0.029*M4+17),F4=5,1.36*M4/(0.033*M4+15.8)),1)
        # F4 = fridge_type, M4 = equivalent_volume
        if fridge_type == 1:
            return round( 1.3 * equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round( 1.3 * equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round(1.3 * equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round(1.3 * equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round(1.36 * equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)

    def calculate_benchmark_consumption(self, equivalent_volume, energy_allowance):
        """計算耗電量基準"""
        return int(equivalent_volume / energy_allowance)
    
    def calculate_future_benchmark_consumption(self, equivalent_volume, future_energy_allowance):
        """計算2027耗電量基準"""
        return int(equivalent_volume / future_energy_allowance)
    
    def calculate_current_efficiency(self, ef_value, thresholds):
        # 確定等級
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        else :
            grade = "5級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        
        return final_percent, grade
    
    def calculate_future_efficiency(self, ef_value, thresholds):
        # 確定等級
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        else :
            grade = "5級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        
        return final_percent, grade


# 在程式中的適當位置（例如在plot_chart函數之前）添加DraggableLine類
class DraggableLine:
    def __init__(self, ax, xdata, ydata, initial_pos, color='red', linestyle='--', linewidth=1, 
                 date_var=None, time_var=None):
        self.ax = ax
        self.xdata = xdata
        self.ydata = ydata
        self.line = ax.axvline(x=initial_pos, color=color, linestyle=linestyle, linewidth=linewidth)
        self.press = None
        self.date_var = date_var
        self.time_var = time_var
        self.cid_press = self.line.figure.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = self.line.figure.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = self.line.figure.canvas.mpl_connect('motion_notify_event', self.on_motion)
        
    def on_press(self, event):
        if event.inaxes != self.ax:
            return
        contains, attrd = self.line.contains(event)
        if not contains:
            return
        self.press = True
        
    def on_motion(self, event):
        if not self.press or event.inaxes != self.ax:
            return
        x_pos = event.xdata
        self.line.set_xdata([x_pos, x_pos])
        self.update_text_boxes(x_pos)
        self.line.figure.canvas.draw()
        
    def on_release(self, event):
        self.press = False
        self.line.figure.canvas.draw()
        
    def get_position(self):
        return self.line.get_xdata()[0]
        
    def update_text_boxes(self, x_pos):
        if self.date_var is not None and self.time_var is not None:
            dt = mdates.num2date(x_pos)
            self.date_var.set(dt.strftime('%Y-%m-%d'))
            self.time_var.set(dt.strftime('%H:%M:%S'))

def plot_chart():
    ax1_color = "lightcyan"
    ax2_color = "lightyellow"
    try:
        global ax1, ax2, ax2a, ax2b, ax2c
        
        # 清除現有繪圖
        fig.clf()

        # 重新建立 gridspec 與子圖，保留 7:3 高度比例
        gs = fig.add_gridspec(2, 1, height_ratios=[7, 3])
        ax1 = fig.add_subplot(gs[0, 0], facecolor=ax1_color)
        ax2 = fig.add_subplot(gs[1, 0], sharex=ax1, facecolor=ax2_color)
        
        # 清除右側的雙 Y 軸
        for ax in fig.axes:
            if ax not in [ax1, ax2]:
                ax.remove()
        
        # 重新創建右側 Y 軸
        ax2a = ax2  # 左側 Y 軸 (功率 P(W))
        ax2b = ax2.twinx()  # 右側第一個 Y 軸 (電壓 U(V))
        ax2c = ax2.twinx()  # 右側第二個 Y 軸 (電流 I(A))
        
        # 調整右側第二個 Y 軸的位置
        ax2c.spines['right'].set_position(('outward', 35))

        # 使用 select_file() 處理後的 df
        global global_df
        if global_df is None or global_df.empty:
            messagebox.showerror("錯誤", "請先選擇檔案！")
            return
        
        df = global_df.copy()

        # 以倒數3列設為電力欄位,分別為U(V), I(A), P(W)
        #power_columns = list(df.columns[-3:])
        power_columns = ['V', 'I', 'W']
        df_power = df[['datetime'] + power_columns]
        
        # 其餘欄位視為溫度資料
        # temp_columns 設定為第二列到倒數第三列(不包括倒數第三列)，並排除 'datetime'
        temp_columns = [c for c in df.columns[1:-3]]
        df_temp = df[['datetime'] + temp_columns]
        #print(df_temp)
        # 繪製圖表
        fig.suptitle(chart_title.get())
        
        # 繪製溫度資料
        for col in temp_columns:
            ax1.plot(df_temp['datetime'], df_temp[col], label=col)
        ax1.set_ylabel("溫度")

        ax1.legend(temp_columns, loc='upper left')
        ax1.set_facecolor(ax1_color)
        ax1.grid(True, linestyle='dashdot')
        ax1.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
        
        # 繪製電力資料
        # 使用 power_columns 的實際列名，順序為：倒數第3列=U(V), 倒數第2列=I(A), 倒數第1列=P(W)
        if len(power_columns) >= 3:
            u_col = power_columns[0]  # 倒數第3列
            i_col = power_columns[1]  # 倒數第2列
            p_col = power_columns[2]  # 倒數第1列
            
            ax2a.plot(df_power['datetime'], df_power[p_col], label=p_col, color='green')
            ax2a.set_ylabel("功率 (P)", color='green')
            ax2a.tick_params(axis='y', labelcolor='green')
            ax2a.grid(True, linestyle='dashdot')

            ax2b.plot(df_power['datetime'], df_power[u_col], label=u_col, color='blue')
            ax2b.set_ylabel("電壓 (U)", color='blue')
            ax2b.tick_params(axis='y', labelcolor='blue')
            ax2b.set_ylim(0, 120)

            ax2c.plot(df_power['datetime'], df_power[i_col], label=i_col, color='red')
            ax2c.set_ylabel("電流 (I)", color='red')
            ax2c.tick_params(axis='y', labelcolor='red')
        else:
            messagebox.showerror("錯誤", "電力欄位不足！")
            return

        ax2.set_facecolor(ax2_color)
        ax2.xaxis.set_major_locator(MaxNLocator(10))
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45)
        plt.tight_layout()
        # 設定 ax2 的 x 軸 datetime 顯示格式為 "d-hh:mm"
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%d-%H:%M'))
        ax2.set_xlim(ax1.get_xlim())

        # 添加可拖動的光棒
        global start_line, end_line
        # start/end 若是空白或格式不對，回退到資料首末筆時間
        start_str = f"{start_date.get()} {start_time.get()}".strip()
        end_str = f"{end_date.get()} {end_time.get()}".strip()
        start_pos = pd.to_datetime(start_str, errors='coerce')
        end_pos = pd.to_datetime(end_str, errors='coerce')
        if pd.isna(start_pos):
            start_pos = df['datetime'].iloc[0]
        if pd.isna(end_pos):
            end_pos = df['datetime'].iloc[-1]
        
        start_line = DraggableLine(ax1, df_temp['datetime'], df_temp[temp_columns[0]], 
                                 start_pos, color='blue', linestyle='--',
                                 date_var=start_date, time_var=start_time)
        end_line = DraggableLine(ax1, df_temp['datetime'], df_temp[temp_columns[0]], 
                               end_pos, color='red', linestyle='--',
                               date_var=end_date, time_var=end_time)
        
        canvas.draw()
        toolbar.update()

    except Exception as e:
        messagebox.showerror("錯誤", f"繪製圖表時發生錯誤：{e}")
        print(f"plot_chart: err：{e}")

def calculate_statistics():
    try:
        # 使用 select_file() 處理後的 df
        global global_df
        if global_df is None or global_df.empty:
            messagebox.showerror("錯誤", "請先選擇檔案！")
            return
        
        df = global_df.copy()
        
        # 以倒數3列設為電力欄位,分別為U(V), I(A), P(W)
        #power_columns = df.columns[-4:-1]
        # 直接指定欄位'21','22','23'為電力欄位
        power_columns = ['V', 'I', 'W']
        if len(power_columns) < 3:
            messagebox.showerror("錯誤", "電力欄位不足！")
            return
        #print(power_columns)
        # 將 start_date, start_time, end_date, end_time 轉換為 datetime 格式
        start = pd.to_datetime(f"{start_date.get()} {start_time.get()}")
        end = pd.to_datetime(f"{end_date.get()} {end_time.get()}")
        # 計算 start 和 end 之間的分鐘數
        minutes_difference = int((end - start).total_seconds() / 60)
        
        # 過濾指定的日期時間範圍，並創建副本
        filtered_df = df[(df['datetime'] >= start) & (df['datetime'] <= end)].copy()

        if filtered_df.empty:
            messagebox.showinfo("結果", "指定範圍內沒有資料！")
            return
        
        # 計算平均值
        averages = filtered_df.mean(numeric_only=True)
        # 計算最大值
        max_values = filtered_df.max(numeric_only=True)
        # 計算最小值
        min_values = filtered_df.min(numeric_only=True)

        # 計算電力啟停周期
        # 使用倒數第1列作為功率欄位 P(W)
        power_column = power_columns[2]  # 倒數第1列
        power_throttle = float(on_off_throttle_entry_var.get())
        if power_column in filtered_df.columns:
            # 確保正確建立 power_on 欄位
            filtered_df.loc[:, 'power_on'] = filtered_df[power_column] >= power_throttle

            # 計算啟停周期次數
            power_cycles = int(filtered_df['power_on'].astype(int).diff().fillna(0).abs().sum() // 2)
            #print(f'power_cycles={power_cycles}')

            # 計算大於等於power_throttle和小於power_throttle的週期數，排除頭尾兩個周期
            mask = filtered_df['power_on']
            groups = (mask != mask.shift()).cumsum()
            segments = pd.DataFrame({
                '狀態': mask,
                '區段編號': groups,
                '時間': filtered_df['datetime']
            }).groupby(['區段編號', '狀態']).agg({'時間': ['min', 'max']}).reset_index()

            # 排除頭尾兩個周期
            if len(segments) > 2:
                segments = segments.iloc[1:-1]

            # 計算每個區段的持續時間
            segments['持續時間'] = (segments[('時間', 'max')] - segments[('時間', 'min')]).dt.total_seconds()

            # 分別計算大於等於power_throttle和小於power_throttle的週期數與平均時間
            above_segments = segments[segments['狀態']]
            below_segments = segments[~segments['狀態']]

            above_count = len(above_segments)
            below_count = len(below_segments)

            above_avg_time = int((above_segments['持續時間'].mean() / 60) if above_count > 0 else 0)
            below_avg_time = int(((below_segments['持續時間'].mean() / 60) + 1) if below_count > 0 else 0) #Rev.2.3.1 修正On/Off低標計算加1分鐘

            # 計算百分比
            if above_avg_time + below_avg_time > 0:
                above_percentage = int((above_avg_time / (above_avg_time + below_avg_time)) * 100)
            else:
                above_percentage = 0
        else:
            power_cycles = "無法計算，缺少 P(W) 欄位"
            
        # 以power_columns最後一列計算 WP(Wh)
        # 將start_pos到end_pos的值相加,除以60得到分鐘數
        total_seconds = int((filtered_df['datetime'].iloc[-1] - filtered_df['datetime'].iloc[0]).total_seconds())
        filter_WH = int(filtered_df['W'].sum() / 60)

        print(f"統計範圍：{start} ~ {end}")

        #print(f"filter_WH: {filter_WH}, total_seconds: {total_seconds}")

        if (total_seconds > 0):
            wp_24h_difference = int((filter_WH / total_seconds) * (24 * 3600))
        else:
            wp_24h_difference = "無法計算，時間範圍不足"
                    
        # 計算能耗
        fan_type = fan_type_var.get()  # 取得風扇類型的狀態
        vf = float(vf_entry_var.get()) if vf_entry_var.get().isdigit() else 0
        vr = float(vr_entry_var.get()) if vr_entry_var.get().isdigit() else 0
        fridge_temp = float(temp_r_entry_var.get()) if temp_r_entry_var.get().replace('.', '', 1).isdigit() else 3  # 冷藏室溫度
        freezer_temp = float(temp_f_entry_var.get()) if temp_f_entry_var.get().replace('.', '', 1).lstrip('-').isdigit() else -18.0  # 冷凍室溫度
        energy_calculator = EnergyCalculator()
        if isinstance(wp_24h_difference, (int, float)) and vf > 0 and vr > 0:
            daily_consumption = round(wp_24h_difference / 1000,3)  # 將 Wh 轉換為 kWh, 取小數3位
            # 計算
            results = energy_calculator.calculate(vf, vr, daily_consumption, fridge_temp, freezer_temp, fan_type)
            # 提取結果
            if results:
                # 打印結果
                print("冰箱能耗計算結果:")
                for key, value in results.items():
                    print(f"{key}: {value}")
        else:
            results = None
            print("無耗電量數據,無法計算能耗")

        # 顯示結果
        result_textbox.delete(1.0, tk.END)  # 清空文字框
        result_textbox.insert(tk.END, f"統計範圍：{start} ~ {end}\n")
        result_textbox.insert(tk.END, "平均/最大/最小：\n")
        # 整合平均值、最大值、最小值為一行輸出
        for column in averages.index:
            avg = averages[column]
            max_v = max_values[column]
            min_v = min_values[column]
            result_textbox.insert(tk.END, f"{column}: {avg:.1f} / {max_v:.1f} / {min_v:.1f}\n")
        result_textbox.insert(tk.END, f"\nON / Off 低標：{power_throttle}\n")
        result_textbox.insert(tk.END, f"ON / Off 周期次數：{power_cycles}\n")
        result_textbox.insert(tk.END, f"On 的平均時間: {above_avg_time} 分\n" if above_count > 0 else f"P(W) >= {power_throttle} 的平均時間: 無資料\n")
        result_textbox.insert(tk.END, f"Off 的平均時間: {below_avg_time} 分\n" if below_count > 0 else f"P(W) < {power_throttle} 的平均時間: 無資料\n")
        result_textbox.insert(tk.END, f"On / Off 百分比: {above_percentage}%\n")
        result_textbox.insert(tk.END, f"\n電力消耗：{filter_WH} w / {minutes_difference} 分\n")
        result_textbox.insert(tk.END, f"24 小時電力消耗：{wp_24h_difference} w\n")
        result_textbox.insert(tk.END, f"\n能耗計算：\n")
        if results:
            for key, value in results.items():
                result_textbox.insert(tk.END, f"{key}: {value}\n")
        else:
            result_textbox.insert(tk.END, "無法計算能耗，請檢查數據\n")
    except Exception as e:
        messagebox.showerror("錯誤", f"計算平均值或電力啟停周期時發生錯誤：{e}")
        print(f"計算平均值或電力啟停周期時發生錯誤：{e}")

# 新增儲存結果的函數
def save_results():
    try:
        # 選擇儲存檔案的路徑
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return  # 如果使用者取消操作，直接返回

        # 取得多行文字框的內容
        content = result_textbox.get(1.0, tk.END).strip()

        # 將內容寫入檔案
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)

        messagebox.showinfo("成功", "結果已成功儲存！")
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存結果時發生錯誤：{e}")


# 選擇檔案的函數
def select_file():
    global global_df
    # choose input format (xls or csv)
    # csv 是以本版儲存的格式, 非'GX20_PW3335.py'儲存的格式
    # xls 是双信冰箱軟體access 97 mdb 的資料表匯出excel 97格式
    # 使用 .get() 方法獲取 tk.StringVar 的值
    format_value = input_format_var.get()
    print(f'select format: {format_value}')
    if format_value == 'xls':
        file_path = filedialog.askopenfilename(filetypes=[("Office 97 files", "*.xls")])
        if file_path:
            csv_path.set(file_path)
            try:
                # 讀取 Office 97 檔案
                df = pd.read_excel(file_path)
                
                # 刪除第一行
                #df = df.iloc[1:].reset_index(drop=True)
                
                # 更改第27-29列欄位名稱為V,I,P
                #df.columns[26:29] = ['V', 'I', 'P']
                # 選取欄位：第2列（索引1）和第7-29列（索引6-28），其餘刪除
                # 列索引從0開始，所以第2列=索引1，第7-29列=索引6-28
                selected_columns = [1] + list(range(6, 29))  # 第2列和第7-29列
                # 確保索引不超出範圍
                available_columns = [col for col in selected_columns if col < len(df.columns)]
                df = df.iloc[:, available_columns]
                # 刪除所有值都是 NaN 的欄位
                df = df.dropna(axis=1, how='all')
                
                # 檢查 df 是否有值
                if df.empty:
                    messagebox.showerror("錯誤", "讀取的資料為空！")
                    global_df = None
                    return
                
                # 將列[0]轉換為 datetime
                if len(df.columns) > 0:
                    first_col = df.columns[0]
                    
                    # 如果第一列是字符串類型，先清理空格和空字符串
                    if df[first_col].dtype == 'object':
                        # 去除首尾空格
                        df[first_col] = df[first_col].astype(str).str.strip()
                        # 將空字符串和 'nan' 字符串轉換為 NaN
                        df[first_col] = df[first_col].replace(['', 'nan', 'NaN', 'None'], pd.NA)
                    
                    # 檢查刪除空值後是否還有資料
                    if df.empty:
                        messagebox.showerror("錯誤", "第一列沒有有效的日期時間資料！")
                        global_df = None
                        return
                    
                    # 使用 errors='coerce' 將無效值轉換為 NaT，然後刪除這些行
                    try:
                        df['datetime'] = pd.to_datetime(df[first_col], errors='coerce')
                    except Exception as e:
                        messagebox.showerror("錯誤", f"日期時間轉換錯誤：{e}")
                        global_df = None
                        return
                    
                    # 刪除 datetime 為 NaT 的行
                    df = df.dropna(subset=['datetime'])
                    # 按 datetime 欄位排序資料
                    df = df.sort_values(by='datetime')
                    # 重置索引
                    df = df.reset_index(drop=True)
                    # 檢查轉換後是否還有資料
                    if df.empty:
                        messagebox.showerror("錯誤", "無法解析日期時間欄位，資料為空！")
                        global_df = None
                        return
                else:
                    messagebox.showerror("錯誤", "資料格式錯誤，缺少時間欄位！")
                    global_df = None
                    return
                
                if '日期' in df.columns and 'datetime' in df.columns:
                    df = df.drop('日期',axis=1)

                #print(df.columns)
                df.rename(
                            columns={
                                '21': 'V',
                                '22': 'I',
                                '23': 'W'
                            },
                            inplace=True
                        )
                # 26/2/9 修正為1分鐘平均值
                # 設為索引，以便使用resample()
                df = df.set_index('datetime')
                df = df.resample('1min').mean(numeric_only=True)
                # 所有數值欄位取小數 1 位
                df = df.round(1)
                # 將索引還原為欄位，並把 datetime 放到最後一欄
                df = df.reset_index()



                #匯出csv檔案到讀取檔案的資料夾
                output_dir = os.path.dirname(file_path)
                # 獲取源檔案名稱（不含擴展名）並改為 .csv
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_file = os.path.join(output_dir, f'{base_name}-1min平均值.csv')
                df.to_csv(output_file, index=False, encoding='utf-8-sig')
                messagebox.showinfo("成功", f"CSV 檔案已成功儲存至：\n{output_file}")

            except Exception as e:
                messagebox.showerror("錯誤", f"處理 Office 97 檔案時發生錯誤：{e}")
                print(f"Select_file: err：{e}")
    else:
        try:
            file_path = filedialog.askopenfilename(filetypes=[("csv data", "*.csv")])
            csv_path.set(file_path)
            df = pd.read_csv(file_path)
            
            #刪除舊版保留的'日期'欄位
            if '日期' in df.columns:
                df = df.drop('日期',axis=1)
            
            if 'datetime' not in df.columns:
                messagebox.showerror("錯誤", f"沒有datetime欄位")
                return
            df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
        except Exception as e:
            messagebox.showerror("錯誤", f"日期時間轉換錯誤：{e}")
            global_df = None
            return

    print(df.head(5))
    # 將處理後的 df 保存到全局變數
    global_df = df.copy()
    # 同步更新 GUI 的起訖時間（避免 plot_chart() 解析空白字串失敗）
    try:
        first_dt = df['datetime'].iloc[0]
        last_dt = df['datetime'].iloc[-1]
        start_date.set(first_dt.strftime('%Y-%m-%d'))
        start_time.set(first_dt.strftime('%H:%M:%S'))
        end_date.set(last_dt.strftime('%Y-%m-%d'))
        end_time.set(last_dt.strftime('%H:%M:%S'))
    except Exception:
        pass

# 定義視窗關閉事件處理函數
def on_closing():
    try:
        result = messagebox.askokcancel("退出", "確定要退出程式嗎？")
        if result:
            root.quit()  # 停止主循環
            root.destroy()  # 關閉 Tkinter 視窗
        # 如果用戶點擊取消，不做任何事（視窗保持開啟）
    except Exception as e:
        # 如果出現錯誤，嘗試直接退出
        try:
            root.quit()
            root.destroy()
        except:
            pass


# 設定 matplotlib 使用的字體
rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 使用微軟正黑體
rcParams['axes.unicode_minus'] = False  # 解決負號無法顯示的問題
rcParams['font.size'] = 8
rcParams['axes.titlesize'] = 8      # 座標軸標題字體大小
rcParams['axes.labelsize'] = 8      # 座標軸標籤字體大小
rcParams['xtick.labelsize'] = 8     # X軸刻度字體大小
rcParams['ytick.labelsize'] = 8     # Y軸刻度字體大小
rcParams['legend.fontsize'] = 8     # 圖例字體大小
rcParams['figure.titlesize'] = 12    # 圖形標題字體大小

# 初始化主視窗
root = tk.Tk()
root.title("Plotter with Statistics for access 97 0.2 (1min回歸平均值計算)")
try:
    root.iconbitmap(resource_path('favicon.ico'))
except Exception:
    pass  # 如果圖標文件不存在，忽略錯誤

# 全域變數
csv_path = tk.StringVar()
chart_title = tk.StringVar()
start_date = tk.StringVar()
start_time = tk.StringVar()
end_date = tk.StringVar()
end_time = tk.StringVar()
input_format_var = tk.StringVar(value="xls")
global_df = None  # 儲存 select_file() 處理後的 DataFrame

# 選單元件
# Create a menu bar
menubar = tk.Menu(root)
root.config(menu=menubar)

# Create a menu
options_menu = tk.Menu(menubar, tearoff=0)

# Variable to store the selected option
#input_format_var = tk.StringVar(value="xls")  # default selection

# Add radiobutton menu items
options_menu.add_radiobutton(
    label="xls",
    variable=input_format_var,
    value="xls",
    #command=on_option_selected
)
options_menu.add_radiobutton(
    label="csv",
    variable=input_format_var,
    value="csv",
    #command=on_option_selected
)
# Add the menu to the menu bar
menubar.add_cascade(label="檔案格式", menu=options_menu)

# GUI 元件
main_frame = tk.Frame(root)  # 使用 Frame 包含文字框
main_frame.grid(row=0, column=0, padx=5, pady=5)
tk.Label(main_frame, text="檔案路徑:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Entry(main_frame, textvariable=csv_path, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(main_frame, text="選擇檔案", command=select_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(main_frame, text="圖表標題:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Entry(main_frame, textvariable=chart_title, width=50).grid(row=1, column=1, padx=5, pady=5)

tk.Button(main_frame, text="繪製圖表", command=lambda: plot_chart()).grid(row=4, column=0, pady=10)
tk.Button(main_frame, text="計算平均值", command=calculate_statistics).grid(row=4, column=1, pady=10)
tk.Button(main_frame, text="儲存結果", command=save_results).grid(row=4, column=2, pady=10)

# 多行文字框顯示結果
result_frame = tk.Frame(root)  # 使用 Frame 包含文字框
result_frame.grid(row=1, column=0, padx=5, pady=5)
# 
# 創建多行文字框
result_textbox = tk.Text(
    result_frame, 
    width=50, 
    height=50, 
    wrap=tk.WORD  # 自動換行
)
result_textbox.grid(row=0, rowspan=13, column=0, padx=5, pady=5)

# 能耗計算用欄位
tk.Label(result_frame, text="能耗計算用:").grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")
tk.Label(result_frame, text="冷凍室容積(L):").grid(row=1, column=1, padx=5, pady=5, sticky="w")
vf_entry_var = tk.StringVar(value="150")
vf_entry = tk.Entry(result_frame, width=5, textvariable=vf_entry_var)
vf_entry.grid(row=1, column=2, padx=5, pady=5, sticky="w")
tk.Label(result_frame, text="冷藏室容積(L):").grid(row=2, column=1, padx=5, pady=5, sticky="w")
vr_entry_var = tk.StringVar(value="350")
vr_entry = tk.Entry(result_frame, width=5, textvariable=vr_entry_var)
vr_entry.grid(row=2, column=2, padx=5, pady=5, sticky="w")
tk.Label(result_frame, text="F:").grid(row=3, column=1, padx=5, pady=5, sticky="w")
temp_f_entry_var = tk.StringVar(value="-18.0")
temp_f_entry = tk.Entry(result_frame, width=5, textvariable=temp_f_entry_var)
temp_f_entry.grid(row=3, column=2, padx=5, pady=5, sticky="w")

tk.Label(result_frame, text="R:").grid(row=4, column=1, padx=5, pady=5, sticky="w")
temp_r_entry_var = tk.StringVar(value="3.0")
temp_r_entry = tk.Entry(result_frame, width=5, textvariable=temp_r_entry_var)
temp_r_entry.grid(row=4, column=2, padx=5, pady=5, sticky="w")
fan_type_var = tk.IntVar(value=1)  # 0: unchecked, 1: checked
fan_type_checkbox = tk.Checkbutton(result_frame, text="風扇式", variable=fan_type_var)
fan_type_checkbox.grid(row=5, column=1, padx=5, pady=5, sticky="w")

tk.Label(result_frame, text="OnOfthrottle").grid(row=6, column=1, padx=5, pady=5, sticky="w")
on_off_throttle_entry_var = tk.StringVar(value="3.0")
on_off_throttle_entry = tk.Entry(result_frame, width=5, textvariable=on_off_throttle_entry_var)
on_off_throttle_entry.grid(row=6, column=2, padx=5, pady=5, sticky="w")

# 分割線
ttk.Separator(result_frame, orient="horizontal").grid(row=7, column=1, sticky="ew", pady=10)
# 開始日期與時間
tk.Label(result_frame, text="資料日期:").grid(row=8, column=1, columnspan=2, padx=5, pady=5, sticky="nsew")
def increment_date_time(var, increment, unit):
    try:
        current_value = pd.to_datetime(var.get())
        if unit == "day":
            new_value = current_value + pd.Timedelta(days=increment)
        elif unit == "hour":
            new_value = current_value + pd.Timedelta(hours=increment)
        var.set(new_value.strftime('%Y-%m-%d' if unit == "day" else '%H:%M'))
    except Exception:
        messagebox.showerror("錯誤", "無效的日期或時間格式！")

def bind_increment(widget, var, unit):
    def on_key(event):
        if event.state & 0x4:  # 檢查是否按下 CTRL 鍵
            if event.keysym == "Up":
                increment_date_time(var, 1, unit)
            elif event.keysym == "Down":
                increment_date_time(var, -1, unit)
    widget.bind("<KeyPress-Up>", on_key)
    widget.bind("<KeyPress-Down>", on_key)


start_date_entry = tk.Entry(result_frame, textvariable=start_date, width=10)
start_date_entry.grid(row=9, column=1, padx=5, pady=5, sticky="w")
bind_increment(start_date_entry, start_date, "day")

start_time_entry = tk.Entry(result_frame, textvariable=start_time, width=10)
start_time_entry.grid(row=10, column=1, padx=5, pady=5, sticky="w")
bind_increment(start_time_entry, start_time, "hour")

end_date_entry = tk.Entry(result_frame, textvariable=end_date, width=10)
end_date_entry.grid(row=11, column=1, padx=5, pady=5, sticky="w")
bind_increment(end_date_entry, end_date, "day")

end_time_entry = tk.Entry(result_frame, textvariable=end_time, width=10)
end_time_entry.grid(row=12, column=1, padx=5, pady=5, sticky="w")
bind_increment(end_time_entry, end_time, "hour")

fig, (ax1, ax2) = plt.subplots(2, 1, sharex=True, figsize=(12, 8), facecolor='lightgray')

canvas = FigureCanvasTkAgg(fig, master=root)
canvas_widget = canvas.get_tk_widget()
canvas_widget.grid(row=0, rowspan=2, column=1, padx=5, pady=5)

# 添加 Matplotlib 的工具列
toolbar_frame = tk.Frame(root)
toolbar_frame.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")
toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
toolbar.update()

# 綁定視窗關閉事件
root.protocol("WM_DELETE_WINDOW", on_closing)

# 啟動主迴圈
root.mainloop()