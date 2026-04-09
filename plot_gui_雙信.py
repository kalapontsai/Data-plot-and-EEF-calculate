#data log plot and EEF simple calculate for MS access 97 data base raw data
#mdb file need to export to excel 97 format (*.xls)
# fork from plot_gui_2_4_1.py
# pip install xlrd>=2.0.0.0 dedicated for office 97 *.xls, can not used for *.xlsx
#------------------------------------------------------------------------------
#Rev.0.1 initial
#Rev.0.1.0.1 delete '日期' from xls and csv file once import to df
#Rev.0.2.0.0 26/2/9 更正原數據筆數間隔非10秒整,能耗計算需先計算1min平均,再累積WH值;以tab20色盤作為溫度線避免重複
#Rev.0.2.0.1 26/4/8 調整能耗計算小數位,以及因為python 四捨五入的誤差
#Rev.0.3.0.0 26/4/8 UI響應式重構，支援螢幕尺寸自適應
#------------------------------------------------------------------------------
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
import sys
import os

def resource_path(relative_path):
    """ 獲取資源的絕對路徑 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(getattr(sys, '_MEIPASS'), relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)


# ============================================================================
# EnergyCalculator - 不動
# ============================================================================
class EnergyCalculator:
    def __init__(self):
        pass

    def current_ef_thresholds(self, energy_allowance, fridge_type):
        if fridge_type == 5:
            threshold_lv1 = round(energy_allowance * 1.72, 1)
            threshold_lv2 = round(energy_allowance * 1.54, 1)
            threshold_lv3 = round(energy_allowance * 1.36, 1)
            threshold_lv4 = round(energy_allowance * 1.18, 1)
        else:
            threshold_lv1 = round(energy_allowance * 1.6, 1)
            threshold_lv2 = round(energy_allowance * 1.45, 1)
            threshold_lv3 = round(energy_allowance * 1.3, 1)
            threshold_lv4 = round(energy_allowance * 1.15, 1)
        return [threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4]

    def future_ef_thresholds(self, energy_allowance, fridge_type):
        if fridge_type == 5:
            threshold_lv1 = round(energy_allowance * 1.294, 1)
            threshold_lv2 = round(energy_allowance * 1.221, 1)
            threshold_lv3 = round(energy_allowance * 1.147, 1)
            threshold_lv4 = round(energy_allowance * 1.074, 1)
        else:
            threshold_lv1 = round(energy_allowance * 1.308, 1)
            threshold_lv2 = round(energy_allowance * 1.231, 1)
            threshold_lv3 = round(energy_allowance * 1.154, 1)
            threshold_lv4 = round(energy_allowance * 1.077, 1)
        return [threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4]

    def calculate(self, VF, VR, daily_consumption, fridge_temp, freezer_temp, fan_type):
        results = {}
        standard_K = 1.78
        measured_K = self.calculate_K_value(freezer_temp, fridge_temp)
        equivalent_volume = self.calculate_equivalent_volume(VR, VF, standard_K)
        measured_equivalent_volume = self.calculate_equivalent_volume(VR, VF, measured_K)
        effective_volume = VR + VF
        fridge_type = self.determine_fridge_type(effective_volume, VR, VF, fan_type)
        energy_allowance = self.calculate_energy_allowance(equivalent_volume, fridge_type)
        future_energy_allowance = self.calculate_future_energy_allowance(equivalent_volume, fridge_type)
        benchmark_consumption = self.calculate_benchmark_consumption(equivalent_volume, energy_allowance)
        future_benchmark_consumption = self.calculate_future_benchmark_consumption(equivalent_volume, future_energy_allowance)
        monthly_consumption = round(daily_consumption * 30, 0)
        ef_value = round(measured_equivalent_volume / monthly_consumption, 1)
        current_ef_thresholds = self.current_ef_thresholds(energy_allowance, fridge_type)
        current_percent, current_grade = self.calculate_current_efficiency(ef_value, current_ef_thresholds)
        future_ef_thresholds = self.future_ef_thresholds(future_energy_allowance, fridge_type)
        future_percent, future_grade = self.calculate_future_efficiency(ef_value, future_ef_thresholds)
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
            '2018耗電量基準(kWh/月)': int(benchmark_consumption),
            '2027耗電量基準(kWh/月)': int(future_benchmark_consumption),
            '實測月耗電量(kWh/月)': int(monthly_consumption),
            '實測EF值': ef_value,
            '2018效率基準百分比(%)': int(current_percent),
            '2018效率等級': current_grade,
            '2027新效率基準百分比(%)': int(future_percent),
            '2027新效率等級': future_grade
        })
        return results

    def calculate_K_value(self, freezer_temp, fridge_temp):
        return round((30 - freezer_temp) / (30 - fridge_temp), 2)

    def calculate_equivalent_volume(self, VR, VF, K):
        return round(VR + (K * VF), 1)

    def determine_fridge_type(self, equivalent_volume, VR, VF, fan_type):
        if VF == 0:
            return 5
        elif equivalent_volume < 400 and fan_type == 1:
            return 1
        elif equivalent_volume >= 400 and fan_type == 1:
            return 2
        elif equivalent_volume < 400 and fan_type == 0:
            return 3
        else:
            return 4

    def calculate_energy_allowance(self, equivalent_volume, fridge_type):
        if fridge_type == 1:
            return round(equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round(equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round(equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round(equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round(equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)

    def calculate_future_energy_allowance(self, equivalent_volume, fridge_type):
        if fridge_type == 1:
            return round(1.3 * equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round(1.3 * equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round(1.3 * equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round(1.3 * equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round(1.36 * equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)

    def calculate_benchmark_consumption(self, equivalent_volume, energy_allowance):
        return round(equivalent_volume / energy_allowance, 0)

    def calculate_future_benchmark_consumption(self, equivalent_volume, future_energy_allowance):
        return round(equivalent_volume / future_energy_allowance, 0)

    def calculate_current_efficiency(self, ef_value, thresholds):
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 0)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 0)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[0] * 100, 0)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[0] * 100, 0)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[0] * 100, 0)
        else:
            grade = "5級"
            final_percent = round(ef_value / thresholds[0] * 100, 0)
        return final_percent, grade

    def calculate_future_efficiency(self, ef_value, thresholds):
        return self.calculate_current_efficiency(ef_value, thresholds)


# ============================================================================
# DraggableLine - 不動
# ============================================================================
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


# ============================================================================
# 視窗尺寸響應式工具
# ============================================================================
def get_fig_size(screen_w, screen_h):
    """根據螢幕尺寸計算 figure 大小（英吋）"""
    # 圖表區域：右側，寬度約 70% 螢幕 width，高度約 85% 螢幕 height
    fw = screen_w * 0.70 / 100  # DPI=100，所以吋 = 像素 / 100
    fh = screen_h * 0.80 / 100
    return fw, fh

def get_textwidget_size(screen_w, screen_h):
    """根據螢幕尺寸計算 Text widget 字元寬度與行數"""
    # 左側面板寬度約 30% 螢幕，扣掉 padding
    char_w = max(30, int(screen_w * 0.28 / 7))   # 7px per char average
    lines = max(20, int(screen_h * 0.75 / 18))   # 18px per line
    return char_w, lines


# ============================================================================
# 核心函數
# ============================================================================
def plot_chart():
    ax1_color = "lightcyan"
    ax2_color = "lightyellow"
    try:
        global ax1, ax2, ax2a, ax2b, ax2c

        fig.clf()

        gs = fig.add_gridspec(2, 1, height_ratios=[7, 3])
        ax1 = fig.add_subplot(gs[0, 0], facecolor=ax1_color)
        ax2 = fig.add_subplot(gs[1, 0], sharex=ax1, facecolor=ax2_color)

        for ax in fig.axes:
            if ax not in [ax1, ax2]:
                ax.remove()

        ax2a = ax2
        ax2b = ax2.twinx()
        ax2c = ax2.twinx()
        ax2c.spines['right'].set_position(('outward', 35))

        global global_df
        if global_df is None or global_df.empty:
            messagebox.showerror("錯誤", "請先選擇檔案！")
            return

        df = global_df.copy()
        power_columns = ['V', 'I', 'W']
        df_power = df[['datetime'] + power_columns]
        temp_columns = [c for c in df.columns[1:-3]]
        df_temp = df[['datetime'] + temp_columns]

        fig.suptitle(chart_title.get())
        colors = plt.cm.tab20.colors

        for idx, col in enumerate(temp_columns):
            color = colors[idx % len(colors)]
            ax1.plot(df_temp['datetime'], df_temp[col], label=col, color=color)
        ax1.set_ylabel("溫度")
        ax1.legend(temp_columns, loc='upper left')
        ax1.set_facecolor(ax1_color)
        ax1.grid(True, linestyle='dashdot')
        ax1.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)

        if len(power_columns) >= 3:
            u_col = power_columns[0]
            i_col = power_columns[1]
            p_col = power_columns[2]

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
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%d-%H:%M'))
        ax2.set_xlim(ax1.get_xlim())

        global start_line, end_line
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
        root.update_idletasks()

    except Exception as e:
        messagebox.showerror("錯誤", f"繪製圖表時發生錯誤：{e}")
        print(f"plot_chart: err：{e}")


def calculate_statistics():
    try:
        global global_df
        if global_df is None or global_df.empty:
            messagebox.showerror("錯誤", "請先選擇檔案！")
            return

        df = global_df.copy()
        power_columns = ['V', 'I', 'W']
        if len(power_columns) < 3:
            messagebox.showerror("錯誤", "電力欄位不足！")
            return

        start = pd.to_datetime(f"{start_date.get()} {start_time.get()}")
        end = pd.to_datetime(f"{end_date.get()} {end_time.get()}")
        minutes_difference = int((end - start).total_seconds() / 60)

        filtered_df = df[(df['datetime'] >= start) & (df['datetime'] <= end)].copy()
        if filtered_df.empty:
            messagebox.showinfo("結果", "指定範圍內沒有資料！")
            return

        averages = filtered_df.mean(numeric_only=True)
        max_values = filtered_df.max(numeric_only=True)
        min_values = filtered_df.min(numeric_only=True)

        power_column = power_columns[2]
        power_throttle = float(on_off_throttle_entry_var.get())
        if power_column in filtered_df.columns:
            filtered_df.loc[:, 'power_on'] = filtered_df[power_column] >= power_throttle
            power_cycles = int(filtered_df['power_on'].astype(int).diff().fillna(0).abs().sum() // 2)

            mask = filtered_df['power_on']
            groups = (mask != mask.shift()).cumsum()
            segments = pd.DataFrame({
                '狀態': mask,
                '區段編號': groups,
                '時間': filtered_df['datetime']
            }).groupby(['區段編號', '狀態']).agg({'時間': ['min', 'max']}).reset_index()

            if len(segments) > 2:
                segments = segments.iloc[1:-1]
            segments['持續時間'] = (segments[('時間', 'max')] - segments[('時間', 'min')]).dt.total_seconds()

            above_segments = segments[segments['狀態']]
            below_segments = segments[~segments['狀態']]
            above_count = len(above_segments)
            below_count = len(below_segments)
            above_avg_time = int((above_segments['持續時間'].mean() / 60) if above_count > 0 else 0)
            below_avg_time = int(((below_segments['持續時間'].mean() / 60) + 1) if below_count > 0 else 0)

            if above_avg_time + below_avg_time > 0:
                above_percentage = int((above_avg_time / (above_avg_time + below_avg_time)) * 100)
            else:
                above_percentage = 0
        else:
            power_cycles = "無法計算，缺少 P(W) 欄位"

        total_seconds = int((filtered_df['datetime'].iloc[-1] - filtered_df['datetime'].iloc[0]).total_seconds())
        filter_WH = int(filtered_df['W'].sum() / 60)

        if total_seconds > 0:
            wp_24h_difference = int((filter_WH / total_seconds) * (24 * 3600))
        else:
            wp_24h_difference = "無法計算，時間範圍不足"

        fan_type = fan_type_var.get()
        vf = float(vf_entry_var.get()) if vf_entry_var.get().replace('.', '', 1).isdigit() else 0
        vr = float(vr_entry_var.get()) if vr_entry_var.get().replace('.', '', 1).isdigit() else 0
        fridge_temp = float(temp_r_entry_var.get()) if temp_r_entry_var.get().replace('.', '', 1).isdigit() else 3
        freezer_temp = float(temp_f_entry_var.get()) if temp_f_entry_var.get().replace('.', '', 1).lstrip('-').isdigit() else -18.0
        energy_calculator = EnergyCalculator()
        if isinstance(wp_24h_difference, (int, float)) and vf > 0 and vr > 0:
            daily_consumption = round(wp_24h_difference / 1000, 3)
            results = energy_calculator.calculate(vf, vr, daily_consumption, fridge_temp, freezer_temp, fan_type)
        else:
            results = None

        result_textbox.delete(1.0, tk.END)
        result_textbox.insert(tk.END, f"統計範圍：{start} ~ {end}\n")
        result_textbox.insert(tk.END, "平均/最大/最小：\n")
        for column in averages.index:
            result_textbox.insert(tk.END, f"{column}: {averages[column]:.1f} / {max_values[column]:.1f} / {min_values[column]:.1f}\n")
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


def save_results():
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return
        content = result_textbox.get(1.0, tk.END).strip()
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)
        messagebox.showinfo("成功", "結果已成功儲存！")
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存結果時發生錯誤：{e}")


def select_file():
    global global_df
    format_value = input_format_var.get()
    if format_value == 'xls':
        file_path = filedialog.askopenfilename(filetypes=[("Office 97 files", "*.xls")])
        if not file_path:
            return
        csv_path.set(file_path)
        try:
            df = pd.read_excel(file_path)
            selected_columns = [1] + list(range(6, 29))
            available_columns = [col for col in selected_columns if col < len(df.columns)]
            df = df.iloc[:, available_columns]
            df = df.dropna(axis=1, how='all')
            if df.empty:
                messagebox.showerror("錯誤", "讀取的資料為空！")
                global_df = None
                return

            if len(df.columns) > 0:
                first_col = df.columns[0]
                if df[first_col].dtype == 'object':
                    df[first_col] = df[first_col].astype(str).str.strip()
                    df[first_col] = df[first_col].replace(['', 'nan', 'NaN', 'None'], pd.NA)
                if df.empty:
                    messagebox.showerror("錯誤", "第一列沒有有效的日期時間資料！")
                    global_df = None
                    return
                try:
                    df['datetime'] = pd.to_datetime(df[first_col], errors='coerce')
                except Exception as e:
                    messagebox.showerror("錯誤", f"日期時間轉換錯誤：{e}")
                    global_df = None
                    return
                df = df.dropna(subset=['datetime'])
                df = df.sort_values(by='datetime')
                df = df.reset_index(drop=True)
                if df.empty:
                    messagebox.showerror("錯誤", "無法解析日期時間欄位，資料為空！")
                    global_df = None
                    return
            else:
                messagebox.showerror("錯誤", "資料格式錯誤，缺少時間欄位！")
                global_df = None
                return

            if '日期' in df.columns and 'datetime' in df.columns:
                df = df.drop('日期', axis=1)

            df.rename(columns={'21': 'V', '22': 'I', '23': 'W'}, inplace=True)
            df = df.set_index('datetime')
            df = df.resample('1min').mean(numeric_only=True)
            df = df.round(1)
            df = df.reset_index()

            output_dir = os.path.dirname(file_path)
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
            if not file_path:
                return
            csv_path.set(file_path)
            df = pd.read_csv(file_path)
            if '日期' in df.columns:
                df = df.drop('日期', axis=1)
            if 'datetime' not in df.columns:
                messagebox.showerror("錯誤", "沒有datetime欄位")
                return
            df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
        except Exception as e:
            messagebox.showerror("錯誤", f"日期時間轉換錯誤：{e}")
            global_df = None
            return

    print(df.head(5))
    global_df = df.copy()
    try:
        first_dt = df['datetime'].iloc[0]
        last_dt = df['datetime'].iloc[-1]
        start_date.set(first_dt.strftime('%Y-%m-%d'))
        start_time.set(first_dt.strftime('%H:%M:%S'))
        end_date.set(last_dt.strftime('%Y-%m-%d'))
        end_time.set(last_dt.strftime('%H:%M:%S'))
    except Exception:
        pass


def on_closing():
    try:
        result = messagebox.askokcancel("退出", "確定要退出程式嗎？")
        if result:
            root.quit()
            root.destroy()
    except Exception:
        try:
            root.quit()
            root.destroy()
        except:
            pass


# ============================================================================
# 視窗縮放響應
# ============================================================================
def on_resize(event):
    """視窗大小改變時，重新計算 figure 和文字框尺寸"""
    w = root.winfo_width()
    h = root.winfo_height()
    if w < 100 or h < 100:
        return

    # 重新計算 figure 大小並更新
    fw, fh = get_fig_size(w, h)
    fig.set_size_inches(fw, fh)
    canvas.draw()

    # 重新計算 text widget size 並更新
    tw, tl = get_textwidget_size(w, h)
    result_textbox.config(width=tw, height=tl)


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
        if event.state & 0x4:
            if event.keysym == "Up":
                increment_date_time(var, 1, unit)
            elif event.keysym == "Down":
                increment_date_time(var, -1, unit)
    widget.bind("<KeyPress-Up>", on_key)
    widget.bind("<KeyPress-Down>", on_key)


# ============================================================================
# GUI 初始化
# ============================================================================
rcParams['font.sans-serif'] = ['Microsoft JhengHei']
rcParams['axes.unicode_minus'] = False
rcParams['font.size'] = 8
rcParams['axes.titlesize'] = 8
rcParams['axes.labelsize'] = 8
rcParams['xtick.labelsize'] = 8
rcParams['ytick.labelsize'] = 8
rcParams['legend.fontsize'] = 8
rcParams['figure.titlesize'] = 12

root = tk.Tk()
root.title("Plotter with Statistics for access 97 0.3 (響應式)")

# 取得螢幕尺寸
screen_w = root.winfo_screenwidth()
screen_h = root.winfo_screenheight()
print(f"螢幕尺寸: {screen_w} x {screen_h}")

# 初始視窗最大化（可用，或改為固定比例）
# root.state('zoomed')

# 全域變數
csv_path = tk.StringVar()
chart_title = tk.StringVar()
start_date = tk.StringVar()
start_time = tk.StringVar()
end_date = tk.StringVar()
end_time = tk.StringVar()
input_format_var = tk.StringVar(value="xls")
global_df = None

# ----- 選單 -----
menubar = tk.Menu(root)
root.config(menu=menubar)
options_menu = tk.Menu(menubar, tearoff=0)
options_menu.add_radiobutton(label="xls", variable=input_format_var, value="xls")
options_menu.add_radiobutton(label="csv", variable=input_format_var, value="csv")
menubar.add_cascade(label="檔案格式", menu=options_menu)

# ----- 主要容器：左右兩欄 -----
root.columnconfigure(0, weight=0)  # 左面板固定
root.columnconfigure(1, weight=1)  # 右側圖表自適應撐滿
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=0)      # toolbar row

# ---- 左側面板（控制區 + 結果文字框 + 參數）----
left_frame = tk.Frame(root, padx=5, pady=5)
left_frame.grid(row=0, column=0, sticky="nsew")
left_frame.columnconfigure(0, weight=1)

# --- 檔案選擇區 ---
file_frame = tk.Frame(left_frame)
file_frame.pack(fill="x", pady=(0, 5))
tk.Label(file_frame, text="檔案路徑:").pack(side="left", padx=2)
tk.Entry(file_frame, textvariable=csv_path, width=30).pack(side="left", fill="x", expand=True, padx=2)
tk.Button(file_frame, text="選擇檔案", command=select_file).pack(side="left", padx=2)

# --- 圖表標題 ---
title_frame = tk.Frame(left_frame)
title_frame.pack(fill="x", pady=(0, 5))
tk.Label(title_frame, text="圖表標題:").pack(side="left", padx=2)
tk.Entry(title_frame, textvariable=chart_title, width=30).pack(side="left", fill="x", expand=True, padx=2)

# --- 按鈕區 ---
btn_frame = tk.Frame(left_frame)
btn_frame.pack(fill="x", pady=(0, 5))
tk.Button(btn_frame, text="繪製圖表", command=plot_chart).pack(side="left", padx=2, fill="x", expand=True)
tk.Button(btn_frame, text="計算平均值", command=calculate_statistics).pack(side="left", padx=2, fill="x", expand=True)
tk.Button(btn_frame, text="儲存結果", command=save_results).pack(side="left", padx=2, fill="x", expand=True)

# --- 結果文字框（響應式尺寸）---
tw, tl = get_textwidget_size(screen_w, screen_h)
result_textbox = tk.Text(left_frame, width=tw, height=tl, wrap=tk.WORD)
result_textbox.pack(fill="both", expand=True,pady=5)

# --- 能耗計算參數 ---
param_frame = tk.Frame(left_frame)
param_frame.pack(fill="x", pady=(5, 0))

# VF / VR row
row1 = tk.Frame(param_frame)
row1.pack(fill="x", pady=2)
tk.Label(row1, text="能耗:", width=8, anchor="e").pack(side="left")
tk.Label(row1, text="VF(L):").pack(side="left", padx=(0,2))
vf_entry_var = tk.StringVar(value="150")
tk.Entry(row1, textvariable=vf_entry_var, width=6).pack(side="left")
tk.Label(row1, text="VR(L):").pack(side="left", padx=(4,2))
vr_entry_var = tk.StringVar(value="350")
tk.Entry(row1, textvariable=vr_entry_var, width=6).pack(side="left")

# Temp F / R row
row2 = tk.Frame(param_frame)
row2.pack(fill="x", pady=2)
tk.Label(row2, text="F溫度:").pack(side="left", padx=(0,2))
temp_f_entry_var = tk.StringVar(value="-18.0")
tk.Entry(row2, textvariable=temp_f_entry_var, width=6).pack(side="left")
tk.Label(row2, text="R溫度:").pack(side="left", padx=(4,2))
temp_r_entry_var = tk.StringVar(value="3.0")
tk.Entry(row2, textvariable=temp_r_entry_var, width=6).pack(side="left")

# Fan type + throttle row
row3 = tk.Frame(param_frame)
row3.pack(fill="x", pady=2)
fan_type_var = tk.IntVar(value=1)
tk.Checkbutton(row3, text="風扇式", variable=fan_type_var).pack(side="left")
tk.Label(row3, text="OnOfTh:").pack(side="left", padx=(8,2))
on_off_throttle_entry_var = tk.StringVar(value="3.0")
tk.Entry(row3, textvariable=on_off_throttle_entry_var, width=6).pack(side="left")

# Separator
ttk.Separator(param_frame, orient="horizontal").pack(fill="x", pady=4)

# --- 日期時間設定 ---
dt_label = tk.Label(param_frame, text="資料日期範圍")
dt_label.pack(anchor="w", pady=(0, 2))

# start date / time
s_frame = tk.Frame(param_frame)
s_frame.pack(fill="x", pady=1)
tk.Label(s_frame, text="開始:", width=5, anchor="e").pack(side="left")
start_date_entry = tk.Entry(s_frame, textvariable=start_date, width=12)
start_date_entry.pack(side="left", padx=2)
bind_increment(start_date_entry, start_date, "day")
start_time_entry = tk.Entry(s_frame, textvariable=start_time, width=10)
start_time_entry.pack(side="left", padx=2)
bind_increment(start_time_entry, start_time, "hour")

# end date / time
e_frame = tk.Frame(param_frame)
e_frame.pack(fill="x", pady=1)
tk.Label(e_frame, text="結束:", width=5, anchor="e").pack(side="left")
end_date_entry = tk.Entry(e_frame, textvariable=end_date, width=12)
end_date_entry.pack(side="left", padx=2)
bind_increment(end_date_entry, end_date, "day")
end_time_entry = tk.Entry(e_frame, textvariable=end_time, width=10)
end_time_entry.pack(side="left", padx=2)
bind_increment(end_time_entry, end_time, "hour")


# ---- 右側：圖表區 ----
right_frame = tk.Frame(root, padx=5, pady=5)
right_frame.grid(row=0, column=1, sticky="nsew")
right_frame.rowconfigure(0, weight=1)
right_frame.columnconfigure(0, weight=1)

fw, fh = get_fig_size(screen_w, screen_h)
fig = plt.figure(figsize=(fw, fh), facecolor='lightgray')

canvas = FigureCanvasTkAgg(fig, master=right_frame)
canvas_widget = canvas.get_tk_widget()
canvas_widget.grid(row=0, column=0, sticky="nsew")

# Toolbar
toolbar_frame = tk.Frame(root)
toolbar_frame.grid(row=1, column=1, padx=5, pady=(0,5), sticky="ew")
toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
toolbar.update()

# 綁定視窗縮放事件
root.bind("<Configure>", on_resize)
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
