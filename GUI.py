#%% 載入需要的模組
from datetime import datetime
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import ttk
from scipy.interpolate import CubicSpline
import numpy as np
from scipy.optimize import minimize

#%% 列舉需要的參數
# position = True  # True = Buyer
# Notional = 1000000  # Principal = 1 million USD
# Startd = datetime.date(2022, 6, 1)
# period = 3
# Endd = datetime.date(2024, 6, 1)
# freq = 2

#%% 載入swap curve data
import pandas as pd
## 讀資料
swap_df = pd.read_csv("/Users/changshuowen/112_2/財務演算法/Final/data/DSWP.csv", parse_dates=['DATE'])
workday = swap_df.DATE

## 把假日的值用前一工作日的 swap rates 遞補
fulld_df = pd.DataFrame({'DATE': pd.date_range(start=workday.iloc[0], end=workday.iloc[-1])})
swap_fdf = pd.merge(fulld_df, swap_df, on='DATE', how='left')
swap_fdf.ffill(inplace=True)

## 將日期設為索引
swap_fdf.set_index('DATE', inplace=True)
tenor_list = [int(x) for x in swap_fdf.columns]
# pd.Timestamp(stdt) in workday.values
# type(swap_fdf.index)
# swap_fdf.loc[stdt]

#%% 設計 Cash Flow Table Calculation Engine
def CF_summation(YTM, f, spot_rate):
    summation = 0
    for index, spot in enumerate(spot_rate):
        summation += (YTM / f) / ((1 + spot / f) ** (index + 1))
    return summation

def spot(tenor, freq, swap_rate):
    spot_rate = []
    for index, (t, ytm) in enumerate(zip(tenor, swap_rate)):
        WIP = (1 + (ytm/100) / freq) / (1 - CF_summation(ytm/100, freq, spot_rate))
        spot = ((WIP) ** (1 / (t * freq)) - 1) * freq
        spot_rate.append(spot)
    return spot_rate

def PV_fac(tenor, freq, spot_rate):
    PV = [(1 / ((1 + spot / freq)** (t * freq))) for t, spot in zip(tenor, spot_rate)]
    return PV

def forward(tenor, freq, spot_rate):
    forward_rate = []
    last_t, last_spot = 0, 0
    for index, (t, spot) in enumerate(zip(tenor, spot_rate)):
        forward = spot if (index == 0) else (((1 + spot / freq) ** (t * freq)) / ((1 + last_spot / freq) ** (last_t * freq)) - 1) * freq
        forward_rate.append(forward)
        last_t, last_spot = t, spot
    return forward_rate

def date_period(start_date, T, freq):
    date_list = []
    pay_date = start_date
    for _ in range(int(T / 12 * freq)):
        pay_date = pay_date + relativedelta(months = 12 / freq)
        date_list.append(pay_date.date())
    return date_list

def float_side(notional, forward_rate, start_date, pay_date, PV_factor):
    PV_floatCF = []
    last_date = start_date.date()
    for index, (forward, date, PV) in enumerate(zip(forward_rate, pay_date, PV_factor)):
        PV_ftCF = notional * forward * ((date - last_date).days / 365) * PV if (index != (len(pay_date) - 1)) else notional * (1 + forward * (date - last_date).days / 365) * PV
        PV_floatCF.append(PV_ftCF)
        last_date = date
    return PV_floatCF

def fixed_side(notional, PV_float_CF, start_date, pay_date, PV_factor, initial_guess):
    # Step 1: 算出當天 Swap 期限結構下的 Fixed Rate
    ## 寫好最佳化用的函數 (NPV)
    PV_float = sum(PV_float_CF)
    def PV_fixed(fixed_rate):
        PV_fixed = 0
        last_date = start_date.date()
        for index, (date, PV) in enumerate(zip(pay_date, PV_factor)):
            PV_fxCF = notional * fixed_rate * ((date - last_date).days / 365) * PV if (index != (len(pay_date) - 1)) else notional * (1 + fixed_rate * (date - last_date).days / 365) * PV
            PV_fixed += PV_fxCF
            last_date = date
        return abs(PV_fixed - PV_float)
    ## 執行最小化
    fixed_rate = minimize(PV_fixed, initial_guess).x.item()

    # Step 2: 輸出 Fixed Cash Flows
    PV_fixedCF = []
    last_date = start_date.date()
    for index, (date, PV) in enumerate(zip(pay_date, PV_factor)):
        PV_fxCF = notional * fixed_rate * ((date - last_date).days / 365) * PV if (index != (len(pay_date) - 1)) else notional * (1 + fixed_rate * (date - last_date).days / 365) * PV
        PV_fixedCF.append(PV_fxCF)
        last_date = date

    return fixed_rate, PV_fixedCF

def CF_table(pars_list):
    # 定義重要變數
    position, notional, startd, T, endd, frequency = pars_list
    # 利用 NCS 建構完整 Swap Curve 曲線
    swap_curve = swap_fdf.loc[startd]
    tenor, rates = swap_curve.index.astype(int), [x for x in swap_curve.values]
    cs = CubicSpline(tenor, rates, bc_type='natural')
    interp_time_points = np.linspace(1 / frequency, int(T / 12), num=int(T / 12 * frequency))
    swap_fcurve = cs(interp_time_points)
    # 計算所有 term structure 及導出最佳化 fixed_rate
    spot_rate = spot(interp_time_points, frequency, swap_fcurve)
    PV_factor = PV_fac(interp_time_points, frequency, spot_rate)
    forward_rate = forward(interp_time_points, frequency, spot_rate)
    pay_date = date_period(startd, T, frequency)
    PV_float_CF = float_side(notional, forward_rate, startd, pay_date, PV_factor)
    fixed_rate, PV_fixed_CF = fixed_side(notional, PV_float_CF, startd, pay_date, PV_factor, spot_rate[0])
    # 輸出現金流量表格
    out_table = pd.DataFrame({
        'Tenor': interp_time_points,
        'Swap Rate': [round(x/100, 6) for x in swap_fcurve],
        'Spot Rate': [round(x, 6) for x in spot_rate],
        'PV': [round(x, 4) for x in PV_factor],
        'Forward Rate': [round(x, 6) for x in forward_rate],
        'Date': pay_date,
        'PV (Fixed CF)': [round(x, 2) for x in PV_float_CF],
        'PV (Float CF)': [round(x, 2) for x in PV_float_CF],
        })
    # 美化一下 DataFrame
    def format_as_percent(x):
        return "{:.4%}".format(x)
    out_table[['Swap Rate', 'Spot Rate', 'Forward Rate']] = out_table[['Swap Rate', 'Spot Rate', 'Forward Rate']].applymap(format_as_percent)
    # 添加重要資訊
    new_row = pd.DataFrame({
        'Tenor': [''],
        'Swap Rate': [''],
        'Spot Rate': [''],
        'PV': ['Fixed Rate:'],
        'Forward Rate': ["{:.4%}".format(fixed_rate)],
        'Date': ['NPV:'],
        'PV (Fixed CF)': [round(sum(PV_fixed_CF), 2)],
        'PV (Float CF)': [round(sum(PV_float_CF), 2)],
        })
    out_table = pd.concat([new_row, out_table], ignore_index=True)
    if not position:
        out_table = out_table.reindex(columns=['Tenor', 'Swap Rate', 'Spot Rate', 'PV',
                                               'Forward Rate', 'Date', 'PV (Float CF)', 'PV (Fixed CF)'])
    return out_table, fixed_rate

#%% 設計介面
def on_entry_click(event, entry, default_text):
    if entry.get() == default_text:
        entry.delete(0, tk.END)  # 清空提示文字
        entry.config(fg='black')

def on_focusout(event, entry, default_text):
    if entry.get() == '':
        entry.delete(0, tk.END)
        entry.insert(0, default_text)
        entry.config(fg='grey')

def label_entry(label_text, text, index):
    ## 創建標題和輸入框
    label = tk.Label(root, text=label_text)
    entry = tk.Entry(root, fg='grey')
    entry.insert(0, text)
    entry.bind('<FocusIn>', lambda event: on_entry_click(event, entry, text))
    entry.bind('<FocusOut>', lambda event: on_focusout(event, entry, text))
    ## 使用 grid 佈局管理器將標題和輸入框放置在表格中
    label.grid(row=index, column=0, padx=5, pady=5, sticky="e")
    entry.grid(row=index, column=1, padx=5, pady=5)
    if (index == 2):
        label = tk.Label(root, text="between 2000/07/03 ~ 2016/10/28 (due to data availability)")
        label.grid(row=index, column=2, padx=5, pady=5, sticky="w")
    elif (index == 3):
        label = tk.Label(root, text="set valid period such that maturity does not exceed 2016/10/28.")
        label.grid(row=index, column=2, padx=5, pady=5, sticky="w")
    elif (index == 4):
        label = tk.Label(root, text="need to be a factor of 'Period'. (1=annual; 2=semi-annual, etc.)")
        label.grid(row=index, column=2, padx=5, pady=5, sticky="w")
    return label_text, entry

def collect_response(entries):
    global text_dict, pars_list
    text_dict = {}
    for i, (label_text, entry) in enumerate(entries):        
        user_input = entry.get()
        entry.config(state="disabled")
        text_dict[label_text] = user_input
        print(f"{label_text} {user_input}")
    ## position
    if text_dict[label_p] == "payer":
        pos = True
    elif text_dict[label_p] == "receiver":
        pos = False
    else:
        raise ValueError("invalid 'Position' input, please enter again.")
    ## notional
    try:
        ntnl = int(text_dict[label_n])
    except ValueError:
        raise ValueError("invalid 'Notional' input, please enter again.")
    ## start date
    try:
        stdt = datetime.strptime(text_dict[label_sd], "%Y/%m/%d")
        if (stdt > datetime(2016, 10, 28) or stdt < datetime(2000, 7, 3)):
            raise TypeError("'StartDate' out of bound, please enter again.")
    except ValueError:
        raise ValueError("invalid 'Start Date' input, please enter again.")
    ## period
    try:
        t = int(text_dict[label_t])
        endt = stdt + relativedelta(months = t)
        if (endt > datetime(2016, 10, 28)):
            raise TypeError("'Maturity' out of bound, please enter valid 'Period'.")
        if (t < 0):
            raise TypeError("please enter positive 'Period'.")
    except ValueError:
        raise ValueError("invalid 'Period' input, please enter again.")
    ## frequency
    try:
        freq = int(text_dict[label_f])
        if freq not in [1, 2, 4, 12]:
            raise TypeError("invalid 'Frequency' input, please enter again.")
        elif (t % freq != 0):
            raise TypeError("'Frequency' is not a factor of 'Period', please enter again.")
    except ValueError:
        raise ValueError("invalid 'Frequency' input, please enter again.")

    pars_list = [pos, ntnl, stdt, t, endt, freq]
    return text_dict, pars_list

def show_results():
    try:
        text_dict, pars_list = collect_response(entries)
    except ValueError as e:
        print(e)
        for widget in root.winfo_children():
            widget.grid_forget()
        # 創建返回按鈕
        back_button = tk.Button(root, text=e, command=create_input_interface)
        back_button.grid(row=len(entries), columnspan=2, pady=20)
        return
    except TypeError as e:
        print(e)
        for widget in root.winfo_children():
            widget.grid_forget()
        # 創建返回按鈕
        back_button = tk.Button(root, text=e, command=create_input_interface)
        back_button.grid(row=len(entries), columnspan=2, pady=20)
        return

    for widget in root.winfo_children():
        widget.grid_forget()
    
    for i, (label, text) in enumerate(text_dict.items()):
        label = tk.Label(root, text=f"{label} {text}")
        label.grid(row=i, column=0, columnspan=2, sticky="w")
    
    matt = (pars_list[2] + relativedelta(months = pars_list[3])).strftime("%Y/%m/%d")
    label = tk.Label(root, text=f"Maturity Date: {matt}")
    label.grid(row=len(entries), column=0, columnspan=2, sticky="w")

    # 於介面輸出表格
    out_table, fixed_rate = CF_table(pars_list)
    columns = out_table.columns.tolist()
    data_rows = out_table.values.tolist()

    treeview = ttk.Treeview(root, columns=columns, show="headings")
    for index, col in enumerate(columns):
        width = 50 if (index == 0) else 100
        treeview.heading(col, text=col)
        treeview.column(col, width=width)
    for index, row in enumerate(data_rows):
        if (index == 0):
            treeview.insert("", "end", values=row, tags=("first_row",))
        else:
            treeview.insert("", "end", values=row)
    treeview.tag_configure("first_row", background="lightblue")

    # ScrollBar
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=treeview.yview)
    treeview.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=len(entries)+1, column=2, sticky="ns")

    treeview.grid(row=len(entries)+1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

    # 創建返回按鈕
    back_button = tk.Button(root, text="Back", command=create_input_interface)
    back_button.grid(row=len(entries) + 2, columnspan=2, pady=20)

    root.geometry("800x450")

def create_input_interface():
    root.geometry("800x300")
    # 清空原有視窗的所有小部件
    for widget in root.winfo_children():
        widget.grid_forget()
    # 保存標籤和輸入框的列表
    global entries
    entries = []
    entries.append(label_entry(label_p, text_p, 0))
    entries.append(label_entry(label_n, text_n, 1))
    entries.append(label_entry(label_sd, text_d, 2))
    entries.append(label_entry(label_t, text_t, 3))
    entries.append(label_entry(label_f, text_f, 4))
    # 創建提交按鈕
    submit_button = tk.Button(root, text="Submit", command=show_results)
    submit_button.grid(row=5, columnspan=2, pady=20)
    root.bind('<Return>', lambda event: submit_button.invoke())

#%% 執行
# 初始化介面
root = tk.Tk()
root.title("Interest Rate Swap Cashflow Demonstration")
root.geometry("800x300")

# 配置輸入框的默認提示文字
label_p = "Position:"
label_n = "Notional:"
label_sd = "Start Date:"
label_t = "Period (M):"
label_f = "Frequency:"

text_p = "ex: payer/receiver"
text_n = "ex: 1000000"
text_d = "ex: 2016/01/03"
text_t = "ex: 1/3/6/12/24"
text_f = "ex: 1/2/4/12"

create_input_interface()

root.mainloop()






















