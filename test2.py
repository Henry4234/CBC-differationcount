import json
import numpy as np
import pandas as pd

import matplotlib.pyplot as plt
from matplotlib.pylab import mpl
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import mplcursors
mpl.rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 中文顯示
mpl.rcParams['axes.unicode_minus'] = False  # 負號顯示

account="admin"

jsonfile = open('testdata\score.json','rb')
# jsonfile = pd.read_json('testdata/score.json')
print(jsonfile)
rawdata = json.load(jsonfile)

aa = rawdata[account]
df = pd.DataFrame(aa)
df["score"] = df["score"].astype("int")
df["timestamp"] = df["timestamp"].astype("datetime64")
oldest = df["timestamp"].head(1)
latest = df["timestamp"].tail(1)
oldest_month = oldest.astype("datetime64[M]")
latest_month = latest.astype("datetime64[M]")
latest_month = latest_month + np.timedelta64(30,'D')
# df[]
# print(df)
# print(aa)
sfigure = plt.figure(figsize=(7, 4), dpi=80, facecolor="#FFEEDD", frameon=True)  #創建繪圖物件f figsize的單位是英寸 像素 = 英寸*解析度
fig1 = sfigure.add_subplot(1, 1, 1)  # 三個引數，依次是：行，列，當前索引
fig1.set_title("應考次數", loc='center', pad=20, fontsize='xx-large', color='black')    #設定標題
fig1.set_xlabel("日期") #確定x坐標軸標題
fig1.xaxis.set_major_formatter(mdates.DateFormatter("%y/%m/%d"))
fig1.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
fig1.xaxis.set_minor_locator(mdates.WeekdayLocator(interval=1))
# fig1.xaxis.set_minor_locator(mdates.MonthLocator(interval=0.5))

sfigure.autofmt_xdate()
fig1.set_xlim(oldest_month,latest_month)
fig1.set_ylim(0,100)
fig1.set_yticks(np.arange(0,110,10))  #設定y坐標軸刻度
# fig1.set_yticklabels(np.arange(0,110,10))
fig1.set_ylabel("成績") #確定y坐標軸標題

fig1.grid(which='major', axis='x', color='gray', linestyle='-', linewidth=0.5, alpha=0.2)  #設定網格
# 創建一副子圖

# 創建資料源：x軸是等間距的一組數
x = df["timestamp"]
y = df["score"]
# x = np.arange(-20 , 20 , 0.1)
# x = ["2020/01/31","2020/02/31","2020/03/31","2020/04/31"]
# x  = fig1.gca().xaxis.set_major_formatter(mdates.DateFormatter("%y/%m/%d"))
# y1 = np.sin(x)
# y2 = np.cos(x)

##方法一:找第三方函示庫 mplcursor
dot = fig1.plot(x, y,'*', color='red', label='成績',markersize=12)  # 畫點
line = fig1.plot(x, y,linewidth = 2, color='red', label='成績',linestyle ="-",markersize=12)  # 畫線
cursor = mplcursors.cursor(dot,hover = True)
cursor.connect("add",lambda sel:sel.annotation.set_text("時間:" + str(df["timestamp"][sel.index]) + ",\n" + "分數:" + str(df["score"][sel.index])))

plt.show()