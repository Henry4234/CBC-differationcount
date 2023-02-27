#authorised by Henry Tsai
import sys
import datetime

import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import messagebox
import customtkinter as ctk

import json
from PIL import Image
import numpy as np
import pandas as pd

import mplcursors
import matplotlib
# from basedesk_admin import Baccount
matplotlib.use("TkAgg")
from matplotlib.pylab import mpl
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

# 設定中文顯示字體
mpl.rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 中文顯示
mpl.rcParams['axes.unicode_minus'] = False  # 負號顯示


def getaccount(acount):
    global Baccount
    Baccount = str(acount)
    return None

class Search:

    def __init__(self,master,oldmaster=None):  
        # 建立登入後視窗  
        self.master = master
        super().__init__()
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.exit(oldmaster))        
        ctk.set_default_color_theme("blue")  
        # 給主視窗設定標題內容  
        self.master.title("成績查詢")  
        self.master.geometry('1200x650')
        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_columnconfigure(1, weight=3)
        self.master.config(background='#FFEEDD') #設定背景色
        ##圖表設置
        self.canvas = tk.Canvas()
        # self.figure = self.create_matplotlib()
        ##框架設置
        self.labelframe_1 = ctk.CTkFrame(self.master, corner_radius=0,fg_color="#FFDCB9",bg_color="#FFEEDD")
        self.labelframe_1.grid(row=0, column=0,rowspan=2, sticky="nsew")
        # self.labelframe_1.grid_rowconfigure(4, weight=1)
        self.labelframe_2 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.labelframe_2.grid(row=0, column=1)
        # self.labelframe_3 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        # self.labelframe_3.grid(row=0,column=3,rowspan=2,padx=20,pady=15)
        # self.labelframe_2.grid_columnconfigure(1, weight=1)
        ##計算全部應考次數&今年度應考次數
        jsonfile = open('testdata\score.json','rb')
        rawdata = json.load(jsonfile)
        count = pd.DataFrame(rawdata[Baccount])
        count["timestamp"] = count["timestamp"].astype("datetime64")
        testcount = count["ID"].count()
        fliter = count["timestamp"] >= "2023-01-01"
        new = count[fliter]
        newtestcount = new["ID"].count()
        # print(count)
        #左手邊功能區
        self.people = ctk.CTkImage(Image.open("assets\score.png"),size=(80,80))
        self.label_1 = ctk.CTkLabel(
                    self.labelframe_1, 
                    image=self.people,
                    compound="top",
                    text = "成績查詢", 
                    fg_color='#FFDCB9',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_1.grid(row=0,column=0,padx=20, pady=20)
        self.label_2 = ctk.CTkLabel(
                    self.labelframe_1, 
                    text = """歡迎!
admin""", 
#                     text = """歡迎!
# %s"""%(Baccount), 
                    fg_color='#FFDCB9',
                    font=('微軟正黑體',22),
                    text_color="#4A4AFF",
                    width=120,height=90
                    )
        self.label_2.grid(row=1,column=0,padx=20, pady=20)
        self.btn_1=ctk.CTkButton(
            self.labelframe_1,
            command = None, 
            text = "成績面板", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_1.grid(row=2,column=0,pady=10,sticky="ew")
        self.btn_2=ctk.CTkButton(
            self.labelframe_1,
            command = None, 
            text = "考題分析", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_2.grid(row=3,column=0,pady=10,sticky="ew")
        self.btn_3=ctk.CTkButton(
            self.labelframe_1,
            command = lambda: self.exit(oldmaster),
            text = "離開", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_3.grid(row=4,column=0,pady=10,sticky="ew")
        self.cc = ctk.CTkLabel(
            self.master, 
            fg_color="#FFEEDD",
            text='@Design by Henry Tsai',
            text_color="#E0E0E0",
            font=("Calibri",12),
            width=120)
        self.cc.grid(row=1,column=1,sticky='se')

        self.label_2 = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = """總共應考次數:
%d 次"""%(testcount), 
                    fg_color='#FFBD9D',
                    corner_radius=8,
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=80
                    )
        self.label_2.grid(row=0,column=0,padx=20,pady=15)

        self.label_3 = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = """今年應考次數:
%d 次"""%(newtestcount), 
                    fg_color='#FFBD9D',
                    corner_radius=8,
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=80
                    )
        self.label_3.grid(row=0,column=1,padx=20,pady=15)
    

    # def loaddb(self):
    #     global db
        jsonfile = open('testdata\score.json','rb')
        rawdata = json.load(jsonfile)
        db = pd.DataFrame(rawdata[Baccount])
        db["score"] = db["score"].astype("int")
        db["timestamp"] = db["timestamp"].astype("datetime64")
    # """創建繪圖物件"""
    # def create_matplotlib(self):
        ##主要設定(包含圖表大小、圖表標題)
        self.figure = Figure(figsize=(7, 4), dpi=80, facecolor="#FFEEDD", frameon=True)  #創建繪圖物件f figsize的單位是英寸 像素 = 英寸*解析度
        fig1 = self.figure.add_subplot(1, 1, 1)  # 三個引數，依次是：行，列，當前索引
        fig1.set_title("應考次數", loc='center', pad=20, fontsize='xx-large', color='black')    #設定標題
        ##X軸範圍需要依據資料有所變化
        oldest_month = db["timestamp"].head(1).astype("datetime64[M]")
        latest_month = db["timestamp"].tail(1).astype("datetime64[M]")
        latest_month = latest_month + np.timedelta64(31,'D')
        ##X軸設定
        fig1.set_xlabel("日期") #確定x坐標軸標題
        fig1.xaxis.set_major_formatter(mdates.DateFormatter("%y/%m/%d"))   #X軸刻度格式
        fig1.xaxis.set_major_locator(mdates.MonthLocator(interval=1))   #X軸主要刻度
        fig1.xaxis.set_minor_locator(mdates.WeekdayLocator(interval=1)) #X軸次要刻度
        fig1.set_xlim(oldest_month,latest_month)    #設定X軸範圍
        self.figure.autofmt_xdate() #讓X軸標籤好看一點
        ##Y軸設定
        fig1.set_ylabel("成績") #確定y坐標軸標題
        fig1.set_ylim(0,100)    #設定Y軸範圍
        fig1.set_yticks(np.arange(0,110,10))    #Y軸刻度
        ##設定網格
        fig1.grid(which='major', axis='x', color='gray', linestyle='-', linewidth=0.5, alpha=0.2)  
        
        ##資料存取
        # x = np.arange(-20 , 20 , 0.1)
        x = db["timestamp"]
        y = db["score"]
        ##將資料繪製至圖表
        dot = fig1.plot(x, y,'o', color='red', label='成績',markersize=6)  # 畫點
        line = fig1.plot(x, y, color='red', linewidth=2, label='成績', linestyle='-')
        cursor = mplcursors.cursor(dot,hover = True)
        cursor.connect("add",lambda sel:sel.annotation.set_text("時間:" + str(db["timestamp"][sel.index]) + ",\n" + "分數:" + str(db["score"][sel.index])))
        
        self.canvas = FigureCanvasTkAgg(self.figure,self.master)
        # self.canvas.draw()
        self.canvas.get_tk_widget().grid(row=0,column=3,rowspan=2,padx=20,pady=15)
        # self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        # self.canvas._tkcanvas.pack()

    def exit(self,oldmaster):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return


def main():  
    root = ctk.CTk()
    S = Search(root)
    # S.loaddb()
    # S.create_matplotlib()
    #  主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  
