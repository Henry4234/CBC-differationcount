from email import message
from re import T
import tkinter as tk
import pandas as pd
from tkinter import RIDGE, DoubleVar, StringVar, ttk, messagebox,IntVar
from turtle import bgcolor, width
import customtkinter  as ctk
from collections import defaultdict
import json
import winsound
from datetime import datetime
import basedesk, basedesk_admin
duration = 50  # millisecond
freq = 600  # Hz
duration_1 = 900  # millisecond
freq_1 = 1000  # Hz
def getaccount(acount):
    global Baccount
    Baccount = str(acount)
    # Baccount = "henry423"
    print(Baccount)
    return None
class Count:    #建立計數器
    def __init__(self,master,oldmaster=None):
        self.master = master
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.back(oldmaster))
        # super().__init__()
        #建立主視窗前先取得初始資料
        jsonfile = open('testdata\data.json','rb')
        rawdata = json.load(jsonfile)
        self.rawdata = pd.DataFrame(rawdata["blood"])
        self.testyear = self.rawdata['year'].unique().tolist()
        self.testlist = []
        # for item in self.rawdata:
        #     self.testdict[item["year"]].append(item["ID"])
        #刪除重複年
        # testno = 0
        # testno = self.testlist.index('B001')
        # rawdata = self.rawdata[testno]["rawdata"]
        # 建立主視窗,用於容納其它元件  
        
        # self.root = ctk.CTk()
        self.master.title("細胞計數器")
        self.master.geometry('800x625')
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.master.config(background='#FCFCFC')
        self.master.bind('<Key>', self.pending)
        #設置框架_考核片選擇
        self.frame_exam = ctk.CTkFrame(self.master)
        self.frame_exam.configure(bg_color='#FCFCFC',fg_color='#FCFCFC',)
        #設置框架_計時器
        self.frame_time = ctk.CTkFrame(self.master)
        self.frame_time.configure(bg_color='#FCFCFC',fg_color='#BEBEBE',corner_radius=8)
        # self.frame_time.grid_rowconfigure(0, weight=1)
        self.frame_time.grid_columnconfigure(1, weight=1)
        #設置框架_考核片資訊
        self.frame_info = ctk.CTkFrame(self.master)
        self.frame_info.configure(bg_color='#FCFCFC',fg_color='#BEBEBE',corner_radius=8)
        #設置框架_counter
        self.frame_counter = ctk.CTkFrame(self.master)
        self.frame_counter.configure(bg_color='#FCFCFC',fg_color='#F0F0F0',width=600,height=400,corner_radius=8)
        self.frame_counter.propagate(0)
        #設置框架_功能區
        self.frame_function = ctk.CTkFrame(self.master)
        self.frame_function.configure(bg_color='#FCFCFC',fg_color='#FCFCFC',)
        #設置細胞總數
        self.totalcount = tk.IntVar()
        self.totalcount.set(0)
        #設置時間
        self.time = ctk.CTkLabel(
            self.frame_time,
            text="00:00:00",
            text_color= "#000000",
            font=('微軟正黑體',36,'bold'),
            fg_color='#BEBEBE',
            width=120
            )
        self.doTick = True
        # self.update_clock()
        #設置標籤
        self.label_exam1 = ctk.CTkLabel(
            self.frame_exam, 
            text = "請選擇考核年度:", 
            fg_color='#FCFCFC',
            font=('微軟正黑體',20),
            text_color="#000000",
            width=200
            )
        self.input_exam1 = ctk.CTkComboBox(
            master = self.frame_exam, 
            # variable=self.testlist,
            command=self.callback_year,
            values=self.testyear,
            text_color='#000000',
            fg_color='#F0F0F0',
            button_color='#0080FF',
            width=120,height=40,
            state="readonly"
            )
        self.label_exam2 = ctk.CTkLabel(
            self.frame_exam, 
            text = "請選擇考核片:", 
            fg_color='#FCFCFC',
            font=('微軟正黑體',20),
            text_color="#000000",
            width=200
            )
        self.input_exam2 = ctk.CTkComboBox(
            master = self.frame_exam, 
            # variable=self.testlist,
            command=self.callback,
            values="",
            text_color='#000000',
            fg_color='#F0F0F0',
            button_color='#0080FF',
            width=120,height=40,
            state="readonly"
            )
        # self.input_exam.bind("<<ComboboxSelected>>", self.callback)
        self.label_totalcount = ctk.CTkLabel(
            self.frame_counter, 
            text = "總數",
            text_color= "#000000",
            font=('微軟正黑體',20,'bold'),
            fg_color='#CCCCCC',
            width=120,height=50,
            corner_radius=8
            )
        self.input_totalcount = ctk.CTkLabel(
            self.frame_counter,
            width=80,height=50,
            # text = '000' ,
            textvariable = self.totalcount,
            font=('微軟正黑體',24,'bold'),
            text_color='#3399CC',
            fg_color='#323232',
            corner_radius=8
            )
        #設置輸入端即開始測驗按鈕
        self.start = ctk.CTkButton(
        self.frame_exam,
        text="開始測驗",
        font=('微軟正黑體',18,'bold'),
        command=self.tstart,
        width=120,height=40,
        bg_color='#FCFCFC',
        fg_color='#0080FF',
            )
        self.clear_btn = ctk.CTkButton(
        self.frame_exam,
        text="清除",
        font=('微軟正黑體',18,'bold'),
        command=self.clear,
        width=120,height=40,
        bg_color='#FCFCFC',
        fg_color='#0080FF',
            )
        #設置上下數的switch
        self.switch_var = ctk.StringVar(value="off")
        self.switch = ctk.CTkSwitch(
            self.frame_function,
            text="上/下數",
            font=('微軟正黑體',18,),
            text_color="#000000",
            variable=self.switch_var,
            command=self.updown,
            onvalue="on",
            offvalue="off",
            button_hover_color="#46A3FF",
            progress_color="#46A3FF"
        )
        #離開計數畫面，回到主畫面
        self.backbtn = ctk.CTkButton(
            self.frame_function,
            text="返回",
            font=('微軟正黑體',18,),
            command= lambda: self.back(oldmaster),
            width=100,height=30,
            bg_color='#FCFCFC',
            fg_color='#0080FF',
        )
        ##計數器歸0
        self.zero = ctk.CTkButton(
            self.frame_function,
            text="歸0",
            font=('微軟正黑體',18,),
            command= self.tozero,
            width=100,height=30,
            bg_color='#FCFCFC',
            fg_color='#0080FF',
        )
        self.btn_sendtest = ctk.CTkButton(
            self.frame_function,
            text="交卷",
            font=('微軟正黑體',18,),
            command= self.sendtest,
            width=100,height=30,
            bg_color='#FCFCFC',
            fg_color='#FF9999',
        )
        self.btn_set = ctk.CTkButton(
            self.frame_function,
            text="參數設定",
            font=('微軟正黑體',18,),
            command= self.set_btn,
            width=100,height=30,
            bg_color='#FCFCFC',
            fg_color='#FF2D2D',
        )
        #建立db(counter每個細胞)
        self.cell_A = tk.IntVar()
        self.cell_A.set(0)
        self.cell_B = tk.IntVar()
        self.cell_B.set(0)
        self.cell_C = tk.IntVar()
        self.cell_C.set(0)
        self.cell_D = tk.IntVar()
        self.cell_D.set(0)
        self.cell_E = tk.IntVar()
        self.cell_E.set(0)
        self.cell_F = tk.IntVar()
        self.cell_F.set(0)
        self.cell_G = tk.IntVar()
        self.cell_G.set(0)
        self.cell_H = tk.IntVar()
        self.cell_H.set(0)
        self.cell_I = tk.IntVar()
        self.cell_I.set(0)
        self.cell_J = tk.IntVar()
        self.cell_J.set(0)
        self.cell_K = tk.IntVar()
        self.cell_K.set(0)
        self.cell_L = tk.IntVar()
        self.cell_L.set(0)
        self.cell_M = tk.IntVar()
        self.cell_M.set(0)
        self.cell_N = tk.IntVar()
        self.cell_N.set(0)
        self.cell_O = tk.IntVar()
        self.cell_O.set(0)
        self.cell_P = tk.IntVar()
        self.cell_P.set(0)
        self.cell_Q = tk.IntVar()
        self.cell_Q.set(0)
        self.cell_R = tk.IntVar()
        self.cell_R.set(0)
        self.cell_S = tk.IntVar()
        self.cell_S.set(0)
        self.cell_T = tk.IntVar()
        self.cell_T.set(0)

        ##建立按鍵與細胞連結
        #字典結構:{key=輸入器輸入字母:value=[x,y,細胞全名,細胞簡稱,連結計數器,百分比]
        self.keybordmatrix = {
        'I':[3,0,'test','TEST',self.cell_A,],
        '4':[3,1,'plasma cell','PLASMA',self.cell_B,],
        '3':[3,2,'abnormal lympho','AB-LYM',self.cell_C,],
        '2':[3,3,'megakery cell','MEGAKA',self.cell_D,],
        '1':[3,4,'nRBC','N-RBC',self.cell_E,],
        'T':[2,0,'abnormal monocyte','AB-MONO',self.cell_F,],
        'R':[2,1,'blast','BLAST',self.cell_G,],
        'E':[2,2,'metamyelocyte','META',self.cell_H,],
        'W':[2,3,'eosinopil','EOSIN',self.cell_I,],
        'Q':[2,4,'plasma cytoid','PLACYT',self.cell_J,],
        'G':[1,0,'promonocyte','PROMO',self.cell_K,],
        'F':[1,1,'promyelocyte','PROMY',self.cell_L,],
        'D':[1,2,'band neutropil','BAND',self.cell_M,],
        'S':[1,3,'basopil','BASO',self.cell_N,],
        'A':[1,4,'atypical lymphocyte','AT-LYM',self.cell_O,],
        'B':[0,0,'hypersegmented neutrophil','HYPERSEG',self.cell_P,],
        'V':[0,1,'myelocyte','MYELO',self.cell_Q,],
        'C':[0,2,'segmented neutrophil','SEG',self.cell_R,],
        'X':[0,3,'lymphocyte','LYM',self.cell_S,],
        'Z':[0,4,'monocyte','MONO',self.cell_T,]
        }
        
        global cell_grids
        cell_grids = [
            [3,0],[3,1],[3,2],[3,3],[3,4],
            [5,0],[5,1],[5,2],[5,3],[5,4],
            [7,0],[7,1],[7,2],[7,3],[7,4],
            [9,0],[9,1],[9,2],[9,3],[9,4],
        ]
        #建立百分比變數
        for key in self.keybordmatrix:
            self.keybordmatrix[key].append(tk.DoubleVar())
            self.keybordmatrix[key][5].set(0.0)
        #去除nRBC記錄至百分比中
        self.keybordmatrix['1'][5] = tk.IntVar()
        self.keybordmatrix['1'][5].set(0)
        #考片資訊
        self.info={
            "WBC":DoubleVar(),
            "RBC":DoubleVar(),
            "HB":DoubleVar(),
            "Hct":DoubleVar(),
            "MCV":DoubleVar(),
            "MCH":DoubleVar(),
            "MCHC":DoubleVar(),
            "RDW":DoubleVar(),
            "Plt":DoubleVar()
        }
        #考片初始值為0
        for key in self.info:
            self.info[key].set(0.0)
        # print(self.info)
        # self.percent_doublevar= []
        # for y in range(0,20):
        #     # percent_name = tk.DoubleVar('self.percent_%d'%(y))
        #     self.percent_doublevar.append(tk.DoubleVar())
        #     self.percent_doublevar[y].set(1.0)
        # print(self.percent_doublevar[y].get())
        #考片標題(CBC DATA,檢驗名稱,檢驗值)
        self.info_title = ctk.CTkLabel(self.frame_info,text = 'CBC DATA:',
            text_color='#000000',
            font=('微軟正黑體',18),
            bg_color='#BEBEBE',
            fg_color='#BEBEBE',
            )
        title2=["檢驗名稱","檢驗值"]
        for i in range(0,2):
            self.info_title2 = ctk.CTkEntry(
                self.frame_info,
                width=80,height=20,
                bg_color='#BEBEBE',
                fg_color='#CCCCCC',
                # text="",
                font=('微軟正黑體',14),
                text_color="#000000"
                )
            self.info_title2.grid(row=1, column=i)
            self.info_title2.insert(ctk.END, title2[i])
            self.info_title2.configure(state='disable')
    #數細胞的矩陣
    # def matrix(self):
        # celllist=[key for key in cell]
        # print(celllist)
        keylist = [key for key in self.keybordmatrix]
        celllist2=[]
        for key in self.keybordmatrix:
            celllist2.append(self.keybordmatrix[key][3])
        # print(celllist2)
        for i in range(2,10):
            for j in range(0,10):
                # print(j)
                if i % 2 == 0:
                    if j % 2 == 0 and int(5*(i / 2 - 1) + (j / 2)) != 0: 
                        #細胞數標籤
                        M1 = ctk.CTkLabel(
                            self.frame_counter,
                            text=celllist2[int(5*(i / 2 - 1) + (j / 2))],
                            font=('微軟正黑體',14),
                            text_color="#000000",
                            fg_color='#CCCCCC',
                            width=60,height=25,
                            corner_radius=8
                            )
                        M1.grid(row=i, column=j,padx=5,pady=4,columnspan=2,sticky='nsew')
                        M1.grid_propagate(0)
                else:
                    if j % 2 == 1 and int(5*(i-3)/2 + (j-1)/2) != 0: 
                        y = int(5*(i-3)/2 + (j-1)/2)
                        # yy = self.percent_doublevar[y]
                        yy = self.keybordmatrix[keylist[y]][5]
                        # print(yy)
                        #細胞百分比標籤
                        M2 = ctk.CTkLabel(
                            self.frame_counter,
                            textvariable= yy,
                            font=('微軟正黑體',14),
                            text_color="#4F4F4F",
                            fg_color='#C4E1FF',
                            width=55,height=25,
                            corner_radius=8
                            )
                        M2.grid(row=i,column=j,padx=10,pady=4)
                        M2.grid_propagate(0)
        #每種細胞的數值標籤
        cell_type_01 = ctk.CTkLabel(
            self.frame_counter,
            text = "",
            font=('微軟正黑體',18),
            fg_color='#F0F0F0',
            width=80,height=80,
            corner_radius=8
            )
        cell_type_01.grid(row=2,rowspan=2, column=0,columnspan=2,padx=2)
        cell_type_01.grid_propagate(0)
        cell_type_02 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_B,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,
            corner_radius=8
            )
        cell_type_02.grid(row=3, column=2,padx=2)
        cell_type_02.grid_propagate(0)
        cell_type_03 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_C,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_03.grid(row=3, column=4,padx=2)
        cell_type_03.grid_propagate(0)
        cell_type_04 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_D,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_04.grid(row=3, column=6,padx=2)
        cell_type_04.grid_propagate(0)
        cell_type_05 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_E,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_05.grid(row=3, column=8,padx=2)
        cell_type_05.grid_propagate(0)
        cell_type_06 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_F,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_06.grid(row=5, column=0,padx=2)
        cell_type_06.grid_propagate(0)
        cell_type_07 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_G,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_07.grid(row=5, column=2,padx=2)
        cell_type_07.grid_propagate(0)
        cell_type_08 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_H,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_08.grid(row=5, column=4,padx=2)
        cell_type_08.grid_propagate(0)
        cell_type_09 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_I,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_09.grid(row=5, column=6,padx=2)
        cell_type_09.grid_propagate(0)
        cell_type_10 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_J,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_10.grid(row=5, column=8,padx=2)
        cell_type_10.grid_propagate(0)
        cell_type_11 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_K,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_11.grid(row=7, column=0,padx=2)
        cell_type_11.grid_propagate(0)
        cell_type_12 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_L,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_12.grid(row=7, column=2,padx=2)
        cell_type_12.grid_propagate(0)
        cell_type_13 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_M,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_13.grid(row=7, column=4,padx=2)
        cell_type_13.grid_propagate(0)
        cell_type_14 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_N,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_14.grid(row=7, column=6,padx=2)
        cell_type_14.grid_propagate(0)
        cell_type_15 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_O,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_15.grid(row=7, column=8,padx=2)
        cell_type_15.grid_propagate(0)
        cell_type_16 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_P,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_16.grid(row=9, column=0,padx=2)
        cell_type_16.grid_propagate(0)
        cell_type_17 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_Q,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_17.grid(row=9, column=2,padx=2)
        cell_type_17.grid_propagate(0)
        cell_type_18 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_R,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_18.grid(row=9, column=4,padx=2)
        cell_type_18.grid_propagate(0)
        cell_type_19 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_S,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_19.grid(row=9, column=6,padx=2)
        cell_type_19.grid_propagate(0)
        cell_type_20 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_T,
            text_color="#FFFFFF",
            font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_20.grid(row=9, column=8,padx=2)
        cell_type_20.grid_propagate(0)
    # def gui_arrang(self):
        self.frame_exam.grid(row=0, column=0,columnspan=2,pady=20,sticky=tk.W)
        self.frame_time.grid(row=1, column=0,padx=3,ipadx=20)
        self.frame_info.grid(row=2, column=0,padx=10)
        self.frame_counter.grid(row=1, column=1,rowspan=2,columnspan=2,pady=30)
        self.frame_counter.grid_propagate(0)
        self.frame_function.grid(row=0, column=2)
        self.time.grid(row=0, column=1,)
        self.label_exam1.grid(row=0, column=0,sticky='e')
        self.input_exam1.grid(row=0, column=1,pady=10)
        self.label_exam2.grid(row=1, column=0,sticky='e')
        self.input_exam2.grid(row=1, column=1,pady=10)
        self.label_totalcount.grid(row=0, column=0,columnspan=10,padx=20,pady=5)
        self.input_totalcount.grid(row=1, column=0,columnspan=10,padx=20,)
        self.info_title.grid(row=0,column=0,columnspan=2,padx=20)
        # self.input.grid(row=2, column=2)
        self.start.grid(row=0, column=2,padx=20)
        self.clear_btn.grid(row=1, column=2,padx=20)
        self.switch.grid(row=0, column=0,columnspan=2,pady=10)
        self.zero.grid(row=1, column=0,pady=5)
        self.backbtn.grid(row=2, column=0,pady=5)
        self.btn_sendtest.grid(row=1, column=1,padx=2,pady=5)
        self.btn_set.grid(row=2, column=1,padx=2,pady=5)
    #考片CBC資訊(infocreate)
        infolist =[key for key in self.info.keys()]
        for key in range(0,len(infolist)):
            M3 = ctk.CTkEntry(
                self.frame_info,
                width=8,
                bg_color='#BEBEBE',
                fg_color='#F0F0F0',
                font=('微軟正黑體',12),
                text_color="#000000"
                )
            M3.insert(tk.END,infolist[key])
            M3.grid(row = key + 2, column = 0,sticky='nsew')
            M3.configure(state='disabled')
        m = 2
        for values in self.info.values():
            M4 = ctk.CTkLabel(
                self.frame_info,
                width=8,
                bg_color='#BEBEBE',
                fg_color='#F0F0F0',
                textvariable= values,
                font=('微軟正黑體',14),
                text_color="#000000"
                )
            M4.grid(row = m, column = 1,sticky='nsew')
            m+=1
        MB = ctk.CTkLabel(
                self.frame_info,
                width=2,
                bg_color='#BEBEBE',
                fg_color='#BEBEBE',
                text= "",
                # text_color="#F0F0F0"
                )
        MB.grid(row = 11, column = 0,columnspan=2,pady=5)
    
    def update_clock(self):
        if not self.doTick:
            return
        a = datetime.now()
        e = str(a - self.open).split('.', 2)[0]
        now = e
        # print(e.strptime("%H:%M:%S"))
        self.time.configure(text=now)
        self.master.after(1000, self.update_clock)
    def stop_clock(self):
        self.doTick = False

    
    #函式判別上/下數、判別是否數錯方向
    def pending(self,event):
        #聲音判別:數到一百顆會diang
        if (self.totalcount.get() + 1) % 100 == 0:
            self.input_totalcount.configure(text_color='red')
            # self.master.bell()
            winsound.Beep(freq_1, duration_1)
        else:
            winsound.Beep(freq, duration)
            self.input_totalcount.configure(text_color='#3399CC')
        #檢查是否有小於0的細胞
        # idiv_cell=[]
        idiv_cell2=[]
        # for h in cell.values():
        #     idiv_cell.append(h.get())
        # print(idiv_cell)
        for key in self.keybordmatrix:
            kk = self.keybordmatrix[key][4]
            idiv_cell2.append(kk.get())
        # print(idiv_cell2)
        result = all(ele >= 0 for ele in idiv_cell2)
        # print(result)
        if str.upper(event.char)=="I":
            self.switch.toggle()
        else:
            x = self.keybordmatrix[str.upper(event.char)][4].get()
        if self.switch_var.get() == "off":
            return self.add_count(event)
        else: 
            if x ==0 or self.totalcount.get() <=0 or result == False:
                tk.messagebox.showerror(title='總數問題', message='細胞已歸0!請切換至上數模式!')
            else:
                return self.minus_count(event)
            
    #函式上數
    def add_count(self,event):
        # x = cell[self.keyboard[event.char]].get()
        x = self.keybordmatrix[str.upper(event.char)][4].get()
        x += 1
        # cell[self.keyboard[event.char]].set(x)
        self.keybordmatrix[str.upper(event.char)][4].set(x)
        y = self.totalcount.get()
        if str.upper(event.char) !='1':
            y+=1
        else:
            pass
        # print(y)
        self.totalcount.set(y)
        tal = self.totalcount.get()
        #新增百分比更新
        for key in self.keybordmatrix:
            if key != '1':
                val = self.keybordmatrix[key][4].get()
                percent = round(val / tal * 100,2)
                self.keybordmatrix[key][5].set(percent)
            else:
                val = self.keybordmatrix[key][4].get()
                self.keybordmatrix[key][5].set(val)
    
    #函式下數
    def minus_count(self,event):
        # x = cell[self.keyboard[event.char]].get()
        x = self.keybordmatrix[str.upper(event.char)][4].get()
        if x == 0 :
            return self.pending
        x -= 1
        # cell[self.keyboard[event.char]].set(x)
        self.keybordmatrix[str.upper(event.char)][4].set(x)
        y = self.totalcount.get()
        # print(y)
        if str.upper(event.char) !='1':
            y-=1
        else:
            pass
        # print(y)
        self.totalcount.set(y)
        tal = self.totalcount.get()
        #新增百分比更新
        for key in self.keybordmatrix:
            if key != '1':
                val = self.keybordmatrix[key][4].get()
                percent = round(val / tal * 100,2)
                self.keybordmatrix[key][5].set(percent)
            else:
                val = self.keybordmatrix[key][4].get()
                self.keybordmatrix[key][5].set(val)

    #上下數switch
    def updown(self,event):
        mode =self.switch_var.get()
        if mode =="off":
            mode =="on"
        else:
            mode =="off"
        self.switch_var.set(mode)
    
    
    #combobox雙層列表
    def callback_year(self,event):  
        testno = self.input_exam1.get()
        fliter = (self.rawdata['year'] == testno)
        self.testlist = self.rawdata[fliter]['ID']
        self.input_exam2.configure(values=self.testlist)
        for key in self.info:
            self.info[key].set(0.0)

    def callback(self,event):  
        testyear = self.input_exam1.get()
        testno = self.input_exam2.get()
        filter1 = (self.rawdata['year'] == testyear)
        filter2 = (self.rawdata['ID'] == testno)
        rawdata = self.rawdata[filter1 & filter2]
        rawdata = rawdata.iloc[0]['rawdata']
        for key,value in rawdata.items():
            self.info[key].set(rawdata[key])
        #轉換考題的時候要把計數規0(tozero)
    
    def tozero(self):
        ##歸0 counter
        for value in self.keybordmatrix.values():
            value[4].set(0)
            value[5].set(0)
        ##歸0 總數
        self.totalcount.set(0)
    def tstart(self):
        ##檢查是否有考片
        testyear = self.input_exam1.get()
        testno = self.input_exam2.get()
        if testyear=="" or testno=="":
            tk.messagebox.showerror(title='土城醫院檢驗科',message="尚未選擇考片!選擇後再開始測驗!")
            return
        if tk.messagebox.askyesno(title='土城醫院檢驗科',message="""確定要開始測驗?
考片年份:%s
考片編號%s"""%(testyear,testno)):
            self.open = datetime.now()
            self.doTick = True
            self.start.configure(state='disabled')
            self.zero.configure(state='disabled')
            self.input_exam1.configure(state='disabled')
            self.input_exam2.configure(state='disabled')
            self.clear_btn.configure(state='disabled')
            self.tozero()
            self.update_clock()
        else:
            return
    def sendtest(self):
        if tk.messagebox.askyesno(title='土城長庚檢驗科', message='確定要交卷?', ):
            testyear = self.input_exam1.get()
            testno = self.input_exam2.get()
            stamp = datetime.now()
            stamp = datetime.strftime(stamp,'%Y/%m/%d %H:%M:%S')
            # print(testno,Baccount,stamp)
            ##dict(ans)
            ans = {}
            for values in self.keybordmatrix.values():
                if values[2] =="test":
                    pass
                else:
                    ans[values[2]] = values[4].get()
            ##建立存入dict
            savedic = {}
            savedic["year"] = testyear
            savedic["ID"] = testno
            savedic["personID"] = "henry423"
            # savedic["personID"] = Baccount
            savedic["timestamp"] = stamp
            savedic["Ans"] = ans
            # print (savedic)
            ##開啟rawdata.json
            jsonfile = open('testdata/rawdata.json','rb')
            a = json.load(jsonfile)
            rawdata = a["blood"]
            rawdata.append(savedic)
            # print(ml)
            # with open('testdata/rawdata.json','w') as r:
            #     json.dump(a,r)
            #     r.close()
            tk.messagebox.showinfo(title='土城長庚檢驗科', message="交卷成功!")
            self.tozero()
            self.start.configure(state=tk.NORMAL)
            self.zero.configure(state=tk.NORMAL)
            self.input_exam1.configure(state='readonly')
            self.input_exam2.configure(state='readonly')
            self.clear_btn.configure(state='normal')
            self.stop_clock()
            
        else:
            return 
    def clear(self):
        if tk.messagebox.askyesno(title="土城醫院檢驗科",message="確定要清除所選擇的考片?"):
            self.input_exam1.set("")
            self.input_exam2.set("")
            for value in self.info.values():
                value.set(0.0)
        else:
            return
    def set_btn(self):
        def OK():
            pass
        def CANCEL():
            pass
        
        self.newWindow = ctk.CTkToplevel()
        self.newWindow.config(background="#F0F0F0")
        self.newWindow.title("設定輸入參數")
        self.newtitle = ctk.CTkLabel(self.newWindow,text="輸入參數設定",font=('微軟正黑體',18),text_color="#000000",height=30,bg_color="#F0F0F0",fg_color="#F0F0F0")
        self.newtitle.grid(row=0,column=0,columnspan=5,sticky="nsew")

        for key,value in self.keybordmatrix.items():
            #上方座標標籤
            inlabel = "(%d,%d) key: %s"%(value[0],value[1],key)
            matrixlabel=ctk.CTkLabel(
                self.newWindow,
                text=inlabel,
                text_color="#000000",
                bg_color="#F0F0F0",
                fg_color="#CCCCCC",
                corner_radius=5,
                width=20,height=30
            )
            #下方按鈕
            matrixbtn=ctk.CTkLabel(
                self.newWindow,
                text=value[3],
                text_color="#000000",
                bg_color="#F0F0F0",
                fg_color="#97CBFF",
                corner_radius=5,
                width=80,height=60
            )
            lrow = 7 - (value[0] * 2)
            nrow = (4 - value[0]) * 2
            ncolumn = value[1]
            matrixlabel.grid(row=lrow,column=ncolumn,padx=10,pady=5)
            matrixbtn.grid(row=nrow,column=ncolumn,padx=10,pady=5)
        ok_btn = ctk.CTkButton(
            self.newWindow,
            text="修改",
            command= OK,
            bg_color="#F0F0F0",
            fg_color="#FF8000",
            corner_radius=5,
            width=80,height=30
        )
        ok_btn.grid(row=9,column=1,pady=5)
        cancel_btn = ctk.CTkButton(
            self.newWindow,
            text="取消",
            command= CANCEL,
            bg_color="#F0F0F0",
            fg_color="#FF8000",
            corner_radius=5,
            width=80,height=30
        )
        cancel_btn.grid(row=9,column=3,pady=5)
    def back(self,oldmaster):
        # if tk.messagebox.askyesno(title='土城長庚檢驗科', message='確定要離開計數畫面，返回主畫面嗎?', ):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return
        

def main():  
    # 初始化物件  
    root = ctk.CTk()
    C = Count(root)  
    # 進行佈局
    # C.matrix() 
    # C.gui_arrang()
    # C.infocreate()
    # 主程式執行
    root.mainloop()  
  
if __name__ == "__main__":  
    main()