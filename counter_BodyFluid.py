import tkinter as tk
from tkinter import RIDGE, DoubleVar, StringVar, ttk, messagebox,IntVar
from turtle import bgcolor, width
from unicodedata import name
from ttkbootstrap import Style
import customtkinter  as ctk
import json
import winsound
duration = 50  # millisecond
freq = 600  # Hz


class Count(ctk.CTk):    #建立計數器

    def __init__(self):
        #建立主事窗前先取得初始資料
        jsonfile = open('testdata\data.json','rb')
        rawdata = json.load(jsonfile)
        self.rawdata = rawdata["blood"]
        self.testlist = []
        for item in self.rawdata:
            self.testlist.append(item["ID"])
        testno = self.testlist.index('B001')
        rawdata = self.rawdata[testno]["rawdata"]
        # 建立主視窗,用於容納其它元件  
        super().__init__()
        # self.root = ctk.CTk()
        self.title("細胞計數器")
        self.geometry('900x600')
        style = Style(theme='flatly')
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.config(background='#FCFCFC')
        self.bind('<Key>', self.pending)
        #設置框架_考核片選擇
        self.frame_exam = ctk.CTkFrame(self)
        self.frame_exam.configure(bg_color='#FCFCFC',fg_color='#FCFCFC',)
        #設置框架_考核片資訊
        self.frame_info = ctk.CTkFrame(self)
        self.frame_info.configure(bg_color='#FCFCFC',fg_color='#BEBEBE',corner_radius=8)
        #設置框架_counter
        self.frame_counter = ctk.CTkFrame(self)
        self.frame_counter.configure(bg_color='#FCFCFC',fg_color='#F0F0F0',width=600,height=400,corner_radius=8)
        self.frame_counter.propagate(0)
        #設置框架_功能區
        self.frame_function = ctk.CTkFrame(self)
        self.frame_function.configure(bg_color='#FCFCFC',fg_color='#FCFCFC',)
        #設置細胞總數
        self.totalcount = tk.IntVar()
        self.totalcount.set(0)
        #設置標籤
        self.label_exam = ctk.CTkLabel(
            self.frame_exam, 
            text = "請選擇考核片:", 
            fg_color='#FCFCFC',
            text_font=('微軟正黑體',16),
            text_color="#000000",
            width=200
            )
        self.input_exam = ttk.Combobox(
            master = self.frame_exam, 
            values=self.testlist,
            width=15,height=40,
            state="readonly"
            )
        self.input_exam.bind("<<ComboboxSelected>>", self.callback)
        self.label_totalcount = ctk.CTkLabel(
            self.frame_counter, 
            text = "總數",
            text_color= "#000000",
            text_font=('微軟正黑體',20,'bold'),
            fg_color='#CCCCCC',
            width=120,height=50,
            corner_radius=8
            )
        self.input_totalcount = ctk.CTkLabel(
            self.frame_counter,
            width=80,height=50,
            # text = '000' ,
            textvariable = self.totalcount,
            text_font=('微軟正黑體',24,'bold'),
            text_color='#3399CC',
            fg_color='#323232',
            corner_radius=8
            )
        #設置輸入端即開始測驗按鈕
        # self.input = ctk.CTkEntry(
        #     self.frame_exam,
        #     width=80,height=20,
        #     bg_color='#AE0000',
        #     fg_color='#AE0000',
        #     )
        self.start = ctk.CTkButton(
        self.frame_exam,
        text="開始測驗",
        text_font=('微軟正黑體',18,'bold'),
        command=None,
        width=120,height=40,
        bg_color='#FCFCFC',
        fg_color='#0080FF',
            )
        #設置上下數的switch
        self.switch_var = ctk.StringVar(value="off")
        self.switch = ctk.CTkSwitch(
            self.frame_function,
            text="上/下數",
            text_font=('微軟正黑體',18,),
            text_color="#000000",
            variable=self.switch_var,
            command=self.updown,
            onvalue="on",
            offvalue="off",
            button_hover_color="#46A3FF",
            progress_color="#46A3FF"
        )
        #建立db
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

        # self.info=[
        #             ('WBC',16.7),
        #             ('RBC',4.65),
        #             ('HB',12.3),
        #             ('Hct',37.5),
        #             ('MCV',80.6),
        #             ('MCH',26.5),
        #             ('MCHC',32.8),
        #             ('RDW',13.3),
        #             ('Plt',301),
        #         ]
        #建立按鍵與細胞連結
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
        self.keyboard = {
            'a':'a',
            'b':'PLASMA',
            'c':'AB-LYM',
            'd':'MEGAKA',
            'e':'N-RBC',
            'f':'AB-MONO',
            'g':'BLAST',
            'h':'META',
            'i':'EOSIN',
            'j':'PLACYT',
            'k':'PROMO',
            'l':'PROMY',
            'm':'BAND',
            'n':'BASO',
            'o':'AT-LYM',
            'p':'HYPERSEG',
            'q':'MYELO',
            'r':'SEG',
            's':'LYM',
            't':'MONO',
        }
        global cell_grids
        cell_grids = [
            [3,0],[3,1],[3,2],[3,3],[3,4],
            [5,0],[5,1],[5,2],[5,3],[5,4],
            [7,0],[7,1],[7,2],[7,3],[7,4],
            [9,0],[9,1],[9,2],[9,3],[9,4],
        ]
        
        for key in self.keybordmatrix:
            self.keybordmatrix[key].append(tk.DoubleVar())
            self.keybordmatrix[key][5].set(0.0)
        #去除nRBC記錄至百分比中
        self.keybordmatrix['1'][5] = tk.IntVar()
        self.keybordmatrix['1'][5].set(0)

        self.info={
            "WBC":tk.DoubleVar(),
            "RBC":tk.DoubleVar(),
            "HB":tk.DoubleVar(),
            "Hct":tk.DoubleVar(),
            "MCV":tk.DoubleVar(),
            "MCH":tk.DoubleVar(),
            "MCHC":tk.DoubleVar(),
            "RDW":tk.DoubleVar(),
            "Plt":tk.DoubleVar()
        }
        for key in self.info:
            # self.info[key].set(rawdata[key])
            self.info[key].set(0.0)
        # print(self.info)
        # self.percent_doublevar= []
        # for y in range(0,20):
        #     # percent_name = tk.DoubleVar('self.percent_%d'%(y))
        #     self.percent_doublevar.append(tk.DoubleVar())
        #     self.percent_doublevar[y].set(1.0)
        # print(self.percent_doublevar[y].get())
        
        self.info_title = ctk.CTkLabel(self.frame_info,text = 'CBC DATA:',
            text_color='#000000',
            text_font=('微軟正黑體',18),
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
                text="",
                text_font=('微軟正黑體',12),
                text_color="#000000"
                )
            self.info_title2.grid(row=1, column=i)
            self.info_title2.insert(ctk.END, title2[i])
            self.info_title2.configure(state='disable')
    #數細胞的矩陣
    def matrix(self):
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
                    if j % 2 == 0: 
                        #欄位名稱
                        M1 = ctk.CTkLabel(
                            self.frame_counter,
                            text=celllist2[int(5*(i / 2 - 1) + (j / 2))],
                            text_font=('微軟正黑體',11),
                            text_color="#000000",
                            fg_color='#CCCCCC',
                            width=60,height=25,
                            corner_radius=8
                            )
                        M1.grid(row=i, column=j,padx=5,pady=4,columnspan=2,sticky='nsew')
                        M1.grid_propagate(0)
                else:
                    if j % 2 == 1: 
                        y = int(5*(i-3)/2 + (j-1)/2)
                        # yy = self.percent_doublevar[y]
                        yy = self.keybordmatrix[keylist[y]][5]
                        # print(yy)
                        M2 = ctk.CTkLabel(
                            self.frame_counter,
                            textvariable= yy,
                            text_font=('微軟正黑體',11),
                            text_color="#4F4F4F",
                            fg_color='#CCCCCC',
                            width=55,height=25,
                            corner_radius=8
                            )
                        M2.grid(row=i,column=j,padx=10,pady=4)
                        M2.grid_propagate(0)
        #每種細胞的數值標籤
        cell_type_01 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_A,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,
            corner_radius=8
            )
        cell_type_01.grid(row=3, column=0,padx=2)
        cell_type_01.grid_propagate(0)
        cell_type_02 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_B,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,
            corner_radius=8
            )
        cell_type_02.grid(row=3, column=2,padx=2)
        cell_type_02.grid_propagate(0)
        cell_type_03 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_C,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_03.grid(row=3, column=4,padx=2)
        cell_type_03.grid_propagate(0)
        cell_type_04 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_D,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_04.grid(row=3, column=6,padx=2)
        cell_type_04.grid_propagate(0)
        cell_type_05 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_E,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_05.grid(row=3, column=8,padx=2)
        cell_type_05.grid_propagate(0)
        cell_type_06 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_F,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_06.grid(row=5, column=0,padx=2)
        cell_type_06.grid_propagate(0)
        cell_type_07 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_G,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_07.grid(row=5, column=2,padx=2)
        cell_type_07.grid_propagate(0)
        cell_type_08 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_H,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_08.grid(row=5, column=4,padx=2)
        cell_type_08.grid_propagate(0)
        cell_type_09 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_I,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_09.grid(row=5, column=6,padx=2)
        cell_type_09.grid_propagate(0)
        cell_type_10 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_J,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_10.grid(row=5, column=8,padx=2)
        cell_type_10.grid_propagate(0)
        cell_type_11 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_K,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_11.grid(row=7, column=0,padx=2)
        cell_type_11.grid_propagate(0)
        cell_type_12 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_L,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_12.grid(row=7, column=2,padx=2)
        cell_type_12.grid_propagate(0)
        cell_type_13 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_M,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_13.grid(row=7, column=4,padx=2)
        cell_type_13.grid_propagate(0)
        cell_type_14 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_N,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_14.grid(row=7, column=6,padx=2)
        cell_type_14.grid_propagate(0)
        cell_type_15 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_O,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_15.grid(row=7, column=8,padx=2)
        cell_type_15.grid_propagate(0)
        cell_type_16 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_P,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_16.grid(row=9, column=0,padx=2)
        cell_type_16.grid_propagate(0)
        cell_type_17 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_Q,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_17.grid(row=9, column=2,padx=2)
        cell_type_17.grid_propagate(0)
        cell_type_18 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_R,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_18.grid(row=9, column=4,padx=2)
        cell_type_18.grid_propagate(0)
        cell_type_19 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_S,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_19.grid(row=9, column=6,padx=2)
        cell_type_19.grid_propagate(0)
        cell_type_20 = ctk.CTkLabel(
            self.frame_counter,
            textvariable = self.cell_T,
            text_font=('微軟正黑體',18),
            fg_color='#323232',
            width=40,corner_radius=8
            )
        cell_type_20.grid(row=9, column=8,padx=2)
        cell_type_20.grid_propagate(0)
    
    #函式判別上/下數、判別是否數錯方向
    def pending(self,event):
        #聲音判別:數到一百顆會diang
        if (self.totalcount.get() + 1) % 100 == 0:
            self.input_totalcount.configure(text_color='red')
            self.bell()
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
    def updown(self):
        mode =self.switch_var.get()
        mode =="off"
        self.switch_var.set(mode)
    #考片CBC資訊
    def infocreate(self):
        # self.entry.delete(0, ctk.END)
        infolist =[key for key in self.info.keys()]
        for key in range(0,len(infolist)):
            M3 = tk.Label(
                self.frame_info,
                relief="ridge",borderwidth=2,
                width=8,
                background='#BEBEBE',
                foreground='#F0F0F0',
                text= infolist[key],
                font=('微軟正黑體',12),
                # text_color="#000000"
                )
            M3.grid(row = key + 2, column = 0,sticky='nsew')
        m = 2
        for values in self.info.values():
            M4 = tk.Label(
                self.frame_info,
                relief="ridge",borderwidth=2,
                width=8,
                background='#BEBEBE',
                foreground='#F0F0F0',
                textvariable= values,
                font=('微軟正黑體',12),
                # text_color="#000000"
                )
            M4.grid(row = m, column = 1,sticky='nsew')
            m+=1
        MB = ctk.CTkLabel(
                self.frame_info,
                width=2,
                background='#BEBEBE',
                foreground='#F0F0F0',
                text= "",
                # text_color="#000000"
                )
        MB.grid(row = 11, column = 0,columnspan=2,)

    def gui_arrang(self):
        self.frame_exam.grid(row=0, column=0,columnspan=2,pady=20,sticky=tk.W)
        self.frame_info.grid(row=1, column=0,padx=20)
        self.frame_counter.grid(row=1, column=1,columnspan=2,pady=30)
        self.frame_counter.grid_propagate(0)
        self.frame_function.grid(row=0, column=2)
        self.label_exam.grid(row=0, column=0)
        self.input_exam.grid(row=0, column=1)
        self.label_totalcount.grid(row=0, column=0,columnspan=10,padx=20,pady=5)
        self.input_totalcount.grid(row=1, column=0,columnspan=10,padx=20,)
        self.info_title.grid(row=0,column=0,columnspan=2,padx=20)
        # self.input.grid(row=2, column=2)
        self.start.grid(row=0, column=2,padx=20)
        self.switch.grid(row=0, column=0,pady=20)
    def callback(self,event):  #combobox雙層列表
        testno = self.input_exam.get()
        testno = self.testlist.index(testno)
        rawdata = self.rawdata[testno]["rawdata"]
        # print(rawdata)
        for key,value in rawdata.items():
            self.info[key].set(rawdata[key])
def main():  
    # 初始化物件  
    C = Count()  
    # 進行佈局
    C.matrix() 
    C.gui_arrang()
    C.infocreate()
    # 主程式執行
    C.mainloop()  
  
if __name__ == "__main__":  
    main()