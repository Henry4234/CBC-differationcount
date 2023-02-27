import os
from textwrap import fill
import tkinter as tk
from tkinter import ttk
from tkinter import RIDGE, DoubleVar, StringVar, ttk, messagebox,IntVar, filedialog
from turtle import bgcolor, width
import customtkinter  as ctk
import json
import basedesk, basedesk_admin
from PIL import Image
import pandas as pd
import json
##取得登入帳號資訊
def getaccount(acount):
    global Baccount
    Baccount = str(acount)
    # print(Baccount)
    return None
##ctk無法有效支援treeview，所以使用entery + loop解決
def createtable(master,totalrows,totalcolumns,lst):
    global entrylst
    entrylst = []
    for i in range(totalrows):
        for j in range(totalcolumns):
            TB = tk.Entry(master, width=9, fg='blue',
                            font=('Arial',10,))
            entrylst.append(TB)
            if i < 14:
                TB.grid(row=i+2, column=j)
                TB.insert(tk.END, lst[i][j])
            elif i >= 14 and i < 28:
                TB.grid(row=i-12, column=j+2)
                TB.insert(tk.END, lst[i][j])
            else:
                TB.grid(row=i-24, column=j+4)
                TB.insert(tk.END, lst[i][j])
            TB.configure(state="disabled")

class Import:

    def __init__(self,master,oldmaster=None):
        #確認輸入類型
        global testtype
        testtype=""
        self.master = master
        super().__init__()
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.back(oldmaster))        
        ctk.set_default_color_theme("blue")
        self.master.title("考題匯入")
        self.master.geometry("1100x600")
        # self.master.config(bg='#FFEEDD')
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        
        ##放入背景
        canvas = ctk.CTkCanvas(
            self.master,
            bg = "#FFEEDD",
            height = 600,
            width = 1100,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )
        canvas.place(x = 0, y = 0)
        #匯入標題
        self.welcome = ctk.CTkLabel(
            self.master,
            text="考題匯入:",
            bg_color='#FFEEDD',
            font=('微軟正黑體',40),
            text_color='#000000',
            fg_color="#FFEEDD",
            width=240
        )
        self.welcome.pack()
        self.welcome.place(relx=0.5,rely=0.05,anchor=tk.CENTER)
        ###建立分頁窗格
        self.tab = ctk.CTkTabview(
            self.master,
            fg_color="#FFDCB9",
            bg_color="#FFEEDD",
            segmented_button_selected_color="#FF6600",
            width=250,height=400
            )
        self.tab.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
        self.tab.add("血液")
        self.tab.add("尿液")
        self.tab.add("體液")
        #建立血液中分頁的框架
        self.labelframe_1 = ctk.CTkFrame(
            self.tab.tab("血液"),
            fg_color="#FFDCB9",
            bg_color="#FFDCB9",
            width=250,height=400
            )
        self.labelframe_1.pack(side="top",fill='both')
        self.labelframe_1.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        #建立尿液中分頁的框架
        self.labelframe_2 = ctk.CTkFrame(
            self.tab.tab("尿液"),
            fg_color="#FFDCB9",
            bg_color="#FFDCB9",
            width=250,height=400
            )
        self.labelframe_2.pack(side="top",fill='both')
        self.labelframe_2.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        #建立體液中分頁的框架
        self.labelframe_3 = ctk.CTkFrame(
            self.tab.tab("體液"),
            fg_color="#FFDCB9",
            bg_color="#FFDCB9",
            width=250,height=400
            )
        self.labelframe_3.pack(side="top",fill='both')
        self.labelframe_3.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        ##匯入成功不成功的框架
        self.labelframe_ok = ctk.CTkFrame(
            self.master,
            fg_color="#8CEA00",
            bg_color="#FFEEDD",
            width=250,height=400,
            corner_radius=8
            )
        self.labelframe_ok.pack()
        self.labelframe_nok = ctk.CTkFrame(
            self.master,
            fg_color="#FF2D2D",
            bg_color="#FFEEDD",
            width=250,height=400,
            corner_radius=8
            )
        self.labelframe_nok.pack()
        ##驗證成功不成功的frame
        self.labelframe_vfok = ctk.CTkFrame(
            self.master,
            fg_color="#8CEA00",
            bg_color="#FFEEDD",
            width=280,height=500,
            corner_radius=8
            )
        self.labelframe_vfok.pack()
        self.labelframe_vfnok = ctk.CTkFrame(
            self.master,
            fg_color="#FF2D2D",
            bg_color="#FFEEDD",
            width=280,height=500,
            corner_radius=8
            )
        self.labelframe_vfnok.pack()
        # self.labelframe_ok.pack(side="top",fill='both')
        # self.labelframe_ok.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        
        # 匯入excel照片及建立tab中按鈕
        self.excel = ctk.CTkImage(Image.open("assets\logoexcel.png"),size=(40,40))
        self.correctcheck = ctk.CTkImage(Image.open("assets\check.png"),size=(60,60))
        self.wrongcheck = ctk.CTkImage(Image.open("assets\\ncheck.png"),size=(60,60))
        self.dbimage = ctk.CTkImage(Image.open("assets\logodb.png"),size=(80,80))
        self.btn_1 = ctk.CTkButton(
            self.labelframe_1,
            text="請選擇檔案...",
            image =self.excel,
            compound="bottom",
            anchor=tk.CENTER,
            command=self.selectfile,
            fg_color="#FFDCB9",
            bg_color='#FFDCB9', 
            hover_color="#FF9224",
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,height=400
        )
        self.btn_1.pack(fill='both')
        self.btn_1.grid(row=0,column=0,sticky='nsew')
        self.btn_2 = ctk.CTkButton(
            self.labelframe_2,
            text="請選擇檔案...",
            image =self.excel,
            compound="bottom",
            anchor=tk.CENTER,
            command=self.selectfile_urin,
            fg_color="#FFDCB9",
            bg_color='#FFDCB9',
            hover_color="#FF9224",
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,height=400
        )
        self.btn_2.pack(fill='both')
        self.btn_2.grid(row=0,column=0,sticky='nsew')
        self.btn_3 = ctk.CTkButton(
            self.labelframe_3,
            text="請選擇檔案...",
            image =self.excel,
            compound="bottom",
            anchor=tk.CENTER,
            command=self.selectfile_bodyfluid,
            fg_color="#FFDCB9",
            bg_color='#FFDCB9',
            hover_color="#FF9224",
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,height=400
        )
        self.btn_3.pack(fill='both')
        self.btn_3.grid(row=0,column=0,sticky='nsew')
        #建立匯入成功圖示
        self.label_ok = ctk.CTkLabel(
            self.labelframe_ok,
            text="匯入成功:",
            bg_color='#FFDCB9',
            font=('微軟正黑體',40),
            text_color='#000000',
            fg_color="#8CEA00",
            # width=250,height=400
        )
        self.label_ok.place(relx=0.5,rely=0.4,anchor=tk.CENTER)
        self.image_ok = ctk.CTkLabel(
            self.labelframe_ok,
            image =self.correctcheck,
            text="",
            bg_color='#FFDCB9',
            fg_color="#8CEA00",
        )
        self.image_ok.place(relx=0.5,rely=0.6,anchor=tk.CENTER)
        #建立匯入失敗圖示
        self.label_nok = ctk.CTkLabel(
            self.labelframe_nok,
            text="匯入失敗...",
            bg_color='#FFDCB9',
            font=('微軟正黑體',40),
            text_color='#000000',
            fg_color="#FF2D2D",
            # width=250,height=400
        )
        self.label_nok.place(relx=0.5,rely=0.4,anchor=tk.CENTER)
        self.image_nok = ctk.CTkLabel(
            self.labelframe_nok,
            image =self.wrongcheck,
            text="",
            bg_color='#FFDCB9',
            fg_color="#FF2D2D",
        )
        self.image_nok.place(relx=0.5,rely=0.6,anchor=tk.CENTER)
        
        #建立驗證成功圖示
        self.label_vfok = ctk.CTkLabel(
            self.labelframe_vfok,
            text="驗證成功:",
            bg_color='#FFDCB9',
            font=('微軟正黑體',40),
            text_color='#000000',
            fg_color="#8CEA00",
            # width=250,height=400
        )
        self.label_vfok.place(relx=0.5,rely=0.4,anchor=tk.CENTER)
        self.image_vfok = ctk.CTkLabel(
            self.labelframe_vfok,
            image =self.correctcheck,
            text="",
            bg_color='#FFDCB9',
            fg_color="#8CEA00",
        )
        self.image_vfok.place(relx=0.5,rely=0.6,anchor=tk.CENTER)
        #建立驗證失敗圖示
        self.label_vfnok = ctk.CTkLabel(
            self.labelframe_vfnok,
            text="驗證失敗...",
            bg_color='#FFDCB9',
            font=('微軟正黑體',40),
            text_color='#000000',
            fg_color="#FF2D2D",
            # width=250,height=400
        )
        self.label_vfnok.place(relx=0.5,rely=0.4,anchor=tk.CENTER)
        self.image_vfnok = ctk.CTkLabel(
            self.labelframe_vfnok,
            image =self.wrongcheck,
            text="",
            bg_color='#FFDCB9',
            fg_color="#FF2D2D",
        )
        self.image_vfnok.place(relx=0.5,rely=0.6,anchor=tk.CENTER)
        #建立範例檔案下載按鈕
        self.btn_4 = ctk.CTkButton(
            self.master,
            text="範例檔案下載",
            anchor=tk.CENTER,
            command=None,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,
        )
        self.btn_4.place(relx=0.15,rely=0.8,anchor=tk.CENTER)
        #建立重新匯入按鈕
        self.btn_r = ctk.CTkButton(
            self.master,
            text="重新匯入",
            anchor=tk.CENTER,
            command=self.refresh,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,
        )
        self.btn_r.place(relx=0.15,rely=0.87,anchor=tk.CENTER)
        #建立返回按鈕
        self.btn_back = ctk.CTkButton(
            self.master,
            text="返回",
            anchor=tk.CENTER,
            command= lambda: self.back(oldmaster),
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,
        )
        self.btn_back.place(relx=0.15,rely=0.94,anchor=tk.CENTER)
        self.updatetojson_btn = ctk.CTkButton(
            self.master,
            text="上傳至資料庫",
            image =self.dbimage,
            compound="bottom",
            anchor=tk.CENTER,
            command=self.upload,
            fg_color="#FFDCB9",
            bg_color='#FFEEDD', 
            hover_color="#FF9224",
            font=('微軟正黑體',24),
            text_color="#000000",
            state=tk.DISABLED,
            width=250,height=380
        )
        self.updatetojson_btn.pack()
        self.updatetojson_btn.place(relx=0.8,rely=0.12,anchor=tk.N)
        #建立查看標籤
        self.label_5 = ctk.CTkLabel(
            self.master,
            text="""查看
>>""",
            anchor=tk.CENTER,
            fg_color="#FFEEDD",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22),
            text_color="#000000",
            width=60,height=120
        )
        self.label_5.place(relx=0.32,rely=0.45,anchor=tk.CENTER)
        #建立上傳標籤
        self.label_6 = ctk.CTkLabel(
            self.master,
            text="""上傳
>>""",
            anchor=tk.CENTER,
            fg_color="#FFEEDD",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22),
            text_color="#000000",
            width=60,height=120
        )
        self.label_6.place(relx=0.65,rely=0.45,anchor=tk.CENTER)
        #建立驗證verifyframe
        self.verifyframe=ctk.CTkFrame(
            self.master,
            fg_color="#FFDCB9",
            bg_color="#FFEEDD",
            width=250,height=380
        )
        self.verifyframe.place(relx=0.49,rely=0.12,anchor=tk.N)
        
        self.cc = ctk.CTkLabel(
            self.master, 
            fg_color="#FFEEDD",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            font=("Calibri",12),
            width=170)
        self.cc.place(relx=0.86,rely=0.98,anchor=tk.W)
        self.labelframe_ok.pack_forget()
        self.labelframe_nok.pack_forget()
        self.labelframe_vfnok.pack_forget()
        self.labelframe_vfok.pack_forget()
    ##選擇檔案後，驗證是否匯入成功
    def selectfile(self):
        global testtype
        self.filename  = filedialog.askopenfilename(initialdir="E:\土城長庚醫院",title="考題匯入",filetypes=(("Excel","*.xlsx"),("CSV UTF-8","*.csv"),("Excel 2003","*.xls"),("all files","*.*")))
        ##建立檔案路徑及檔案名稱核對使用
        self.labelframe_1.pack_forget()
        self.labelframe_2.pack_forget()
        self.labelframe_3.pack_forget()
        #在verifyframe中加入讀取資訊
        self.label_path = ctk.CTkLabel(
            self.verifyframe, 
            text = "檔案路徑:",
            fg_color='#FFDCB9',
            font=('微軟正黑體',20),
            text_color="#000000",
            )
        self.label_path.grid(row=0,column=0,padx=10,columnspan=6,sticky='w')
        self.selectfile_path = ctk.CTkTextbox(
            self.verifyframe, 
            fg_color='#FFDCB9',
            font=('微軟正黑體',14),
            text_color="#000000",
            width=220,height=40
            )
        self.selectfile_path.insert('0.0',self.filename)
        self.selectfile_path.configure(state="disabled")
        self.selectfile_path.grid(row=1,column=0,columnspan=6)
        #擷取excel裡面資訊，放入entry中，核對是否正確
        file_path = self.filename
        try:
            excel_filename = r"{}".format(file_path)
            # print(basename)
            if excel_filename[-4:] == ".csv":
                self.df = pd.read_csv(excel_filename,skiprows=2)
                testname = pd.read_csv(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
            else:
                self.df = pd.read_excel(excel_filename,skiprows=2)
                testname = pd.read_excel(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
        except ValueError:
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            tk.messagebox.showerror('土城長庚醫院檢驗科', message='資料格式不符，請重新選擇檔案')
            return None
        except FileNotFoundError:
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            tk.messagebox.showerror("Information", message='未選擇檔案，請選擇檔案後再重新驗證!')
            return None
        self.testyear = str(testname.columns[1])
        self.testname = testname.iloc[0,1]
        db = self.df.to_numpy().tolist()
        # db =[["WBC",0],["RBC",2],["HB",3],["Hct",4],["MCV",5],["MCH",6],["MCHC",7],["RDW",8],["Plt",9],["plasma cell",0]...]
        createtable(self.verifyframe,28,2,lst=db)
        #加入驗證按鈕，後續匯入database中
        self.btn_6 = ctk.CTkButton(
            self.verifyframe,
            text="驗證",
            anchor=tk.CENTER,
            command=self.verifyfile,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=190,
        )
        self.btn_6.grid(row=16,column=0,columnspan=6,padx=30,pady=20,sticky='nsew')
        testtype="blood"
    ########################################尿液考題匯入########################################
    ##選擇檔案後，驗證是否匯入成功(尿液)
    def selectfile_urin(self):
        global testtype
        self.filename  = filedialog.askopenfilename(initialdir="E:\土城長庚醫院",title="考題匯入",filetypes=(("Excel","*.xlsx"),("CSV UTF-8","*.csv"),("Excel 2003","*.xls"),("all files","*.*")))
        ##建立檔案路徑及檔案名稱核對使用
        self.labelframe_1.pack_forget()
        self.labelframe_2.pack_forget()
        self.labelframe_3.pack_forget()
        #在verifyframe中加入讀取資訊
        self.label_path = ctk.CTkLabel(
            self.verifyframe, 
            text = "檔案路徑:",
            fg_color='#FFDCB9',
            font=('微軟正黑體',20),
            text_color="#000000",
            )
        self.label_path.grid(row=0,column=0,padx=10,columnspan=6,sticky='w')
        self.selectfile_path = ctk.CTkTextbox(
            self.verifyframe, 
            fg_color='#FFDCB9',
            font=('微軟正黑體',14),
            text_color="#000000",
            width=220,height=40
            )
        self.selectfile_path.insert('0.0',self.filename)
        self.selectfile_path.configure(state="disabled")
        self.selectfile_path.grid(row=1,column=0,columnspan=6)
        #擷取excel裡面資訊，放入entry中，核對是否正確
        file_path = self.filename
        try:
            excel_filename = r"{}".format(file_path)
            if excel_filename[-4:] == ".csv":
                self.df = pd.read_csv(excel_filename,skiprows=1)
                testname = pd.read_csv(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
            else:
                self.df = pd.read_excel(excel_filename,skiprows=1)
                testname = pd.read_excel(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
        except ValueError:
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            tk.messagebox.showerror('土城長庚醫院檢驗科', message='資料格式不符，請重新選擇檔案')
            return None
        except FileNotFoundError:
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            tk.messagebox.showerror("Information", message='未選擇檔案，請選擇檔案後再重新驗證!')
            return None
        self.testname = testname.columns[1]
        # print(self.testname)
        db = self.df.to_numpy().tolist()
        # db =[["WBC",0],["RBC",2],["HB",3],["Hct",4],["MCV",5],["MCH",6],["MCHC",7],["RDW",8],["Plt",9]]
        # print(db)
        ##這個要記得改
        createtable(self.verifyframe,28,2,lst=db)
        #加入驗證按鈕，後續匯入database中
        self.btn_6 = ctk.CTkButton(
            self.verifyframe,
            text="驗證",
            anchor=tk.CENTER,
            command=self.verifyfile_urine,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=190,
        )
        self.btn_6.grid(row=16,column=0,columnspan=6,padx=30,pady=20,sticky='nsew')
        testtype="urine"
    ########################################體液考題匯入########################################
    ##選擇檔案後，驗證是否匯入成功(尿液)
    def selectfile_bodyfluid(self):
        global testtype
        self.filename  = filedialog.askopenfilename(initialdir="E:\土城長庚醫院",title="考題匯入",filetypes=(("Excel","*.xlsx"),("CSV UTF-8","*.csv"),("Excel 2003","*.xls"),("all files","*.*")))
        ##建立檔案路徑及檔案名稱核對使用
        self.labelframe_1.pack_forget()
        self.labelframe_2.pack_forget()
        self.labelframe_3.pack_forget()
        #在verifyframe中加入讀取資訊
        self.label_path = ctk.CTkLabel(
            self.verifyframe, 
            text = "檔案路徑:",
            fg_color='#FFDCB9',
            font=('微軟正黑體',20),
            text_color="#000000",
            )
        self.label_path.grid(row=0,column=0,padx=10,columnspan=6,sticky='w')
        self.selectfile_path = ctk.CTkTextbox(
            self.verifyframe, 
            fg_color='#FFDCB9',
            font=('微軟正黑體',14),
            text_color="#000000",
            width=220,height=40
            )
        self.selectfile_path.insert('0.0',self.filename)
        self.selectfile_path.configure(state="disabled")
        self.selectfile_path.grid(row=1,column=0,columnspan=6)
        #擷取excel裡面資訊，放入entry中，核對是否正確
        file_path = self.filename
        try:
            excel_filename = r"{}".format(file_path)
            if excel_filename[-4:] == ".csv":
                self.df = pd.read_csv(excel_filename,skiprows=1)
                testname = pd.read_csv(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
            else:
                self.df = pd.read_excel(excel_filename,skiprows=1)
                testname = pd.read_excel(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
        except ValueError:
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            tk.messagebox.showerror('土城長庚醫院檢驗科', message='資料格式不符，請重新選擇檔案')
            return None
        except FileNotFoundError:
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            tk.messagebox.showerror("Information", message='未選擇檔案，請選擇檔案後再重新驗證!')
            return None
        self.testname = testname.columns[1]
        # print(self.testname)
        db = self.df.to_numpy().tolist()
        # db =[["WBC",0],["RBC",2],["HB",3],["Hct",4],["MCV",5],["MCH",6],["MCHC",7],["RDW",8],["Plt",9]]
        # print(db)
        ##這個要記得改
        createtable(self.verifyframe,28,2,lst=db)
        #加入驗證按鈕，後續匯入database中
        self.btn_6 = ctk.CTkButton(
            self.verifyframe,
            text="驗證",
            anchor=tk.CENTER,
            command=self.verifyfile_bf,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=190,
        )
        self.btn_6.grid(row=16,column=0,columnspan=6,padx=30,pady=20,sticky='nsew')
        testtype="bodyfluid"
    ##驗證檔案是否符合格式，是否可以存入JSON
    def verifyfile(self):
        global testtype
        # file_path = self.filename
        # print(file_path)
        ##對比項目
        verify=[]
        jsonfile = open('testdata\data.json','rb')
        rawdata = json.load(jsonfile)
        yearIDdict={}
        ID = []
        if testtype == "blood":
            self.rawdataID = rawdata["blood"]
            self.rawdata = rawdata["blood"][0]["rawdata"]
            self.rawdata_a = rawdata["blood"][0]["Ans"]
        elif testtype == "urine":
            self.rawdataID = rawdata["urine"]
            self.rawdata = rawdata["urine"][0]["rawdata"]
            self.rawdata_a = rawdata["urine"][0]["Ans"]
        else:
            self.rawdataID = rawdata["bodyfluid"]
            self.rawdata = rawdata["bodyfluid"][0]["rawdata"]
            self.rawdata_a = rawdata["bodyfluid"][0]["Ans"]
        # for item in self.rawdataID:
        #     ID.append(item["ID"])
        for item in self.rawdataID:
            if item["year"] not in yearIDdict:
                yearIDdict[item["year"]] = list()
            yearIDdict[item["year"]].append(item["ID"])
        # print(yearIDdict)
        # fliter = (self.rawdata['year'] == testno)
        verify = [key for key in self.rawdata.keys()]
        for key in self.rawdata_a.keys():
            verify.append(key)
        # print(verify)
        ##對比是否有重複值在JSON內        
        if self.testyear in yearIDdict:
            if self.testname in yearIDdict[self.testyear]:
                self.labelframe_vfnok.place(relx=0.49,rely=0.5,anchor=tk.CENTER)
                self.labelframe_vfnok.tkraise()
                tk.messagebox.showerror(title='土城長庚檢驗科', message="檔案重複!請檢查檔案後重新輸入!!")
                return None
        else:
            pass
        ##對比匯入資料
        try:
            itemcompair = self.df["項目"]
        except KeyError:
            self.labelframe_vfnok.place(relx=0.49,rely=0.5,anchor=tk.CENTER)
            self.labelframe_vfnok.tkraise()
            tk.messagebox.showerror(title='土城長庚檢驗科', message="檔案錯誤!請檢查檔案後重新輸入!!")
        # print(itemcompair)
        self.verifyframe.pack_forget()
        for j in range(0,len(itemcompair)):
            if itemcompair[j] == verify[j]:
                continue
            else:
                self.labelframe_vfnok.place(relx=0.49,rely=0.5,anchor=tk.CENTER)
                self.labelframe_vfnok.tkraise()
                tk.messagebox.showerror(title='土城長庚檢驗科', message="檔案錯誤!請檢查檔案後重新輸入!!")
                return None
        self.labelframe_vfok.pack()
        self.labelframe_vfok.place(relx=0.49,rely=0.5,anchor=tk.CENTER)
        self.labelframe_vfok.tkraise()
        self.updatetojson_btn.configure(state=tk.NORMAL)
    ##返回basdesk按鈕    
    def back(self,oldmaster):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return
    ##重新整理，以利再次上傳
    def refresh(self):
        #清除匯入成功或者不成功的框架
        self.labelframe_ok.place_forget()
        self.labelframe_nok.place_forget()
        #清除驗證成功或者不成功的框架
        self.labelframe_vfok.place_forget()
        self.labelframe_vfnok.place_forget()
        ##把匯入檔案框架提至上層
        self.tab.tkraise()
        
        self.verifyframe.place(relx=0.49,rely=0.12,anchor=tk.N)
        self.verifyframe.tkraise()
        #清除檔案路徑
        self.selectfile_path.configure(state='normal')
        self.selectfile_path.delete(0.0,tk.END)
        self.selectfile_path.configure(state='disable')
        #清除格子
        for entry in entrylst:
            entry.configure(state="normal")
            entry.delete(0,tk.END)
            entry.configure(state="disabled")
        #驗證按鈕反灰
        self.btn_6.configure(state="disabled")
    ##更新至JSON
    def upload(self):
        #加入的rawdata
        rawdata = self.df.to_numpy().tolist()
        dict_r,dict_a = {},{}
        for i in range(0,9):
            dict_r[rawdata[i][0]] = rawdata[i][1]
        for j in range(9,len(rawdata)):
            dict_a[rawdata[j][0]] = rawdata[j][1]
        # print(dict_r)
        # print(dict_a)
        # print(self.testname)
        #uploadfile = 準備要寫入json的dict
        uploadfile={}
        uploadfile["year"] = self.testyear
        uploadfile["ID"] = self.testname
        uploadfile["testtype"] = "blood"
        uploadfile["rawdata"] = dict_r
        uploadfile["Ans"] = dict_a
        # print(uploadfile)
        aa = open('testdata\data.json','r+')
        jsonfile = json.load(aa)
        json_blood = jsonfile["blood"]
        json_blood.append(uploadfile)
        aa.seek(0,0)
        aa.truncate()   ##清空JSON檔案
        aa.write(json.dumps(jsonfile))
        # print(a)
        tk.messagebox.showinfo(title='土城長庚檢驗科', message="考片建立成功!!")
def main():  
    root = ctk.CTk()
    I = Import(root)
    # root['bg']='#FFEEDD'
    # B.gui_arrang()
    # 主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  