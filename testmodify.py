#authorised by Henry Tsai
import sys
import tkinter as tk
import pandas as pd
from tkinter import IntVar, simpledialog,messagebox,DoubleVar,StringVar
import customtkinter as ctk
from setuptools import Command
from verifyAccount import changepw2,addaccount,delaccount,editaccount
from PIL import Image

import pyodbc
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import sqlalchemy as sa

import pyodbc
#建立與mySQL連線資料
connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=bloodtest;UID=sa;PWD=1234"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.OperationalError:
    connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)

connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)

class Modify:
    
    def __init__(self,master,oldmaster=None):  
        # 建立登入後視窗  
        self.master = master
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.exit(oldmaster))
        super().__init__()
        ctk.set_default_color_theme("blue")
        # jsonfile = open('testdata\data_new.json','rb')
        # rawdata = json.load(jsonfile)
        # self.rawdata = pd.DataFrame(rawdata["blood"])
        with coxn.cursor() as cursor:
            cursor.execute("SELECT DISTINCT [year] FROM [bloodtest].[dbo].[bloodinfo];")
        # self.testyear = self.rawdata['year'].unique().tolist()
            self.testyear = [str(row[0]) for row in cursor.fetchall()]
        # 給主視窗設定標題內容  
        self.master.title("考題設定")  
        self.master.geometry('900x650')
        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_columnconfigure(1, weight=3)
        self.master.config(background='#FFEEDD') #設定背景色
        ##框架設置
        self.labelframe_1 = ctk.CTkFrame(self.master, corner_radius=0,fg_color="#FFDCB9",bg_color="#FFEEDD")
        self.labelframe_1.grid(row=0, column=0,rowspan=2, sticky="nsew")
        # self.labelframe_1.grid_rowconfigure(4, weight=1)
        self.labelframe_2 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.labelframe_2.grid(row=0, column=1)
        self.frame_2s = ctk.CTkFrame(self.labelframe_2,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.frame_2s.grid(row=2, column=2,columnspan=3,rowspan=4,)
        self.labelframe_3 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.frame_3s = ctk.CTkFrame(self.labelframe_3,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.frame_3s.grid(row=2, column=2,columnspan=3,rowspan=4,padx=20)
        self.labelframe_4 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        # self.labelframe_2.grid_columnconfigure(1, weight=1)
        ##data格式(包含CBC資訊跟答案)
        self.testinfo={
            "WBC":DoubleVar(),
            "RBC":DoubleVar(),
            "HB":DoubleVar(),
            "Hct":DoubleVar(),
            "MCV":DoubleVar(),
            "MCH":DoubleVar(),
            "MCHC":DoubleVar(),
            "RDW":DoubleVar(),
            "plt":DoubleVar()
        }
        self.ansinfo={
            "plasma cell": DoubleVar(),
            "abnormal lympho": DoubleVar(),
            "megakaryocyte": DoubleVar(),
            "nRBC": DoubleVar(),
            "blast": DoubleVar(),
            "metamyelocyte": DoubleVar(),
            "eosinophil": DoubleVar(),
            "plasmacytoid": DoubleVar(),
            "promonocyte": DoubleVar(),
            "promyelocyte": DoubleVar(),
            "band neutropil": DoubleVar(),
            "basopil": DoubleVar(),
            "atypical lymphocyte": DoubleVar(),
            "hypersegmented neutrophil": DoubleVar(),
            "myelocyte": DoubleVar(),
            "segmented neutrophil": DoubleVar(),
            "lymphocyte": DoubleVar(),
            "monocyte": DoubleVar()
        }
        self.ageinfo={
            "age": IntVar()
        }
        self.commentinfo={
            "comment": StringVar(),
        }
        ##左手邊功能區
        self.people = ctk.CTkImage(Image.open("assets\pages.png"),size=(120,120))
        self.label_1 = ctk.CTkLabel(
                    self.labelframe_1, 
                    image=self.people,
                    compound="top",
                    text = "考題設定", 
                    fg_color='#FFDCB9',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_1.grid(row=0,column=0,padx=20, pady=20)
        #血液考題設定
        self.btn_1=ctk.CTkButton(
            self.labelframe_1,
            command = self.switch1, 
            text = "血液考題設定", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_1.grid(row=1,column=0,pady=10,sticky="ew")
        #血液考題參數設定
        self.btn_1t=ctk.CTkButton(
            self.labelframe_1,
            command = self.changefigure, 
            text = "血液考題參數設定", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_1t.grid(row=2,column=0,pady=10,sticky="ew")
        #尿液考題設定
        self.btn_2=ctk.CTkButton(
            self.labelframe_1,
            command = None, 
            text = "尿液考題設定", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_2.grid(row=3,column=0,pady=10,sticky="ew")
        #離開按鈕
        self.btn_3=ctk.CTkButton(
            self.labelframe_1,
            command = lambda: self.exit(oldmaster), 
            text = "返回", 
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
        ##(f2)血液考題設定
        #(f2)標題
        self.label_2 = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "血液考題設定", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_2.grid(row=0,column=0,columnspan=5,sticky='nsew')
        #(f2)年份標籤 & 下拉欄位
        self.label_year = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "請選擇考核年度:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_year.grid(row=1,column=0,pady=10,sticky='w')
        self.input_year = ctk.CTkComboBox(
                    master = self.labelframe_2,
                    # command= None,
                    command= self.updatelist,
                    values=self.testyear,
                    text_color='#000000',
                    fg_color='#F0F0F0',
                    button_color='#FF9900',
                    state="readonly",
                    width=120,height=40,
                    font=('微軟正黑體',16,)
                    )
        self.input_year.grid(row=1,column=1,padx=10,pady=10,columnspan=2,sticky='w')
        #(f2)左邊考題選擇listbox & 旁邊捲軸
        self.scrollbar = tk.Scrollbar(self.labelframe_2)
        
        self.test_listbox = tk.Listbox(
            self.labelframe_2,
            yscrollcommand=self.scrollbar.set,
            height=10,width=8,
            font=('微軟正黑體',16))
        self.test_listbox.grid(row=2,column=0,rowspan=2,sticky='nsew',)
        self.scrollbar.grid(row=2,column=1,rowspan=2,sticky='nsew')
        self.scrollbar.config(command=self.test_listbox.yview)
        #(f2)點選listbox之後連結事件self.listbox_event
        self.test_listbox.bind("<<ListboxSelect>>", self.listbox_event)        
        #加上基本資訊小標題
        self.title2 = ctk.CTkLabel(
            self.labelframe_2, 
            text = "基本資料:", 
            fg_color='#FFEEDD',
            font=('微軟正黑體',18),
            text_color="#000000",
            width=140,height=40
            )
        self.title2.grid(row=2,column=2,sticky='nw')
        ##(f2s)旁邊編輯窗格
        self.testinfolist = [key for key in self.testinfo.keys()]
        self.ansinfolist = [key for key in self.ansinfo.keys()]
        global entrylst
        entrylst = []
        #loop 0,放入項目(WBC/RBC/Hct...)
        for j in range(0,9):
            #第一欄
            self.modifytable = ctk.CTkEntry(
                self.frame_2s,
                width=100,height=20,
                bg_color='#FFEEDD',
                fg_color='#FFCC99',
                # text="",
                font=('微軟正黑體',12),
                text_color="#000000"
                )
            try:
                self.modifytable.insert(tk.END,self.testinfolist[j])
            #補空格補到10格
            except IndexError:
                self.modifytable = ctk.CTkEntry(
                self.frame_2s,
                width=100,height=20,
                bg_color='#FFEEDD',
                fg_color='#FFCC99',
                # text="",
                font=('微軟正黑體',12),
                text_color="#000000"
                )
            finally:
                self.modifytable.grid(row=j,column=0)
                self.modifytable.configure(state="disabled")
        #loop 2,放入性別(gender)
        self.modifytable_gender = ctk.CTkEntry(
                self.frame_2s,
                width=100,height=50,
                bg_color='#FFEEDD',
                fg_color='#FFCC99',
                # text="",
                font=('微軟正黑體',16),
                text_color="#000000"
                )
        self.modifytable_gender.insert(tk.END,"性別")
        self.modifytable_gender.grid(row=0,column=2,rowspan=2,sticky='n')
        self.modifytable_gender.configure(state="disabled")
        #loop 4,放入年齡(comment)
        self.modifytable_age = ctk.CTkEntry(
                self.frame_2s,
                width=100,height=50,
                bg_color='#FFEEDD',
                fg_color='#FFCC99',
                # text="",
                font=('微軟正黑體',16),
                text_color="#000000"
                )
        self.modifytable_age.insert(tk.END,"年齡")
        self.modifytable_age.grid(row=2,column=2,rowspan=2,sticky='n')
        self.modifytable_age.configure(state="disabled")
        #loop 6,放入備註(comment)
        self.modifytable_comment = ctk.CTkEntry(
                self.frame_2s,
                width=100,height=120,
                bg_color='#FFEEDD',
                fg_color='#FFCC99',
                # text="",
                font=('微軟正黑體',16),
                text_color="#000000"
                )
        self.modifytable_comment.insert(tk.END,"其他備註")
        self.modifytable_comment.grid(row=4,column=2,rowspan=5,sticky='n')
        self.modifytable_comment.configure(state="disabled")
        #loop 1,放入變數(self.testinfo.values())
        m = 0
        for values in self.testinfo.values():
            self.modifytable = ctk.CTkEntry(
                    self.frame_2s,
                    width=50,height=20,
                    bg_color='#FFEEDD',
                    fg_color='#FFFFFF',
                    textvariable= values,
                    font=('微軟正黑體',14),
                    text_color="#000000"
                    )
            self.modifytable.grid(row = m, column = 1,sticky='nsew')
            self.modifytable.configure(state="disabled")
            entrylst.append(self.modifytable)
            m += 1
        #loop 3,放入性別combobox欄位(self.testinfo.values())
        self.modifytable_gender_val = ctk.CTkComboBox(
                    self.frame_2s,
                    width=130,height=50,
                    button_color='#FF9900',
                    bg_color='#FFEEDD',
                    fg_color='#FFFFFF',
                    values=["male","female"],
                    font=('微軟正黑體',16),
                    text_color="#000000"
                    )
        self.modifytable_gender_val.grid(row = 0, column = 3,rowspan=2,sticky='n')
        self.modifytable_gender_val.configure(state="disabled")
        #loop 5,放入年齡變數(self.ageinfo.values())
        self.modifytable = ctk.CTkEntry(
                    self.frame_2s,
                    width=130,height=50,
                    bg_color='#FFEEDD',
                    fg_color='#FFFFFF',
                    textvariable= self.ageinfo['age'],
                    font=('微軟正黑體',16),
                    text_color="#000000"
                    )
        self.modifytable.grid(row = 2, column = 3,rowspan=2,sticky='n')
        self.modifytable.configure(state="disabled")
        entrylst.append(self.modifytable)
        #loop 7,放入其他備註變數(self.commoninfo.values())
        self.modifytable_comment_val = ctk.CTkTextbox(
                    self.frame_2s,
                    width=130,height=120,
                    border_width=2,
                    bg_color='#FFEEDD',
                    fg_color='#FFFFFF',
                    # textvariable= self.commentinfo['comment'],
                    font=('微軟正黑體',16),
                    text_color="#000000"
                    )
        self.modifytable_comment_val.grid(row = 4, column = 3,rowspan=5,sticky='n')
        self.modifytable_comment_val.configure(state="disabled")
        entrylst.append(self.modifytable_comment_val)
        #(f2)編輯/確認/清除按鈕
        self.edit_btn = ctk.CTkButton(
            self.labelframe_2,
            text = "編輯",
            command=self.edittable,
            height=30,
            fg_color='#FF6600',
            text_color='#000000')
        self.edit_btn.grid(row=4,column=2,padx=20,pady=15)
        self.yes_btn = ctk.CTkButton(
            self.labelframe_2,
            text = "確定",
            command=self.editdata,
            height=30,
            fg_color='#FF9900',
            text_color='#000000',
            state='disabled')
        self.yes_btn.grid(row=4,column=3,padx=20,pady=15)
        self.clear_btn = ctk.CTkButton(
            self.labelframe_2,
            text = "清除",
            command=self.clear,
            height=30,
            fg_color='#FF9900',
            text_color='#000000')
        self.clear_btn.grid(row=4,column=4,padx=20,pady=15)

##左手邊按鍵區域 血液考題設定/血液參數設定/尿液考題設定/離開
    #(f3)參數設定介面介面
    def changefigure(self):
        for value in self.ansinfo.values():
            value.set(0)        # self.labelframe_2.grid_forget()
        entrylst.clear()
        self.slcid = StringVar()
        ##把frame3放到frame1上
        self.labelframe_3.grid(row=0, column=1,sticky='nsew')
        self.labelframe_3.columnconfigure(2,weight=1)
        self.label_3 = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "考題參數設定", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120,
                    )
        self.label_3.grid(row=0,column=0,columnspan=6,sticky='nsew')
        #(f3)年份標籤 & 下拉欄位
        self.f3_label_year = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "請選擇考核年度:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.f3_label_year.grid(row=1,column=0,pady=10,sticky='w')
        self.f3_input_year = ctk.CTkComboBox(
                    master = self.labelframe_3,
                    # command= None,
                    command= self.f3_updatelist,
                    values=self.testyear,
                    text_color='#000000',
                    fg_color='#F0F0F0',
                    button_color='#FF9900',
                    state="readonly",
                    width=120,height=40,
                    font=('微軟正黑體',16,)
                    )
        self.f3_input_year.grid(row=1,column=1,padx=10,pady=10,columnspan=2,sticky='w')
        #(f3)左邊考題選擇listbox & 旁邊捲軸
        self.scrollbar = tk.Scrollbar(self.labelframe_3)
        
        self.f3_test_listbox = tk.Listbox(
            self.labelframe_3,
            yscrollcommand=self.scrollbar.set,
            height=10,width=8,
            font=('微軟正黑體',16))
        self.f3_test_listbox.grid(row=2,column=0,rowspan=2,sticky='nsew',)
        self.scrollbar.grid(row=2,column=1,rowspan=2,sticky='nsew')
        self.scrollbar.config(command=self.f3_test_listbox.yview)
        #(f3)點選listbox之後連結事件self.listbox_event
        self.f3_test_listbox.bind("<<ListboxSelect>>", self.f3_listbox_event)
        self.title3 = ctk.CTkLabel(
            self.labelframe_3, 
            text = "考題參數:", 
            fg_color='#FFEEDD',
            font=('微軟正黑體',18),
            text_color="#000000",
            width=140,height=40
            )
        self.title3.grid(row=2,column=2,sticky='nw')
        #loop 0,2,4 放入項目(cell type)
        for h in range(0,6,2):
            for k in range(0,9):
                if h == 0:
                    self.modifytable = ctk.CTkEntry(
                            self.frame_3s,
                            width=100,height=20,
                            bg_color='#FFEEDD',
                            fg_color='#FFCC99',
                            # text="",
                            font=('微軟正黑體',12),
                            text_color="#000000"
                            )
                    self.modifytable.insert(tk.END,self.ansinfolist[k])
                    self.modifytable.grid(row=k,column=h)
                    self.modifytable.configure(state="disabled")
                elif h == 2:
                    self.modifytable = ctk.CTkEntry(
                        self.frame_3s,
                        width=100,height=20,
                        bg_color='#FFEEDD',
                        fg_color='#FFCC99',
                        # text="",
                        font=('微軟正黑體',12),
                        text_color="#000000"
                        )
                    #從ansinfolist[11]開始跳
                    try:
                        self.modifytable.insert(tk.END,self.ansinfolist[k+9])
                    #補空格補到10格
                    except IndexError:
                        self.modifytable = ctk.CTkEntry(
                        self.frame_3s,
                        width=100,height=20,
                        bg_color='#FFEEDD',
                        fg_color='#FFCC99',
                        # text="",
                        font=('微軟正黑體',12),
                        text_color="#000000"
                        )
                    finally:
                        self.modifytable.grid(row=k,column=h)
                        self.modifytable.configure(state="disabled")
        #(f3)loop 1,3,5放入變數(self.testinfo.values())
        m = 0
        for values in self.ansinfo.values():
            self.modifytable = ctk.CTkEntry(
                    self.frame_3s,
                    width=50,height=20,
                    bg_color='#FFEEDD',
                    fg_color='#FFFFFF',
                    textvariable= values,
                    font=('微軟正黑體',14),
                    text_color="#000000"
                    )
            if m<9:
                self.modifytable.grid(row = m, column = 1,sticky='nsew')
                self.modifytable.configure(state="disabled")
                entrylst.append(self.modifytable)
                m += 1
            else:
                self.modifytable.grid(row = m-9, column = 3,sticky='nsew')
                self.modifytable.configure(state="disabled")
                entrylst.append(self.modifytable)
                m += 1
        ##(f3)變數補空格
        # self.modifytable = ctk.CTkEntry(
        #             self.frame_3s,
        #             width=50,height=20,
        #             bg_color='#FFEEDD',
        #             fg_color='#FFFFFF',
        #             font=('微軟正黑體',14),
        #             text_color="#000000"
        #             )
        # self.modifytable.grid(row = 9, column = 3,sticky='nsew')
        # self.modifytable.configure(state="disabled")
        #(f3)編輯/確認/清除按鈕
        self.f3_edit_btn = ctk.CTkButton(
            self.labelframe_3,
            text = "編輯",
            command=self.f3_edittable,
            height=30,
            fg_color='#FF6600',
            text_color='#000000')
        self.f3_edit_btn.grid(row=4,column=2,padx=20,pady=15)
        self.f3_yes_btn = ctk.CTkButton(
            self.labelframe_3,
            text = "確定",
            command=self.f3_editdata,
            height=30,
            fg_color='#FF9900',
            text_color='#000000',
            state='disabled')
        self.f3_yes_btn.grid(row=4,column=3,padx=20,pady=15)
        self.f3_clear_btn = ctk.CTkButton(
            self.labelframe_3,
            text = "清除",
            command=self.f3_clear,
            height=30,
            fg_color='#FF9900',
            text_color='#000000')
        self.f3_clear_btn.grid(row=4,column=4,padx=20,pady=15)

    #(f1)離開
    def exit(self,oldmaster):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return 
##(f2)三個功能鍵"編輯"/"確定"/"清除"
    #(f2)編輯
    def edittable(self):
        #取得考片ID
        self.focusyear = self.input_year.get()
        idx = self.test_listbox.curselection()
        if self.focusyear=="" :
            tk.messagebox.showerror(title="土城醫院檢驗科",message="尚未選擇年份!!請選擇後再開始編輯!")
            return
        elif idx ==():
            tk.messagebox.showerror(title="土城醫院檢驗科",message="尚未選擇考片!!請選擇後再開始編輯!")
            return
        else:
            if tk.messagebox.askyesno(title='土城醫院檢驗科', message='確定編輯?'):    
                self.focustest = self.test_listbox.get(idx)
                #鎖住combobox & listbox不要選到其他的考片
                self.input_year.configure(state="disabled")
                self.test_listbox.configure(state="disabled")
                #打開確定按鈕
                self.yes_btn.configure(state="normal")
                for value in entrylst:
                    value.configure(state="normal")
                #開啟combobox
                self.modifytable_gender_val.configure(state="normal")
            else:
                return
    #(f2)確定
    def editdata(self):
        if tk.messagebox.askyesno(title='土城醫院檢驗科', message='改完了?'):
            # i = self.rawdata.index
            #年份跟ID做為filter
            # filter1 = (self.rawdata['year'] == self.focusyear)
            # filter2 = (self.rawdata['ID'] == self.focustest)
            #把原本的資料調進來
            # index = filter1 & filter2
            # result = i[index]
            # result.tolist()
            # result=result[0]
            #建立一個新的dict，裝修改過後的testdata資料
            dict_rawdata={}
            for key,value in self.testinfo.items():
                A = value.get()
                dict_rawdata[key]=A
            for key,value in dict_rawdata.items():
                with coxn.cursor() as cursor:
                    query = "UPDATE [bloodtest].[dbo].[bloodinfo_cbc] SET [%s] = %.2f WHERE [smear_id]='%s';"%(key,value,self.focustest)
                    cursor.execute(query)
                coxn.commit()
            # self.rawdata.at[result,'rawdata'] = dict_rawdata
            #建立一個新val，裝修改過後的gender資料
            m_gender = self.modifytable_gender_val.get()
            #建立一個新val，裝修改過後的age資料
            m_age = self.ageinfo['age'].get()
            #建立一個新val，裝修改過後的comment資料
            m_comment = self.modifytable_comment_val.get(0.0,ctk.END)
            with coxn.cursor() as cursor:
                query = "UPDATE [bloodtest].[dbo].[bloodinfo] SET [gender] = '%s',[age] = %d, [comment]='%s' WHERE [smear_id]='%s';"%(m_gender,m_age,m_comment,self.focustest)
                cursor.execute(query)
            coxn.commit()
            
            # add = self.rawdata.to_json(orient="records")
            # jsonfile = open('testdata\data_new.json','rb')
            # a = json.load(jsonfile)
            # a["blood"] = json.loads(add)
            # with open('testdata/data_new.json','w',encoding='utf8') as r:
            #     json.dump(a,r,ensure_ascii=False)
            #     r.close()
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='修改成功!')
            self.input_year.configure(state="normal")
            self.test_listbox.configure(state="normal")
            self.yes_btn.configure(state="disabled")
            for value in entrylst:
                value.configure(state="disabled")
        else:
            return
    #(f2)清除
    def clear(self):
        if tk.messagebox.askyesno(title='土城醫院檢驗科', message='確定清除?'):
            self.input_year.set("")
            self.test_listbox.delete(0,tk.END)
            for value in self.testinfo.values():
                value.set(0.0)
            self.ageinfo['age'].set(0)
            self.modifytable_comment_val.configure(state='normal')
            self.modifytable_comment_val.delete("0.0",ctk.END)
            self.modifytable_comment_val.configure(state='disable')
        else:
            return
##(f2)事件連結
    #選擇年份之後，更新listbox(JSON裡面有的考題)
    def updatelist(self,event):
        #建立年分與smearid的dict
        with coxn.cursor() as cursor:
            cursor.execute('SELECT "year","smear_id" FROM [bloodtest].[dbo].[bloodinfo];')
            sql_yearid = cursor.fetchall()
        yearid = dict()
        for row in sql_yearid:
            sql_year, smear_id = row
            sql_year = str(sql_year)
            if sql_year not in yearid:
                yearid[sql_year] = []
            yearid[sql_year].append(smear_id)
        #先看一下是哪一個介面再呼叫updatelist
        #取得combobox選擇的年份
        year = self.input_year.get()   
        #建立一個filter把年份對的篩選出來
        # fliter = (self.rawdata['year'] == year)
        # self.testlist = self.rawdata[fliter]["ID"]
        self.testlist = yearid[year]
        #放入資料之前，先把listbox清空
        self.test_listbox.delete(0,tk.END)
        for i in self.testlist:
            self.test_listbox.insert(tk.END,i)
    #listbox選擇後，更新self.testinfo內的變數
    def listbox_event(self,event):
        #先看一下是哪一個介面再呼叫listbox_event

        #取得combobox選擇的年份
        # year = self.input_year.get()
        #取得listbox選擇的考題
        idx = self.test_listbox.curselection()
        try:
            slcid = self.test_listbox.get(idx)
        except tk.TclError:
            return
        #建立兩個filter篩選出考題
        # fliter1 = (self.rawdata['year'] == year)
        # fliter2 = (self.rawdata['ID'] == slcid)
        # datalist = self.rawdata[fliter1 & fliter2]
        with engine.begin() as conn:
            filter_test = "SELECT * FROM [bloodtest].[dbo].[bloodinfo_cbc] WHERE [smear_id]='%s';"%(slcid)
            datadict1 = pd.read_sql_query(sa.text(filter_test), conn).to_dict('records')[0]
            del datadict1['smear_id']
        #兩個dict分別放CBC DATA & Ans
        with coxn.cursor() as cursor:
            query = "SELECT [gender],[age],[comment] FROM [bloodtest].[dbo].[bloodinfo] WHERE [smear_id]='%s';"%(slcid)
            cursor.execute(query)
            datadict2 = cursor.fetchall()
            # print(datadict2)
        # datadict1 = datalist.iloc[0]["rawdata"]
        # datadict2 = datalist.iloc[0]["gender"]
        # datadict3 = datalist.iloc[0]["age"]
        # datadict4 = datalist.iloc[0]["comment"]
        #CBCRAW DATA數據更改
        for key,value in datadict1.items():
            self.testinfo[key].set(value)
        #性別combobox更改
        self.modifytable_gender_val.configure(state="normal")
        if datadict2[0][0] == "female":
            self.modifytable_gender_val.set("female")
        else:
            self.modifytable_gender_val.set("male")
        self.modifytable_gender_val.configure(state="disabled")
        #年齡數據更改
        self.ageinfo["age"].set(datadict2[0][1])
        #其他備註數據更改
        self.modifytable_comment_val.configure(state="normal")
        self.modifytable_comment_val.delete("0.0",ctk.END)
        self.modifytable_comment_val.insert(ctk.END,datadict2[0][2])
        self.modifytable_comment_val.configure(state="disabled")
        # self.commentinfo["comment"].set(datadict4)
##(f3)三個功能鍵"編輯"/"確定"/"清除"
    #(f3)編輯
    def f3_edittable(self):
        #取得考片ID
        self.focusyear = self.f3_input_year.get()
        idx = self.f3_test_listbox.curselection()
        if self.focusyear=="" :
            tk.messagebox.showerror(title="土城醫院檢驗科",message="尚未選擇年份!!請選擇後再開始編輯!")
            return
        elif idx ==():
            tk.messagebox.showerror(title="土城醫院檢驗科",message="尚未選擇考片!!請選擇後再開始編輯!")
            return
        else:
            if tk.messagebox.askyesno(title='土城醫院檢驗科', message='確定編輯?'):    
                self.focustest = self.f3_test_listbox.get(idx)
                #鎖住combobox & listbox不要選到其他的考片
                self.f3_input_year.configure(state="disabled")
                self.f3_test_listbox.configure(state="disabled")
                #打開確定按鈕
                self.f3_yes_btn.configure(state="normal")
                for value in entrylst:
                    value.configure(state="normal")
            else:
                return
    #(f3)確定
    def f3_editdata(self):
        if tk.messagebox.askyesno(title='土城醫院檢驗科', message='改完了?'):
            # i = self.rawdata.index
            # #年份跟ID做為filter
            # filter1 = (self.rawdata['year'] == self.focusyear)
            # filter2 = (self.rawdata['ID'] == self.focustest)
            # #把原本的資料調進來
            # index = filter1 & filter2
            # result = i[index]
            # result.tolist()
            # result=result[0]
            #建立一個新的dict，裝修改過後的testdata資料
            dict_rawdata={}
            for key,value in self.ansinfo.items():
                A = value.get()
                dict_rawdata[key]=A
            for key,value in dict_rawdata.items():
                with coxn.cursor() as cursor:
                    query = "UPDATE [bloodtest].[dbo].[bloodinfo_ans] SET [%s] = %.2f WHERE [smear_id]='%s';"%(key,value,self.focustest)
                    cursor.execute(query)
                coxn.commit()
            # self.rawdata.at[result,'Ans'] = dict_rawdata
            # add = self.rawdata.to_json(orient="records")
            # jsonfile = open('testdata\data_new.json','rb')
            # a = json.load(jsonfile)
            # a["blood"] = json.loads(add)
            # with open('testdata/data_new.json','w',encoding='utf8') as r:
            #     json.dump(a,r,ensure_ascii=False)
            #     r.close()
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='修改成功!')
            self.f3_input_year.configure(state="normal")
            self.f3_test_listbox.configure(state="normal")
            self.f3_yes_btn.configure(state="disabled")
            for value in entrylst:
                value.configure(state="disabled")
        else:
            return
    #(f3)清除
    def f3_clear(self):
        if tk.messagebox.askyesno(title='土城醫院檢驗科', message='確定清除?'):
            self.f3_input_year.set("")
            self.f3_test_listbox.delete(0,tk.END)
            for value in self.ansinfo.values():
                value.set(0)
        else:
            return
##(f3)事件連結
    #(f3)選擇年份之後，更新listbox(JSON裡面有的考題)
    def f3_updatelist(self,event):
        #取得combobox選擇的年份
        with coxn.cursor() as cursor:
            cursor.execute('SELECT "year","smear_id" FROM [bloodtest].[dbo].[bloodinfo];')
            sql_yearid = cursor.fetchall()
        yearid = dict()
        for row in sql_yearid:
            sql_year, smear_id = row
            sql_year = str(sql_year)
            if sql_year not in yearid:
                yearid[sql_year] = []
            yearid[sql_year].append(smear_id)
        year = self.f3_input_year.get()   
        #建立一個filter把年份對的篩選出來
        # fliter = (self.rawdata['year'] == year)
        # self.testlist = self.rawdata[fliter]["ID"]
        self.testlist = yearid[year]
        #放入資料之前，先把listbox清空
        self.f3_test_listbox.delete(0,tk.END)
        for i in self.testlist:
            self.f3_test_listbox.insert(tk.END,i)
    #(f3)listbox選擇後，更新self.testinfo內的變數
    def f3_listbox_event(self,event):
        #取得combobox選擇的年份
        # year = self.f3_input_year.get()
        #取得listbox選擇的考題
        idx = self.f3_test_listbox.curselection()
        try:
            slcid = self.f3_test_listbox.get(idx)
        except tk.TclError:
            return
        #建立兩個filter篩選出考題
        # fliter1 = (self.rawdata['year'] == year)
        # fliter2 = (self.rawdata['ID'] == slcid)
        # datalist = self.rawdata[fliter1 & fliter2]
        #兩個dict分別放CBC DATA & Ans
        # datadict2 = datalist.iloc[0]["Ans"]
        # with engine.begin() as conn:
        #     filter_test = "SELECT * FROM [bloodtest].[dbo].[bloodinfo_ans] WHERE [smear_id]='%s';"%(slcid)
        #     datadict2 = pd.read_sql_query(sa.text(filter_test), conn).to_dict('records')[0]
        #     del datadict2['smear_id']
        with coxn.cursor() as cursor:
            filter_test = "SELECT [celltype],[value] FROM [bloodtest].[dbo].[bloodinfo_ans2] WHERE [smear_id]='%s';"%(slcid)
            cursor.execute(filter_test)
            datadict2 = dict(cursor.fetchall())
        for key,value in datadict2.items():
            self.ansinfo[key].set(value)
##(switch)從血液參數設定切回考題設定
    def switch1(self):
        self.labelframe_3.grid_forget()
        entrylst.clear()
        self.labelframe_1.tkraise()
        # self.frame_2s.tkraise()
    
def main():  
    root = ctk.CTk()
    M = Modify(root)
    #  主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  