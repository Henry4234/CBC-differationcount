import os,shutil
import tkinter as tk
from tkinter import ttk
from tkinter import RIDGE, DoubleVar, StringVar, ttk, messagebox,IntVar, filedialog
from turtle import bgcolor, width
import customtkinter  as ctk
import basedesk, basedesk_admin,json,pyodbc
from PIL import Image
import pandas as pd
##取得登入帳號資訊
def getaccount(acount):
    global Baccount
    Baccount = str(acount)
    # print(Baccount)
    return None

connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=bloodtest;UID=sa;PWD=1234"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.OperationalError:
    connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)

####年份跟ID####
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

##ctk無法有效支援treeview，所以使用entery + loop解決
global entrylst
entrylst = []
def createtable(master,startrow,totalrows,totalcolumns,lst):
    
    for i in range(totalrows):
        for j in range(totalcolumns):
            TB = tk.Entry(master, width=9, fg='blue',
                            font=('Arial',10,))
            entrylst.append(TB)
            if i < 5:
                TB.grid(row=startrow, column=j)
                TB.insert(tk.END, lst[i][j])
            elif i >= 5 and i < 9:
                TB.grid(row=startrow-5, column=j+2)
                TB.insert(tk.END, lst[i][j])
            else:
                pass
            TB.configure(state="disabled")
        startrow += 1
def createtable_ans(master,startrow,totalrows,totalcolumns,lst):
    for i in range(totalrows):
        for j in range(0,totalcolumns):
            TB = tk.Entry(master, width=9, fg='blue',
                            font=('Arial',10,))
            entrylst.append(TB)
            TB.grid(row=startrow, column=j,sticky='ew')
            TB.insert(tk.END, lst[i][j])
            TB.configure(state="disabled")
        startrow += 1

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
        self.master.geometry("1100x850")
        self.master.config(background='#FFEEDD')
        # self.master.config(bg='#FFEEDD')
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

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
            width=250,height=500
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
            width=250,height=500
            )
        self.labelframe_1.pack(side="top",fill='both')
        self.labelframe_1.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        #建立尿液中分頁的框架
        self.labelframe_2 = ctk.CTkFrame(
            self.tab.tab("尿液"),
            fg_color="#FFDCB9",
            bg_color="#FFDCB9",
            width=250,height=500
            )
        self.labelframe_2.pack(side="top",fill='both')
        self.labelframe_2.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        #建立體液中分頁的框架
        self.labelframe_3 = ctk.CTkFrame(
            self.tab.tab("體液"),
            fg_color="#FFDCB9",
            bg_color="#FFDCB9",
            width=250,height=500
            )
        self.labelframe_3.pack(side="top",fill='both')
        self.labelframe_3.place(relx=0.5,rely=0.5,anchor=tk.CENTER)
        ##匯入成功不成功的框架
        self.labelframe_ok = ctk.CTkFrame(
            self.master,
            fg_color="#8CEA00",
            bg_color="#FFEEDD",
            width=250,height=500,
            corner_radius=8
            )
        self.labelframe_ok.pack()
        self.labelframe_nok = ctk.CTkFrame(
            self.master,
            fg_color="#FF2D2D",
            bg_color="#FFEEDD",
            width=250,height=500,
            corner_radius=8
            )
        self.labelframe_nok.pack()
        ##驗證成功不成功的frame
        self.labelframe_vfok = ctk.CTkFrame(
            self.master,
            fg_color="#8CEA00",
            bg_color="#FFEEDD",
            width=280,height=710,
            corner_radius=8
            )
        self.labelframe_vfok.pack()
        self.labelframe_vfnok = ctk.CTkFrame(
            self.master,
            fg_color="#FF2D2D",
            bg_color="#FFEEDD",
            width=280,height=710,
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
            width=250,height=500
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
            width=250,height=500
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
            width=250,height=500
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
            # width=250,height=500
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
            # width=250,height=500
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
            # width=250,height=500
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
            # width=250,height=500
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
            command=self.download_file,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22),
            text_color="#000000",
            width=250,
        )
        self.btn_4.place(relx=0.15,rely=0.78,anchor=tk.CENTER)
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
        self.btn_r.place(relx=0.15,rely=0.84,anchor=tk.CENTER)
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
        self.btn_back.place(relx=0.15,rely=0.9,anchor=tk.CENTER)
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
            width=250,height=480
        )
        self.updatetojson_btn.pack()
        self.updatetojson_btn.place(relx=0.8,rely=0.13,anchor=tk.N)
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
            width=250,height=480
        )
        self.verifyframe.place(relx=0.49,rely=0.13,anchor=tk.N)
        
        self.cc = ctk.CTkLabel(
            self.master, 
            fg_color="#FFEEDD",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            font=("Calibri",12),
            width=170)
        self.cc.place(relx=0.86,rely=1,anchor=tk.W)
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
            width=220,height=30
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
                self.df = pd.read_csv(excel_filename,skiprows=6)
                testname = pd.read_csv(excel_filename)
                self.labelframe_ok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)  #把匯入成功的frame蓋上去
                self.labelframe_ok.tkraise()
            else:
                self.df = pd.read_excel(excel_filename,skiprows=6)
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

        ##建立檔案資訊
        self.label_info = ctk.CTkLabel(
            self.verifyframe, 
            text = "檔案資訊:",
            fg_color='#FFDCB9',
            font=('微軟正黑體',20),
            text_color="#000000",
            )
        self.label_info.grid(row=2,column=0,padx=10,columnspan=6,sticky='w')
        ##檔案年分
        # TB = tk.Entry(self.verifyframe, width=9, fg='blue',
        #                     font=('Arial',10,))
        # TB.grid(row=i+3, column=j)
        self.testyear = str(testname.columns[1])
        self.testname = testname.iloc[0,1]
        self.gender = testname.iloc[1,1]
        self.age = testname.iloc[2,1]
        self.comment = testname.iloc[3,1]
        self.count_value = testname.iloc[4,1]
        info_lst = ["label_year","val_year","label_testname","val_testname","label_gender","val_gender","label_age","val_age","label_count","val_count"]
        info_name = ["考片年份",self.testyear,"考片編號",self.testname,"性別",self.gender,"年齡",self.age,"數細胞數",self.count_value]
        for i in range(0,3):
            for j in range(0,4):
                info_lst[i] = tk.Entry(self.verifyframe, width=9, fg='blue',font=('Arial',10,))
                entrylst.append(info_lst[i])
                if i >= 2 :
                    if j == 0:
                        info_lst[i].grid(row=i+3, column=j)
                        info_lst[i].insert(tk.END, info_name[4*i + j])
                    if j == 1:
                        info_lst[i].grid(row=i+3, column=j,columnspan=3,sticky='ew')
                        info_lst[i].insert(tk.END, info_name[4*i + j])
                    else:
                        pass
                else:
                    info_lst[i].grid(row=i+3, column=j)
                    info_lst[i].insert(tk.END, info_name[4*i + j])
                info_lst[i].configure(state="disabled")
        label_comment = tk.Entry(self.verifyframe, width=9, fg='blue',font=('Arial',10,),)
        val_comment = tk.Entry(self.verifyframe,width=28, fg='blue',font=('Arial',10,),)
        entrylst.append(label_comment) 
        entrylst.append(val_comment)
        label_comment.grid(row=6, column=0)
        val_comment.grid(row=6, column=1,columnspan=3)
        label_comment.insert(tk.END,"備註")
        val_comment.insert(tk.END,self.comment)
        label_comment.configure(state="disabled")
        val_comment.configure(state="disabled")
        #整理database資料
        db = self.df.T.apply(lambda x: x.dropna().tolist()).tolist()
        # print(db)
        # db = self.df.to_numpy().tolist()
        # db =[["WBC",0],["RBC",2],["HB",3],["Hct",4],["MCV",5],["MCH",6],["MCHC",7],["RDW",8],["Plt",9],["plasma cell",0]...]
        try:
            createtable(self.verifyframe,7,9,2,lst=db)
        except IndexError:
            tk.messagebox.showerror("土城長庚醫院檢驗科", message='檔案格式不符，請檢查後再上傳')
            self.labelframe_nok.place(relx=0.15,rely=0.42,anchor=tk.CENTER)
            self.labelframe_nok.tkraise()
            return None
        LA_1 = tk.Entry(self.verifyframe, width=9, fg='blue',font=('Arial',10,))
        LA_1.grid(row=12, column=0,sticky='ew')
        LA_1.insert(tk.END,"細胞")
        LA_2 = tk.Entry(self.verifyframe, width=9, fg='blue',font=('Arial',10,))
        LA_2.grid(row=12, column=1,sticky='ew')
        LA_2.insert(tk.END,"數值")
        LA_3 = tk.Entry(self.verifyframe, width=9, fg='blue',font=('Arial',10,))
        LA_3.grid(row=12, column=2,sticky='ew')
        LA_3.insert(tk.END,"must")
        LA_4 = tk.Entry(self.verifyframe, width=9, fg='blue',font=('Arial',10,))
        LA_4.grid(row=12, column=3,sticky='ew')
        LA_4.insert(tk.END,"mustnot")
        del db[0:9]
        createtable_ans(self.verifyframe,13,18,4,lst=db)
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
        self.btn_6.grid(row=31,column=0,columnspan=6,padx=30,pady=10,sticky='nsew')
        testtype="blood"
#↓↓↓↓↓##################################尿液考題匯入##################################↓↓↓↓↓#
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
#↑↑↑↑↑##################################尿液考題匯入##################################↑↑↑↑↑#
#↓↓↓↓↓##################################體液考題匯入##################################↓↓↓↓↓#
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
            self.labelframe_ok.place_forget()
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
#↑↑↑↑↑##################################體液考題匯入##################################↑↑↑↑↑#
    ##驗證檔案是否符合格式，是否可以存入JSON
    def verifyfile(self):
        global testtype
        # file_path = self.filename
        # print(file_path)
        ##對比項目
        # yearIDdict={}
        ##rawdataID = 建立一個dict，key=year，values=ID，去確保沒有重複上傳
        ##這個可以用yearid這個dict解決
        ##                                        ┌rawdata = ["WBC","RBC",...]
        ##merge成一個list:verify，然後比對傳入資料┤
        ##                                        └rawdata_a = ["plasma cell","abnormal lympho"]
        verify = ["WBC","RBC","HB","Hct","MCV","MCH","MCHC","RDW","Plt","plasma cell","abnormal lympho","megakaryocyte","nRBC","blast","metamyelocyte","eosinopil","plasmacytoid","promonocyte","promyelocyte","band neutropil","basopil","atypical lymphocyte","hypersegmented neutrophil","myelocyte","segmented neutrophil","lymphocyte","monocyte"]
        # print(verify)
        ##對比是否有重複值在JSON內        
        if self.testyear in yearid:
            if self.testname in yearid[self.testyear]:
                self.labelframe_vfnok.place(relx=0.49,rely=0.42,anchor=tk.CENTER)
                self.labelframe_vfnok.tkraise()
                tk.messagebox.showerror(title='土城長庚檢驗科', message="檔案重複!請檢查檔案後重新輸入!!")
                return None
        else:
            pass
        ##對比匯入資料
        try:
            itemcompair = self.df["項目"]
        except KeyError:
            self.labelframe_vfnok.place(relx=0.49,rely=0.42,anchor=tk.CENTER)
            self.labelframe_vfnok.tkraise()
            tk.messagebox.showerror(title='土城長庚檢驗科', message="檔案錯誤!請檢查檔案後重新輸入!!")
            return
        # print(itemcompair)
        self.verifyframe.pack_forget()

        for j in range(0,len(itemcompair)):
            if itemcompair[j] == verify[j]:
                continue
            else:
                self.labelframe_vfnok.place(relx=0.49,rely=0.42,anchor=tk.CENTER)
                self.labelframe_vfnok.tkraise()
                tk.messagebox.showerror(title='土城長庚檢驗科', message="檔案錯誤!請檢查檔案後重新輸入!!")
                return None
        ##還要驗證加起來是否100 跟 must & mustnot有沒有打架
        v_db = self.df.drop([i for i in range(0,9)])
        # print(v_db)
        #加總100
        sum_val = v_db["數值"].sum()
        if round(sum_val) < 101 and round(sum_val) >= 99:
            pass
        else:
            self.labelframe_vfnok.place(relx=0.49,rely=0.42,anchor=tk.CENTER)
            self.labelframe_vfnok.tkraise()
            tk.messagebox.showerror(title='土城長庚檢驗科', message="細胞加總非100!請檢查檔案後重新輸入!!")
            return None
        #must/mustnot檢查
        must_lst = v_db["must"].tolist()
        mustnot_lst = v_db["mustnot"].tolist()
        for j in range(0,len(must_lst)):
            if must_lst[j] != mustnot_lst[j] or must_lst[j]==0 and mustnot_lst[j]==0:
                continue
            else:
                self.labelframe_vfnok.place(relx=0.49,rely=0.42,anchor=tk.CENTER)
                self.labelframe_vfnok.tkraise()
                tk.messagebox.showerror(title='土城長庚檢驗科', message="不可同時存在must/mustnot!請檢查檔案後重新輸入!!")
                return None
        self.labelframe_vfok.pack()
        self.labelframe_vfok.place(relx=0.49,rely=0.54,anchor=tk.CENTER)
        self.labelframe_vfok.tkraise()
        self.updatetojson_btn.configure(state=tk.NORMAL)
    ##下載範例檔案
    def download_file(self):
        dl_path = filedialog.askdirectory()
        if dl_path:
            template="testdata\\template_CBCDATA.xlsx"
            # template = template.replace("\\","/")
            destination_path = os.path.join(dl_path, "template_CBCDATA.xlsx")
            try:
                shutil.copy(template, destination_path)
                tk.messagebox.showinfo(title='土城醫院檢驗科', message=f"文件已成功保存到 {destination_path}")
            except Exception as e:
                tk.messagebox.showerror(title='土城醫院檢驗科', message=f"保存文件时出错：{e}")
        else:
            return
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
        
        self.verifyframe.place(relx=0.49,rely=0.13,anchor=tk.N)
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
        try:
            self.btn_6.configure(state="disabled")
        except AttributeError:
            pass
    ##(X)更新至JSON
    ##(O)更新至SQL
        #要更新3個table:["bloodinfo","bloodinfo_cbc","bloodinfo_ans"]
        #先整理資料格式
    def upload(self):
        #加入的rawdata
        ##建立不同table的匯入數值
        lst_r = self.df["數值"].head(9).tolist()
        str_r = ','.join(str(i) for i in lst_r)
        lst_a = self.df["數值"].tail(18).tolist()
        str_a = ','.join(str(i) for i in lst_a)
        ans_upload = self.df.T.apply(lambda x: x.dropna().tolist()).tolist()
        del ans_upload[0:9]
        print(ans_upload)
        # print(rawdata,lst_r,lst_a)
        #匯入SQL
        with coxn.cursor() as cursor:
        #匯入bloodinfo table
            bloodinfo_input = """INSERT INTO [bloodtest].[dbo].[bloodinfo]
            ("year","smear_id","gender","age","count_value","comment")
            VALUES (%s,'%s','%s',%d,%d,'%s');""" %(self.testyear,self.testname,self.gender,self.age,self.count_value,self.comment)
        #匯入bloodinfo_cbc table
            bloodinfocbc_input = """INSERT INTO [bloodtest].[dbo].[bloodinfo_cbc]
            ("smear_id","WBC","RBC","HB","Hct","MCV","MCH","MCHC","RDW","plt")
            VALUES ('%s',%s);""" %(self.testname,str_r)
            cursor.execute(bloodinfo_input)
            cursor.execute(bloodinfocbc_input)
        #匯入bloodinfo_ans2 table
            for item in range(0,len(ans_upload)):
                bloodinfoans_input = """INSERT INTO bloodtest.dbo.bloodinfo_ans2 ([smear_id],[celltype],[value],[must],[mustnot]) 
                VALUES('%s','%s',%.2f,%d,%d);"""%(self.testname,ans_upload[item][0],ans_upload[item][1],ans_upload[item][2],ans_upload[item][3])
            #     bloodinfoans_input = """INSERT INTO [bloodtest].[dbo].[bloodinfo_ans]
            # ("smear_id","plasma cell","abnormal lympho","megakaryocyte","nRBC","blast","metamyelocyte","eosinophil","plasmacytoid","promonocyte","promyelocyte","band neutropil","basopil","atypical lymphocyte","hypersegmented neutrophil","myelocyte","segmented neutrophil","lymphocyte","monocyte") 
            # VALUES ('%s',%s);""" %(self.testname,str_a)
                # print(bloodinfoans_input)
                cursor.execute(bloodinfoans_input)
        coxn.commit()
        # for i in range(0,9):
        #     dict_r[rawdata[i][0]] = rawdata[i][1]
        # for j in range(9,len(rawdata)):
        #     dict_a[rawdata[j][0]] = rawdata[j][1]
        #     dict_l[rawdata[j][0]] = rawdata[j][2]
        #     dict_u[rawdata[j][0]] = rawdata[j][3]
        # print(dict_r)
        # print(dict_a)
        # print(self.testname)
        #uploadfile = 準備要寫入json的dict
        # uploadfile={}
        # uploadfile["year"] = self.testyear
        # uploadfile["ID"] = self.testname
        # uploadfile["testtype"] = "blood"
        # uploadfile["gender"] = self.gender
        # uploadfile["age"] = self.age
        # uploadfile["rawdata"] = dict_r
        # uploadfile["comment"] = self.comment
        # uploadfile["Ans"] = dict_a
        # uploadfile["criteria_lower"] = dict_u
        # uploadfile["criteria_upper"] = dict_l
        # print(uploadfile)
        # aa = open('testdata\data_new.json','r+',encoding="utf8")
        # jsonfile = json.load(aa)
        # json_blood = jsonfile["blood"]
        # json_blood.append(uploadfile)
        # # print(json_blood)
        # aa.seek(0,0)
        # aa.truncate()   ##清空JSON檔案
        # aa.write(json.dumps(jsonfile,ensure_ascii=False))
        # print(aa)
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