import sys,os
import tkinter as tk
import customtkinter  as ctk
from tksheet import Sheet
from tkinter import ttk
from tkinter import RIDGE, DoubleVar, StringVar, ttk, messagebox,IntVar, filedialog

from turtle import width
from PIL import Image
import json, pyodbc,basedesk, basedesk_admin
import pandas as pd
import numpy as np

from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import sqlalchemy as sa

# def getaccount(acount):
#     global Baccount
#     Baccount = str(acount)
#     # print(Baccount)
#     return None
#SQL連線設定
connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=220.133.50.28;DATABASE=bloodtest;UID=cgmh;PWD=B[-!wYJ(E_i7Aj3r"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.InterfaceError:
    # connection_string = "DRIVER={ODBC Driver 11 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
    connection_string = "DRIVER={SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)
#取出年份(唯一)
with coxn.cursor() as cursor:
    sql = """SELECT DISTINCT [bloodinfo].[year] 
FROM [bloodtest].[dbo].[test_data]
JOIN [bloodtest].[dbo].[bloodinfo]
ON [test_data].[smear_id]=[bloodinfo].[smear_id];"""
    cursor.execute(sql)
    un_year = cursor.fetchall()
un_year = [e[0] for e in un_year]
str_un_year =[]
for item in un_year:
    item  = str(item)
    str_un_year.append(item)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class SCORE_CAL:

    def __init__(self,master,oldmaster=None):
        self.master = master
        super().__init__()
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.back(oldmaster))        
        ctk.set_default_color_theme("blue")
        self.master.title("考題成績計算")
        self.master.geometry("1200x650")
        self.master.config(background='#FFEEDD')
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=2)
        #框架設計
        self.labelframe_1 = ctk.CTkFrame(self.master,corner_radius=0,fg_color="#FFDCB9",bg_color="#FFEEDD")
        self.labelframe_1.grid(row=0, column=0, rowspan=2,sticky='nsew')
        self.labelframe_2 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.labelframe_2.grid(row=0, column=1)
        self.testicon = ctk.CTkImage(Image.open(resource_path("assets\pages.png")),size=(120,120))
        with coxn.cursor() as cursor:
            cursor.execute("SELECT [院區] FROM [bloodtest].[dbo].[hospital_code];")
            self.hospital = [str(row[0]) for row in cursor.fetchall()]
        self.hospital = ["","All"] + self.hospital
        #引入henry計算Range
        filename = resource_path("./chart/rangechart.json")
        self.rangechart = pd.read_json(filename)
        # self.celllst = Ans
        #計算用的list
        global anslst,entrylst, mustlst, mustnotlst
        anslst,entrylst, mustlst, mustnotlst=[],[],[],[]
        #匯入標題
        self.welcome = ctk.CTkLabel(
                    self.labelframe_1, 
                    image=self.testicon,
                    compound="top",
                    text = "考題成績試算", 
                    # fg_color='#FFDCB9',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.welcome.grid(row=0,column=0,padx=10, pady=20)
        self.all_combobox = ctk.CTkComboBox(
                    self.labelframe_1, 
                    command=self.callback_all,
                    fg_color='#FFDCB9',
                    button_color='#FF9900',
                    values=self.hospital,
                    font=('微軟正黑體',22),
                    text_color="#000000",
                    width=120,height=50
                    )
        self.all_combobox.grid(row=1,column=0,padx=20, pady=20,sticky='ew')
        self.year_combobox = ctk.CTkComboBox(
                    self.labelframe_1, 
                    fg_color='#FFDCB9',
                    button_color='#FF9900',
                    values=str_un_year,
                    font=('微軟正黑體',22),
                    text_color="#000000",
                    width=120,height=50
                    )
        self.year_combobox.grid(row=2,column=0,padx=20, pady=20,sticky='nsew')
        self.btn_search = ctk.CTkButton(
            self.labelframe_1,
            text="搜尋",
            anchor=tk.CENTER,
            command=self.search_gai,
            fg_color="#FF9224",
            bg_color='#FFDCB9',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,height=20,
        )
        self.btn_search.grid(row=3,column=0,padx=20,pady=30,sticky='ew')
        self.frame_testtable=ctk.CTkFrame(
            self.labelframe_2,
            fg_color="#FFEEDD",
            bg_color="#FFEEDD",
            height=150,width=900,
            )
        self.frame_testtable.grid(row=0,column=0,columnspan=2,padx=10,pady=20,)
        self.frame_anstable=ctk.CTkFrame(
            self.labelframe_2,
            fg_color="#000000",
            bg_color="#FFDCB9",
            width=500,
            )
        self.frame_anstable.grid(row=1,column=0,rowspan=7,padx=5,pady=20,)
        self.btn_to_sht2 = ctk.CTkButton(
            self.labelframe_2,
            text="↙選擇DATA",
            anchor=tk.CENTER,
            command=self.to_sht2,
            fg_color="#66B3FF",
            bg_color='#FFDCB9',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,height=20,
        )
        self.btn_to_sht2.grid(row=1,column=1,)
        #標籤_計算細胞量
        self.label_comment = ctk.CTkLabel(
            self.labelframe_2,
            width=120,height=20,
            bg_color="#FFEEDD",
            fg_color="#FFEEDD",
            text="考生備註:",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_comment.grid(row=2,column=1,)
        self.input_testcomment = ctk.CTkTextbox(
            self.labelframe_2,
            height=80,corner_radius=8,
            fg_color='#FFFFFF',
            font=('微軟正黑體',16),
            text_color="#000000",
            state='disable'
            )
        self.input_testcomment.grid(row=3,column=1,)
        self.label_tscore = ctk.CTkLabel(
            self.labelframe_2,
            width=120,
            bg_color="#FFEEDD",
            fg_color="#FFEEDD",
            text="得分",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_tscore.grid(row=4,column=1,pady=0)
        self.combobox_tscore = ctk.CTkComboBox(
                    self.labelframe_2, 
                    fg_color='#FFDCB9',
                    button_color='#FF9900',
                    values=["0","1","2"],
                    font=('微軟正黑體',22),
                    text_color='#000000',
                    width=120,height=40
                    )
        self.combobox_tscore.grid(row=5,column=1,)
        self.btn_cal = ctk.CTkButton(
            self.labelframe_2,
            text="計算結果",
            anchor=tk.CENTER,
            command=self.calculate_cell,
            fg_color="#FF9224",
            bg_color='#FFDCB9',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,height=20,
        )
        self.btn_cal.grid(row=6,column=1,)
        self.btn_upload = ctk.CTkButton(
            self.labelframe_2,
            text="上傳結果",
            anchor=tk.CENTER,
            command=self.upload_score,
            fg_color="#FF9224",
            bg_color='#FFDCB9',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,height=20,
        )
        self.btn_upload.grid(row=7,column=1,)
        #清除資料
        self.btn_clear = ctk.CTkButton(
            self.labelframe_1,
            text="清除搜尋結果",
            anchor=tk.CENTER,
            command=self.clear,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,
        )
        self.btn_clear.grid(row=4,column=0,padx=20,sticky='ew')
        #aaa
        self.btn_aaa = ctk.CTkButton(
            self.labelframe_1,
            text="※先不要按※",
            anchor=tk.CENTER,
            command=self.multi_upload,
            fg_color="#7FBEEB",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,
        )
        self.btn_aaa.grid(row=5,column=0,padx=20,pady=20,sticky='ew')
        #返回主介面
        self.btn_back = ctk.CTkButton(
            self.labelframe_1,
            text="返回",
            anchor=tk.CENTER,
            command=lambda: self.back(oldmaster),
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,
        )
        self.btn_back.grid(row=6,column=0,padx=20,pady=10,sticky='ew')
        ##tksheet_1結果表格設計
        self.sht_result = Sheet(
            self.frame_testtable,
            column_width=175,
            width=900,height=150,

        )
        self.sht_result.headers(["院區","考生姓名","考片_ID","考試次數","交卷時間"])
        self.sht_result.enable_bindings(('single_select',"row_select","drag_select"))
        
        self.sht_result.grid(row=0,column=0,sticky='nsew')
        ##tksheet_2點選後跳出欄位
        self.sht_ans = Sheet(
            self.frame_anstable,
            column_width=115,
            width=650,height=400,
        )
        self.sht_ans.headers(["細胞","答案數值","must_chk","mustnot_chk","考生DATA"])
        self.sht_ans.enable_bindings()
        self.sht_ans.grid(row=0,column=0,sticky='ew')
    def callback_all(self,event):
        h_site = self.all_combobox.get()
        if h_site == "All":
            self.year_combobox.set("")
            self.year_combobox.configure(values = [])
            self.year_combobox.configure(state = 'disable')
        else:
            self.year_combobox.configure(state = 'normal')
            self.year_combobox.configure(values = str_un_year)
            

##################↓↓↓↓↓搜尋: 查看當前所選擇的院區以及年分↓↓↓↓↓##################
    def search_gai(self):
        h_site = self.all_combobox.get()
        testyear = self.year_combobox.get()
        if h_site =="":
            tk.messagebox.showerror(title='檢驗醫學部(科)', message='請選擇院區!所有院區請選擇ALL')
        elif h_site =="All":
            srh = """SELECT [hospital_code].[院區],[id].[ac],[test_data].[smear_id],[test_data].[count],[test_data].[timestamp] 
            FROM [bloodtest].[dbo].[test_data] 
            JOIN [bloodtest].[dbo].[id] 
            ON [id].[No] = [test_data].[test_id]
            JOIN [bloodtest].[dbo].[hospital_code] 
            ON [hospital_code].[code] = [id].[院區]
            WHERE [score] IS NULL;"""
        else:
            srh = """SELECT [bloodtest].[dbo].[hospital_code].[院區],[bloodtest].[dbo].[id].[ac],[bloodtest].[dbo].[test_data].[smear_id],[bloodtest].[dbo].[test_data].[count],[bloodtest].[dbo].[test_data].[timestamp] 
            FROM [bloodtest].[dbo].[test_data] 
            JOIN [bloodtest].[dbo].[id] ON [id].[No] = [test_data].[test_id]
            JOIN [bloodtest].[dbo].[bloodinfo] ON [bloodinfo].[smear_id] = [test_data].[smear_id]
            JOIN [bloodtest].[dbo].[hospital_code] ON [hospital_code].[code] = [id].[院區]
            WHERE [hospital_code].[院區] = '%s' AND [bloodinfo].[year] = %d 
            AND [score] IS NULL;"""%(h_site,int(testyear))
        with coxn.cursor() as cursor:
            cursor.execute(srh)
            data_sht_1 = cursor.fetchall()
        # print(data_sht_1)
        if data_sht_1 ==[]:
            tk.messagebox.showinfo(title='檢驗醫學部(科)',message="搜尋結果並無須要計算結果之考片") 
            return
        else:
            self.sht_result.set_sheet_data(data_sht_1,reset_col_positions=True,reset_row_positions=True)
        # return
##################↑↑↑↑↑搜尋: 查看當前所選擇的院區以及年分↑↑↑↑↑##################
#####↓↓↓↓↓輸入tksheet_2: 將選好的考試答案，加入tksheet_2與答案一起比對↓↓↓↓↓#####
    def to_sht2(self):
        try:
            last_selected = self.sht_result.get_currently_selected()
        # print(last_selected)
            #row_data= ["院區",ac,id,cellcount_value]
            self.row_data = self.sht_result.get_row_data(last_selected.row)
        
        except AttributeError:
           tk.messagebox.showerror(title='檢驗醫學部(科)',message="先查詢批改資訊後，選擇後再進行批改!") 
           return
        except IndexError:
            tk.messagebox.showerror(title='檢驗醫學部(科)',message="先查詢批改資訊後，選擇後再進行批改!") 
            return
        if tk.messagebox.askyesno(title='檢驗醫學部(科)',message="上述選擇正確?"):
            # print (self.row_data)
            merge =[]
            #取出血片解答
            with coxn.cursor() as cursor:
                sql_tk2 = """SELECT [celltype],[value],[must],[mustnot] 
FROM [bloodtest].[dbo].[bloodinfo_ans2]
WHERE [smear_id] ='%s';"""%(self.row_data[2])
                cursor.execute(sql_tk2)
                self.data_sht_2 = cursor.fetchall()
                ##self.data_sht_2格式 =[('plasma cell',2.0,1,0)]
            # print(self.data_sht_2)
            #取出百分比
            with coxn.cursor() as cursor:
                sql_per = "SELECT [count_value] FROM [bloodtest].[dbo].[bloodinfo] WHERE [smear_id]='%s';"%(self.row_data[2])
                cursor.execute(sql_per)
                n = cursor.fetchone()
            self.data_count_val= int(n[0])
            #取出學生考試答案
            with coxn.cursor() as cursor:
                sql_tk2s = """SELECT [plasma cell],[abnormal lympho],[megakaryocyte],[nRBC],[blast],[metamyelocyte],[eosinophil],[plasmacytoid],[promonocyte],[promyelocyte],[band neutropil],[basopil],[atypical lymphocyte],[hypersegmented neutrophil],[myelocyte],[segmented neutrophil],[lymphocyte],[monocyte]
                FROM [bloodtest].[dbo].[test_data] 
                JOIN [bloodtest].[dbo].[id]
                ON [id].[No] = [test_data].[test_id]
                WHERE [test_id] IN (SELECT [id].[No] WHERE [id].[ac]='%s')
                AND [smear_id]='%s' 
                AND [count]=%d;"""%(self.row_data[1],self.row_data[2],self.row_data[3])
                cursor.execute(sql_tk2s)
                data_sht_2s = cursor.fetchone()
            #百分比轉換
            self.cal_percent = [] 
            for i in range(0,len(data_sht_2s)):
                a = data_sht_2s[i] / self.data_count_val *100
                self.cal_percent.append(a)
            # print(data_sht_2s)
            #取出學生備註
            with coxn.cursor() as cursor:
                sql_tk2c = """SELECT [comment]
                FROM [bloodtest].[dbo].[test_data] 
                JOIN [bloodtest].[dbo].[id]
                ON [id].[No] = [test_data].[test_id]
                WHERE [test_id] IN (SELECT [id].[No] WHERE [id].[ac]='%s')
                AND [smear_id]='%s' 
                AND [count]=%d;"""%(self.row_data[1],self.row_data[2],self.row_data[3])
                cursor.execute(sql_tk2c)
                data_comment = cursor.fetchone()
            merge =[(x[0], x[1], x[2], x[3], y) for x, y in zip(self.data_sht_2, self.cal_percent)]
            # print(merge)
            #放入tksheets_2中
            self.sht_ans.dehighlight_all()
            self.sht_ans.headers(["細胞","答案數值","must_chk","mustnot_chk","考生DATA"])
            self.sht_ans.set_sheet_data(merge,reset_col_positions=True,reset_row_positions=True)
            #更改備註
            self.input_testcomment.configure(state='normal')
            self.input_testcomment.delete("0.0",ctk.END)
            try:
                self.input_testcomment.insert(ctk.END,data_comment[0])
            except tk.TclError:
                pass
            self.input_testcomment.configure(state='disable')
#####↑↑↑↑↑輸入tksheet_2: 將選好的考試答案，加入tksheet_2與答案一起比對↑↑↑↑↑#####

#####↓↓↓↓↓##################計算成績: 餵進去計算成績##################↓↓↓↓↓#####
    def calculate_cell(self):
        must_chk,mustnot_chk,ablym_chk,final_score = self.calculate(ans_tuple=self.data_sht_2,entry_lst=self.cal_percent,data_count_val=self.data_count_val)
        #重新放入tksheets_2中
        self.data_sht_2_lst = [list(element) for element in self.data_sht_2]    #原先的data_sht_2為tuple結構，不可以append，所以轉換為list
        for cellname in self.data_sht_2_lst:    #利用建立好的ansrangedict去查表，把range加進
            lowrange = self.ansrangedict[cellname[0]][1]
            uprange = self.ansrangedict[cellname[0]][2]
            cellname.append(lowrange)
            cellname.append(uprange)
        # print(self.data_sht_2_lst)
        merge = [(x[0], x[1], x[4], x[5], y,  x[2], x[3]) for x, y in zip(self.data_sht_2_lst, self.cal_percent)]
        self.sht_ans.set_sheet_data(merge,reset_col_positions=True,reset_row_positions=True)
        self.sht_ans.headers(["細胞","答案數值","上限","下限","考生DATA","must_chk","mustnot_chk"],redraw = True)
        
        #加入bg顏色輔助判讀(range)
        for ansno in range(0,len(self.chkrange)):
            if self.chkrange[ansno] != 1:
                #看是偏大還是偏小
                if self.data_sht_2_lst[ansno][1] > self.cal_percent[ansno]:
                    self.sht_ans.highlight_cells(row = ansno, column = 4,bg = '#92D050')
                else:
                    self.sht_ans.highlight_cells(row = ansno, column = 4,bg = '#FFCC99')
            else:
                pass
        #加入bg顏色輔助判讀(must & mustnot & ablym_chk)
        if must_chk =="" or must_chk ==True:
            pass
        else:
            #先把must的位置叫出來
            must_loc = [musttrue[2] for musttrue in self.data_sht_2_lst]
            must_loc = [index for index, value in enumerate(must_loc) if value]
            #再一一比對must是否符合
            for loc in must_loc:
                if self.data_sht_2_lst[loc][5] > self.cal_percent[loc]:
                    self.sht_ans.highlight_cells(row = loc, column = 4,bg = '#FF2D2D')
                else:
                    pass
        if mustnot_chk =="" or mustnot_chk ==True:
            pass
        else:
            #先把mustnot的位置叫出來           
            mustnot_loc = [mustnottrue[3] for mustnottrue in self.data_sht_2_lst]
            mustnot_loc = [index for index, value in enumerate(mustnot_loc) if value]
            #再一一比對mustnot是否符合
            for loc in mustnot_loc:
                if self.cal_percent[loc] > 0:
                    self.sht_ans.highlight_cells(row = loc, column = 4,bg = '#FF2D2D')
                else:
                    pass
        if ablym_chk =="" or ablym_chk ==True:
            pass
        else:
            self.sht_ans.highlight_cells(row = 0, column = 4,bg = '#FF2D2D')
            self.sht_ans.highlight_cells(row = 1, column = 4,bg = '#FF2D2D')

        tk.messagebox.showinfo(title='檢驗醫學部(科)', message="""已完成計算!
(空白表示不需檢查)
必須打到細胞檢查: %s
不可打到細胞檢查: %s
plasmacell+abn-Lym:%s
最後總分:%d"""%(must_chk,mustnot_chk,ablym_chk,final_score))
        #將三個boolin值轉換準備傳至sql
        chk_lst=[must_chk,mustnot_chk,ablym_chk]
        bool_lst = []
        for item in chk_lst:
            if item =="":
                bool_lst.append(f"NULL")
            elif item == 'False':
                bool_lst.append(0)
            else :
                bool_lst.append(1)
        self.must_chk = bool_lst[0]
        self.mustnot_chk = bool_lst[1]
        self.ablym_chk = bool_lst[2]
        # print(self.must_chk,self.mustnot_chk,self.ablym_chk)
        #加入tksheets
        # print(self.chkrange)
        #加入combobox_
        self.combobox_tscore.set(final_score)
#####↑↑↑↑↑##################計算成績: 餵進去計算成績##################↑↑↑↑↑#####
    def calculate(self,ans_tuple,entry_lst,data_count_val):
        # if self.chk_val()==False or self.chk_chkbox()==False:
        #     return
        anslst = [item[1] for item in ans_tuple]
        entrylst = entry_lst
        self.mustlst=[item[2] for item in ans_tuple]
        self.mustnotlst=[item[3] for item in ans_tuple]
        self.ansrangedict={item[0]:[item[1]] for item in ans_tuple}
        # print(anslst)
        # print(entrylst)
        # print(self.mustlst)        
        # print(self.mustnotlst)
        # print(self.ansrangedict)
        #收細胞 
        mtl,mnl=[],[]
        for k in range(0,len(self.mustlst)):
            if self.mustlst[k] == 0:
                pass
            else:
                mtl.append(k)
        for g in range(0,len(self.mustnotlst)):
            if self.mustnotlst[g] == 0:
                pass
            else:
                mnl.append(g)
        # print(mtl,mnl)
        #如果mustlist裡面有plasma cell or abn-lym的話，先把他們排除，再考慮他們有沒有扣分
        if 0 in mtl or 1 in mtl:
            if 0 in mtl:
                mtl.remove(0)
                plasmaabnormal = 0
            else:
                mtl.remove(1)
                plasmaabnormal = 1
        else:
            plasmaabnormal = "False"
        #設定n=100/200/500/1000
        ncount = data_count_val
        #設定好之後去json裡面調出相對應的格式db
        colname_1 = "%s_lower"%(str(ncount))
        colname_2 = "%s_upper"%(str(ncount))
        db = self.rangechart[[colname_1,colname_2]]
        #將解答的range填入self.ansrangedict
        for ansamount in self.ansrangedict.values():
            num = round(ansamount[0])
            tt = db.iloc[num]
            lower = tt[colname_1]
            upper = tt[colname_2]
            ansamount.append(lower)
            ansamount.append(upper)
        #18個細胞的matrix
        self.chkrange = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        #3個chk
        ablym_chk,must_chk,mustnot_chk = "","",""
        anscal, studentcal = [], []
        for value in anslst:
            anscal.append(float(value))
        for value in entrylst:
            studentcal.append(float(value))
        #Rule I: 必須要打到，不能打到細胞chk
        #檢查如果有plasma跟abnormal-lym的問題的話要先處理
        if plasmaabnormal!="False":
            #如果plasmaabnormal=0，也就是mustlst裡面有plasmacell
            if plasmaabnormal==0:
                num= round(anscal[0])
            #如果plasmaabnormal=1，也就是mustlst裡面有abnormallympho
            elif plasmaabnormal==1:
                num = round(anscal[1])
            #去抓rangechart中的相對整數
            tt = db.iloc[num]
            lower = tt[colname_1]
            cellmerge = studentcal[0] + studentcal[1]
            if cellmerge >= lower:
                ablym_chk = "True"
            else:
                ablym_chk = "False"
            #檢查必要cell
        for i in range(0,len(mtl)):
            num = round(anscal[mtl[i]])
            tt = db.iloc[num]
            lower = tt[colname_1]
            if studentcal[mtl[i]] >= lower:
                must_chk = "True"
            else:
                must_chk = "False"
                break
            #檢查不必要cell
        for j in range(0,len(mnl)):
            if studentcal[mnl[j]] != 0:
                mustnot_chk = "False"
                break
            else:
                mustnot_chk="True"
                continue
        # print("must:"+ must_chk)
        # print("mustnot:"+ mustnot_chk)
        # print("ab:" + ablym_chk)
        #如果兩個判定其中一個False，return final_score=0，其餘照舊
    #Rule II: 確定18個細胞都在Range中
        #exceptions: 如果plasma cell 跟abnormal-lym是有打勾的
        if plasmaabnormal ==0 or plasmaabnormal==1:
            c_idx = 2
        else:
            c_idx = 0
        for c in range(c_idx,len(anscal)):
            #取答案整數num，然後去db裡面抓取range
            num = round(anscal[c])
            tt = db.iloc[num]
            lower = tt[colname_1]
            upper = tt[colname_2]
            # print(tt)
            #check student_ans isin range
            std = studentcal[c]
            if std >= lower and std <= upper:
                self.chkrange[c] += 1
            else:
                pass
        #要檢查是不是只有plasma cell 跟abnormal lympho不再range內，如果是的話，綜合加起來一起評估
        if plasmaabnormal==0 or plasmaabnormal==1:
            if sum(self.chkrange[1::]) == 16:
                final_score = 2
            else:
                final_score = 1
        elif sum(self.chkrange) == 18:
            final_score = 2
        else:
            final_score = 1
        if must_chk == "False" or mustnot_chk == "False" or ablym_chk == "False":
            final_score=0
        else:
            pass
        # print(self.chkrange)
        # print(final_score)
    #Rule III: 看看有沒有blast & abn-Lym 重複出現chk
        if studentcal[1] > 0 and studentcal[4] > 0:
            final_score=0
            return must_chk,mustnot_chk,ablym_chk,final_score
        else:
            return must_chk,mustnot_chk,ablym_chk,final_score
#####↑↑↑↑↑##################計算成績: 餵進去計算成績##################↑↑↑↑↑#####
#####↓↓↓↓↓##############多次成績計算，自動上傳############↓↓↓↓↓#####
    def multi_upload(self):
        ##先檢查有沒有搜尋結果，如果沒有的話show error
        if self.sht_result.get_sheet_data()==[]:
            tk.messagebox.showerror(title='檢驗醫學部(科)',message="尚未選擇範圍!!請搜尋後再次啟用!") 
            return
        if tk.messagebox.askyesno(title='檢驗醫學部(科)',message="確定要多重計算後，自動上傳?"):
            multi_lst = pd.DataFrame(self.sht_result.get_sheet_data())
            ##迴圈搜尋結果
            for i in range(0,len(multi_lst)):
                #(multi)取出血片解答
                with coxn.cursor() as cursor:
                    sql_tk2 = """SELECT [celltype],[value],[must],[mustnot] 
    FROM [bloodtest].[dbo].[bloodinfo_ans2]
    WHERE [smear_id] ='%s';"""%(multi_lst.iloc[i][2])
                    cursor.execute(sql_tk2)
                    self.data_sht_2 = cursor.fetchall()
                    ##self.data_sht_2格式 =[('plasma cell',2.0,1,0)]
                #(multi)取出百分比
                with coxn.cursor() as cursor:
                    sql_per = "SELECT [count_value] FROM [bloodtest].[dbo].[bloodinfo] WHERE [smear_id]='%s';"%(multi_lst.iloc[i][2])
                    cursor.execute(sql_per)
                    n = cursor.fetchone()
                self.data_count_val= int(n[0])
                #(multi)取出學生考試答案
                with coxn.cursor() as cursor:
                    sql_tk2s = """SELECT [plasma cell],[abnormal lympho],[megakaryocyte],[nRBC],[blast],[metamyelocyte],[eosinophil],[plasmacytoid],[promonocyte],[promyelocyte],[band neutropil],[basopil],[atypical lymphocyte],[hypersegmented neutrophil],[myelocyte],[segmented neutrophil],[lymphocyte],[monocyte]
                    FROM [bloodtest].[dbo].[test_data] 
                    JOIN [bloodtest].[dbo].[id]
                    ON [id].[No] = [test_data].[test_id]
                    WHERE [test_id] IN (SELECT [id].[No] WHERE [id].[ac]='%s')
                    AND [smear_id]='%s' 
                    AND [count]=%d;"""%(multi_lst.iloc[i][1],multi_lst.iloc[i][2],multi_lst.iloc[i][3])
                    cursor.execute(sql_tk2s)
                    data_sht_2s = cursor.fetchone()
                #(multi)百分比轉換
                self.cal_percent = [] 
                for j in range(0,len(data_sht_2s)):
                    #data_count_value為考片所需打的顆數
                    ##data_sht_2s[j]為考生每一個細胞(raw data)
                    #相除即為百分比
                    a = data_sht_2s[j] / self.data_count_val *100   
                    self.cal_percent.append(a)
                ##丟進去calculate計算後傳出結果，must/mustnot/ablym回傳布林值，final score回傳[0,1,2]
                must_chk,mustnot_chk,ablym_chk,final_score = self.calculate(ans_tuple=self.data_sht_2,entry_lst=self.cal_percent,data_count_val=self.data_count_val)
                # print(must_chk,mustnot_chk,ablym_chk,final_score)
                ## NULL轉換，如果值為""的話，需要轉換為文字NULL，不然無法上傳SQL
                chk_lst=[must_chk,mustnot_chk,ablym_chk]
                bool_lst = []
                for item in chk_lst:    #依序把chk_lst中的布林值更換為NULL
                    if item =="":
                        bool_lst.append(f"NULL")
                    elif item == 'False':
                        bool_lst.append(0)
                    else :
                        bool_lst.append(1)
                # print(bool_lst)
                self.must_chk = bool_lst[0]
                self.mustnot_chk = bool_lst[1]
                self.ablym_chk = bool_lst[2]
                ##(multi)上傳第一個表格[bloodtest].[dbo].[test_data]
                ##更新表格後面的: 必須打到的細胞布林值[must]、不可打到的細胞布林值[mustnot]、abnormal lympho跟plasma cell的布林值[abn_lym]、總分[score],
                with coxn.cursor() as cursor:
                    upload_final = """UPDATE [bloodtest].[dbo].[test_data]
    SET [test_data].[score] =%d , [test_data].[must] = %s,[test_data].[mustnot] = %s, [test_data].[abn_lym] = %s
    WHERE [test_data].[test_id]
    IN(SELECT [id].[No] FROM [bloodtest].[dbo].[id] WHERE [id].[ac]='%s')
    AND [test_data].[smear_id]='%s'
    AND [test_data].[count] = %d;"""%(final_score,self.must_chk,self.mustnot_chk,self.ablym_chk,multi_lst.iloc[i][1],multi_lst.iloc[i][2],multi_lst.iloc[i][3])
                    cursor.execute(upload_final)
                coxn.commit()
                ##上傳第二個表格
                #先找ac對應的id
                with coxn.cursor() as cursor:
                    cursor.execute("SELECT [No] FROM [bloodtest].[dbo].[id] WHERE [id].[ac] = '%s'"%(multi_lst.iloc[i][1]))
                    t_id = cursor.fetchone()
                t_id = t_id[0]
                #跟細胞名稱zip一起
                cell_item =['plasma cell','abnormal lympho','megakaryocyte','nRBC','blast','metamyelocyte','eosinophil','plasmacytoid','promonocyte','promyelocyte','band neutropil','basopil','atypical lymphocyte','hypersegmented neutrophil','myelocyte','segmented neutrophil','lymphocyte','monocyte']
                cell_score_matrix = list(zip(cell_item,self.chkrange,self.cal_percent))
                # print(cell_score_matrix)
                #迴圈上傳(為了不要有重複值的關係，所以需要加入IF ELSE判斷句)
                for j in range(0,len(cell_score_matrix)):
                    with coxn.cursor() as cursor:
                        upload_final_2 = """IF NOT EXISTS (
                                SELECT [smear_id],[count],[celltype]
                                FROM [bloodtest].[dbo].[blood_final] 
                                WHERE [smear_id]='%s' AND [count]=%d AND [test_id]= %d AND [celltype]= '%s'
                                )
                            INSERT INTO [bloodtest].[dbo].[blood_final] ([test_id],[smear_id],[count],[celltype],[matrix_value],[percent_value])
                            VALUES(%d,'%s',%d,'%s',%d,%.2f)
                        ELSE 
                            UPDATE [bloodtest].[dbo].[blood_final] 
                            SET [matrix_value]=%d,[percent_value]=%.2f
                            WHERE[smear_id]='%s' AND [count]=%d AND [test_id]=%d AND [celltype]='%s';"""%(multi_lst.iloc[i][2],multi_lst.iloc[i][3],t_id,cell_score_matrix[j][0],t_id,multi_lst.iloc[i][2],multi_lst.iloc[i][3],cell_score_matrix[j][0],cell_score_matrix[j][1],cell_score_matrix[j][2],cell_score_matrix[j][1],cell_score_matrix[j][2],multi_lst.iloc[i][2],multi_lst.iloc[i][3],t_id,cell_score_matrix[j][0])
                        cursor.execute(upload_final_2)
                    coxn.commit()
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message="搜尋結果已全部計算總成績且完成上傳!!")
            self.clear()
            return
#####↑↑↑↑↑##############多次成績計算，自動上傳#############↑↑↑↑↑#####
#####↓↓↓↓↓##############成績上傳: 上傳至sql [test_data]##############↓↓↓↓↓#####
    def upload_score(self):
        ##需要資訊:院區、考核人員、考片ID、考核次數、最後總分
        final_score = int(self.combobox_tscore.get())
        show = [x for x in self.row_data]
        if tk.messagebox.askyesno(title='檢驗醫學部(科)',message="""請確認以下資訊:
院區: %s
考核人員: %s
考片ID: %s
考核次數: %d
上傳總分: %d                        
確定要上傳成績?"""%(show[0],show[1],show[2],show[3],final_score)):
            ##上傳第一個表格[bloodtest].[dbo].[test_data]
            ##更新表格後面的: 必須打到的細胞布林值[must]、不可打到的細胞布林值[mustnot]、abnormal lympho跟plasma cell的布林值[abn_lym]、總分[score],
            with coxn.cursor() as cursor:
                upload_final = """UPDATE [bloodtest].[dbo].[test_data]
SET [test_data].[score] =%d , [test_data].[must] = %s,[test_data].[mustnot] = %s, [test_data].[abn_lym] = %s
WHERE [test_data].[test_id]
IN(SELECT [id].[No] FROM [bloodtest].[dbo].[id] WHERE [id].[ac]='%s')
AND [test_data].[smear_id]='%s'
AND [test_data].[count] = %d;"""%(final_score,self.must_chk,self.mustnot_chk,self.ablym_chk,show[1],show[2],show[3])
                # print(upload_final)
                cursor.execute(upload_final)
            coxn.commit()
            ##上傳第二個表格
            
            #先找ac對應的id
            with coxn.cursor() as cursor:
                cursor.execute("SELECT [No] FROM [bloodtest].[dbo].[id] WHERE [id].[ac] = '%s'"%(show[1]))
                t_id = cursor.fetchone()
            t_id = t_id[0]
            #跟細胞名稱zip一起
            cell_item =['plasma cell','abnormal lympho','megakaryocyte','nRBC','blast','metamyelocyte','eosinophil','plasmacytoid','promonocyte','promyelocyte','band neutropil','basopil','atypical lymphocyte','hypersegmented neutrophil','myelocyte','segmented neutrophil','lymphocyte','monocyte']
            cell_score_matrix = list(zip(cell_item,self.chkrange,self.cal_percent))
            # print(cell_score_matrix)
            #迴圈上傳(為了不要有重複值的關係，所以需要加入IF ELSE判斷句)
            for j in range(0,len(cell_score_matrix)):
                with coxn.cursor() as cursor:
                    # upload_final_2 = """INSERT INTO bloodtest.dbo.blood_final ([test_id],[smear_id],[count],[celltype],[matrix_value],[percent_value])
                    # VALUES(%d,'%s',%d,'%s',%d,%.2f);"""%(t_id,show[2],show[3],cell_score_matrix[j][0],cell_score_matrix[j][1],cell_score_matrix[j][2])
                    upload_final_2 = """IF NOT EXISTS (
                                        SELECT [smear_id],[count],[celltype]
                                        FROM [bloodtest].[dbo].[blood_final] 
                                        WHERE [smear_id]='%s' AND [count]=%d AND [test_id]= %d AND [celltype]= '%s'
                                        )
                                INSERT INTO [bloodtest].[dbo].[blood_final] ([test_id],[smear_id],[count],[celltype],[matrix_value],[percent_value])
                                VALUES(%d,'%s',%d,'%s',%d,%.2f)
                            ELSE 
                                UPDATE [bloodtest].[dbo].[blood_final] 
                                SET [matrix_value]=%d,[percent_value]=%.2f
                                WHERE[smear_id]='%s' AND [count]=%d AND [test_id]=%d AND [celltype]='%s';"""%(show[2],show[3],t_id,cell_score_matrix[j][0],t_id,show[2],show[3],cell_score_matrix[j][0],cell_score_matrix[j][1],cell_score_matrix[j][2],cell_score_matrix[j][1],cell_score_matrix[j][2],show[2],show[3],t_id,cell_score_matrix[j][0])
                    cursor.execute(upload_final_2)
                coxn.commit()
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message="已完成上傳!")
            self.upload_clear()
            return
        else:
            pass
#####↑↑↑↑↑##############成績上傳: 上傳至sql [test_data]##############↑↑↑↑↑#####
##清除按鈕
    def clear(self):
        ##清除兩個combobox
        self.all_combobox.set("")
        self.year_combobox.set("")
        ##清除兩個tksheets
        self.sht_result.set_sheet_data()
        self.sht_ans.set_sheet_data()
        ##清除一個備註內容
        self.input_testcomment.configure(state='normal')
        self.input_testcomment.delete("0.0",ctk.END)
        self.input_testcomment.configure(state='disable')
        ##清除總分combobox
        self.combobox_tscore.set("")
##清除上傳後reload
    def upload_clear(self):
        ##清除兩個tksheets
        self.sht_result.set_sheet_data()
        self.sht_ans.set_sheet_data()
        ##清除一個備註內容
        self.input_testcomment.configure(state='normal')
        self.input_testcomment.delete("0.0",ctk.END)
        self.input_testcomment.configure(state='disable')
        ##清除總分combobox
        self.combobox_tscore.set("")
        ##讓clear重新用搜尋reload一次資料
        self.search_gai()
##返回basdesk按鈕    
    def back(self,oldmaster):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return


def main():  
    root = ctk.CTk()
    SC = SCORE_CAL(root)
    # root['bg']='#FFEEDD'
    # B.gui_arrang()
    # 主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  