import sys,os,io,random
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox
import customtkinter as ctk
from collections import defaultdict
import json,pyodbc
from PIL import Image,ImageTk

#抓取帳號
def getaccount(acount):
    global Baccount
    Baccount = str(acount) 
    # Baccount = "henry423"
    # print(Baccount)
    return None
def getlevel(level):
    global test_level
    test_level = int(level) 
    return None
##連線SQL
connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=220.133.50.28;DATABASE=bloodtest;UID=cgmh;PWD=B[-!wYJ(E_i7Aj3r"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.InterfaceError:
    # connection_string = "DRIVER={ODBC Driver 11 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
    connection_string = "DRIVER={SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)
##取得細胞名稱
with coxn.cursor() as cursor:
    unique_celltype = "SELECT DISTINCT [celltype] FROM [bloodtest].[dbo].[cell];"
    cursor.execute(unique_celltype)
    celltype = cursor.fetchall()
celltype = [e[0] for e in celltype]
##取得細胞簡介
with coxn.cursor() as cursor:
    sql_cellinfo = "SELECT [cell_name],[info] FROM [bloodtest].[dbo].[cell_info];"
    cursor.execute(sql_cellinfo)
    cell_info = cursor.fetchall()
cell_info = {key: value.replace("\r", "") for key, value in cell_info}
# print(cell_info)

##利用細胞名稱，隨機取三個不同種類的細胞
test_lst =[]
def generate_test(test_level):
    global test_no,test_final,test_dict
    for item in celltype:
        with coxn.cursor() as cursor:
            if test_level==1:   #難度為初階的話
                pass
            else:   #難度為中階
                #如果是太難太少的細胞，就選兩個就好
                if item =="atypical lymphocyte" or item=="abnormal lympho" or item=="promonocyte" or item=="basopil":
                    ans = """SELECT TOP 2 [celltype],[image] FROM [bloodtest].[dbo].[cell]
                            WHERE [celltype]= '%s' 
                            ORDER BY NEWID();"""%(item)
                else:
                    ans = """SELECT TOP 3 [celltype],[image] FROM [bloodtest].[dbo].[cell]
                            WHERE [celltype]= '%s' 
                            ORDER BY NEWID();"""%(item)
            cursor.execute(ans)
            ans = cursor.fetchall()
        for i in ans:
            test_lst.append(i)
    ##生成1-50隨機且不重複的序列號碼列表，當作題目編號
    test_no = random.sample(range(1,51),50)
    test_final = zip(test_no,test_lst)
    test_dict = {key: value for key, value in test_final}
    ##test_dict格式 test_dict={1:["abnormal lymphocyte",<圖片二進位的碼>]}
    # print(test_lst)
    # print(test_dict)


class IMAGEPRACTICE:

    def __init__(self,master,oldmaster=None):
        self.master = master
        super().__init__()
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.exit(oldmaster))
        ctk.set_default_color_theme("blue")  
        # 給主視窗設定標題內容  
        self.master.title("圖庫練習")  
        self.master.geometry('950x650')
        self.master.grid_rowconfigure(0, weight=1)
        # self.master.grid_columnconfigure(0, weight=3)
        self.master.config(background='#FFEEDD') #設定背景色
        ##框架設置
        self.labelframe_1 = ctk.CTkFrame(self.master, corner_radius=0,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.labelframe_1.grid(row=0, column=0,columnspan=2, sticky="nsew")
        # self.labelframe_1.grid_rowconfigure(0, weight=1)
        self.labelframe_1.grid_columnconfigure(0, weight=1)
        self.labelframe_2 = ctk.CTkFrame(self.master,fg_color="#FFDCB9",bg_color="#FFEEDD")
        self.labelframe_2.grid(row=1, column=0,columnspan=2,sticky="nsew")
        self.labelframe_3 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.labelframe_result = ctk.CTkFrame(self.master,fg_color="#FFDCB9",bg_color="#FFEEDD")
        if test_level==1:
            self.label_start = ctk.CTkLabel(
                self.labelframe_1, 
                fg_color="#FFEEDD",
                text='圖庫練習_初階',
                text_color="#000000",
                font=("微軟正黑體",60),
                width=120)
        else:
            self.label_start = ctk.CTkLabel(
                self.labelframe_1, 
                fg_color="#FFEEDD",
                text='圖庫練習_中階',
                text_color="#000000",
                font=("微軟正黑體",60),
                width=120)
        self.label_start.grid(row =0, column = 0,pady=100, sticky="ew",columnspan=2)
        self.btn_start = ctk.CTkButton(
            self.labelframe_1,
            command = self.start, 
            text = "開始測驗", 
            fg_color='#E6883B',
            hover_color='#FFCC99',
            width=200,height=60,
            corner_radius=8,
            font=('微軟正黑體',22,'bold'),
            text_color="#000000"
            )
        self.btn_start.grid(row=2, column=0,columnspan=2)


        ##建立按鍵與細胞連結
        #字典結構:{key=輸入器輸入字母:value=[x,y,細胞全名,細胞簡稱,連結計數器,百分比]

        self.keybordmatrix = {
        'I':[0,0,1,'test','TEST',],
        '4':[0,1,2,'plasma cell','PLASMA',],
        '3':[0,2,3,'abnormal lympho','AB-LYM',],
        '2':[0,3,4,'megakaryocyte','MEGAKA',],
        '1':[0,4,5,'nRBC','N-RBC',],
        'R':[1,1,6,'blast','BLAST',],
        'E':[1,2,7,'metamyelocyte','META',],
        'W':[1,3,8,'eosinophil','EOSIN',],
        'Q':[1,4,9,'giant plt','GIANTPLT',],
        'G':[2,0,10,'promonocyte','PROMO',],
        'F':[2,1,11,'promyelocyte','PROMY',],
        'D':[2,2,12,'band neutropil','BAND',],
        'S':[2,3,13,'basopil','BASO',],
        'A':[2,4,14,'atypical lymphocyte','AT-LYM',],
        'B':[3,0,15,'hypersegmented neutrophil','HYPERSEG',],
        'V':[3,1,16,'myelocyte','MYELO',],
        'C':[3,2,17,'segmented neutrophil','SEG',],
        'X':[3,3,18,'lymphocyte','LYM',],
        'Z':[3,4,19,'monocyte','MONO',]
        }
        self.ans_var = tk.StringVar()
        #細胞的radio_btn矩陣
        self.radiolst =[]
        for value in self.keybordmatrix.values():
            if value[2]==1: #skip過test_btn
                pass
            else:
                R1 = ctk.CTkRadioButton(
                    self.labelframe_2,
                    variable= self.ans_var,
                    value=value[2],
                    text=value[3],
                    font=('微軟正黑體',14),
                    text_color="#000000",
                    fg_color='#F55536',
                    width=180,height=25,
                    state='disabled'
                    )
                R1.grid(row=value[0], column=value[1],padx=5,pady=4,sticky='nsew')
                R1.grid_propagate(0)
                self.radiolst.append(R1)
        ##答案儲存dict
        self.tempt_ans={}
        for i in range(1,51):
            self.tempt_ans[i]= tk.IntVar()
        #交卷按鈕
        self.btn_submmit = ctk.CTkButton(
            self.master,
            text="交卷",
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            command= self.submmit_1,
            bg_color="#F0F0F0",
            fg_color="#FF8000",
            corner_radius=5,
            width=80,height=30,
            state='disabled'
        )
        self.btn_submmit.grid(row=2,column=1,padx=5,pady=4,sticky='nsew',)
        # self.btn_submmit.grid(row=4, column=3,columnspan=2,padx=5,pady=4,sticky='nsew')
        
        #返回主介面
        self.btn_exit = ctk.CTkButton(
            self.master,
            text="返回",
            font=('微軟正黑體',22,'bold'),
            command= lambda: self.exit(oldmaster),
            bg_color="#F0F0F0",
            fg_color="#FF8000",
            corner_radius=5,
            width=80,height=30
        )
        self.btn_exit.grid(row=2,column=0,padx=5,pady=4,sticky='nsew')
        # self.btn_exit.grid(row=4, column=0,padx=5,pady=4,sticky='nsew')
        generate_test(test_level)



    def start(self):
        self.labelframe_3.grid(row=0, column=0,columnspan=2,sticky='nsew')
        self.labelframe_3.grid_columnconfigure(2,weight=1)
        ##設定鍵盤快捷鍵
        self.master.bind('<Key>', self.bind_select)
        for item in self.radiolst:
            item.configure(state='normal')
        self.btn_submmit.configure(state='normal')
        self.label_question = ctk.CTkLabel(
                    master = self.labelframe_3, 
                    text = "題目:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=60
                    )
        self.label_question.grid(row=0,column=0)
        self.combobox_no = ctk.CTkComboBox(
                    master = self.labelframe_3, 
                    values=[str(n) for n in range(1,51)], 
                    width=70,height=30,
                    button_color='#FF9900',
                    command=self.no_callback,
                    )
        self.combobox_no.grid(row=0,column=1,pady=10,sticky='e')
        self.label_no = ctk.CTkLabel(
                    master = self.labelframe_3, 
                    text = "/50", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=60
                    )
        self.label_no.grid(row=0,column=2,sticky='w')
        self.label_click = ctk.CTkLabel(
                    master = self.labelframe_3, 
                    text = "(點擊照片即可放大)", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',22),
                    text_color="#000000",
                    width=150,height=30
                    )
        self.label_click.grid(row=1,column=0,columnspan=2,padx=10)
        self.btn_prev = ctk.CTkButton(
            self.labelframe_3,
            text="▲上一題",
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            command= self.prev_pic,
            bg_color="#F0F0F0",
            fg_color="#FF8000",
            corner_radius=5,
            width=150,height=30
        )
        self.btn_prev.grid(row=3,column=0,columnspan=2,pady=10)
        self.btn_next = ctk.CTkButton(
            self.labelframe_3,
            text="▼下一題",
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            command= self.next_pic,
            bg_color="#F0F0F0",
            fg_color="#FF8000",
            corner_radius=5,
            width=150,height=30
        )
        self.btn_next.grid(row=3,column=4,padx=15,pady=10)
        self.fr_canvas = ctk.CTkImage(Image.open(io.BytesIO(test_dict[1][1])),size=(368,369))
        self.image = ctk.CTkLabel(  
                    self.labelframe_3, 
                    image=self.fr_canvas,
                    text="",
                    # width=120,height=120,
                    fg_color='#000000',
                    bg_color='#000000',
                    )
        self.image.grid(row=1,column=2,rowspan=3,padx=10,pady =5,)
        self.image .bind("<Double-Button-1>", self.clickzoomin)
    def clickzoomin(self,event):
        ##考慮到後會加上"*"導致ValueError，所以用try...except
        try:
            no_now = int(self.combobox_no.get())
        except ValueError:
            no_now = int(self.combobox_no.get().rstrip("*"))
        window_zoomin = ctk.CTkToplevel()
        window_zoomin.title("圖片放大")
        window_zoomin.geometry("600x600")
        fr_zoomin = ctk.CTkFrame(window_zoomin)
        self.zoomin_img = ctk.CTkImage(Image.open(io.BytesIO(test_dict[no_now][1])),size=(600,600))
        self.label_zoomin = ctk.CTkLabel(  
                    window_zoomin, 
                    image=self.zoomin_img,
                    text="",
                    width=600,height=600,
                    fg_color='#000000',
                    bg_color='#000000',
                    )
        self.label_zoomin.pack()
    
    def no_callback(self,event):
        self.master.focus_set()
        ##考慮到後會加上"*"導致ValueError，所以用try...except
        try:
            no_now = int(self.combobox_no.get())
        except ValueError:
            no_now = int(self.combobox_no.get().rstrip("*"))
        # self.combobox_no.set(str(no_now))
        ##換照片
        self.fr_canvas.configure(light_image = Image.open(io.BytesIO(test_dict[no_now][1])))
        ##把記憶答案叫出來
        q_2 = int(self.tempt_ans[no_now].get())
        # print(q_2)
        if q_2 != 0:
            self.radiolst[q_2-1].select()
        else:
            self.ans_var.set(0)
        
        try:
            self.var_info.set(cell_info[self.std_ans[no_now]])
            self.var_stdans.set(self.std_ans[no_now])   #標準答案dict
            self.var_yrans.set(self.lst_yrans[no_now-1])    #學生的答案
        except AttributeError:
            pass
        return
    ##<key>快捷鍵選取
    def bind_select(self,event):
        ##slc = 按鍵選擇的按鈕號碼
        slc = self.keybordmatrix[str.upper(event.char)][2]
        self.radiolst[slc-1].select()
        self.next_pic()
    ##功能下一題
    def next_pic(self):
        no_now = int(self.combobox_no.get())
        if no_now < 50:
            ##記憶答案內容
            q_1 =int(self.ans_var.get())
            self.tempt_ans[no_now].set(q_1)
            self.ans_var.set(0)
            # print(self.tempt_ans[no_now].get())
            no_now += 1
        else:
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message="這一題為最後一題囉")
            return
        ##換照片
        self.combobox_no.set(str(no_now))
        self.fr_canvas.configure(light_image = Image.open(io.BytesIO(test_dict[no_now][1])))
        ##把記憶的答案內容叫出來，如果沒有的話，那就給空值(0)
        q_2 = int(self.tempt_ans[no_now].get())
        # print(q_2)
        if q_2 != 0:
            self.radiolst[q_2-1].select()
        else:
            pass
        # self.image.configure(image=self.fr_canvas)
        return
    ##解答下一題
    def next_ans(self):
        ##考慮到後會加上"*"導致ValueError，所以用try...except
        try:
            no_now = int(self.combobox_no.get())
        except ValueError:
            no_now = int(self.combobox_no.get().rstrip("*"))
        if no_now < 50:
            no_now += 1
            self.var_info.set(cell_info[self.std_ans[no_now]])
        else:
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message="這一題為最後一題囉")
            return
        #先知道下一題是錯是對，要從self.lst_tf去查
        if self.lst_tf[no_now-1]==0:
            self.combobox_no.set(str(no_now) + "*") #換combobox的題號
        else:
            self.combobox_no.set(str(no_now)) #換combobox的題號
        self.fr_canvas.configure(light_image = Image.open(io.BytesIO(test_dict[no_now][1])))    ##換照片
        ##從標準答案裡面抓答案跟學生答案裡面撈他寫的答案
        self.var_stdans.set(self.std_ans[no_now])   #標準答案dict
        self.var_yrans.set(self.lst_yrans[no_now-1])  #學生的答案

    ##功能上一題
    def prev_pic(self):
        no_now = int(self.combobox_no.get())
        if no_now > 1:
            ##記憶答案內容
            q_1 = self.ans_var.get()
            self.tempt_ans[no_now].set(q_1)
            self.ans_var.set(0)
            no_now -= 1
        else:
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message="這一題為第一題喔")
            return
        ##換照片
        self.combobox_no.set(str(no_now))
        self.fr_canvas.configure(light_image = Image.open(io.BytesIO(test_dict[no_now][1])))
        ##把記憶的答案內容叫出來，如果沒有的話，那就給空值(0)
        q_2 = int(self.tempt_ans[no_now].get())
        # print(q_2)
        if q_2 != 0:
            self.radiolst[q_2-1].select()
        else:
            pass
        return
    ##解答上一題
    def prev_ans(self):
        ##考慮到後會加上"*"導致ValueError，所以用try...except
        try:
            no_now = int(self.combobox_no.get())
        except ValueError:
            no_now = int(self.combobox_no.get().rstrip("*"))
        if no_now > 1:
            no_now -= 1
            self.var_info.set(cell_info[self.std_ans[no_now]])
        else:
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message="這一題為第一題喔")
            return
        #先知道下一題是錯是對，要從self.lst_tf去查
        if self.lst_tf[no_now-1]==0:
            self.combobox_no.set(str(no_now) + "*") #換combobox的題號
        else:
            self.combobox_no.set(str(no_now)) #換combobox的題號
        self.fr_canvas.configure(light_image = Image.open(io.BytesIO(test_dict[no_now][1])))    ##換照片
        ##從標準答案裡面抓答案跟學生答案裡面撈他寫的答案
        self.var_stdans.set(self.std_ans[no_now])    #標準答案dict
        self.var_yrans.set(self.lst_yrans[no_now-1])  #學生的答案

    ##(f2)交卷
    def submmit_1(self):
        #把所有學生的答案list再一起
        all_ans=[]
        for value in self.tempt_ans.values():
            a = value.get()
            all_ans.append(a)
        # print(all_ans)
        #建立未填答的list
        none_ans  = [index+1 for (index, item) in enumerate(all_ans) if item == 0]
        ##檢查是否有未填答的
        if len(none_ans)!=0:
            if tk.messagebox.askyesno(title='檢驗醫學部(科)', message="""尚未答題:
%s
要繼續交卷嗎?"""%(str(none_ans))):
                self.submmit_final(all_ans)
        else:
            if tk.messagebox.askyesno(title='檢驗醫學部(科)', message="確定交卷?"):
                self.submmit_final(all_ans)
    ##改考卷
    def submmit_final(self,input_ans):
        #把交卷按鈕btn_submmit鎖起來
        self.btn_submmit.configure(state='disabled')
        #建立一個dictionary，std_ans = {key = 題號, value = str(正確答案)}
        self.std_ans={}
        for key, value in test_dict.items():
            self.std_ans[key] = value[0]
        # print(self.std_ans)
        #再建立一個dict給上面str(正確答案)轉換成數字
        ans_no={}
        for value in self.keybordmatrix.values():
            ans_no[value[3]] = value[2]
        no_ans ={}
        for value in self.keybordmatrix.values():
            no_ans[value[2]] = value[3]
        # print(ans_no)
        #利用for迴圈，把上面std_ans原本是string的正確答案換成int
        std_ans_no={}
        for key in self.std_ans:
            std_ans_no[key] = ans_no[self.std_ans[key]]
        #把input_ans轉換成str在檢討的時候會用到
        self.lst_yrans=[]
        for k in range(0,len(input_ans)):
            if input_ans[k] == 0:
                self.lst_yrans.append("未選擇")    
            else:
                self.lst_yrans.append(no_ans[input_ans[k]])
        ##跟input_ans對答案，也就是學生的答案
        #學生的答案是list，現在要用dictionary對答案，
        # print(self.lst_yrans)
        self.statistics= defaultdict(list)  #建立self.statistics={}提供上傳至SQL
        self.lst_tf = []      #self.lst_tf:指儲存true or false的訊息
        for j in range(0,len(input_ans)):
            r_ans = std_ans_no[j+1]    #標準答案
            name_ans = self.std_ans[j+1]    #標準答案(string)
            student_ans = input_ans[j]  #學生答案
            if r_ans == student_ans:    #比對相不相吻合
                self.lst_tf.append(1)
                self.statistics[name_ans].append(1)
            else:
                self.lst_tf.append(0)
                self.statistics[name_ans].append(0)
        count_true = self.lst_tf.count(1)
        
        # print(self.statistics)
        self.upload_cell_result()   #上傳至SQL的definition
        tk.messagebox.showinfo(title='檢驗醫學部(科)', message="""恭喜完成!
共50題，
答對 %d 題
答錯 %d 題"""%(count_true,50-count_true))
        if tk.messagebox.askyesno(title='檢驗醫學部(科)', message="是否要觀看檢討?"):
            self.view_result()
        else:
            return
    ##建立一個新的介面，把答案show在上面 
    def view_result(self):
        self.labelframe_2.grid_forget() #把radio_button
        self.master.unbind("<Key>") #解除按鍵關聯
        self.labelframe_result.grid(row=1, column=0,columnspan=2,sticky="nsew")
        self.fr_canvas.configure(size=(270,270))
        self.master.grid_rowconfigure(1, weight=2)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)
        ##建立兩個stringvar，放入學生答案跟標準答案的
        self.var_stdans = tk.StringVar()
        self.var_yrans = tk.StringVar()
        self.var_info = tk.StringVar()
        ##更改上/下一題的function，從next_pic改成next_ans
        self.btn_prev.configure(command=self.prev_ans)
        self.btn_next.configure(command=self.next_ans)
        ##更改combobox_no裡面的數字，改成錯誤的會打"*"
        new_combobox_no=[]
        for i in range(1,51):
            if self.lst_tf[i-1] == 0:
                new_combobox_no.append(str(i) + "*")
            else:
                new_combobox_no.append(str(i))
        # print(new_combobox_no)
        self.combobox_no.configure(values=new_combobox_no)
        
        ##標籤設計
        self.label_yrans = ctk.CTkLabel(
                    master = self.labelframe_result, 
                    text = "你的答案:", 
                    fg_color='#FFDCB9',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=60
                    )
        self.label_yrans.grid(row=0, column=0)
        self.label_stdans = ctk.CTkLabel(
                    master = self.labelframe_result, 
                    text = "標準答案:", 
                    fg_color='#FFDCB9',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=60
                    )
        self.label_stdans.grid(row=0, column=2)
        self.input_yrans = ctk.CTkLabel(
                self.labelframe_result,
                width=120,
                textvariable=self.var_yrans,
                bg_color='#FFDCB9',
                font=('微軟正黑體',24),
                text_color="#000000"
                )
        self.input_yrans.grid(row=0, column=1,padx=5,sticky='w')
        self.input_stdans = ctk.CTkLabel(
                self.labelframe_result,
                width=120,
                textvariable = self.var_stdans,
                bg_color='#FFDCB9',
                font=('微軟正黑體',24),
                text_color="#000000"
                )
        self.input_stdans.grid(row=0, column=3,padx=5,sticky='w')
        self.input_info = ctk.CTkLabel(
                self.labelframe_result,
                textvariable = self.var_info,
                bg_color='#FFDCB9',
                font=('微軟正黑體',20),
                text_color="#000000",
                anchor='n',
                justify='left',
                height=120
                )
        self.input_info.grid(row=1,column=0,columnspan=4,padx=5,sticky='w')
        ##歸第1題
        if self.lst_tf[0] == 1: #要先檢查第一題有沒有錯
            self.combobox_no.set(1) #combobox歸1
        else:
            self.combobox_no.set("1*") #combobox歸1*
        self.fr_canvas.configure(light_image = Image.open(io.BytesIO(test_dict[1][1]))) #圖片改到第一張
        self.var_stdans.set(self.std_ans[1])   #標準答案dict
        self.var_yrans.set(self.lst_yrans[0])  #學生的答案
        self.var_info.set(cell_info[self.std_ans[1]])  #細胞解釋
    ##上傳結果至SQL中
    def upload_cell_result(self):
        ##確認輸入者id
        with coxn.cursor() as cursor:
            cursor.execute("SELECT No FROM[bloodtest].[dbo].[id] WHERE [ac]='%s';" %(Baccount))
            ac = cursor.fetchone()[0]
        ##檢查為第幾次練習
        with coxn.cursor() as cursor:
            cursor.execute("SELECT MAX(count) FROM[bloodtest].[dbo].[cell_test_data] WHERE [id]=%d;" %(ac))
            sql_count = cursor.fetchone()[0]
        #如果練習次數不等於空值的話，那就繼續壘加上去
        if sql_count != None:
            sql_count +=1
        else:   #如果是空值的話，帶入第一次
            sql_count =1
        #先將defaultdic轉換為一般dic
        self.statistics = dict(self.statistics)
        #將value中的list結果變成百分比(0,0.33,0.5,0.66)
        dict_percent_statistics={}
        for key, value in self.statistics.items():
            per_result = round(sum(value)/len(value),2)
            dict_percent_statistics[key] = per_result

        print(dict_percent_statistics)
        for key, value in dict_percent_statistics.items():
            #建立上傳SQL語句
            sql_upload_var="""INSERT INTO [bloodtest].[dbo].[cell_test_data] ("id","count","celltype","rate") 
                VALUES (%d,%d,'%s',%.2f);"""%(ac,sql_count,key,value)
            with coxn.cursor() as cursor:
                cursor.execute(sql_upload_var)
            coxn.commit()
        
        # return
    ##返回
    def exit(self,oldmaster):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return



def main():  
    root = ctk.CTk()
    IP = IMAGEPRACTICE(root)
    #  主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  