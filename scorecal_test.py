import sys,os
import tkinter as tk
import customtkinter  as ctk
from tksheet import Sheet
from tkinter import ttk
from tkinter import RIDGE, DoubleVar, StringVar, ttk, messagebox,IntVar, filedialog
from turtle import width
import pandas as pd
import pyodbc

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=220.133.50.28;DATABASE=bloodtest;UID=cgmh;PWD=B[-!wYJ(E_i7Aj3r"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.InterfaceError:
    # connection_string = "DRIVER={ODBC Driver 11 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
    connection_string = "DRIVER={SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)

with coxn.cursor() as cursor:
    cursor.execute("SELECT [celltype] FROM [bloodtest].[dbo].[bloodinfo_ans2] WHERE [smear_id] = 'B0000_0000';")
    Ans = cursor.fetchall()
Ans = [e[0] for e in Ans]
# def getaccount(acount):
#     global Baccount
#     Baccount = str(acount)
#     # print(Baccount)
#     return None

class SCORE_TEST:

    def __init__(self,master,oldmaster=None):
        self.master = master
        super().__init__()
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.back(oldmaster))        
        ctk.set_default_color_theme("blue")
        self.master.title("考題成績試算")
        self.master.geometry("1600x850")
        self.master.config(background='#FFEEDD')
        filename = resource_path("./chart/rangechart.json")
        self.rangechart = pd.read_json(filename)
        # jsonfile = open('testdata\data_new.json','rb')
        # rawdata = json.load(jsonfile)
        # self.rawdata = pd.DataFrame(rawdata["blood"])
        # print(self.rawdata)
        # self.celllst = self.rawdata["Ans"][0]
        self.celllst = Ans
        # print(self.celllst)
        global anslst,entrylst, mustlst, mustnotlst
        anslst,entrylst, mustlst, mustnotlst=[],[],[],[]
        #匯入標題
        self.welcome = ctk.CTkLabel(
            self.master,
            text="考題成績試算:",
            bg_color='#FFEEDD',
            font=('微軟正黑體',40),
            text_color='#000000',
            fg_color="#FFEEDD",
            width=240
        )
        self.welcome.grid(pady=20,row=0,column=0,columnspan=20,sticky='n')
        self.frame_1=ctk.CTkFrame(
            self.master,
            fg_color="#FFDCB9",
            bg_color="#FFEEDD",
            height=200,width=1570,
            )
        self.frame_1.place(relx=0.5,rely=0.3, anchor=tk.CENTER)
        self.frame_2=ctk.CTkFrame(
            self.master,
            fg_color="#FFDCB9",
            bg_color="#FFEEDD",
            height=210,width=1570,
            )
        self.frame_2.place(relx=0.5,rely=0.58, anchor=tk.CENTER)
        self.frame_table=ctk.CTkFrame(
            self.master,
            fg_color="#000000",
            bg_color="#000000",
            height=130,width=1050,
            )
        self.frame_table.grid(row=13,column=1,columnspan=9,rowspan=3,pady=30,)
        #標籤_計算細胞量
        self.label_N = ctk.CTkLabel(
            self.master,
            width=120,height=50,
            bg_color="#FFEEDD",
            fg_color="#FFEEDD",
            text="計算細胞量:",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_N.grid(row=1,column=0,pady=5,sticky='e',)
        #val_計算細胞量
        self.input_N = ctk.CTkComboBox(
            self.master,
            width=120,height=50,
            bg_color="#FFEEDD",
            fg_color="#C4E1E1",
            button_color="#FF9224",
            values=["","100", "200","500"],
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.input_N.grid(row=1,column=1,sticky='e')
        self.label_N_amount = ctk.CTkLabel(
            self.master,
            width=120,height=50,
            bg_color="#FFEEDD",
            fg_color="#FFEEDD",
            text="/片",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_N_amount.grid(row=1,column=2,pady=10,sticky='w',)
        #標籤_必須要打到的標題
        self.label_must = ctk.CTkLabel(
            self.master,
            width=200,height=50,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="必須要打到:",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_must.grid(row=2,column=0,padx=20,sticky='ew',)
        #標籤_不能打到的標題
        self.label_mustnot = ctk.CTkLabel(
            self.master,
            width=200,height=50,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="不可打到:",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_mustnot.grid(row=3,column=0,padx=20,sticky='ew',)
        #標籤_項目名稱
        self.label_item = ctk.CTkLabel(
            self.master,
            width=200,height=30,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="細胞項目名稱:",
            font=('微軟正黑體',22),
            text_color="#000000"
        )
        self.label_item.grid(row=4,column=0,padx=20,sticky='ew',)
        #標籤_答案百分比
        self.label_ans = ctk.CTkLabel(
            self.master,
            width=200,height=30,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="正確百分比數(%)",
            font=('微軟正黑體',22),
            text_color="#000000"
        )
        self.label_ans.grid(row=5,column=0,padx=20,sticky='ew',)
        #標籤_考生百分比
        self.label_val = ctk.CTkLabel(
            self.master,
            width=200,height=30,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="測試百分比數(%)",
            font=('微軟正黑體',22),
            text_color="#000000"
        )
        self.label_val.grid(row=6,column=0,padx=20,sticky='ew',)
        #spacer
        self.label_spacer = ctk.CTkLabel(
            self.master,
            width=200,height=20,
            bg_color="#FFEEDD",
            fg_color="#FFEEDD",
            text="",
        )
        self.label_spacer.grid(row=7,column=0,pady=10,sticky='ew',)
        #標籤_必須要打到的標題_第二行
        self.label_must_2 = ctk.CTkLabel(
            self.master,
            width=200,height=50,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="必須要打到:",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_must_2.grid(row=8,column=0,padx=20,pady=5,sticky='s',)
        #標籤_不能打到的標題_第二行
        self.label_mustnot_2 = ctk.CTkLabel(
            self.master,
            width=200,height=50,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="不可打到:",
            font=('微軟正黑體',32),
            text_color="#000000"
        )
        self.label_mustnot_2.grid(row=9,column=0,padx=20,sticky='ew',)
        #標籤_項目名稱_第二行
        self.label_item_2 = ctk.CTkLabel(
            self.master,
            width=200,height=30,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="細胞項目名稱:",
            font=('微軟正黑體',22),
            text_color="#000000"
        )
        self.label_item_2.grid(row=10,column=0,padx=20,sticky='ew',)
        #標籤_答案百分比_第二行
        self.label_ans_2 = ctk.CTkLabel(
            self.master,
            width=200,height=30,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="正確百分比數(%)",
            font=('微軟正黑體',22),
            text_color="#000000"
        )
        self.label_ans_2.grid(row=11,column=0,padx=20,sticky='ew',)
        #標籤_答案百分比_第二行
        self.label_val_2 = ctk.CTkLabel(
            self.master,
            width=200,height=30,
            bg_color="#FFEEDD",
            fg_color="#FFDCB9",
            text="測試百分比數(%)",
            font=('微軟正黑體',22),
            text_color="#000000"
        )
        self.label_val_2.grid(row=12,column=0,padx=20,sticky='ew',)
        m=1
        for item in self.celllst:
            ##細胞標籤
            self.label_cell = ctk.CTkLabel(
                self.master,
                width=100,height=12,
                bg_color="#FFDCB9",
                fg_color="#FFDCB9",
                text=item,
                font=('微軟正黑體',14),
                text_color="#000000",
            )
            #必須打到chkbox
            self.cell_chk_must = ctk.CTkCheckBox(
                self.master,
                width=10,height=12,
                checkbox_width=40,checkbox_height=40,
                bg_color="#FFDCB9",
                text="",
            )
            ##不能打到chkbox
            self.cell_chk_mustnot = ctk.CTkCheckBox(
                self.master,
                width=10,height=12,
                checkbox_width=40,checkbox_height=40,
                bg_color="#FFDCB9",
                text="",
            )
            ##標準答案
            self.cell_ans = ctk.CTkEntry(
                self.master,
                width=100,height=20,
                bg_color='#FFDCB9',
                fg_color='#C4E1E1',
                justify='center',
                # text="0",
                font=('微軟正黑體',14),
                text_color="#000000"
                )
            self.cell_ans.insert(tk.END,0)
            ##學生答案
            self.cell_val = ctk.CTkEntry(
                self.master,
                width=100,height=20,
                bg_color='#FFDCB9',
                fg_color='#C4E1E1',
                justify='center',
                # text="0",
                font=('微軟正黑體',14),
                text_color="#000000"
                )
            self.cell_val.insert(tk.END,0)
            ##nRBC細胞與megakaryocyte細胞是不需計算百分比的，要予以排除    
            anslst.append(self.cell_ans)
            entrylst.append(self.cell_val)
            ##chkbox歸lst管理
            mustlst.append(self.cell_chk_must)
            mustnotlst.append(self.cell_chk_mustnot)
            #放置框架中
            if m<10:
                self.label_cell.grid(row=4,column=m,padx=6,sticky='n',)
                self.cell_chk_must.grid(row=2,column=m,sticky='ns',)
                self.cell_chk_mustnot.grid(row=3,column=m,sticky='ns',)
                self.cell_ans.grid(row=5,column=m,sticky='ns',)
                self.cell_val.grid(row=6,column=m,sticky='ns',)
                m+=1
            else:
                self.label_cell.grid(row=10,column=m-9,padx=6,sticky='nsew',)
                self.cell_chk_must.grid(row=8,column=m-9,sticky='ns',)
                self.cell_chk_mustnot.grid(row=9,column=m-9,sticky='ns',)
                self.cell_ans.grid(row=11,column=m-9,sticky='ns',)
                self.cell_val.grid(row=12,column=m-9,sticky='ns',)
                m+=1
        #計算按鈕
        self.btn_OK = ctk.CTkButton(
            self.master,
            text="計算",
            anchor=tk.CENTER,
            command=self.calculate_cell,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=80,height=150,
        )
        self.btn_OK.grid(row=13,column=0,rowspan=2,padx=50,pady=30,sticky='ew')
        #清除資料
        self.btn_clear = ctk.CTkButton(
            self.master,
            text="清除",
            anchor=tk.CENTER,
            command=self.clear,
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,
        )
        self.btn_clear.grid(row=13,column=10,pady=40,sticky='s')
        #返回主介面
        self.btn_back = ctk.CTkButton(
            self.master,
            text="返回",
            anchor=tk.CENTER,
            command=lambda: self.back(oldmaster),
            fg_color="#FF9224",
            bg_color='#FFEEDD',
            font=('微軟正黑體',22,'bold'),
            text_color="#000000",
            width=120,
        )
        self.btn_back.grid(row=14,column=10,pady=40,sticky='s')
        ##結果表格設計
        self.result = Sheet(
            self.frame_table,
            # data = [[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(50)] for r in range(500)],
            column_width=180,
            width=1200,height=180,

        )
        self.result.headers(self.celllst)
        self.result.enable_bindings()
        self.result.grid(row=0,column=0,sticky='nsew')
#####↓↓↓↓↓流程: 檢查UI上填寫的資料，確認沒錯誤後，餵進去計算成績↓↓↓↓↓#####
    #檢查輸入資料_是否皆為數字且下面百分比部分加起來是100
    def chk_val(self):
        #檢查細胞數量是否為100/200/500
        amount = self.input_N.get()
        
        if int(amount) == 100 or int(amount) == 200 or int(amount) == 500 :
            pass
        else:
            tk.messagebox.showerror(title='檢驗醫學部(科)', message='請在計算細胞量中輸入數字!\n請輸入100/200/500!')
            self.input_N.delete(0,tk.END)
            return False            
        for value in entrylst:
            val = value.get()
            try:
                float(val)
            except ValueError:
                tk.messagebox.showerror(title='檢驗醫學部(科)', message='請在細胞百分比中輸入數字!')
                return False
        #檢查_是否百分比的部分加起來為100
        tal = 0
        tal_lst=[]
        for value in anslst:
            val = value.get()
            try:
                float(val)
            except ValueError:
                tk.messagebox.showerror(title='檢驗醫學部(科)', message='''請在細胞百分比中輸入數字!
如果沒有此細胞請填0''')
                return False
            val = int(val)
            tal_lst.append(val)
        ##nRBC細胞與megakaryocyte細胞是不需計算百分比的，要予以排除
        tal_lst.pop(2)
        tal_lst.pop(2)
        tal = sum(tal_lst)
        if tal != 100:
            tk.messagebox.showerror(title='檢驗醫學部(科)', message='解答百分比加總不為100，請重新輸入!')
        # print(tal)
            return False
    #檢查輸入資料_chkbox上下不可同時勾取
    def chk_chkbox(self):
        self.mustlst=[]
        self.mustnotlst=[]
        for boolin in mustlst:
            aa = boolin.get()
            self.mustlst.append(aa)
        for boolin in mustnotlst:
            bb = boolin.get()
            self.mustnotlst.append(bb)
        # print(l_1,self.mustnotlst)
        for i in range(0,len(self.mustlst)):
            if self.mustlst[i] + self.mustnotlst[i] >=2:
                tk.messagebox.showerror(title='檢驗醫學部(科)', message='勾選重複!請勿重複勾選!')
                return False
            else:
                pass
        return
#####↑↑↑↑↑流程: 檢查UI上填寫的資料，確認沒錯誤後，餵進去計算成績↑↑↑↑↑#####    
#####↓↓↓↓↓###############計算成績: 餵進去計算成績###############↓↓↓↓↓#####
    def calculate_cell(self):
        if self.chk_val()==False or self.chk_chkbox()==False:
            return
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
        print(mtl,mnl)
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
        ncount = self.input_N.get()
        #設定好之後去json裡面調出相對應的格式db
        colname_1 = "%s_lower"%(str(ncount))
        colname_2 = "%s_upper"%(str(ncount))
        db = self.rangechart[[colname_1,colname_2]]
        #18個細胞的matrix
        chkrange = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        #3個chk
        ablym_chk,must_chk,mustnot_chk = "","",""
        anscal, studentcal = [], []
        for value in anslst:
            anscal.append(float(value.get()))
        for value in entrylst:
            studentcal.append(float(value.get()))
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
        print("must:"+ must_chk)
        print("mustnot:"+ mustnot_chk)
        print("ab:" + ablym_chk)
        #如果兩個判定其中一個False，return final_score=0，其餘照舊
    #Rule II: 確定19個細胞都在Range中
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
                chkrange[c] += 1
            else:
                pass
        #要檢查是不是只有plasma cell 跟abnormal lympho不再range內，如果是的話，綜合加起來一起評估
        if plasmaabnormal==0 or plasmaabnormal==1:
            if sum(chkrange[1::]) == 16:
                final_score = 2
            else:
                final_score = 1
        elif sum(chkrange) == 18:
            final_score = 2
        else:
            final_score = 1
        if must_chk == "False" or mustnot_chk == "False" or ablym_chk == "False":
            final_score=0
        else:
            pass
        print(chkrange)
        print(final_score)
        #Rule III: 看看有沒有blast & abn-Lym 重複出現chk
        if studentcal[1] > 0 and studentcal[5] > 0:
            final_score=0
        else:
            pass
        tk.messagebox.showinfo(title='檢驗醫學部(科)', message="""已完成計算!
必須打到細胞檢查: %s
不可打到細胞檢查: %s
plasmacell+abn-Lym(空白表示不需檢查):%s
最後總分:%d"""%(must_chk,mustnot_chk,ablym_chk,final_score))
        #加入tksheets
        self.result.insert_row(values=tuple(chkrange))
#####↑↑↑↑↑###############計算成績: 餵進去計算成績###############↑↑↑↑↑#####
##清除按鈕
    def clear(self):
        ##清除細胞總數選擇
        self.input_N.set("")
        ##清除chkbox中選中的
        for item in mustlst:
            item.deselect()
        for item in mustnotlst:
            item.deselect()
        
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
    SC = SCORE_TEST(root)
    # root['bg']='#FFEEDD'
    # B.gui_arrang()
    # 主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  