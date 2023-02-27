#authorised by Henry Tsai
from email import message
from operator import index
import sys
import tkinter as tk
import pandas as pd
from tkinter import IntVar, simpledialog,messagebox,DoubleVar
import customtkinter as ctk
import json
from setuptools import Command
from verifyAccount import changepw2,addaccount,delaccount,editaccount
from PIL import Image


class Modify:
    
    def __init__(self,master,oldmaster=None):  
        # 建立登入後視窗  
        self.master = master
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.exit(oldmaster))
        super().__init__()
        ctk.set_default_color_theme("blue")
        jsonfile = open('testdata\data.json','rb')
        rawdata = json.load(jsonfile)
        self.rawdata = pd.DataFrame(rawdata["blood"])
        self.testyear = self.rawdata['year'].unique().tolist()
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
        self.frame_2s.grid(row=1, column=2,columnspan=3,rowspan=4,)
        self.labelframe_3 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
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
            "Plt":DoubleVar()
        }
        self.ansinfo={
            "plasma cell": IntVar(),
            "abnormal lympho": IntVar(),
            "megakery cell": IntVar(),
            "nRBC": IntVar(),
            "abnormal monocyte": IntVar(),
            "blast": IntVar(),
            "metamyelocyte": IntVar(),
            "eosinopil": IntVar(),
            "plasma cytoid": IntVar(),
            "promonocyte": IntVar(),
            "promyelocyte": IntVar(),
            "band neutropil": IntVar(),
            "basopil": IntVar(),
            "atypical lymphocyte": IntVar(),
            "hypersegmented neutrophil": IntVar(),
            "myelocyte": IntVar(),
            "segmented neutrophil": IntVar(),
            "lymphocyte": IntVar(),
            "monocyte": IntVar()
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
            font=('微軟正黑體',20))
        self.test_listbox.grid(row=2,column=0,rowspan=2,sticky='nsew',)
        self.scrollbar.grid(row=2,column=1,rowspan=2,sticky='nsew')
        self.scrollbar.config(command=self.test_listbox.yview)
        #(f2)點選listbox之後連結事件self.listbox_event
        self.test_listbox.bind("<<ListboxSelect>>", self.listbox_event)        
        ##(f2s)旁邊編輯窗格
        testinfolist = [key for key in self.testinfo.keys()]
        ansinfolist = [key for key in self.ansinfo.keys()]
        global entrylst
        entrylst = []
        #loop 0,2,4放入項目(WBC/RBC/Hct...)
        for i in range(0,6,2):
            for j in range(0,10):
                #第一欄
                if i == 0:
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
                        self.modifytable.insert(tk.END,testinfolist[j])
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
                        self.modifytable.grid(row=j,column=i)
                        self.modifytable.configure(state="disabled")
                #第三欄
                elif i==2:
                    self.modifytable = ctk.CTkEntry(
                        self.frame_2s,
                        width=100,height=20,
                        bg_color='#FFEEDD',
                        fg_color='#FFCC99',
                        # text="",
                        font=('微軟正黑體',12),
                        text_color="#000000"
                        )
                    self.modifytable.insert(tk.END,ansinfolist[j])
                    self.modifytable.grid(row=j,column=i)
                    self.modifytable.configure(state="disabled")
                #第五欄
                elif i==4:
                    self.modifytable = ctk.CTkEntry(
                        self.frame_2s,
                        width=100,height=20,
                        bg_color='#FFEEDD',
                        fg_color='#FFCC99',
                        # text="",
                        font=('微軟正黑體',12),
                        text_color="#000000"
                        )
                    #從ansinfolist[11]開始跳
                    try:
                        self.modifytable.insert(tk.END,ansinfolist[j+10])
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
                        self.modifytable.grid(row=j,column=i)
                        self.modifytable.configure(state="disabled")
        #loop 1,3,5放入變數(self.testinfo.values())
        for i in range(1,7,2):
            m = 0
            if i == 1:
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
                    self.modifytable.grid(row = m, column = i,sticky='nsew')
                    self.modifytable.configure(state="disabled")
                    entrylst.append(self.modifytable)
                    m += 1
            elif i == 3:
                for values in self.ansinfo.values():
                    self.modifytable = ctk.CTkEntry(
                            self.frame_2s,
                            width=50,height=20,
                            bg_color='#FFEEDD',
                            fg_color='#FFFFFF',
                            textvariable= values,
                            font=('微軟正黑體',14),
                            text_color="#000000"
                            )
                    if m<10:
                        self.modifytable.grid(row = m, column = i,sticky='nsew')
                        self.modifytable.configure(state="disabled")
                        entrylst.append(self.modifytable)
                        m += 1
                    else:
                        self.modifytable.grid(row = m-10, column = i+2,sticky='nsew')
                        self.modifytable.configure(state="disabled")
                        entrylst.append(self.modifytable)
                        m += 1
        #補空格
        for i in range(1,9,4):
            MB = ctk.CTkEntry(
                self.frame_2s,
                width=50,height=20,
                bg_color='#FFEEDD',
                fg_color='#FFFFFF',
                state='disabled',
                )
            MB.grid(row = 9, column = i,sticky='nsew')
            MB.configure(state="disabled")
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
        # self.labelframe_2.grid_forget()
        self.slcid = tk.StringVar()
        ##把frame3放到frame1上
        self.labelframe_3.grid(row=0, column=1,sticky='nsew')
        self.labelframe_3.columnconfigure(0,weight=1)
        self.label_3 = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "參數設定", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_3.grid(row=0,column=0,columnspan=5,sticky='ew')
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
            else:
                return
    #(f2)確定
    def editdata(self):
        if tk.messagebox.askyesno(title='土城醫院檢驗科', message='改完了?'):
            i = self.rawdata.index
            #年份跟ID做為filter
            filter1 = (self.rawdata['year'] == self.focusyear)
            filter2 = (self.rawdata['ID'] == self.focustest)
            #把原本的資料調進來
            index = filter1 & filter2
            result = i[index]
            result.tolist()
            result=result[0]
            #建立一個新的dict，裝修改過後的資料
            dict_rawdata={}
            for key,value in self.testinfo.items():
                A = value.get()
                dict_rawdata[key]=A
            self.rawdata.at[result,'rawdata'] = dict_rawdata
            add = self.rawdata.to_json(orient="records")
            jsonfile = open('testdata\data.json','rb')
            a = json.load(jsonfile)
            a["blood"] = json.loads(add)
            with open('testdata/data.json','w') as r:
                json.dump(a,r)
                r.close()
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
            for value in self.ansinfo.values():
                value.set(0)
        else:
            return
##(f2)事件連結
    #選擇年份之後，更新listbox(JSON裡面有的考題)
    def updatelist(self,event):
        #取得combobox選擇的年份
        year = self.input_year.get()
        #建立一個filter把年份對的篩選出來
        fliter = (self.rawdata['year'] == year)
        self.testlist = self.rawdata[fliter]["ID"]
        #放入資料之前，先把listbox清空
        self.test_listbox.delete(0,tk.END)
        for i in self.testlist:
            self.test_listbox.insert(tk.END,i)
    #listbox選擇後，更新self.testinfo內的變數
    def listbox_event(self,event):
        #取得combobox選擇的年份
        year = self.input_year.get()
        #取得listbox選擇的考題
        idx = self.test_listbox.curselection()
        try:
            slcid = self.test_listbox.get(idx)
        except tk.TclError:
            return
        #建立兩個filter篩選出考題
        fliter1 = (self.rawdata['year'] == year)
        fliter2 = (self.rawdata['ID'] == slcid)
        datalist = self.rawdata[fliter1 & fliter2]
        #兩個dict分別放CBC DATA & Ans
        datadict1 = datalist.iloc[0]["rawdata"]
        datadict2 = datalist.iloc[0]["Ans"]
        for key,value in datadict1.items():
            self.testinfo[key].set(value)
        for key,value in datadict2.items():
            self.ansinfo[key].set(value)
##(switch)從血液參數設定切回考題設定
    def switch1(self):
        self.labelframe_3.grid_forget()
        self.labelframe_1.tkraise()
        # self.frame_2s.tkraise()
    
def main():  
    root = ctk.CTk()
    M = Modify(root)
    #  主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  