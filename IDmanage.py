#authorised by Henry Tsai
import sys
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox
import customtkinter as ctk
import json,pyodbc
from setuptools import Command
from verifyAccount import addaccount_sql,delaccount,editaccount,changepw_sql
from PIL import Image


connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=bloodtest;UID=sa;PWD=1234"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.OperationalError:
    connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)
##擷取SQL ID table中有的院區
with coxn.cursor() as cursor:
    unique_hos = "SELECT DISTINCT [院區] FROM [bloodtest].[dbo].[id];"
    cursor.execute(unique_hos)
    hos = cursor.fetchall()
hos = [e[0] for e in hos]

class ID:
    
    def __init__(self,master,oldmaster=None):  
        # 建立登入後視窗  
        self.master = master
        super().__init__()
        self.master.protocol("WM_DELETE_WINDOW",lambda: self.exit(oldmaster))
        ctk.set_default_color_theme("blue")  
        # 給主視窗設定標題內容  
        self.master.title("帳號管理")  
        self.master.geometry('800x650')
        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_columnconfigure(1, weight=3)
        self.master.config(background='#FFEEDD') #設定背景色
        ##框架設置
        self.labelframe_1 = ctk.CTkFrame(self.master, corner_radius=0,fg_color="#FFDCB9",bg_color="#FFEEDD")
        self.labelframe_1.grid(row=0, column=0,rowspan=2, sticky="nsew")
        self.labelframe_1.grid_rowconfigure(4, weight=1)
        self.labelframe_2 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        self.labelframe_2.grid(row=0, column=1)
        self.labelframe_3 = ctk.CTkFrame(self.master,fg_color="#FFEEDD",bg_color="#FFEEDD")
        # self.labelframe_2.grid_columnconfigure(1, weight=1)
##左手邊功能區
        self.people = ctk.CTkImage(Image.open("assets\person.png"),size=(160,120))
        self.label_1 = ctk.CTkLabel(
                    self.labelframe_1, 
                    image=self.people,
                    compound="top",
                    text = "帳號管理", 
                    fg_color='#FFDCB9',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_1.grid(row=0,column=0,padx=20, pady=20)
        #新增帳號按鈕
        self.btn_1=ctk.CTkButton(
            self.labelframe_1,
            command = self.switch1, 
            text = "新增帳號", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_1.grid(row=1,column=0,pady=10,sticky="ew")
        #修改/刪除按鈕
        self.btn_2=ctk.CTkButton(
            self.labelframe_1,
            command = self.changeedit, 
            text = "修改/刪除帳號", 
            fg_color='#FFDCB9',
            hover_color='#FFCC99',
            width=100,height=40,
            corner_radius=8,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_2.grid(row=2,column=0,pady=10,sticky="ew")
        #離開按鈕
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
        self.btn_3.grid(row=3,column=0,pady=10,sticky="ew")
        self.cc = ctk.CTkLabel(
            self.master, 
            fg_color="#FFEEDD",
            text='@Design by Henry Tsai',
            text_color="#E0E0E0",
            font=("Calibri",12),
            width=120)
        self.cc.grid(row=1,column=1,sticky='se') 
##新增帳號frame2
        self.label_2 = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "新增帳號", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_2.grid(row=0,column=0,columnspan=2,sticky='nsew')
        self.label_branch = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "院區:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_branch.grid(row=1,column=0,pady=10,sticky='e')
        self.label_ID = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "帳號:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_ID.grid(row=2,column=0,pady=10,sticky='e')
        self.label_pw = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "密碼:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_pw.grid(row=3,column=0,pady=10,sticky='e')
        self.label_pw2 = ctk.CTkLabel(
                    self.labelframe_2, 
                    text = "再次輸入密碼:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_pw2.grid(row=4,column=0,pady=10,sticky='e')
        #(f2)院區滾動欄位
        self.input_branch = ctk.CTkComboBox(
                    self.labelframe_2, 
                    values=hos,
                    # fg_color='#000000',
                    button_color='#FF9900',
                    width=120,height=40,
                    # font=('微軟正黑體',16,)
                    )
        self.input_branch.grid(row=1,column=1,pady=10,sticky='w')
        #(f2)帳號輸入欄位
        self.input_ID = ctk.CTkEntry(
                    self.labelframe_2, 
                    width=120,height=40
                    )
        self.input_ID.grid(row=2,column=1,pady=10,sticky='w')
        #(f2)密碼輸入欄位
        self.input_pw = ctk.CTkEntry(
                    self.labelframe_2, 
                    width=120,height=40,
                    show='*'
                    )
        self.input_pw.grid(row=3,column=1,pady=10,sticky='w')
        self.input_pw2 = ctk.CTkEntry(
                    self.labelframe_2, 
                    width=120,height=40,
                    show='*'
                    )
        self.input_pw2.grid(row=4,column=1,pady=10,sticky='w')
        #(f2)確認/清除按鈕
        self.f2_okbtn = ctk.CTkButton(
            self.labelframe_2,
            text = "確定",
            command=self.addaccount,
            height=30,
            fg_color='#FF9900',
            text_color='#000000')
        self.f2_okbtn.grid(row=5,column=0,padx=40,pady=15)
        self.f2_clrbtn = ctk.CTkButton(
            self.labelframe_2,
            text = "清除",
            command=self.clear,
            height=30,
            fg_color='#FF9900',
            text_color='#000000')
        self.f2_clrbtn.grid(row=5,column=1,padx=40,pady=15)
##(f2)新增帳號介面功能區
    #(f2)清除欄位功能
    def clear(self):
        self.input_ID.delete(0,tk.END)
        self.input_pw.delete(0,tk.END)
        self.input_pw2.delete(0,tk.END)
    #(f2)新增帳號的功能鍵
    def addaccount(self):
        branch = self.input_branch.get()
        account = self.input_ID.get()
        pw1 = self.input_pw.get()
        pw2 = self.input_pw2.get()
        ##檢查密碼是否重複
        if pw1 !=pw2:
            tk.messagebox.showerror(title='土城醫院檢驗科', message='密碼兩次不一致，請重新輸入!')
            self.input_pw.delete(0,tk.END)
            self.input_pw2.delete(0,tk.END)
            return None
        else:
            result = addaccount_sql(branch,account,pw1)
        if result =="duplicate":
            tk.messagebox.showerror(title='土城醫院檢驗科', message='帳號重複!請重新輸入!')
            self.clear()
        elif result =="success":
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='帳號建立成功!')
            self.clear()
##(f3)修改/刪除帳號介面
    def changeedit(self):
        ##(f3)listbox轉換
        def listbox_event(event):
            idx = self.id_listbox.curselection()
            slcid = self.id_listbox.get(idx)
            self.input_ID2.configure(state='normal')
            self.input_ID2.delete(0,tk.END)
            self.input_pwe2.delete(0,tk.END)
            self.input_ID2.insert(tk.END,slcid)
            self.input_pwe2.insert(tk.END,'*******')
            self.input_ID2.configure(state='disabled')
            self.chgpw_btn.configure(state='disabled')
        ##把frame3放到frame2上
        self.labelframe_3.grid(row=0, column=1,sticky='nsew')
        self.labelframe_3.columnconfigure(0,weight=1)
        self.label_3 = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "修改/刪除帳號", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',36),
                    text_color="#000000",
                    width=120,height=120
                    )
        self.label_3.grid(row=0,column=0,columnspan=5,sticky='ew')
        self.f3_label_branch = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "院區:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.f3_label_branch.grid(row=1,column=1,columnspan=2,)
        self.f3_input_branch = ctk.CTkComboBox(
                    self.labelframe_3, 
                    values=hos,
                    command=self.updatelist,
                    button_color='#FF9900',
                    width=120,height=40,
                    # font=('微軟正黑體',16,)
                    )
        self.f3_input_branch.grid(row=1,column=3,sticky='w')
        #spacer
        self.spacer = ctk.CTkLabel(
            self.labelframe_3,
            text=" ",
            width=10
        )
        self.spacer.grid(row=2,column=0,rowspan=3,)
        ##(f3)左手邊tk.listbox & 卷軸
        self.scrollbar = tk.Scrollbar(self.labelframe_3)
        self.scrollbar.grid(row=2,column=2,rowspan=3,sticky='nsew')
        self.id_listbox = tk.Listbox(
            self.labelframe_3,
            yscrollcommand=self.scrollbar.set,
            height=10,width=8,
            font=('微軟正黑體',20)
        )
        # self.updatelist()
        self.id_listbox.grid(row=2,column=1,rowspan=3,sticky='nsew',)
        self.scrollbar.config(command=self.id_listbox.yview)
        #(f3)listbox bind
        self.id_listbox.bind("<<ListboxSelect>>", listbox_event)        
        #(f3)右手邊點左邊listbox後會跳出帳戶資訊
        self.label_ID2 = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "帳號:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_ID2.grid(row=2,column=3,pady=10,sticky='se')
        self.label_pw2 = ctk.CTkLabel(
                    self.labelframe_3, 
                    text = "密碼:", 
                    fg_color='#FFEEDD',
                    font=('微軟正黑體',20),
                    text_color="#000000",
                    width=140,height=40
                    )
        self.label_pw2.grid(row=3,column=3,pady=10,sticky='ne')
        self.input_ID2 = ctk.CTkEntry(
                    self.labelframe_3, 
                    width=120,height=40,
                    state='disabled'
                    )
        self.input_ID2.grid(row=2,column=4,pady=10,sticky='sw')
        #(f3)double clicked ID bind
        self.input_ID2 .bind("<Double-Button-1>", self.clicklabelac)
        self.input_pwe2 = ctk.CTkEntry(
                    self.labelframe_3, 
                    width=120,height=40
                    )
        self.input_pwe2.grid(row=3,column=4,pady=10,sticky='nw')
        #(f3)double clicked password bind
        self.input_pwe2.bind("<Double-Button-1>", self.clicklabelpw)
        #(f3)下面修改按鍵
        self.chgpw_btn = ctk.CTkButton(
            self.labelframe_3,
            text = "修改帳號",
            command=self.editac,
            height=30,
            fg_color='#FF9900',
            text_color='#000000',
            state='disabled')
        self.chgpw_btn.grid(row=4,column=3,padx=40,sticky='n')
        #(f3)刪除帳號按鍵
        self.del_btn = ctk.CTkButton(
            self.labelframe_3,
            text = "刪除",
            command=self.delac,
            height=30,
            fg_color='#FF9900',
            text_color='#000000')
        self.del_btn.grid(row=4,column=4,padx=40,sticky='n')
##(f3)修改帳號介面功能區    
    #(f3)更新listbox(MSSQL裡面有的帳號)
    def updatelist(self,event):
        branch = self.f3_input_branch.get()
        with coxn.cursor() as cursor:
            query = "SELECT ac,pw  FROM [bloodtest].[dbo].[id] WHERE [院區]='%s';"%(branch)
            cursor.execute(query)
            acpw = dict(cursor.fetchall())
        # jsonfile = open('in.json','rb')
        # a = json.load(jsonfile)
        # ml = a['member']
        idlist = [i for i in acpw.keys()]
        # for i in ml:
        #     x = i['ID']
        #     idlist.append(x)
        self.id_listbox.delete(0,tk.END)
        for i in idlist:
            self.id_listbox.insert(tk.END, i)
    #清除list中資料
    def clearlist(self):
        self.id_listbox.delete(0,tk.END)
        return
    #(f3)點兩下帳號可以修改帳號
    def clicklabelac(self,event):
        self.toedit = self.input_ID2.get()
        if self.toedit !="":
            self.input_ID2.configure(state='normal')
            tk.messagebox.showinfo(
                title='土城醫院檢驗科', 
                message="""請修改帳號框中帳號""")
            self.chgpw_btn.configure(state='normal')
        else:
            tk.messagebox.showerror(title='土城醫院檢驗科', message="尚未選擇帳號!請選擇後再逕行修改!")
            return
    #(f3)修改帳號功能鍵
    def editac(self):
        editid = self.input_ID2.get()
        if editid !="":
            if tk.messagebox.askyesno(title='土城醫院檢驗科', message='確認將%s修改為%s?'%(self.toedit,editid)):
                state = editaccount(self.toedit,editid)
            else:
                return
            if state == "success":
                tk.messagebox.showinfo(title='土城醫院檢驗科', message='修改成功!')
                self.input_ID2.configure(state='normal')
                self.input_ID2.delete(0,tk.END)
                self.input_ID2.configure(state='disable')
                self.input_pwe2.delete(0,tk.END)
                self.clearlist()
                self.chgpw_btn.configure(state='disabled')
            else:
                tk.messagebox.showerror(title='土城醫院檢驗科', message='發生未知錯誤，請聯絡管理員')
        else:
            tk.messagebox.showerror(title='土城醫院檢驗科', message='尚未選擇帳號!請選擇後再逕行修改!')
            return
    #(f3)點兩下密碼可以修改密碼
    def clicklabelpw(self,event):
        #toplevel裡的確認鍵
        def ok():
            account = self.input_ID2.get()
            newpw = self.input_newpw.get()
            newpw2 = self.input_newpw2.get()
            if newpw == newpw2:
                changeResult = changepw_sql(account=account,newpassword=newpw)
                print(changeResult)
                if changeResult == "success":
                    tk.messagebox.showinfo(title='土城醫院檢驗科', message='修改成功!')
                    self.newWindow.destroy() 
            else:
                tk.messagebox.showinfo(title='土城醫院檢驗科', message='新密碼不一致，請重新輸入!')
                self.input_newpw.delete(0,tk.END)
                self.input_newpw2.delete(0,tk.END)
        #toplevel裡的取消鍵
        def cancel():
            self.newWindow.destroy() 
        #創建一個新視窗newWindow
        account = self.input_ID2.get()
        self.newWindow = ctk.CTkToplevel()
        self.newWindow.title("修改密碼")
        self.label_newpw = ctk.CTkLabel(self.newWindow,text='請輸入新密碼: ',font=('微軟正黑體',18),height=30, width=120)
        self.label_newpw.grid(row=0,column=0,padx=10,pady=15)
        self.input_newpw = ctk.CTkEntry(self.newWindow,height=30, width=120,show='*')
        self.input_newpw.grid(row=0,column=1,padx=10,pady=15)
        self.label_newpw2 = ctk.CTkLabel(self.newWindow,text='請再次輸入新密碼: ',font=('微軟正黑體',18),height=30, width=120)
        self.label_newpw2.grid(row=1,column=0,padx=10,pady=15)
        self.input_newpw2 = ctk.CTkEntry(self.newWindow,height=30, width=120,show='*')
        self.input_newpw2.grid(row=1,column=1,padx=10,pady=15)
        self.chgpw_btn = ctk.CTkButton(self.newWindow,text = "確定",command=ok,height=30, width=120)
        self.chgpw_btn.grid(row=2,column=0,padx=40,pady=15)
        self.cancel_btn = ctk.CTkButton(self.newWindow,text = "取消",command=cancel,height=30, width=120)
        self.cancel_btn.grid(row=2,column=1,padx=40,pady=15)
    #(f3)刪除帳號功能鍵
    def delac(self):
        id = self.input_ID2.get()
        if tk.messagebox.askyesno(title='土城醫院檢驗科', message='確認刪除?'):
            state = delaccount(id)
        else:
            return
        if state == "success":
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='刪除成功!')
            self.input_ID2.configure(state='normal')
            self.input_ID2.delete(0,tk.END)
            self.input_ID2.configure(state='disable')
            self.input_pwe2.delete(0,tk.END)
            self.clearlist()
        else:
            tk.messagebox.showerror(title='土城醫院檢驗科', message='發生未知錯誤，請聯絡管理員')
    

    #(switch)從修改刪除帳號切回新增帳號
    def switch1(self):
        self.labelframe_3.grid_forget()
        self.labelframe_1.tkraise()
    #(f1)離開介面
    def exit(self,oldmaster):
        try:
            oldmaster.deiconify()
            self.master.withdraw()
        except AttributeError:
            self.master.destroy()
            return

def main():  
    root = ctk.CTk()
    B = ID(root)
    #  主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  