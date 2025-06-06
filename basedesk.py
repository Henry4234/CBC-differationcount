#authorised by Henry Tsai
import sys
from threading import activeCount
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import messagebox
import customtkinter as ctk
import subprocess

from setuptools import Command
from verifyAccount import changepw_sql
# from counter import getaccount
# global account
# account = sys.argv[1]
def get_accountpermission(acount,govid,hoscode):
    global Baccount,Govid,Hos_code
    Baccount = str(acount)
    Govid = str(govid)
    Hos_code = str(hoscode)

    return None

class Basedesk:
    
    def __init__(self,master,oldmaster=None):  
        # 建立登入後視窗  
        self.master = master
        super().__init__()
        ctk.set_default_color_theme("dark-blue")  
        # 給主視窗設定標題內容  
        self.master.title("形態學考核及教育程式-v5.3")  
        self.master.geometry('800x650')
        self.master.config(background='#FFEEDD') #設定背景色
        global account_1
        s = ttk.Style()
        s.configure('Red.TLabelframe.Label', font=('微軟正黑體', 12))
        s.configure('Red.TLabelframe.Label', foreground ='#FFFFFF')
        s.configure('Red.TLabelframe.Label', background='#FFEEDD')
        s.configure("1.basedesk",background = "#FFEEDD")
        self.hellow_label = ctk.CTkLabel(
            self.master, 
            # text = "歡迎回來",
            text = "歡迎回來 %s"%(Baccount),
            fg_color='#FFEEDD',
            font=('微軟正黑體',40),
            text_color="#000000",
            width=240
            )
        self.hellow_label.pack()
        self.label_1 = ctk.CTkLabel(
            self.master, 
            text = "請選擇需要使用的功能:", 
            fg_color='#FFEEDD',
            font=('微軟正黑體',36),
            text_color="#000000",
            width=300
            )
        self.label_1.pack()
        ##框架設置
        self.labelframe_1 = ctk.CTkFrame(
            self.master,
            # text='1. 考核介面',
            fg_color="#FFDCB9",
            bg_color="#FFEEDD"
            )
        self.labelframe_1.pack()
        self.labelframe_2 = ctk.CTkFrame(
            self.master,
            # text='2. 成績查詢',
            fg_color="#FFDCB9",
            bg_color="#FFEEDD"
            )
        self.labelframe_2.pack()
        self.labelframe_3 = ctk.CTkFrame(
            self.master,
            # text='3. 練習介面',
            fg_color="#95C8EF",
            bg_color="#FFEEDD"
            )
        self.labelframe_3.pack()
        self.labelframe_link = ctk.CTkFrame(
            self.master,
            # text='3. 數位學習網',
            fg_color="#BDF5CB",
            bg_color="#FFEEDD"
            )
        self.labelframe_link.pack()
        ##各類功能選單
        self.button_blood=ctk.CTkButton(
            self.labelframe_1, 
            command = self.bloodcounter, 
            text = "血液考核",
            fg_color='#FF9900', 
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000",
            )
        self.button_blood.grid(row = 0,column = 0,padx=10, pady=15)
        self.button_urin=ctk.CTkButton(
            self.labelframe_1, 
            command = self.urinesedimentcounter, 
            text = "尿沉渣考核",
            fg_color='#FF9900', 
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000",
            )
        self.button_urin.grid(row = 0,column = 1,padx=10, pady=15)
        self.button_bodyfluid=ctk.CTkButton(
            self.labelframe_1, 
            command = self.bodyfluidcounter, 
            text = "體液考核",
            fg_color='#FF9900', 
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000"
            )
        self.button_bodyfluid.grid(row = 0,column = 2,padx=10, pady=15)
        self.button_practise=ctk.CTkButton(
            self.labelframe_3, 
            command = self.practise, 
            text = "練習考核",
            fg_color='#A8DEF0', 
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000"
            )
        self.button_practise.grid(row = 0,column = 0,padx=10, pady=10)
        self.button_image_pracrice_1=ctk.CTkButton(
            self.labelframe_3, 
            # command = lambda:self.image_practise(1), 
            command= None,
            text = "圖庫練習_初階",
            fg_color='#A8DEF0', 
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000"
            )
        self.button_image_pracrice_1.grid(row = 0,column = 1,padx=10, pady=10)
        self.button_image_pracrice_2=ctk.CTkButton(
            self.labelframe_3, 
            command = lambda:self.image_practise(2), 
            text = "圖庫練習_中階",
            fg_color='#A8DEF0', 
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000"
            )
        self.button_image_pracrice_2.grid(row = 0,column = 2,padx=10, pady=10)
        self.button_search=ctk.CTkButton(
            self.labelframe_2, 
            command = self.scoresearch, 
            text = "成績查詢",
            fg_color='#FF9900',
            width=200,height=70,
            font=('微軟正黑體',26),
            text_color="#000000",
            )
        self.button_search.pack(padx=10, pady=25)
        self.button_link=ctk.CTkButton(
            self.labelframe_link,
            command = self.openlearning_link, 
            text = "醫檢數位學習網", 
            fg_color='#64E864',
            width=200,height=70,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_link.pack(padx=10, pady=25)
        self.button_logout=ctk.CTkButton(
            self.master, 
            command = lambda: self.logout_interface(oldmaster), 
            text = "登出", 
            fg_color='#FF6600',
            bg_color='#FFEEDD',
            width=180,height=40,
            font=('微軟正黑體',24,'bold'),
            text_color="#000000",
            state='disabled'
            )
        self.button_logout.pack(padx=10, pady=25)
        self.button_exit=ctk.CTkButton(
            self.master, 
            command = lambda: self.exit_interface(oldmaster), 
            text = "結束使用", 
            fg_color='#FF6600',
            bg_color='#FFEEDD',
            width=180,height=40,
            font=('微軟正黑體',24,'bold'),
            text_color="#000000"
            )
        self.button_exit.pack()
        self.button_changepw=ctk.CTkButton(
            self.master, 
            command = self.changepw, 
            text = "更改密碼", 
            fg_color='#FF6600',
            bg_color='#FFEEDD',
            width=180,height=40,
            font=('微軟正黑體',24,'bold'),
            text_color="#000000"
            )
        self.button_changepw.pack()
        self.cc = ctk.CTkLabel(
            self.master, 
            fg_color="#FFEEDD",
            text='@Design by Henry Tsai',
            text_color="#E0E0E0",
            font=("Calibri",12),
            width=150)
        self.cc.pack()
        self.versioninfo = ctk.CTkLabel(
            self.master, 
            fg_color="#FFEEDD",
            bg_color='#FFEEDD',
            text='@version -5.3',
            text_color="#000000",
            font=("Calibri",12),
            width=80)
        self.versioninfo.pack()
    # gui_arrang
        self.hellow_label.place(relx=0.5, rely=0.05, anchor=tk.CENTER)
        self.label_1.place(relx=0.5, rely=0.13, anchor=tk.CENTER)
        self.button_logout.place(relx=0.2,rely=0.9,anchor=tk.CENTER)
        self.button_exit.place(relx=0.5,rely=0.9,anchor=tk.CENTER)
        self.button_changepw.place(relx=0.8,rely=0.9,anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE) 
        self.versioninfo.place(relx=0, rely=1,anchor=tk.SW) 
        self.labelframe_1.place(relx=0.5,rely=0.34, anchor=tk.CENTER)
        self.labelframe_2.place(relx=0.775,rely=0.7, anchor=tk.CENTER)
        self.labelframe_3.place(relx=0.5,rely=0.5, anchor=tk.CENTER)
        self.labelframe_link.place(relx=0.225,rely=0.7, anchor=tk.CENTER)

    def changepw(self):
        def ok():
            oldpw = self.input_oldpw.get()
            newpw = self.input_newpw.get()
            newpw2 = self.input_newpw2.get()
            # print(oldpw,newpw,newpw2)
            if newpw == newpw2:
                changeResult = changepw_sql(account=Baccount,newpassword=newpw,oldpassword=oldpw)
                if changeResult =="wrongoldpassword":
                    tk.messagebox.showerror(title='檢驗醫學部(科)', message='舊密碼錯誤，請重新輸入!')
                    self.input_oldpw.delete(0,tk.END)
                    self.input_newpw.delete(0,tk.END)
                    self.input_newpw2.delete(0,tk.END)
                elif changeResult =="success":
                    tk.messagebox.showinfo(title='檢驗醫學部(科)', message='密碼更新成功!')
                    self.input_oldpw.delete(0,tk.END)
                    self.input_newpw.delete(0,tk.END)
                    self.input_newpw2.delete(0,tk.END)
                    self.newWindow.destroy()
            else:
                tk.messagebox.showinfo(title='檢驗醫學部(科)', message='新密碼不一致，請重新輸入!')
                self.input_newpw.delete(0,tk.END)
                self.input_newpw2.delete(0,tk.END)
        def cancel():
            self.newWindow.destroy()   
        self.newWindow = ctk.CTkToplevel()
        self.label_oldpw = ctk.CTkLabel(self.newWindow,text='請輸入舊密碼: ',font=('微軟正黑體',18),height=30, width=120)
        self.label_oldpw.grid(row=0,column=0,padx=10,pady=15)
        self.input_oldpw = ctk.CTkEntry(self.newWindow,height=30, width=120,show='*')
        self.input_oldpw.grid(row=0,column=1,padx=10,pady=15)
        self.label_newpw = ctk.CTkLabel(self.newWindow,text='請輸入新密碼: ',font=('微軟正黑體',18),height=30, width=120)
        self.label_newpw.grid(row=1,column=0,padx=10,pady=15)
        self.input_newpw = ctk.CTkEntry(self.newWindow,height=30, width=120,show='*')
        self.input_newpw.grid(row=1,column=1,padx=10,pady=15)
        self.label_newpw2 = ctk.CTkLabel(self.newWindow,text='請再次輸入新密碼: ',font=('微軟正黑體',18),height=30, width=120)
        self.label_newpw2.grid(row=2,column=0,padx=10,pady=15)
        self.input_newpw2 = ctk.CTkEntry(self.newWindow,height=30, width=120,show='*')
        self.input_newpw2.grid(row=2,column=1,padx=10,pady=15)
        self.chgpw_btn = ctk.CTkButton(self.newWindow,text = "確定",command=ok,height=30, width=120)
        self.chgpw_btn.grid(row=3,column=0,padx=40,pady=15)
        self.cancel_btn = ctk.CTkButton(self.newWindow,text = "取消",command=cancel,height=30, width=120)
        self.cancel_btn.grid(row=3,column=1,padx=40,pady=15)

    def bloodcounter(self):
        import counter
        from counter import Count
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        counter.getaccount(Baccount,Govid)
        C = Count(self.newWindow,self.master)
    def urinesedimentcounter(self):
        self.destroy()
        from counter_Urin import Count
        C = Count()
        C.matrix() 
        C.gui_arrang()
        C.infocreate()
        C.mainloop()
    def bodyfluidcounter(self):
        self.destroy()
        from counter_BodyFluid import Count
        C = Count()
        C.matrix() 
        C.gui_arrang()
        C.infocreate()
        C.mainloop()
    def practise(self):
        import counter_practise
        from counter_practise import PRACTISE
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        counter_practise.getaccount(Baccount,Govid,Hos_code)
        P = PRACTISE(self.newWindow,self.master)
    def image_practise(self,level):
        import practice_image
        from practice_image import IMAGEPRACTICE
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        practice_image.getaccount(Baccount)
        practice_image.getlevel(level)
        IP = IMAGEPRACTICE(self.newWindow,self.master)
    def openlearning_link(self):   
        import webbrowser
        url="https://cghlabm.cgmh.org.tw/course/view.php?id=884"
        webbrowser.open(url,new=2)
    def scoresearch(self):
        import ScoreSearch
        from ScoreSearch import Search
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        ScoreSearch.getaccount(Baccount)
        S = Search(self.newWindow,self.master)
    def logout_interface(self,oldmaster):
        if tk.messagebox.askyesno(title='檢驗醫學部(科)', message='確定要登出嗎?', ):
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='已登出!')
            oldmaster.deiconify()
            self.master.destroy()
        else:
            return
    def exit_interface(self,oldmaster):
        if tk.messagebox.askyesno(title='檢驗醫學部(科)', message='確定要離開嗎?', ):
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='結束能力試驗!')
            self.master.destroy()
            try:
                oldmaster.destroy()
            except AttributeError:
                pass
        else:
            return

def main():  
    root = ctk.CTk()
    B = Basedesk(root)
    # B.gui_arrang()
    # 主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  