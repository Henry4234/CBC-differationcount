#authorised by Henry Tsai
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import messagebox
import customtkinter as ctk
import subprocess
from verifyAccount import changepw_sql
from setuptools import Command

# from verifyAccount import changepw
# global account
# account = sys.argv[1]

def get_accountpermission(acount,permission,govid,hoscode):
    global Baccount,Permission,Govid,Hos_code
    Permission = str(permission)
    Baccount = str(acount)
    Govid = str(govid)
    Hos_code = str(hoscode)
    return None
class Basedesk_Admin:
    
    def __init__(self,master,oldmaster=None):  
        self.master = master
        # 建立登入後視窗  
        super().__init__()
        # self.root = ctk.CTk()
        ctk.set_default_color_theme("dark-blue")  
        # 給主視窗設定標題內容  
        self.master.title("形態學考核及教育程式-v5.3")  
        self.master.geometry('650x600')
        self.master.config(background='#FFEEDD') #設定背景色
        # global account_1
        s = ttk.Style()
        s.configure('Red.TLabelframe.Label', font=('微軟正黑體', 12))
        s.configure('Red.TLabelframe.Label', foreground ='#FFFFFF')
        s.configure('Red.TLabelframe.Label', background='#FFEEDD')
        self.hello_label_user = ctk.CTkLabel(
            self.master, 
            # text = "歡迎回來",
            text = "歡迎回來 %s"%(Baccount),
            fg_color='#FFEEDD',
            font=('微軟正黑體',36,'bold'),
            text_color="#000000",
            width=240
            )
        self.hello_label_user.pack()
        self.hello_label_function = ctk.CTkLabel(
            self.master, 
            text = "請選擇需要使用的功能:", 
            fg_color='#FFEEDD',
            font=('微軟正黑體',32),
            text_color="#000000",
            width=300
            )
        self.hello_label_function.pack()
        ##框架設置
        self.labelframe_1 = ctk.CTkFrame(
            self.master,
            # text='1. 考核介面',
            fg_color="#FFDCB9",
            bg_color="#FFEEDD")
        self.labelframe_1.pack()
        self.labelframe_2 = ctk.CTkFrame(
            self.master,
            # text='2. 成績查詢',
            fg_color="#FFDCB9",
            bg_color="#FFEEDD")
        self.labelframe_2.pack()
        self.labelframe_3 = ctk.CTkFrame(
            self.master,
            # text='3. 考核匯入/成績匯出',
            fg_color="#FFDCB9",
            bg_color="#FFEEDD")
        self.labelframe_3.pack()
        self.labelframe_4 = ctk.CTkFrame(
            self.master,
            # text='4. 帳號管理',
            fg_color="#FFDCB9",
            bg_color="#FFEEDD")
        self.labelframe_4.pack()
        self.labelframe_5 = ctk.CTkFrame(
            self.master,
            # text='5. 練習介面',
            fg_color="#95C8EF",
            bg_color="#FFEEDD")
        self.labelframe_5.pack()
        self.labelframe_link = ctk.CTkFrame(
            self.master,
            # text='5. 數位學習網',
            fg_color="#BDF5CB",
            bg_color="#FFEEDD")
        self.labelframe_link.pack()
        ##各類功能選單
        self.button_blood=ctk.CTkButton(
            self.labelframe_1, 
            command = self.bloodcounter, 
            text = "血液考核",
            fg_color='#FF9900', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000",
            )
        self.button_blood.pack(padx=10, pady=15)
        self.button_urin=ctk.CTkButton(
            self.labelframe_1, 
            command = self.urinesedimentcounter, 
            text = "尿沉渣考核",
            fg_color='#FF9900', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000",
            )
        self.button_urin.pack(padx=10, pady=15)
        self.button_bodyfluid=ctk.CTkButton(
            self.labelframe_1, 
            command = self.bodyfluidcounter, 
            text = "體液考核",
            fg_color='#FF9900', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_bodyfluid.pack(padx=10, pady=15)
        self.button_practise=ctk.CTkButton(
            self.labelframe_5, 
            command = self.practise, 
            text = "練習考核",
            fg_color='#A8DEF0', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_practise.grid(row=0,column=0,padx=10, pady=15)
        self.button_image_pracrice=ctk.CTkButton(
            self.labelframe_5, 
            command = lambda:self.image_practise(2), 
            text = "圖庫練習_中階",
            fg_color='#A8DEF0', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_image_pracrice.grid(row=0,column=1,padx=10, pady=15)
        self.button_2=ctk.CTkButton(
            self.labelframe_2, 
            command = self.scoresearch,
            text = "個人成績查詢",
            fg_color='#FF9900', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000",
            )
        # self.button_2.pack(padx=10, pady=15)
        self.btn_allscore=ctk.CTkButton(
            self.labelframe_2, 
            command = None, 
            text = "全體成績查詢",
            fg_color='#FF9900', 
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.btn_allscore.pack(padx=10, pady=15)
        self.button_import=ctk.CTkButton(
            self.labelframe_3,
            command = self.scoreimport, 
            text = "考題匯入", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_import.pack(padx=10, pady=15)
        self.button_test_config=ctk.CTkButton(
            self.labelframe_3,
            command = self.testmodify, 
            text = "考題參數設定", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_test_config.pack(padx=10, pady=15)
        self.button_hos_import=ctk.CTkButton(
            self.labelframe_3,
            command = self.hos_scoreimport, 
            text = "(各院區)\n考題匯入", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_hos_import.pack(padx=10, pady=15)
        self.button_hos_test_config=ctk.CTkButton(
            self.labelframe_3,
            command = self.hos_testmodify, 
            text = "(各院區)\n考題參數設定", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_hos_test_config.pack(padx=10, pady=15)
        self.button_export=ctk.CTkButton(
            self.labelframe_2,
            command = self.output_score, 
            text = "成績匯出", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_export.pack(padx=10, pady=15)
        self.button_scoretest=ctk.CTkButton(
            self.labelframe_2,
            command = self.scorecal_test, 
            text = "成績試算", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_scoretest.pack(padx=10, pady=15)
        self.button_score=ctk.CTkButton(
            self.labelframe_2,
            command = self.scorecal, 
            text = "成績計算", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_score.pack(padx=10, pady=15)
        self.button_4=ctk.CTkButton(
            self.labelframe_4,
            command = self.IDmanage, 
            # text = "新增/刪除帳號", 
            text = "帳號管理", 
            fg_color='#FF9900',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_4.pack(padx=10, pady=15)
        self.button_link=ctk.CTkButton(
            self.labelframe_link,
            command = self.openlearning_link, 
            text = "醫檢數位學習網", 
            fg_color='#64E864',
            width=160,height=40,
            font=('微軟正黑體',22),
            text_color="#000000"
            )
        self.button_link.pack(padx=10, pady=15)
        self.button_5=ctk.CTkButton(
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
        self.button_5.pack()
        self.button_6=ctk.CTkButton(
            self.master, 
            command = lambda: self.exit_interface(oldmaster),
            text = "結束使用", 
            fg_color='#FF6600',
            bg_color='#FFEEDD',
            width=180,height=40,
            font=('微軟正黑體',24,'bold'),
            text_color="#000000"
            )
        self.button_6.pack()
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
            bg_color='#FFEEDD',
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
    # def gui_arrang(self):
        self.hello_label_user.place(relx=0.5, rely=0.05, anchor=tk.CENTER)
        self.hello_label_function.place(relx=0.5, rely=0.13, anchor=tk.CENTER)
        self.button_5.place(relx=0.2,rely=0.9,anchor=tk.CENTER)
        self.button_6.place(relx=0.5,rely=0.9,anchor=tk.CENTER)
        self.button_changepw.place(relx=0.8,rely=0.9,anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE) 
        self.versioninfo.place(relx=0, rely=1,anchor=tk.SW) 
        self.labelframe_1.place(relx=0.17,rely=0.395, anchor=tk.CENTER)
        self.labelframe_2.place(relx=0.52,rely=0.455, anchor=tk.CENTER)
        self.labelframe_3.place(relx=0.85,rely=0.49, anchor=tk.CENTER)
        self.labelframe_4.place(relx=0.85,rely=0.78, anchor=tk.CENTER)
        self.labelframe_5.place(relx=0.31,rely=0.78, anchor=tk.CENTER)
        self.labelframe_link.place(relx=0.17,rely=0.65, anchor=tk.CENTER)
        self.permission_able()
    def permission_able(self):
        if Permission=="useradmin":
            self.labelframe_5.place_forget()
            self.labelframe_4.place(relx=0.17,rely=0.78, anchor=tk.CENTER)
        elif Permission=="master":
            self.button_5.configure(state="normal")
        elif Permission=="primarysupervisor":
            self.labelframe_4.place_forget()
            self.button_4.configure(state="disabled")
        elif Permission=="secondarysupervisor":
            self.labelframe_4.place_forget()
            self.button_import.configure(state="disabled")
            self.button_test_config.configure(state="disabled")
            self.btn_allscore.configure(state="disabled")

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
    def scoresearch(self):
        import ScoreSearch
        from ScoreSearch import Search
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        ScoreSearch.getaccount(Baccount)
        S = Search(self.newWindow,self.master)
    def scoreimport(self):
        import scoreImport
        from scoreImport import Import
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        scoreImport.gethos('D')
        I = Import(self.newWindow,self.master)
        # self.newWindow['bg'] = "#FFEEDD"
    def testmodify(self):
        import testmodify
        from testmodify import Modify
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        testmodify.gethos('D')
        M = Modify(self.newWindow,self.master)
    def hos_scoreimport(self):
        import scoreImport
        from scoreImport import Import
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        scoreImport.gethos(Hos_code)
        I = Import(self.newWindow,self.master)
        # self.newWindow['bg'] = "#FFEEDD"
    def hos_testmodify(self):
        import testmodify
        from testmodify import Modify
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        testmodify.gethos(Hos_code)
        M = Modify(self.newWindow,self.master)
    def output_score(self):
        from output_score import OUTPUT_SCORE
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        OPS = OUTPUT_SCORE(self.newWindow,self.master)
    def scorecal_test(self):
        from scorecal_test import SCORE_TEST
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        CAL = SCORE_TEST(self.newWindow,self.master)
    def scorecal(self):
        from score_calculation import SCORE_CAL
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        CALT = SCORE_CAL(self.newWindow,self.master)
    def IDmanage(self):
        from IDmanage import ID
        self.master.withdraw() #把basedesk隱藏
        self.newWindow = ctk.CTkToplevel()
        I = ID(self.newWindow,self.master)
    def openlearning_link(self):   
        import webbrowser
        url="https://cghlabm.cgmh.org.tw/course/view.php?id=884"
        webbrowser.open(url,new=2)
    def logout_interface(self,oldmaster):
        if tk.messagebox.askyesno(title='檢驗醫學部(科)', message='確定要登出嗎?', ):
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='已登出!')
            oldmaster.deiconify()
            self.master.destroy()
            # subprocess.run("python main.py", shell=True)
            # from main import Login
            # root = ctk.CTk()
            # L = Login(root)  
            # root.mainloop()
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
    B = Basedesk_Admin(root)
    #  主程式執行  
    root.mainloop()  
  
  
if __name__ == '__main__':  
    main()  