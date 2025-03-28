#authorised by Henry Tsai
import sys,os
import tkinter as tk
from tkinter import messagebox
from tkinter import font
from typing import Sized
import customtkinter  as ctk
import verifyAccount
import basedesk, basedesk_admin

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class Login:    #建立登入介面  
    
    #初始化設定__init__
    def __init__(self,master):
        self.master = master
        # 建立主視窗,用於容納其它元件
        super().__init__()
        # self.root = ctk.CTk()
        ctk.set_default_color_theme("dark-blue")  
        #設定字型
        # 給主視窗設定標題內容  
        self.master.title("檢驗醫學部(科)")  
        self.master.geometry('600x400')
        self.master.config(background='#323232')
        # self.account_2 = None
        self.master.bind('<Return>', self.callback)
        ####################取得連線IP####################
        # 分割字符串
        parts = verifyAccount.connection_string.split(";")
        # 找到包含SERVER的部分
        server_part = [part for part in parts if "SERVER=" in part][0]
        # 从SERVER部分提取IP地址
        server_ip = server_part.split("=")[1]
        self.ip = ctk.StringVar()
        self.ip.set(server_ip)
        ####################取得連線IP####################
        #建立圖片
        self.canvas = tk.Canvas(self.master, height=56, width=379,background="#323232",highlightthickness=0)#建立畫布
        # self.canvas.comfig(highlightthickness=0)  
        self.image_file = tk.PhotoImage(file = resource_path('assets\logo.png'))#載入圖片檔案  
        self.image = self.canvas.create_image(0,0, anchor='nw', image=self.image_file)#將圖片置於畫布上  
        self.canvas.pack(side='top')#放置畫布（為上端）  
        #建立標題
        self.label_title = ctk.CTkLabel(self.master, text='歡迎使用考核系統',height=30,font=('微軟正黑體',26),text_color="#FFFFFF",fg_color="#323232")  
        #建立一個`label`名為`Account: `  
        self.label_account = ctk.CTkLabel(self.master, text='帳號: ',height=30,font=('微軟正黑體',18),text_color="#FFFFFF",fg_color="#323232",bg_color="#323232")  
        #建立一個`label`名為`Password: `  
        self.label_password = ctk.CTkLabel(self.master, text='密碼: ',height=30,font=('微軟正黑體',18),text_color="#FFFFFF",fg_color="#323232",bg_color="#323232")
        # 建立一個賬號輸入框,並設定尺寸  
        self.input_account = ctk.CTkEntry(self.master,bg_color="#323232",height=30, width=120)
        self.cc = ctk.CTkLabel(
            self.master, 
            fg_color="#323232",
            text='@Design by Henry Tsai',
            text_color="#8E8E8E",
            font=("Calibri",12),
            width=170)
        self.label_ip = ctk.CTkLabel(
            self.master, 
            fg_color="#323232",
            text='SQL Server IP:',
            text_color="#8E8E8E",
            font=("Calibri",12),
            bg_color="#323232"
            )
        self.var_ip = ctk.CTkLabel(
            self.master, 
            fg_color="#323232",
            textvariable=self.ip,
            text_color="#8E8E8E",
            font=("Calibri",12),
            bg_color="#323232"
            )
        # 建立一個密碼輸入框,並設定尺寸  
        self.input_password = ctk.CTkEntry(self.master, show='*', bg_color="#323232",height=30, width=120)  
        # 建立一個登入系統的按鈕  
        self.login_button = ctk.CTkButton(self.master, command = self.backstage_interface, text = "登入", width=60,font=('微軟正黑體',20),fg_color="#E6883B",bg_color="#323232")
        self.login_button.pack()
        # 建立一個退出系統的按鈕  
        self.exit_button = ctk.CTkButton(self.master, command = self.exit_interface, text = "退出", width=60,font=('微軟正黑體',20),fg_color="#E6883B",bg_color="#323232")
        self.exit_button.pack()
        # 完成佈局
    # def gui_arrang(self):  
        self.canvas.place(relx=0.5, rely=0.15, anchor=tk.CENTER)
        self.label_title.place(relx=0.5, rely=0.4, anchor=tk.CENTER)  
        self.label_account.place(relx=0.4, rely=0.6, anchor=tk.CENTER)  
        self.label_password.place(relx=0.4, rely=0.75, anchor=tk.CENTER)  
        self.input_account.place(relx=0.55, rely=0.6, anchor=tk.CENTER)  
        self.input_password.place(relx=0.55, rely=0.75, anchor=tk.CENTER)  
        self.login_button.place(relx=0.4, rely=0.9, anchor=tk.CENTER)  
        self.exit_button.place(relx=0.6, rely=0.9, anchor=tk.CENTER)
        self.cc.place(relx=1, rely=1,anchor=tk.SE)
        self.label_ip.place(relx=0, rely=1,anchor=tk.SW)
        self.var_ip.place(relx=0.125, rely=1,anchor=tk.SW)
        
    # 退出介面  
    def exit_interface(self):  
        self.master.destroy()

    # 進行登入資訊驗證  
    def backstage_interface(self):  
        # with open('pw.pickle','wb') as usr_file:
        #     usrs_info={'admin':'admin','t001':'t001'}
        #     pickle.dump(usrs_info,usr_file)
        # # global account
        global account
        # global account_1
        account = self.input_account.get()
        password = self.input_password.get()
        # idreturn(account)
        #對賬戶資訊進行驗證，普通使用者返回user，管理員返回master，賬戶錯誤返回noAccount，密碼錯誤返回noPassword  
        verifyResult = verifyAccount.verifyAccountData_sql(account,password)  
        if verifyResult=='master':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入管理介面')
            self.input_account.delete(0,tk.END)
            self.input_password.delete(0,tk.END)
            self.loginuseradmin()
        elif verifyResult=='user':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入使用者介面')
            self.input_account.delete(0,tk.END)
            self.input_password.delete(0,tk.END)
            self.loginuser()   
        elif verifyResult=='noAccount':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='該賬號不存在請重新輸入!')
            self.input_account.delete(0,tk.END)
            self.input_password.delete(0,tk.END)
        elif verifyResult=='noPassword':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='賬號/密碼錯誤請重新輸入!')
            self.input_password.delete(0,tk.END)
        elif verifyResult=='empty':
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='未輸入賬號/密碼!')
    def callback(self, event):  #按Enter鍵自動連結登入
        self.backstage_interface()
    def loginuser(self):
        self.master.withdraw()
        
        self.newWindow = ctk.CTkToplevel()
        # root = ctk.CTk()
        basedesk.get_accountpermission(account)
        B = basedesk.Basedesk(self.newWindow,self.master)
        # self.master.destroy()
        # B = basedesk.Basedesk()
        # B.gui_arrang()
        # # B.mainloop()
        # B.deiconify()
        
        # command = "python basedesk.py " + account
        # subprocess.run(command, shell=True)
    def loginuseradmin(self):
        self.master.withdraw()
        # self.master.destroy()
        self.newWindow = ctk.CTkToplevel()
        basedesk_admin.get_accountpermission(account,permission='master')
        B = basedesk_admin.Basedesk_Admin(self.newWindow,self.master)
        # self.master.destroy()
        # command = "python basedesk_admin.py " + account
        # subprocess.run(command, shell=True)
def main():  
    # 初始化物件  
    root = ctk.CTk()
    L = Login(root)  
    # 進行佈局 
    # L.gui_arrang()  
    # 主程式執行
    # L.destroy
    root.mainloop()
if __name__ == '__main__':  
    main()