#authorised by Henry Tsai
import sys,os
import verifyAccount
import basedesk, basedesk_admin
import customtkinter  as ctk
import tkinter as tk

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

ht = sys.argv[1]
ac = sys.argv[2]
govid = sys.argv[3]

# print(ac,ht)

class LMS:

    def __init__(self):
        hospital_code = verifyAccount.hos_matrix(ht)
        verifyResult = verifyAccount.verifyAccountData_lms(govid,ac,hospital_code)
        if verifyResult=='administrator':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入admin介面')
            self.loginuseradmin()
        elif verifyResult=='useradmin':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入使用者管理介面')
            self.loginuseradmin()
        elif verifyResult=='primarysupervisor':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入主要管理者介面')
            self.loginuseradmin()
        elif verifyResult=='secondarysupervisor':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入各院區管理者介面')
            self.loginuseradmin()
        elif verifyResult=='user':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入使用者介面')
            self.loginuser()
        elif verifyResult=='noGovid':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='先前缺少身分證字號資料!已補上!進入使用者介面')
            stat = verifyAccount.noGovid_lms(ac,hospital_code,govid)
            if stat == "success":
                tk.messagebox.showinfo(title='檢驗醫學部(科)', message='新增成功!進入使用者介面')
                self.loginuser()
        elif verifyResult=='noAccount':
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='初次登入!新增後進入使用者介面')
            stat = verifyAccount.addaccount_lms(ac,ht,govid)
            if stat == "success":
                tk.messagebox.showinfo(title='檢驗醫學部(科)', message='新增使用者成功!')
                self.loginuser()
        
        elif verifyResult=='empty':
            tk.messagebox.showerror(title='檢驗醫學部(科)', message='參數錯誤，請聯繫管理人員!')
            return
    #user登入
    def loginuser(self):
        self.newWindow = ctk.CTk()
        # root = ctk.CTk()
        basedesk.getaccount(ac)
        B = basedesk.Basedesk(self.newWindow)
        self.newWindow.mainloop()
    #admin登入
    def loginuseradmin(self):
        self.newWindow = ctk.CTk()
        basedesk_admin.getaccount(ac)
        B = basedesk_admin.Basedesk_Admin(self.newWindow)
        self.newWindow.mainloop()



def main():  
    # 初始化物件  
    L = LMS()  
if __name__ == '__main__':  
    main()