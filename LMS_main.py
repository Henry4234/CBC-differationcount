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

# ac = sys.argv[1]
# ht = sys.argv[2]
ac = ""
ht = "土城"

print(ac,ht)

class LMS:

    def __init__(self):
        verifyResult = verifyAccount.verifyAccountData_lms(ac,ht)
        print(verifyResult)
        if verifyResult=='master':  
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='進入管理介面')
            self.loginuseradmin()
        elif verifyResult=='user':  
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='進入使用者介面')
            self.loginuser()
        elif verifyResult=='nohos' or verifyResult=='noAccount':
            tk.messagebox.showinfo(title='土城醫院檢驗科', message='初次登入!新增後進入使用者介面')
            stat = verifyAccount.addaccount_lms(ac,ht)
            if stat == "success":
                tk.messagebox.showinfo(title='土城醫院檢驗科', message='新增使用者成功!')
                self.loginuser()
        elif verifyResult=='empty':
            tk.messagebox.showerror(title='土城醫院檢驗科', message='參數錯誤，請聯繫管理人員!')
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