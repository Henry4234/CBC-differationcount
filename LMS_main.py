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
        self.hospital_code = verifyAccount.hos_matrix(ht)
        self.verifyResult = verifyAccount.verifyAccountData_lms(govid,ac,self.hospital_code)
        self.login_result()
    def login_result(self):
        if self.verifyResult=='administrator':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入admin介面')
            self.loginuseradmin(self.verifyResult)
        elif self.verifyResult=='useradmin':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入使用者管理介面')
            self.loginuseradmin(self.verifyResult)
        elif self.verifyResult=='primarysupervisor':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入主要管理者介面')
            self.loginuseradmin(self.verifyResult)
        elif self.verifyResult=='secondarysupervisor':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入各院區管理者介面')
            self.loginuseradmin(self.verifyResult)
        elif self.verifyResult=='user':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='進入使用者介面')
            self.loginuser()
        elif self.verifyResult=='noGovid':  
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='先前缺少身分證字號資料!已補上!進入使用者介面')
            stat = verifyAccount.noGovid_lms(ac,self.hospital_code,govid)
            if stat == "success":
                tk.messagebox.showinfo(title='檢驗醫學部(科)', message='新增成功!進入介面')
                self.verifyResult = verifyAccount.verifyAccountData_lms(govid,ac,self.hospital_code)
                return self.login_result()
        elif self.verifyResult=='noAccount':
            tk.messagebox.showinfo(title='檢驗醫學部(科)', message='初次登入!新增後進入使用者介面')
            stat = verifyAccount.addaccount_lms(ac,ht,govid)
            if stat == "success":
                tk.messagebox.showinfo(title='檢驗醫學部(科)', message='新增使用者成功!')
                self.verifyResult='user'
                self.loginuser()
        
        elif self.verifyResult=='empty':
            tk.messagebox.showerror(title='檢驗醫學部(科)', message='參數錯誤，請聯繫管理人員!')
            return
    #user登入
    def loginuser(self):
        self.newWindow = ctk.CTk()
        # root = ctk.CTk()
        basedesk.get_accountpermission(ac)
        B = basedesk.Basedesk(self.newWindow)
        self.newWindow.mainloop()
    #admin登入
    def loginuseradmin(self,permission):
        self.newWindow = ctk.CTk()
        basedesk_admin.get_accountpermission(ac,permission)
        B = basedesk_admin.Basedesk_Admin(self.newWindow)
        self.newWindow.mainloop()



def main():  
    # 初始化物件  
    L = LMS()
if __name__ == '__main__':  
    main()