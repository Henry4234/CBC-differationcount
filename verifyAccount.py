import pandas as pd
import sqlalchemy as sa


import bcrypt
import json
import pymysql
##bcrypt pw

import pyodbc
#建立與mySQL連線資料
connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=220.133.50.28;DATABASE=bloodtest;UID=cgmh;PWD=B[-!wYJ(E_i7Aj3r"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.InterfaceError:
    #如果有遇到InterfaceError的話，應該是在院內網路的環境，更改ip
    # connection_string = "DRIVER={ODBC Driver 11 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
    connection_string = "DRIVER={SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)


##建立dict 包含帳號密碼acpw
with coxn.cursor() as cursor:
    query = "SELECT ac,pw  FROM [bloodtest].[dbo].[id];"
    cursor.execute(query)
    acpw = dict(cursor.fetchall())

##建立dict 包含院區帳號acht
with coxn.cursor() as cursor:
    query = "SELECT ac,院區  FROM [bloodtest].[dbo].[id];"
    cursor.execute(query)
    acht = dict(cursor.fetchall())
# print(acht)
def refresh_acpw():
    global acpw
    query = "SELECT ac,pw  FROM [bloodtest].[dbo].[id];"
    cursor.execute(query)
    acpw = dict(cursor.fetchall())
# with conn.cursor() as cursor:
#     cursor.execute("SELECT `ac`, `pw` FROM `id`;")
# acpw = cursor.fetchall()
# acpw = dict((x, y) for x, y in acpw)

# print(acpw)

def verifyAccountData(account,password):
    md = {}
    jsonfile = open('in.json','rb')
    a = json.load(jsonfile)
    ml = a['member']
    for i in ml:
        x = i['ID']
        y = i['token']
        md[x] = y
    # print(md)
    if account in md:
        ddd = md[account]
        ddd = bytes(ddd,encoding="utf8")
        password = bytes(password,encoding="utf8")
        if account == "admin" and bcrypt.checkpw(password,ddd)==True:
            return "master"
        elif bcrypt.checkpw(password,ddd)==True:
            return "user"
        else:
            return "noPassword"
    #使用者名稱密碼不能為空
    elif account=='' or password=='' :
        return "empty"
    #不在資料庫中彈出是否註冊的框
    else:
        return "noAccount"
##(desert)change pw(for manager)
def changepw(account,newpassword,oldpassword):
    jsonfile = open('in.json','rb')
    a = json.load(jsonfile)
    newpassword = bytes(newpassword,encoding="utf8")
    oldpassword = bytes(oldpassword,encoding="utf8")
    ml = a['member']
    md = {}
    for i in ml:
        x = i['ID']
        y = i['token']
        md[x] = y
    IL = [key for key in md.keys()]
    idx = IL.index(account)
    #verifyold
    old = bytes(ml[idx]['token'],encoding="utf8")
    if bcrypt.checkpw(oldpassword,old)==True:
        pass
    else:
        return "wrongoldpassword"
    #gen
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(newpassword,salt)
    ml[idx]['token'] = hashed.decode('utf-8')
    with open('in.json','w') as r:
        json.dump(a,r)
        r.close()
    return "success"
##(desert)change pw(for user)
def changepw2(account,newpassword):
    jsonfile = open('in.json','rb')
    a = json.load(jsonfile)
    newpassword = bytes(newpassword,encoding="utf8")
    ml = a['member']
    md = {}
    for i in ml:
        x = i['ID']
        y = i['token']
        md[x] = y
    IL = [key for key in md.keys()]
    idx = IL.index(account)
    #gen
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(newpassword,salt)
    ml[idx]['token'] = hashed.decode('utf-8')
    with open('in.json','w') as r:
        json.dump(a,r)
        r.close()
    return "success"
##(desert)add account(for manager)
def addaccount(account,password):
    jsonfile = open('in.json','rb')
    a = json.load(jsonfile)
    password = bytes(password,encoding="utf8")
    ml = a['member']
    md = {}
    for i in ml:
        x = i['ID']
        y = i['token']
        md[x] = y
    IL = [key for key in md.keys()]
    if account in IL:
        return "duplicate"
    else:
        #gen
        salt = bcrypt.gensalt()
        hashed = bcrypt.hashpw(password,salt)
        indc = {"ID":account,"token":hashed.decode('utf-8')}
        ml.append(indc)
        with open('in.json','w') as r:
            json.dump(a,r)
            r.close()
        return "success"
##(UPDATE)delete account(for manager)
def delaccount(account):
    with coxn.cursor() as cursor:
        sql = """DELETE FROM [bloodtest].[dbo].[id] WHERE [ac]='%s';"""%(account)
        cursor.execute(sql)
    coxn.commit()
    # jsonfile = open('in.json','rb')
    # a = json.load(jsonfile)
    # ml = a['member']
    # md = {}
    # for i in ml:
    #     x = i['ID']
    #     y = i['token']
    #     md[x] = y
    # IL = [key for key in md.keys()]
    # idx = IL.index(account)
    # ml.pop(idx)
    # # print(ml)
    # with open('in.json','w') as r:
    #     json.dump(a,r)
    #     r.close()
    return "success"
##(UPDATE)edit account(for manager)
def editaccount(oldaccount,newaccount):
    with coxn.cursor() as cursor:
        sql = """UPDATE [bloodtest].[dbo].[id] SET [ac] = '%s' WHERE [ac]='%s';"""%(newaccount,oldaccount)
        cursor.execute(sql)
    coxn.commit()
    # jsonfile = open('in.json','rb')
    # a = json.load(jsonfile)
    # ml = a['member']
    # md = {}
    # for i in ml:
    #     x = i['ID']
    #     y = i['token']
    #     md[x] = y
    # IL = [key for key in md.keys()]
    # idx = IL.index(oldaccount)
    # ml[idx]["ID"] = newaccount
    # # print(ml)
    # with open('in.json','w') as r:
    #     json.dump(a,r)
    #     r.close()
    return "success"

def verifyAccountData_sql(account,password):
    # with open('pw.pickle','rb') as usr_file:
    #         usrs_info=pickle.load(usr_file)
    if account in acpw:
        ddd = acpw[account]
        ddd = bytes(ddd,encoding="utf8")
        password = bytes(password,encoding="utf8")
        if account == "admin" and bcrypt.checkpw(password,ddd)==True:
            return "master"
        elif bcrypt.checkpw(password,ddd)==True:
            return "user"
        else:
            return "noPassword"
    #使用者名稱密碼不能為空
    elif account=='' or password=='' :
        return "empty"
    #不在資料庫中彈出是否註冊的框
    else:
        return "noAccount"

def changepw_sql(account,newpassword,oldpassword=None):
    newpassword = bytes(newpassword,encoding="utf8")
    if oldpassword == None:
        pass
    else:
        oldpassword = bytes(oldpassword,encoding="utf8")
        #verifyold
        old = bytes(acpw[account],encoding="utf8")
        if bcrypt.checkpw(oldpassword,old)==True:
            pass
        else:
            return "wrongoldpassword"
    #gen
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(newpassword,salt)
    update = hashed.decode('utf-8')
    with coxn.cursor() as cursor:
        sql = "UPDATE [bloodtest].[dbo].[id] SET [pw] = '%s' WHERE [ac]='%s';"%(update,account)
        cursor.execute(sql)
    coxn.commit()
    return "success"

def addaccount_sql(branch, account,password):
    password = bytes(password,encoding="utf8")
    refresh_acpw()
    if account in acpw:
        return "duplicate"
    else:
        #gen
        salt = bcrypt.gensalt()
        hashed = bcrypt.hashpw(password,salt)
        update = hashed.decode('utf-8')
        with coxn.cursor() as cursor:
            sql = """INSERT INTO [bloodtest].[dbo].[id] ("院區","ac","pw") VALUES ('%s','%s','%s');"""%(branch,account,update)
            cursor.execute(sql)
        coxn.commit()
        return "success"
#連接登入LMS
def verifyAccountData_lms(account,hospital):
    # with open('pw.pickle','rb') as usr_file:
    #         usrs_info=pickle.load(usr_file)
    if account in acht:
        hos = acht[account]
        
        if account == "admin" and hos == hospital:
            return "master"
        elif hos == hospital:
            return "user"
        else:
            return "nohos"
    #使用者名稱密碼不能為空
    elif account=='' or hospital=='' :
        return "empty"
    #不在資料庫中彈出是否註冊的框
    else:
        return "noAccount"
    
def addaccount_lms(account,hospital):
    with coxn.cursor() as cursor:
            sql = """INSERT INTO [bloodtest].[dbo].[id] ("院區","ac") VALUES ('%s','%s');"""%(hospital,account)
            cursor.execute(sql)
    coxn.commit()
    return "success"
