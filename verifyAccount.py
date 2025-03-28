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


##建立dict-> key:govid身分證字號, value:[account,permission]
#整合進refresh_sql_data

# with coxn.cursor() as cursor:
#     query = "SELECT govid,ac,[permission].[permission_en],[id].[院區] FROM [bloodtest].[dbo].[id] JOIN [bloodtest].[dbo].[permission] ON [permission].[Order] = [id].[permission];"
#     cursor.execute(query)
#     govid_ac_permission = cursor.fetchall()
# # print(govid_ac_permission)

# no_govid=[item for item in govid_ac_permission if item[0]==""]
# govid_ac_permission=[item for item in govid_ac_permission if item[0]!=""]

# dict_id ={key[0]:key[1:] for key in govid_ac_permission}
#dict_id = {"govid": ('ac','permission','院區'),...}
# print(dict_id)
# print(no_govid)


##建立hospital_code_dict 包含院區帳號acht
with coxn.cursor() as cursor:
    query = "SELECT [院區],code FROM [bloodtest].[dbo].[hospital_code];"
    cursor.execute(query)
    ht_code = dict(cursor.fetchall())
ht_code['林口'] = [ht_code['林口'],'4','L']
ht_code['土城'] = [ht_code['土城'],'C']
ht_code['大里仁愛'] = [ht_code['大里仁愛'],'I']

# print(ht_code)

def refresh_acpw():
    global acpw
    query = "SELECT ac,pw  FROM [bloodtest].[dbo].[id];"
    cursor.execute(query)
    acpw = dict(cursor.fetchall())
# with conn.cursor() as cursor:
#     cursor.execute("SELECT `ac`, `pw` FROM `id`;")
# acpw = cursor.fetchall()
# acpw = dict((x, y) for x, y in acpw)


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
    #轉成代碼
    branch = hos_rematrix(branch)
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
def hos_matrix(hospital_code):
    for key, val in ht_code.items():
        # 如果值是list，檢查是否包含輸入值
        if isinstance(val, list) and hospital_code in val:
            return key
        # 如果值是單一值，直接比較
        elif val == hospital_code:
            return key
    return "找不到對應的地點"

def hos_rematrix(hospital_name):
    #如果回傳的是一個list的話
    if isinstance(ht_code[hospital_name], list):
        return str(ht_code[hospital_name][0])
    else:
        return str(ht_code[hospital_name])


def refresh_sql_data():
    global no_govid,govid_ac_permission,dict_id
    with coxn.cursor() as cursor:
        query = "SELECT govid,ac,[permission].[permission_en],[id].[院區] FROM [bloodtest].[dbo].[id] JOIN [bloodtest].[dbo].[permission] ON [permission].[Order] = [id].[permission];"
        cursor.execute(query)
        govid_ac_permission = cursor.fetchall()

    no_govid=[item for item in govid_ac_permission if item[0]==None or item[0]==""]
    govid_ac_permission=[item for item in govid_ac_permission if item[0]!=""]

    dict_id ={key[0]:key[1:] for key in govid_ac_permission}

#連接登入LMS
def verifyAccountData_lms(govid,account,hospital_name):
    refresh_sql_data()
    hospital_code = hos_rematrix(hospital_name)
    #傳進來的hospital會是代碼，需要經過轉換
    if govid in dict_id:
        if dict_id[govid][0]==account and dict_id[govid][2]==hospital_code:
            return dict_id[govid][1]
    #不在資料庫中彈出是否註冊的框
    else:
        for no_govid_item in no_govid:
            # print(no_govid_item[1])
            # print(account)
            # print(no_govid_item[3])
            # print(hospital_code)
            if no_govid_item[1] == account and no_govid_item[3] == hospital_code:
                return "noGovid"
            else:
                continue
        return "noAccount"

#如果table[id]找不到account，或者有account，但沒有院區(可能為重名)，那就INSERT一個帳號給他
def addaccount_lms(account,hospitalcode,govid):
    if hospitalcode=='4'or hospitalcode=='L'or hospitalcode=='C'or hospitalcode=='I':
        hospitalname = hos_matrix(hospitalcode)
        hospitalcode = hos_rematrix(hospitalname)
    with coxn.cursor() as cursor:
        sql = """INSERT INTO [bloodtest].[dbo].[id] ("院區","ac","govid",permission) VALUES ('%s','%s','%s',4);"""%(hospitalcode,account,govid)
        cursor.execute(sql)
    coxn.commit()
    return "success"

def noGovid_lms(account,hospital_name,govid):
    hospitalcode = hos_rematrix(hospital_name)
    with coxn.cursor() as cursor:
        sql = """UPDATE [bloodtest].[dbo].[id] SET [govid]='%s' WHERE [ac]='%s' AND [院區]='%s' ;"""%(govid,account,hospitalcode)
        cursor.execute(sql)
    coxn.commit()
    return "success"

refresh_sql_data()