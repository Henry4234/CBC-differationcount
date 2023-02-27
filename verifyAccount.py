import bcrypt
import json
##bcrypt pw
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
##change pw(for manager)
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
##change pw(for user)
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
##add account(for manager)
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
##delete account(for manager)
def delaccount(account):
    jsonfile = open('in.json','rb')
    a = json.load(jsonfile)
    ml = a['member']
    md = {}
    for i in ml:
        x = i['ID']
        y = i['token']
        md[x] = y
    IL = [key for key in md.keys()]
    idx = IL.index(account)
    ml.pop(idx)
    # print(ml)
    with open('in.json','w') as r:
        json.dump(a,r)
        r.close()
    return "success"
##edit account(for manager)
def editaccount(oldaccount,newaccount):
    jsonfile = open('in.json','rb')
    a = json.load(jsonfile)
    ml = a['member']
    md = {}
    for i in ml:
        x = i['ID']
        y = i['token']
        md[x] = y
    IL = [key for key in md.keys()]
    idx = IL.index(oldaccount)
    ml[idx]["ID"] = newaccount
    # print(ml)
    with open('in.json','w') as r:
        json.dump(a,r)
        r.close()
    return "success"