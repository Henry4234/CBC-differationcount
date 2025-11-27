# CBC-differationcount

# 0. 環境建置
- 安裝.exe
    - `python -m venv urvenvname`
    - `.\urvenvname\Scripts\activate.ps1` 進入venv
    - `pip install -r requirements.txt ` 安裝package

# 1. 登入介面
![螢幕擷取畫面 2022-12-03 154438](https://user-images.githubusercontent.com/102476562/221456935-3b566723-6ad1-46a7-ba02-b01485901e61.png)

login GUI
- 利用`bcrypt`將帳號密碼儲存至本地
- 利用Enter快捷鍵登入

# 2-1. 功能介面(管理員)

![螢幕擷取畫面 2024-03-26 092247](https://github.com/Henry4234/CBC-differationcount/assets/102476562/ea9712a8-5cbd-44de-84ca-7cfac29508f4)

v2.1 版本新增功能

1. 考核介面
    1. 練習模式
2. 成績查詢
    1. 成績計算
    2. 成績試算
--------------------------------------------
base desk GUI (admin)
1. 考核介面
    1. 血液考核
    2. 尿沉渣考核
    3. 體液考核
2. 成績查詢
    1. 成績查詢
    2. 全體成績查詢
3. 考題匯入/成績匯出
    1. 考題匯入
    2. 考題設定
4. 帳號管理
    1. 新增修改帳號
`basedesk_admin.getaccount(account)`函式可以擷取登入帳號，顯示在歡迎回來後方，也利後續存取db
# 2-2. 功能介面(user)
![螢幕擷取畫面 2022-12-13 175317](https://user-images.githubusercontent.com/102476562/221457361-198da60c-def7-4cbd-8cfb-5100df971a85.png)

base desk GUI
1. 考核介面
    1. 血液考核
    2. 尿沉渣考核
    3. 體液考核
2. 成績查詢
    1. 成績查詢
# 3-1.a. 血液考核介面
![螢幕擷取畫面 2023-02-19 030653](https://user-images.githubusercontent.com/102476562/221457747-a1971578-7850-46a9-80c7-dcde14f3df4f.png)

Counter GUI
主要分四個frame: 考核片選擇、功能區、考核片資訊、計數器
- 考核片選擇
    1. 考核年度選擇
    2. 考核片選擇
    3. 開始測驗
- 功能區
    1. 上/下數
    2. 歸0
    3. 交卷
    4. 返回
- 考核片資訊
    1. WBC
    2. RBC
    3. Hct
    4. MCV
    5. MCH
    6. MCHC
    7. RDW
    8. Plt
- 計數器
    1. 總數
# 3-1.b. 尿沉渣考核介面
更新中...
# 3-1.c. 體液考核介面
更新中...
# 3-2.a. 成績查詢介面
![螢幕擷取畫面 2023-02-19 035240](https://user-images.githubusercontent.com/102476562/221458313-df38a04c-2e12-421f-9625-ed48dc2d674d.png)

ScoreSearch GUI
成績查詢
- 成績面板
- 考題分析(更新中)
- 離開(返回功能介面)
會利用`basedesk.py`傳入的`account`去搜尋`rawdata.json`裡面相符帳號的紀錄，統計總共考了幾次，今年度考了幾次

在右手邊利用`matplotlib`劃一個dashboard

增加游標事件(`mplcursors`)，可以在指標指到該點時，顯示相關資訊
(e.g.成績上傳時間/分數)
# 3-2.b. 全體成績查詢介面
更新中...
# 3-3.a. 考題匯入介面
![螢幕擷取畫面 2022-12-29 221417](https://user-images.githubusercontent.com/102476562/221458789-d37c9241-0152-4305-bb16-b028967f237b.png)

scoreImport GUI
選擇匯入檔案
- 選擇匯入檔案
    - 血液
    - 尿液
    - 體液
- 查看匯入資訊
- 上傳資料庫

範例檔案： [test_CBCDATA.xlsx](https://github.com/Henry4234/CBC-differationcount/files/10835580/test_CBCDATA.xlsx)

1. 上傳範例檔案檢查：副檔名格式 `.xlsx / .cvs`，如果正確會有匯入成功字樣，反之會有匯入失敗。
    
    同時，中間驗證視窗中也會顯示匯入檔案路徑及匯入考題資訊，並且在下方多出驗證按鈕
```json
{
    "blood": [
  {
      "ID": "B001",
      "testtype": "blood",
      "rawdata": {
      "WBC": "16.7",
      "RBC": "4.65",
      "HB": "12.3",
      "Hct": "37.5",
      "MCV": "80.6",
      "MCH": "26.5",
      "MCHC": "32.8",
      "RDW": "13.3",
      "Plt": "301"
      }
  }
  ]
}
```
1. 上傳範例資料驗證兩件事：
    1. 比對上傳檔案中，是否為資料庫要的資料(e.g. Hct/MCV/MCHC…)
    2. 比對資料庫中，是否有重複考片資料
    
    ```python
    self.rawdata = rawdata["blood"][0]["rawdata"]
    self.rawdata_a = rawdata["blood"][0]["Ans"]
    .
    .
    .
    if self.testname in ID:
    ```
    
    驗證通過後會有驗證成功字樣，反之會有驗證失敗。驗證成功後，右邊上傳至資料庫按鈕解鎖。
2. 上傳至資料庫：按下上傳至資料庫按鈕後，即可上傳至資料庫中。
# 3-3.b. 考題設定介面
![螢幕擷取畫面 2023-02-19 030614](https://user-images.githubusercontent.com/102476562/221460189-bfaf86f8-e6ad-4ae2-8a8c-5e2dc36964a8.png)

testmodify GUI
考題設定
- 血液考題設定
- 尿液考題設定(更新中)
- 返回(返回功能介面)
考題設定:先選擇考核年度
會在下方`listbox`中跳出目前database該年上傳的考片資料，選擇後可以在旁邊格子內看到`CBC DATA`數值跟 `Ans`答案數值，

需要編輯的話，按下方第一顆按鍵編輯`self.edit_btn command=self.clicktable`即可編輯

修改完畢後按確定`self.yes_btn command=self.editdata` 完成修改

如果選錯想要清除按清除 `self.clear_btn command=self.clear` 清除所選擇的考片
# 3-3.c. 成績匯出介面
更新中...
# 3-4.a. 新增/刪除帳號介面
![螢幕擷取畫面 2022-12-29 221125](https://user-images.githubusercontent.com/102476562/221460481-1869bd8f-6b12-4b27-a812-2609330350fa.png)

IDmanage GUI
- 新增帳號
- 修改/刪除帳號
- 離開(返回功能介面)
`input_ID`帳號的變數，`input_pw`為密碼變數，會先比對兩次輸入密碼是否相同，確認後利用`bcrypt`hash後將資訊存入`in.json`中

如果需要新增帳號，在上方輸入帳號密碼，密碼重複輸入一次，按下確定即可。

如果輸入錯誤，按下清除即可清除格子內輸入內容
![螢幕擷取畫面 2023-02-19 040027](https://user-images.githubusercontent.com/102476562/221460354-08e2c209-b1f7-40ab-a976-a4c56e37c09f.png)

修改/刪除帳號中可以在左側的選單中，選取所有存在`in.json`中的帳號資訊，選取後右邊欄位會顯示目前選擇的帳號資訊，

點擊兩下帳號視窗可以進行帳號修改。修改完畢後，按下左下「帳號修改」即可修改帳號。
點擊兩下密碼視窗可以在彈出視窗進行密碼修改。
![螢幕擷取畫面 2023-02-23 040527](https://user-images.githubusercontent.com/102476562/221460557-b7c332de-040c-4933-bff4-5927c17e14ba.png)

點擊右下方「刪除」，直接刪除帳號。
