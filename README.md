[![en](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/jonatasemidio/multilanguage-readme-pattern/blob/master/README_EN.md)

# CBC-differationcount

# 專案結構
```
CBC-differationcount/  (LMS — 血液抹片鏡檢考試系統)
│
├── 核心程式
│   ├── main.py                          # 舊版登入入口(V3.1後已廢止)
│   ├── LMS_main.py                      # 主程式入口（含 SQL 整合）
│   ├── basedesk.py                      # 考生桌面介面
│   └── basedesk_admin.py                # 管理員桌面介面
│
├── 計數模組
│   ├── counter.py                       # 血球計數（主）
│   ├── counter_BodyFluid.py             # 體液計數
│   ├── counter_Urin.py                  # 尿液計數
│   └── counter_practise.py              # 練習模式計數
│
├── 帳號與 ID 管理
│   ├── IDmanage.py                      # 考生 ID 管理
│   └── verifyAccount.py                 # 帳號驗證
│
├── 成績模組
│   ├── score_calculation.py             # 成績計算
│   ├── scoreImport.py                   # 考題匯入
│   ├── ScoreSearch.py                   # 成績查詢()
│   └── output_score.py                  # 成績輸出（Excel）
│
├── 圖片模組
│   └── practice_image.py               # 練習題圖片管理
│
├── 設定 / 設定檔
│   ├── requirements.txt                 # Python 套件清單
│   ├── install.txt                      # 安裝說明
│   └── README.md                        # 專案說明
│
├── assets/                             # UI 圖示與圖片資源
│   ├── logo.png / logodb.png / logoexcel.png
│   ├── check.png / ncheck.png          # 勾選狀態圖示
│   ├── person.png / score.png / pages.png
│   ├── maketest.png / output.png
│   └── iconn.png / iconn.webp
│
├── chart/
│   └── rangechart.json                 # 95細胞範圍區間設定
│
├── testdata/                           # 開發測試資料
│   ├── data.json / data_new.json / rawdata.json / score.json #測試用資料(JSON)
│   ├── 95.xlsx                         # 95細胞範圍區間(xlsx)
│   └── template_CBCDATA.xlsx           # CBC上傳考片模板(xlsx)
│
├── 開發 / 測試腳本（非正式）
│   ├── connect_test.py                 # DB 連線測試
│   ├── pd_test.py                      # pandas 功能測試
│   ├── pyxl_test.py                    # openpyxl 功能測試
│   └── test.py / test2.py / testmodify.py / scorecal_test.py
└
```
# 更新日誌
 - 儲存於CHANGELOG.md

## 0. 環境建置
- 安裝.exe
    - `python -m venv urvenvname`
    - `.\urvenvname\Scripts\activate.ps1` 進入venv
    - `pip install -r requirements.txt ` 安裝package

## 1. 登入介面 <code style="color : red">(V3.1後已廢除，修改為LMS_main.py入口)</font>
![螢幕擷取畫面 2022-12-03 154438](https://user-images.githubusercontent.com/102476562/221456935-3b566723-6ad1-46a7-ba02-b01485901e61.png)

login GUI
- 利用`bcrypt`將帳號密碼儲存至本地
- 利用Enter快捷鍵登入

## 1. LMS登入介面(LMS_main.py入口)

<img width="865" height="216" alt="image" src="https://github.com/user-attachments/assets/0941bea6-ffd3-4c54-8aac-2654a154b70b" />

 - 直接進入LMS中點擊`形態學考核及教育程式`
 - 或於terminal中輸入以下資訊: ```python V account IDpw```
    - V: 第一變數為院區代碼
    - account: 使用LMS中帳號
    - IDpw: 使用身分證字號


## 2-1. 功能介面(管理員)

<img width="652" height="689" alt="螢幕擷取畫面 2026-05-15 200836" src="https://github.com/user-attachments/assets/e0f86f4e-1524-499f-aaf9-446665251095" />

--------------------------------------------
base desk GUI (admin)
1. 考核介面
    1. 血液考核
    2. 尿沉渣考核
    3. 體液考核
2. 成績查詢
    1. 全體成績查詢
    2. 成績匯出
    3. 成績試算
    4. 成績計算
3. 考題匯入
    1. 考題匯入
    2. 考題參數設定
    3. (各院區)考題匯入
    4. (各院區)考題參數設定
4. 帳號管理
    1. 新增修改帳號
5. 醫檢數位學習網
    1. 醫檢數位學習網連結
6. 練習介面
    1. 練習考核介面
    2. 圖庫練習_中階

`basedesk_admin.getaccount(account)`函式可以擷取登入帳號，顯示在歡迎回來後方，也利後續存取db
## 2-2. 功能介面(user)
<img width="802" height="682" alt="螢幕擷取畫面 2026-05-15 202505" src="https://github.com/user-attachments/assets/c5da8077-a787-44ff-9db6-9ca62fd6fa7a" />


base desk GUI
1. 考核介面
    1. 血液考核
    2. 尿沉渣考核
    3. 體液考核
2. 練習介面
    1. 練習考核介面
    2. 圖庫練習_初階
    3. 圖庫練習_中階
3. 醫檢數位學習網
    1. 醫檢數位學習網連結
4. 成績查詢
    1. 成績查詢

## 3-1.a. 血液考核介面
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
## 3-1.b. 尿沉渣考核介面
更新中...
## 3-1.c. 體液考核介面
更新中...
## 3-2.a. 成績查詢介面<font color=#FF0000>(V5.4後已廢除，修改SAS_BI連結)</font>
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
## 3-2.b. 全體成績查詢介面

連結至[SAS](https://cghsasva2.cgmh.org.tw/SASVisualAnalytics/)

## 3-3.a. 考題匯入介面

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
## 3-3.b. 考題設定介面
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
## 3-3.c. 成績匯出介面

output_score GUI

<img width="1202" height="682" alt="螢幕擷取畫面 2024-11-27 032237" src="https://github.com/user-attachments/assets/1d5dd739-c140-4e41-b87d-bab6dc11a9af" />


### 功能介紹

成績匯出介面主要用於匯出目前已批改完成，並且具有最終成績的資料。

- 可依照院區與年份進行分項搜尋
- 可選擇單一院區，也可選擇 `ALL` 匯出所有院區資料
- 可預覽目前搜尋到的成績結果
- 可依照不同需求匯出不同格式的 Excel 檔案
- 匯出結果皆儲存為 `.xlsx`

---

### 操作步驟

1. 點選左上第一個欄位，選擇院區  
   - 若要搜尋或匯出所有院區，請選擇 `ALL`

2. 點選左上第二個欄位，選擇對應年份

3. 點擊「搜尋」

4. 搜尋完成後，可於右上方預覽所選取的成績結果

5. 依照需求選擇匯出格式

---

### 匯出格式

成績匯出介面目前提供三種匯出格式：

1. 匯出考試結果
2. 匯出考生原始結果
3. 匯出 SAS

---

#### 1. 匯出考試結果

「匯出考試結果」會匯出目前已批改完成的考試結果，總共包含兩個表格。

#### 第一個表格：考片答案資訊

此表格為考片的標準答案資訊，包含「必須要打到的細胞」以及「不可打到的細胞」。

| 欄位名稱 | 說明 |
|---|---|
| 考片ID | 考片編號 |
| 必須要打到的細胞 / 不可打到細胞 | 此細胞是否屬於必須出現或不可出現的答案設定 |
| 各項細胞 | 對應的細胞種類 |

#### 第二個表格：考生考試結果

此表格為依照院區與年份搜尋後所選取的考試結果。

| 欄位名稱 | 說明 |
|---|---|
| 院區 | 考生所屬院區 |
| 姓名 | 考生姓名 |
| 年份 | 考試年份 |
| 考片ID | 考片編號 |
| count | 考試次數 |
| celltype | 細胞種類 |
| matrix_value | 該細胞是否在允許範圍內 |
| timestamp | 交卷時間 |

---

#### 2. 匯出考生原始結果

「匯出考生原始結果」會匯出考生作答的百分比結果，並同時附上該考片答案的上下限範圍，總共包含兩個表格。

#### 第一個表格：考片原始解答

此表格為考片原始答案資料。

| 欄位名稱 | 說明 |
|---|---|
| 考片ID | 考片編號 |
| count_value | 該考片應計數的細胞總數 |
| 各項細胞 | 各細胞種類的標準答案數值或比例 |

#### 第二個表格：考生百分比結果

此表格與「匯出考試結果」不同，主要匯出考生的百分比結果，並附上該考片各細胞的上下限。

| 欄位名稱 | 說明 |
|---|---|
| 院區 | 考生所屬院區 |
| 姓名 | 考生姓名 |
| year | 考試年份 |
| 考片ID | 考片編號 |
| count | 考試次數 |
| celltype | 細胞種類 |
| percent_value | 考生百分比結果 |
| timestamp | 交卷時間 |
| lower | 下限值 |
| upper | 上限值 |

---

#### 3. 匯出 SAS

「匯出 SAS」會將上述考試結果與原始百分比結果整合為單一表格，主要提供給 SAS 作為 BI 讀取與後續分析使用。

#### 匯出欄位

| 欄位名稱 | 說明 |
|---|---|
| 院區 | 考生所屬院區 |
| 姓名 | 考生姓名 |
| 年份 | 考試年份 |
| 考片ID | 考片編號 |
| count | 考試次數 |
| celltype | 細胞種類 |
| cellpercent | 考生百分比結果 |
| ansvalue | 百分比解答 |
| percent_lower | 下限值 |
| percent_upper | 上限值 |
| matrix_value | 該細胞是否在允許範圍內 |
| must | 必須要打到的細胞 |
| mustnot | 不可打到的細胞 |
| abn_lym | 是否納入異常淋巴球相加計算 |
| score | 最後總分 |
| timestamp | 交卷時間 |

---
## 3-4.a. 新增/刪除帳號介面
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

點擊兩下密碼視窗可以在彈出視窗進行密碼修改。<font color=#FF0000>(V3.0後已廢除，修改為LMS_main.py入口後，利用身分證字號作為pw)</font>

![螢幕擷取畫面 2023-02-23 040527](https://user-images.githubusercontent.com/102476562/221460557-b7c332de-040c-4933-bff4-5927c17e14ba.png)

點擊右下方「刪除」，直接刪除帳號。
