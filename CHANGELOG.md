# Changelog

---

## [V5.5] - 2025-11-27

### Documentation

- `6e24724` — Update README with environment setup steps
  - 補充環境建置步驟。
  - 新增 / 更新 venv environment 操作說明。

### Changed

- `e3cc9e6` — Merge pull request #11 from Henry4234/V5.5
  - 合併 V5.5 版本。
  - `score_calculation.py`：上傳清除功能調整為不清除院區與年份。
  - `output_score.py`：院區欄位 `all_combobox` 與 SQL 院區資料連動。
  - `score_calculation.py`：院區欄位 `all_combobox` 與 SQL 院區資料連動。
  - `score_calculation.py`：新增 `calculate_cell`。

- `70ae12f` — V5.5
  - 調整 `score_calculation.py` 上傳清除邏輯，避免清除院區與年份。
  - 調整 `output_score.py` 院區欄位，使 `all_combobox` 與 SQL 院區資料連動。
  - 調整 `score_calculation.py` 院區欄位，使 `all_combobox` 與 SQL 院區資料連動。

---

## [V5.4] - 2025-09-25

### Added

- `e704d14` — Merge pull request #10 from Henry4234/V5.4
  - 合併 V5.4 版本。
  - 新增「全體成績查詢」連結。
  - 修正成績匯出的資料夾選擇視窗。

- `37a081c` — V5.4
  - 新增「全體成績查詢」連結。
  - 修正成績匯出的資料夾選擇視窗。

---

## [V5.3] - 2025-05-25

### Added

- `eb23442` — Merge pull request #9 from Henry4234/V5.3
  - 合併 V5.3 版本。
  - `basedesk_admin.py`：新增 Moodle 學習網欄位。
  - `basedesk.py`：新增 Moodle 學習網欄位。

### Fixed

- `eb23442` — Merge pull request #9 from Henry4234/V5.3
  - 修復練習模式中會出現檢驗醫學部考片的問題。
  - 修復血液考核模式中會出現各院區考片的問題。

### Removed

- `eb23442` — Merge pull request #9 from Henry4234/V5.3
  - 汰除 `counter.py` 使用 `testdata.json` 的資料來源。

### Commit Details

- `da0e29c` — V5.3
  - 修復練習模式中會出現檢驗醫學部考片的問題。
  - 修復血液考核模式中會出現各院區考片的問題。
  - 汰除 `counter.py` 使用 `testdata.json` 的資料來源。

---

## [V5.2] - 2025-05-25

### Added

- `16f9fbd` — Merge pull request #8 from Henry4234/V5.1
  - 合併 V5.2 版本。
  - `basedesk_admin.py`：新增 Moodle 學習網欄位。
  - `basedesk.py`：新增 Moodle 學習網欄位。

### Fixed

- `16f9fbd` — Merge pull request #8 from Henry4234/V5.1
  - 修正 `practice_image.py` 未填答提醒題號錯誤。
  - 修正 `LMS_main.py` 中跨院區帳號驗證邏輯，跨院區帳號排除院區驗證。

### Commit Details

- `ded7601` — V5.2
  - 修正 `practice_image.py` 未填答提醒題號錯誤。
  - 修正 `LMS_main.py` 中跨院區帳號排除院區驗證。
  - `basedesk_admin.py`：新增 Moodle 學習網欄位。
  - `basedesk.py`：新增 Moodle 學習網欄位。

---

## [V5.0] - 2025-05-25

### Added

- `ba920eb` — Merge pull request #7 from Henry4234/V5.0
  - 合併 V5.0 版本。
  - 新增角色權限系統，分為五種角色：
    - 程式管理員：`administrator`
    - 角色管理員：`useradmin`
    - 主要管理者：`primarysupervisor`
    - 次要管理者：`secondarysupervisor`
    - 使用者：`user`
  - `IDmanage.py`：新增角色欄位與修改權限條件。
  - `LMS_main.py`：新增 `govid` 身分證字號驗證。

### Changed

- `ba920eb` — Merge pull request #7 from Henry4234/V5.0
  - 更新 SQL 欄位。
  - 更新 `counter.py` 片語系統。
  - `counter.py` 上傳結果識別依據由 `ac` 改為 `govid`。
  - `scoreImport.py`：依權限邏輯調整 SQL 上傳語句。
  - `testmodify.py`：依權限邏輯調整 SQL 上傳語句。

### Commit Details

- `b544621` — V5.0
  - 更新 SQL 欄位。

---

## [V4.2] - 2025-03-28

### Added

- `f82b677` — Merge pull request #6 from Henry4234/V4.2
  - 合併 V4.2 版本。
  - `basedesk_admin.py`：使用帳號登入時，允許點擊「登出」按鈕。

### Changed

- `f82b677` — Merge pull request #6 from Henry4234/V4.2
  - 更新 `basedesk.py` 版本資訊。
  - 更新 `basedesk_admin.py` 版本資訊。

### Fixed

- `f82b677` — Merge pull request #6 from Henry4234/V4.2
  - 修正 `LMS_main.py` 的 `__init__` 未加入 `login_result`，導致無法辨識身分的問題。
  - 修正 `main.py` 使用 admin 登入時回傳參數 `permission='master'` 的問題。
  - 修正 `verifyAccount.py` 因 MS SQL 版本不同，SQL `NULL` 回傳為 `None` 或空字串時，無法正常建立 `nogovid` dict 的問題。

### Commit Details

- `fcbf15c` — V4.2
  - V4.2 版本更新 commit。

---

## [V4.1] - 2025-03-27

### Fixed

- `a61438b` — Merge pull request #5 from Henry4234/V4.1
  - 合併 V4.1 版本。
  - 修正 `practice_image.py` 缺少 `generate_test`，導致 `test_dict` 無法讀取的問題。
  - 修正 `LMS_main.py` 中 `login_result`，將 `verifyAccountData_lms` 單獨生成一個函式。
  - 排除缺少身分證字號資料的帳號進入使用者介面。
  - 修正 `verifyAccount.py` 的 `refresh_sql_data`，重新取得資料庫資料。

### Commit Details

- `9c12dec` — V4.1
  - 修正 `practice_image.py` 缺少 `generate_test`，導致 `test_dict` 無法讀取的問題。
  - 修正 `LMS_main.py` 中 `login_result`，將 `verifyAccountData_lms` 單獨生成一個函式。
  - 排除缺少身分證字號資料的帳號進入使用者介面。
  - 修正 `verifyAccount.py` 的 `refresh_sql_data`，重新取得資料庫資料。

---

## [V4.0] - 2025-03-23

### Added

- `74e1671` — Merge pull request #4 from Henry4234/V4.0
  - 合併 V4.0 版本。
  - SQL 新增 `permission` table，用於角色管理名稱。
  - SQL `id` table 新增 `permission` 角色欄位。
  - SQL `id` table 新增 `govid` 身分證字號欄位。
  - 新增登入角色權限分類，共五種使用者角色。
  - `basedesk.py`：新增圖庫練習初階功能。

### Changed

- `74e1671` — Merge pull request #4 from Henry4234/V4.0
  - `LMS_main.py`：更新登入角色分類。

### Commit Details

- `19e89f0` — v4.0
  - SQL 新增 `permission` table。
  - SQL `id` table 新增 `permission` 與 `govid` 欄位。
  - 新增角色權限。
  - `basedesk.py`：新增圖庫練習初階。
  - `LMS_main.py`：更新登入角色分類。

---

## [V3.1] - 2024-10-30

### Added

- `f444d8a` — Merge pull request #3 from Henry4234/V3.0
  - 合併 V3.1 版本。
  - SQL 新增 `hospital_code` table，用於院區轉換。
  - `count.py` / `count_practice.py` 新增片語系統，雙擊可選擇片語。

### Changed

- `f444d8a` — Merge pull request #3 from Henry4234/V3.0
  - 訊息視窗標題由「土城長庚檢驗科」修正為「檢驗醫學部(科)」。
  - `LMS_main.py`：修正參數，更新為「院區 / 姓名 / id」。

### Commit Details

- `3c58484` — V3.1
  - 修正 `msgbox_title`：「土城長庚檢驗科」改為「檢驗醫學部(科)」。
  - `LMS_main.py`：修正參數，更新為「院區 / 姓名 / id」。
  - SQL 新增 `hospital_code` table，用於院區轉換。
  - `count_practice.py` 新增片語系統，雙擊可選擇片語。

---

## [V3.0] - 2024-08-10

### Added

- `2b8f219` — 部屬至LMS
  - 部署至 LMS。
  - 新增 `sys.argv`，讓程式可擷取參數資料。

### Changed

- `99ec581` — V3.0
  - V3.0 版本更新。
  - 加入 `resource_path()` 類型的資源路徑處理邏輯，以支援開發環境與 PyInstaller 打包後環境。
  - 多個 GUI / 功能模組改用打包相容的資源路徑。
  - 調整 SQL 連線與部分考題 / 計數資料讀取邏輯。

- `a29d633` — 刪除測試參數
  - 刪除測試用參數。
  - 將 `ac`、`ht` 改為由 `sys.argv[1]` / `sys.argv[2]` 取得。

---

## [V2.1] - 2024-03-26 / 2024-03-22

### Documentation

- `3b8ad74` — Update README.md
  - README 進行 V2.1 相關修改。

### Changed

- `680b71e` — Merge pull request #2 from Henry4234/V2.1
  - 合併 V2.1 版本。

- `1cbd9c0` — V2.1
  - 更新修改答案參數。
  - 主要影響檔案：`testmodify.py`。

---

## [V2.0 / Major Update] - 2024-03-20

### Changed

- `7b774f1` — Merge pull request #1 from Henry4234/v2.0
  - 合併 V2.0 重大更新。

- `7a00e01` — 重大更新
  - 更新連線方式。
  - 將 data 儲存至 MS SQL Server。

---

## [Pre-release / Early Development] - 2023-03-01 / 2023-02-27

### Documentation

- `86b9f5e` — Update README.md
  - 更新 README 文件。

- `573d84b` — Update README.md
  - 更新 README 文件。

### Added

- `71531e6` — add new document
  - 新增文件。
  - commit body：`none`

- `af491c2` — Initial commit
  - 專案初始 commit。

---

## Commit Index

| Date | Commit | Section | Summary |
|---|---:|---|---|
| 2026-05-15 | `704503a` | Unreleased / Documentation | Update README with new image and GUI section |
| 2026-05-15 | `092b0a4` | Unreleased / Documentation | Update README with new screenshot and version features |
| 2025-11-27 | `6e24724` | V5.5 | Update README with environment setup steps |
| 2025-11-27 | `e3cc9e6` | V5.5 | Merge pull request #11 |
| 2025-11-27 | `70ae12f` | V5.5 | V5.5 |
| 2025-09-25 | `e704d14` | V5.4 | Merge pull request #10 |
| 2025-09-25 | `37a081c` | V5.4 | V5.4 |
| 2025-05-25 | `eb23442` | V5.3 | Merge pull request #9 |
| 2025-05-25 | `da0e29c` | V5.3 | V5.3 |
| 2025-05-25 | `16f9fbd` | V5.2 | Merge pull request #8 |
| 2025-05-25 | `ded7601` | V5.2 | V5.2 |
| 2025-05-25 | `ba920eb` | V5.0 | Merge pull request #7 |
| 2025-04-25 | `b544621` | V5.0 | V5.0 |
| 2025-03-28 | `f82b677` | V4.2 | Merge pull request #6 |
| 2025-03-28 | `fcbf15c` | V4.2 | V4.2 |
| 2025-03-27 | `a61438b` | V4.1 | Merge pull request #5 |
| 2025-03-27 | `9c12dec` | V4.1 | V4.1 |
| 2025-03-23 | `74e1671` | V4.0 | Merge pull request #4 |
| 2025-03-23 | `19e89f0` | V4.0 | v4.0 |
| 2024-10-30 | `f444d8a` | V3.1 | Merge pull request #3 |
| 2024-10-30 | `3c58484` | V3.1 | V3.1 |
| 2024-08-10 | `a29d633` | V3.0 | 刪除測試參數 |
| 2024-08-10 | `99ec581` | V3.0 | V3.0 |
| 2024-08-10 | `2b8f219` | V3.0 | 部屬至LMS |
| 2024-03-26 | `3b8ad74` | V2.1 | Update README.md |
| 2024-03-22 | `680b71e` | V2.1 | Merge pull request #2 |
| 2024-03-22 | `1cbd9c0` | V2.1 | V2.1 |
| 2024-03-20 | `7b774f1` | V2.0 | Merge pull request #1 |
| 2024-03-20 | `7a00e01` | V2.0 | 重大更新 |
| 2023-03-01 | `86b9f5e` | Early Development | Update README.md |
| 2023-02-27 | `573d84b` | Early Development | Update README.md |
| 2023-02-27 | `71531e6` | Early Development | add new document |
| 2023-02-27 | `af491c2` | Early Development | Initial commit |
