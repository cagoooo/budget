# 動支及黏存單自動產生系統

> 桃園市龍潭區石門國民小學 — 開源行政工具　　![版本](https://img.shields.io/badge/版本-v1.2.0-blue) ![授權](https://img.shields.io/badge/授權-MIT-green)

上傳廠商報價單 PDF，自動解析品項，一鍵產出格式完整的動支及黏存單 Excel 檔案。

---

## 功能特色

- **PDF 自動解析**：拖曳或點選上傳廠商報價單 PDF，自動辨識品名規格、數量、單價
- **多行品名合併**：正確處理跨行的商品名稱（如含型號的耗材品名）
- **格式完整保留**：本地模式透過 Python openpyxl 保留原始欄寬、列高、框線、合併儲存格
- **公式自動計算**：金額、合計、民國年份、用途說明等欄位均保留原始 Excel 公式
- **雙模式運作**：本地 Python 伺服器（格式最完整）＋ GitHub Pages 靜態模式（免安裝）
- **支援兩種表單類型**：預算內 / 代收代辦
- **RWD 響應式設計**：桌機、平板（768px）、手機（375px）均可正常操作
- **隱私安全**：所有資料僅在瀏覽器端或本機處理，不上傳任何伺服器

---

## 快速開始

### 方式一：GitHub Pages 靜態模式（免安裝）

直接開啟部署好的 GitHub Pages 網址即可使用。無需安裝任何軟體。

> 注意：靜態模式使用 SheetJS 在瀏覽器端填寫 Excel，資料正確但部分格式（欄寬、框線細節）可能略有差異。

### 方式二：本地 Python 伺服器模式（格式最完整）

**需求**
- Python 3.8 以上
- openpyxl（首次啟動會自動安裝）

**步驟**

```bash
# 1. 進入專案資料夾
cd Budget

# 2. 啟動伺服器
python run.py

# 3. 瀏覽器會自動開啟 http://localhost:8000
```

啟動後畫面右上角會顯示綠色「本地模式｜格式完整保留」徽章，代表完整格式輸出模式已啟用。

---

## 使用流程

```
步驟 1  →  步驟 2  →  步驟 3  →  步驟 4
上傳 PDF    確認品項    填寫資訊    下載 Excel
```

### 步驟 1：上傳報價單 PDF

將廠商報價單 PDF 拖曳至上傳區，或點選「選擇檔案」按鈕。

系統支援常見的廠商報價單格式，自動辨識包含以下欄位的表格：
- 品名及規格（含跨行商品名稱）
- 數量
- 單價

### 步驟 2：確認解析結果

系統解析完成後，顯示可編輯的品項表格。

- **直接點選**表格中的文字即可修改
- **新增品項**按鈕可手動增加未解析到的品項
- **刪除按鈕**可移除不需要的列
- 最多支援 **8 筆品項**（超過部分不會填入 Excel）

### 步驟 3：填寫動支單資訊

| 欄位 | 說明 |
|------|------|
| 表單類型 | 選擇「預算內」或「代收代辦」 |
| 單位別 | 選擇申請單位（總務處、教務處⋯） |
| 用途說明 | 填寫採購用途（會自動填入 Excel 對應欄位） |
| 日期 | 填寫民國年月日（年份預設帶入當年） |
| 預算科目 | 選擇一級科目後可選對應二級科目 |

### 步驟 4：產生並下載

點選「產生並下載 Excel 檔案」按鈕，瀏覽器自動下載填寫完成的 `動支及黏存單_*.xlsx` 檔案。

---

## 專案結構

```
Budget/
├── index.html              # 主頁面（4 步驟 UI 流程）
├── run.py                  # 本地 Python 伺服器（ThreadedHTTPServer）
├── fill_excel.py           # openpyxl 格式保留填寫腳本
├── favicon.svg             # 網站圖示（現代瀏覽器）
├── apple-touch-icon.png    # iOS 主畫面圖示（180×180）
├── og-image.png            # 社群分享預覽圖（1200×630）
├── css/
│   └── style.css           # 完整 RWD 樣式
├── js/
│   ├── app.js              # 主應用程式邏輯
│   ├── pdf-parser.js       # PDF.js 座標式解析模組
│   └── excel-generator.js  # 雙模式 Excel 產生模組
└── template/
    └── template.xlsx       # 原始動支及黏存單範本（包含兩個工作表）
```

---

## 雙模式說明

靜態模式已升級為使用 **ExcelJS**，兩種模式均可完整保留原始格式。

| 功能 | 本地 Python 模式 | GitHub Pages 靜態模式 |
|------|:--------------:|:-------------------:|
| 欄寬、列高保留 | ✅ 完整 | ✅ 完整 |
| 框線保留 | ✅ 完整 | ✅ 完整 |
| 合併儲存格 | ✅ 完整 | ✅ 完整 |
| 公式保留 | ✅ | ✅ |
| 免安裝 | ❌ 需 Python | ✅ |
| 啟動方式 | `python run.py` | 直接開啟網址 |
| 模式徽章顏色 | 綠色 | 黃色 |

---

## 部署到 GitHub Pages

### 第一次部署

```bash
# 1. 在 GitHub 建立新 Repository（例如：budget-tool）

# 2. 初始化並推送
git init
git add .
git commit -m "init: 動支及黏存單自動產生系統"
git remote add origin https://github.com/你的帳號/budget-tool.git
git push -u origin main

# 3. 在 GitHub Repository 設定頁面
#    Settings → Pages → Source → Deploy from a branch → main / (root)
#    儲存後等待約 1-2 分鐘部署完成
```

### 更新 OG 圖片 URL（重要）

部署後請在 `index.html` 中將 OG 圖片改為**絕對路徑**，LINE / FB 分享才能正確抓取預覽圖：

```html
<!-- 將 og-image.png 改為完整 URL -->
<meta property="og:image" content="https://你的帳號.github.io/budget-tool/og-image.png">
<meta property="og:url" content="https://你的帳號.github.io/budget-tool/">
<meta property="twitter:image" content="https://你的帳號.github.io/budget-tool/og-image.png">
```

### 後續更新

```bash
git add .
git commit -m "update: 說明更新內容"
git push
```

---

## 社群分享卡片

在 LINE 或 Facebook 分享連結時，會自動顯示預覽卡片：

- **標題**：動支及黏存單自動產生系統
- **說明**：上傳廠商報價單 PDF，自動解析品項並一鍵產出格式完整的動支及黏存單 Excel
- **預覽圖**：`og-image.png`（1200 × 630 px）

> 分享前請確認 `og:image` 已設定為完整的絕對 URL（見上方說明）。

---

## 技術說明

| 技術 | 版本 | 用途 |
|------|------|------|
| PDF.js | 3.11.174 | PDF 文字座標擷取 |
| SheetJS (xlsx) | 0.18.5 | 靜態模式 Excel 讀寫 |
| openpyxl | 最新版 | 本地模式 Excel 格式保留填寫 |
| Python | 3.8+ | 本地 HTTP 伺服器 |
| Noto Sans TC | Google Fonts | 繁體中文字型 |

---

## 常見問題

**Q：解析結果不正確，怎麼辦？**
A：可在步驟 2 直接點選表格中的儲存格手動修改，或刪除/新增品項列。

**Q：PDF 解析失敗或品項為空？**
A：部分 PDF 以圖片方式儲存（掃描件）無法解析文字。請確認 PDF 可以正常選取文字。

**Q：下載的 Excel 格式跟原始範本有差異？**
A：請使用本地 Python 模式（`python run.py`），格式保留最完整。

**Q：可以用在其他學校嗎？**
A：可以。修改 `template/template.xlsx`（換成你們的範本），並調整 `fill_excel.py` 中的欄位對應設定即可。

**Q：資料會上傳到哪裡嗎？**
A：不會。所有 PDF 解析與 Excel 產生均在本機（本地模式）或瀏覽器端（靜態模式）完成，不傳送任何資料到外部伺服器。

---

## 授權

MIT License — 開源免費，歡迎修改與再利用。

---

*桃園市龍潭區石門國民小學 行政工具專案*
