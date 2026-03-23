# 未來優化改良建議

> 動支及黏存單自動產生系統 — 開發路線圖
> Made with 💝 by [阿凱老師](https://www.smes.tyc.edu.tw/modules/tadnews/page.php?ncsn=11&nsn=16#a5)
> 最後更新：2026-03-23（v1.3.0）

---

## 目前已完成進度總覽

| 版本 | 功能 | 狀態 |
|------|------|------|
| v1.0.0 | PDF 拖曳上傳與解析 | ✅ 完成 |
| v1.0.0 | 座標式欄位偵測（品名/數量/單價） | ✅ 完成 |
| v1.0.0 | 多行品名合併 | ✅ 完成 |
| v1.0.0 | 可編輯品項表格 | ✅ 完成 |
| v1.0.0 | 預算內 / 代收代辦 雙表單 | ✅ 完成 |
| v1.0.0 | 多層預算科目下拉選單 | ✅ 完成 |
| v1.0.0 | RWD 響應式設計（桌機/平板/手機） | ✅ 完成 |
| v1.0.0 | GitHub Pages 靜態部署 | ✅ 完成 |
| v1.1.0 | 本地 Python 模式（openpyxl 格式保留） | ✅ 完成 |
| v1.1.0 | 多執行緒 HTTP 伺服器（ThreadingMixIn） | ✅ 完成 |
| v1.1.0 | 模式自動偵測徽章（本地/靜態） | ✅ 完成 |
| v1.1.0 | Excel 開啟時正確顯示工作表（activeTab 修正） | ✅ 完成 |
| v1.2.0 | favicon.svg 網站圖示 | ✅ 完成 |
| v1.2.0 | OG 社群分享卡片（LINE / FB） | ✅ 完成 |
| v1.2.0 | apple-touch-icon（iOS 主畫面） | ✅ 完成 |
| v1.2.0 | Cache busting（?v= 版本參數） | ✅ 完成 |
| v1.3.0 | JSZip 直接修補 XML（取代 ExcelJS） | ✅ 完成 |
| v1.3.0 | 根除 Excel XML 損壞問題 | ✅ 完成 |
| v1.3.0 | Footer 版權資訊（阿凱老師超連結） | ✅ 完成 |

---

## 未來優化建議

以下依「重要性」與「實作難度」分級，供參考排程。

---

### 🔴 優先等級 A：實用性最高，建議優先實作

---

#### A1. 掃描式 PDF 支援（OCR）

**問題**：目前只支援「可選取文字」的 PDF。掃描件（如紙本報價單拍照轉 PDF）無法解析。

**解決方案**：整合 [Tesseract.js](https://tesseract.projectnaptha.com/)（免費 OCR 引擎，瀏覽器端執行）。

**實作概念**：
```
PDF 上傳
  ↓ 嘗試 PDF.js 文字擷取
  ↓ 若文字量過少（< 10 字元）→ 判定為掃描件
  ↓ PDF.js 轉 canvas 圖片
  ↓ Tesseract.js OCR 辨識繁體中文
  ↓ 再送入現有欄位解析邏輯
```

**CDN**：`https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js`

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥🔥（極高）

---

#### A2. 多頁 PDF 支援

**問題**：目前只解析 PDF 第一頁。有些廠商報價單品項跨多頁。

**解決方案**：在 `pdf-parser.js` 中迴圈解析所有頁面，合併結果後統一送入欄位偵測。

**實作概念**：
```javascript
// 目前：只讀第 1 頁
const page = await pdf.getPage(1);

// 改為：讀所有頁
for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    // 累積所有 items
}
```

**難度**：⭐（容易）　**效益**：🔥🔥（高）

---

#### A3. 品項數量上限提升（超過 8 筆）

**問題**：Excel 範本目前只有 8 列品項。若廠商報價超過 8 筆，多餘品項會被丟棄。

**解決方案**：
- **短期**：在 UI 加上醒目警告「共 N 筆，僅填入前 8 筆，請手動補填剩餘品項」
- **長期**：修改 Excel 範本，增加更多品項列（需同步更新 `fill_excel.py` 的 `ITEM_ROWS` 陣列與 `excel-generator.js`）

**難度**：⭐（容易，UI 警告）/ ⭐⭐⭐（修改範本）　**效益**：🔥🔥（高）

---

#### A4. 表單資料本地記憶（LocalStorage）

**問題**：每次開啟網頁，單位別、預算科目等欄位都要重新選擇，重複操作耗時。

**解決方案**：使用 `localStorage` 記住上次填寫的選項。

**實作概念**：
```javascript
// 儲存
localStorage.setItem('budget_unit', selectedUnit);
localStorage.setItem('budget_category', selectedCategory);

// 讀取（頁面載入時）
const lastUnit = localStorage.getItem('budget_unit');
if (lastUnit) unitSelect.value = lastUnit;
```

**難度**：⭐（容易）　**效益**：🔥🔥（高）

---

#### A5. 列印 / 直接輸出 PDF 功能

**問題**：目前只能下載 Excel，還需要再開 Excel 列印。

**解決方案**：整合 [jsPDF](https://github.com/parallax/jsPDF) + html2canvas，或直接呼叫 `window.print()` 搭配列印 CSS。

**簡易方案**（不需額外套件）：
```css
/* 在 style.css 加入 */
@media print {
    /* 隱藏非必要元素，只顯示表單預覽 */
    .step-upload, .step-review, .mode-badge { display: none; }
    .print-preview { display: block; }
}
```

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

### 🟡 優先等級 B：提升體驗，建議中期實作

---

#### B1. 多份報價單批次處理

**問題**：每次只能處理一份 PDF，若有多家廠商比價需逐一上傳。

**解決方案**：支援一次上傳多個 PDF，系統自動產生多份 Excel 並打包成 ZIP 下載。

**實作概念**：
```javascript
// input 加入 multiple 屬性
<input type="file" accept=".pdf" multiple>

// 對每個檔案迴圈處理，最後用 JSZip 打包
const masterZip = new JSZip();
for (const file of files) {
    const items = await parseFile(file);
    const xlsx = await generateExcel(items);
    masterZip.file(`動支及黏存單_${filename}.xlsx`, xlsx);
}
masterZip.generateAsync({type:'blob'}).then(downloadZip);
```

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### B2. 解析結果信心分數 + 顏色標示

**問題**：解析後使用者不知道哪些品項解析正確、哪些可能有誤。

**解決方案**：在品項表格中加入信心度標示：
- 🟢 綠色：品名、數量、單價均完整
- 🟡 黃色：缺少數量或單價（需補填）
- 🔴 紅色：僅有品名，數量/單價缺失

**難度**：⭐（容易）　**效益**：🔥🔥（高）

---

#### B3. 深色模式（Dark Mode）

**問題**：長時間使用白色介面對眼睛較疲勞，特別是夜晚辦公。

**解決方案**：CSS 變數已有設計，加入 `prefers-color-scheme` 自動切換即可。

```css
@media (prefers-color-scheme: dark) {
    :root {
        --bg: #1e1e2e;
        --surface: #2a2a3e;
        --text: #e2e2f0;
        --border: #3a3a5c;
    }
}
```

**難度**：⭐⭐（中低）　**效益**：🔥（中）

---

#### B4. 品項手動新增優化：快速輸入模式

**問題**：目前手動新增品項需一格一格點選輸入，效率較低。

**解決方案**：支援貼入純文字或 Excel 複製的 Tab 分隔格式自動分割：

```
品名          數量  單價
KYOCERA 碳粉  2     3500
HP 墨水匣     5     280
```

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

#### B5. 用途說明智慧推薦

**問題**：「用途說明」欄位每次都要手動輸入，重複填寫相似內容。

**解決方案**：根據選擇的預算科目，自動帶入常用的用途說明範本，使用者可直接選用或修改。

```javascript
const PURPOSE_TEMPLATES = {
    '材料及用品費': [
        '購買辦公耗材（碳粉匣、墨水匣）供行政人員使用',
        '購買教學材料供課程使用',
    ],
    '設備及投資': [
        '購置資訊設備提升教學品質',
    ]
};
```

**難度**：⭐（容易）　**效益**：🔥🔥（高）

---

#### B6. Excel 填寫預覽面板

**問題**：下載後才能看到結果，若有錯誤需重新操作。

**解決方案**：在步驟 4 前加入「預覽」面板，用 HTML 表格模擬 Excel 外觀（不需真正開啟 Excel），讓使用者確認後再下載。

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

### 🟢 優先等級 C：進階功能，長期規劃

---

#### C1. 歷史記錄功能

**問題**：過去產出的動支單無法在系統中查詢。

**解決方案**：使用 `localStorage` 或 `IndexedDB` 儲存每次產出的記錄（品項清單、日期、金額合計），在網頁上顯示「最近 10 筆」歷史清單。

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### C2. 支援更多 Excel 範本（推廣至其他學校）

**問題**：目前只支援石門國小的動支及黏存單範本。

**解決方案**：
- 提供「範本管理」介面，讓使用者上傳自己的 Excel 範本
- 透過「欄位對應」設定介面，指定哪個儲存格對應什麼資料
- 儲存設定後即可自動套用

**難度**：⭐⭐⭐⭐（高）　**效益**：🔥🔥🔥（極高，可推廣給更多學校）

---

#### C3. LINE 機器人整合

**問題**：老師在外面拿到報價單，需要回到電腦才能使用系統。

**解決方案**：建立 LINE Bot，老師直接在 LINE 傳送 PDF 圖片，Bot 透過 OCR 解析後回傳填好的 Excel。

**技術**：LINE Messaging API（免費）+ Google Cloud Functions 或 Render.com（免費方案）+ Tesseract OCR

**難度**：⭐⭐⭐⭐⭐（高）　**效益**：🔥🔥🔥（極高）

---

#### C4. 自動填寫「附件」核取方塊

**問題**：Excel 表單中的「附件」欄位（發票、收據、驗收報告等核取方塊）目前需手動勾選。

**解決方案**：在步驟 3 加入附件類型的多選選項，自動填寫 Excel 中對應的核取方塊儲存格。

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

#### C5. 支援上傳 Excel / CSV 格式報價單

**問題**：部分廠商提供的是 Excel 或 CSV 格式報價單，無法用 PDF 解析。

**解決方案**：在上傳區加入 `.xlsx` / `.csv` 支援，用 JSZip 直接解析 Excel XML，或用 Papa Parse 解析 CSV。

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### C6. PWA（Progressive Web App）離線支援

**問題**：在網路不穩定的環境，無法正常使用網頁版。

**解決方案**：加入 Service Worker，將所有靜態資源（HTML、JS、CSS、範本 xlsx）快取至本機，支援完全離線使用。

```javascript
// service-worker.js
self.addEventListener('install', e => {
    e.waitUntil(caches.open('budget-v1.3.0').then(cache =>
        cache.addAll(['/', '/css/style.css', '/js/app.js',
                      '/template/template.xlsx'])
    ));
});
```

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

#### C7. 多使用者雲端版本

**問題**：目前每台電腦各自獨立，無法共用歷史記錄或設定。

**解決方案**：整合 Supabase（免費方案）作為後端：
- 使用者以 Google 帳號登入（OAuth）
- 歷史記錄、常用範本、預設設定同步至雲端
- 多位行政人員共用一套設定

**難度**：⭐⭐⭐⭐⭐（高）　**效益**：🔥🔥🔥（極高）

---

#### C8. 自動更新版本通知

**問題**：使用者不知道系統有新版本，可能使用舊版（有 bug）。

**解決方案**：網頁載入時比對 GitHub 上的版本號，若有新版則顯示更新提示。

```javascript
const latest = await fetch('https://api.github.com/repos/cagoooo/budget/releases/latest')
    .then(r => r.json()).then(j => j.tag_name);
if (latest !== CURRENT_VERSION) showUpdateBanner(latest);
```

**難度**：⭐（容易）　**效益**：🔥（中）

---

## 技術債（建議清理）

| 項目 | 說明 | 難度 |
|------|------|------|
| PDF 解析欄位邊界 | 目前依靠 header 行座標推算，對非標準報價單可能失準 | ⭐⭐⭐⭐ |
| fill_excel.py 欄位對應 | 硬編碼於程式中，建議改為 JSON 設定檔，方便其他學校自訂 | ⭐⭐ |
| 自動化測試 | 目前無自動化測試，建議加入 PDF 解析單元測試與 Excel 輸出驗證 | ⭐⭐⭐ |
| 錯誤追蹤 | 使用者端的 JS 錯誤無法收集，建議整合 Sentry（免費方案） | ⭐⭐ |

---

## 建議實作順序

```
短期（1-2 週）
├── A2. 多頁 PDF 支援              ← 容易，效益高
├── A3. 品項超過 8 筆 UI 警告      ← 極容易
├── A4. LocalStorage 記憶設定      ← 容易，每次用都省時
├── B2. 解析信心分數顏色標示        ← 容易，提升易用性
└── B5. 用途說明智慧推薦            ← 容易，省時

中期（1 個月）
├── A1. 掃描 PDF OCR 支援          ← 最常被需要
├── A5. 列印 / PDF 輸出            ← 減少操作步驟
├── B4. 快速貼入輸入模式            ← 提升效率
├── B6. Excel 填寫預覽面板          ← 減少錯誤
└── C4. 附件核取方塊自動填寫        ← 表單完整度更高

長期（3-6 個月）
├── C2. 多範本支援                  ← 可推廣其他學校
├── C6. PWA 離線支援                ← 提升可靠性
├── C5. Excel/CSV 報價單支援        ← 擴大相容性
└── C1. 歷史記錄功能                ← 查詢追蹤
```

---

*如有任何建議或想優先實作某項功能，歡迎聯繫阿凱老師！*

Made with 💝 by [阿凱老師](https://www.smes.tyc.edu.tw/modules/tadnews/page.php?ncsn=11&nsn=16#a5)
