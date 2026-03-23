# 未來優化改良建議

> 動支及黏存單自動產生系統 — 開發路線圖
> Made with 💝 by [阿凱老師](https://www.smes.tyc.edu.tw/modules/tadnews/page.php?ncsn=11&nsn=16#a5)
> 最後更新：2026-03-23（v1.4.3）

---

## 目前已完成進度總覽

| 版本 | 功能 | 狀態 |
|------|------|------|
| v1.0.0 | PDF 拖曳上傳與解析 | ✅ 完成 |
| v1.0.0 | 座標式欄位偵測（品名/數量/單價） | ✅ 完成 |
| v1.0.0 | 多行品名合併 | ✅ 完成 |
| v1.0.0 | 可編輯品項表格（點選直接修改） | ✅ 完成 |
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
| v1.3.0 | JSZip 直接修補 XML（取代 ExcelJS，根除 XML 損壞） | ✅ 完成 |
| v1.3.0 | Footer 版權資訊（阿凱老師超連結） | ✅ 完成 |
| v1.4.0 | **A1** 掃描版 PDF OCR 支援（Tesseract.js，按需載入） | ✅ 完成 |
| v1.4.0 | **A2** 多頁 PDF 支援（全頁迴圈解析） | ✅ 完成 |
| v1.4.0 | **A3** 超過 8 筆品項 UI 警告（紅色警告框） | ✅ 完成 |
| v1.4.0 | **A4** LocalStorage 記憶單位別、表單類型、預算科目 | ✅ 完成 |
| v1.4.0 | 所屬年度（B2）自動填入當年民國年 | ✅ 完成 |
| v1.4.3 | PDF 欄位邊界精確修正（數量欄左界固定，避免單位欄干擾） | ✅ 完成 |

---

## 未來優化建議

以下依「重要性」與「實作難度」分級，供參考排程。

---

### 🔴 優先等級 A：實用性最高，建議優先實作

---

#### A1. 列印 / 直接輸出 PDF 功能

**問題**：目前只能下載 Excel，還需要再開 Excel 列印，多了一個步驟。

**解決方案一（極簡，推薦）**：直接在瀏覽器呼叫 `window.print()`，搭配列印專用 CSS，將步驟 4 的「匯出預覽」美化成接近表單的外觀後列印。

```css
@media print {
    header, .step-section:not(#step4), footer, .generate-actions { display: none; }
    .preview-card { box-shadow: none; border: 1px solid #ccc; }
    body { font-size: 12pt; }
}
```

**解決方案二（精確）**：整合 [jsPDF](https://github.com/parallax/jsPDF)，將已產出的 Excel 範本資料直接轉為 PDF，外觀與 Excel 完全一致。

**難度**：⭐（方案一）/ ⭐⭐⭐（方案二）　**效益**：🔥🔥（高）

---

#### A2. 解析結果信心分數 + 顏色標示

**問題**：解析完成後，使用者無法直觀判斷哪些品項解析正確、哪些有疑問需要人工確認。

**解決方案**：在品項表格每列左側加入色點或 icon，根據欄位完整度自動標示：

| 狀態 | 條件 | 顏色 |
|------|------|------|
| 完整 | 品名 + 數量 + 單價 均有值 | 🟢 綠色 |
| 需確認 | 缺少數量或單價其中一項 | 🟡 黃色 |
| 需補填 | 只有品名，數量/單價皆缺 | 🔴 紅色 |

```javascript
function getConfidenceLevel(item) {
    const hasName  = item.name.trim().length > 0;
    const hasQty   = item.quantity > 0;
    const hasPrice = item.unitPrice > 0;
    if (hasName && hasQty && hasPrice) return 'ok';
    if (hasName && (hasQty || hasPrice))  return 'warn';
    return 'error';
}
```

**難度**：⭐（容易）　**效益**：🔥🔥（高）

---

#### A3. 用途說明智慧推薦

**問題**：「用途說明」每次都要手動輸入，重複填寫相似內容耗時，且容易拼錯常用詞彙。

**解決方案**：根據已選的預算科目自動帶入常用範本，使用者點選即套用，可再手動修改。

```javascript
const PURPOSE_TEMPLATES = {
    '材料及用品費': [
        '購置辦公耗材（碳粉匣、墨水匣）供行政人員使用',
        '購置教學材料供授課使用',
        '購置清潔用品供環境整潔使用',
    ],
    '服務費用': [
        '委託廠商維修設備恢復正常運作',
        '支付網路電話費用維持行政通訊',
    ],
    '其他設備': [
        '購置資訊設備提升教學品質',
        '購置辦公設備改善行政效率',
    ],
    '無形資產': [
        '購置授權軟體供行政人員使用',
    ]
};
```

顯示方式：用途說明欄位下方出現可點選的灰色建議按鈕，點選後自動填入，不強制使用。

**難度**：⭐（容易）　**效益**：🔥🔥（高）

---

#### A4. 品項快速貼入（Tab 分隔格式）

**問題**：目前手動新增品項需一格一格點選，若從 Excel 或文字整理好的資料貼入，每次都要逐格重填。

**解決方案**：在品項表格上方加入「批次貼入」文字區，支援從 Excel 複製後直接貼入的 Tab 分隔格式：

```
品名及規格         數量  單價
Kingston SSD 240G  2     1500
HP 碳粉匣 CF217A   3     890
```

```javascript
// 監聽品項表格的 paste 事件
parsedBody.addEventListener('paste', (e) => {
    const text = e.clipboardData.getData('text/plain');
    const rows = text.trim().split('\n').map(r => r.split('\t'));
    // 略過標題行，逐列解析並新增到表格
    rows.forEach(cols => addItemRow({
        name: cols[0] || '',
        quantity: parseFloat(cols[1]) || 0,
        unitPrice: parseFloat(cols[2]) || 0
    }));
});
```

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

### 🟡 優先等級 B：提升體驗，建議中期實作

---

#### B1. 多份報價單批次處理

**問題**：每次只能處理一份 PDF。若有多家廠商比價，需逐一上傳、逐一下載，重複操作耗時。

**解決方案**：支援一次選取多個 PDF，系統自動依序解析並產出多份 Excel，最後打包成 ZIP 下載。

```javascript
// HTML
<input type="file" accept=".pdf" multiple id="pdfInput">

// JS：一次處理多個 PDF
const files = Array.from(e.target.files);
const zip = new JSZip();
for (const file of files) {
    const data = await PDFParser.parse(file);
    const xlsx = await ExcelGenerator.generate({ items: data.items, ...formValues });
    zip.file(`動支及黏存單_${file.name.replace('.pdf','')}.xlsx`, xlsx);
}
const blob = await zip.generateAsync({ type: 'blob' });
downloadBlob(blob, '動支及黏存單_批次.zip');
```

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### B2. Excel 填寫結果預覽面板

**問題**：下載 Excel 後才能看到結果，若有欄位填錯需重新操作整個流程。

**解決方案**：步驟 4 加入 HTML 表格「模擬預覽」，完整呈現動支單的填寫結果（含欄位名稱、品項、日期、科目等），確認無誤再下載。

| 欄位 | 預覽內容 |
|------|------|
| 所屬年度 | 115 |
| 單位別 | 總務處 |
| 用途說明 | 購置辦公耗材 |
| 品項 1 | Kingston SSD 240G × 2 台 = NT$3,000 |
| 合計金額 | NT$3,000 |

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### B3. 歷史記錄功能

**問題**：過去產出的動支單無法在系統中查詢，須另行保存或翻找下載資料夾。

**解決方案**：用 `IndexedDB`（支援較大儲存量）記錄每次產出的摘要：日期、金額、品項數量、用途說明，在頁面側欄或獨立頁面顯示「最近 20 筆」。

```javascript
const record = {
    id: Date.now(),
    date: `${year}/${month}/${day}`,
    type: templateType,
    unit: unitSelect,
    purpose: purpose,
    itemCount: items.length,
    total: items.reduce((s, i) => s + i.subtotal, 0)
};
await db.records.add(record);
```

點選歷史記錄可快速還原上次的「表單設定（科目/單位/日期）」，節省重複填寫時間。

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### B4. 附件核取方塊自動填寫

**問題**：動支單 Excel 中的「附件」欄位（發票、統一收據、驗收報告等核取方塊）目前需手動在 Excel 中勾選。

**解決方案**：在步驟 3 加入多選核取方塊，系統自動在對應的 Excel 儲存格填入「✓」或「V」：

```
□ 統一發票  □ 收據  □ 廠商報價單  □ 驗收報告  □ 其他：＿＿＿
```

需先確認 Excel 範本中附件核取方塊對應的儲存格位置（需配合 XML 分析）。

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

#### B5. 深色模式（Dark Mode）

**問題**：長時間使用白色介面對眼睛較疲勞，特別是夜晚辦公或使用深色系電腦。

**解決方案**：CSS 變數已設計完整，新增 `prefers-color-scheme` 媒體查詢即可自動切換，不需 JS。

```css
@media (prefers-color-scheme: dark) {
    :root {
        --gray-50:  #111827;
        --gray-100: #1f2937;
        --gray-800: #f3f4f6;
        --gray-900: #f9fafb;
    }
    .step-section, .form-card, .preview-card, .parsed-info-card {
        background: #1f2937;
        border-color: #374151;
    }
}
```

也可加入手動切換按鈕（🌙/☀️），方便不使用深色系統設定的使用者切換。

**難度**：⭐⭐（中低）　**效益**：🔥（中）

---

#### B6. OCR 結果手動框選修正

**問題**：掃描版 PDF 的 OCR 結果可能部分辨識錯誤，目前只能在表格中逐格手動修改，無法看到 PDF 原始影像對照確認。

**解決方案**：OCR 完成後，在步驟 2 左側顯示 PDF 頁面縮圖，右側顯示解析結果表格，使用者可直接對照修改。當點擊某品項列時，左側 PDF 縮圖高亮對應區域。

**難度**：⭐⭐⭐⭐（高）　**效益**：🔥🔥（高，搭配 A1 使用）

---

### 🟢 優先等級 C：進階功能，長期規劃

---

#### C1. 支援上傳 Excel / CSV 格式報價單

**問題**：部分廠商提供的是 Excel 或 CSV 格式報價單，目前無法解析。

**解決方案**：
- **CSV**：用 [Papa Parse](https://www.papaparse.com/)（免費）解析 CSV，再送入現有欄位對應邏輯
- **Excel**：用現有的 JSZip 解析 `.xlsx` ZIP 結構，讀取 `sheet1.xml` 內的儲存格值

```javascript
// CSV 支援
if (file.name.endsWith('.csv')) {
    const text = await file.text();
    const result = Papa.parse(text, { header: true, skipEmptyLines: true });
    return mapCsvToItems(result.data);
}
```

**難度**：⭐⭐⭐（中等）　**效益**：🔥🔥（高）

---

#### C2. 支援更多 Excel 範本（推廣至其他學校）

**問題**：目前只支援石門國小的動支及黏存單。其他學校若有不同格式的表單，無法直接使用。

**解決方案**：提供「範本管理」功能：
1. 使用者上傳自己學校的 Excel 範本
2. 透過視覺化「欄位對應」介面點選儲存格，指定哪格填品名、哪格填數量
3. 設定儲存至 `localStorage`，之後自動套用

這是推廣給全台學校的關鍵功能，效益極高。

**難度**：⭐⭐⭐⭐（高）　**效益**：🔥🔥🔥（極高）

---

#### C3. PWA（Progressive Web App）離線支援

**問題**：在網路不穩定（如偏遠地區學校）或出差時，無法使用 GitHub Pages 版本。

**解決方案**：加入 Service Worker，將所有靜態資源快取至本機，支援完全離線使用。

```javascript
// service-worker.js
const CACHE = 'budget-v1.4.0';
const ASSETS = ['/', '/css/style.css', '/js/app.js',
    '/js/pdf-parser.js', '/js/excel-generator.js', '/template/template.xlsx'];

self.addEventListener('install', e =>
    e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)))
);
self.addEventListener('fetch', e =>
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)))
);
```

安裝後可從桌面直接開啟，體驗與 App 一致，無需瀏覽器網址列。

**難度**：⭐⭐（中低）　**效益**：🔥🔥（高）

---

#### C4. LINE 機器人整合

**問題**：老師在外面取得報價單時，需要回到電腦才能使用系統，行動端操作不便。

**解決方案**：建立 LINE Bot，老師直接在 LINE 傳送 PDF 或拍照的報價單影像，Bot 自動回傳填寫完成的 Excel 附件。

**技術架構**：
```
老師 LINE 傳圖/PDF
  ↓ LINE Messaging API Webhook
  ↓ Cloud Functions（Google / Render.com 免費方案）
  ↓ Tesseract OCR 辨識
  ↓ 套用 Excel 範本
  ↓ 回傳 Excel 附件給老師
```

**難度**：⭐⭐⭐⭐⭐（高）　**效益**：🔥🔥🔥（極高）

---

#### C5. 自動更新版本通知

**問題**：使用者不知道系統有新版本，可能長期使用有 bug 的舊版。

**解決方案**：網頁載入時比對 GitHub Releases API 版本號，若有新版顯示更新提示 banner。

```javascript
async function checkUpdate() {
    const CURRENT = 'v1.4.0';
    const resp = await fetch(
        'https://api.github.com/repos/cagoooo/budget/releases/latest'
    ).catch(() => null);
    if (!resp?.ok) return;
    const { tag_name } = await resp.json();
    if (tag_name && tag_name !== CURRENT) {
        showToast(`🎉 新版本 ${tag_name} 已發布，請按 Ctrl+Shift+R 更新`, 'info');
    }
}
```

**難度**：⭐（容易）　**效益**：🔥（中）

---

#### C6. 多使用者雲端版本

**問題**：目前每台電腦各自獨立，歷史記錄、預設設定無法在多台電腦間共用。

**解決方案**：整合 [Supabase](https://supabase.com/)（免費方案）作為後端：
- Google 帳號 OAuth 登入（不需自建帳號系統）
- 歷史記錄、常用科目、預設設定同步至雲端
- 全校行政人員共用同一套設定（如預算科目清單集中維護）

**難度**：⭐⭐⭐⭐⭐（高）　**效益**：🔥🔥🔥（極高）

---

## 技術債（建議清理）

| 項目 | 說明 | 難度 |
|------|------|------|
| PDF 解析欄位邊界 | 依靠 header 行 X 座標推算，對非標準格式報價單可能失準 | ⭐⭐⭐⭐ |
| OCR 文字解析器 | 目前 `parseOcrText` 為通用規則，針對特定廠商格式準確率有限 | ⭐⭐⭐ |
| fill_excel.py 欄位對應 | 硬編碼於程式中，建議改為 JSON 設定檔，方便其他學校自訂 | ⭐⭐ |
| 自動化測試 | 目前無自動化測試，建議加入 PDF 解析單元測試與 Excel 輸出驗證 | ⭐⭐⭐ |
| 錯誤追蹤 | 使用者端 JS 錯誤無法收集，建議整合 Sentry（免費方案） | ⭐⭐ |
| Tesseract 語言包快取 | 每次首次使用掃描 PDF 需重新下載語言包，可考慮 Cache API 持久化 | ⭐⭐ |

---

## 建議實作順序

```
短期（1-2 週）
├── A1. 列印 / PDF 輸出（window.print 方案）   ← 極容易，高頻需求
├── A2. 解析信心分數顏色標示                   ← 容易，提升易用性
├── A3. 用途說明智慧推薦                       ← 容易，每次省時
├── A4. 品項快速貼入（Tab 分隔）               ← 中低難度，大幅提升效率
└── C5. 自動更新版本通知                       ← 極容易，解決使用舊版問題

中期（1 個月）
├── B2. Excel 填寫結果預覽面板                 ← 減少填錯需重來的情況
├── B3. 歷史記錄功能（IndexedDB）             ← 方便查詢追蹤
├── B4. 附件核取方塊自動填寫                   ← 表單完整度更高
├── B6. OCR 結果對照 PDF 縮圖                  ← 搭配 A1 OCR 使用
└── C3. PWA 離線支援                           ← 提升可靠性

長期（3-6 個月）
├── C1. Excel/CSV 報價單支援                   ← 擴大相容性
├── C2. 多範本支援（推廣其他學校）             ← 最大推廣效益
├── C4. LINE 機器人整合                        ← 行動端突破
└── C6. 多使用者雲端版本                       ← 全校協作
```

---

*如有任何建議或想優先實作某項功能，歡迎聯繫阿凱老師！*

Made with 💝 by [阿凱老師](https://www.smes.tyc.edu.tw/modules/tadnews/page.php?ncsn=11&nsn=16#a5)
