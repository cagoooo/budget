# 後續優化改良建議

> 依優先順序排列，標示所需技術複雜度與預估效益。

---

## 一、高優先 — 功能穩健性

### 1.1 PDF 解析準確率提升

**現況**：使用座標式解析（X/Y 軸邊界分欄），對格式相近的廠商 PDF 準確率高，但少數廠商版面差異較大時可能解析失敗。

**改良方向**

- **關鍵字比對輔助**：偵測到「品名」「規格」「數量」「單價」等關鍵字時，動態調整欄位邊界
- **多廠商 Profile 記憶**：記錄曾成功解析的廠商版面資訊（localStorage），下次同廠商直接套用
- **OCR 備援**：整合 [Tesseract.js](https://github.com/naptha/tesseract.js) 處理掃描式 PDF（目前完全無法解析）

```
預估難度：中
預估效益：★★★★★
```

---

### 1.2 Excel 範本版本管理

**現況**：只支援單一範本（`template/template.xlsx`），若學校更換新版公文範本需手動替換檔案。

**改良方向**

- 支援多個範本版本（例如按年度命名：`template_114.xlsx`、`template_115.xlsx`）
- 在 UI 新增「範本版本」選單
- 提供「自訂範本上傳」功能（用戶可上傳自己的 xlsx 作為一次性範本）

```
預估難度：低
預估效益：★★★★☆
```

---

### 1.3 輸入資料驗證強化

**現況**：數量、單價欄位若輸入非數字不會即時提示，只在產生時才報錯。

**改良方向**

- 即時驗證（contenteditable 儲存格 `input` 事件）：非數字立即標紅
- 金額上限警告（超過一定金額顯示提示）
- 品名超過 Excel 儲存格寬度時自動截斷警告

```
預估難度：低
預估效益：★★★☆☆
```

---

## 二、中優先 — 使用體驗提升

### 2.1 歷史記錄功能（localStorage）

讓使用者不需每次重新選擇固定欄位。

**儲存項目**
- 最後使用的「單位別」
- 最後使用的「表單類型」（預算內 / 代收代辦）
- 最後使用的「預算科目」

**實作方式**

```javascript
// 寫入
localStorage.setItem('lastUnit', unitSelect.value);

// 讀取（頁面載入時）
const saved = localStorage.getItem('lastUnit');
if (saved) unitSelect.value = saved;
```

```
預估難度：低
預估效益：★★★★☆
```

---

### 2.2 批次處理多份 PDF

**現況**：每次只能處理一份 PDF。

**改良方向**

- 支援一次上傳多份 PDF（多個廠商報價）
- 各份 PDF 解析結果獨立顯示於分頁或手風琴
- 可分別為各份報價選擇不同的動支單參數，批次下載

```
預估難度：高
預估效益：★★★★☆
```

---

### 2.3 匯出 PDF 功能

除了 Excel，提供直接匯出 PDF 版本動支單的選項。

**實作方式**

- 方案 A：使用 [jsPDF](https://github.com/parallax/jsPDF) + [html2canvas] 截圖轉 PDF（快速但品質較差）
- 方案 B：在本地模式透過 LibreOffice 將填好的 xlsx 轉換為 PDF（品質最佳）
  ```python
  subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', output_xlsx])
  ```

```
預估難度：中（方案A） / 高（方案B）
預估效益：★★★☆☆
```

---

### 2.4 品項表格拖曳排序

讓使用者可以拖曳調整品項順序。

**推薦套件**：[SortableJS](https://sortablejs.github.io/Sortable/)（CDN 引入，輕量）

```html
<script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"></script>
```

```javascript
Sortable.create(document.getElementById('parsedBody'), { animation: 150 });
```

```
預估難度：低
預估效益：★★★☆☆
```

---

## 三、部署與維運

### 3.1 GitHub Actions 自動部署

**現況**：需手動 `git push` 後等待 GitHub Pages 部署。

**改良方向**：建立 `.github/workflows/deploy.yml`，在每次推送 `main` 分支時自動觸發部署並執行簡易 smoke test。

```yaml
# .github/workflows/deploy.yml
name: Deploy to GitHub Pages
on:
  push:
    branches: [main]
jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Deploy
        uses: peaceiris/actions-gh-pages@v4
        with:
          github_token: ${{ secrets.GITHUB_TOKEN }}
          publish_dir: ./
```

```
預估難度：低
預估效益：★★★★☆
```

---

### 3.2 正式版本號管理

在頁面顯示版本號，方便追蹤與回報問題。

**實作**：在 `package.json`（或 `version.txt`）記錄版本，頁面 footer 顯示 `v1.2.0`。

---

### 3.3 錯誤回報機制

讓使用者一鍵將解析錯誤回報至 GitHub Issues。

```javascript
const issueUrl = `https://github.com/你的帳號/budget-tool/issues/new?title=${encodeURIComponent('PDF解析問題')}&body=${encodeURIComponent(errorInfo)}`;
window.open(issueUrl, '_blank');
```

---

## 四、進階功能（長期規劃）

### 4.1 AI 輔助解析（Claude API）

對解析失敗或結果不確定的 PDF，可整合 Claude API 進行語意輔助辨識：

- 將 PDF 文字送至 Claude，請 Claude 以 JSON 格式回傳品項清單
- 適用於版面特殊、欄位無表頭的報價單

**注意**：需妥善處理 API Key 安全性，建議只在本地模式實作（在 `run.py` 伺服器端呼叫 API，避免 Key 外洩）。

```
預估難度：中
預估效益：★★★★★（對難解析 PDF 效果極佳）
```

---

### 4.2 資料庫記錄（SQLite）

在本地模式加入 SQLite 資料庫，記錄每次產出的動支單基本資訊（品項、金額、日期），方便日後查詢或產出統計報表。

```python
# 在 fill_excel.py 或 run.py 中加入
import sqlite3
conn = sqlite3.connect('records.db')
conn.execute('''INSERT INTO records (date, unit, total, items_json) VALUES (?,?,?,?)''',
             (date, unit, total, json.dumps(items)))
```

```
預估難度：中
預估效益：★★★☆☆
```

---

### 4.3 PWA（Progressive Web App）

讓使用者可以將本系統「安裝」到手機或桌面，像 App 一樣開啟。

**需要新增的檔案**

- `manifest.json`：App 名稱、圖示、顯示模式設定
- `sw.js`：Service Worker，快取靜態資源（離線可用）

```json
// manifest.json
{
  "name": "動支及黏存單自動產生系統",
  "short_name": "動支單",
  "icons": [{"src": "apple-touch-icon.png", "sizes": "180x180", "type": "image/png"}],
  "start_url": ".",
  "display": "standalone",
  "theme_color": "#2563eb",
  "background_color": "#eff6ff"
}
```

```
預估難度：低（基本 PWA）/ 中（離線快取）
預估效益：★★★☆☆
```

---

## 優先順序總結

| 建議項目 | 難度 | 效益 | 建議時程 |
|----------|:----:|:----:|----------|
| 1.1 PDF 解析準確率 | 中 | ★★★★★ | 第一批 |
| 2.1 歷史記錄（localStorage） | 低 | ★★★★☆ | 第一批 |
| 1.2 多範本支援 | 低 | ★★★★☆ | 第一批 |
| 3.1 GitHub Actions 部署 | 低 | ★★★★☆ | 第一批 |
| 2.4 品項拖曳排序 | 低 | ★★★☆☆ | 第二批 |
| 1.3 輸入驗證強化 | 低 | ★★★☆☆ | 第二批 |
| 2.3 匯出 PDF | 中 | ★★★☆☆ | 第二批 |
| 4.3 PWA 離線支援 | 低-中 | ★★★☆☆ | 第三批 |
| 2.2 批次多份 PDF | 高 | ★★★★☆ | 第三批 |
| 4.1 AI 輔助解析 | 中 | ★★★★★ | 長期規劃 |
| 4.2 SQLite 記錄 | 中 | ★★★☆☆ | 長期規劃 |

> ✅ **已完成**：靜態模式 ExcelJS 格式完整保留（v1.1.0）
> ✅ **已完成**：修正 Excel 開啟顯示錯誤工作表導致內容空白問題（v1.2.0）

---

*本文件由 Claude AI 輔助撰寫，供石門國小行政工具開發參考。*
