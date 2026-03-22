/**
 * Excel 動支及黏存單產生模組
 *
 * 雙模式運作：
 *  1. 本地模式（run.py）：呼叫 /api/generate，由 Python openpyxl 填寫
 *     → 完整保留原始格式（欄寬、列高、框線、合併儲存格）
 *  2. 靜態模式（GitHub Pages）：使用 ExcelJS 在瀏覽器端填寫
 *     → 完整保留原始格式（ExcelJS 支援樣式讀寫）
 *
 * 公式欄位（兩種模式皆不修改）：
 *   代收代辦: P4, H5-O5, F13, G15/G17/.../G29, G31, L31
 *   預算內:   O4, G5-N5, F13, G15/G19/.../G29, G31, K31
 *
 * 品項列（row 15,17,19,21,23,25,27,29）：
 *   A欄=名稱及規格, C欄=數量, E欄=單價
 *   代收代辦用途說明=J15(col10), 預算內用途說明=I15(col9)
 */

const ExcelGenerator = (() => {

    const ITEM_ROWS = [15, 17, 19, 21, 23, 25, 27, 29];

    // ===== 模式偵測 =====

    let _modeCache = null;

    async function isPythonMode() {
        if (_modeCache !== null) return _modeCache;
        try {
            const controller = new AbortController();
            const timer = setTimeout(() => controller.abort(), 3000);
            const resp = await fetch('/api/ping', {
                method: 'GET',
                signal: controller.signal
            });
            clearTimeout(timer);
            _modeCache = resp.ok;
            return _modeCache;
        } catch {
            _modeCache = false;
            return false;
        }
    }

    // ===== 本地 Python 模式 =====

    async function generateViaPython(params) {
        const resp = await fetch('/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(params)
        });

        if (!resp.ok) {
            const msg = await resp.text().catch(() => '');
            throw new Error(`伺服器錯誤 ${resp.status}：${msg}`);
        }

        const buffer = await resp.arrayBuffer();
        return new Uint8Array(buffer);
    }

    // ===== 靜態 ExcelJS 模式 =====

    async function loadTemplate() {
        const response = await fetch('template/template.xlsx');
        if (!response.ok) throw new Error('無法載入範本檔案');
        return await response.arrayBuffer();
    }

    /**
     * 安全設定儲存格值：
     * - 公式儲存格：不覆蓋（保留原始公式）
     * - value=null 時：清除資料（不清除公式儲存格）
     */
    function setVal(ws, row, col, value) {
        const cell = ws.getCell(row, col);
        const isFormula = cell.value && typeof cell.value === 'object' && cell.value.formula;
        if (isFormula) return; // 公式儲存格永不覆蓋
        cell.value = value;
    }

    async function generateViaExcelJS(params) {
        const templateData = await loadTemplate();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(templateData);

        const sheetName = params.templateType;
        const ws = workbook.getWorksheet(sheetName);
        if (!ws) throw new Error(`找不到工作表「${sheetName}」`);

        // 設定開啟時顯示正確的工作表（0-indexed）
        const activeIdx = workbook.worksheets.findIndex(s => s.name === sheetName);
        workbook.views = [{ activeTab: activeIdx }];

        // 清除另一張 sheet 的舊資料，避免混淆
        const otherName = sheetName === '預算內' ? '代收代辦' : '預算內';
        const wsOther = workbook.getWorksheet(otherName);
        if (wsOther) {
            for (const row of ITEM_ROWS) {
                wsOther.getCell(row, 1).value = null;
                wsOther.getCell(row, 3).value = null;
                wsOther.getCell(row, 5).value = null;
            }
        }

        const items = params.items.slice(0, 8);

        // 清除品項舊資料
        for (const row of ITEM_ROWS) {
            setVal(ws, row, 1, null);  // A = 名稱及規格
            setVal(ws, row, 3, null);  // C = 數量
            setVal(ws, row, 5, null);  // E = 單價
        }

        // 填入品項
        for (let i = 0; i < items.length; i++) {
            const row = ITEM_ROWS[i];
            const item = items[i];
            if (item.name)      setVal(ws, row, 1, item.name);
            if (item.quantity)  setVal(ws, row, 3, Number(item.quantity));
            if (item.unitPrice) setVal(ws, row, 5, Number(item.unitPrice));
        }

        // 用途說明（第 15 列）
        if (params.purpose) {
            const purposeCol = sheetName === '代收代辦' ? 10 : 9; // J=10, I=9
            setVal(ws, 15, purposeCol, params.purpose);
        }

        // 單位別（B13）
        if (params.unit) setVal(ws, 13, 2, params.unit);

        // 月、日（年由公式自動帶入）
        if (sheetName === '代收代辦') {
            if (params.month) setVal(ws, 13, 10, Number(params.month)); // J13
            if (params.day)   setVal(ws, 13, 12, Number(params.day));   // L13
        } else {
            if (params.month) setVal(ws, 13, 9,  Number(params.month)); // I13
            if (params.day)   setVal(ws, 13, 11, Number(params.day));   // K13
        }

        // 預算科目（B4, B5）
        if (params.budgetCategory)    setVal(ws, 4, 2, params.budgetCategory);
        if (params.budgetSubCategory) setVal(ws, 5, 2, params.budgetSubCategory);

        const buffer = await workbook.xlsx.writeBuffer();
        return new Uint8Array(buffer);
    }

    // ===== 主要入口 =====

    async function generate(params) {
        const pythonAvailable = await isPythonMode();

        if (pythonAvailable) {
            return await generateViaPython(params);
        } else {
            return await generateViaExcelJS(params);
        }
    }

    function downloadFile(data, filename) {
        const blob = new Blob([data], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    async function getMode() {
        return (await isPythonMode()) ? 'python' : 'static';
    }

    return { generate, downloadFile, getMode };
})();
