/**
 * Excel 動支及黏存單產生模組
 *
 * 雙模式運作：
 *  1. 本地模式（run.py）：呼叫 /api/generate，由 Python openpyxl 填寫
 *     → 完整保留原始格式（欄寬、列高、框線、合併儲存格）
 *  2. 靜態模式（GitHub Pages）：使用 SheetJS 在瀏覽器端填寫
 *     → 資料正確，但部分格式可能與原始不同
 *
 * 公式欄位（兩種模式皆不修改）：
 *   代收代辦: P4, H5-O5, F13, G15/G17/.../G29, G31, L31
 *   預算內:   O4, G5-N5, F13, G15/G19/.../G29, G31, K31
 *
 * 品項列（row 15,17,19,21,23,25,27,29）：
 *   A欄=名稱及規格, C欄=數量, E欄=單價
 *   代收代辦用途說明=J15, 預算內用途說明=I15
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

    // ===== 靜態 SheetJS 模式 =====

    async function loadTemplate() {
        const response = await fetch('template/template.xlsx');
        if (!response.ok) throw new Error('無法載入範本檔案');
        return await response.arrayBuffer();
    }

    function cellRef(col, row) {
        return XLSX.utils.encode_cell({ c: col, r: row - 1 });
    }

    function setCellValue(ws, ref, value, type) {
        if (!ws[ref]) ws[ref] = {};
        ws[ref].v = value;
        ws[ref].t = type || (typeof value === 'number' ? 'n' : 's');
        delete ws[ref].w;
    }

    function clearCell(ws, ref) {
        if (ws[ref] && !ws[ref].f) delete ws[ref];
    }

    async function generateViaSheetJS(params) {
        const templateData = await loadTemplate();
        const wb = XLSX.read(templateData, {
            type: 'array',
            cellFormula: true,
            cellStyles: true,
            bookVBA: true
        });

        const sheetName = params.templateType;
        const ws = wb.Sheets[sheetName];
        if (!ws) throw new Error(`找不到工作表「${sheetName}」`);

        // 確保 range 足夠
        const range = XLSX.utils.decode_range(ws['!ref']);
        if (range.e.r < 34) range.e.r = 34;
        if (range.e.c < 22) range.e.c = 22;
        ws['!ref'] = XLSX.utils.encode_range(range);

        const items = params.items.slice(0, 8);

        // 清除舊品項
        for (const row of ITEM_ROWS) {
            clearCell(ws, cellRef(0, row));
            clearCell(ws, cellRef(2, row));
            clearCell(ws, cellRef(4, row));
        }

        // 填入品項
        for (let i = 0; i < items.length; i++) {
            const row = ITEM_ROWS[i];
            const item = items[i];
            if (item.name) setCellValue(ws, cellRef(0, row), item.name, 's');
            if (item.quantity) setCellValue(ws, cellRef(2, row), item.quantity, 'n');
            if (item.unitPrice) setCellValue(ws, cellRef(4, row), item.unitPrice, 'n');
        }

        // 用途說明
        if (params.purpose) {
            const purposeCol = sheetName === '代收代辦' ? 9 : 8;
            setCellValue(ws, cellRef(purposeCol, 15), params.purpose, 's');
        }

        // 單位別 B13
        if (params.unit) setCellValue(ws, cellRef(1, 13), params.unit, 's');

        // 日期（月、日）
        if (sheetName === '代收代辦') {
            if (params.month) setCellValue(ws, cellRef(9, 13), params.month, 'n');
            if (params.day) setCellValue(ws, cellRef(11, 13), params.day, 'n');
        } else {
            if (params.month) setCellValue(ws, cellRef(8, 13), params.month, 'n');
            if (params.day) setCellValue(ws, cellRef(10, 13), params.day, 'n');
        }

        // 預算科目
        if (params.budgetCategory) setCellValue(ws, cellRef(1, 4), params.budgetCategory, 's');
        if (params.budgetSubCategory) setCellValue(ws, cellRef(1, 5), params.budgetSubCategory, 's');

        return XLSX.write(wb, {
            bookType: 'xlsx',
            type: 'array',
            cellFormula: true,
            cellStyles: true
        });
    }

    // ===== 主要入口 =====

    async function generate(params) {
        const pythonAvailable = await isPythonMode();

        if (pythonAvailable) {
            return await generateViaPython(params);
        } else {
            return await generateViaSheetJS(params);
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
