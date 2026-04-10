/**
 * Excel 動支及黏存單產生模組
 *
 * 雙模式運作：
 *  1. 本地模式（run.py）：呼叫 /api/generate，由 Python openpyxl 填寫
 *     → 完整保留原始格式（欄寬、列高、框線、合併儲存格）
 *  2. 靜態模式（GitHub Pages）：使用 JSZip 直接修補工作表 XML
 *     → 完整保留原始格式（不重新生成 XML，只替換特定儲存格值）
 *
 * 公式欄位（兩種模式皆不修改）：
 *   代收代辦: P4, H5-O5, F13, G15/G17/.../G29, G31, L31
 *   預算內:   O4, G5-N5, F13, G15/G19/.../G29, G31, K31
 *
 * 品項列（row 15,17,19,21,23,25,27,29）：
 *   A欄=名稱及規格, C欄=數量, E欄=單價
 *   代收代辦用途說明=J15, 預算內用途說明=I15
 *   代收代辦月=J13 日=L13, 預算內月=I13 日=K13
 */

const ExcelGenerator = (() => {

    const ITEM_ROWS = [15, 17, 19, 21, 23, 25, 27, 29];

    // sheet 檔案對應（由 workbook.xml 確認）
    const SHEET_FILES = {
        '代收代辦': 'xl/worksheets/sheet1.xml',
        '預算內':   'xl/worksheets/sheet2.xml'
    };

    // 各 sheet 的目標儲存格設定
    const SHEET_CELLS = {
        '預算內': {
            purposeCell: 'I15',
            monthCell:   'I13',
            dayCell:     'K13'
        },
        '代收代辦': {
            purposeCell: 'J15',
            monthCell:   'J13',
            dayCell:     'L13'
        }
    };

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

    // ===== 靜態 JSZip XML 修補模式 =====

    let _templateCache = null;

    async function loadTemplate() {
        if (_templateCache) return _templateCache;
        const response = await fetch('template/template.xlsx');
        if (!response.ok) throw new Error('無法載入範本檔案');
        _templateCache = await response.arrayBuffer();
        return _templateCache;
    }

    /** XML 特殊字元跳脫 */
    function escXml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /**
     * 設定 inlineStr 型儲存格的文字值。
     * 支援：
     *  (a) 原本就是 inlineStr → 替換文字
     *  (b) 原本是空的 numeric (t="n") → 改為 inlineStr
     */
    function patchStr(xml, cellRef, value) {
        const esc = escXml(value);
        // (a) 已是 inlineStr
        const rxInline = new RegExp(
            `(<c r="${cellRef}"[^>]*t="inlineStr"[^>]*><is><t[^>]*>)[^<]*(</t></is></c>)`
        );
        if (rxInline.test(xml)) {
            return xml.replace(rxInline, `$1${esc}$2`);
        }
        // (b) 空的 numeric 儲存格 <c r="..." s="..." t="n"></c>
        const rxEmptyN = new RegExp(
            `(<c r="${cellRef}")([^>]*)t="n"([^>]*>)(</c>)`
        );
        if (rxEmptyN.test(xml)) {
            return xml.replace(
                rxEmptyN,
                `$1$2t="inlineStr"$3<is><t xml:space="preserve">${esc}</t></is></c>`
            );
        }
        // (c) numeric 有值的儲存格 <c ...><v>...</v></c>
        const rxNumV = new RegExp(
            `(<c r="${cellRef}")([^>]*)t="n"([^>]*>)<v>[^<]*</v>(</c>)`
        );
        return xml.replace(
            rxNumV,
            `$1$2t="inlineStr"$3<is><t xml:space="preserve">${esc}</t></is></c>`
        );
    }

    /**
     * 清除 inlineStr 或 numeric 儲存格的值（保留樣式與型別）。
     */
    function clearStr(xml, cellRef) {
        // inlineStr → 清空文字
        const rxInline = new RegExp(
            `(<c r="${cellRef}"[^>]*t="inlineStr"[^>]*><is><t[^>]*>)[^<]*(</t></is></c>)`
        );
        if (rxInline.test(xml)) {
            return xml.replace(rxInline, '$1$2');
        }
        // numeric 有 <v> → 移除 <v>
        return xml.replace(
            new RegExp(`(<c r="${cellRef}"[^>]*>)<v>[^<]*</v>(</c>)`),
            '$1$2'
        );
    }

    /**
     * 設定數值型儲存格的值。
     * 支援：有 <v> 的儲存格（替換）與空儲存格（插入）。
     */
    function patchNum(xml, cellRef, value) {
        // 有 <v> → 替換
        const rxHas = new RegExp(`(<c r="${cellRef}"[^>]*>)<v>[^<]*</v>(</c>)`);
        if (rxHas.test(xml)) {
            return xml.replace(rxHas, `$1<v>${value}</v>$2`);
        }
        // 空儲存格 → 插入 <v>
        return xml.replace(
            new RegExp(`(<c r="${cellRef}"[^>]*>)(</c>)`),
            `$1<v>${value}</v>$2`
        );
    }

    /** 清除數值型儲存格的值 */
    function clearNum(xml, cellRef) {
        return xml.replace(
            new RegExp(`(<c r="${cellRef}"[^>]*>)<v>[^<]*</v>(</c>)`),
            '$1$2'
        );
    }

    async function generateViaJSZip(params) {
        const templateData = await loadTemplate();
        const zip = await JSZip.loadAsync(templateData);

        const sheetName = params.templateType;
        const cells = SHEET_CELLS[sheetName];
        const sheetFile = SHEET_FILES[sheetName];
        const otherFile = sheetName === '預算內'
            ? SHEET_FILES['代收代辦']
            : SHEET_FILES['預算內'];

        if (!sheetFile) throw new Error(`未知的工作表類型：${sheetName}`);

        // ===== 修補目標工作表 =====
        let xml = await zip.file(sheetFile).async('string');

        // 清除所有品項列（避免殘留舊資料）
        for (const row of ITEM_ROWS) {
            xml = clearStr(xml, `A${row}`);
            xml = clearNum(xml,  `C${row}`);
            xml = clearNum(xml,  `E${row}`);
        }

        // 填入品項
        const items = params.items.slice(0, 8);
        for (let i = 0; i < items.length; i++) {
            const row = ITEM_ROWS[i];
            const item = items[i];
            if (item.name)      xml = patchStr(xml, `A${row}`, item.name);
            if (item.quantity)  xml = patchNum(xml,  `C${row}`, Number(item.quantity));
            if (item.unitPrice) xml = patchNum(xml,  `E${row}`, Number(item.unitPrice));
        }

        // 用途說明
        if (params.purpose) xml = patchStr(xml, cells.purposeCell, params.purpose);

        // 黏存單用途說明格：確保自動換行格式
        // 代收代辦(sheet1) P4 原樣式 s=147 無 wrapText → 改為 s=148（同屬性但有 wrapText）
        xml = xml.replace(/<c r="P4" s="147"/, '<c r="P4" s="148"');
        // 移除 row 4 固定高度限制，讓 Excel 依內容自動展開
        xml = xml.replace(/(<row r="4"[^>]*)\bcustomHeight="1"\b/, '$1');

        // 單位別（B13）
        if (params.unit) xml = patchStr(xml, 'B13', params.unit);

        // 月、日（年由公式 YEAR(TODAY())-1911 自動帶入）
        if (params.month) xml = patchNum(xml, cells.monthCell, Number(params.month));
        if (params.day)   xml = patchNum(xml, cells.dayCell,   Number(params.day));

        // 所屬年度（B2）
        if (params.year) xml = patchNum(xml, 'B2', Number(params.year));

        // 預算科目（B4, B5）
        if (params.budgetCategory)    xml = patchStr(xml, 'B4', params.budgetCategory);
        if (params.budgetSubCategory) xml = patchStr(xml, 'B5', params.budgetSubCategory);

        zip.file(sheetFile, xml);

        // ===== 清除另一張工作表的品項殘留 =====
        if (otherFile && zip.file(otherFile)) {
            let xmlOther = await zip.file(otherFile).async('string');
            for (const row of ITEM_ROWS) {
                xmlOther = clearStr(xmlOther, `A${row}`);
                xmlOther = clearNum(xmlOther,  `C${row}`);
                xmlOther = clearNum(xmlOther,  `E${row}`);
            }
            zip.file(otherFile, xmlOther);
        }

        // ===== 設定開啟時顯示正確工作表（workbook.xml） =====
        const activeIdx = sheetName === '預算內' ? 1 : 0;
        let wbXml = await zip.file('xl/workbook.xml').async('string');
        wbXml = wbXml.replace(
            /(<workbookView[^>]*)activeTab="\d+"([^>]*\/>)/,
            `$1activeTab="${activeIdx}"$2`
        );
        if (!wbXml.includes('activeTab=')) {
            wbXml = wbXml.replace(
                /(<workbookView)([^>]*\/>)/,
                `$1 activeTab="${activeIdx}"$2`
            );
        }
        zip.file('xl/workbook.xml', wbXml);

        const content = await zip.generateAsync({
            type: 'uint8array',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
        return content;
    }

    // ===== 主要入口 =====

    async function generate(params) {
        const pythonAvailable = await isPythonMode();
        if (pythonAvailable) {
            return await generateViaPython(params);
        } else {
            return await generateViaJSZip(params);
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
