/**
 * PDF 報價單解析模組
 * 從廠商報價單 PDF 中擷取品名、數量、單價等欄位
 */

const PDFParser = (() => {
    // 設定 PDF.js worker
    if (typeof pdfjsLib !== 'undefined') {
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    /**
     * 從 PDF 檔案擷取全部文字內容（含座標資訊）
     */
    async function extractTextFromPDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const allItems = [];

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            allItems.push(...textContent.items);
        }

        return allItems;
    }

    /**
     * 將文字項目依 Y 座標分群為列
     */
    function groupIntoRows(items) {
        if (items.length === 0) return [];

        // 依 Y 座標排序（由上到下，Y 值由大到小）
        const sorted = [...items].sort((a, b) => b.transform[5] - a.transform[5]);

        const rows = [];
        let currentRow = [sorted[0]];
        let currentY = sorted[0].transform[5];

        for (let i = 1; i < sorted.length; i++) {
            const item = sorted[i];
            const y = item.transform[5];
            // 同一列的 Y 座標差距在 5 以內
            if (Math.abs(y - currentY) < 5) {
                currentRow.push(item);
            } else {
                // 列內依 X 座標排序
                currentRow.sort((a, b) => a.transform[4] - b.transform[4]);
                rows.push(currentRow);
                currentRow = [item];
                currentY = y;
            }
        }
        currentRow.sort((a, b) => a.transform[4] - b.transform[4]);
        rows.push(currentRow);

        return rows;
    }

    /**
     * 從列文字中合併取得完整字串
     */
    function rowToText(row) {
        return row.map(item => item.str.trim()).filter(s => s).join(' ');
    }

    /**
     * 解析報價單結構化資料
     */
    function parseQuotation(items) {
        const rows = groupIntoRows(items);
        const result = {
            vendorName: '',
            quoteDate: '',
            totalAmount: 0,
            items: []
        };

        // 取得全部文字列
        const textRows = rows.map(r => rowToText(r));

        // 擷取廠商名稱（通常在第一行或包含「公司」的行）
        for (const text of textRows) {
            if (text.match(/有限公司|股份有限公司|企業社|工程行/) && !text.includes('確認')) {
                const match = text.match(/([^\s]+(?:有限公司|股份有限公司|企業社|工程行))/);
                if (match) {
                    result.vendorName = match[1];
                    break;
                }
            }
        }

        // 擷取報價日期
        for (const text of textRows) {
            const dateMatch = text.match(/單據日期[:\s]*(\d{4}\/\d{2}\/\d{2})/);
            if (dateMatch) {
                result.quoteDate = dateMatch[1];
                break;
            }
        }

        // 擷取總金額
        for (const text of textRows) {
            const totalMatch = text.match(/總計金額[:\s]*([\d,]+)/);
            if (totalMatch) {
                result.totalAmount = parseInt(totalMatch[1].replace(/,/g, ''), 10);
                break;
            }
        }
        if (result.totalAmount === 0) {
            for (const text of textRows) {
                const totalMatch = text.match(/合計金額[:\s]*([\d,]+)/);
                if (totalMatch) {
                    result.totalAmount = parseInt(totalMatch[1].replace(/,/g, ''), 10);
                    break;
                }
            }
        }

        // 找出表頭列（含「品名」）以及「以下空白」列
        let headerRowIdx = -1;
        let endRowIdx = textRows.length;

        for (let i = 0; i < textRows.length; i++) {
            const t = textRows[i];
            // 同一行有「品名」且有數量/單價/售價/小計其中一個即可
            if (t.includes('品') && t.includes('名') &&
                (t.includes('數量') || t.includes('單位') || t.includes('單價') ||
                 t.includes('售價') || t.includes('小計'))) {
                headerRowIdx = i;
            }
            if (t.includes('以下空白')) {
                endRowIdx = i;
                break;
            }
        }

        if (headerRowIdx === -1) {
            // 嘗試尋找包含「序」的表頭
            for (let i = 0; i < textRows.length; i++) {
                if (textRows[i].match(/^序\s/) && textRows[i].includes('貨品編號')) {
                    headerRowIdx = i;
                    break;
                }
            }
        }

        if (headerRowIdx === -1) return result;

        // 分析表頭欄位位置（使用 X 座標）
        // 支援跨兩行的表頭（如「採購\n數量」分列）
        let headerItemsForAnalysis = [...rows[headerRowIdx]];
        let dataStartRow = headerRowIdx + 1;
        if (dataStartRow < endRowIdx) {
            const nextText = textRows[dataStartRow];
            // 若下一行含欄位關鍵字且沒有大數字（非資料列）→ 視為表頭延續列
            const hasColKeywords = nextText.includes('數量') || nextText.includes('定價') ||
                                   nextText.includes('交期') || nextText.includes('單位');
            const hasProductData = /\d{5,}/.test(nextText); // 5位以上數字 → 資料列
            if (hasColKeywords && !hasProductData) {
                headerItemsForAnalysis = [...headerItemsForAnalysis, ...rows[dataStartRow]];
                dataStartRow++;
            }
        }
        const headerPositions = analyzeHeaderPositions(headerItemsForAnalysis);

        // 解析資料列（支援多行品名合併）
        for (let i = dataStartRow; i < endRowIdx; i++) {
            const row = rows[i];
            const text = rowToText(row);

            // 跳過空白列
            if (!text.trim() || text === '以下空白') continue;

            // 解析品項
            const item = parseItemRow(row, headerPositions, text);
            if (item) {
                if (item.quantity !== 0 || item.unitPrice !== 0) {
                    // 有數量或單價的正常品項列（排除數量=0 的「不採購」品項）
                    if (item.quantity > 0 || item.unitPrice > 0) {
                        result.items.push(item);
                    }
                } else if (item.name && result.items.length > 0) {
                    // 沒有數量/單價但有名稱：視為上一品項的延續（多行品名）
                    result.items[result.items.length - 1].name += ' ' + item.name;
                } else if (item.name) {
                    result.items.push(item);
                }
            } else {
                // parseItemRow 回傳 null，但可能是品名延續列
                // 檢查此列是否只有文字落在品名區域
                const contName = extractContinuationName(row, headerPositions);
                if (contName && result.items.length > 0) {
                    result.items[result.items.length - 1].name += ' ' + contName;
                }
            }
        }

        return result;
    }

    /**
     * 分析表頭各欄位的 X 座標範圍
     * 使用欄位邊界（兩欄位中間值）來劃分區域
     */
    function analyzeHeaderPositions(headerRow) {
        const positions = {};

        for (const item of headerRow) {
            const text = item.str.trim();
            const x = item.transform[4];
            if (text === '序') positions.seqX = x;
            if (text.includes('貨品編號')) positions.codeX = x;
            if (text.includes('品') && text.includes('名')) positions.nameX = x;
            // 數量欄：支援「數量」「採購」「採購數量」等變體
            if (text === '數量' || text === '採購' || text === '採購數量') positions.qtyX = x;
            if (text === '單位') positions.unitX = x;
            // 單價欄：支援「單價」「售價」
            if (text === '單價' || text === '售價') positions.priceX = x;
            // 折扣欄：偵測後可精確限制單價右界
            if (text === '折扣' || text === '折讓') positions.discountX = x;
            if (text === '小計') positions.subtotalX = x;
            if (text.includes('附註')) positions.noteX = x;
            // 作者 / 出版社欄：偵測到後可縮小品名右界
            if (text === '作者') positions.authorX = x;
            if (text === '出版社') positions.publisherX = x;
        }

        const qtyX = positions.qtyX || 300;
        const unitX = positions.unitX || null;
        const priceX = positions.priceX || (unitX ? unitX + 40 : qtyX + 40);
        const discountX = positions.discountX || null;
        const subtotalX = positions.subtotalX || priceX + 60;

        // ---- 使用「相鄰欄位中點」作為邊界，避免固定偏移量造成範圍重疊 ----

        // 品名右界：使用品名欄與作者欄的中點（避免作者資料混入品名）
        const nameEndX = positions.authorX && positions.nameX
            ? (positions.nameX + positions.authorX) / 2
            : positions.authorX
                ? positions.authorX - 5
                : qtyX - 5;

        // 「單位」欄是否夾在「數量」與「單價」之間（qty < unit < price）
        const hasUnitBetween = unitX && unitX > qtyX && unitX < priceX;

        // 數量右界：若單位欄夾在中間，用 qty↔unit 中點；否則用 qty↔price 中點
        const qtyEndX = hasUnitBetween
            ? (qtyX + unitX) / 2
            : (qtyX + priceX) / 2;

        // 單價左界：若單位欄夾在中間，用 unit↔price 中點；否則用 qty↔price 中點
        const priceStartX = hasUnitBetween
            ? (unitX + priceX) / 2
            : (qtyX + priceX) / 2;

        // 單價右界：若有折扣欄則取 price↔discount 中點，否則取 price↔subtotal 中點
        const priceEndX = discountX
            ? (priceX + discountX) / 2
            : (priceX + subtotalX) / 2;

        // 小計左界：若有折扣欄則取 discount↔subtotal 中點
        const subtotalStartX = discountX
            ? (discountX + subtotalX) / 2
            : subtotalX - 15;

        return {
            // 品名區域
            nameStart: (positions.codeX || 40) + 30,
            nameEnd: nameEndX,
            // 數量區域：左界固定為 qtyX-20（避免受 unitX 位置影響）
            qtyStart: qtyX - 20,
            qtyEnd: qtyEndX,
            // 單價區域
            priceStart: priceStartX,
            priceEnd: priceEndX,
            // 小計區域
            subtotalStart: subtotalStartX,
            subtotalEnd: (positions.noteX || subtotalX + 80) - 5,
            // 序號/貨品編號區域
            codeEnd: (positions.codeX || 40) + 80
        };
    }

    /**
     * 判斷文字是否為貨品編號格式（如 EER-0000005, DISCOUNT-0000002 等）
     */
    function isProductCode(text) {
        return /^[A-Z]+-\d{4,}$/i.test(text);
    }

    /**
     * 判斷文字是否為序號（純數字 1-2 位）
     */
    function isSeqNumber(text) {
        return /^\d{1,2}$/.test(text);
    }

    /**
     * 判斷文字是否為 ISBN-13（978/979 開頭的 13 位數字）
     */
    function isISBN13(text) {
        return /^97[89]\d{10}$/.test(text.replace(/[-\s]/g, ''));
    }

    /**
     * 判斷文字是否為應忽略的數字型雜訊（如定價、ISBN、百分比）
     * 用於品名欄位過濾
     */
    function isNameNoise(text) {
        if (isISBN13(text)) return true;          // ISBN-13
        if (/^\d+%$/.test(text)) return true;     // 折扣百分比 (79%)
        if (/^\d{3,}$/.test(text)) return true;   // 3位以上純數字（定價、編號等）
        if (/^\d{1,3}(,\d{3})+$/.test(text)) return true; // 千分位數字 (1,188)
        return false;
    }

    /**
     * 解析單一品項列
     */
    function parseItemRow(row, bounds, fullText) {
        const nameTexts = [];
        let qty = 0;
        let price = 0;
        let subtotal = 0;

        for (const item of row) {
            const x = item.transform[4];
            const text = item.str.trim();
            if (!text) continue;

            // 根據 X 座標判斷此文字屬於哪個欄位
            if (x >= bounds.qtyStart && x < bounds.qtyEnd) {
                // 數量區域：過濾 ISBN、百分比、異常大數（如 ISBN 流入）
                if (!isISBN13(text) && !/^\d+%$/.test(text)) {
                    const parsed = parseFloat(text.replace(/,/g, ''));
                    if (!isNaN(parsed) && parsed < 100000) qty = parsed;
                }
            } else if (x >= bounds.priceStart && x < bounds.priceEnd) {
                // 單價區域：過濾百分比（折扣欄）
                if (!/^\d+%$/.test(text)) {
                    const parsed = parseFloat(text.replace(/,/g, ''));
                    if (!isNaN(parsed)) price = parsed;
                }
            } else if (x >= bounds.subtotalStart && x <= bounds.subtotalEnd) {
                // 小計區域
                const parsed = parseFloat(text.replace(/,/g, ''));
                if (!isNaN(parsed)) subtotal = parsed;
            } else if (x >= bounds.nameStart - 30 && x < bounds.nameEnd) {
                // 品名區域（較寬鬆的左邊界）
                // 過濾：序號、貨品編號、ISBN、定價數字、百分比等雜訊
                if (!isSeqNumber(text) && !isProductCode(text) && !isNameNoise(text)) {
                    nameTexts.push(text);
                }
            }
            // 其他位置的文字（序號、貨品編號、單位、附註）直接忽略
        }

        const name = nameTexts.join(' ').trim();
        if (!name && qty === 0 && price === 0) return null;
        if (!name) return null;

        if (subtotal === 0 && qty !== 0 && price !== 0) {
            subtotal = Math.round(qty * price);
        }

        return { name, quantity: qty, unitPrice: price, subtotal };
    }

    /**
     * 擷取延續行的品名文字（用於多行品名合併）
     */
    function extractContinuationName(row, bounds) {
        const texts = [];

        for (const item of row) {
            const x = item.transform[4];
            const text = item.str.trim();
            if (!text) continue;

            if (x >= bounds.nameStart - 30 && x < bounds.nameEnd) {
                if (!isSeqNumber(text) && !isProductCode(text)) {
                    texts.push(text);
                }
            }
        }

        return texts.join(' ').trim();
    }

    // ===== OCR 支援（掃描版 PDF）=====

    /**
     * 判斷是否為掃描版 PDF（文字稀少）
     */
    function isScannedPDF(textItems) {
        const totalChars = textItems.reduce((sum, item) => sum + (item.str || '').length, 0);
        return totalChars < 80;
    }

    /**
     * 動態載入 Tesseract.js CDN（按需載入）
     */
    function loadTesseract() {
        return new Promise((resolve, reject) => {
            if (typeof Tesseract !== 'undefined') { resolve(); return; }
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/tesseract.js@4/dist/tesseract.min.js';
            script.onload = resolve;
            script.onerror = () => reject(new Error('無法載入 OCR 引擎，請確認網路連線'));
            document.head.appendChild(script);
        });
    }

    /**
     * 將 PDF 頁面渲染成 canvas（scale=2 提高 OCR 精度）
     */
    async function renderPageToCanvas(page, scale = 2) {
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');
        await page.render({ canvasContext: ctx, viewport }).promise;
        return canvas;
    }

    /**
     * 使用 Tesseract.js OCR 辨識掃描 PDF 的全部頁面
     */
    async function extractTextViaOCR(file, onProgress) {
        await loadTesseract();
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const numPages = pdf.numPages;
        const texts = [];

        onProgress && onProgress('正在初始化 OCR 引擎（首次使用需下載語言包約 5MB）...');
        const worker = await Tesseract.createWorker('chi_tra+eng', 1, {
            logger: m => {
                if (!onProgress) return;
                if (m.status === 'loading tesseract core') onProgress('正在載入 OCR 核心...');
                else if (m.status === 'loading language traineddata') onProgress('正在下載中文語言包...');
                else if (m.status === 'recognizing text') {
                    // 不在這裡更新，由外層 for 迴圈顯示頁碼
                }
            }
        });

        try {
            for (let i = 1; i <= numPages; i++) {
                onProgress && onProgress(`正在 OCR 辨識第 ${i} / ${numPages} 頁，請稍候...`);
                const page = await pdf.getPage(i);
                const canvas = await renderPageToCanvas(page, 2);
                const { data: { text } } = await worker.recognize(canvas);
                texts.push(text);
            }
        } finally {
            await worker.terminate();
        }

        return texts.join('\n');
    }

    /**
     * 解析 OCR 輸出文字為品項（比 parseByText 更寬鬆）
     */
    function parseOcrText(ocrText) {
        const lines = ocrText.split('\n').map(l => l.trim()).filter(l => l.length > 1);

        // 先嘗試原有的 parseByText 邏輯
        const byText = parseByText(lines);
        if (byText.length > 0) return byText;

        // 更寬鬆的備用解析：每行找數字群組
        const items = [];
        let headerFound = false;

        for (const line of lines) {
            if (line.includes('以下空白')) break;

            if (!headerFound) {
                if (line.includes('品') && (line.includes('數量') || line.includes('單價') || line.includes('合計'))) {
                    headerFound = true;
                }
                continue;
            }

            // 找出行內所有數字（忽略逗號格式）
            const nums = [...line.matchAll(/([\d,]+(?:\.\d+)?)/g)]
                .map(m => parseFloat(m[1].replace(/,/g, '')))
                .filter(n => !isNaN(n) && n > 0);

            if (nums.length < 2) continue;

            // 移除數字和單位詞，剩下的視為品名
            let name = line
                .replace(/[\d,]+(?:\.\d+)?/g, '')
                .replace(/[台個套式支張組份箱包瓶罐片條]/g, '')
                .replace(/\s+/g, ' ')
                .trim();

            // 移除開頭的序號和雜字
            name = name.replace(/^[\s\d.\-|]+/, '').trim();

            if (name.length >= 2) {
                items.push({
                    name,
                    quantity: nums[0],
                    unitPrice: nums[1],
                    subtotal: Math.round(nums[0] * nums[1])
                });
            }
        }

        return items;
    }

    /**
     * 備用：純文字模式解析（當座標解析失敗時使用）
     */
    function parseByText(textRows) {
        const items = [];

        for (const text of textRows) {
            // 嘗試匹配：品名 數量 單位 單價 小計
            const match = text.match(
                /(.+?)\s+([\d.]+)\s+[台個套式支張組份箱包瓶罐片條]\s+([\d,.]+)\s+([\d,.-]+)/
            );
            if (match) {
                const name = match[1].replace(/^[\d]+\s+[A-Z]+-\d+\s+/i, '').trim();
                if (name && !name.includes('以下空白')) {
                    items.push({
                        name: name,
                        quantity: parseFloat(match[2]),
                        unitPrice: parseFloat(match[3].replace(/,/g, '')),
                        subtotal: parseFloat(match[4].replace(/,/g, ''))
                    });
                }
            }
        }

        return items;
    }

    /**
     * 主要入口：解析 PDF 檔案
     * @param {File} file - PDF 檔案
     * @param {Function} [onProgress] - 進度回呼 (message: string) => void
     */
    async function parse(file, onProgress) {
        const textItems = await extractTextFromPDF(file);

        // A1：偵測掃描版 PDF，啟用 OCR
        if (isScannedPDF(textItems)) {
            onProgress && onProgress('偵測到掃描版 PDF，正在啟動 OCR 辨識...');
            try {
                const ocrText = await extractTextViaOCR(file, onProgress);
                onProgress && onProgress('OCR 完成，正在解析品項...');
                const result = {
                    vendorName: '',
                    quoteDate: '',
                    totalAmount: 0,
                    items: parseOcrText(ocrText)
                };
                return result;
            } catch (err) {
                console.warn('OCR 失敗，改用純文字模式：', err);
                // 繼續使用一般解析（可能結果為空，讓使用者手動新增）
            }
        }

        const result = parseQuotation(textItems);

        // 座標解析失敗時改用純文字模式
        if (result.items.length === 0) {
            const rows = groupIntoRows(textItems);
            const textRows = rows.map(r => rowToText(r));
            result.items = parseByText(textRows);
        }

        return result;
    }

    return { parse };
})();
