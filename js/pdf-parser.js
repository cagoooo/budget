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

        // 找出表頭列（含「品 名」或「品名」）以及「以下空白」列
        let headerRowIdx = -1;
        let endRowIdx = textRows.length;

        for (let i = 0; i < textRows.length; i++) {
            const t = textRows[i];
            if (t.includes('品') && t.includes('名') && (t.includes('數量') || t.includes('單位') || t.includes('單價'))) {
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
        const headerRow = rows[headerRowIdx];
        const headerPositions = analyzeHeaderPositions(headerRow);

        // 解析資料列（支援多行品名合併）
        for (let i = headerRowIdx + 1; i < endRowIdx; i++) {
            const row = rows[i];
            const text = rowToText(row);

            // 跳過空白列
            if (!text.trim() || text === '以下空白') continue;

            // 解析品項
            const item = parseItemRow(row, headerPositions, text);
            if (item) {
                if (item.quantity !== 0 || item.unitPrice !== 0) {
                    // 有數量或單價的正常品項列
                    result.items.push(item);
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
            if (text === '數量') positions.qtyX = x;
            if (text === '單位') positions.unitX = x;
            if (text === '單價') positions.priceX = x;
            if (text === '小計') positions.subtotalX = x;
            if (text.includes('附註')) positions.noteX = x;
        }

        // 計算各欄位的邊界值（使用相鄰欄位的中點）
        const qtyX = positions.qtyX || 300;
        const unitX = positions.unitX || qtyX + 40;
        const priceX = positions.priceX || unitX + 40;
        const subtotalX = positions.subtotalX || priceX + 60;

        return {
            // 品名區域：貨品編號右側到數量左側
            nameStart: (positions.codeX || 40) + 30,
            nameEnd: qtyX - 5,
            // 數量區域
            qtyStart: qtyX - 15,
            qtyEnd: unitX - 5,
            // 單價區域
            priceStart: priceX - 30,
            priceEnd: subtotalX - 10,
            // 小計區域
            subtotalStart: subtotalX - 15,
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
                // 數量區域
                const parsed = parseFloat(text);
                if (!isNaN(parsed)) qty = parsed;
            } else if (x >= bounds.priceStart && x < bounds.priceEnd) {
                // 單價區域
                const parsed = parseFloat(text.replace(/,/g, ''));
                if (!isNaN(parsed)) price = parsed;
            } else if (x >= bounds.subtotalStart && x <= bounds.subtotalEnd) {
                // 小計區域
                const parsed = parseFloat(text.replace(/,/g, ''));
                if (!isNaN(parsed)) subtotal = parsed;
            } else if (x >= bounds.nameStart - 30 && x < bounds.nameEnd) {
                // 品名區域（較寬鬆的左邊界，因為品名可能從較左位置開始）
                if (!isSeqNumber(text) && !isProductCode(text)) {
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
     */
    async function parse(file) {
        const items = await extractTextFromPDF(file);
        const result = parseQuotation(items);

        // 如果座標解析沒有找到品項，嘗試文字模式
        if (result.items.length === 0) {
            const rows = groupIntoRows(items);
            const textRows = rows.map(r => rowToText(r));
            result.items = parseByText(textRows);
        }

        return result;
    }

    return { parse };
})();
