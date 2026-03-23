/**
 * 動支及黏存單自動產生系統 - 主程式
 */

document.addEventListener('DOMContentLoaded', () => {
    // ===== DOM 元素 =====
    const uploadArea = document.getElementById('uploadArea');
    const pdfInput = document.getElementById('pdfInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const removeFileBtn = document.getElementById('removeFile');
    const step2 = document.getElementById('step2');
    const step3 = document.getElementById('step3');
    const step4 = document.getElementById('step4');
    const parsedBody = document.getElementById('parsedBody');
    const addItemBtn = document.getElementById('addItemBtn');
    const generateBtn = document.getElementById('generateBtn');
    const loadingOverlay = document.getElementById('loadingOverlay');
    const loadingText = document.getElementById('loadingText');
    const templateType = document.getElementById('templateType');
    const purposeInput = document.getElementById('purposeInput');

    // 預算科目選單資料（從 Excel 擷取）
    const budgetCategories = {
        '預算內': {
            '用人費用': ['113職員薪金','114工員工資','121聘用人員薪金','122約僱職員薪金','124兼職人員酬金','131加班費','151考績獎金','152年終獎金','15Y 其他獎金','161職員退休及離職金','162工員退休及離職金','181分擔員工保險費','183傷病醫藥費','18Y其他福利費'],
            '服務費用': ['212工作場所電費','214工作場所水費','221郵費','222電話費','224數據通信費','231國內旅費','235貨物運費','241印刷及裝訂費','252一般房屋修護費','254其他建築修護費','255機械及設備修護費','257什項設備修護費','26Y其他保險費','276佣金、匯費、經理費及手續費','279外包費','27D計時與計件人員酬金','27F體育活動費','285講課鐘點稿費出席審查及查詢費','287委託檢驗(定)試驗認證費','288委託考選訓練費','28A電子計算機軟體服務費','291公共關係費'],
            '材料及用品費': ['315設備零件','321辦公(事務)用品','322報章什誌','323農業與園藝用品及環境美化費','328醫療用品','32Y其他'],
            '租金、償債與利息': ['451機器租賃'],
            '稅捐及規費': ['612一般土地地價稅'],
            '會費、捐助、補助、分攤、照護、救濟與交流活動費': ['713職業團體會費','726獎助學員生給與','73Y分擔其他費用','743獎勵費用','751技能競賽'],
            '其他': ['91Y其他'],
            '營建及修建工程': ['513擴充改良房屋建築及設備'],
            '交通及運輸設備': ['515購置交通及運輸設備'],
            '其他設備': ['514購置機械及設備','516購置什項設備'],
            '無形資產': ['521購置電腦軟體']
        },
        '代收代辦': {
            '應付代收款': [],
            '押標金': [],
            '履約保證金': [],
            '保固金': [],
            '暫收及待結轉帳項': [],
            '應付保證品': []
        }
    };

    // 狀態
    let parsedData = null;
    let currentFile = null;

    // ===== 初始化 =====
    initDate();
    initLocalStorage();   // A4：含 updateCategorySelects()
    initModeDetection();

    // ===== 事件綁定 =====

    // 拖放上傳
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0 && files[0].type === 'application/pdf') {
            handleFile(files[0]);
        } else {
            showToast('請上傳 PDF 格式的檔案', 'error');
        }
    });

    // 點擊上傳區域
    uploadArea.addEventListener('click', (e) => {
        if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
            pdfInput.click();
        }
    });

    pdfInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // 移除檔案
    removeFileBtn.addEventListener('click', () => {
        resetAll();
    });

    // 新增品項
    addItemBtn.addEventListener('click', () => {
        const currentCount = parsedBody.querySelectorAll('tr').length;
        if (currentCount >= 8) {
            showToast('最多只能填入 8 筆品項', 'error');
            return;
        }
        addItemRow({ name: '', quantity: 0, unitPrice: 0, subtotal: 0 }, currentCount + 1);
        updatePreview();
    });

    // 表單類型切換
    templateType.addEventListener('change', () => {
        updateCategorySelects();
        updatePreview();
        saveSettings();
    });

    // 產生 Excel
    generateBtn.addEventListener('click', () => {
        generateExcel();
    });

    // 監聽表單變更以更新預覽
    document.getElementById('unitSelect').addEventListener('change', () => { updatePreview(); saveSettings(); });
    purposeInput.addEventListener('input', updatePreview);
    document.getElementById('dateYear').addEventListener('input', updatePreview);
    document.getElementById('dateMonth').addEventListener('input', updatePreview);
    document.getElementById('dateDay').addEventListener('input', updatePreview);
    document.getElementById('budgetCategory').addEventListener('change', () => {
        updateSubCategory('預算內');
        updatePreview();
        saveSettings();
    });
    document.getElementById('budgetSubCategory').addEventListener('change', () => { updatePreview(); saveSettings(); });
    document.getElementById('agencyCategory').addEventListener('change', () => {
        updateSubCategory('代收代辦');
        updatePreview();
        saveSettings();
    });
    document.getElementById('agencySubCategory').addEventListener('change', () => { updatePreview(); saveSettings(); });

    // ===== 核心函式 =====

    async function handleFile(file) {
        currentFile = file;
        fileName.textContent = file.name;
        fileInfo.style.display = 'block';
        uploadArea.style.display = 'none';

        showLoading('正在解析 PDF 報價單...');

        try {
            parsedData = await PDFParser.parse(file, (msg) => showLoading(msg));
            hideLoading();

            const count = parsedData.items.length;
            if (count === 0) {
                showToast('未能從 PDF 中解析出品項資料，請手動新增', 'error');
            } else {
                showToast(`成功解析 ${count} 筆品項`, 'success');
            }

            // A3：超過 8 筆品項警告
            const warningBox = document.getElementById('overLimitWarning');
            if (count > 8) {
                document.getElementById('totalItemCount').textContent = count;
                warningBox.style.display = 'flex';
            } else {
                warningBox.style.display = 'none';
            }

            displayParsedData(parsedData);
            showSteps();
            updateCategorySelects();
            updatePreview();
        } catch (err) {
            hideLoading();
            showToast('PDF 解析失敗：' + err.message, 'error');
            console.error(err);
        }
    }

    function displayParsedData(data) {
        // 報價單資訊
        document.getElementById('vendorName').textContent = data.vendorName || '-';
        document.getElementById('quoteDate').textContent = data.quoteDate || '-';
        document.getElementById('totalAmount').textContent = data.totalAmount
            ? `NT$ ${data.totalAmount.toLocaleString()}`
            : '-';

        // 品項表格
        parsedBody.innerHTML = '';
        data.items.forEach((item, idx) => {
            addItemRow(item, idx + 1);
        });
    }

    function addItemRow(item, seq) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-seq">${seq}</td>
            <td class="col-name" contenteditable="true" data-field="name">${escapeHtml(item.name)}</td>
            <td class="col-qty" contenteditable="true" data-field="quantity">${item.quantity || ''}</td>
            <td class="col-price" contenteditable="true" data-field="unitPrice">${item.unitPrice ? item.unitPrice.toLocaleString() : ''}</td>
            <td class="col-subtotal">${item.subtotal ? item.subtotal.toLocaleString() : ''}</td>
            <td class="col-action">
                <button class="btn-delete-row" title="刪除此品項">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                    </svg>
                </button>
            </td>
        `;

        // 刪除按鈕
        tr.querySelector('.btn-delete-row').addEventListener('click', () => {
            tr.remove();
            reindexRows();
            updatePreview();
        });

        // 編輯事件：更新小計
        const editableCells = tr.querySelectorAll('[contenteditable]');
        editableCells.forEach(cell => {
            cell.addEventListener('blur', () => {
                updateRowSubtotal(tr);
                updatePreview();
            });
            cell.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    cell.blur();
                }
            });
        });

        parsedBody.appendChild(tr);
    }

    function updateRowSubtotal(tr) {
        const qty = parseNumber(tr.querySelector('[data-field="quantity"]').textContent);
        const price = parseNumber(tr.querySelector('[data-field="unitPrice"]').textContent);
        const subtotal = Math.round(qty * price);
        tr.querySelector('.col-subtotal').textContent = subtotal ? subtotal.toLocaleString() : '';
    }

    function reindexRows() {
        const rows = parsedBody.querySelectorAll('tr');
        rows.forEach((tr, idx) => {
            tr.querySelector('.col-seq').textContent = idx + 1;
        });
    }

    function getItemsFromTable() {
        const rows = parsedBody.querySelectorAll('tr');
        const items = [];
        rows.forEach(tr => {
            const name = tr.querySelector('[data-field="name"]').textContent.trim();
            const qty = parseNumber(tr.querySelector('[data-field="quantity"]').textContent);
            const price = parseNumber(tr.querySelector('[data-field="unitPrice"]').textContent);
            if (name) {
                items.push({
                    name: name,
                    quantity: qty,
                    unitPrice: price,
                    subtotal: Math.round(qty * price)
                });
            }
        });
        return items;
    }

    function showSteps() {
        step2.style.display = 'block';
        step3.style.display = 'block';
        step4.style.display = 'block';
        // 滾動到解析結果
        step2.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function resetAll() {
        currentFile = null;
        parsedData = null;
        pdfInput.value = '';
        fileInfo.style.display = 'none';
        uploadArea.style.display = 'block';
        step2.style.display = 'none';
        step3.style.display = 'none';
        step4.style.display = 'none';
        parsedBody.innerHTML = '';
    }

    // ===== 預算科目選單 =====

    function updateCategorySelects() {
        const type = templateType.value;
        const budgetGroup = document.getElementById('budgetCategoryGroup');
        const budgetSubGroup = document.getElementById('budgetSubCategoryGroup');
        const agencyGroup = document.getElementById('agencyCategoryGroup');
        const agencySubGroup = document.getElementById('agencySubCategoryGroup');

        if (type === '預算內') {
            budgetGroup.style.display = 'block';
            budgetSubGroup.style.display = 'block';
            agencyGroup.style.display = 'none';
            agencySubGroup.style.display = 'none';
            populateCategories('預算內');
        } else {
            budgetGroup.style.display = 'none';
            budgetSubGroup.style.display = 'none';
            agencyGroup.style.display = 'block';
            agencySubGroup.style.display = 'block';
            populateCategories('代收代辦');
        }
    }

    function populateCategories(type) {
        const cats = budgetCategories[type];
        const selectId = type === '預算內' ? 'budgetCategory' : 'agencyCategory';
        const select = document.getElementById(selectId);

        select.innerHTML = '<option value="">-- 請選擇一級科目 --</option>';
        for (const cat of Object.keys(cats)) {
            const opt = document.createElement('option');
            opt.value = cat;
            opt.textContent = cat;
            select.appendChild(opt);
        }

        // 清空二級
        const subSelectId = type === '預算內' ? 'budgetSubCategory' : 'agencySubCategory';
        const subSelect = document.getElementById(subSelectId);
        subSelect.innerHTML = '<option value="">-- 請先選擇一級科目 --</option>';
    }

    function updateSubCategory(type) {
        const selectId = type === '預算內' ? 'budgetCategory' : 'agencyCategory';
        const subSelectId = type === '預算內' ? 'budgetSubCategory' : 'agencySubCategory';
        const select = document.getElementById(selectId);
        const subSelect = document.getElementById(subSelectId);
        const cat = select.value;
        const cats = budgetCategories[type];

        subSelect.innerHTML = '<option value="">-- 請選擇二級科目 --</option>';

        if (cat && cats[cat]) {
            for (const sub of cats[cat]) {
                if (sub) {
                    const opt = document.createElement('option');
                    opt.value = sub;
                    opt.textContent = sub;
                    subSelect.appendChild(opt);
                }
            }
        }
    }

    // ===== A4：LocalStorage 記憶設定 =====

    const LS_PREFIX = 'smes_budget_';
    const LS_FIELDS = ['templateType', 'unitSelect', 'budgetCategory', 'budgetSubCategory', 'agencyCategory', 'agencySubCategory'];

    function saveSettings() {
        try {
            LS_FIELDS.forEach(id => {
                const el = document.getElementById(id);
                if (el) localStorage.setItem(LS_PREFIX + id, el.value);
            });
        } catch {}
    }

    function initLocalStorage() {
        try {
            // 還原表單類型（需先還原才能初始化科目選單）
            const savedType = localStorage.getItem(LS_PREFIX + 'templateType');
            if (savedType) {
                const el = document.getElementById('templateType');
                if (el) el.value = savedType;
            }

            // 還原單位別
            const savedUnit = localStorage.getItem(LS_PREFIX + 'unitSelect');
            if (savedUnit) {
                const el = document.getElementById('unitSelect');
                if (el) el.value = savedUnit;
            }

            // 初始化科目選單（基於已還原的表單類型）
            updateCategorySelects();

            // 還原一級科目
            const type = document.getElementById('templateType').value;
            const catId = type === '預算內' ? 'budgetCategory' : 'agencyCategory';
            const savedCat = localStorage.getItem(LS_PREFIX + catId);
            if (savedCat) {
                const el = document.getElementById(catId);
                if (el) {
                    el.value = savedCat;
                    updateSubCategory(type);
                }
            }

            // 還原二級科目
            const subCatId = type === '預算內' ? 'budgetSubCategory' : 'agencySubCategory';
            const savedSubCat = localStorage.getItem(LS_PREFIX + subCatId);
            if (savedSubCat) {
                const el = document.getElementById(subCatId);
                if (el) el.value = savedSubCat;
            }
        } catch {}
    }

    // ===== 日期初始化 =====

    function initDate() {
        const now = new Date();
        const rocYear = now.getFullYear() - 1911;
        document.getElementById('dateYear').value = rocYear;
        document.getElementById('dateMonth').value = now.getMonth() + 1;
        document.getElementById('dateDay').value = now.getDate();
    }

    // ===== 預覽 =====

    function updatePreview() {
        const items = getItemsFromTable();
        const summary = document.getElementById('previewSummary');

        if (items.length === 0) {
            summary.innerHTML = '<p style="color:var(--gray-400);text-align:center;">尚無品項資料</p>';
            return;
        }

        const type = templateType.value;
        const unit = document.getElementById('unitSelect').value;
        const purpose = purposeInput.value;
        const year = document.getElementById('dateYear').value;
        const month = document.getElementById('dateMonth').value;
        const day = document.getElementById('dateDay').value;

        let total = 0;
        let itemsHtml = '';
        items.forEach((item, i) => {
            total += item.subtotal;
            itemsHtml += `
                <div class="preview-row">
                    <span class="preview-label">${i + 1}. ${escapeHtml(item.name)}</span>
                    <span class="preview-value">NT$ ${item.subtotal.toLocaleString()}</span>
                </div>`;
        });

        summary.innerHTML = `
            <div class="preview-row">
                <span class="preview-label">表單類型</span>
                <span class="preview-value">${type}</span>
            </div>
            <div class="preview-row">
                <span class="preview-label">單位別</span>
                <span class="preview-value">${unit.trim()}</span>
            </div>
            <div class="preview-row">
                <span class="preview-label">日期</span>
                <span class="preview-value">${year || '?'}年${month || '?'}月${day || '?'}日</span>
            </div>
            <div class="preview-row">
                <span class="preview-label">用途說明</span>
                <span class="preview-value">${purpose || '(未填寫)'}</span>
            </div>
            <hr style="border:none;border-top:1px solid var(--gray-200);margin:8px 0;">
            ${itemsHtml}
            <div class="preview-row preview-total">
                <span class="preview-label">合計金額</span>
                <span class="preview-value">NT$ ${total.toLocaleString()}</span>
            </div>
        `;
    }

    // ===== 產生 Excel =====

    async function generateExcel() {
        const items = getItemsFromTable();
        if (items.length === 0) {
            showToast('請至少填入一筆品項', 'error');
            return;
        }

        const purpose = purposeInput.value.trim();
        if (!purpose) {
            showToast('請填寫用途說明', 'error');
            purposeInput.focus();
            return;
        }

        const type = templateType.value;
        const unit = document.getElementById('unitSelect').value;
        const year = parseInt(document.getElementById('dateYear').value, 10);
        const month = parseInt(document.getElementById('dateMonth').value, 10);
        const day = parseInt(document.getElementById('dateDay').value, 10);

        let budgetCat = '';
        let budgetSub = '';
        if (type === '預算內') {
            budgetCat = document.getElementById('budgetCategory').value;
            budgetSub = document.getElementById('budgetSubCategory').value;
        } else {
            budgetCat = document.getElementById('agencyCategory').value;
            budgetSub = document.getElementById('agencySubCategory').value;
        }

        showLoading('正在產生動支及黏存單...');

        try {
            const data = await ExcelGenerator.generate({
                items: items,
                templateType: type,
                unit: unit,
                purpose: purpose,
                year: year,
                month: month,
                day: day,
                budgetCategory: budgetCat,
                budgetSubCategory: budgetSub
            });

            const dateStr = `${year}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
            const filename = `動支及黏存單_${type}_${dateStr}.xlsx`;
            ExcelGenerator.downloadFile(data, filename);

            hideLoading();
            showToast('Excel 檔案已產生並開始下載', 'success');
        } catch (err) {
            hideLoading();
            showToast('產生失敗：' + err.message, 'error');
            console.error(err);
        }
    }

    // ===== 工具函式 =====

    function parseNumber(text) {
        if (!text) return 0;
        const cleaned = text.replace(/,/g, '').trim();
        const num = parseFloat(cleaned);
        return isNaN(num) ? 0 : num;
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function showLoading(text) {
        loadingText.textContent = text;
        loadingOverlay.style.display = 'flex';
    }

    function hideLoading() {
        loadingOverlay.style.display = 'none';
    }

    function showToast(message, type = 'info') {
        const container = document.getElementById('toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        container.appendChild(toast);

        setTimeout(() => {
            toast.classList.add('fade-out');
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }

    // ===== 模式偵測 =====

    async function initModeDetection() {
        const badge = document.getElementById('modeBadge');
        try {
            const mode = await ExcelGenerator.getMode();
            badge.style.display = 'inline-flex';
            if (mode === 'python') {
                badge.className = 'mode-badge mode-python';
                badge.innerHTML = `
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor">
                        <circle cx="12" cy="12" r="10"/>
                    </svg>
                    本地模式｜格式完整保留`;
            } else {
                badge.className = 'mode-badge mode-static';
                badge.innerHTML = `
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor">
                        <circle cx="12" cy="12" r="10"/>
                    </svg>
                    靜態模式｜格式完整保留`;
            }
        } catch {
            // 靜默失敗
        }
    }
});
