// BOR 排班系統 - 前端互動腳本（無狀態版本）

document.addEventListener('DOMContentLoaded', function() {
    // 狀態管理（只在前端保存，不傳到後端儲存）
    const state = {
        currentStep: 1,
        file: null,  // 保存檔案物件
        filename: null,
        staffList: [],
        identityMap: {}
    };

    // DOM 元素
    const elements = {
        // 步驟
        steps: document.querySelectorAll('.step'),
        stepConnectors: document.querySelectorAll('.step-connector'),
        stepPanels: document.querySelectorAll('.step-panel'),
        
        // 步驟 1
        uploadArea: document.getElementById('upload-area'),
        fileInput: document.getElementById('file-input'),
        fileInfo: document.getElementById('file-info'),
        fileName: document.getElementById('file-name'),
        fileStaff: document.getElementById('file-staff'),
        clearFile: document.getElementById('clear-file'),
        templateUploadArea: document.getElementById('template-upload-area'),
        templateInput: document.getElementById('template-input'),
        templateInfo: document.getElementById('template-info'),
        templateName: document.getElementById('template-name'),
        clearTemplate: document.getElementById('clear-template'),
        btnNext1: document.getElementById('btn-next-1'),
        
        // 步驟 2
        yearSelect: document.getElementById('year-select'),
        monthSelect: document.getElementById('month-select'),
        staffOrder: document.getElementById('staff-order'),
        btnAutoFill: document.getElementById('btn-auto-fill'),
        btnClearStaff: document.getElementById('btn-clear-staff'),
        staffTags: document.getElementById('staff-tags'),
        loadedCount: document.getElementById('loaded-count'),
        btnBack2: document.getElementById('btn-back-2'),
        btnGenerate: document.getElementById('btn-generate'),
        
        // 步驟 3
        resultIcon: document.getElementById('result-icon'),
        resultTitle: document.getElementById('result-title'),
        resultDesc: document.getElementById('result-desc'),
        downloadBox: document.getElementById('download-box'),
        downloadFilename: document.getElementById('download-filename'),
        downloadBtn: document.getElementById('download-btn'),
        btnRestart: document.getElementById('btn-restart'),
        
        // Loading
        loadingOverlay: document.getElementById('loading-overlay'),
        loadingText: document.getElementById('loading-text')
    };

    // 初始化
    init();

    function init() {
        initYearMonthSelects();
        bindEvents();
        // 隱藏模板上傳區域（無狀態版本不需要）
        const templateSection = document.querySelector('.template-section');
        if (templateSection) {
            templateSection.style.display = 'none';
        }
    }

    // 初始化年月選單
    function initYearMonthSelects() {
        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth() + 1;

        // 年份選單（前一年到後一年）
        for (let year = currentYear - 1; year <= currentYear + 1; year++) {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = `${year} 年`;
            if (year === currentYear) option.selected = true;
            elements.yearSelect.appendChild(option);
        }

        // 月份選單
        for (let month = 1; month <= 12; month++) {
            const option = document.createElement('option');
            option.value = month;
            option.textContent = `${month} 月`;
            if (month === currentMonth) option.selected = true;
            elements.monthSelect.appendChild(option);
        }
    }

    // 綁定事件
    function bindEvents() {
        // 主檔案上傳
        elements.uploadArea.addEventListener('click', () => elements.fileInput.click());
        elements.fileInput.addEventListener('change', handleFileSelect);
        elements.clearFile.addEventListener('click', clearFile);
        
        // 拖放上傳
        elements.uploadArea.addEventListener('dragover', handleDragOver);
        elements.uploadArea.addEventListener('dragleave', handleDragLeave);
        elements.uploadArea.addEventListener('drop', handleDrop);
        
        // 步驟導航
        elements.btnNext1.addEventListener('click', () => goToStep(2));
        elements.btnBack2.addEventListener('click', () => goToStep(1));
        elements.btnGenerate.addEventListener('click', generateSchedule);
        elements.btnRestart.addEventListener('click', restart);
        
        // 人員排序
        elements.btnAutoFill.addEventListener('click', autoFillStaff);
        elements.btnClearStaff.addEventListener('click', () => {
            elements.staffOrder.value = '';
        });
    }

    // 拖放處理
    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        elements.uploadArea.classList.add('drag-over');
    }

    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        elements.uploadArea.classList.remove('drag-over');
    }

    function handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        elements.uploadArea.classList.remove('drag-over');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            elements.fileInput.files = files;
            handleFileSelect({ target: elements.fileInput });
        }
    }

    // 檔案選擇處理（只預覽，不儲存到伺服器）
    async function handleFileSelect(e) {
        const file = e.target.files[0];
        if (!file) return;

        // 保存檔案到前端狀態
        state.file = file;
        state.filename = file.name;

        showLoading('正在讀取檔案...');
        
        const formData = new FormData();
        formData.append('file', file);

        try {
            // 只預覽檔案內容，不儲存
            const response = await fetch('/api/preview', {
                method: 'POST',
                body: formData
            });
            
            const result = await response.json();
            
            if (result.success) {
                state.staffList = result.staff_list;
                state.identityMap = result.identity_map;
                
                // 更新 UI
                elements.fileName.textContent = result.filename;
                elements.fileStaff.textContent = `已載入 ${result.staff_count} 位人員`;
                elements.uploadArea.style.display = 'none';
                elements.fileInfo.style.display = 'flex';
                elements.btnNext1.disabled = false;
                
                // 更新人員預覽
                updateStaffPreview();
            } else {
                alert('讀取失敗：' + result.error);
                clearFileState();
            }
        } catch (error) {
            alert('讀取失敗：' + error.message);
            clearFileState();
        } finally {
            hideLoading();
        }
    }

    // 清除檔案狀態
    function clearFileState() {
        state.file = null;
        state.filename = null;
        state.staffList = [];
        state.identityMap = {};
    }

    // 清除檔案
    function clearFile(e) {
        e.stopPropagation();
        clearFileState();
        
        elements.fileInput.value = '';
        elements.uploadArea.style.display = 'block';
        elements.fileInfo.style.display = 'none';
        elements.btnNext1.disabled = true;
        elements.staffTags.innerHTML = '';
        elements.loadedCount.textContent = '0';
    }

    // 更新人員預覽
    function updateStaffPreview() {
        elements.staffTags.innerHTML = '';
        elements.loadedCount.textContent = state.staffList.length;
        
        state.staffList.forEach(name => {
            const tag = document.createElement('span');
            tag.className = 'staff-tag';
            tag.textContent = name;
            
            // 根據身分設定樣式
            const identity = state.identityMap[name];
            if (identity === '公職') {
                tag.classList.add('public');
            } else if (identity === '契約') {
                tag.classList.add('contract');
            }
            
            elements.staffTags.appendChild(tag);
        });
    }

    // 自動填入人員
    function autoFillStaff() {
        if (state.staffList.length > 0) {
            elements.staffOrder.value = state.staffList.join('\n');
        }
    }

    // 步驟切換
    function goToStep(step) {
        state.currentStep = step;
        
        // 更新步驟指示器
        elements.steps.forEach((el, index) => {
            el.classList.remove('active', 'completed');
            if (index + 1 < step) {
                el.classList.add('completed');
            } else if (index + 1 === step) {
                el.classList.add('active');
            }
        });
        
        // 更新連接線
        elements.stepConnectors.forEach((el, index) => {
            if (index + 1 < step) {
                el.classList.add('filled');
            } else {
                el.classList.remove('filled');
            }
        });
        
        // 切換面板
        document.getElementById(`step${step}-panel`).style.display = 'block';
        elements.stepPanels.forEach((el, index) => {
            if (index + 1 !== step) {
                el.style.display = 'none';
            }
        });
        
        // 滾動到頂部
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // 產生排班表（重新上傳檔案，不依賴伺服器儲存）
    async function generateSchedule() {
        if (!state.file) {
            alert('請先上傳檔案');
            goToStep(1);
            return;
        }

        const year = parseInt(elements.yearSelect.value);
        const month = parseInt(elements.monthSelect.value);
        const staffOrder = elements.staffOrder.value;

        showLoading('正在產生排班表...');

        try {
            // 使用 FormData 重新傳送檔案
            const formData = new FormData();
            formData.append('file', state.file);
            formData.append('year', year);
            formData.append('month', month);
            formData.append('staff_order', staffOrder);

            const response = await fetch('/api/generate', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                // 取得檔案名稱
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = `BOR_${year}${String(month).padStart(2, '0')}_排班表.xlsx`;
                if (contentDisposition) {
                    const filenameMatch = contentDisposition.match(/filename\*?=(?:UTF-8'')?([^;\n]*)/i);
                    if (filenameMatch && filenameMatch[1]) {
                        filename = decodeURIComponent(filenameMatch[1].replace(/['"]/g, ''));
                    }
                }

                // 下載檔案
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                
                // 更新結果頁面
                elements.resultIcon.className = 'result-icon success';
                elements.resultIcon.innerHTML = `
                    <svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="32" cy="32" r="28" stroke="currentColor" stroke-width="3"/>
                        <path d="M20 32L28 40L44 24" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                `;
                elements.resultTitle.textContent = '排班表已產生！';
                elements.resultDesc.textContent = '您的排班表已經處理完成，點擊下方按鈕下載檔案';
                elements.downloadFilename.textContent = filename;
                elements.downloadBtn.href = url;
                elements.downloadBtn.download = filename;
                elements.downloadBox.style.display = 'flex';
                
                goToStep(3);
            } else {
                const result = await response.json();
                showError(result.error || '產生失敗');
            }
        } catch (error) {
            showError(error.message);
        } finally {
            hideLoading();
        }
    }

    // 顯示錯誤
    function showError(message) {
        elements.resultIcon.className = 'result-icon error';
        elements.resultIcon.innerHTML = `
            <svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
                <circle cx="32" cy="32" r="28" stroke="currentColor" stroke-width="3"/>
                <path d="M24 24L40 40M40 24L24 40" stroke="currentColor" stroke-width="3" stroke-linecap="round"/>
            </svg>
        `;
        elements.resultTitle.textContent = '產生失敗';
        elements.resultDesc.textContent = message;
        elements.downloadBox.style.display = 'none';
        
        goToStep(3);
    }

    // 重新開始
    function restart() {
        // 重置狀態
        state.currentStep = 1;
        clearFileState();
        
        // 重置 UI
        elements.fileInput.value = '';
        elements.uploadArea.style.display = 'block';
        elements.fileInfo.style.display = 'none';
        elements.btnNext1.disabled = true;
        elements.staffOrder.value = '';
        elements.staffTags.innerHTML = '';
        elements.loadedCount.textContent = '0';
        
        goToStep(1);
    }

    // Loading 控制
    function showLoading(text = '正在處理中...') {
        elements.loadingText.textContent = text;
        elements.loadingOverlay.style.display = 'flex';
    }

    function hideLoading() {
        elements.loadingOverlay.style.display = 'none';
    }
});
