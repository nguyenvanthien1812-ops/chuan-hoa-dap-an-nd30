/**
 * app.js - Main UI Controller & Workflow Orchestrator
 * Pure Client-Side Document Normalization Engine
 */

document.addEventListener('DOMContentLoaded', () => {
    // State
    let currentTab = 'general';
    let selectedFiles = [];
    let lastProcessedBlob = null;
    let lastProcessedFileName = '';

    // DOM Elements
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const toolTabs = document.querySelectorAll('.tool-tab');
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const fileListContainer = document.getElementById('fileListContainer');
    const btnProcess = document.getElementById('btnProcess');
    const btnProcessText = document.getElementById('btnProcessText');
    const progressBox = document.getElementById('progressBox');
    const progressBarFill = document.getElementById('progressBarFill');
    const progressPercent = document.getElementById('progressPercent');
    const progressStatus = document.getElementById('progressStatus');
    const resultsCard = document.getElementById('resultsCard');
    const btnDownload = document.getElementById('btnDownload');
    const toastContainer = document.getElementById('toastContainer');

    const configCardTitle = document.getElementById('configCardTitle');
    const modeBadge = document.getElementById('modeBadge');
    const uploadCardTitle = document.getElementById('uploadCardTitle');

    // Tab Option Stacks
    const optionStacks = {
        general: document.getElementById('optGeneral'),
        english: document.getElementById('optEnglish'),
        nd30: document.getElementById('optNd30'),
        compactor: document.getElementById('optCompactor'),
        batch: document.getElementById('optBatch')
    };

    // =========================================================================
    // 1. Theme Management (Dark / Light)
    // =========================================================================
    function initTheme() {
        const savedTheme = localStorage.getItem('doc_theme') || 'dark';
        document.documentElement.setAttribute('data-theme', savedTheme);
        updateThemeIcon(savedTheme);
    }

    function toggleTheme() {
        const current = document.documentElement.getAttribute('data-theme') || 'dark';
        const next = current === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', next);
        localStorage.setItem('doc_theme', next);
        updateThemeIcon(next);
    }

    function updateThemeIcon(theme) {
        if (!themeToggleBtn) return;
        themeToggleBtn.innerHTML = theme === 'dark' 
            ? '<i class="fa-solid fa-moon"></i>' 
            : '<i class="fa-solid fa-sun"></i>';
    }

    themeToggleBtn.addEventListener('click', toggleTheme);
    initTheme();

    // =========================================================================
    // 2. Tab Navigation
    // =========================================================================
    toolTabs.forEach(tabBtn => {
        tabBtn.addEventListener('click', () => {
            const tab = tabBtn.getAttribute('data-tab');
            switchTab(tab);
        });
    });

    function switchTab(tab) {
        currentTab = tab;
        toolTabs.forEach(btn => {
            btn.classList.toggle('active', btn.getAttribute('data-tab') === tab);
        });

        const mainWorkbench = document.getElementById('mainWorkbench');
        const templatesView = document.getElementById('templatesView');

        if (tab === 'templates') {
            if (mainWorkbench) mainWorkbench.style.display = 'none';
            if (templatesView) {
                templatesView.style.display = 'block';
                renderTemplates();
            }
            return;
        } else {
            if (mainWorkbench) mainWorkbench.style.display = 'grid';
            if (templatesView) templatesView.style.display = 'none';
        }

        // Hide all option stacks, show active
        Object.keys(optionStacks).forEach(k => {
            if (optionStacks[k]) {
                optionStacks[k].style.display = (k === tab) ? 'flex' : 'none';
            }
        });

        // Update titles & labels
        const modeInfo = {
            general: { title: 'Đề Thi Chung', cardTitle: 'Tùy Chọn Đề Thi Chung', btn: 'Chuẩn Hóa Đề Thi', multi: false },
            english: { title: 'Đề Tiếng Anh', cardTitle: 'Tùy Chọn Đề Tiếng Anh', btn: 'Chuẩn Hóa Đề Tiếng Anh', multi: false },
            nd30: { title: 'Nghị Định 30', cardTitle: 'Quy Chuẩn Nghị Định 30', btn: 'Chuẩn Hóa Văn Bản NĐ 30', multi: false },
            compactor: { title: 'Dồn Dòng', cardTitle: 'Tùy Chọn Dồn Dòng Phương Án', btn: 'Dồn Dòng Phương Án', multi: false },
            batch: { title: 'Xử Lý Hàng Loạt', cardTitle: 'Cấu Hình Xử Lý Hàng Loạt', btn: 'Bắt Đầu Xử Lý Tất Cả', multi: true }
        };

        const info = modeInfo[tab] || modeInfo.general;
        modeBadge.textContent = info.title;
        configCardTitle.textContent = info.cardTitle;
        btnProcessText.textContent = info.btn;
        fileInput.multiple = info.multi;
        uploadCardTitle.textContent = info.multi ? 'Chọn nhiều file Word (.docx)' : 'Chọn file Word (.docx)';

        // Reset results display on tab change
        resultsCard.style.display = 'none';
    }

    // =========================================================================
    // 2.1. Template Gallery Renderer & Downloader
    // =========================================================================
    function renderTemplates() {
        const grid = document.getElementById('templateCardsGrid');
        if (!grid || grid.children.length > 0) return; // Đã render rồi thì bỏ qua

        if (!window.TemplateGenerator || !window.TemplateGenerator.TEMPLATES) return;

        grid.innerHTML = window.TemplateGenerator.TEMPLATES.map(tpl => `
            <div class="template-card">
                <div class="template-card-top">
                    <div class="template-icon-box">
                        <i class="fa-solid ${tpl.icon}"></i>
                    </div>
                    <span class="template-category-badge">${tpl.badge}</span>
                </div>
                <div class="template-info">
                    <h4 class="template-name">${tpl.title}</h4>
                    <p class="template-desc">${tpl.desc}</p>
                </div>
                <div class="template-actions">
                    <button class="btn-template-download" data-tpl-id="${tpl.id}" data-tpl-name="${tpl.fileName}">
                        <i class="fa-solid fa-download"></i> Tải Mẫu Word (.docx)
                    </button>
                </div>
            </div>
        `).join('');

        // Gắn sự kiện tải template
        grid.querySelectorAll('.btn-template-download').forEach(btn => {
            btn.addEventListener('click', async () => {
                const tplId = btn.getAttribute('data-tpl-id');
                const tplName = btn.getAttribute('data-tpl-name') || `${tplId}.docx`;
                
                const originalHtml = btn.innerHTML;
                btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang tạo file...';
                btn.disabled = true;

                try {
                    const blob = await window.TemplateGenerator.generateNd30TemplateDocx(tplId);
                    saveAs(blob, tplName);
                    showToast(`Đã tải thành công mẫu: ${tplName}`, 'success');
                } catch (e) {
                    console.error('Error generating template:', e);
                    showToast('Lỗi khi tạo biểu mẫu: ' + e.message, 'error');
                } finally {
                    btn.innerHTML = originalHtml;
                    btn.disabled = false;
                }
            });
        });
    }

    // =========================================================================
    // 3. Drag & Drop and File Selection
    // =========================================================================
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('drag-over');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('drag-over');
        });
    });

    dropzone.addEventListener('drop', (e) => {
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith('.docx'));
        if (files.length === 0) {
            showToast('Vui lòng kéo thả file định dạng Microsoft Word (.docx)', 'error');
            return;
        }
        handleFilesSelected(files);
    });

    fileInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files).filter(f => f.name.toLowerCase().endsWith('.docx'));
        if (files.length > 0) {
            handleFilesSelected(files);
        }
    });

    function handleFilesSelected(files) {
        if (currentTab === 'batch') {
            selectedFiles = [...selectedFiles, ...files];
        } else {
            selectedFiles = [files[0]];
        }
        renderFileList();
        btnProcess.disabled = (selectedFiles.length === 0);
        resultsCard.style.display = 'none';
    }

    function renderFileList() {
        if (selectedFiles.length === 0) {
            fileListContainer.style.display = 'none';
            fileListContainer.innerHTML = '';
            btnProcess.disabled = true;
            return;
        }

        fileListContainer.style.display = 'flex';
        fileListContainer.innerHTML = '';

        selectedFiles.forEach((file, idx) => {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <div class="file-info">
                    <i class="fa-solid fa-file-word file-icon"></i>
                    <div>
                        <div class="file-name" title="${file.name}">${file.name}</div>
                        <div class="file-size">${formatFileSize(file.size)}</div>
                    </div>
                </div>
                <button type="button" class="btn-remove-file" title="Xóa file" data-idx="${idx}">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            `;
            fileListContainer.appendChild(item);
        });

        // Add remove handlers
        fileListContainer.querySelectorAll('.btn-remove-file').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(btn.getAttribute('data-idx'), 10);
                selectedFiles.splice(idx, 1);
                renderFileList();
            });
        });
    }

    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
    }

    // =========================================================================
    // 4. Processing & Execution
    // =========================================================================
    btnProcess.addEventListener('click', async () => {
        if (selectedFiles.length === 0) return;

        // Check license authorization before processing
        if (window.FirebaseAuth && typeof window.FirebaseAuth.checkLicenseAccess === 'function') {
            const license = window.FirebaseAuth.checkLicenseAccess();
            if (!license.allowed) {
                showToast(license.message || 'Bạn chưa có quyền sử dụng tính năng này.', 'error');
                if (window.FirebaseAuth.renderLicenseGateBanner) {
                    window.FirebaseAuth.renderLicenseGateBanner();
                }
                return;
            }
        }

        btnProcess.disabled = true;
        progressBox.style.display = 'flex';
        resultsCard.style.display = 'none';
        await updateProgress(10, 'Đang đọc file Word...');

        const startTime = performance.now();

        try {
            if (currentTab === 'batch') {
                await processBatchFiles(startTime);
            } else {
                await processSingleFile(selectedFiles[0], startTime);
            }
        } catch (err) {
            console.error('Processing error:', err);
            showToast('Lỗi khi xử lý file: ' + err.message, 'error');
            await updateProgress(0, 'Đã xảy ra lỗi: ' + err.message);
        } finally {
            btnProcess.disabled = false;
        }
    });

    async function processSingleFile(file, startTime) {
        await updateProgress(25, 'Đang đọc file và cấu trúc OpenXML...');
        const arrayBuffer = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);

        await updateProgress(50, 'Đang phân tích và thực thi chuẩn hóa...');
        let result = null;

        if (currentTab === 'general') {
            const auto_key = document.getElementById('genAutoKey').checked;
            const bold_as_correct = document.getElementById('genBoldAsCorrect').checked;
            const choice_layout = document.querySelector('input[name="genChoiceLayout"]:checked')?.value || 'auto';
            const choice_indent = document.getElementById('genChoiceIndent')?.checked !== false;
            const indent_dxa = choice_indent ? parseInt(document.getElementById('genChoiceIndentSize')?.value || '567', 10) : 0;

            result = await window.Normalizer.normalizeDocx(zip, {
                auto_key,
                bold_as_correct,
                choice_layout,
                choice_indent,
                indent_dxa
            });
        } else if (currentTab === 'english') {
            const auto_key = document.getElementById('engAutoKey').checked;
            const bold_as_correct = document.getElementById('engBoldAsCorrect').checked;
            const choice_layout = document.querySelector('input[name="engChoiceLayout"]:checked')?.value || 'auto';
            const choice_indent = document.getElementById('engChoiceIndent')?.checked !== false;
            const indent_dxa = choice_indent ? parseInt(document.getElementById('engChoiceIndentSize')?.value || '567', 10) : 0;

            result = await window.EnglishNormalizer.normalizeEnglishDocx(zip, {
                auto_key,
                bold_as_correct,
                choice_layout,
                choice_indent,
                indent_dxa
            });
        } else if (currentTab === 'nd30') {
            const margin = document.getElementById('nd30Margin')?.checked !== false;
            const page_num = document.getElementById('nd30PageNum')?.checked !== false;
            const headings = document.getElementById('nd30Headings')?.checked !== false;
            const paragraph = document.getElementById('nd30Paragraph')?.checked !== false;
            const font = document.getElementById('nd30Font')?.checked !== false;
            const convert_fonts = document.getElementById('nd30ConvertFonts')?.checked !== false;
            const clean_punctuation = document.getElementById('nd30CleanPunctuation')?.checked !== false;
            const remove_empty_paras = document.getElementById('nd30RemoveEmptyParas')?.checked !== false;

            result = await window.Nd30Normalizer.normalizeDocumentNd30(zip, {
                margin,
                page_num,
                headings,
                paragraph,
                font,
                convert_fonts,
                clean_punctuation,
                remove_empty_paras
            });
        } else if (currentTab === 'compactor') {
            const compactorMode = document.querySelector('input[name="compactorMode"]:checked')?.value || 'auto';
            const choice_indent = document.getElementById('compactorChoiceIndent')?.checked !== false;
            const indent_dxa = choice_indent ? parseInt(document.getElementById('compactorChoiceIndentSize')?.value || '567', 10) : 0;
            result = await window.ChoiceCompactor.compactDocxChoices(zip, compactorMode, indent_dxa);
        }

        await updateProgress(85, 'Đang đóng gói file kết quả...');
        const outBlob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        const elapsedSec = ((performance.now() - startTime) / 1000).toFixed(2);
        await updateProgress(100, 'Hoàn thành trong ' + elapsedSec + 's');

        // Suffix mapping
        const suffixMap = {
            general: '_ChuanHoa.docx',
            english: '_English_ChuanHoa.docx',
            nd30: '_ND30.docx',
            compactor: '_DonDong.docx'
        };
        const suffix = suffixMap[currentTab] || '_ChuanHoa.docx';
        const baseName = file.name.replace(/\.docx$/i, '');
        lastProcessedFileName = `${baseName}${suffix}`;
        lastProcessedBlob = outBlob;

        // Display results
        showResults(result, elapsedSec);
        showToast(result.message || 'Chuẩn hóa file thành công!', 'success');
    }

    async function processBatchFiles(startTime) {
        const batchEngine = document.querySelector('input[name="batchEngine"]:checked')?.value || 'general';
        const total = selectedFiles.length;
        const outZip = new JSZip();

        let totalKeysFound = 0;
        let totalCompacted = 0;

        const batch_choice_indent = document.getElementById('batchChoiceIndent')?.checked !== false;
        const batch_indent_dxa = batch_choice_indent ? parseInt(document.getElementById('batchChoiceIndentSize')?.value || '567', 10) : 0;

        for (let i = 0; i < total; i++) {
            const file = selectedFiles[i];
            const percent = Math.round(((i + 0.2) / total) * 100);
            await updateProgress(percent, `Đang xử lý (${i + 1}/${total}): ${file.name}`);

            const arrayBuffer = await file.arrayBuffer();
            const zip = await JSZip.loadAsync(arrayBuffer);

            let res = null;
            if (batchEngine === 'general') {
                res = await window.Normalizer.normalizeDocx(zip, { 
                    auto_key: true, 
                    bold_as_correct: true, 
                    choice_layout: 'auto',
                    choice_indent: batch_choice_indent,
                    indent_dxa: batch_indent_dxa
                });
            } else if (batchEngine === 'english') {
                res = await window.EnglishNormalizer.normalizeEnglishDocx(zip, { 
                    auto_key: true, 
                    bold_as_correct: true, 
                    choice_layout: 'auto',
                    choice_indent: batch_choice_indent,
                    indent_dxa: batch_indent_dxa
                });
            } else if (batchEngine === 'nd30') {
                res = await window.Nd30Normalizer.normalizeDocumentNd30(zip, { 
                    margin: true, 
                    page_num: true, 
                    headings: true, 
                    paragraph: true, 
                    font: true,
                    convert_fonts: true,
                    clean_punctuation: true,
                    remove_empty_paras: true
                });
            } else {
                res = await window.ChoiceCompactor.compactDocxChoices(zip, 'auto', batch_indent_dxa);
            }

            if (res && res.keyCount) totalKeysFound += res.keyCount;
            if (res && res.compactedCount) totalCompacted += res.compactedCount;

            const fileBlob = await zip.generateAsync({ type: 'blob' });
            const baseName = file.name.replace(/\.docx$/i, '');
            outZip.file(`${baseName}_ChuanHoa.docx`, fileBlob);
        }

        await updateProgress(95, 'Đang nén toàn bộ tệp ZIP...');
        const zipBlob = await outZip.generateAsync({ type: 'blob' });
        const elapsedSec = ((performance.now() - startTime) / 1000).toFixed(2);
        await updateProgress(100, `Hoàn tất ${total} files trong ${elapsedSec}s`);

        lastProcessedFileName = `KetQua_ChuanHoa_${Date.now()}.zip`;
        lastProcessedBlob = zipBlob;

        showResults({
            keyCount: totalKeysFound,
            compactedCount: totalCompacted,
            isBatch: true,
            totalFiles: total
        }, elapsedSec);

        showToast(`Đã chuẩn hóa thành công toàn bộ ${total} file!`, 'success');
    }

    async function updateProgress(percent, text) {
        progressBarFill.style.width = percent + '%';
        progressPercent.textContent = percent + '%';
        progressStatus.textContent = text;
        // Nhường luồng cho trình duyệt vẽ giao diện mượt mà
        await new Promise(resolve => setTimeout(resolve, 30));
    }

    function showResults(result, elapsedSec) {
        resultsCard.style.display = 'block';
        document.getElementById('resTimeBadge').textContent = elapsedSec + 's';

        const statKeysFound = document.getElementById('statKeysFound');
        const statCompacted = document.getElementById('statCompacted');
        const statStatus = document.getElementById('statStatus');

        statKeysFound.textContent = result.keyCount !== undefined ? result.keyCount : '—';
        statCompacted.textContent = result.compactedCount !== undefined ? result.compactedCount : '—';
        statStatus.textContent = '100%';

        // Answer Key Inspector Grid
        const keysContainer = document.getElementById('keysInspectorContainer');
        const keysGrid = document.getElementById('keysBadgeGrid');
        keysGrid.innerHTML = '';

        if (result.answerKeys && Object.keys(result.answerKeys).length > 0) {
            let hasAnyKey = false;
            Object.entries(result.answerKeys).forEach(([secIdx, secKeys]) => {
                if (Object.keys(secKeys).length > 0) {
                    hasAnyKey = true;
                    Object.entries(secKeys).sort((a, b) => parseInt(a[0], 10) - parseInt(b[0], 10)).forEach(([qNum, qVal]) => {
                        const badge = document.createElement('span');
                        badge.className = 'key-badge';
                        badge.innerHTML = `<span class="q-num">C${qNum}:</span> <span class="q-val">${qVal}</span>`;
                        keysGrid.appendChild(badge);
                    });
                }
            });

            keysContainer.style.display = hasAnyKey ? 'block' : 'none';
        } else {
            keysContainer.style.display = 'none';
        }

        // Update download button label
        if (result.isBatch) {
            btnDownload.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Tải Về Tất Cả File (.zip)';
        } else {
            btnDownload.innerHTML = '<i class="fa-solid fa-download"></i> Tải Về File Chuẩn Hóa (.docx)';
        }
    }

    // =========================================================================
    // 5. Download Handlers
    // =========================================================================
    btnDownload.addEventListener('click', () => {
        if (!lastProcessedBlob) return;
        if (typeof saveAs !== 'undefined') {
            saveAs(lastProcessedBlob, lastProcessedFileName);
        } else {
            const url = URL.createObjectURL(lastProcessedBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = lastProcessedFileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 1000);
        }
        showToast('Đang tải xuống file...', 'info');
    });

    // =========================================================================
    // 6. Toast Notifications
    // =========================================================================
    function showToast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;

        const iconMap = {
            success: 'fa-circle-check',
            error: 'fa-circle-exclamation',
            info: 'fa-circle-info'
        };
        const icon = iconMap[type] || iconMap.info;

        toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
        toastContainer.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateY(20px)';
            toast.style.transition = 'all 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    window.showToast = showToast;
});
