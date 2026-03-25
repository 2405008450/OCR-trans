'use strict';

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview = document.getElementById('filePreview');
const fileSummary = document.getElementById('fileSummary');
const fileChipRow = document.getElementById('fileChipRow');
const btnRemove = document.getElementById('btnRemove');
const btnProcess = document.getElementById('btnProcess');
const btnNewTask = document.getElementById('btnNewTask');
const processingMode = document.getElementById('processingMode');
const modeTip = document.getElementById('modeTip');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const streamLogWrap = document.getElementById('streamLogWrap');
const streamLog = document.getElementById('streamLog');
const resultStats = document.getElementById('resultStats');
const resultFileList = document.getElementById('resultFileList');

let selectedFiles = [];
let pollingTimer = null;
let config = { processing_modes: {} };

const MODE_TIPS = {
    single: '单图处理只能上传 1 张图片。',
    merge: '多图合并会把多张属于同一驾驶证的图片合并成 1 份译文。',
    batch: '多图批量会为每张图片分别生成一份译文。',
};

init();

async function init() {
    bindEvents();
    await loadConfig();
    updateModeTip();
    updateProcessButton();
}

function bindEvents() {
    uploadArea?.addEventListener('click', () => fileInput?.click());
    uploadArea?.addEventListener('dragover', handleDragOver);
    uploadArea?.addEventListener('drop', handleDrop);
    fileInput?.addEventListener('change', handleFileSelect);
    btnRemove?.addEventListener('click', (event) => {
        event.stopPropagation();
        clearFiles();
    });
    btnProcess?.addEventListener('click', processFiles);
    btnNewTask?.addEventListener('click', resetPage);
    processingMode?.addEventListener('change', () => {
        updateModeTip();
        updateProcessButton();
    });
}

async function loadConfig() {
    try {
        const response = await fetch('/task/drivers-license/config');
        if (!response.ok) throw new Error(`配置加载失败: ${response.status}`);
        config = await response.json();
    } catch (error) {
        console.error(error);
        config = {
            processing_modes: {
                single: { label: '单图处理' },
                merge: { label: '多图合并' },
                batch: { label: '多图批量' },
            },
        };
    }

    processingMode.innerHTML = '';
    Object.entries(config.processing_modes || {}).forEach(([value, info]) => {
        processingMode.add(new Option(info.label || value, value));
    });
    processingMode.value = config.processing_modes.merge ? 'merge' : Object.keys(config.processing_modes)[0] || 'merge';
}

function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    uploadArea.style.borderColor = 'var(--primary-color)';
}

function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    uploadArea.style.borderColor = 'var(--border-color)';
    const files = Array.from(event.dataTransfer.files || []);
    if (files.length > 0) handleFiles(files);
}

function handleFileSelect(event) {
    const files = Array.from(event.target.files || []);
    if (files.length > 0) handleFiles(files);
}

function handleFiles(files) {
    const allowedExtensions = ['.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'];
    const invalid = files.find((file) => !allowedExtensions.some((ext) => file.name.toLowerCase().endsWith(ext)));
    if (invalid) {
        alert(`不支持的图片格式: ${invalid.name}`);
        return;
    }

    selectedFiles = files;
    fileSummary.textContent = `已选择 ${files.length} 张图片`;
    fileChipRow.innerHTML = files.map((file) => `<span class="file-chip"><i class="fas fa-image"></i>${escapeHtml(file.name)}</span>`).join('');
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    updateProcessButton();
}

function clearFiles() {
    selectedFiles = [];
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    fileChipRow.innerHTML = '';
    updateProcessButton();
}

function updateModeTip() {
    modeTip.textContent = MODE_TIPS[processingMode.value] || '';
}

function updateProcessButton() {
    const mode = processingMode.value;
    const fileCount = selectedFiles.length;
    btnProcess.disabled = fileCount === 0 || (mode === 'single' && fileCount !== 1);
}

async function processFiles() {
    if (!selectedFiles.length) return;

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    updateProgress(5, '正在提交任务...');

    try {
        const formData = new FormData();
        selectedFiles.forEach((file) => formData.append('files', file));
        const params = new URLSearchParams({ processing_mode: processingMode.value });

        const response = await fetch(`/task/drivers-license?${params.toString()}`, {
            method: 'POST',
            body: formData,
        });

        if (!response.ok) {
            const message = await safeReadError(response);
            throw new Error(message || `提交失败: ${response.status}`);
        }

        const data = await response.json();
        startPolling(data.task_id);
    } catch (error) {
        showFailure(error.message);
    }
}

function startPolling(taskId) {
    stopPolling();
    pollStatus(taskId);
    pollingTimer = setInterval(() => pollStatus(taskId), 2500);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
}

async function pollStatus(taskId) {
    try {
        const response = await fetch(`/task/drivers-license/status/${taskId}`);
        if (!response.ok) throw new Error(`状态查询失败: ${response.status}`);
        const data = await response.json();

        updateProgress(data.progress || 0, data.message || '正在处理...');
        syncLog(data.stream_log || data.result?.stream_log || '');

        if (data.status === 'done') {
            stopPolling();
            showResult(data.result || {});
        } else if (data.status === 'failed') {
            stopPolling();
            showFailure(data.error || data.message || '处理失败', data.stream_log || '');
        }
    } catch (error) {
        stopPolling();
        showFailure(error.message);
    }
}

function updateProgress(percent, message) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;
    processingStatus.textContent = message;
}

function syncLog(text) {
    if (!text) return;
    streamLogWrap.style.display = 'block';
    streamLog.textContent = text;
    streamLog.scrollTop = streamLog.scrollHeight;
}

function clearLog() {
    streamLogWrap.style.display = 'none';
    streamLog.textContent = '';
}

function showFailure(message, logText = '') {
    syncLog(logText);
    updateProgress(100, message || '处理失败');
}

function showResult(result) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'grid';

    const modeLabel = (config.processing_modes?.[result.processing_mode] || {}).label || result.processing_mode || '-';
    const items = Array.isArray(result.items) ? result.items : [];
    const doneCount = items.filter((item) => item.status === 'done' || item.status === 'merged').length;

    resultStats.innerHTML = `
        <div class="stat-card">
            <i class="fas fa-images"></i>
            <h3>${result.input_count || selectedFiles.length}</h3>
            <p>处理图片数</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-layer-group"></i>
            <h3>${escapeHtml(modeLabel)}</h3>
            <p>处理模式</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-file-word"></i>
            <h3>${doneCount}</h3>
            <p>成功输出数</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-check-circle"></i>
            <h3>成功</h3>
            <p>任务状态</p>
        </div>
    `;

    const fileItems = [];
    if (result.output_docx) {
        fileItems.push({
            name: '驾驶证译文.docx',
            path: result.output_docx,
            desc: result.processing_mode === 'merge' ? '合并输出' : '单图输出',
        });
    }
    items.forEach((item) => {
        if (item.output_docx) {
            fileItems.push({
                name: item.input_filename || '输出文件',
                path: item.output_docx,
                desc: item.status === 'merged' ? '合并来源图片' : '对应输出 Word',
            });
        }
    });

    if (!fileItems.length) {
        resultFileList.innerHTML = '<div class="result-file-item"><div class="result-file-meta"><strong>没有可下载文件</strong><span>请在任务看板查看详细日志。</span></div></div>';
        return;
    }

    resultFileList.innerHTML = fileItems.map((item) => `
        <div class="result-file-item">
            <div class="result-file-meta">
                <strong>${escapeHtml(item.name)}</strong>
                <span>${escapeHtml(item.desc)}</span>
            </div>
            <a class="download-btn" href="/${item.path}" download>
                <i class="fas fa-download"></i> 下载
            </a>
        </div>
    `).join('');
}

function resetPage() {
    clearFiles();
    clearLog();
    stopPolling();
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    updateProgress(0, '等待开始处理');
}

async function safeReadError(response) {
    try {
        const payload = await response.json();
        return payload?.detail?.error || payload?.detail || payload?.message || '';
    } catch (_) {
        return '';
    }
}

function escapeHtml(value) {
    return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
