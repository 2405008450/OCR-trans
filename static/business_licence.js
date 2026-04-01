'use strict';

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview = document.getElementById('filePreview');
const fileName = document.getElementById('fileName');
const fileStatus = document.getElementById('fileStatus');
const btnRemove = document.getElementById('btnRemove');
const btnProcess = document.getElementById('btnProcess');
const btnNewTask = document.getElementById('btnNewTask');
const modelSelect = document.getElementById('modelSelect');
const geminiRouteSelect = document.getElementById('geminiRouteSelect');

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

let selectedFile = null;
let pollingTimer = null;
let modelConfig = {};
let routeConfig = {};
let defaultModel = 'google/gemini-3.1-pro-preview';
let defaultRoute = 'openrouter';
let etaHint = null;

const ETA_TIME_ZONE = 'Asia/Shanghai';
const ALLOWED_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp', '.gif'];

init();

async function init() {
    bindEvents();
    await loadConfig();
    updateProcessButton();
}

function bindEvents() {
    uploadArea?.addEventListener('click', () => fileInput?.click());
    uploadArea?.addEventListener('dragover', handleDragOver);
    uploadArea?.addEventListener('drop', handleDrop);
    fileInput?.addEventListener('change', handleFileSelect);
    btnRemove?.addEventListener('click', (event) => {
        event.stopPropagation();
        clearFile();
    });
    btnProcess?.addEventListener('click', processFile);
    btnNewTask?.addEventListener('click', resetPage);
}

async function loadConfig() {
    try {
        const response = await fetch('/task/business-licence/config');
        if (!response.ok) {
            throw new Error(`配置加载失败: ${response.status}`);
        }

        const data = await response.json();
        modelConfig = data.models || {};
        routeConfig = data.routes || {};
        defaultModel = data.default_model || defaultModel;
        defaultRoute = data.default_route || defaultRoute;
    } catch (error) {
        console.error(error);
        modelConfig = {
            'google/gemini-3.1-pro-preview': { label: 'Gemini 3.1 Pro' },
            'google/gemini-3-flash-preview': { label: 'Gemini 3 Flash' },
        };
        routeConfig = {
            openrouter: { label: 'OpenRouter' },
            google: { label: 'Google' },
        };
    }

    renderModels();
    renderRoutes();
}

function renderModels() {
    modelSelect.innerHTML = '';
    Object.entries(modelConfig).forEach(([value, info]) => {
        modelSelect.add(new Option(info.label || value, value));
    });

    const fallback = Object.keys(modelConfig)[0] || defaultModel;
    modelSelect.value = modelConfig[defaultModel] ? defaultModel : fallback;
}

function renderRoutes() {
    geminiRouteSelect.innerHTML = '';
    Object.entries(routeConfig).forEach(([value, info]) => {
        geminiRouteSelect.add(new Option(info.label || value, value));
    });

    const fallback = Object.keys(routeConfig)[0] || defaultRoute;
    geminiRouteSelect.value = routeConfig[defaultRoute] ? defaultRoute : fallback;
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
    const file = event.dataTransfer.files?.[0];
    if (file) {
        handleFile(file);
    }
}

function handleFileSelect(event) {
    const file = event.target.files?.[0];
    if (file) {
        handleFile(file);
    }
}

function handleFile(file) {
    if (!ALLOWED_EXTENSIONS.some((ext) => file.name.toLowerCase().endsWith(ext))) {
        alert(`不支持的图片格式: ${file.name}`);
        return;
    }

    selectedFile = file;
    fileName.textContent = `${file.name} (${formatFileSize(file.size)})`;
    fileStatus.textContent = '文件已就绪';
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    updateProcessButton();
}

function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    updateProcessButton();
}

function updateProcessButton() {
    btnProcess.disabled = !selectedFile;
}

async function processFile() {
    if (!selectedFile) {
        return;
    }

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    updateProgress(5, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('file', selectedFile);

        const params = new URLSearchParams({
            model: modelSelect.value,
            gemini_route: geminiRouteSelect.value,
        });

        const response = await fetch(`/task/business-licence?${params.toString()}`, {
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
    if (!pollingTimer) {
        return;
    }

    clearInterval(pollingTimer);
    pollingTimer = null;
}

async function pollStatus(taskId) {
    try {
        const response = await fetch(`/task/business-licence/status/${taskId}`);
        if (!response.ok) {
            throw new Error(`状态查询失败: ${response.status}`);
        }

        const data = await response.json();
        updateProgress(data.progress || 0, data.message || '正在处理...', data);
        syncLog(data.stream_log || data.result?.stream_log || '');

        if (data.status === 'done') {
            stopPolling();
            showResult(data.result || {});
            return;
        }

        if (data.status === 'failed') {
            stopPolling();
            showFailure(data.error || data.message || '处理失败', data.stream_log || '');
        }
    } catch (error) {
        stopPolling();
        showFailure(error.message);
    }
}

function updateProgress(percent, message, task = null) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;
    processingStatus.textContent = message;
    updateEtaHint(task);
}

function ensureEtaHint() {
    if (etaHint && etaHint.isConnected) {
        return etaHint;
    }

    const card = processingSection?.querySelector('.processing-card') || processingSection;
    if (!card) {
        return null;
    }

    etaHint = document.createElement('div');
    etaHint.className = 'eta-hint';
    etaHint.style.cssText = 'margin-top:10px;color:var(--text-secondary, var(--muted, #94a3b8));font-size:13px;';
    etaHint.textContent = '预计完成时间：计算中...';

    if (processingStatus?.parentNode) {
        processingStatus.parentNode.insertBefore(etaHint, processingStatus.nextSibling);
    } else {
        card.appendChild(etaHint);
    }

    return etaHint;
}

function updateEtaHint(task) {
    const el = ensureEtaHint();
    if (!el) {
        return;
    }

    const text = buildEtaText(task);
    if (!text) {
        el.style.display = 'none';
        el.textContent = '';
        return;
    }

    el.style.display = 'block';
    el.textContent = text;
}

function buildEtaText(task) {
    if (!task) {
        return '预计完成时间：计算中...';
    }

    if (task.status === 'failed' || task.status === 'cancelled') {
        return '';
    }

    if (task.status === 'done' && task.finished_at) {
        return `预计完成时间：${formatEtaMinute(task.finished_at)}`;
    }

    if (task.status === 'queued') {
        return '预计完成时间：排队中，开始处理后计算';
    }

    const progress = Number(task.progress ?? 0);
    if (!Number.isFinite(progress) || progress <= 0 || progress >= 100 || !task.created_at) {
        return '预计完成时间：计算中...';
    }

    const createdAt = parseServerTime(task.created_at);
    if (Number.isNaN(createdAt.getTime())) {
        return '预计完成时间：计算中...';
    }

    const elapsedMs = Date.now() - createdAt.getTime();
    if (elapsedMs <= 0) {
        return '预计完成时间：计算中...';
    }

    const estimatedFinishedAt = new Date(createdAt.getTime() + elapsedMs / (progress / 100));
    return `预计完成时间：${formatEtaDate(estimatedFinishedAt)}`;
}

function parseServerTime(iso) {
    if (!iso) {
        return new Date(NaN);
    }

    const normalized = /([zZ]|[+\-]\d{2}:\d{2})$/.test(iso) ? iso : `${iso}Z`;
    return new Date(normalized);
}

function formatEtaMinute(iso) {
    const date = parseServerTime(iso);
    if (Number.isNaN(date.getTime())) {
        return '-';
    }

    return formatEtaDate(date);
}

function formatEtaDate(date) {
    const parts = new Intl.DateTimeFormat('zh-CN', {
        timeZone: ETA_TIME_ZONE,
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
    }).formatToParts(date);

    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return `${values.month}-${values.day} ${values.hour}:${values.minute}`;
}

function syncLog(text) {
    if (!text) {
        return;
    }

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

    resultStats.innerHTML = `
        <div class="stat-card">
            <i class="fas fa-file-image"></i>
            <h3>${escapeHtml(result.filename || '-')}</h3>
            <p>源文件</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml((modelConfig[result.model] || {}).label || result.model || '-')}</h3>
            <p>识别模型</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-route"></i>
            <h3>${escapeHtml((routeConfig[result.gemini_route] || {}).label || result.gemini_route || '-')}</h3>
            <p>调用线路</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-check-circle"></i>
            <h3>成功</h3>
            <p>任务状态</p>
        </div>
    `;

    if (!result.output_docx) {
        resultFileList.innerHTML = '<div class="result-file-empty">当前任务没有可下载文件。</div>';
        return;
    }

    resultFileList.innerHTML = `
        <div class="result-file-item">
            <div class="result-file-meta">
                <strong>营业执照译文.docx</strong>
                <span>最终 Word 译文</span>
            </div>
            <a class="download-btn" href="/${result.output_docx}" download>
                <i class="fas fa-download"></i> 下载
            </a>
        </div>
    `;
}

function resetPage() {
    clearFile();
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

function formatFileSize(size) {
    if (size < 1024) {
        return `${size} B`;
    }

    if (size < 1024 * 1024) {
        return `${(size / 1024).toFixed(1)} KB`;
    }

    return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

function escapeHtml(value) {
    return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
