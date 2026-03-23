let selectedFile = null;
let pollingTimer = null;
let modelConfig = {};
let routeConfig = {};
let defaultModel = 'google/gemini-3-flash-preview';
let defaultRoute = 'google';

const MODEL_DISPLAY_NAMES = {
    'gemini-3.1-flash-lite-preview': '极速版V2',
    'google/gemini-3.1-flash-lite-preview': '极速版V2',
    'google/gemini-3-flash-preview': '快速版V2',
    'google/gemini-3.1-pro-preview': '增强版V2',
    'Google Gemini 3 Flash Preview': '快速版V2',
    'Google Gemini 3.1 Pro Preview': '增强版V2',
};

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview = document.getElementById('filePreview');
const fileName = document.getElementById('fileName');
const btnRemove = document.getElementById('btnRemove');
const btnProcess = document.getElementById('btnProcess');
const btnNewTask = document.getElementById('btnNewTask');
const modelSelect = document.getElementById('modelSelect');
let geminiRouteSelect = document.getElementById('geminiRouteSelect');
const modelLabel = document.getElementById('modelLabel');
const modelDesc = document.getElementById('modelDesc');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

let streamLogWrap = null;
let streamLogEl = null;
let retryBtn = null;

init();

async function init() {
    ensureGeminiRouteSelect();
    ensureLogPanel();
    bindEvents();
    await loadConfig();
    updateModelInfo();
}

function ensureGeminiRouteSelect() {
    if (geminiRouteSelect) return;
    const panel = document.querySelector('.options-panel');
    if (!panel) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'option-group option-card';
    wrapper.innerHTML = [
        '<label for="geminiRouteSelect">路线切换</label>',
        '<div class="field-wrap">',
        '<i class="fas fa-route"></i>',
        '<select id="geminiRouteSelect"></select>',
        '</div>',
    ].join('');
    panel.insertBefore(wrapper, panel.children[1] || null);
    geminiRouteSelect = document.getElementById('geminiRouteSelect');
}

function ensureLogPanel() {
    streamLogWrap = document.getElementById('streamLogWrap');
    streamLogEl = document.getElementById('streamLog');
    if (streamLogWrap && streamLogEl) return;

    const card = processingSection?.querySelector('.processing-card');
    if (!card) return;

    streamLogWrap = document.createElement('div');
    streamLogWrap.id = 'streamLogWrap';
    streamLogWrap.className = 'stream-log-wrap';
    streamLogWrap.style.display = 'none';
    streamLogWrap.innerHTML = [
        '<div class="stream-log-head"><i class="fas fa-terminal"></i><span>后端日志</span></div>',
        '<pre id="streamLog" class="stream-log"></pre>',
    ].join('');
    card.appendChild(streamLogWrap);
    streamLogEl = document.getElementById('streamLog');
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
    modelSelect?.addEventListener('change', updateModelInfo);
}

async function loadConfig() {
    try {
        const response = await fetch('/task/pdf2docx/config');
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
        routeConfig = {
            google: { label: '线路1' },
            openrouter: { label: '线路2' },
        };
        modelConfig = {
            'gemini-3.1-flash-lite-preview': {
                label: '极速版V2',
                description: '更轻量的极速 OCR 模型，适合追求速度的 PDF / 图片转 Word 场景。',
            },
            'google/gemini-3-flash-preview': {
                label: '快速版V2',
                description: '速度更快，适合常规 PDF / 图片转 Word 场景。',
            },
            'google/gemini-3.1-pro-preview': {
                label: '增强版V2',
                description: '更强的复杂版面理解能力，适合高难度文档。',
            },
        };
    }

    renderModels();
    renderRoutes();
}

function renderModels() {
    modelSelect.innerHTML = '';
    Object.entries(modelConfig).forEach(([value, info]) => {
        modelSelect.add(new Option(getModelDisplayName(info.label || value), value));
    });
    modelSelect.value = modelConfig[defaultModel] ? defaultModel : Object.keys(modelConfig)[0];
}

function renderRoutes() {
    geminiRouteSelect.innerHTML = '';
    Object.entries(routeConfig).forEach(([value, info]) => {
        geminiRouteSelect.add(new Option(info.label || value, value));
    });
    geminiRouteSelect.value = routeConfig[defaultRoute] ? defaultRoute : Object.keys(routeConfig)[0];
}

function updateModelInfo() {
    const model = modelSelect.value;
    const info = modelConfig[model] || {};
    modelLabel.textContent = getModelDisplayName(info.label || model);
    modelDesc.textContent = info.description || '';
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
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(event) {
    const files = event.target.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFile(file) {
    const allowedExtensions = ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'];
    const lowerName = file.name.toLowerCase();
    const matched = allowedExtensions.some((ext) => lowerName.endsWith(ext));
    if (!matched) {
        alert('仅支持 PDF、PNG、JPG、JPEG、BMP、GIF、WEBP 文件。');
        return;
    }

    selectedFile = file;
    fileName.textContent = `${file.name} (${formatFileSize(file.size)})`;
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    btnProcess.disabled = false;
}

function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    btnProcess.disabled = true;
}

async function processFile() {
    if (!selectedFile) return;

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    removeRetryButton();
    updateProgress(5, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('file', selectedFile);

        const params = new URLSearchParams({
            model: modelSelect.value,
            gemini_route: geminiRouteSelect.value,
        });

        const response = await fetch(`/task/pdf2docx?${params.toString()}`, {
            method: 'POST',
            body: formData,
        });

        if (!response.ok) {
            const errorText = await safeReadError(response);
            throw new Error(errorText || `提交失败: ${response.status}`);
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
    pollingTimer = setInterval(() => pollStatus(taskId), 2000);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
}

async function pollStatus(taskId) {
    try {
        const response = await fetch(`/task/pdf2docx/status/${taskId}`);
        if (!response.ok) {
            throw new Error(`状态查询失败: ${response.status}`);
        }

        const data = await response.json();
        updateProgress(data.progress || 0, data.message || '正在处理中...');
        syncLog(data.stream_log || data.result?.stream_log || '');

        if (data.status === 'done') {
            stopPolling();
            showResult(data.result || {});
        } else if (data.status === 'failed') {
            stopPolling();
            showFailure(data.error || '转换失败', data.stream_log || '');
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

function syncLog(logText) {
    if (!streamLogWrap || !streamLogEl || !logText) return;
    streamLogWrap.style.display = 'block';
    streamLogEl.textContent = logText;
    streamLogEl.scrollTop = streamLogEl.scrollHeight;
}

function clearLog() {
    if (!streamLogWrap || !streamLogEl) return;
    streamLogWrap.style.display = 'none';
    streamLogEl.textContent = '';
}

function showFailure(message, logText = '') {
    processingSection.style.display = 'block';
    syncLog(logText);
    updateProgress(100, message || '处理失败');
    processingStatus.textContent = message || '处理失败';
    removeRetryButton();

    retryBtn = document.createElement('button');
    retryBtn.className = 'btn-secondary';
    retryBtn.style.marginTop = '18px';
    retryBtn.innerHTML = '<i class="fas fa-rotate-right"></i> 重新开始';
    retryBtn.addEventListener('click', resetPage);
    processingSection.querySelector('.processing-card')?.appendChild(retryBtn);
}

function removeRetryButton() {
    if (retryBtn) {
        retryBtn.remove();
        retryBtn = null;
    }
}

function showResult(result) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';

    resultStats.innerHTML = `
        <div class="stat-card">
            <i class="fas fa-file-alt"></i>
            <h3>${escapeHtml(result.filename || '-')}</h3>
            <p>源文件</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml(getModelDisplayName(result.model || '-'))}</h3>
            <p>使用模型</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-file-word"></i>
            <h3>DOCX</h3>
            <p>输出格式</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-check-circle"></i>
            <h3>成功</h3>
            <p>任务状态</p>
        </div>
    `;

    resultGrid.innerHTML = `
        <div class="result-item">
            <h3>下载结果</h3>
            <div class="download-links">
                <a href="/${result.output_docx}" download class="download-btn">
                    <i class="fas fa-file-word"></i> 下载 Word 文档
                </a>
                <a href="/${result.raw_output_txt}" download class="download-btn">
                    <i class="fas fa-file-lines"></i> 下载原始 OCR 文本
                </a>
            </div>
        </div>
    `;
}

function resetPage() {
    clearFile();
    stopPolling();
    clearLog();
    removeRetryButton();
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
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
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

function getModelDisplayName(name) {
    return MODEL_DISPLAY_NAMES[name] || name;
}
