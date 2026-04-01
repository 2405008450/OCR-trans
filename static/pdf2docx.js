let selectedFiles = [];
let pollingTimer = null;
let modelConfig = {};
let routeConfig = {};
let defaultModel = 'google/gemini-3-flash-preview';
let defaultRoute = 'openrouter';

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
const fileStatus = document.getElementById('fileStatus');
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
const batchSection = document.getElementById('batchSection');
const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

const fileListPanel = document.getElementById('fileListPanel');
const fileListItems = document.getElementById('fileListItems');
const fileCountEl = document.getElementById('fileCount');

let streamLogWrap = null;
let streamLogEl = null;
let retryBtn = null;
let batchPollingTimer = null;

const ETA_TIME_ZONE = 'Asia/Shanghai';
let etaHint = null;

function ensureEtaHint() {
    if (etaHint && etaHint.isConnected) return etaHint;
    const card = processingSection?.querySelector('.processing-card') || processingSection;
    if (!card) return null;
    etaHint = document.createElement('div');
    etaHint.className = 'eta-hint';
    etaHint.style.cssText = 'margin-top:10px;color:var(--text-secondary, var(--muted, #94a3b8));font-size:13px;';
    etaHint.textContent = '预计完成时间：计算中...';
    const anchor = processingStatus || null;
    if (anchor?.parentNode) anchor.parentNode.insertBefore(etaHint, anchor.nextSibling);
    else card.appendChild(etaHint);
    return etaHint;
}

function updateEtaHint(task) {
    const el = ensureEtaHint();
    if (!el) return;
    const text = buildEtaText(task);
    if (!text) { el.style.display = 'none'; el.textContent = ''; return; }
    el.style.display = 'block';
    el.textContent = text;
}

function buildEtaText(task) {
    if (!task) return '预计完成时间：计算中...';
    if (task.status === 'failed' || task.status === 'cancelled') return '';
    if (task.status === 'done' && task.finished_at) return `预计完成时间：${formatEtaMinute(task.finished_at)}`;
    if (task.status === 'queued') return '预计完成时间：排队中，开始处理后计算';
    const progress = Number(task.progress ?? 0);
    if (!Number.isFinite(progress) || progress <= 0 || progress >= 100 || !task.created_at) return '预计完成时间：计算中...';
    const createdAt = parseServerTime(task.created_at);
    if (Number.isNaN(createdAt.getTime())) return '预计完成时间：计算中...';
    const elapsedMs = Date.now() - createdAt.getTime();
    if (elapsedMs <= 0) return '预计完成时间：计算中...';
    const estimatedFinishedAt = new Date(createdAt.getTime() + (elapsedMs / (progress / 100)));
    return `预计完成时间：${formatEtaDate(estimatedFinishedAt)}`;
}

function parseServerTime(iso) {
    if (!iso) return new Date(NaN);
    const normalized = /([zZ]|[+\-]\d{2}:\d{2})$/.test(iso) ? iso : `${iso}Z`;
    return new Date(normalized);
}
function formatEtaMinute(iso) { const d = parseServerTime(iso); return Number.isNaN(d.getTime()) ? '-' : formatEtaDate(d); }
function formatEtaDate(date) {
    const parts = new Intl.DateTimeFormat('zh-CN', { timeZone: ETA_TIME_ZONE, month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit', hour12:false }).formatToParts(date);
    const v = Object.fromEntries(parts.map((p) => [p.type, p.value]));
    return `${v.month}-${v.day} ${v.hour}:${v.minute}`;
}

init();

async function init() {
    ensureGeminiRouteSelect();
    ensureLogPanel();
    bindEvents();
    await loadConfig();
    updateModelInfo();
}

function ensureGeminiRouteSelect() {
    const routeCard = document.getElementById('geminiRouteGroup');
    if (routeCard) routeCard.style.display = 'none';
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
    streamLogWrap.innerHTML = '<div class="stream-log-head"><i class="fas fa-terminal"></i><span>后端日志</span></div><pre id="streamLog" class="stream-log"></pre>';
    card.appendChild(streamLogWrap);
    streamLogEl = document.getElementById('streamLog');
}

function bindEvents() {
    uploadArea?.addEventListener('click', () => fileInput?.click());
    uploadArea?.addEventListener('dragover', handleDragOver);
    uploadArea?.addEventListener('drop', handleDrop);
    fileInput?.addEventListener('change', handleFileSelect);
    btnRemove?.addEventListener('click', (e) => { e.stopPropagation(); clearFiles(); });
    btnProcess?.addEventListener('click', processFiles);
    btnNewTask?.addEventListener('click', resetPage);
    document.getElementById('btnBatchNewTask')?.addEventListener('click', resetPage);
    modelSelect?.addEventListener('change', updateModelInfo);
}

async function loadConfig() {
    try {
        const response = await fetch('/task/pdf2docx/config');
        if (!response.ok) throw new Error(`配置加载失败: ${response.status}`);
        const data = await response.json();
        modelConfig = data.models || {};
        routeConfig = data.routes || {};
        defaultModel = data.default_model || defaultModel;
        defaultRoute = data.default_route || defaultRoute;
    } catch (error) {
        console.error(error);
        routeConfig = { google: { label: '\u7ebf\u8def1' }, openrouter: { label: '\u7ebf\u8def2' } };
        modelConfig = {
            'gemini-3.1-flash-lite-preview': { label: '极速版V2', description: '更轻量的极速 OCR 模型。' },
            'google/gemini-3-flash-preview': { label: '快速版V2', description: '速度更快，适合常规场景。' },
            'google/gemini-3.1-pro-preview': { label: '增强版V2', description: '更强的复杂版面理解能力。' },
        };
    }
    renderModels();
    renderRoutes();
}

function renderModels() {
    modelSelect.innerHTML = '';
    Object.entries(modelConfig).forEach(([value, info]) => modelSelect.add(new Option(getModelDisplayName(info.label || value), value)));
    modelSelect.value = modelConfig[defaultModel] ? defaultModel : Object.keys(modelConfig)[0];
}
function renderRoutes() {
    if (!geminiRouteSelect) return;
    geminiRouteSelect.innerHTML = '';
    Object.entries(routeConfig).forEach(([value, info]) => geminiRouteSelect.add(new Option(info.label || value, value)));
    geminiRouteSelect.value = routeConfig[defaultRoute] ? defaultRoute : Object.keys(routeConfig)[0];
}
function updateModelInfo() {
    const model = modelSelect.value;
    const info = modelConfig[model] || {};
    modelLabel.textContent = getModelDisplayName(info.label || model);
    modelDesc.textContent = info.description || '';
}

function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); uploadArea.style.borderColor = 'var(--primary-color)'; }
function handleDrop(e) {
    e.preventDefault(); e.stopPropagation();
    uploadArea.style.borderColor = 'var(--border-color)';
    addFiles(Array.from(e.dataTransfer.files));
}
function handleFileSelect(e) { addFiles(Array.from(e.target.files)); }

function addFiles(newFiles) {
    const allowedExtensions = ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'];
    for (const file of newFiles) {
        const lowerName = file.name.toLowerCase();
        if (!allowedExtensions.some((ext) => lowerName.endsWith(ext))) {
            alert(`不支持的文件格式：${file.name}`);
            continue;
        }
        if (!selectedFiles.some((f) => f.name === file.name && f.size === file.size)) {
            selectedFiles.push(file);
        }
    }
    if (selectedFiles.length === 0) return;
    fileName.textContent = selectedFiles.length === 1
        ? `${selectedFiles[0].name} (${formatFileSize(selectedFiles[0].size)})`
        : `已选择 ${selectedFiles.length} 个文件`;
    if (fileStatus) fileStatus.textContent = selectedFiles.length === 1
        ? '文件已就绪，可开始转换'
        : `共 ${selectedFiles.length} 个文件，将逐个创建任务队列处理`;
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    btnProcess.disabled = false;
    renderFileList();
}

function renderFileList() {
    if (selectedFiles.length <= 1) { fileListPanel.style.display = 'none'; return; }
    fileListPanel.style.display = '';
    fileCountEl.textContent = selectedFiles.length;
    fileListItems.innerHTML = selectedFiles.map((file, idx) => `
        <div style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-radius:10px;background:linear-gradient(180deg,#f8fbfe,#f1f6fb);border:1px solid rgba(88,112,138,0.12);">
            <i class="fas fa-file" style="color:var(--primary-color);"></i>
            <span style="flex:1;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(file.name)}</span>
            <span style="font-size:12px;color:var(--text-secondary);">${formatFileSize(file.size)}</span>
            <button onclick="removeFileAt(${idx})" style="border:none;background:none;color:#ef4444;cursor:pointer;font-size:14px;padding:4px;" title="移除"><i class="fas fa-times"></i></button>
        </div>
    `).join('');
}

function removeFileAt(idx) {
    selectedFiles.splice(idx, 1);
    if (selectedFiles.length === 0) { clearFiles(); return; }
    fileName.textContent = selectedFiles.length === 1
        ? `${selectedFiles[0].name} (${formatFileSize(selectedFiles[0].size)})`
        : `已选择 ${selectedFiles.length} 个文件`;
    if (fileStatus) fileStatus.textContent = selectedFiles.length === 1
        ? '文件已就绪，可开始转换'
        : `共 ${selectedFiles.length} 个文件，将逐个创建任务队列处理`;
    renderFileList();
}

function clearFiles() {
    selectedFiles = [];
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    fileListPanel.style.display = 'none';
    btnProcess.disabled = true;
}

async function processFiles() {
    if (selectedFiles.length === 0) return;
    if (selectedFiles.length === 1) { await processSingleFile(); return; }
    await processBatchFiles();
}

async function processSingleFile() {
    const formData = new FormData();
    formData.append('file', selectedFiles[0]);
    const params = new URLSearchParams({ model: modelSelect.value, gemini_route: geminiRouteSelect?.value || defaultRoute });

    let response;
    try { response = await fetch(`/task/pdf2docx?${params.toString()}`, { method: 'POST', body: formData }); }
    catch (err) { showFailure(err.message); return; }

    if (!response.ok) {
        const detail = await safeReadError(response);
        showPageLimitModal(detail || `提交失败: ${response.status}`);
        return;
    }

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    removeRetryButton();
    updateProgress(5, '正在提交任务...');

    try {
        const data = await response.json();
        startPolling(data.task_id);
    } catch (error) { showFailure(error.message); }
}

async function processBatchFiles() {
    uploadSection.style.display = 'none';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'block';

    const params = new URLSearchParams({ model: modelSelect.value, gemini_route: geminiRouteSelect?.value || defaultRoute });
    const formData = new FormData();
    selectedFiles.forEach((file) => formData.append('files', file));

    try {
        const resp = await fetch(`/task/pdf2docx/batch?${params}`, { method: 'POST', body: formData });
        if (!resp.ok) {
            let msg = `提交失败: ${resp.status}`;
            try { const ed = await resp.json(); msg = ed?.detail?.error || ed?.detail || msg; } catch (_) {}
            throw new Error(msg);
        }
        const data = await resp.json();
        startBatchPolling(data.tasks || []);
    } catch (error) { alert(`批量提交失败: ${error.message}`); resetPage(); }
}

function startBatchPolling(tasks) {
    stopBatchPolling();
    const taskStates = tasks.map((t) => ({
        filename: t.filename, task_id: t.task_id, submitStatus: t.status, error: t.error || null,
        status: t.status === 'ACCEPTED' ? 'queued' : 'failed', progress: 0,
        message: t.status === 'ACCEPTED' ? '排队中...' : (t.error || '提交失败'), result: null,
    }));
    renderBatchTaskList(taskStates);
    batchPollingTimer = setInterval(async () => {
        let allDone = true;
        for (const task of taskStates) {
            if (!task.task_id || ['done','failed','cancelled'].includes(task.status)) continue;
            allDone = false;
            try {
                const resp = await fetch(`/task/pdf2docx/status/${task.task_id}`);
                if (!resp.ok) continue;
                const data = await resp.json();
                task.status = data.status === 'processing' ? 'running' : data.status;
                task.progress = data.progress || 0;
                task.message = data.message || '';
                if (data.status === 'done') { task.status = 'done'; task.progress = 100; task.result = data.result; }
                else if (data.status === 'failed') { task.status = 'failed'; task.message = data.error || '处理失败'; }
            } catch (_) {}
        }
        renderBatchTaskList(taskStates);
        if (allDone || taskStates.every((t) => ['done','failed','cancelled'].includes(t.status) || !t.task_id)) stopBatchPolling();
    }, 2500);
}

function stopBatchPolling() { if (batchPollingTimer) { clearInterval(batchPollingTimer); batchPollingTimer = null; } }

function renderBatchTaskList(taskStates) {
    const doneCount = taskStates.filter((t) => t.status === 'done').length;
    const failedCount = taskStates.filter((t) => t.status === 'failed').length;
    const totalCount = taskStates.length;
    document.getElementById('batchStats').innerHTML = `
        <div class="stat-card"><i class="fas fa-files"></i><h3>${totalCount}</h3><p>总文件数</p></div>
        <div class="stat-card"><i class="fas fa-check-circle" style="color:var(--success-color)"></i><h3>${doneCount}</h3><p>已完成</p></div>
        <div class="stat-card"><i class="fas fa-spinner fa-spin" style="color:var(--primary-color)"></i><h3>${totalCount - doneCount - failedCount}</h3><p>处理中</p></div>
        ${failedCount > 0 ? `<div class="stat-card"><i class="fas fa-times-circle" style="color:var(--danger-color)"></i><h3>${failedCount}</h3><p>失败</p></div>` : ''}
    `;
    document.getElementById('batchTaskList').innerHTML = taskStates.map((task) => {
        const icon = task.status === 'done' ? '<i class="fas fa-check-circle" style="color:var(--success-color)"></i>'
            : task.status === 'failed' ? '<i class="fas fa-times-circle" style="color:var(--danger-color)"></i>'
            : task.status === 'running' ? '<i class="fas fa-spinner fa-spin" style="color:var(--primary-color)"></i>'
            : '<i class="fas fa-clock" style="color:var(--warning-color)"></i>';
        const dl = task.status === 'done' && task.result?.output_docx
            ? `<a href="/${task.result.output_docx}" download style="color:var(--success-color);font-size:12px;text-decoration:none;white-space:nowrap;"><i class="fas fa-download"></i> DOCX</a>` : '';
        return `<div style="display:flex;align-items:center;gap:14px;padding:14px 18px;border-radius:16px;background:var(--card-bg);border:1px solid var(--border-color);box-shadow:var(--shadow-sm);">
            ${icon}
            <div style="flex:1;min-width:0;">
                <div style="font-size:14px;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(task.filename)}</div>
                <div style="font-size:12px;color:var(--text-secondary);margin-top:4px;">${escapeHtml(task.message)}</div>
            </div>
            <div style="min-width:60px;text-align:right;">
                <div style="font-size:13px;font-weight:700;">${task.progress}%</div>
                <div style="width:60px;height:4px;border-radius:2px;background:#e4edf5;margin-top:4px;overflow:hidden;">
                    <div style="width:${task.progress}%;height:100%;border-radius:inherit;background:linear-gradient(90deg,var(--primary-color),var(--success-color));transition:width .3s;"></div>
                </div>
            </div>
            ${dl ? `<div>${dl}</div>` : ''}
        </div>`;
    }).join('');
}

function showPageLimitModal(message) {
    const modal = document.getElementById('pagelimitModal');
    const msgEl = document.getElementById('pageLimit-msg');
    const closeBtn = document.getElementById('pageLimit-close');
    if (!modal || !msgEl || !closeBtn) return;
    msgEl.textContent = message;
    modal.style.display = 'flex';
    const close = () => { modal.style.display = 'none'; closeBtn.removeEventListener('click', close); };
    closeBtn.addEventListener('click', close);
}

function startPolling(taskId) { stopPolling(); pollStatus(taskId); pollingTimer = setInterval(() => pollStatus(taskId), 2000); }
function stopPolling() { if (pollingTimer) { clearInterval(pollingTimer); pollingTimer = null; } }

async function pollStatus(taskId) {
    try {
        const response = await fetch(`/task/pdf2docx/status/${taskId}`);
        if (!response.ok) throw new Error(`状态查询失败: ${response.status}`);
        const data = await response.json();
        updateProgress(data.progress || 0, data.message || '正在处理中...', data);
        syncLog(data.stream_log || data.result?.stream_log || '');
        if (data.status === 'done') { stopPolling(); showResult(data.result || {}); }
        else if (data.status === 'failed') { stopPolling(); showFailure(data.error || '转换失败', data.stream_log || ''); }
    } catch (error) { stopPolling(); showFailure(error.message); }
}

function updateProgress(percent, message, task = null) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;
    processingStatus.textContent = message;
    updateEtaHint(task);
}

function syncLog(logText) { if (!streamLogWrap || !streamLogEl || !logText) return; streamLogWrap.style.display = 'block'; streamLogEl.textContent = logText; streamLogEl.scrollTop = streamLogEl.scrollHeight; }
function clearLog() { if (!streamLogWrap || !streamLogEl) return; streamLogWrap.style.display = 'none'; streamLogEl.textContent = ''; }

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
function removeRetryButton() { if (retryBtn) { retryBtn.remove(); retryBtn = null; } }

function showResult(result) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';
    resultStats.innerHTML = `
        <div class="stat-card"><i class="fas fa-file-alt"></i><h3>${escapeHtml(result.filename || '-')}</h3><p>源文件</p></div>
        <div class="stat-card"><i class="fas fa-robot"></i><h3>${escapeHtml(getModelDisplayName(result.model || '-'))}</h3><p>使用模型</p></div>
        <div class="stat-card"><i class="fas fa-file-word"></i><h3>DOCX</h3><p>输出格式</p></div>
        <div class="stat-card"><i class="fas fa-check-circle"></i><h3>成功</h3><p>任务状态</p></div>
    `;
    resultGrid.innerHTML = `<div class="result-item"><h3>下载结果</h3><div class="download-links">
        <a href="/${result.output_docx}" download class="download-btn"><i class="fas fa-file-word"></i> 下载 Word 文档</a>
        <a href="/${result.raw_output_txt}" download class="download-btn"><i class="fas fa-file-lines"></i> 下载原始 OCR 文本</a>
    </div></div>`;
}

function resetPage() {
    clearFiles();
    stopPolling();
    stopBatchPolling();
    clearLog();
    removeRetryButton();
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'none';
    updateProgress(0, '等待开始处理');
}

async function safeReadError(response) { try { const p = await response.json(); return p?.detail?.error || p?.detail || p?.message || ''; } catch (_) { return ''; } }
function formatFileSize(size) { if (size < 1024) return `${size} B`; if (size < 1024*1024) return `${(size/1024).toFixed(1)} KB`; return `${(size/1024/1024).toFixed(1)} MB`; }
function escapeHtml(v) { return String(v).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;'); }
function getModelDisplayName(name) { return MODEL_DISPLAY_NAMES[name] || name; }
