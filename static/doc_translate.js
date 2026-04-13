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
const btnBatchDownloadAll = document.getElementById('btnBatchDownloadAll');
const modelSelect = document.getElementById('modelSelect');
let geminiRouteSelect = document.getElementById('geminiRouteSelect');
const sourceLangSelect = document.getElementById('sourceLangSelect');
const targetLangGroup = document.getElementById('targetLangGroup');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const batchSection = document.getElementById('batchSection');
const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const translationResults = document.getElementById('translationResults');

const fileListPanel = document.getElementById('fileListPanel');
const fileListItems = document.getElementById('fileListItems');
const fileCountEl = document.getElementById('fileCount');

let selectedFiles = [];
let pollingTimer = null;
let modelConfig = {};
let languageConfig = {};
let routeConfig = {};
let defaultModel = 'google/gemini-3-flash-preview';
let defaultRoute = 'openrouter';
let allowedExtensions = ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tif', '.tiff', '.doc', '.docx'];
let streamLogWrap = null;
let streamLogEl = null;
let retryBtn = null;
let batchPollingTimer = null;
let batchTaskStates = [];

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

const MODEL_DISPLAY_NAMES = {
    'google/gemini-3-flash-preview': '快速版V2',
    'google/gemini-3.1-pro-preview': '增强版V2',
};

init();

async function init() {
    ensureGeminiRouteSelect();
    ensureLogPanel();
    bindEvents();
    await loadConfig();
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
    btnBatchDownloadAll?.addEventListener('click', downloadAllBatchResults);
    sourceLangSelect?.addEventListener('change', renderTargetLanguages);
}

async function loadConfig() {
    try {
        const response = await fetch('/task/doc-translate/config');
        if (!response.ok) throw new Error(`配置加载失败: ${response.status}`);
        const data = await response.json();
        modelConfig = data.models || {};
        defaultModel = data.default_model || defaultModel;
        languageConfig = data.languages || {};
        routeConfig = data.routes || {};
        defaultRoute = data.default_route || defaultRoute;
        if (Array.isArray(data.allowed_extensions) && data.allowed_extensions.length > 0) {
            allowedExtensions = data.allowed_extensions.map((ext) => String(ext).toLowerCase());
        }
    } catch (error) {
        console.error(error);
        modelConfig = {
            'google/gemini-3-flash-preview': { label: '快速版V2', description: '速度更快，适合常规 OCR。' },
            'google/gemini-3.1-pro-preview': { label: '增强版V2', description: '复杂版面表现更稳。' },
        };
        routeConfig = { google: { label: '\u7ebf\u8def1' }, openrouter: { label: '\u7ebf\u8def2' } };
        languageConfig = { zh:{name:'中文'}, en:{name:'英文'}, ja:{name:'日文'}, ko:{name:'韩文'}, es:{name:'西班牙文'}, fr:{name:'法文'}, de:{name:'德文'}, ru:{name:'俄文'} };
    }
    renderModels();
    renderRoutes();
    renderSourceLanguages();
    renderTargetLanguages();
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
function renderSourceLanguages() {
    sourceLangSelect.innerHTML = '';
    Object.entries(languageConfig).forEach(([code, info]) => sourceLangSelect.add(new Option(info.name, code)));
    sourceLangSelect.value = languageConfig.zh ? 'zh' : Object.keys(languageConfig)[0];
}
function renderTargetLanguages() {
    targetLangGroup.innerHTML = '';
    const sourceLang = sourceLangSelect.value;
    Object.entries(languageConfig).forEach(([code, info]) => {
        if (code === sourceLang) return;
        const chip = document.createElement('label');
        chip.className = 'lang-chip';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox'; checkbox.value = code;
        if (code === 'en') { checkbox.checked = true; chip.classList.add('active'); }
        const checkIcon = document.createElement('span');
        checkIcon.className = 'chip-check';
        checkIcon.innerHTML = '<i class="fas fa-check"></i>';
        const nameSpan = document.createElement('span');
        nameSpan.textContent = info.name;
        chip.appendChild(checkbox); chip.appendChild(checkIcon); chip.appendChild(nameSpan);
        chip.addEventListener('click', (e) => { e.preventDefault(); checkbox.checked = !checkbox.checked; chip.classList.toggle('active', checkbox.checked); updateProcessButton(); });
        targetLangGroup.appendChild(chip);
    });
    updateProcessButton();
}

function getSelectedTargetLangs() {
    return Array.from(targetLangGroup.querySelectorAll('input[type="checkbox"]:checked')).map((el) => el.value);
}
function updateProcessButton() {
    btnProcess.disabled = !(selectedFiles.length > 0 && getSelectedTargetLangs().length > 0);
}

function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); uploadArea.style.borderColor = 'var(--primary-color)'; }
function handleDrop(e) {
    e.preventDefault(); e.stopPropagation();
    uploadArea.style.borderColor = 'var(--border-color)';
    addFiles(Array.from(e.dataTransfer.files));
}
function handleFileSelect(e) { addFiles(Array.from(e.target.files)); }

function addFiles(newFiles) {
    for (const file of newFiles) {
        const lowerName = file.name.toLowerCase();
        if (!allowedExtensions.some((ext) => lowerName.endsWith(ext))) { alert(`不支持的文件格式：${file.name}`); continue; }
        if (!selectedFiles.some((f) => f.name === file.name && f.size === file.size)) selectedFiles.push(file);
    }
    if (selectedFiles.length === 0) return;
    fileName.textContent = selectedFiles.length === 1 ? `${selectedFiles[0].name} (${formatFileSize(selectedFiles[0].size)})` : `已选择 ${selectedFiles.length} 个文件`;
    if (fileStatus) fileStatus.textContent = selectedFiles.length === 1 ? '文件已就绪' : `共 ${selectedFiles.length} 个文件，将逐个创建任务队列处理`;
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    updateProcessButton();
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
    fileName.textContent = selectedFiles.length === 1 ? `${selectedFiles[0].name} (${formatFileSize(selectedFiles[0].size)})` : `已选择 ${selectedFiles.length} 个文件`;
    if (fileStatus) fileStatus.textContent = selectedFiles.length === 1 ? '文件已就绪' : `共 ${selectedFiles.length} 个文件，将逐个创建任务队列处理`;
    updateProcessButton();
    renderFileList();
}

function clearFiles() {
    selectedFiles = [];
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    fileListPanel.style.display = 'none';
    updateProcessButton();
}

async function processFiles() {
    if (selectedFiles.length === 0) return;
    const targetLangs = getSelectedTargetLangs();
    if (targetLangs.length === 0) { alert('请至少选择一种目标翻译语言。'); return; }
    if (selectedFiles.length === 1) { await processSingleFile(targetLangs); return; }
    await processBatchFiles(targetLangs);
}

async function processSingleFile(targetLangs) {
    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog(); removeRetryButton();
    updateProgress(5, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('file', selectedFiles[0]);
        const params = new URLSearchParams({ source_lang: sourceLangSelect.value, target_langs: targetLangs.join(','), ocr_model: modelSelect.value, gemini_route: geminiRouteSelect?.value || defaultRoute });
        const response = await fetch(`/task/doc-translate?${params.toString()}`, { method: 'POST', body: formData });
        if (!response.ok) { const t = await safeReadError(response); throw new Error(t || `提交失败: ${response.status}`); }
        const data = await response.json();
        startPolling(data.task_id);
    } catch (error) { showFailure(error.message); }
}

async function processBatchFiles(targetLangs) {
    uploadSection.style.display = 'none';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'block';

    const params = new URLSearchParams({ source_lang: sourceLangSelect.value, target_langs: targetLangs.join(','), ocr_model: modelSelect.value, gemini_route: geminiRouteSelect?.value || defaultRoute });
    const formData = new FormData();
    selectedFiles.forEach((file) => formData.append('files', file));

    try {
        const resp = await fetch(`/task/doc-translate/batch?${params}`, { method: 'POST', body: formData });
        if (!resp.ok) { let msg = `提交失败: ${resp.status}`; try { const ed = await resp.json(); msg = ed?.detail?.error || ed?.detail || msg; } catch (_) {} throw new Error(msg); }
        const data = await resp.json();
        startBatchPolling(data.tasks || []);
    } catch (error) { alert(`批量提交失败: ${error.message}`); resetPage(); }
}

function startBatchPolling(tasks) {
    stopBatchPolling();
    batchTaskStates = tasks.map((t) => ({
        filename: t.filename, task_id: t.task_id, submitStatus: t.status, error: t.error || null,
        status: t.status === 'ACCEPTED' ? 'queued' : 'failed', progress: 0,
        message: t.status === 'ACCEPTED' ? '排队中...' : (t.error || '提交失败'), result: null,
    }));
    renderBatchTaskList(batchTaskStates);
    batchPollingTimer = setInterval(async () => {
        let allDone = true;
        for (const task of batchTaskStates) {
            if (!task.task_id || ['done','failed','cancelled'].includes(task.status)) continue;
            allDone = false;
            try {
                const resp = await fetch(`/task/doc-translate/status/${task.task_id}`);
                if (!resp.ok) continue;
                const data = await resp.json();
                task.status = data.status === 'processing' ? 'running' : data.status;
                task.progress = data.progress || 0;
                task.message = data.message || '';
                if (data.status === 'done') { task.status = 'done'; task.progress = 100; task.result = data.result; }
                else if (data.status === 'failed') { task.status = 'failed'; task.message = data.error || '处理失败'; }
            } catch (_) {}
        }
        renderBatchTaskList(batchTaskStates);
        if (allDone || batchTaskStates.every((t) => ['done','failed','cancelled'].includes(t.status) || !t.task_id)) stopBatchPolling();
    }, 2500);
}

function stopBatchPolling() {
    if (batchPollingTimer) {
        clearInterval(batchPollingTimer);
        batchPollingTimer = null;
    }
}

function getBatchDownloadableTaskIds() {
    return batchTaskStates
        .filter((task) => task.status === 'done' && hasBatchDocx(task))
        .map((task) => task.task_id)
        .filter(Boolean);
}

function hasBatchDocx(task) {
    const translations = task?.result?.translations || {};
    return Object.values(translations).some((item) => item?.output_docx);
}

function updateBatchDownloadButton() {
    if (!btnBatchDownloadAll) return;
    const readyTaskIds = getBatchDownloadableTaskIds();
    if (!batchSection || batchSection.style.display === 'none' || readyTaskIds.length === 0) {
        btnBatchDownloadAll.style.display = 'none';
        btnBatchDownloadAll.disabled = true;
        btnBatchDownloadAll.innerHTML = '<i class="fas fa-box-archive"></i> 全部下载';
        return;
    }

    btnBatchDownloadAll.style.display = '';
    btnBatchDownloadAll.disabled = false;
    btnBatchDownloadAll.innerHTML = `<i class="fas fa-box-archive"></i> 全部下载（${readyTaskIds.length}）`;
}

function renderBatchTaskList(taskStates) {
    const doneCount = taskStates.filter((t) => t.status === 'done').length;
    const failedCount = taskStates.filter((t) => t.status === 'failed').length;
    const totalCount = taskStates.length;
    const remainingCount = Math.max(totalCount - doneCount - failedCount, 0);
    const allSettled = totalCount > 0 && remainingCount === 0;
    const progressIcon = allSettled
        ? '<i class="fas fa-check-circle" style="color:var(--success-color)"></i>'
        : '<i class="fas fa-spinner fa-spin" style="color:var(--primary-color)"></i>';
    const progressValue = allSettled ? totalCount : remainingCount;
    const progressLabel = allSettled ? (failedCount > 0 ? '已结束' : '全部完成') : '处理中';
    document.getElementById('batchStats').innerHTML = `
        <div class="stat-card"><i class="fas fa-files"></i><h3>${totalCount}</h3><p>总文件数</p></div>
        <div class="stat-card"><i class="fas fa-check-circle" style="color:var(--success-color)"></i><h3>${doneCount}</h3><p>已完成</p></div>
        <div class="stat-card">${progressIcon}<h3>${progressValue}</h3><p>${progressLabel}</p></div>
        ${failedCount > 0 ? `<div class="stat-card"><i class="fas fa-times-circle" style="color:var(--danger-color)"></i><h3>${failedCount}</h3><p>失败</p></div>` : ''}
    `;
    document.getElementById('batchTaskList').innerHTML = taskStates.map((task) => {
        const icon = task.status === 'done' ? '<i class="fas fa-check-circle" style="color:var(--success-color)"></i>'
            : task.status === 'failed' ? '<i class="fas fa-times-circle" style="color:var(--danger-color)"></i>'
            : task.status === 'running' ? '<i class="fas fa-spinner fa-spin" style="color:var(--primary-color)"></i>'
            : '<i class="fas fa-clock" style="color:var(--warning-color)"></i>';
        let dlLinks = '';
        if (task.status === 'done' && task.result) {
            const translations = task.result.translations || {};
            const items = Object.values(translations);
            dlLinks = items.filter((i) => i.output_docx).map((i) =>
                `<a href="/${i.output_docx}" download style="color:var(--success-color);font-size:12px;text-decoration:none;white-space:nowrap;"><i class="fas fa-download"></i> ${escapeHtml(i.lang_name || 'DOCX')}</a>`
            ).join(' ');
        }
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
            ${dlLinks ? `<div style="display:flex;gap:6px;flex-wrap:wrap;">${dlLinks}</div>` : ''}
        </div>`;
    }).join('');
    updateBatchDownloadButton();
}

async function downloadAllBatchResults() {
    const taskIds = getBatchDownloadableTaskIds();
    if (!taskIds.length) {
        alert('当前还没有可打包下载的已完成文档。');
        return;
    }

    if (btnBatchDownloadAll) {
        btnBatchDownloadAll.disabled = true;
        btnBatchDownloadAll.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 正在打包...';
    }

    try {
        const response = await fetch('/task/batch-download', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                task_ids: taskIds,
                extensions: ['.docx'],
                archive_name: '通用证件批量结果.zip',
            }),
        });
        if (!response.ok) {
            const message = await safeReadError(response);
            throw new Error(message || `打包失败: ${response.status}`);
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = '通用证件批量结果.zip';
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
    } catch (error) {
        alert(`全部下载失败: ${error.message}`);
    } finally {
        updateBatchDownloadButton();
    }
}

function startPolling(taskId) { stopPolling(); pollStatus(taskId); pollingTimer = setInterval(() => pollStatus(taskId), 2500); }
function stopPolling() { if (pollingTimer) { clearInterval(pollingTimer); pollingTimer = null; } }

async function pollStatus(taskId) {
    try {
        const response = await fetch(`/task/doc-translate/status/${taskId}`);
        if (!response.ok) throw new Error(`状态查询失败: ${response.status}`);
        const data = await response.json();
        updateProgress(data.progress || 0, data.message || '正在处理中...', data);
        syncLog(data.stream_log || data.result?.stream_log || '');
        if (data.status === 'done') { stopPolling(); showResult(data.result || {}); }
        else if (data.status === 'failed') { stopPolling(); showFailure(data.error || '翻译失败', data.stream_log || ''); }
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
    const translations = result.translations || {};
    const langCount = Object.keys(translations).length;
    resultStats.innerHTML = `
        <div class="stat-card"><i class="fas fa-file-alt"></i><h3>${escapeHtml(result.filename || '-')}</h3><p>源文件</p></div>
        <div class="stat-card"><i class="fas fa-robot"></i><h3>${escapeHtml(getModelDisplayName(result.ocr_model || '-'))}</h3><p>OCR 模型</p></div>
        <div class="stat-card"><i class="fas fa-globe"></i><h3>${langCount} 种语言</h3><p>翻译输出</p></div>
        <div class="stat-card"><i class="fas fa-check-circle"></i><h3>成功</h3><p>任务状态</p></div>
    `;
    let html = '';
    if (result.raw_output_txt) {
        html += `<div class="translation-result-item"><div style="display:flex;align-items:center;gap:12px;"><span class="lang-badge"><i class="fas fa-file-lines"></i> 原始文本</span><span style="color:var(--text-secondary);font-size:13px;">OCR 识别结果</span></div><div class="download-actions"><a href="/${result.raw_output_txt}" download class="download-btn"><i class="fas fa-download"></i> 下载 TXT</a></div></div>`;
    }
    Object.entries(translations).forEach(([langCode, langResult]) => {
        html += `<div class="translation-result-item"><div style="display:flex;align-items:center;gap:12px;"><span class="lang-badge"><i class="fas fa-language"></i> ${escapeHtml(langResult.lang_name || langCode)}</span></div><div class="download-actions"><a href="/${langResult.output_docx}" download class="download-btn"><i class="fas fa-file-word"></i> 下载 Word</a><a href="/${langResult.translated_txt}" download class="download-btn"><i class="fas fa-file-lines"></i> 下载译文文本</a></div></div>`;
    });
    translationResults.innerHTML = html;
}

function resetPage() {
    clearFiles();
    stopPolling();
    stopBatchPolling();
    batchTaskStates = [];
    clearLog();
    removeRetryButton();
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'none';
    updateProgress(0, '等待开始处理');
    updateBatchDownloadButton();
}

async function safeReadError(response) { try { const p = await response.json(); return p?.detail?.error || p?.detail || p?.message || ''; } catch (_) { return ''; } }
function formatFileSize(size) { if (size < 1024) return `${size} B`; if (size < 1024*1024) return `${(size/1024).toFixed(1)} KB`; return `${(size/1024/1024).toFixed(1)} MB`; }
function escapeHtml(v) { return String(v).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;'); }
function getModelDisplayName(name) { return MODEL_DISPLAY_NAMES[name] || name; }
