const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const singleFileInput = document.getElementById('singleFile');
const modeDoubleRadio = document.getElementById('modeDouble');
const modeSingleRadio = document.getElementById('modeSingle');
const modeHint = document.getElementById('modeHint');
const uploadDesc = document.getElementById('uploadDesc');
const pageSubtitle = document.getElementById('pageSubtitle');
const modelSelect = document.getElementById('modelSelect');
const modelDesc = document.getElementById('modelDesc');
const btnRunCheck = document.getElementById('btnRunCheck');
const btnReset = document.getElementById('btnReset');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const resultSummary = document.getElementById('resultSummary');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

const doubleFileFields = document.getElementById('doubleFileFields');
const singleFileFields = document.getElementById('singleFileFields');

const progressBar = document.getElementById('progressBar');
const progressPercent = document.getElementById('progressPercent');
const progressDetails = document.getElementById('progressDetails');
const processingTitle = document.getElementById('processingTitle');
const processingText = document.getElementById('processingText');

const POLL_INTERVAL = 1000;
const ETA_TIME_ZONE = 'Asia/Shanghai';
const MODEL_DISPLAY_NAMES = {
    'gemini-3-flash-preview': '快速版V2',
    'google/gemini-3-flash-preview': '快速版V2',
    'gemini-3.1-pro-preview': '增强版V2',
    'google/gemini-3.1-pro-preview': '增强版V2',
};

let pollingTimer = null;
let modelConfig = {};
let defaultRoute = 'openrouter';
let defaultModel = 'gemini-3.1-pro-preview';
let defaultMode = 'double';
let currentMode = 'double';
let modeConfig = {};
let singleFileExtensions = ['.docx', '.doc', '.pdf', '.xlsx', '.pptx'];
let doubleFileExtensions = ['.docx', '.doc'];
let streamLogWrap = null;
let streamLogEl = null;
let etaHint = null;

init();

async function init() {
    ensureLogPanel();
    ensureEtaHint();
    bindEvents();
    await loadConfig();
    applyMode(defaultMode);
}

function bindEvents() {
    btnRunCheck?.addEventListener('click', runNumberCheck);
    btnReset?.addEventListener('click', resetPage);
    modeDoubleRadio?.addEventListener('change', () => applyMode('double'));
    modeSingleRadio?.addEventListener('change', () => applyMode('single'));
    modelSelect?.addEventListener('change', updateModelInfo);
}

async function loadConfig() {
    try {
        const resp = await fetch('/task/number-check/config');
        if (!resp.ok) throw new Error(`配置加载失败: ${resp.status}`);
        const data = await resp.json();
        modelConfig = data.models || {};
        defaultRoute = data.default_route || defaultRoute;
        defaultModel = data.default_model || defaultModel;
        defaultMode = data.default_mode || defaultMode;
        modeConfig = data.modes || {};
        singleFileExtensions = data.single_file_extensions || singleFileExtensions;
        doubleFileExtensions = data.double_file_extensions || doubleFileExtensions;
    } catch (error) {
        console.error(error);
        modelConfig = {
            'google/gemini-3-flash-preview': {
                label: '快速版V2',
                description: '速度更快，适合常规数字专检场景。',
            },
            'gemini-3.1-pro-preview': {
                label: '增强版V2',
                description: '推理更强，适合复杂编号和上下文判断场景。',
            },
        };
        modeConfig = {
            double: {
                label: '双文件模式',
                description: '上传原文和译文两个 DOC / DOCX 文件，输出修订版译文。',
            },
            single: {
                label: '单文件模式',
                description: '上传一个双语对照文件；DOC / DOCX 可生成修订版。',
            },
        };
    }

    singleFileInput.accept = singleFileExtensions.join(',');
    originalFileInput.accept = doubleFileExtensions.join(',');
    translatedFileInput.accept = doubleFileExtensions.join(',');
    renderModels();
    updateModelInfo();
}

function renderModels() {
    if (!modelSelect) return;
    modelSelect.innerHTML = '';
    Object.entries(modelConfig).forEach(([value, info]) => {
        modelSelect.add(new Option(getModelDisplayName(info.label || value), value));
    });
    const fallback = Object.keys(modelConfig)[0] || defaultModel;
    modelSelect.value = modelConfig[defaultModel] ? defaultModel : fallback;
}

function updateModelInfo() {
    if (!modelDesc) return;
    const currentModel = modelSelect?.value || defaultModel;
    const info = modelConfig[currentModel] || {};
    modelDesc.textContent = info.description || '';
}

function getSelectedMode() {
    if (modeSingleRadio?.checked) return 'single';
    return 'double';
}

function applyMode(mode) {
    currentMode = mode === 'single' ? 'single' : 'double';
    if (modeDoubleRadio) modeDoubleRadio.checked = currentMode === 'double';
    if (modeSingleRadio) modeSingleRadio.checked = currentMode === 'single';

    const singleMode = currentMode === 'single';
    doubleFileFields.style.display = singleMode ? 'none' : 'block';
    singleFileFields.style.display = singleMode ? 'block' : 'none';

    const currentModeConfig = modeConfig[currentMode] || {};
    const description = currentModeConfig.description || '';
    modeHint.textContent = description;
    pageSubtitle.textContent = singleMode
        ? '上传一个双语对照文件，自动检查数值并在 DOC / DOCX 输出修订版'
        : '上传原文与译文两个 DOC / DOCX，自动对比并批注修复';
    uploadDesc.textContent = singleMode
        ? `当前为单文件模式，支持 ${singleFileExtensions.join(' / ')}`
        : `当前为双文件模式，支持 ${doubleFileExtensions.join(' / ')}`;
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

function ensureEtaHint() {
    if (etaHint && etaHint.isConnected) return etaHint;
    const card = processingSection?.querySelector('.processing-card') || processingSection;
    if (!card) return null;
    etaHint = document.createElement('div');
    etaHint.className = 'eta-hint';
    etaHint.style.cssText = 'margin-top:10px;color:var(--text-secondary, var(--muted, #94a3b8));font-size:13px;';
    etaHint.textContent = '预计完成时间：计算中...';
    const anchor = processingText || null;
    if (anchor?.parentNode) {
        anchor.parentNode.insertBefore(etaHint, anchor.nextSibling);
    } else {
        card.appendChild(etaHint);
    }
    return etaHint;
}

function updateEtaHint(task) {
    const el = ensureEtaHint();
    if (!el) return;
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
    if (!task) return '预计完成时间：计算中...';
    if (task.status === 'failed' || task.status === 'cancelled') return '';
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
    if (Number.isNaN(createdAt.getTime())) return '预计完成时间：计算中...';

    const elapsedMs = Date.now() - createdAt.getTime();
    if (elapsedMs <= 0) return '预计完成时间：计算中...';

    const estimatedTotalMs = elapsedMs / (progress / 100);
    const estimatedFinishedAt = new Date(createdAt.getTime() + estimatedTotalMs);
    return `预计完成时间：${formatEtaDate(estimatedFinishedAt)}`;
}

function parseServerTime(iso) {
    if (!iso) return new Date(NaN);
    const normalized = /([zZ]|[+\-]\d{2}:\d{2})$/.test(iso) ? iso : `${iso}Z`;
    return new Date(normalized);
}

function formatEtaMinute(iso) {
    const date = parseServerTime(iso);
    if (Number.isNaN(date.getTime())) return '-';
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

async function runNumberCheck() {
    const mode = getSelectedMode();
    const formData = new FormData();

    if (mode === 'single') {
        const singleFile = singleFileInput.files[0];
        if (!singleFile) {
            alert('请选择一个双语对照文件');
            return;
        }
        formData.append('single_file', singleFile);
    } else {
        const originalFile = originalFileInput.files[0];
        const translatedFile = translatedFileInput.files[0];
        if (!originalFile || !translatedFile) {
            alert('请同时选择原文和译文 DOCX 文件');
            return;
        }
        formData.append('original_file', originalFile);
        formData.append('translated_file', translatedFile);
    }

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    updateProgressUI(0, '正在提交任务...');

    try {
        const params = new URLSearchParams({
            mode,
            gemini_route: defaultRoute,
            model_name: modelSelect?.value || defaultModel,
        });

        const resp = await fetch(`/task/number-check?${params.toString()}`, {
            method: 'POST',
            body: formData,
        });

        if (!resp.ok) {
            const detailMsg = await safeReadError(resp);
            throw new Error(detailMsg || `请求失败: ${resp.status}`);
        }

        const submitResp = await resp.json();
        if (submitResp.status === 'ACCEPTED' && submitResp.task_id) {
            updateProgressUI(5, '任务已提交，正在后台处理...');
            startPolling(submitResp.task_id);
            return;
        }

        showResult(submitResp);
    } catch (error) {
        showFailure(`数字专检失败: ${error.message}`);
    }
}

function updateProgressUI(progress, message, details = [], task = null) {
    progressBar.style.setProperty('--progress', `${progress}%`);
    progressPercent.textContent = `${progress}%`;
    processingTitle.textContent = message || '数字专检处理中...';
    processingText.textContent = message || '正在处理...';
    updateEtaHint(task);

    if (details?.length) {
        progressDetails.innerHTML = details.map((item) => `<div class="detail-item">${escapeHtml(item)}</div>`).join('');
    } else {
        progressDetails.innerHTML = `<div class="detail-item">${escapeHtml(message || '正在处理...')}</div>`;
    }
}

async function pollTaskStatus(taskId) {
    try {
        const resp = await fetch(`/task/number-check/status/${taskId}`);
        if (!resp.ok) throw new Error(`获取任务状态失败: ${resp.status}`);

        const status = await resp.json();
        updateProgressUI(status.progress || 0, status.message || '正在处理...', status.details || [], status);
        syncLog(status.stream_log || status.result?.stream_log || '');

        if (status.status === 'done') {
            stopPolling();
            showResult(status.result || status);
        } else if (status.status === 'failed') {
            stopPolling();
            showFailure(`数字专检失败: ${status.error || '未知错误'}`, status.stream_log || '');
        }
        return status;
    } catch (error) {
        stopPolling();
        showFailure(error.message);
        return null;
    }
}

function startPolling(taskId) {
    stopPolling();
    pollTaskStatus(taskId);
    pollingTimer = setInterval(() => {
        pollTaskStatus(taskId);
    }, POLL_INTERVAL);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
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
    updateProgressUI(100, message || '处理失败');
}

function showResult(data) {
    processingSection.style.display = 'none';
    uploadSection.style.display = 'none';
    resultSection.style.display = 'block';

    const mode = data.mode || currentMode;
    const reportCounts = data.report_counts || {};
    const stats = data.stats || {};
    const fixStats = data.fix_stats || {};
    const totalIssues = Number(
        stats.total_issues
        ?? ((reportCounts.body_issues || 0) + (reportCounts.header_issues || 0) + (reportCounts.footer_issues || 0))
    );

    resultSummary.textContent = data.summary || (mode === 'single' ? '单文件检查已完成。' : '双文件检查已完成。');
    resultStats.innerHTML = buildStatsHtml(mode, stats, fixStats, reportCounts, totalIssues, data.model_name || defaultModel);
    resultGrid.innerHTML = buildResultHtml(data, mode);
}

function buildStatsHtml(mode, stats, fixStats, reportCounts, totalIssues, modelName) {
    if (mode === 'single') {
        const issueTotal = stats.total_issues ?? totalIssues;
        const bodyIssues = stats.body_issues ?? reportCounts.body_issues ?? 0;
        const headerIssues = stats.header_issues ?? reportCounts.header_issues ?? 0;
        const footerIssues = stats.footer_issues ?? reportCounts.footer_issues ?? 0;
        const fixSuccess = fixStats.success ?? 0;
        return `
            <div class="stat-card">
                <i class="fas fa-list-check"></i>
                <h3>${issueTotal}</h3>
                <p>问题总数</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-file-lines"></i>
                <h3>${bodyIssues}</h3>
                <p>正文问题</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-heading"></i>
                <h3>${headerIssues}</h3>
                <p>页眉问题</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-shoe-prints"></i>
                <h3>${footerIssues}</h3>
                <p>页脚问题</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-wand-magic-sparkles"></i>
                <h3>${fixSuccess}</h3>
                <p>已修订条数</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-robot"></i>
                <h3>${escapeHtml(getModelDisplayName(modelName))}</h3>
                <p>模型</p>
            </div>
        `;
    }

    return `
        <div class="stat-card">
            <i class="fas fa-check"></i>
            <h3>${stats.success ?? 0}</h3>
            <p>成功</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-times"></i>
            <h3>${stats.failed ?? 0}</h3>
            <p>失败</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-forward"></i>
            <h3>${stats.skipped ?? 0}</h3>
            <p>跳过</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-file-lines"></i>
            <h3>${(reportCounts.body_issues || 0) + (reportCounts.header_issues || 0) + (reportCounts.footer_issues || 0)}</h3>
            <p>报告问题数</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml(getModelDisplayName(modelName))}</h3>
            <p>模型</p>
        </div>
    `;
}

function buildResultHtml(data, mode) {
    const links = [];
    if (data.corrected_docx) {
        links.push(
            `<a href="/${data.corrected_docx}" download class="download-btn"><i class="fas fa-file-word"></i> ${mode === 'single' ? '下载修订版文档' : '下载修订版译文'}</a>`
        );
    }

    const reports = data.reports || {};
    if (reports.body_json) {
        links.push(`<a href="/${reports.body_json}" download class="download-btn"><i class="fas fa-file-code"></i> 正文报告 JSON</a>`);
    }
    if (reports.header_json) {
        links.push(`<a href="/${reports.header_json}" download class="download-btn"><i class="fas fa-file-code"></i> 页眉报告 JSON</a>`);
    }
    if (reports.footer_json) {
        links.push(`<a href="/${reports.footer_json}" download class="download-btn"><i class="fas fa-file-code"></i> 页脚报告 JSON</a>`);
    }

    if (!links.length) {
        links.push('<div class="detail-item">当前没有可下载输出文件。</div>');
    }

    return `
        <div class="result-item">
            <h3>输出文件</h3>
            <div class="download-links">
                ${links.join('')}
            </div>
        </div>
    `;
}

function resetPage() {
    originalFileInput.value = '';
    translatedFileInput.value = '';
    singleFileInput.value = '';
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    clearLog();
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
    resultSummary.textContent = '';
    resultStats.innerHTML = '';
    resultGrid.innerHTML = '';
    if (modelSelect) {
        modelSelect.value = modelConfig[defaultModel] ? defaultModel : (Object.keys(modelConfig)[0] || defaultModel);
    }
    updateModelInfo();
    applyMode(defaultMode);
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

function getModelDisplayName(name) {
    return MODEL_DISPLAY_NAMES[name] || modelConfig[name]?.label || name;
}
