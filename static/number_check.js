const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const btnRunCheck = document.getElementById('btnRunCheck');
const btnReset = document.getElementById('btnReset');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

const progressBar = document.getElementById('progressBar');
const progressPercent = document.getElementById('progressPercent');
const progressDetails = document.getElementById('progressDetails');
const processingTitle = document.getElementById('processingTitle');
const processingText = document.getElementById('processingText');

const POLL_INTERVAL = 1000;

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
    const anchor = typeof processingText !== 'undefined' && processingText ? processingText : null;
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
    if (Number.isNaN(createdAt.getTime())) {
        return '预计完成时间：计算中...';
    }

    const elapsedMs = Date.now() - createdAt.getTime();
    if (elapsedMs <= 0) {
        return '预计完成时间：计算中...';
    }

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
const MODEL_DISPLAY_NAMES = {
    'gemini-3-flash-preview': '快速版V2',
    'google/gemini-3-flash-preview': '快速版V2',
    'gemini-3.1-pro-preview': '增强版V2',
    'google/gemini-3.1-pro-preview': '增强版V2',
    'Gemini 3.1 Pro Preview': '增强版V2',
};

let pollingTimer = null;
let routeConfig = {};
let modelConfig = {};
let defaultRoute = 'openrouter';
let defaultModel = 'gemini-3.1-pro-preview';
let geminiRouteSelect = null;
let streamLogWrap = null;
let streamLogEl = null;

init();

async function init() {
    ensureOptionControls();
    ensureLogPanel();
    ensureEtaHint();
    bindEvents();
    await loadConfig();
}

function bindEvents() {
    btnRunCheck?.addEventListener('click', runNumberCheck);
    btnReset?.addEventListener('click', resetPage);
}

function ensureOptionControls() {
    const panel = document.querySelector('.options-panel');
    if (!panel) return;

    if (!document.getElementById('geminiRouteSelect')) {
        const routeGroup = document.createElement('div');
        routeGroup.className = 'option-group option-card';
        routeGroup.style.gridColumn = '1 / -1';
        routeGroup.innerHTML = [
            '<label for="geminiRouteSelect">路线切换</label>',
            '<div class="field-wrap">',
            '<i class="fas fa-route"></i>',
            '<select id="geminiRouteSelect"></select>',
            '</div>',
        ].join('');
        panel.appendChild(routeGroup);
    }

    geminiRouteSelect = document.getElementById('geminiRouteSelect');
    const routeGroup = geminiRouteSelect?.closest('.option-group');
    if (routeGroup) routeGroup.style.display = ''; 
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

async function loadConfig() {
    try {
        const resp = await fetch('/task/number-check/config');
        if (!resp.ok) throw new Error(`配置加载失败: ${resp.status}`);
        const data = await resp.json();
        routeConfig = data.routes || {};
        modelConfig = data.models || {};
        defaultRoute = data.default_route || defaultRoute;
        defaultModel = data.default_model || defaultModel;
    } catch (error) {
        console.error(error);
        routeConfig = {
            google: { label: '\u7ebf\u8def1' },
            openrouter: { label: '\u7ebf\u8def2' },
            google_ai_studio: { label: '\u7ebf\u8def3' },
        };
        modelConfig = {
            'gemini-3.1-pro-preview': {
                label: '增强版V2',
                description: '推理更强，适合复杂编号和上下文判断场景。',
            },
        };
    }

    renderRouteOptions();
    defaultModel = modelConfig['gemini-3.1-pro-preview']
        ? 'gemini-3.1-pro-preview'
        : (Object.keys(modelConfig)[0] || defaultModel);
}

function renderRouteOptions() {
    if (!geminiRouteSelect) return;
    geminiRouteSelect.innerHTML = '';
    Object.entries(routeConfig).forEach(([value, info]) => {
        geminiRouteSelect.add(new Option(info.label || value, value));
    });
    geminiRouteSelect.value = routeConfig[defaultRoute] ? defaultRoute : Object.keys(routeConfig)[0];
}

async function runNumberCheck() {
    const originalFile = originalFileInput.files[0];
    const translatedFile = translatedFileInput.files[0];

    if (!originalFile || !translatedFile) {
        alert('请同时选择原文和译文 .docx 文件');
        return;
    }

    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    updateProgressUI(0, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('original_file', originalFile);
        formData.append('translated_file', translatedFile);

        const params = new URLSearchParams({
            gemini_route: geminiRouteSelect?.value || defaultRoute,
            model_name: 'gemini-3.1-pro-preview',
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
        if (!resp.ok) {
            throw new Error(`获取任务状态失败: ${resp.status}`);
        }

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
    resultSection.style.display = 'block';

    const stats = data.stats || {};
    resultStats.innerHTML = `
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
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml(getModelDisplayName(data.model_name || defaultModel))}</h3>
            <p>模型</p>
        </div>
    `;

    const reports = data.reports || {};
    resultGrid.innerHTML = `
        <div class="result-item">
            <h3>输出文件</h3>
            <div class="download-links">
                <a href="/${data.corrected_docx}" download class="download-btn">
                    <i class="fas fa-file-word"></i> 下载修复后译文
                </a>
                ${reports.body_json ? `<a href="/${reports.body_json}" download class="download-btn"><i class="fas fa-file-code"></i> 正文报告 JSON</a>` : ''}
                ${reports.header_json ? `<a href="/${reports.header_json}" download class="download-btn"><i class="fas fa-file-code"></i> 页眉报告 JSON</a>` : ''}
                ${reports.footer_json ? `<a href="/${reports.footer_json}" download class="download-btn"><i class="fas fa-file-code"></i> 页脚报告 JSON</a>` : ''}
            </div>
        </div>
    `;
}

function resetPage() {
    originalFileInput.value = '';
    translatedFileInput.value = '';
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    clearLog();
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
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
    return MODEL_DISPLAY_NAMES[name] || name;
}
