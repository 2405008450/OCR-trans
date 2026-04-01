let originalFileInput, translatedFileInput, ruleFileInput, useAiRuleCheckbox, geminiRouteSelect, btnRun, btnReset;
let uploadSection, processingSection, resultSection, resultStats, resultGrid;
let progressBar, progressPercent, progressDetails, processingTitle, processingText;
let ruleEditorModal, ruleContentArea, streamLogWrap, streamLogEl;
const POLL_INTERVAL = 1500;
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
let pollingTimer = null;
let currentTaskId = null;
let sessionRuleContent = null;
let retryBtn = null;
let routeConfig = {};
let defaultRoute = 'openrouter';

function initElements() {
    originalFileInput = document.getElementById('originalFile');
    translatedFileInput = document.getElementById('translatedFile');
    ruleFileInput = document.getElementById('ruleFile');
    useAiRuleCheckbox = document.getElementById('useAiRule');
    btnRun = document.getElementById('btnRun');
    btnReset = document.getElementById('btnReset');
    uploadSection = document.getElementById('uploadSection');
    processingSection = document.getElementById('processingSection');
    resultSection = document.getElementById('resultSection');
    resultStats = document.getElementById('resultStats');
    resultGrid = document.getElementById('resultGrid');
    progressBar = document.getElementById('progressBar');
    progressPercent = document.getElementById('progressPercent');
    progressDetails = document.getElementById('progressDetails');
    processingTitle = document.getElementById('processingTitle');
    processingText = document.getElementById('processingText');
    ruleEditorModal = document.getElementById('ruleEditorModal');
    ruleContentArea = document.getElementById('ruleContent');

    ensureRouteControl();
    ensureLogPanel();

    btnRun?.addEventListener('click', runZhongfanyi);
    btnReset?.addEventListener('click', resetPage);
    document.getElementById('btnEditRule')?.addEventListener('click', openRuleEditor);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', async () => { initElements(); await loadRouteConfig(); });
} else {
    initElements();
    loadRouteConfig();
}

function ensureRouteControl() {
    geminiRouteSelect = null;
}

async function loadRouteConfig() {
    try {
        const resp = await fetch('/task/zhongfanyi/config');
        if (!resp.ok) throw new Error(`\u914d\u7f6e\u52a0\u8f7d\u5931\u8d25: ${resp.status}`);
        const data = await resp.json();
        routeConfig = data.routes || {};
        defaultRoute = data.default_route || defaultRoute;
    } catch (error) {
        console.error(error);
        routeConfig = {
            google: { label: '\u7ebf\u8def1' },
            openrouter: { label: '\u7ebf\u8def2' },
        };
    }
    renderRouteOptions();
}

function renderRouteOptions() {
    if (!geminiRouteSelect) return;
    geminiRouteSelect.innerHTML = '';
    Object.entries(routeConfig).forEach(([value, info]) => {
        geminiRouteSelect.add(new Option(info.label || value, value));
    });
    geminiRouteSelect.value = routeConfig[defaultRoute] ? defaultRoute : Object.keys(routeConfig)[0];
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

async function runZhongfanyi() {
    const originalFile = originalFileInput.files[0];
    const translatedFile = translatedFileInput.files[0];

    if (!originalFile || !translatedFile) {
        alert('请同时选择原文和译文文件，支持 .docx / .doc / .pdf。');
        return;
    }

    const useAiRule = useAiRuleCheckbox.checked;
    const ruleFile = ruleFileInput.files[0];
    if (useAiRule && !ruleFile && !(sessionRuleContent && sessionRuleContent.trim())) {
        alert('已勾选使用 AI 规则，请上传规则文件或在规则编辑器中填写规则内容。');
        return;
    }

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    removeRetryButton();
    updateProgressUI(0, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('original_file', originalFile);
        formData.append('translated_file', translatedFile);
        if (ruleFile) formData.append('rule_file', ruleFile);
        if (sessionRuleContent && sessionRuleContent.trim()) {
            formData.append('session_rule_content', sessionRuleContent.trim());
        }

        const url = `/task/zhongfanyi?use_ai_rule=${useAiRule ? 'true' : 'false'}&gemini_route=${encodeURIComponent(geminiRouteSelect?.value || defaultRoute)}`;
        const resp = await fetch(url, { method: 'POST', body: formData });

        if (!resp.ok) {
            let detailMsg = '';
            try {
                const errJson = await resp.json();
                const detail = errJson?.detail;
                if (typeof detail === 'string') detailMsg = detail;
                else if (detail && typeof detail === 'object') {
                    detailMsg = detail.error || (Array.isArray(detail.traceback) ? detail.traceback.join('\n') : '');
                }
            } catch (_) {}
            throw new Error(detailMsg || `请求失败: ${resp.status}`);
        }

        const submitResp = await resp.json();
        if (submitResp.status === 'ACCEPTED' && submitResp.task_id) {
            updateProgressUI(5, '任务已提交，正在后台处理...');
            startPolling(submitResp.task_id);
        } else if (submitResp.result) {
            showResult(submitResp.result);
        } else {
            resetPage();
        }
    } catch (err) {
        showFailure(`中翻译专检失败: ${err.message}`);
    }
}

function updateProgressUI(progress, message, details) {
    progressBar.style.setProperty('--progress', `${progress}%`);
    progressPercent.textContent = `${progress}%`;
    processingTitle.textContent = message || '中翻译专检处理中...';
    processingText.textContent = message || '正在处理...';
    if (details && details.length) {
        progressDetails.innerHTML = details.map((d) => `<div class="detail-item">${escapeHtml(d)}</div>`).join('');
    } else {
        progressDetails.innerHTML = `<div class="detail-item">${escapeHtml(message || '')}</div>`;
    }
}

async function pollTaskStatus(taskId) {
    try {
        const resp = await fetch(`/task/zhongfanyi/status/${taskId}`);
        if (!resp.ok) return null;
        const status = await resp.json();

        updateProgressUI(status.progress || 0, status.message || '正在处理...', status.details || []);
        syncLog(status.stream_log || status.result?.stream_log || '');

        if (status.status === 'done') {
            stopPolling();
            if (status.result) showResult(status.result);
        } else if (status.status === 'failed') {
            stopPolling();
            showFailure(`中翻译专检失败: ${status.error || '未知错误'}`, status.stream_log || '');
        }
        return status;
    } catch (err) {
        console.error('轮询状态出错:', err);
        return null;
    }
}

function startPolling(taskId) {
    currentTaskId = taskId;
    pollTaskStatus(taskId);
    pollingTimer = setInterval(() => pollTaskStatus(taskId), POLL_INTERVAL);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
    currentTaskId = null;
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
    updateProgressUI(100, message);
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

function showResult(data) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';

    const stats = data.stats || {};
    resultStats.innerHTML =
        `<div class="stat-card"><i class="fas fa-check"></i><h3>${stats.success ?? 0}</h3><p>成功</p></div>` +
        `<div class="stat-card"><i class="fas fa-times"></i><h3>${stats.failed ?? 0}</h3><p>失败</p></div>` +
        `<div class="stat-card"><i class="fas fa-forward"></i><h3>${stats.skipped ?? 0}</h3><p>跳过</p></div>`;

    const reports = data.reports || {};
    const correctedPath = data.corrected_docx ? `/${data.corrected_docx.replace(/^\/+/, '')}` : '';
    let links = `<a href="${correctedPath}" download class="download-btn"><i class="fas fa-file-word"></i> 下载修复后译文</a>`;
    if (reports['正文_json']) links += `<a href="/${reports['正文_json']}" download class="download-btn"><i class="fas fa-file-code"></i> 正文报告 JSON</a>`;
    if (reports['页眉_json']) links += `<a href="/${reports['页眉_json']}" download class="download-btn"><i class="fas fa-file-code"></i> 页眉报告 JSON</a>`;
    if (reports['页脚_json']) links += `<a href="/${reports['页脚_json']}" download class="download-btn"><i class="fas fa-file-code"></i> 页脚报告 JSON</a>`;

    resultGrid.innerHTML = `<div class="result-item"><h3>输出文件</h3><div class="download-links">${links}</div></div>`;
}

function resetPage() {
    originalFileInput.value = '';
    translatedFileInput.value = '';
    ruleFileInput.value = '';
    useAiRuleCheckbox.checked = false;
    sessionRuleContent = null;
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    clearLog();
    removeRetryButton();
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
}

function openRuleEditor() {
    if (!ruleEditorModal || !ruleContentArea) {
        alert('页面元素未就绪，请刷新后重试。');
        return;
    }
    ruleEditorModal.style.display = 'flex';
    loadRuleContent();
}

function closeRuleEditor() {
    if (ruleEditorModal) ruleEditorModal.style.display = 'none';
}

async function loadRuleContent() {
    if (!ruleContentArea) return;
    const radio = document.querySelector('input[name="ruleType"]:checked');
    const ruleType = radio ? radio.value : 'custom';
    ruleContentArea.value = '加载中...';
    ruleContentArea.disabled = true;
    try {
        const resp = await fetch(`/task/zhongfanyi/rule?rule_type=${ruleType}`);
        if (!resp.ok) throw new Error('加载失败');
        const data = await resp.json();
        ruleContentArea.value = data.content;
    } catch (err) {
        ruleContentArea.value = `加载规则文件失败: ${err.message}`;
    } finally {
        ruleContentArea.disabled = false;
    }
}

function saveRuleContent() {
    if (!ruleContentArea) return;
    sessionRuleContent = ruleContentArea.value;
    alert('已保存为本次任务使用，不会修改磁盘上的规则文件。');
    closeRuleEditor();
}

document.addEventListener('click', function(event) {
    if (ruleEditorModal && event.target === ruleEditorModal) {
        closeRuleEditor();
    }
});

function escapeHtml(value) {
    return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}


