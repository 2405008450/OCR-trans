const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const singleFileInput = document.getElementById('singleFile');
const modeDoubleRadio = document.getElementById('modeDouble');
const modeSingleRadio = document.getElementById('modeSingle');
const modeHint = document.getElementById('modeHint');
const uploadDesc = document.getElementById('uploadDesc');
const pageSubtitle = document.getElementById('pageSubtitle');
const useAiRuleCheckbox = document.getElementById('useAiRule');
const ruleFileInput = document.getElementById('ruleFile');
const btnRun = document.getElementById('btnRun');
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

const ruleEditorModal = document.getElementById('ruleEditorModal');
const ruleContentArea = document.getElementById('ruleContent');

const POLL_INTERVAL = 1500;
const ETA_TIME_ZONE = 'Asia/Shanghai';

let pollingTimer = null;
let defaultRoute = 'openrouter';
let defaultMode = 'double';
let currentMode = 'double';
let modeConfig = {};
let singleFileExtensions = ['.docx', '.doc', '.pdf', '.xlsx', '.xls', '.pptx'];
let doubleFileExtensions = ['.docx', '.doc', '.pdf', '.xlsx', '.xls', '.pptx'];
let streamLogWrap = null;
let streamLogEl = null;
let etaHint = null;
let sessionRuleContent = null;
let retryBtn = null;

init();

async function init() {
    ensureLogPanel();
    ensureEtaHint();
    bindEvents();
    await loadConfig();
    applyMode(defaultMode);
}

function bindEvents() {
    btnRun?.addEventListener('click', runZhongfanyi);
    btnReset?.addEventListener('click', resetPage);
    modeDoubleRadio?.addEventListener('change', () => applyMode('double'));
    modeSingleRadio?.addEventListener('change', () => applyMode('single'));
    document.getElementById('btnEditRule')?.addEventListener('click', openRuleEditor);
    document.addEventListener('click', (event) => {
        if (ruleEditorModal && event.target === ruleEditorModal) {
            closeRuleEditor();
        }
    });
}

async function loadConfig() {
    try {
        const resp = await fetch('/task/zhongfanyi/config');
        if (!resp.ok) throw new Error(`配置加载失败: ${resp.status}`);
        const data = await resp.json();
        defaultRoute = data.default_route || defaultRoute;
        defaultMode = data.default_mode || defaultMode;
        modeConfig = data.modes || {};
        singleFileExtensions = data.single_file_extensions || singleFileExtensions;
        doubleFileExtensions = data.double_file_extensions || doubleFileExtensions;
    } catch (error) {
        console.error(error);
        modeConfig = {
            double: {
                label: '双文件模式',
                description: '上传原文和译文两个文件，自动对比并输出修复结果。',
            },
            single: {
                label: '单文件模式',
                description: '上传一个双语对照文件，自动输出 JSON / Excel 报告。',
            },
        };
    }

    singleFileInput.accept = singleFileExtensions.join(',');
    originalFileInput.accept = doubleFileExtensions.join(',');
    translatedFileInput.accept = doubleFileExtensions.join(',');
}

function applyMode(mode) {
    currentMode = mode === 'single' ? 'single' : 'double';
    if (modeDoubleRadio) modeDoubleRadio.checked = currentMode === 'double';
    if (modeSingleRadio) modeSingleRadio.checked = currentMode === 'single';

    const singleMode = currentMode === 'single';
    doubleFileFields.style.display = singleMode ? 'none' : 'block';
    singleFileFields.style.display = singleMode ? 'block' : 'none';

    const currentModeConfig = modeConfig[currentMode] || {};
    modeHint.textContent = currentModeConfig.description || '';
    pageSubtitle.textContent = singleMode
        ? '上传一个双语对照文件，自动输出 JSON / Excel 报告，并按文件类型生成修订版或批注版'
        : '上传原文与译文两个文件，自动对比并输出修订版或批注版结果';
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

async function runZhongfanyi() {
    const mode = currentMode;
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
            alert('请同时选择原文和译文文件');
            return;
        }
        formData.append('original_file', originalFile);
        formData.append('translated_file', translatedFile);
    }

    const useAiRule = useAiRuleCheckbox.checked;
    const ruleFile = ruleFileInput.files[0];
    if (useAiRule && !ruleFile && !(sessionRuleContent && sessionRuleContent.trim())) {
        alert('已勾选使用 AI 规则，请上传规则文件或在规则编辑器中填写规则内容。');
        return;
    }
    if (ruleFile) formData.append('rule_file', ruleFile);
    if (sessionRuleContent && sessionRuleContent.trim()) {
        formData.append('session_rule_content', sessionRuleContent.trim());
    }

    uploadSection.style.display = 'none';
    resultSection.style.display = 'none';
    processingSection.style.display = 'block';
    clearLog();
    removeRetryButton();
    updateProgressUI(0, '正在提交任务...');

    try {
        const params = new URLSearchParams({
            mode,
            use_ai_rule: useAiRule ? 'true' : 'false',
            gemini_route: defaultRoute,
        });

        const resp = await fetch(`/task/zhongfanyi?${params.toString()}`, {
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

        showResult(submitResp.result || submitResp);
    } catch (error) {
        showFailure(`中翻译专检失败: ${error.message}`);
    }
}

function updateProgressUI(progress, message, details = [], task = null) {
    progressBar.style.setProperty('--progress', `${progress}%`);
    progressPercent.textContent = `${progress}%`;
    processingTitle.textContent = message || '中翻译专检处理中...';
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
        const resp = await fetch(`/task/zhongfanyi/status/${taskId}`);
        if (!resp.ok) throw new Error(`获取任务状态失败: ${resp.status}`);

        const status = await resp.json();
        updateProgressUI(status.progress || 0, status.message || '正在处理...', status.details || [], status);
        syncLog(status.stream_log || status.result?.stream_log || '');

        if (status.status === 'done') {
            stopPolling();
            showResult(status.result || status);
        } else if (status.status === 'failed') {
            stopPolling();
            showFailure(`中翻译专检失败: ${status.error || '未知错误'}`, status.stream_log || '');
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
    uploadSection.style.display = 'none';
    resultSection.style.display = 'block';

    resultSummary.textContent = data.summary || '';

    const stats = data.stats || {};
    const reportCounts = data.report_counts || {};
    resultStats.innerHTML =
        `<div class="stat-card"><i class="fas fa-check"></i><h3>${stats.success ?? 0}</h3><p>成功修复</p></div>` +
        `<div class="stat-card"><i class="fas fa-times"></i><h3>${stats.failed ?? 0}</h3><p>替换失败</p></div>` +
        `<div class="stat-card"><i class="fas fa-forward"></i><h3>${stats.skipped ?? 0}</h3><p>跳过</p></div>` +
        `<div class="stat-card"><i class="fas fa-list-check"></i><h3>${data.total_issues ?? (reportCounts.body_issues ?? 0) + (reportCounts.header_issues ?? 0) + (reportCounts.footer_issues ?? 0)}</h3><p>报告问题数</p></div>`;

    const links = [];
    appendOutputLink(links, data.corrected_docx, 'fa-file-word', '下载修复后文档');
    appendOutputLink(links, data.annotated_pdf, 'fa-file-pdf', '下载批注版 PDF');
    appendOutputLink(links, data.annotated_excel, 'fa-file-excel', '下载批注版 Excel');
    appendOutputLink(links, data.annotated_pptx, 'fa-file-powerpoint', '下载批注版 PPTX');
    if (!links.length) {
        appendOutputLink(links, data.output_file, 'fa-file-arrow-down', '下载输出文件');
    }

    const reports = data.reports || {};
    appendReportLink(links, reports['正文_json'], 'fa-file-code', `正文报告 JSON${buildCountSuffix(reportCounts.body_issues)}`);
    appendReportLink(links, reports['页眉_json'], 'fa-file-code', `页眉报告 JSON${buildCountSuffix(reportCounts.header_issues)}`);
    appendReportLink(links, reports['页脚_json'], 'fa-file-code', `页脚报告 JSON${buildCountSuffix(reportCounts.footer_issues)}`);
    appendReportLink(links, reports['正文_excel'], 'fa-file-excel', '正文报告 Excel');
    appendReportLink(links, reports['页眉_excel'], 'fa-file-excel', '页眉报告 Excel');
    appendReportLink(links, reports['页脚_excel'], 'fa-file-excel', '页脚报告 Excel');

    resultGrid.innerHTML = `<div class="result-item"><h3>输出文件</h3><div class="download-links">${links.join('')}</div></div>`;
}

function appendOutputLink(list, rawPath, icon, label) {
    if (!rawPath) return;
    const path = `/${String(rawPath).replace(/^\/+/, '')}`;
    list.push(`<a href="${path}" download class="download-btn"><i class="fas ${icon}"></i> ${escapeHtml(label)}</a>`);
}

function appendReportLink(list, rawPath, icon, label) {
    if (!rawPath) return;
    appendOutputLink(list, rawPath, icon, label);
}

function buildCountSuffix(value) {
    const count = Number(value ?? 0);
    if (!Number.isFinite(count) || count <= 0) return '';
    return `（${count}条）`;
}

function resetPage() {
    originalFileInput.value = '';
    translatedFileInput.value = '';
    singleFileInput.value = '';
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
    resultSummary.textContent = '';
    resultStats.innerHTML = '';
    resultGrid.innerHTML = '';
    applyMode(defaultMode);
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

async function safeReadError(resp) {
    try {
        const errJson = await resp.json();
        const detail = errJson?.detail;
        if (typeof detail === 'string') return detail;
        if (detail && typeof detail === 'object') {
            return detail.error || (Array.isArray(detail.traceback) ? detail.traceback.join('\n') : '');
        }
        return '';
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

window.closeRuleEditor = closeRuleEditor;
window.loadRuleContent = loadRuleContent;
window.saveRuleContent = saveRuleContent;
