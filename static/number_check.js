const alignmentFileInput = document.getElementById('alignmentFile');
const targetFileInput = document.getElementById('targetFile');
const sourceHfFileInput = document.getElementById('sourceHfFile');
const sourceFileInput = document.getElementById('sourceFile');
const directTargetFileInput = document.getElementById('directTargetFile');
const modeAlignmentRadio = document.getElementById('modeAlignment');
const modeDirectRadio = document.getElementById('modeDirect');
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

const alignmentFields = document.getElementById('alignmentFields');
const directFields = document.getElementById('directFields');
const progressBar = document.getElementById('progressBar');
const progressPercent = document.getElementById('progressPercent');
const progressDetails = document.getElementById('progressDetails');
const processingTitle = document.getElementById('processingTitle');
const processingText = document.getElementById('processingText');

const POLL_INTERVAL = 1000;
const MODEL_DISPLAY_NAMES = {
    'gemini-3-flash-preview': '快速版V2',
    'google/gemini-3-flash-preview': '快速版V2',
    'gemini-3.5-flash': '新模型',
    'google/gemini-3.5-flash': '新模型',
    'gemini-3.1-pro-preview': '增强版V2',
    'google/gemini-3.1-pro-preview': '增强版V2',
};

let pollingTimer = null;
let modelConfig = {};
let modeConfig = {};
let defaultRoute = 'openrouter';
let defaultModel = 'gemini-3.1-pro-preview';
let defaultMode = 'alignment';
let currentMode = 'alignment';
let streamLogWrap = null;
let streamLogEl = null;

init();

async function init() {
    ensureLogPanel();
    bindEvents();
    await loadConfig();
    applyMode(defaultMode);
}

function bindEvents() {
    btnRunCheck?.addEventListener('click', runNumberCheck);
    btnReset?.addEventListener('click', resetPage);
    modeAlignmentRadio?.addEventListener('change', () => applyMode('alignment'));
    modeDirectRadio?.addEventListener('change', () => applyMode('direct'));
    modelSelect?.addEventListener('change', updateModelInfo);
}

async function loadConfig() {
    try {
        const resp = await fetch('/task/number-check/config');
        if (!resp.ok) throw new Error(`配置加载失败: ${resp.status}`);
        const data = await resp.json();
        modelConfig = data.models || {};
        modeConfig = data.modes || {};
        defaultRoute = data.default_route || defaultRoute;
        defaultModel = data.default_model || defaultModel;
        defaultMode = data.default_mode || defaultMode;

        alignmentFileInput.accept = (data.alignment_file_extensions || ['.xlsx']).join(',');
        targetFileInput.accept = (data.target_file_extensions || ['.docx', '.doc', '.xlsx', '.pptx', '.pdf']).join(',');
        sourceHfFileInput.accept = (data.source_hf_file_extensions || ['.docx', '.doc']).join(',');
        sourceFileInput.accept = (data.direct_file_extensions || ['.docx', '.doc', '.xlsx', '.pptx']).join(',');
        directTargetFileInput.accept = (data.direct_file_extensions || ['.docx', '.doc', '.xlsx', '.pptx']).join(',');
    } catch (error) {
        console.error(error);
        modelConfig = {
            'gemini-3-flash-preview': { label: '快速版V2', description: '速度更快，适合常规数字核对场景。' },
            'gemini-3.5-flash': { label: '新模型', description: 'OpenRouter 新模型，适合常规数字核对场景。' },
            'gemini-3.1-pro-preview': { label: '增强版V2', description: '推理更强，适合复杂编号和上下文判断场景。' },
        };
        modeConfig = {
            alignment: { description: '上传含“原文”“译文”两列的 Excel，可选上传译文文件生成修订版。' },
            direct: { description: '上传原文和译文文件，由新版数检程序直接抽取内容。' },
        };
    }

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
    const currentModel = modelSelect?.value || defaultModel;
    const info = modelConfig[currentModel] || {};
    modelDesc.textContent = info.description || '';
}

function getSelectedMode() {
    return modeDirectRadio?.checked ? 'direct' : 'alignment';
}

function applyMode(mode) {
    currentMode = mode === 'direct' ? 'direct' : 'alignment';
    modeAlignmentRadio.checked = currentMode === 'alignment';
    modeDirectRadio.checked = currentMode === 'direct';

    const directMode = currentMode === 'direct';
    alignmentFields.style.display = directMode ? 'none' : 'block';
    directFields.style.display = directMode ? 'block' : 'none';

    const description = modeConfig[currentMode]?.description || (
        directMode
            ? '上传原文和译文文件，由新版数检程序直接抽取内容。'
            : '上传含“原文”“译文”两列的 Excel，可选上传译文文件生成修订版。'
    );
    modeHint.textContent = description;
    uploadDesc.textContent = directMode
        ? '原文和译文结构一致时可使用该模式；PDF 建议先制作对照 Excel。'
        : '对照 Excel 是新版数检推荐输入；上传译文文件后会额外输出修订版。';
    pageSubtitle.textContent = directMode
        ? '直接从原文和译文文件抽取内容，生成数值检查报告和可用修订文件。'
        : '基于已对齐 Excel 执行规则检查与 AI 复核，适合排版复杂或 PDF 来源文件。';
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

async function runNumberCheck() {
    const mode = getSelectedMode();
    const formData = new FormData();

    if (mode === 'alignment') {
        const alignmentFile = alignmentFileInput.files[0];
        if (!alignmentFile) {
            alert('请选择对照 Excel 文件。');
            return;
        }
        formData.append('alignment_file', alignmentFile);
        if (targetFileInput.files[0]) formData.append('target_file', targetFileInput.files[0]);
        if (sourceHfFileInput.files[0]) formData.append('source_hf_file', sourceHfFileInput.files[0]);
    } else {
        const sourceFile = sourceFileInput.files[0];
        const targetFile = directTargetFileInput.files[0];
        if (!sourceFile || !targetFile) {
            alert('请选择原文文件和译文文件。');
            return;
        }
        formData.append('source_file', sourceFile);
        formData.append('target_file', targetFile);
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

async function pollTaskStatus(taskId) {
    try {
        const resp = await fetch(`/task/number-check/status/${taskId}`);
        if (!resp.ok) throw new Error(`获取任务状态失败: ${resp.status}`);

        const status = await resp.json();
        updateProgressUI(status.progress || 0, status.message || '正在处理...', status.details || []);
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
    pollingTimer = setInterval(() => pollTaskStatus(taskId), POLL_INTERVAL);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
}

function updateProgressUI(progress, message, details = []) {
    const safeProgress = Math.max(0, Math.min(100, Number(progress) || 0));
    progressBar.style.background = `linear-gradient(90deg, var(--primary-color) ${safeProgress}%, #e4edf5 ${safeProgress}%)`;
    progressPercent.textContent = `${safeProgress}%`;
    processingTitle.textContent = message || '数字专检处理中...';
    processingText.textContent = message || '正在处理...';
    progressDetails.innerHTML = details.length
        ? details.map((item) => `<div class="detail-item">${escapeHtml(item)}</div>`).join('')
        : `<div class="detail-item">${escapeHtml(message || '正在处理...')}</div>`;
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

    const reportCounts = data.report_counts || {};
    const stats = data.stats || {};
    const totalIssues = Number(stats.total_issues ?? reportCounts.total_issues ?? 0);

    resultSummary.textContent = data.summary || '新版数字专检已完成。';
    resultStats.innerHTML = buildStatsHtml(stats, reportCounts, totalIssues, data.model_name || defaultModel);
    resultGrid.innerHTML = buildResultHtml(data);
}

function buildStatsHtml(stats, reportCounts, totalIssues, modelName) {
    const bodyIssues = stats.body_issues ?? reportCounts.body_issues ?? 0;
    const headerIssues = stats.header_issues ?? reportCounts.header_issues ?? 0;
    const footerIssues = stats.footer_issues ?? reportCounts.footer_issues ?? 0;
    return `
        <div class="stat-card">
            <i class="fas fa-list-check"></i>
            <h3>${totalIssues}</h3>
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
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml(getModelDisplayName(modelName))}</h3>
            <p>模型</p>
        </div>
    `;
}

function buildResultHtml(data) {
    const links = [];
    const reports = data.reports || {};

    addDownload(links, data.revised_file || data.corrected_docx, 'fa-file-pen', '下载修订文件');
    addDownload(links, reports.report_excel, 'fa-file-excel', '下载综合 Excel 报告');
    addDownload(links, reports.alignment_excel, 'fa-table', '下载生成的对照 Excel');
    addDownload(links, reports.alignment_json, 'fa-file-code', '下载对照 JSON');
    addDownload(links, reports.body_json, 'fa-file-code', '正文检查 JSON');
    addDownload(links, reports.body_errors_json, 'fa-bug', '正文错误 JSON');
    addDownload(links, reports.body_flat_errors_json, 'fa-list', '正文错误明细 JSON');
    addDownload(links, reports.header_json, 'fa-heading', '页眉检查 JSON');
    addDownload(links, reports.footer_json, 'fa-shoe-prints', '页脚检查 JSON');

    if (!links.length) {
        links.push('<div class="detail-item">当前没有可下载的输出文件。</div>');
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

function addDownload(links, path, icon, label) {
    if (!path) return;
    links.push(`<a href="/${path}" download class="download-btn"><i class="fas ${icon}"></i> ${escapeHtml(label)}</a>`);
}

function resetPage() {
    alignmentFileInput.value = '';
    targetFileInput.value = '';
    sourceHfFileInput.value = '';
    sourceFileInput.value = '';
    directTargetFileInput.value = '';
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    clearLog();
    updateProgressUI(0, '');
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
