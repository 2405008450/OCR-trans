'use strict';

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
const sourceLangSelect = document.getElementById('sourceLangSelect');
const targetLangGroup = document.getElementById('targetLangGroup');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const translationResults = document.getElementById('translationResults');

let selectedFile = null;
let pollingTimer = null;
let modelConfig = {};
let languageConfig = {};
let routeConfig = {};
let defaultModel = 'google/gemini-3-flash-preview';
let defaultRoute = 'google';
let streamLogWrap = null;
let streamLogEl = null;
let retryBtn = null;

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
    } catch (error) {
        console.error(error);
        modelConfig = {
            'google/gemini-3-flash-preview': { label: '快速版V2', description: '速度更快，适合常规 OCR。' },
            'google/gemini-3.1-pro-preview': { label: '增强版V2', description: '复杂版面表现更稳。' },
        };
        routeConfig = {
            google: { label: '线路1' },
            openrouter: { label: '线路2' },
        };
        languageConfig = {
            zh: { name: '中文' },
            en: { name: '英文' },
            ja: { name: '日文' },
            ko: { name: '韩文' },
            es: { name: '西班牙文' },
            fr: { name: '法文' },
            de: { name: '德文' },
            ru: { name: '俄文' },
        };
    }

    renderModels();
    renderRoutes();
    renderSourceLanguages();
    renderTargetLanguages();
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

function renderSourceLanguages() {
    sourceLangSelect.innerHTML = '';
    Object.entries(languageConfig).forEach(([code, info]) => {
        sourceLangSelect.add(new Option(info.name, code));
    });
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
        checkbox.type = 'checkbox';
        checkbox.value = code;
        if (code === 'en') {
            checkbox.checked = true;
            chip.classList.add('active');
        }

        const checkIcon = document.createElement('span');
        checkIcon.className = 'chip-check';
        checkIcon.innerHTML = '<i class="fas fa-check"></i>';

        const nameSpan = document.createElement('span');
        nameSpan.textContent = info.name;

        chip.appendChild(checkbox);
        chip.appendChild(checkIcon);
        chip.appendChild(nameSpan);
        chip.addEventListener('click', (event) => {
            event.preventDefault();
            checkbox.checked = !checkbox.checked;
            chip.classList.toggle('active', checkbox.checked);
            updateProcessButton();
        });

        targetLangGroup.appendChild(chip);
    });

    updateProcessButton();
}

function getSelectedTargetLangs() {
    return Array.from(targetLangGroup.querySelectorAll('input[type="checkbox"]:checked')).map((el) => el.value);
}

function updateProcessButton() {
    btnProcess.disabled = !(selectedFile && getSelectedTargetLangs().length > 0);
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
    if (files.length > 0) handleFile(files[0]);
}

function handleFileSelect(event) {
    const files = event.target.files;
    if (files.length > 0) handleFile(files[0]);
}

function handleFile(file) {
    const allowedExtensions = ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tif', '.tiff'];
    const lowerName = file.name.toLowerCase();
    const matched = allowedExtensions.some((ext) => lowerName.endsWith(ext));
    if (!matched) {
        alert('仅支持 PDF、PNG、JPG、JPEG、BMP、GIF、WEBP、TIF 文件。');
        return;
    }

    selectedFile = file;
    fileName.textContent = `${file.name} (${formatFileSize(file.size)})`;
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

async function processFile() {
    if (!selectedFile) return;

    const targetLangs = getSelectedTargetLangs();
    if (targetLangs.length === 0) {
        alert('请至少选择一种目标翻译语言。');
        return;
    }

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
            source_lang: sourceLangSelect.value,
            target_langs: targetLangs.join(','),
            ocr_model: modelSelect.value,
            gemini_route: geminiRouteSelect.value,
        });

        const response = await fetch(`/task/doc-translate?${params.toString()}`, {
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
        const response = await fetch(`/task/doc-translate/status/${taskId}`);
        if (!response.ok) throw new Error(`状态查询失败: ${response.status}`);

        const data = await response.json();
        updateProgress(data.progress || 0, data.message || '正在处理中...');
        syncLog(data.stream_log || data.result?.stream_log || '');

        if (data.status === 'done') {
            stopPolling();
            showResult(data.result || {});
        } else if (data.status === 'failed') {
            stopPolling();
            showFailure(data.error || '翻译失败', data.stream_log || '');
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

    const translations = result.translations || {};
    const langCount = Object.keys(translations).length;

    resultStats.innerHTML = `
        <div class="stat-card">
            <i class="fas fa-file-alt"></i>
            <h3>${escapeHtml(result.filename || '-')}</h3>
            <p>源文件</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml(getModelDisplayName(result.ocr_model || '-'))}</h3>
            <p>OCR 模型</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-globe"></i>
            <h3>${langCount} 种语言</h3>
            <p>翻译输出</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-check-circle"></i>
            <h3>成功</h3>
            <p>任务状态</p>
        </div>
    `;

    let html = '';
    if (result.raw_output_txt) {
        html += `
            <div class="translation-result-item">
                <div style="display:flex;align-items:center;gap:12px;">
                    <span class="lang-badge"><i class="fas fa-file-lines"></i> 原始文本</span>
                    <span style="color:var(--text-secondary);font-size:13px;">OCR 识别结果</span>
                </div>
                <div class="download-actions">
                    <a href="/${result.raw_output_txt}" download class="download-btn">
                        <i class="fas fa-download"></i> 下载 TXT
                    </a>
                </div>
            </div>
        `;
    }

    Object.entries(translations).forEach(([langCode, langResult]) => {
        html += `
            <div class="translation-result-item">
                <div style="display:flex;align-items:center;gap:12px;">
                    <span class="lang-badge"><i class="fas fa-language"></i> ${escapeHtml(langResult.lang_name || langCode)}</span>
                </div>
                <div class="download-actions">
                    <a href="/${langResult.output_docx}" download class="download-btn">
                        <i class="fas fa-file-word"></i> 下载 Word
                    </a>
                    <a href="/${langResult.translated_txt}" download class="download-btn">
                        <i class="fas fa-file-lines"></i> 下载译文文本
                    </a>
                </div>
            </div>
        `;
    });

    translationResults.innerHTML = html;
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
