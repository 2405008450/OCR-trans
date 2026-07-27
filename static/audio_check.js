const PROMPT_STORAGE_KEY = 'audio-check-prompts-v2';
const POLL_INTERVAL_MS = 1500;

const elements = {
    file: document.getElementById('audioFile'),
    dropZone: document.getElementById('dropZone'),
    fileSummary: document.getElementById('fileSummary'),
    fileName: document.getElementById('fileName'),
    fileMeta: document.getElementById('fileMeta'),
    preview: document.getElementById('audioPreview'),
    model: document.getElementById('modelSelect'),
    route: document.getElementById('routeSelect'),
    modelDescription: document.getElementById('modelDescription'),
    routeDescription: document.getElementById('routeDescription'),
    temperature: document.getElementById('temperature'),
    temperatureValue: document.getElementById('temperatureValue'),
    maxTokens: document.getElementById('maxTokens'),
    timeout: document.getElementById('timeoutSeconds'),
    systemPrompt: document.getElementById('systemPrompt'),
    userPrompt: document.getElementById('userPrompt'),
    resetPrompts: document.getElementById('resetPrompts'),
    run: document.getElementById('runCheck'),
    cancel: document.getElementById('cancelTask'),
    error: document.getElementById('errorBox'),
    progressPanel: document.getElementById('progressPanel'),
    progressTitle: document.getElementById('progressTitle'),
    progressFill: document.getElementById('progressFill'),
    progressMessage: document.getElementById('progressMessage'),
    progressPercent: document.getElementById('progressPercent'),
    resultPanel: document.getElementById('resultPanel'),
    markdown: document.getElementById('markdownResult'),
    download: document.getElementById('downloadReport'),
    formatHint: document.getElementById('formatHint'),
};

let config = null;
let selectedFile = null;
let currentTaskId = null;
let pollTimer = null;
let previewUrl = null;
let taskActive = false;

function showError(message) {
    elements.error.textContent = message || '发生未知错误';
    elements.error.style.display = 'block';
}

function clearError() {
    elements.error.textContent = '';
    elements.error.style.display = 'none';
}

function formatBytes(bytes) {
    if (!Number.isFinite(bytes)) return '-';
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

function setSelectedFile(file) {
    clearError();
    if (!file) return;
    const extension = `.${String(file.name).split('.').pop().toLowerCase()}`;
    const allowed = config?.allowed_extensions || ['.wav', '.mp3', '.aiff', '.aac', '.ogg', '.flac'];
    const maxBytes = (config?.max_file_mb || 20) * 1024 * 1024;
    if (!allowed.includes(extension)) {
        showError(`不支持 ${extension} 格式，仅支持 ${allowed.join('、')}`);
        return;
    }
    if (!file.size) {
        showError('音频文件不能为空');
        return;
    }
    if (file.size > maxBytes) {
        showError(`文件超过 ${config?.max_file_mb || 20}MB 上传限制`);
        return;
    }
    selectedFile = file;
    elements.fileName.textContent = file.name;
    elements.fileMeta.textContent = `${formatBytes(file.size)} · ${file.type || extension}`;
    elements.fileSummary.style.display = 'block';
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = URL.createObjectURL(file);
    elements.preview.src = previewUrl;
    elements.preview.style.display = 'block';
}

function populateSelect(select, values, defaultValue) {
    select.innerHTML = '';
    Object.entries(values || {}).forEach(([value, info]) => {
        select.add(new Option(info.label || value, value));
    });
    const fallback = Object.keys(values || {})[0] || '';
    select.value = values?.[defaultValue] ? defaultValue : fallback;
}

function updateDescriptions() {
    elements.modelDescription.textContent = config?.models?.[elements.model.value]?.description || '';
    elements.routeDescription.textContent = config?.routes?.[elements.route.value]?.description || '';
}

function loadStoredPrompts() {
    try {
        const stored = JSON.parse(localStorage.getItem(PROMPT_STORAGE_KEY) || 'null');
        elements.systemPrompt.value = stored?.system_prompt || config.default_system_prompt || '';
        elements.userPrompt.value = stored?.user_prompt || config.default_user_prompt || '';
    } catch (_) {
        elements.systemPrompt.value = config.default_system_prompt || '';
        elements.userPrompt.value = config.default_user_prompt || '';
    }
}

function storePrompts() {
    localStorage.setItem(PROMPT_STORAGE_KEY, JSON.stringify({
        system_prompt: elements.systemPrompt.value,
        user_prompt: elements.userPrompt.value,
    }));
}

async function loadConfig() {
    const response = await fetch('/task/audio-check/config');
    if (!response.ok) throw new Error(`配置加载失败：${response.status}`);
    config = await response.json();
    populateSelect(elements.model, config.models, config.default_model);
    populateSelect(elements.route, config.routes, config.default_route);
    const defaults = config.defaults || {};
    elements.temperature.value = defaults.temperature ?? 0.1;
    elements.temperatureValue.textContent = elements.temperature.value;
    elements.maxTokens.value = defaults.max_output_tokens ?? 32768;
    elements.timeout.value = defaults.timeout_seconds ?? 120;
    elements.file.accept = (config.allowed_extensions || []).join(',');
    elements.formatHint.textContent = `支持 ${(config.allowed_extensions || []).map(item => item.slice(1).toUpperCase()).join('、')}，单文件不超过 ${config.max_file_mb || 20}MB。`;
    loadStoredPrompts();
    updateDescriptions();
    if (!Object.keys(config.routes || {}).length) {
        showError('当前没有已配置的 Gemini 调用线路，请先配置 Vertex、AI Studio 或 OpenRouter。');
        elements.run.disabled = true;
    }
}

function updateProgress(task) {
    const progress = Math.max(0, Math.min(100, Number(task.progress || 0)));
    elements.progressPanel.style.display = 'block';
    elements.progressFill.style.width = `${progress}%`;
    elements.progressPercent.textContent = `${progress}%`;
    elements.progressMessage.textContent = task.message || (task.status === 'queued' ? '任务排队中' : '处理中');
    elements.progressTitle.textContent = task.status === 'queued'
        ? `任务排队中${task.queue_position ? `（第 ${task.queue_position} 位）` : ''}`
        : '正在检查音频';
}

function renderResult(result, taskId) {
    elements.markdown.innerHTML = result.analysis_html || '<p>模型未返回可显示结果。</p>';
    if (result.report_markdown) {
        const params = new URLSearchParams({
            file_path: result.report_markdown,
            download_name: result.report_markdown.split('/').pop() || '音频质量检查报告.md',
        });
        elements.download.href = `/task/${encodeURIComponent(taskId)}/download?${params}`;
        elements.download.style.display = 'inline-flex';
    } else {
        elements.download.style.display = 'none';
    }
    elements.resultPanel.style.display = 'block';
    elements.resultPanel.scrollIntoView({behavior: 'smooth', block: 'start'});
}

async function pollStatus() {
    if (!currentTaskId) return;
    try {
        const response = await fetch(`/task/audio-check/status/${encodeURIComponent(currentTaskId)}`);
        if (!response.ok) throw new Error(`状态读取失败：${response.status}`);
        const task = await response.json();
        updateProgress(task);
        if (task.status === 'done') {
            taskActive = false;
            stopPolling();
            elements.progressFill.style.width = '100%';
            elements.progressPercent.textContent = '100%';
            elements.progressTitle.textContent = '质检完成';
            elements.cancel.style.display = 'none';
            elements.run.disabled = false;
            renderResult(task.result || {}, currentTaskId);
        } else if (task.status === 'failed' || task.status === 'cancelled') {
            taskActive = false;
            stopPolling();
            showError(task.error || task.message || (task.status === 'cancelled' ? '任务已取消' : '任务处理失败'));
            elements.cancel.style.display = 'none';
            elements.run.disabled = false;
        }
    } catch (error) {
        stopPolling();
        showError(error.message);
        elements.run.disabled = false;
        elements.cancel.style.display = 'none';
    }
}

function stopPolling() {
    if (pollTimer) clearInterval(pollTimer);
    pollTimer = null;
}

async function submitCheck() {
    clearError();
    elements.resultPanel.style.display = 'none';
    if (!selectedFile) {
        showError('请先选择一个音频文件');
        return;
    }
    if (!elements.systemPrompt.value.trim() || !elements.userPrompt.value.trim()) {
        showError('系统提示词和本次检查要求不能为空');
        return;
    }
    storePrompts();
    const data = new FormData();
    data.append('file', selectedFile);
    data.append('system_prompt', elements.systemPrompt.value);
    data.append('user_prompt', elements.userPrompt.value);
    data.append('model_name', elements.model.value);
    data.append('gemini_route', elements.route.value);
    data.append('temperature', elements.temperature.value);
    data.append('max_output_tokens', elements.maxTokens.value);
    data.append('timeout_seconds', elements.timeout.value);

    elements.run.disabled = true;
    elements.progressPanel.style.display = 'block';
    updateProgress({status: 'queued', progress: 0, message: '正在上传音频'});
    try {
        const response = await fetch('/task/audio-check', {method: 'POST', body: data});
        const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload.detail || `提交失败：${response.status}`);
        currentTaskId = payload.task_id;
        taskActive = true;
        elements.cancel.style.display = 'inline-flex';
        stopPolling();
        await pollStatus();
        if (taskActive && !pollTimer) pollTimer = setInterval(pollStatus, POLL_INTERVAL_MS);
    } catch (error) {
        showError(error.message);
        elements.run.disabled = false;
        elements.cancel.style.display = 'none';
    }
}

async function cancelCurrentTask() {
    if (!currentTaskId) return;
    elements.cancel.disabled = true;
    try {
        const response = await fetch(`/task/${encodeURIComponent(currentTaskId)}/cancel`, {method: 'POST'});
        if (!response.ok) throw new Error('取消请求失败');
        elements.progressMessage.textContent = '已提交取消请求';
    } catch (error) {
        showError(error.message);
    } finally {
        elements.cancel.disabled = false;
    }
}

function bindEvents() {
    elements.dropZone.addEventListener('click', () => elements.file.click());
    elements.dropZone.addEventListener('keydown', event => {
        if (event.key === 'Enter' || event.key === ' ') elements.file.click();
    });
    elements.file.addEventListener('change', () => setSelectedFile(elements.file.files?.[0]));
    ['dragenter', 'dragover'].forEach(name => elements.dropZone.addEventListener(name, event => {
        event.preventDefault();
        elements.dropZone.classList.add('dragging');
    }));
    ['dragleave', 'drop'].forEach(name => elements.dropZone.addEventListener(name, event => {
        event.preventDefault();
        elements.dropZone.classList.remove('dragging');
    }));
    elements.dropZone.addEventListener('drop', event => setSelectedFile(event.dataTransfer?.files?.[0]));
    elements.model.addEventListener('change', updateDescriptions);
    elements.route.addEventListener('change', updateDescriptions);
    elements.temperature.addEventListener('input', () => {
        elements.temperatureValue.textContent = elements.temperature.value;
    });
    elements.systemPrompt.addEventListener('input', storePrompts);
    elements.userPrompt.addEventListener('input', storePrompts);
    elements.resetPrompts.addEventListener('click', () => {
        elements.systemPrompt.value = config.default_system_prompt || '';
        elements.userPrompt.value = config.default_user_prompt || '';
        storePrompts();
    });
    elements.run.addEventListener('click', submitCheck);
    elements.cancel.addEventListener('click', cancelCurrentTask);
    window.addEventListener('beforeunload', () => {
        stopPolling();
        if (previewUrl) URL.revokeObjectURL(previewUrl);
    });
}

async function init() {
    bindEvents();
    try {
        await loadConfig();
    } catch (error) {
        showError(error.message);
        elements.run.disabled = true;
    }
}

init();
