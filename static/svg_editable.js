const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const fileName = document.getElementById('fileName');
const processBtn = document.getElementById('processBtn');
const modelSelect = document.getElementById('modelSelect');
const modelHint = document.getElementById('modelHint');
const threshold = document.getElementById('threshold');
const thresholdValue = document.getElementById('thresholdValue');
const statusPanel = document.getElementById('status');
const statusText = document.getElementById('statusText');
const progressBar = document.getElementById('progressBar');
const progressValue = document.getElementById('progressValue');
const errorText = document.getElementById('errorText');
const resultPanel = document.getElementById('result');
const downloads = document.getElementById('downloads');
let selectedFile = null;
let config = {};
let pollingTimer = null;

init();

async function init() {
    bindEvents();
    try {
        const response = await fetch('/task/svg-editable/config');
        if (!response.ok) throw new Error('配置加载失败');
        config = await response.json();
        renderModels();
        threshold.value = Math.round((config.confidence_threshold || 0.82) * 100);
        updateThreshold();
    } catch (error) { showError(error.message); }
}

function bindEvents() {
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', event => { event.preventDefault(); uploadArea.classList.add('drag'); });
    uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('drag'));
    uploadArea.addEventListener('drop', event => { event.preventDefault(); uploadArea.classList.remove('drag'); setFile(event.dataTransfer.files[0]); });
    fileInput.addEventListener('change', () => setFile(fileInput.files[0]));
    threshold.addEventListener('input', updateThreshold);
    modelSelect.addEventListener('change', updateModelHint);
    processBtn.addEventListener('click', submitTask);
}

function renderModels() {
    modelSelect.innerHTML = '';
    Object.entries(config.models || {}).filter(([value]) => value !== 'google/gemini-3-flash-preview').forEach(([value, info]) => {
        const option = new Option(info.label || value, value);
        option.title = value;
        modelSelect.add(option);
    });
    modelSelect.value = config.default_model || modelSelect.options[0]?.value || '';
    updateModelHint();
}

function updateModelHint() { modelHint.textContent = (config.models || {})[modelSelect.value]?.description || ''; modelSelect.title = modelSelect.value; }
function updateThreshold() { thresholdValue.textContent = `${threshold.value}%`; }

function setFile(file) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.svg')) return showError('请选择 .svg 文件');
    if (file.size > 10 * 1024 * 1024) return showError('SVG 文件不能超过 10MB');
    selectedFile = file;
    fileName.textContent = `${file.name} · ${formatBytes(file.size)}`;
    processBtn.disabled = false;
    errorText.textContent = '';
    resultPanel.style.display = 'none';
}

async function submitTask() {
    if (!selectedFile || processBtn.disabled) return;
    processBtn.disabled = true;
    resultPanel.style.display = 'none';
    statusPanel.style.display = 'block';
    errorText.textContent = '';
    updateProgress(2, '正在提交任务…');
    const form = new FormData();
    form.append('file', selectedFile);
    const params = new URLSearchParams({ model:modelSelect.value, gemini_route:config.default_route || 'openrouter', confidence_threshold:(Number(threshold.value) / 100).toFixed(2) });
    try {
        const response = await fetch(`/task/svg-editable?${params}`, { method:'POST', body:form });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || '任务提交失败');
        poll(payload.task_id);
    } catch (error) { showError(error.message); processBtn.disabled = false; }
}

function poll(taskId) {
    clearInterval(pollingTimer);
    const check = async () => {
        try {
            const response = await fetch(`/task/svg-editable/status/${taskId}`);
            if (!response.ok) throw new Error('无法读取任务状态');
            const task = await response.json();
            updateProgress(Number(task.progress || 0), task.message || task.status || '处理中…');
            if (task.status === 'done') {
                clearInterval(pollingTimer); renderResult(taskId, task); processBtn.disabled = false;
            } else if (['failed','cancelled'].includes(task.status)) {
                clearInterval(pollingTimer); showError(task.error || task.message || '处理失败'); processBtn.disabled = false;
            }
        } catch (error) { clearInterval(pollingTimer); showError(error.message); processBtn.disabled = false; }
    };
    check();
    pollingTimer = setInterval(check, 1500);
}

function renderResult(taskId, task) {
    updateProgress(100, '处理完成');
    const result = task.result || {};
    document.getElementById('convertedCount').textContent = result.converted_line_count || 0;
    document.getElementById('hiddenCount').textContent = result.hidden_path_count || 0;
    document.getElementById('existingCount').textContent = result.existing_text_count || 0;
    downloads.innerHTML = (task.output_files || []).map(item => {
        const name = item.name || item.path.split(/[\\/]/).pop();
        const url = `/task/${encodeURIComponent(taskId)}/download?file_path=${encodeURIComponent(item.path)}&download_name=${encodeURIComponent(name)}`;
        const icon = name.endsWith('.svg') ? 'fa-pen-ruler' : name.endsWith('.json') ? 'fa-list-check' : 'fa-image';
        const label = name.endsWith('.svg') ? '下载可编辑 SVG' : name.endsWith('.json') ? '下载复核 JSON' : '下载预览图';
        return `<a class="download" href="${url}"><i class="fas ${icon}"></i>${label}</a>`;
    }).join('');
    resultPanel.style.display = 'block';
}

function updateProgress(value, message) {
    const normalized = Math.max(0, Math.min(100, value));
    progressBar.style.width = `${normalized}%`;
    progressValue.textContent = `${normalized}%`;
    statusText.textContent = message;
}

function showError(message) {
    statusPanel.style.display = 'block';
    errorText.textContent = message || '发生未知错误';
    statusText.textContent = '处理未完成';
}

function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
