let selectedFile = null;
let pollingTimer = null;
let modelConfig = {};
let defaultModel = 'google/gemini-3-flash-preview';
const MODEL_DISPLAY_NAMES = {
    'google/gemini-2.5-flash': '快速版V1',
    'google/gemini-2.5-pro': '增强版V1',
    'google/gemini-3-flash-preview': '快速版V2',
    'google/gemini-3.1-pro-preview': '增强版V2',
    'Google Gemini 2.5 Flash': '快速版V1',
    'Google Gemini 2.5 Pro': '增强版V1',
    'Google Gemini 3 Flash Preview': '快速版V2',
    'Google Gemini 3.1 Pro Preview': '增强版V2',
};

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview = document.getElementById('filePreview');
const fileName = document.getElementById('fileName');
const btnRemove = document.getElementById('btnRemove');
const btnProcess = document.getElementById('btnProcess');
const btnNewTask = document.getElementById('btnNewTask');
const modelSelect = document.getElementById('modelSelect');
const modelLabel = document.getElementById('modelLabel');
const modelDesc = document.getElementById('modelDesc');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

init();

async function init() {
    bindEvents();
    await loadConfig();
    updateModelInfo();
}

function bindEvents() {
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);
    btnRemove.addEventListener('click', (event) => {
        event.stopPropagation();
        clearFile();
    });
    btnProcess.addEventListener('click', processFile);
    btnNewTask.addEventListener('click', resetPage);
    modelSelect.addEventListener('change', updateModelInfo);
}

async function loadConfig() {
    try {
        const response = await fetch('/task/pdf2docx/config');
        if (!response.ok) {
            throw new Error(`配置加载失败: ${response.status}`);
        }
        const data = await response.json();
        modelConfig = data.models || {};
        defaultModel = data.default_model || defaultModel;
    } catch (error) {
        console.error(error);
        modelConfig = {
            'google/gemini-3-flash-preview': {
                label: getModelDisplayName('google/gemini-3-flash-preview'),
                description: '速度更快，适合常规 PDF / 图片转 Word 场景。',
            },
            'google/gemini-3.1-pro-preview': {
                label: getModelDisplayName('google/gemini-3.1-pro-preview'),
                description: '更强调复杂版面与细节理解，适合高难度文档。',
            },
        };
    }

    modelSelect.innerHTML = '';
    Object.entries(modelConfig).forEach(([value, info]) => {
        modelSelect.add(new Option(getModelDisplayName(info.label || value), value));
    });
    modelSelect.value = modelConfig[defaultModel] ? defaultModel : Object.keys(modelConfig)[0];
}

function updateModelInfo() {
    const model = modelSelect.value;
    const info = modelConfig[model] || {};
    modelLabel.textContent = getModelDisplayName(info.label || model);
    modelDesc.textContent = info.description || '';
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
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(event) {
    const files = event.target.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFile(file) {
    const allowedExtensions = ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'];
    const lowerName = file.name.toLowerCase();
    const matched = allowedExtensions.some((ext) => lowerName.endsWith(ext));
    if (!matched) {
        alert('仅支持 PDF、PNG、JPG、JPEG、BMP、GIF、WEBP 文件。');
        return;
    }

    selectedFile = file;
    fileName.textContent = `${file.name} (${formatFileSize(file.size)})`;
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    btnProcess.disabled = false;
}

function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    btnProcess.disabled = true;
}

async function processFile() {
    if (!selectedFile) {
        return;
    }

    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    updateProgress(5, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('file', selectedFile);

        const params = new URLSearchParams({
            model: modelSelect.value,
        });

        const response = await fetch(`/task/pdf2docx?${params.toString()}`, {
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
        alert(`处理失败: ${error.message}`);
        resetPage();
    }
}

function startPolling(taskId) {
    stopPolling();
    pollStatus(taskId);
    pollingTimer = setInterval(() => pollStatus(taskId), 2000);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
}

async function pollStatus(taskId) {
    try {
        const response = await fetch(`/task/pdf2docx/status/${taskId}`);
        if (!response.ok) {
            throw new Error(`状态查询失败: ${response.status}`);
        }

        const data = await response.json();
        updateProgress(data.progress || 0, data.message || '正在处理中...');

        if (data.status === 'done') {
            stopPolling();
            showResult(data.result || {});
        } else if (data.status === 'failed') {
            stopPolling();
            throw new Error(data.error || '转换失败');
        }
    } catch (error) {
        stopPolling();
        alert(`处理失败: ${error.message}`);
        resetPage();
    }
}

function updateProgress(percent, message) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;
    processingStatus.textContent = message;
}

function showResult(result) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';

    resultStats.innerHTML = `
        <div class="stat-card">
            <i class="fas fa-file-alt"></i>
            <h3>${escapeHtml(result.filename || '-')}</h3>
            <p>源文件</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-robot"></i>
            <h3>${escapeHtml(getModelDisplayName(result.model || '-'))}</h3>
            <p>使用模型</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-file-word"></i>
            <h3>DOCX</h3>
            <p>输出格式</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-check-circle"></i>
            <h3>成功</h3>
            <p>任务状态</p>
        </div>
    `;

    resultGrid.innerHTML = `
        <div class="result-item">
            <h3>下载结果</h3>
            <div class="download-links">
                <a href="/${result.output_docx}" download class="download-btn">
                    <i class="fas fa-file-word"></i> 下载 Word 文档
                </a>
                <a href="/${result.raw_output_txt}" download class="download-btn">
                    <i class="fas fa-file-lines"></i> 下载原始 OCR 文本
                </a>
            </div>
        </div>
    `;
}

function resetPage() {
    clearFile();
    stopPolling();
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    updateProgress(0, '正在提交任务...');
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
    if (size < 1024) {
        return `${size} B`;
    }
    if (size < 1024 * 1024) {
        return `${(size / 1024).toFixed(1)} KB`;
    }
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
