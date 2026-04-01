// ========== 全局变量 ==========
let selectedFiles = [];

// ========== DOM 元素 ==========
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview = document.getElementById('filePreview');
const previewImage = document.getElementById('previewImage');
const fileName = document.getElementById('fileName');
const fileStatus = document.getElementById('fileStatus');
const btnRemove = document.getElementById('btnRemove');
const btnProcess = document.getElementById('btnProcess');
const btnNewTask = document.getElementById('btnNewTask');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const batchSection = document.getElementById('batchSection');

const fromLang = document.getElementById('fromLang');
const toLang = document.getElementById('toLang');
const cardSide = document.getElementById('cardSide');
const enableVisualization = document.getElementById('enableVisualization');

const fileListPanel = document.getElementById('fileListPanel');
const fileListItems = document.getElementById('fileListItems');
const fileCount = document.getElementById('fileCount');

const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

function checkedOrDefault(el, defaultValue = false) {
    return el ? !!el.checked : defaultValue;
}

function valueOrDefault(el, defaultValue = '') {
    return el ? el.value : defaultValue;
}

// ========== 事件监听 ==========
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);
btnRemove.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFiles();
});
btnProcess.addEventListener('click', processFiles);
btnNewTask.addEventListener('click', resetApp);
document.getElementById('btnBatchNewTask')?.addEventListener('click', resetApp);

// ========== 文件处理函数 ==========
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.style.borderColor = 'var(--primary-color)';
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.style.borderColor = 'var(--border-color)';
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) addFiles(files);
}

function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    if (files.length > 0) addFiles(files);
}

function addFiles(newFiles) {
    const validTypes = ['image/jpeg', 'image/jpg', 'image/png', 'image/bmp', 'image/tiff'];
    const maxSize = 50 * 1024 * 1024;

    for (const file of newFiles) {
        if (!validTypes.includes(file.type) && !file.type.startsWith('image/')) {
            alert(`不支持的文件类型：${file.name}`);
            continue;
        }
        if (file.size > maxSize) {
            alert(`文件太大：${file.name}（最大 50MB）`);
            continue;
        }
        if (!selectedFiles.some((f) => f.name === file.name && f.size === file.size)) {
            selectedFiles.push(file);
        }
    }

    if (selectedFiles.length === 0) return;

    if (selectedFiles.length === 1) {
        const reader = new FileReader();
        reader.onload = (e) => { previewImage.src = e.target.result; };
        reader.readAsDataURL(selectedFiles[0]);
        previewImage.style.display = '';
    } else {
        previewImage.style.display = 'none';
    }

    fileName.textContent = selectedFiles.length === 1
        ? selectedFiles[0].name
        : `已选择 ${selectedFiles.length} 个文件`;
    fileStatus.textContent = selectedFiles.length === 1
        ? '文件已就绪，可直接开始处理'
        : `共 ${selectedFiles.length} 个文件，将逐个创建任务并行处理`;

    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    btnProcess.disabled = false;
    renderFileList();
}

function renderFileList() {
    if (selectedFiles.length <= 1) {
        fileListPanel.style.display = 'none';
        return;
    }
    fileListPanel.style.display = '';
    fileCount.textContent = selectedFiles.length;
    fileListItems.innerHTML = selectedFiles.map((file, idx) => `
        <div style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-radius:10px;background:linear-gradient(180deg,#f8fbfe,#f1f6fb);border:1px solid rgba(88,112,138,0.12);">
            <i class="fas fa-image" style="color:var(--primary-color);"></i>
            <span style="flex:1;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(file.name)}</span>
            <span style="font-size:12px;color:var(--text-secondary);">${formatFileSize(file.size)}</span>
            <button onclick="removeFileAt(${idx})" style="border:none;background:none;color:var(--danger-color);cursor:pointer;font-size:14px;padding:4px;" title="移除"><i class="fas fa-times"></i></button>
        </div>
    `).join('');
}

function removeFileAt(idx) {
    selectedFiles.splice(idx, 1);
    if (selectedFiles.length === 0) {
        clearFiles();
    } else {
        fileName.textContent = selectedFiles.length === 1
            ? selectedFiles[0].name
            : `已选择 ${selectedFiles.length} 个文件`;
        fileStatus.textContent = selectedFiles.length === 1
            ? '文件已就绪，可直接开始处理'
            : `共 ${selectedFiles.length} 个文件，将逐个创建任务并行处理`;
        if (selectedFiles.length === 1) {
            const reader = new FileReader();
            reader.onload = (e) => { previewImage.src = e.target.result; };
            reader.readAsDataURL(selectedFiles[0]);
            previewImage.style.display = '';
        }
        renderFileList();
    }
}

function clearFiles() {
    selectedFiles = [];
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    fileListPanel.style.display = 'none';
    btnProcess.disabled = true;
}

// ========== 构建参数 ==========
function buildParams() {
    const params = new URLSearchParams({
        from_lang: valueOrDefault(fromLang, 'zh'),
        to_lang: valueOrDefault(toLang, 'en'),
        enable_visualization: checkedOrDefault(enableVisualization, true),
        card_side: valueOrDefault(cardSide, 'front'),
    });
    return params;
}

// ========== 处理文件 ==========
async function processFiles() {
    if (selectedFiles.length === 0) return;

    if (selectedFiles.length === 1) {
        await processSingleFile();
    } else {
        await processBatchFiles();
    }
}

async function processSingleFile() {
    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    ensureEtaHint();

    let progress = 0;
    const progressInterval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        updateProgress(progress);
    }, 500);

    try {
        const formData = new FormData();
        formData.append('file', selectedFiles[0]);
        const params = buildParams();

        const submitResp = await fetch(`/task/run?${params}`, { method: 'POST', body: formData });
        if (!submitResp.ok) {
            let errorMessage = `提交失败: ${submitResp.status}`;
            try { const ed = await submitResp.json(); errorMessage = ed?.detail?.error || ed?.detail || ed?.message || errorMessage; } catch (_) {}
            throw new Error(errorMessage);
        }

        const submitData = await submitResp.json();
        const taskId = submitData.task_id;

        const queueOverlay = document.getElementById('queueOverlay');
        const MAX_POLLS = 150;
        let polls = 0;
        const result = await new Promise((resolve, reject) => {
            let pollInterval = null;
            const doPoll = async () => {
                polls++;
                if (polls > MAX_POLLS) { clearInterval(pollInterval); if (queueOverlay) queueOverlay.style.display = 'none'; reject(new Error('处理超时，请重试')); return; }
                try {
                    const statusResp = await fetch(`/task/run/status/${taskId}`);
                    if (!statusResp.ok) { clearInterval(pollInterval); if (queueOverlay) queueOverlay.style.display = 'none'; reject(new Error(`状态查询失败: ${statusResp.status}`)); return; }
                    const statusData = await statusResp.json();
                    if (statusData.status === 'queued') { if (queueOverlay) queueOverlay.style.display = 'flex'; }
                    else if (statusData.status === 'processing') { if (queueOverlay) queueOverlay.style.display = 'none'; }
                    else if (statusData.status === 'done') { clearInterval(pollInterval); if (queueOverlay) queueOverlay.style.display = 'none'; resolve(statusData.result); }
                    else if (statusData.status === 'error' || statusData.status === 'failed') { clearInterval(pollInterval); if (queueOverlay) queueOverlay.style.display = 'none'; reject(new Error(statusData.error || statusData.message || '处理失败')); }
                } catch (e) { clearInterval(pollInterval); if (queueOverlay) queueOverlay.style.display = 'none'; reject(e); }
            };
            doPoll();
            pollInterval = setInterval(doPoll, 2000);
        });

        clearInterval(progressInterval);
        updateProgress(100);
        setTimeout(() => displayResult(result), 500);
    } catch (error) {
        clearInterval(progressInterval);
        const message = error?.message === 'Failed to fetch'
            ? '无法连接后端接口。请确认当前页面是否由 FastAPI 服务提供，并检查后端服务、端口、反向代理或浏览器控制台中的网络/CORS 错误。'
            : error.message;
        alert(`处理失败: ${message}`);
        resetApp();
    }
}

async function processBatchFiles() {
    uploadSection.style.display = 'none';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'block';

    const params = buildParams();
    const formData = new FormData();
    selectedFiles.forEach((file) => formData.append('files', file));

    try {
        const resp = await fetch(`/task/run/batch?${params}`, { method: 'POST', body: formData });
        if (!resp.ok) {
            let msg = `提交失败: ${resp.status}`;
            try { const ed = await resp.json(); msg = ed?.detail?.error || ed?.detail || msg; } catch (_) {}
            throw new Error(msg);
        }

        const data = await resp.json();
        const tasks = data.tasks || [];
        renderBatchStats(tasks);
        startBatchPolling(tasks);
    } catch (error) {
        alert(`批量提交失败: ${error.message}`);
        resetApp();
    }
}

let batchPollingTimer = null;

function renderBatchStats(tasks) {
    const total = tasks.length;
    const accepted = tasks.filter((t) => t.status === 'ACCEPTED').length;
    const failed = tasks.filter((t) => t.status === 'FAILED').length;
    document.getElementById('batchStats').innerHTML = `
        <div class="stat-card"><i class="fas fa-files"></i><h3>${total}</h3><p>总文件数</p></div>
        <div class="stat-card"><i class="fas fa-check-circle"></i><h3>${accepted}</h3><p>已提交</p></div>
        ${failed > 0 ? `<div class="stat-card"><i class="fas fa-times-circle"></i><h3>${failed}</h3><p>提交失败</p></div>` : ''}
    `;
}

function startBatchPolling(tasks) {
    stopBatchPolling();
    const taskStates = tasks.map((t) => ({
        filename: t.filename,
        task_id: t.task_id,
        submitStatus: t.status,
        error: t.error || null,
        status: t.status === 'ACCEPTED' ? 'queued' : 'failed',
        progress: 0,
        message: t.status === 'ACCEPTED' ? '排队中...' : (t.error || '提交失败'),
        result: null,
    }));

    renderBatchTaskList(taskStates);

    batchPollingTimer = setInterval(async () => {
        let allDone = true;
        for (const task of taskStates) {
            if (!task.task_id || task.status === 'done' || task.status === 'failed' || task.status === 'cancelled') continue;
            allDone = false;
            try {
                const resp = await fetch(`/task/run/status/${task.task_id}`);
                if (!resp.ok) continue;
                const data = await resp.json();
                task.status = data.status === 'processing' ? 'running' : data.status;
                task.progress = data.progress || 0;
                task.message = data.message || '';
                if (data.status === 'done') {
                    task.status = 'done';
                    task.progress = 100;
                    task.result = data.result;
                } else if (data.status === 'failed') {
                    task.status = 'failed';
                    task.message = data.error || '处理失败';
                }
            } catch (_) {}
        }
        renderBatchTaskList(taskStates);
        if (allDone || taskStates.every((t) => ['done', 'failed', 'cancelled'].includes(t.status) || !t.task_id)) {
            stopBatchPolling();
        }
    }, 2000);
}

function stopBatchPolling() {
    if (batchPollingTimer) { clearInterval(batchPollingTimer); batchPollingTimer = null; }
}

function renderBatchTaskList(taskStates) {
    const container = document.getElementById('batchTaskList');
    const doneCount = taskStates.filter((t) => t.status === 'done').length;
    const failedCount = taskStates.filter((t) => t.status === 'failed' || (t.submitStatus === 'FAILED')).length;
    const totalCount = taskStates.length;

    const batchStatsEl = document.getElementById('batchStats');
    batchStatsEl.innerHTML = `
        <div class="stat-card"><i class="fas fa-files"></i><h3>${totalCount}</h3><p>总文件数</p></div>
        <div class="stat-card"><i class="fas fa-check-circle" style="color:var(--success-color)"></i><h3>${doneCount}</h3><p>已完成</p></div>
        <div class="stat-card"><i class="fas fa-spinner fa-spin" style="color:var(--primary-color)"></i><h3>${totalCount - doneCount - failedCount}</h3><p>处理中</p></div>
        ${failedCount > 0 ? `<div class="stat-card"><i class="fas fa-times-circle" style="color:var(--danger-color)"></i><h3>${failedCount}</h3><p>失败</p></div>` : ''}
    `;

    container.innerHTML = taskStates.map((task) => {
        const statusIcon = task.status === 'done' ? '<i class="fas fa-check-circle" style="color:var(--success-color)"></i>'
            : task.status === 'failed' ? '<i class="fas fa-times-circle" style="color:var(--danger-color)"></i>'
            : task.status === 'running' ? '<i class="fas fa-spinner fa-spin" style="color:var(--primary-color)"></i>'
            : '<i class="fas fa-clock" style="color:var(--warning-color)"></i>';

        const downloadLinks = task.status === 'done' && task.result?.results
            ? task.result.results.map((item) =>
                item.translated_image ? `<a href="/${item.translated_image}" download style="color:var(--success-color);font-size:12px;text-decoration:none;"><i class="fas fa-download"></i> 下载</a>` : ''
            ).join(' ')
            : '';

        return `<div style="display:flex;align-items:center;gap:14px;padding:14px 18px;border-radius:16px;background:var(--card-bg);border:1px solid var(--border-color);box-shadow:var(--shadow-sm);">
            ${statusIcon}
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
            ${downloadLinks ? `<div style="display:flex;gap:8px;">${downloadLinks}</div>` : ''}
        </div>`;
    }).join('');
}

function updateProgress(percent) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;

    if (percent < 30) processingStatus.textContent = '正在读取身份证文件...';
    else if (percent < 60) processingStatus.textContent = '正在进行OCR识别...';
    else if (percent < 90) processingStatus.textContent = '正在翻译文本...';
    else processingStatus.textContent = '正在生成结果...';
}

// ========== 显示结果 ==========
function displayResult(data) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';

    resultStats.innerHTML = `
        <div class="stat-card"><i class="fas fa-file-alt"></i><h3>${data.filename}</h3><p>文件名</p></div>
        <div class="stat-card"><i class="fas fa-stamp"></i><h3>身份证</h3><p>证件类型</p></div>
        <div class="stat-card"><i class="fas fa-images"></i><h3>${data.total_images}</h3><p>处理图片数</p></div>
        <div class="stat-card"><i class="fas fa-check-circle"></i><h3>成功</h3><p>处理状态</p></div>
    `;

    let gridHtml = '';
    data.results.forEach((item, index) => {
        gridHtml += `
            <div class="result-item">
                <h3>图片 ${index + 1}</h3>
                <div class="image-comparison">
                    ${item.corrected_image ? `<div class="image-box"><h4>矫正后的图片</h4><img src="/${item.corrected_image}" alt="矫正后" onclick="window.open('/${item.corrected_image}', '_blank')"></div>` : ''}
                    ${item.visualization_image ? `<div class="image-box"><h4>OCR识别可视化</h4><img src="/${item.visualization_image}" alt="可视化" onclick="window.open('/${item.visualization_image}', '_blank')"></div>` : ''}
                    <div class="image-box"><h4>翻译后的图片</h4><img src="/${item.translated_image}" alt="翻译后" onclick="window.open('/${item.translated_image}', '_blank')"></div>
                </div>
                <div class="download-links">
                    <a href="/${item.translated_image}" download class="download-btn"><i class="fas fa-download"></i> 下载翻译图片</a>
                    ${item.ocr_json ? `<a href="/${item.ocr_json}" download class="download-btn"><i class="fas fa-file-code"></i> 下载OCR数据</a>` : ''}
                    ${item.translated_json ? `<a href="/${item.translated_json}" download class="download-btn"><i class="fas fa-language"></i> 下载翻译数据</a>` : ''}
                </div>
            </div>`;
    });
    resultGrid.innerHTML = gridHtml;
}

// ========== 重置应用 ==========
function resetApp() {
    clearFiles();
    stopBatchPolling();
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    batchSection.style.display = 'none';
    progressFill.style.width = '0%';
    progressText.textContent = '0%';
    const queueOverlay = document.getElementById('queueOverlay');
    if (queueOverlay) queueOverlay.style.display = 'none';
}

function ensureEtaHint() {}

function escapeHtml(value) {
    const div = document.createElement('div');
    div.textContent = value ?? '';
    return div.innerHTML;
}

function formatFileSize(size) {
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
    return `${(size / 1024 / 1024).toFixed(1)} MB`;
}
