const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const fileSelection = document.getElementById('fileSelection');
const fileList = document.getElementById('fileList');
const fileSummary = document.getElementById('fileSummary');
const clearFilesButton = document.getElementById('clearFiles');
const submitButton = document.getElementById('submitButton');
const outputFormat = document.getElementById('outputFormat');
const taskPanel = document.getElementById('taskPanel');
const taskList = document.getElementById('taskList');
const taskSummary = document.getElementById('taskSummary');
const taskCaption = document.getElementById('taskCaption');
const downloadAllButton = document.getElementById('downloadAll');
const newBatchButton = document.getElementById('newBatch');

let selectedFiles = [];
let taskStates = [];
let pollTimer = null;
let isSubmitting = false;
let submittedFormat = 'word';
let config = { max_files: 50, upload_max_mb: 95, default_output_format: 'word' };

const terminalStatuses = new Set(['done', 'failed', 'cancelled']);

function escapeHtml(value) {
    return String(value ?? '')
        .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;').replaceAll("'", '&#39;');
}

function formatBytes(bytes) {
    const size = Number(bytes || 0);
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
    return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

function totalSelectedBytes() {
    return selectedFiles.reduce((sum, file) => sum + Number(file.size || 0), 0);
}

function fileIdentity(file) {
    return `${file.name}\u0000${file.size}\u0000${file.lastModified}`;
}

function validateFiles(files) {
    const invalid = files.find((file) => !file.name.toLowerCase().endsWith('.msg'));
    if (invalid) return `仅支持 .msg 文件：“${invalid.name}”格式不正确。`;
    const identities = new Set(selectedFiles.map(fileIdentity));
    const uniqueNewFiles = files.filter((file) => !identities.has(fileIdentity(file)));
    if (selectedFiles.length + uniqueNewFiles.length > config.max_files) {
        return `单次最多选择 ${config.max_files} 个 MSG 文件。`;
    }
    const nextTotal = totalSelectedBytes() + uniqueNewFiles.reduce((sum, file) => sum + Number(file.size || 0), 0);
    if (nextTotal > config.upload_max_mb * 1024 * 1024) {
        return `单次上传总大小不能超过 ${config.upload_max_mb} MB，当前选择后约 ${formatBytes(nextTotal)}。`;
    }
    return '';
}

function addFiles(fileCollection) {
    const files = Array.from(fileCollection || []);
    if (!files.length) return;
    const error = validateFiles(files);
    if (error) {
        window.alert(error);
        return;
    }
    const existing = new Set(selectedFiles.map(fileIdentity));
    for (const file of files) {
        const identity = fileIdentity(file);
        if (!existing.has(identity)) {
            selectedFiles.push(file);
            existing.add(identity);
        }
    }
    renderSelectedFiles();
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    renderSelectedFiles();
}

function clearSelectedFiles() {
    selectedFiles = [];
    fileInput.value = '';
    renderSelectedFiles();
}

function renderSelectedFiles() {
    const hasFiles = selectedFiles.length > 0;
    fileSelection.hidden = !hasFiles;
    submitButton.disabled = !hasFiles || isSubmitting;
    if (!hasFiles) {
        fileList.innerHTML = '';
        return;
    }
    fileSummary.textContent = `已选择 ${selectedFiles.length} 个文件，合计 ${formatBytes(totalSelectedBytes())}`;
    fileList.innerHTML = selectedFiles.map((file, index) => `
        <div class="file-row">
            <div class="file-icon"><i class="fas fa-envelope"></i></div>
            <div><div class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</div><div class="file-size">${formatBytes(file.size)}</div></div>
            <button class="remove-file" type="button" data-index="${index}" aria-label="删除 ${escapeHtml(file.name)}"><i class="fas fa-xmark"></i></button>
        </div>`).join('');
}

function statusInfo(status) {
    if (status === 'done') return { icon: 'fa-check', cls: 'done', text: '已完成' };
    if (status === 'failed') return { icon: 'fa-xmark', cls: 'failed', text: '失败' };
    if (status === 'cancelled') return { icon: 'fa-ban', cls: 'failed', text: '已取消' };
    if (status === 'running' || status === 'processing') return { icon: 'fa-spinner fa-spin', cls: '', text: '处理中' };
    return { icon: 'fa-clock', cls: '', text: '排队中' };
}

function resultDownloadLink(task, key, label, icon) {
    const path = task.result?.[key];
    if (!path || !task.task_id) return '';
    const name = String(path).replaceAll('\\', '/').split('/').pop() || label;
    const params = new URLSearchParams({ file_path: path, download_name: name });
    return `<a class="download-link" href="/task/${encodeURIComponent(task.task_id)}/download?${params.toString()}"><i class="fas ${icon}"></i> ${label}</a>`;
}

function renderTasks() {
    const total = taskStates.length;
    const done = taskStates.filter((task) => task.status === 'done').length;
    const failed = taskStates.filter((task) => ['failed', 'cancelled'].includes(task.status)).length;
    const active = Math.max(total - done - failed, 0);
    taskSummary.innerHTML = [
        `<span class="summary-pill">总计 ${total}</span>`,
        `<span class="summary-pill">完成 ${done}</span>`,
        active ? `<span class="summary-pill">处理中 ${active}</span>` : '',
        failed ? `<span class="summary-pill">失败 ${failed}</span>` : '',
    ].join('');
    taskCaption.textContent = total && active === 0 ? '本批任务已全部结束，可单独下载或打包下载。' : '任务由服务器队列依次处理，页面可保持打开。';

    if (!total) {
        taskList.innerHTML = '<div class="empty-list">暂无任务</div>';
        downloadAllButton.hidden = true;
        return;
    }
    taskList.innerHTML = taskStates.map((task) => {
        const info = statusInfo(task.status);
        const progress = task.status === 'done' ? 100 : Math.max(0, Math.min(100, Number(task.progress || 0)));
        const downloads = task.status === 'done'
            ? `${resultDownloadLink(task, 'output_docx', '下载 Word', 'fa-file-word')}${resultDownloadLink(task, 'output_pdf', '下载 PDF', 'fa-file-pdf')}`
            : '';
        const subject = task.result?.subject ? `主题：${task.result.subject}` : '';
        const bodyInfo = task.result ? `正文来源：${String(task.result.body_format || '-').toUpperCase()} · 内嵌图片 ${task.result.inline_image_count || 0} 张` : '';
        const details = [subject, bodyInfo].filter(Boolean).join('；');
        const warnings = Array.isArray(task.result?.warnings) && task.result.warnings.length
            ? `<div class="warning-box"><strong>转换提示</strong>\n${escapeHtml(task.result.warnings.join('\n'))}</div>` : '';
        const log = task.status === 'failed'
            ? `<div class="error-log">${escapeHtml(task.error || task.message || '转换失败')}${task.stream_log ? `\n\n${escapeHtml(task.stream_log)}` : ''}</div>` : '';
        return `<article class="task-row">
            <div class="task-head">
                <div class="task-status-icon ${info.cls}"><i class="fas ${info.icon}"></i></div>
                <div><div class="file-name" title="${escapeHtml(task.filename)}">${escapeHtml(task.filename)}</div><div class="task-message">${escapeHtml(task.message || info.text)}${details ? `<br>${escapeHtml(details)}` : ''}</div></div>
                <strong>${escapeHtml(info.text)} · ${progress}%</strong>
            </div>
            <div class="progress-track"><div class="progress-fill" style="width:${progress}%"></div></div>
            ${downloads ? `<div class="task-result">${downloads}</div>` : ''}${warnings}${log}
        </article>`;
    }).join('');

    downloadAllButton.hidden = done === 0;
    downloadAllButton.disabled = done === 0;
    downloadAllButton.innerHTML = `<i class="fas fa-box-archive"></i> 批量下载（${done}）`;
}

async function readError(response) {
    try {
        const payload = await response.json();
        return typeof payload.detail === 'string' ? payload.detail : (payload.detail?.error || payload.message || '');
    } catch (_) {
        return '';
    }
}

async function submitTasks() {
    if (isSubmitting || !selectedFiles.length) return;
    const validationError = validateFiles([]);
    if (validationError) {
        window.alert(validationError);
        return;
    }
    isSubmitting = true;
    submitButton.disabled = true;
    outputFormat.disabled = true;
    submittedFormat = outputFormat.value;
    taskPanel.hidden = false;
    taskStates = selectedFiles.map((file) => ({ filename: file.name, status: 'queued', progress: 0, message: '正在提交…' }));
    renderTasks();
    taskPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });

    const formData = new FormData();
    const isSingle = selectedFiles.length === 1;
    if (isSingle) formData.append('file', selectedFiles[0]);
    else selectedFiles.forEach((file) => formData.append('files', file));
    const endpoint = isSingle ? '/task/msg-convert' : '/task/msg-convert/batch';
    const params = new URLSearchParams({ output_format: submittedFormat });

    try {
        const response = await fetch(`${endpoint}?${params.toString()}`, { method: 'POST', body: formData });
        if (!response.ok) throw new Error((await readError(response)) || `提交失败（${response.status}）`);
        const payload = await response.json();
        if (isSingle) {
            taskStates = [{
                filename: selectedFiles[0].name,
                task_id: payload.task_id,
                status: 'queued', progress: 0,
                message: payload.deduped ? '已复用正在处理的相同任务' : '已进入任务队列',
            }];
        } else {
            taskStates = (payload.tasks || []).map((task) => ({
                filename: task.filename || '未命名.msg', task_id: task.task_id,
                status: task.status === 'ACCEPTED' ? 'queued' : 'failed', progress: 0,
                message: task.status === 'ACCEPTED' ? (task.deduped ? '已复用相同任务' : '已进入任务队列') : (task.error || '提交失败'),
                error: task.error || '', batch_id: task.batch_id || payload.batch_id,
            }));
        }
        renderTasks();
        await pollTasks();
        startPolling();
    } catch (error) {
        taskStates = [{ filename: selectedFiles.map((file) => file.name).join('、'), status: 'failed', progress: 0, message: error.message, error: error.message }];
        renderTasks();
    } finally {
        isSubmitting = false;
        submitButton.disabled = true;
    }
}

async function pollOne(task) {
    if (!task.task_id || terminalStatuses.has(task.status)) return;
    try {
        const response = await fetch(`/task/msg-convert/status/${encodeURIComponent(task.task_id)}`);
        if (!response.ok) throw new Error((await readError(response)) || `状态查询失败（${response.status}）`);
        const payload = await response.json();
        task.status = payload.status === 'processing' ? 'running' : payload.status;
        task.progress = Number(payload.progress || 0);
        task.message = payload.message || task.message;
        task.result = payload.result || null;
        task.error = payload.error || '';
        task.stream_log = payload.stream_log || '';
    } catch (error) {
        task.message = `状态更新失败：${error.message}`;
    }
}

async function pollTasks() {
    await Promise.all(taskStates.map(pollOne));
    renderTasks();
    if (taskStates.length && taskStates.every((task) => !task.task_id || terminalStatuses.has(task.status))) stopPolling();
}

function startPolling() {
    stopPolling();
    if (taskStates.some((task) => task.task_id && !terminalStatuses.has(task.status))) {
        pollTimer = window.setInterval(pollTasks, 2000);
    }
}

function stopPolling() {
    if (pollTimer) window.clearInterval(pollTimer);
    pollTimer = null;
}

async function downloadAll() {
    const completed = taskStates.filter((task) => task.status === 'done' && task.task_id);
    if (!completed.length) return;
    const extensions = submittedFormat === 'both' ? ['.docx', '.pdf'] : submittedFormat === 'pdf' ? ['.pdf'] : ['.docx'];
    downloadAllButton.disabled = true;
    downloadAllButton.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 正在打包…';
    try {
        const response = await fetch('/task/batch-download', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ task_ids: completed.map((task) => task.task_id), extensions, archive_name: 'MSG转文档批量结果.zip' }),
        });
        if (!response.ok) throw new Error((await readError(response)) || `打包失败（${response.status}）`);
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url; anchor.download = 'MSG转文档批量结果.zip'; document.body.appendChild(anchor); anchor.click(); anchor.remove();
        URL.revokeObjectURL(url);
    } catch (error) {
        window.alert(`批量下载失败：${error.message}`);
    } finally {
        renderTasks();
    }
}

function prepareNextBatch() {
    clearSelectedFiles();
    outputFormat.disabled = false;
    submitButton.disabled = true;
    dropZone.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

async function loadConfig() {
    try {
        const response = await fetch('/task/msg-convert/config');
        if (!response.ok) return;
        config = { ...config, ...(await response.json()) };
        outputFormat.value = config.default_output_format || 'word';
        document.getElementById('limitMetric').textContent = `${config.max_files} 个 / ${config.upload_max_mb} MB`;
        dropZone.querySelector('span').textContent = `最多 ${config.max_files} 个文件，单次总大小不超过 ${config.upload_max_mb} MB`;
    } catch (_) {}
}

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', (event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); fileInput.click(); } });
dropZone.addEventListener('dragover', (event) => { event.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (event) => { event.preventDefault(); dropZone.classList.remove('dragover'); addFiles(event.dataTransfer.files); });
fileInput.addEventListener('change', () => { addFiles(fileInput.files); fileInput.value = ''; });
fileList.addEventListener('click', (event) => { const button = event.target.closest('.remove-file'); if (button) removeFile(Number(button.dataset.index)); });
clearFilesButton.addEventListener('click', clearSelectedFiles);
submitButton.addEventListener('click', submitTasks);
downloadAllButton.addEventListener('click', downloadAll);
newBatchButton.addEventListener('click', prepareNextBatch);

loadConfig();
renderSelectedFiles();
