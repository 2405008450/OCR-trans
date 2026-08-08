const POLL_INTERVAL_MS = 1500;

const el = {
    file: document.getElementById('audioFile'), dropZone: document.getElementById('dropZone'),
    fileSummary: document.getElementById('fileSummary'), fileName: document.getElementById('fileName'),
    fileMeta: document.getElementById('fileMeta'), preview: document.getElementById('audioPreview'),
    language: document.getElementById('languageSelect'), enableItn: document.getElementById('enableItn'),
    run: document.getElementById('runTranscription'), cancel: document.getElementById('cancelTask'),
    error: document.getElementById('errorBox'), progressPanel: document.getElementById('progressPanel'),
    progressTitle: document.getElementById('progressTitle'), progressFill: document.getElementById('progressFill'),
    progressMessage: document.getElementById('progressMessage'), progressPercent: document.getElementById('progressPercent'),
    resultPanel: document.getElementById('resultPanel'), resultMeta: document.getElementById('resultMeta'),
    downloadGrid: document.getElementById('downloadGrid'), transcript: document.getElementById('transcriptView'),
    timeline: document.getElementById('timelineView'), formatHint: document.getElementById('formatHint'),
};

let config = null, selectedFile = null, currentTaskId = null, pollTimer = null, previewUrl = null, taskActive = false;

function showError(message) { el.error.textContent = message || '发生未知错误'; el.error.style.display = 'block'; }
function clearError() { el.error.textContent = ''; el.error.style.display = 'none'; }
function syncRunState() { el.run.disabled = taskActive || !selectedFile || !config?.configured; }
function formatBytes(bytes) { return bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(2)} MB`; }
function formatTime(value) {
    const total = Math.max(0, Number(value) || 0), hours = Math.floor(total / 3600), minutes = Math.floor((total % 3600) / 60), seconds = total % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${seconds.toFixed(3).padStart(6, '0')}`;
}

function setSelectedFile(file) {
    clearError();
    if (!file) { syncRunState(); return; }
    const extension = `.${String(file.name).split('.').pop().toLowerCase()}`;
    const allowed = config?.allowed_extensions || [];
    const maxBytes = (config?.max_file_mb || 200) * 1048576;
    if (!allowed.includes(extension)) return showError(`不支持 ${extension} 格式，仅支持 ${allowed.join('、')}`);
    if (!file.size) return showError('音频文件不能为空');
    if (file.size > maxBytes) return showError(`文件超过 ${config?.max_file_mb || 200} MB 上传限制`);
    selectedFile = file; el.fileName.textContent = file.name; el.fileMeta.textContent = `${formatBytes(file.size)} · ${file.type || extension}`; el.fileSummary.style.display = 'block';
    if (previewUrl) URL.revokeObjectURL(previewUrl); previewUrl = URL.createObjectURL(file); el.preview.src = previewUrl; el.preview.style.display = 'block';
    syncRunState();
}

async function loadConfig() {
    const response = await fetch('/task/audio-transcription/config');
    if (!response.ok) throw new Error(`配置加载失败：${response.status}`);
    config = await response.json();
    el.language.innerHTML = '';
    Object.entries(config.languages || {}).forEach(([value, label]) => el.language.add(new Option(label, value)));
    el.language.value = config.defaults?.language || 'auto'; el.enableItn.checked = config.defaults?.enable_itn !== false;
    el.file.accept = (config.allowed_extensions || []).join(',');
    el.formatHint.textContent = `支持 ${(config.allowed_extensions || []).map(item => item.slice(1).toUpperCase()).join('、')}，单文件不超过 ${config.max_file_mb || 200} MB。`;
    if (!config.configured) showError('服务器尚未配置 DASHSCOPE_API_KEY，暂时无法提交转写。');
    syncRunState();
}

function updateProgress(task) {
    const progress = Math.max(0, Math.min(100, Number(task.progress || 0)));
    el.progressPanel.style.display = 'block'; el.progressFill.style.width = `${progress}%`; el.progressPercent.textContent = `${progress}%`;
    el.progressMessage.textContent = task.message || (task.status === 'queued' ? '任务排队中' : '处理中');
    el.progressTitle.textContent = task.status === 'queued' ? `任务排队中${task.queue_position ? `（第 ${task.queue_position} 位）` : ''}` : '正在转写音频';
}

function buildDownload(taskId, path, label, primary = false) {
    if (!path) return null;
    const link = document.createElement('a'); link.className = `download-link${primary ? ' primary' : ''}`;
    const name = path.split('/').pop() || label;
    link.href = `/task/${encodeURIComponent(taskId)}/download?${new URLSearchParams({file_path: path, download_name: name})}`;
    link.innerHTML = `<i class="fas fa-download"></i> ${label}`; return link;
}

function renderResult(result, taskId) {
    el.resultMeta.innerHTML = '';
    [`模型：${result.model_name || '-'}`, `字幕段：${result.segment_count ?? 0}`, `词/字：${result.word_count ?? 0}`, `语言：${result.language || 'auto'}`, `时间轴：${result.timeline_source === 'word_timestamps_resegmented' ? '逐词智能切分' : '模型原始句段'}`].forEach(text => { const item = document.createElement('span'); item.textContent = text; el.resultMeta.appendChild(item); });
    el.downloadGrid.innerHTML = '';
    const downloads = [
        [result.archive_zip, '下载全部 ZIP', true], [result.timeline_txt, '时间轴 TXT'], [result.plain_txt, '纯文本'],
        [result.srt, 'SRT 字幕'], [result.vtt, 'VTT 字幕'], [result.word_tsv, '逐词 TSV'], [result.result_json, '完整 JSON'],
    ];
    downloads.forEach(([path, label, primary]) => { const link = buildDownload(taskId, path, label, primary); if (link) el.downloadGrid.appendChild(link); });
    el.transcript.textContent = result.text || '没有识别到文本'; el.timeline.innerHTML = '';
    (result.segments || []).forEach(segment => {
        const row = document.createElement('div'); row.className = 'timeline-row';
        const time = document.createElement('div'); time.className = 'timeline-time'; time.textContent = `${formatTime(segment.start)} → ${formatTime(segment.end)}`;
        const text = document.createElement('div'); text.textContent = segment.text || ''; row.append(time, text); el.timeline.appendChild(row);
    });
    if (result.timestamps_truncated_in_preview) { const note = document.createElement('div'); note.className = 'muted'; note.textContent = '页面只展示前 500 个句段，下载文件包含完整结果。'; el.timeline.appendChild(note); }
    el.resultPanel.style.display = 'block'; el.resultPanel.scrollIntoView({behavior: 'smooth', block: 'start'});
}

function stopPolling() { if (pollTimer) clearInterval(pollTimer); pollTimer = null; }
async function pollStatus() {
    if (!currentTaskId) return;
    try {
        const response = await fetch(`/task/audio-transcription/status/${encodeURIComponent(currentTaskId)}`);
        if (!response.ok) throw new Error(`状态读取失败：${response.status}`);
        const task = await response.json(); updateProgress(task);
        if (task.status === 'done') {
            taskActive = false; stopPolling(); el.progressFill.style.width = '100%'; el.progressPercent.textContent = '100%'; el.progressTitle.textContent = '转写完成'; el.cancel.style.display = 'none'; syncRunState(); renderResult(task.result || {}, currentTaskId);
        } else if (task.status === 'failed' || task.status === 'cancelled') {
            taskActive = false; stopPolling(); showError(task.error || task.message || (task.status === 'cancelled' ? '任务已取消' : '任务处理失败')); el.cancel.style.display = 'none'; syncRunState();
        }
    } catch (error) { taskActive = false; stopPolling(); showError(error.message); syncRunState(); el.cancel.style.display = 'none'; }
}

async function submitTranscription() {
    clearError(); el.resultPanel.style.display = 'none';
    if (!selectedFile) return showError('请先选择一个音频文件');
    const data = new FormData(); data.append('file', selectedFile); data.append('language', el.language.value); data.append('enable_itn', el.enableItn.checked ? 'true' : 'false');
    taskActive = true; syncRunState(); updateProgress({status: 'queued', progress: 0, message: '正在上传音频'});
    try {
        const response = await fetch('/task/audio-transcription', {method: 'POST', body: data}); const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload.detail || `提交失败：${response.status}`);
        currentTaskId = payload.task_id; el.cancel.style.display = 'inline-flex'; stopPolling(); await pollStatus(); if (taskActive && !pollTimer) pollTimer = setInterval(pollStatus, POLL_INTERVAL_MS);
    } catch (error) { taskActive = false; showError(error.message); syncRunState(); el.cancel.style.display = 'none'; }
}

async function cancelTask() {
    if (!currentTaskId) return;
    el.cancel.disabled = true;
    try { const response = await fetch(`/task/${encodeURIComponent(currentTaskId)}/cancel`, {method: 'POST'}); if (!response.ok) throw new Error('取消请求失败'); el.progressMessage.textContent = '已提交取消请求'; }
    catch (error) { showError(error.message); } finally { el.cancel.disabled = false; }
}

function bindEvents() {
    el.dropZone.addEventListener('click', () => el.file.click()); el.dropZone.addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') el.file.click(); }); el.file.addEventListener('change', () => setSelectedFile(el.file.files?.[0]));
    ['dragenter','dragover'].forEach(name => el.dropZone.addEventListener(name, event => { event.preventDefault(); el.dropZone.classList.add('dragging'); }));
    ['dragleave','drop'].forEach(name => el.dropZone.addEventListener(name, event => { event.preventDefault(); el.dropZone.classList.remove('dragging'); })); el.dropZone.addEventListener('drop', event => setSelectedFile(event.dataTransfer?.files?.[0]));
    el.run.addEventListener('click', submitTranscription); el.cancel.addEventListener('click', cancelTask); window.addEventListener('beforeunload', () => { stopPolling(); if (previewUrl) URL.revokeObjectURL(previewUrl); });
}

async function init() { bindEvents(); try { await loadConfig(); } catch (error) { showError(error.message); el.run.disabled = true; } }
init();
