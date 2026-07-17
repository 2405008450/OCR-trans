'use strict';

let config = { allowed_extensions: ['.docx','.doc','.xlsx','.xls','.pptx','.ppt'], upload_max_mb: 95, max_files: 50 };
let selectedFiles = [];
let tasks = [];
let pollTimer = null;

const byId = (id) => document.getElementById(id);
const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
const selectedStyle = (name) => document.querySelector(`input[name="${name}"]:checked`)?.value || 'british';
const formatSize = (bytes) => bytes < 1024 * 1024 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`;

document.querySelectorAll('.tab-button').forEach((button) => button.addEventListener('click', () => {
    document.querySelectorAll('.tab-button').forEach((item) => item.classList.toggle('active', item === button));
    document.querySelectorAll('.tab-panel').forEach((panel) => panel.classList.toggle('active', panel.id === button.dataset.tab));
}));

async function loadConfig() {
    try {
        const response = await fetch('/task/english-variant/config');
        if (!response.ok) throw new Error(`配置加载失败：${response.status}`);
        config = await response.json();
        const stats = config.stats || {};
        byId('dictionaryInfo').textContent = `词库 ${config.dictionary_version} · ${stats.british_to_american_rules || 0}/${stats.american_to_british_rules || 0} 条规则`;
        byId('uploadHint').textContent = `支持 ${config.allowed_extensions.join(' / ').toUpperCase()}，最多 ${config.max_files} 个，总计不超过 ${config.upload_max_mb} MB`;
    } catch (error) {
        byId('dictionaryInfo').textContent = '词库配置加载失败';
        console.error(error);
    }
}

function renderSummary(element, result) {
    const ambiguous = result.ambiguous_hits || [];
    element.hidden = false;
    element.innerHTML = `<strong>已替换 ${result.replacement_count || 0} 处</strong>，涉及 ${result.distinct_rule_count || 0} 条规则；跳过 ${result.ambiguous_hit_count || 0} 处歧义词。` +
        (ambiguous.length ? `<ul class="ambiguity-list">${ambiguous.slice(0, 8).map((item) => `<li>${esc(item.term)}：${esc(item.candidates.join(' / '))}（${item.count} 处）</li>`).join('')}</ul>` : '');
}

byId('convertTextButton').addEventListener('click', async () => {
    const button = byId('convertTextButton');
    button.disabled = true;
    try {
        const response = await fetch('/task/english-variant/text', {
            method: 'POST', headers: {'Content-Type':'application/json'},
            body: JSON.stringify({ text: byId('sourceText').value, target_style: selectedStyle('textStyle') }),
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.detail || '转换失败');
        byId('resultText').value = result.converted_text;
        renderSummary(byId('textSummary'), result);
    } catch (error) { alert(error.message); }
    finally { button.disabled = false; }
});

byId('copyTextButton').addEventListener('click', async () => {
    const value = byId('resultText').value;
    if (!value) return;
    await navigator.clipboard.writeText(value);
    byId('copyTextButton').textContent = '已复制';
    setTimeout(() => { byId('copyTextButton').innerHTML = '<i class="fas fa-copy"></i> 复制结果'; }, 1200);
});

const dropzone = byId('dropzone');
dropzone.addEventListener('click', () => byId('fileInput').click());
dropzone.addEventListener('dragover', (event) => { event.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (event) => { event.preventDefault(); dropzone.classList.remove('dragover'); addFiles(event.dataTransfer.files); });
byId('fileInput').addEventListener('change', (event) => addFiles(event.target.files));

function addFiles(fileList) {
    const existing = new Set(selectedFiles.map((file) => `${file.name}:${file.size}:${file.lastModified}`));
    for (const file of fileList) {
        const ext = `.${file.name.split('.').pop().toLowerCase()}`;
        const key = `${file.name}:${file.size}:${file.lastModified}`;
        if (config.allowed_extensions.includes(ext) && !existing.has(key) && selectedFiles.length < config.max_files) {
            selectedFiles.push(file); existing.add(key);
        }
    }
    renderFiles();
}

function renderFiles() {
    byId('fileList').innerHTML = selectedFiles.map((file, index) => `<div class="file-row"><div><strong>${esc(file.name)}</strong><br><small>${formatSize(file.size)}</small></div><button class="variant-button secondary" type="button" data-remove="${index}">移除</button></div>`).join('');
    byId('fileList').querySelectorAll('[data-remove]').forEach((button) => button.addEventListener('click', () => { selectedFiles.splice(Number(button.dataset.remove), 1); renderFiles(); }));
    byId('submitFilesButton').disabled = selectedFiles.length === 0;
}

byId('clearFilesButton').addEventListener('click', () => { selectedFiles = []; byId('fileInput').value = ''; renderFiles(); });

byId('submitFilesButton').addEventListener('click', async () => {
    const total = selectedFiles.reduce((sum, file) => sum + file.size, 0);
    if (total > config.upload_max_mb * 1024 * 1024) return alert(`文件总大小不能超过 ${config.upload_max_mb} MB`);
    const button = byId('submitFilesButton'); button.disabled = true;
    const form = new FormData();
    const isBatch = selectedFiles.length > 1;
    selectedFiles.forEach((file) => form.append(isBatch ? 'files' : 'file', file));
    try {
        const response = await fetch(`/task/english-variant${isBatch ? '/batch' : ''}?target_style=${encodeURIComponent(selectedStyle('fileStyle'))}`, {method:'POST', body:form});
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.detail || '提交失败');
        tasks = isBatch ? payload.tasks.filter((item) => item.task_id) : [{filename:selectedFiles[0].name, task_id:payload.task_id, status:'ACCEPTED'}];
        renderTasks(); startPolling();
    } catch (error) { alert(error.message); button.disabled = false; }
});

function statusLabel(status) { return ({queued:'排队中',processing:'处理中',done:'已完成',failed:'失败',cancelled:'已取消',ACCEPTED:'已提交'})[status] || status; }
function downloadUrl(task, result) {
    const params = new URLSearchParams({file_path:result.output_file, download_name:result.output_filename || '转换结果'});
    return `/task/${encodeURIComponent(task.task_id)}/download?${params}`;
}
function renderTasks() {
    byId('taskList').innerHTML = tasks.map((task) => {
        const result = task.result || {};
        const summary = task.status === 'done' ? `替换 ${result.replacement_count || 0} 处，涉及 ${result.distinct_rule_count || 0} 条词汇规则，跳过歧义 ${result.ambiguous_hit_count || 0} 处` : (task.message || '');
        const link = task.status === 'done' && result.output_file ? `<a class="download-link" href="${downloadUrl(task,result)}"><i class="fas fa-download"></i> 下载</a>` : '';
        return `<div class="task-row"><div class="task-copy"><strong>${esc(task.filename || task.task_id)}</strong><small>${esc(statusLabel(task.status))}${summary ? ` · ${esc(summary)}` : ''}</small></div>${link}</div>`;
    }).join('');
    byId('batchDownloadButton').hidden = !tasks.some((task) => task.status === 'done');
}

function startPolling() {
    if (pollTimer) clearInterval(pollTimer);
    pollTasks(); pollTimer = setInterval(pollTasks, 1500);
}
async function pollTasks() {
    await Promise.all(tasks.map(async (task) => {
        if (['done','failed','cancelled'].includes(task.status)) return;
        try {
            const response = await fetch(`/task/english-variant/status/${encodeURIComponent(task.task_id)}`);
            if (!response.ok) return;
            const payload = await response.json();
            Object.assign(task, payload);
        } catch (_) {}
    }));
    renderTasks();
    if (tasks.length && tasks.every((task) => ['done','failed','cancelled'].includes(task.status))) {
        clearInterval(pollTimer); pollTimer = null; byId('submitFilesButton').disabled = false;
    }
}

byId('batchDownloadButton').addEventListener('click', async () => {
    const doneIds = tasks.filter((task) => task.status === 'done').map((task) => task.task_id);
    const response = await fetch('/task/batch-download', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({task_ids:doneIds, archive_name:'英美式英语转换结果.zip'})});
    if (!response.ok) return alert('打包下载失败');
    const blob = await response.blob(); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=url; a.download='英美式英语转换结果.zip'; a.click(); URL.revokeObjectURL(url);
});

loadConfig();
