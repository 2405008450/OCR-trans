const POLL_MS = 1600;

let currentTaskId = '';
let pollTimer = null;

const statusTextMap = {
  queued: '排队中',
  processing: '统计中',
  done: '已完成',
  failed: '失败',
  cancelled: '已取消',
};

const fileStatusText = {
  counted: '已统计',
  needs_ocr: '需 OCR',
  needs_cad_parser: '需 CAD 解析',
  skipped: '已跳过',
  failed: '失败',
};

function init() {
  document.getElementById('submitBtn').addEventListener('click', submitTask);
  document.getElementById('resetBtn').addEventListener('click', resetPage);
  loadConfig();
}

async function loadConfig() {
  try {
    const response = await fetch('/task/word-count/config');
    if (!response.ok) return;
    const config = await response.json();
    renderConfig(config);
  } catch (error) {
    console.error('loadConfig', error);
  }
}

function renderConfig(config) {
  const roots = Array.isArray(config.allowed_roots) ? config.allowed_roots : [];
  const rootsEl = document.getElementById('allowedRoots');
  rootsEl.innerHTML = roots.length
    ? roots.map((item) => {
        const scopeOnly = Boolean(item.scope_only);
        const icon = scopeOnly ? 'fa-circle-info' : (item.exists ? 'fa-circle-check' : 'fa-triangle-exclamation');
        const color = scopeOnly ? 'var(--cyan)' : (item.exists ? 'var(--green)' : 'var(--amber)');
        const labelBase = scopeOnly ? `${item.path}（白名单前缀）` : item.path;
        const label = item.mount_path ? `${labelBase} -> ${item.mount_path}` : labelBase;
        return `<div class="root-item"><i class="fas ${icon}" style="color:${color}"></i><span title="${escAttr(label)}">${escHtml(label)}</span></div>`;
      }).join('')
    : '<div class="root-item"><i class="fas fa-triangle-exclamation" style="color:var(--amber)"></i><span>未读取到允许根目录</span></div>';

  const countable = (config.countable_extensions || []).join(' / ');
  const images = (config.image_extensions || []).join(' / ');
  const cad = (config.cad_extensions || []).join(' / ');
  document.getElementById('formatHint').innerHTML = [
    `实际统计：${escHtml(countable)}`,
    `扩展入口：图片 ${escHtml(images)}；CAD ${escHtml(cad)}`,
    `限制：最多 ${Number(config.max_files || 0).toLocaleString()} 个文件，单文件 ${Number(config.max_file_mb || 0)} MB`,
  ].join('<br>');
}

async function submitTask() {
  const directoryPath = document.getElementById('directoryPath').value.trim();
  if (!directoryPath) {
    alert('请填写目录路径');
    return;
  }

  const submitBtn = document.getElementById('submitBtn');
  submitBtn.disabled = true;
  setStatus({ status: 'queued', progress: 0, message: '正在提交任务...' });

  const extensions = parseExtensions(document.getElementById('extensionsInput').value);
  const body = {
    directory_path: directoryPath,
    recursive: document.getElementById('recursiveInput').checked,
    include_hidden: document.getElementById('hiddenInput').checked,
    extensions,
  };

  try {
    const response = await fetch('/task/word-count', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const payload = await response.json();
    if (!response.ok) {
      alert(payload.detail || '提交失败');
      submitBtn.disabled = false;
      setStatus({ status: 'failed', progress: 0, message: payload.detail || '提交失败' });
      return;
    }
    currentTaskId = payload.task_id;
    document.getElementById('taskIdText').textContent = currentTaskId ? `任务 ${currentTaskId.slice(0, 8)}` : '';
    startPolling();
    await loadStatus();
  } catch (error) {
    console.error('submitTask', error);
    alert('提交失败，请检查服务是否可用');
    submitBtn.disabled = false;
    setStatus({ status: 'failed', progress: 0, message: '提交失败' });
  }
}

function parseExtensions(raw) {
  const items = String(raw || '')
    .split(/[,\s，、]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => item.startsWith('.') ? item.toLowerCase() : `.${item.toLowerCase()}`);
  return items.length ? items : null;
}

function startPolling() {
  stopPolling();
  pollTimer = setInterval(loadStatus, POLL_MS);
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

async function loadStatus() {
  if (!currentTaskId) return;
  try {
    const response = await fetch(`/task/word-count/status/${encodeURIComponent(currentTaskId)}`);
    if (!response.ok) return;
    const task = await response.json();
    setStatus(task);
    renderResult(task.result, task);
    if (['done', 'failed', 'cancelled'].includes(task.status)) {
      stopPolling();
      document.getElementById('submitBtn').disabled = false;
    }
  } catch (error) {
    console.error('loadStatus', error);
  }
}

function setStatus(task) {
  const progress = Number(task.progress || 0);
  document.getElementById('progressFill').style.width = `${Math.max(0, Math.min(100, progress))}%`;
  document.getElementById('progressText').textContent = `${Math.round(progress)}%`;
  const status = statusTextMap[task.status] || task.status || '等待中';
  const message = task.message ? `：${task.message}` : '';
  document.getElementById('statusText').textContent = `${status}${message}`;
}

function renderResult(result, task) {
  if (!result || typeof result !== 'object') {
    return;
  }
  const summary = result.summary || {};
  document.getElementById('mainWords').textContent = formatNumber(summary.total_main_word_count);
  document.getElementById('extraWords').textContent = formatNumber(summary.total_extra_word_count);
  document.getElementById('countedFiles').textContent = formatNumber(summary.counted_files);
  const issueCount = Number(summary.failed_files || 0)
    + Number(summary.skipped_files || 0)
    + Number(summary.needs_ocr_files || 0)
    + Number(summary.needs_cad_parser_files || 0);
  document.getElementById('issueFiles').textContent = formatNumber(issueCount);

  renderDownloads(result, task);
  renderFiles(result.files || []);
}

function renderDownloads(result, task) {
  const area = document.getElementById('downloadArea');
  const links = [];
  if (result.report_excel) {
    links.push(`<a class="download-link" href="${downloadUrl(task.task_id, result.report_excel, '字数统计报告.xlsx')}"><i class="fas fa-file-excel"></i> 下载 Excel</a>`);
  }
  if (result.report_json) {
    links.push(`<a class="download-link secondary" href="${downloadUrl(task.task_id, result.report_json, '字数统计结果.json')}"><i class="fas fa-code"></i> 下载 JSON</a>`);
  }
  area.innerHTML = links.join('');
  area.style.display = links.length ? 'flex' : 'none';
}

function renderFiles(files) {
  const tbody = document.getElementById('fileRows');
  if (!files.length) {
    tbody.innerHTML = '<tr><td colspan="5"><div class="empty">暂无文件明细</div></td></tr>';
    return;
  }
  const visible = files.slice(0, 200);
  tbody.innerHTML = visible.map((item) => {
    const status = item.status || '';
    const statusLabel = fileStatusText[status] || status || '-';
    const message = item.error || item.warning || item.message || '';
    return `<tr>
      <td title="${escAttr(item.file_path || '')}">${escHtml(item.relative_path || item.filename || '-')}</td>
      <td><span class="badge ${escAttr(status)}">${escHtml(statusLabel)}</span></td>
      <td>${formatNumber(item.main_word_count)}</td>
      <td>${formatNumber(item.extra_word_count)}</td>
      <td>${escHtml(message)}</td>
    </tr>`;
  }).join('');
}

function downloadUrl(taskId, filePath, name) {
  return `/task/${encodeURIComponent(taskId)}/download?file_path=${encodeURIComponent(filePath)}&download_name=${encodeURIComponent(name)}`;
}

function resetPage() {
  stopPolling();
  currentTaskId = '';
  document.getElementById('directoryPath').value = '';
  document.getElementById('extensionsInput').value = '';
  document.getElementById('recursiveInput').checked = true;
  document.getElementById('hiddenInput').checked = false;
  document.getElementById('taskIdText').textContent = '';
  document.getElementById('downloadArea').style.display = 'none';
  document.getElementById('fileRows').innerHTML = '<tr><td colspan="5"><div class="empty">暂无结果</div></td></tr>';
  document.getElementById('mainWords').textContent = '0';
  document.getElementById('extraWords').textContent = '0';
  document.getElementById('countedFiles').textContent = '0';
  document.getElementById('issueFiles').textContent = '0';
  document.getElementById('submitBtn').disabled = false;
  setStatus({ status: '', progress: 0, message: '等待提交' });
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function escHtml(value) {
  const div = document.createElement('div');
  div.textContent = value ?? '';
  return div.innerHTML;
}

function escAttr(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

init();
