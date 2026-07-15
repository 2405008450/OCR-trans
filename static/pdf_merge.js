const POLL_MS = 1400;

let discoveredFiles = [];
let currentTaskId = '';
let pollTimer = null;
let draggedIndex = null;

const statusTextMap = {
  queued: '排队中',
  processing: '合并中',
  done: '已完成',
  failed: '失败',
  cancelled: '已取消',
};

function init() {
  document.getElementById('scanBtn').addEventListener('click', scanDirectory);
  document.getElementById('resetBtn').addEventListener('click', resetPage);
  document.getElementById('submitBtn').addEventListener('click', submitMerge);
  document.getElementById('selectAllBtn').addEventListener('click', () => setAllSelected(true));
  document.getElementById('clearSelectionBtn').addEventListener('click', () => setAllSelected(false));
  loadConfig();
}

async function loadConfig() {
  try {
    const response = await fetch('/task/pdf-merge/config');
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
  document.getElementById('limitHint').textContent = `限制：最多 ${Number(config.max_files || 0)} 个文件，单文件 ${Number(config.max_file_mb || 0)} MB，总计 ${Number(config.max_total_mb || 0)} MB。结果保存到本地后提供下载。`;
}

async function scanDirectory() {
  const directoryPath = document.getElementById('directoryPath').value.trim();
  if (!directoryPath) {
    showScanMessage('请先填写共享目录路径。', 'error');
    return;
  }

  const scanBtn = document.getElementById('scanBtn');
  scanBtn.disabled = true;
  showScanMessage('正在扫描共享目录，请稍候...', '');
  try {
    const response = await fetch('/task/pdf-merge/discover', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        directory_path: directoryPath,
        recursive: document.getElementById('recursiveInput').checked,
      }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '扫描失败');

    discoveredFiles = (payload.files || []).map((item) => ({ ...item, selected: true }));
    renderFileList();
    if (payload.truncated) {
      showScanMessage(`已显示前 ${payload.max_files} 个 PDF，目录中的文件数量超过上限，请缩小扫描范围。`, 'warning');
    } else if (!discoveredFiles.length) {
      showScanMessage('该目录中没有找到 PDF 文件。', 'warning');
    } else {
      showScanMessage(`已找到 ${discoveredFiles.length} 个 PDF，默认全部选中。`, 'success');
    }
  } catch (error) {
    discoveredFiles = [];
    renderFileList();
    showScanMessage(error.message || '扫描失败，请检查共享路径和服务权限。', 'error');
  } finally {
    scanBtn.disabled = false;
  }
}

function renderFileList() {
  const list = document.getElementById('fileList');
  if (!discoveredFiles.length) {
    list.innerHTML = '<div class="empty"><i class="fas fa-file-pdf"></i><br>没有可显示的 PDF 文件。</div>';
  } else {
    let selectedOrder = 0;
    list.innerHTML = discoveredFiles.map((item, index) => {
      if (item.selected) selectedOrder += 1;
      const order = item.selected ? selectedOrder : '-';
      return `<div class="file-row ${item.selected ? 'is-selected' : 'is-unselected'}" draggable="true" data-index="${index}">
        <span class="drag-handle" title="拖动调整顺序"><i class="fas fa-grip-vertical"></i></span>
        <input class="file-select" type="checkbox" ${item.selected ? 'checked' : ''} aria-label="选择 ${escAttr(item.name || item.relative_path)}" />
        <div class="file-copy" title="${escAttr(item.relative_path)}">
          <div class="file-name"><span class="order-number">${order}.</span> ${escHtml(item.name || item.relative_path)}</div>
          <div class="file-path">${escHtml(item.relative_path)}</div>
        </div>
        <div class="file-size">${formatBytes(item.size)}</div>
        <div class="order-actions">
          <button class="icon-btn move-up" title="上移" ${index === 0 ? 'disabled' : ''}><i class="fas fa-chevron-up"></i></button>
          <button class="icon-btn move-down" title="下移" ${index === discoveredFiles.length - 1 ? 'disabled' : ''}><i class="fas fa-chevron-down"></i></button>
        </div>
      </div>`;
    }).join('');
    bindFileRowEvents();
  }
  updateSelectionSummary();
}

function bindFileRowEvents() {
  document.querySelectorAll('.file-row').forEach((row) => {
    const index = Number(row.dataset.index);
    row.querySelector('.file-select').addEventListener('change', (event) => {
      discoveredFiles[index].selected = event.target.checked;
      renderFileList();
    });
    row.querySelector('.move-up').addEventListener('click', () => moveFile(index, -1));
    row.querySelector('.move-down').addEventListener('click', () => moveFile(index, 1));
    row.addEventListener('dragstart', () => {
      draggedIndex = index;
      row.classList.add('is-dragging');
    });
    row.addEventListener('dragend', () => {
      draggedIndex = null;
      row.classList.remove('is-dragging');
    });
    row.addEventListener('dragover', (event) => event.preventDefault());
    row.addEventListener('drop', (event) => {
      event.preventDefault();
      if (draggedIndex === null || draggedIndex === index) return;
      const [moved] = discoveredFiles.splice(draggedIndex, 1);
      discoveredFiles.splice(index, 0, moved);
      renderFileList();
    });
  });
}

function moveFile(index, offset) {
  const target = index + offset;
  if (target < 0 || target >= discoveredFiles.length) return;
  [discoveredFiles[index], discoveredFiles[target]] = [discoveredFiles[target], discoveredFiles[index]];
  renderFileList();
}

function setAllSelected(selected) {
  discoveredFiles.forEach((item) => { item.selected = selected; });
  renderFileList();
}

function updateSelectionSummary() {
  const selected = discoveredFiles.filter((item) => item.selected);
  const totalSize = selected.reduce((sum, item) => sum + Number(item.size || 0), 0);
  document.getElementById('selectedCount').textContent = selected.length.toLocaleString();
  document.getElementById('selectedSize').textContent = formatBytes(totalSize);
  document.getElementById('submitBtn').disabled = selected.length < 2 || Boolean(currentTaskId && pollTimer);
  document.getElementById('selectAllBtn').disabled = !discoveredFiles.length;
  document.getElementById('clearSelectionBtn').disabled = !discoveredFiles.length;
}

async function submitMerge() {
  const selected = discoveredFiles.filter((item) => item.selected);
  if (selected.length < 2) {
    showScanMessage('请至少选择 2 个 PDF 文件。', 'error');
    return;
  }
  const outputFilename = document.getElementById('outputFilename').value.trim() || '合并结果.pdf';
  const submitBtn = document.getElementById('submitBtn');
  submitBtn.disabled = true;
  setStatus({ status: 'queued', progress: 0, message: '正在提交任务...' });

  try {
    const response = await fetch('/task/pdf-merge', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        directory_path: document.getElementById('directoryPath').value.trim(),
        relative_paths: selected.map((item) => item.relative_path),
        output_filename: outputFilename,
      }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '提交失败');
    currentTaskId = payload.task_id;
    document.getElementById('taskIdText').textContent = currentTaskId ? `任务 ${currentTaskId.slice(0, 8)}` : '';
    startPolling();
    await loadStatus();
  } catch (error) {
    showScanMessage(error.message || '提交失败，请检查服务是否可用。', 'error');
    setStatus({ status: 'failed', progress: 0, message: error.message || '提交失败' });
    submitBtn.disabled = false;
  }
}

function startPolling() {
  stopPolling();
  pollTimer = setInterval(loadStatus, POLL_MS);
  updateSelectionSummary();
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
    const response = await fetch(`/task/pdf-merge/status/${encodeURIComponent(currentTaskId)}`);
    if (!response.ok) return;
    const task = await response.json();
    setStatus(task);
    renderResult(task.result, task);
    if (['done', 'failed', 'cancelled'].includes(task.status)) {
      stopPolling();
      updateSelectionSummary();
    }
  } catch (error) {
    console.error('loadStatus', error);
  }
}

function setStatus(task) {
  const progress = Math.max(0, Math.min(100, Number(task.progress || 0)));
  document.getElementById('progressFill').style.width = `${progress}%`;
  document.getElementById('progressText').textContent = `${Math.round(progress)}%`;
  const status = statusTextMap[task.status] || task.status || '等待中';
  const detail = task.error || task.message || '';
  document.getElementById('statusText').textContent = detail ? `${status}：${detail}` : status;
}

function renderResult(result, task) {
  if (!result || typeof result !== 'object') return;
  document.getElementById('resultFileCount').textContent = Number(result.input_file_count || 0).toLocaleString();
  document.getElementById('resultPageCount').textContent = Number(result.total_pages || 0).toLocaleString();
  document.getElementById('resultInputSize').textContent = formatBytes(result.input_total_size);
  document.getElementById('resultOutputSize').textContent = formatBytes(result.output_size);

  const area = document.getElementById('downloadArea');
  if (result.output_pdf) {
    area.innerHTML = `<a class="download-link" href="${downloadUrl(task.task_id, result.output_pdf, result.output_filename || '合并结果.pdf')}"><i class="fas fa-download"></i> 下载合并 PDF</a>`;
    area.style.display = 'flex';
  }

  const files = Array.isArray(result.files) ? result.files : [];
  const table = document.getElementById('resultTable');
  document.getElementById('resultRows').innerHTML = files.map((item, index) => `<tr><td>${index + 1}. ${escHtml(item.relative_path || item.filename || '-')}</td><td>${Number(item.page_count || 0).toLocaleString()}</td></tr>`).join('');
  table.classList.toggle('is-hidden', !files.length);
}

function resetPage() {
  stopPolling();
  currentTaskId = '';
  discoveredFiles = [];
  document.getElementById('directoryPath').value = '';
  document.getElementById('recursiveInput').checked = true;
  document.getElementById('outputFilename').value = '合并结果.pdf';
  document.getElementById('taskIdText').textContent = '';
  document.getElementById('downloadArea').style.display = 'none';
  document.getElementById('downloadArea').innerHTML = '';
  document.getElementById('resultTable').classList.add('is-hidden');
  document.getElementById('resultRows').innerHTML = '';
  ['resultFileCount', 'resultPageCount'].forEach((id) => { document.getElementById(id).textContent = '0'; });
  ['resultInputSize', 'resultOutputSize'].forEach((id) => { document.getElementById(id).textContent = '0 B'; });
  hideScanMessage();
  renderFileList();
  setStatus({ status: '', progress: 0, message: '等待提交' });
}

function showScanMessage(message, type) {
  const element = document.getElementById('scanMessage');
  element.textContent = message;
  element.className = `notice ${type || ''}`.trim();
}

function hideScanMessage() {
  document.getElementById('scanMessage').className = 'notice is-hidden';
}

function downloadUrl(taskId, filePath, name) {
  return `/task/${encodeURIComponent(taskId)}/download?file_path=${encodeURIComponent(filePath)}&download_name=${encodeURIComponent(name)}`;
}

function formatBytes(value) {
  let bytes = Number(value || 0);
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let unit = 0;
  while (bytes >= 1024 && unit < units.length - 1) {
    bytes /= 1024;
    unit += 1;
  }
  const digits = unit === 0 || bytes >= 100 ? 0 : (bytes >= 10 ? 1 : 2);
  return `${bytes.toFixed(digits)} ${units[unit]}`;
}

function escHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;',
  }[char]));
}

function escAttr(value) {
  return escHtml(value).replace(/`/g, '&#096;');
}

document.addEventListener('DOMContentLoaded', init);
