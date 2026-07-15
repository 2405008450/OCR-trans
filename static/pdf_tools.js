const POLL_MS = 1400;
const SUPPORTED_OPERATIONS = ['merge', 'split', 'compress', 'extract', 'delete', 'rotate'];

let discoveredFiles = [];
let selectedRelativePath = '';
let currentOperation = getInitialOperation();
let currentTaskOperation = '';
let currentTaskId = '';
let pollTimer = null;
let draggedIndex = null;
let compressionModes = {};

const statusTextMap = { queued: '排队中', processing: '处理中', done: '已完成', failed: '失败', cancelled: '已取消' };
const operationLabels = { merge: '合并', split: '拆分', compress: '压缩', extract: '提取页面', delete: '删除页面', rotate: '旋转页面' };

function getInitialOperation() {
  const requested = new URLSearchParams(window.location.search).get('operation');
  return SUPPORTED_OPERATIONS.includes(requested) ? requested : 'merge';
}

function init() {
  document.getElementById('scanBtn').addEventListener('click', scanDirectory);
  document.getElementById('resetBtn').addEventListener('click', resetPage);
  document.getElementById('submitBtn').addEventListener('click', submitTask);
  document.getElementById('selectAllBtn').addEventListener('click', () => setAllSelected(true));
  document.getElementById('clearSelectionBtn').addEventListener('click', () => setAllSelected(false));
  document.getElementById('splitMode').addEventListener('change', updateSplitMode);
  document.getElementById('compressionMode').addEventListener('change', updateCompressionHint);
  document.getElementById('pageSelectionMode').addEventListener('change', updatePageMode);
  document.querySelectorAll('.op-button').forEach((button) => button.addEventListener('click', () => setOperation(button.dataset.operation)));
  setOperation(currentOperation);
  loadConfig();
}

async function loadConfig() {
  try {
    const response = await fetch('/task/pdf-tools/config');
    if (!response.ok) return;
    const config = await response.json();
    compressionModes = config.compression_modes || {};
    renderConfig(config);
    updateCompressionHint();
  } catch (error) { console.error('loadConfig', error); }
}

function renderConfig(config) {
  const roots = Array.isArray(config.allowed_roots) ? config.allowed_roots : [];
  document.getElementById('allowedRoots').innerHTML = roots.map((item) => {
    const scopeOnly = Boolean(item.scope_only);
    const icon = scopeOnly ? 'fa-circle-info' : (item.exists ? 'fa-circle-check' : 'fa-triangle-exclamation');
    const color = scopeOnly ? 'var(--cyan)' : (item.exists ? 'var(--green)' : 'var(--amber)');
    const label = item.mount_path ? `${item.path} -> ${item.mount_path}` : item.path;
    return `<div class="root-item"><i class="fas ${icon}" style="color:${color}"></i><span>${escHtml(label)}</span></div>`;
  }).join('') || '<div class="root-item">未读取到允许根目录</div>';
  document.getElementById('limitHint').textContent = `最多扫描 ${Number(config.max_files || 0)} 个 PDF，单文件上限 ${Number(config.max_file_mb || 0)} MB，合并总大小上限 ${Number(config.max_total_mb || 0)} MB。`;
}

async function scanDirectory() {
  const directoryPath = document.getElementById('directoryPath').value.trim();
  if (!directoryPath) { showMessage('请先填写共享目录路径。', 'error'); return; }
  const button = document.getElementById('scanBtn');
  button.disabled = true;
  showMessage('正在扫描目录，请稍候...', '');
  try {
    const response = await fetch('/task/pdf-tools/discover', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ directory_path: directoryPath, recursive: document.getElementById('recursiveInput').checked }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '扫描失败');
    discoveredFiles = (payload.files || []).map((item) => ({ ...item, selected: true }));
    selectedRelativePath = discoveredFiles[0]?.relative_path || '';
    renderFiles();
    if (payload.truncated) {
      showMessage(`已显示前 ${payload.max_files} 个 PDF，目录文件超过上限，请缩小扫描范围。`, '');
    } else if (!discoveredFiles.length) {
      showMessage('目录中没有找到 PDF。', '');
    } else if (currentOperation === 'merge') {
      showMessage(`已找到 ${discoveredFiles.length} 个 PDF，默认全部选中，可调整合并顺序。`, 'success');
    } else {
      showMessage(`已找到 ${discoveredFiles.length} 个 PDF，请选择一个源文件。`, 'success');
    }
  } catch (error) {
    discoveredFiles = [];
    selectedRelativePath = '';
    renderFiles();
    showMessage(error.message || '扫描失败。', 'error');
  } finally {
    button.disabled = false;
    updateSubmitState();
  }
}

function renderFiles() {
  document.getElementById('fileCount').textContent = `${discoveredFiles.length} 个`;
  const list = document.getElementById('fileList');
  if (!discoveredFiles.length) {
    list.innerHTML = '<div class="empty"><i class="fas fa-file-pdf"></i><br>没有可选择的 PDF。</div>';
  } else if (currentOperation === 'merge') {
    let selectedOrder = 0;
    list.innerHTML = discoveredFiles.map((item, index) => {
      if (item.selected) selectedOrder += 1;
      const order = item.selected ? selectedOrder : '-';
      return `<div class="file-row merge-row ${item.selected ? 'is-selected' : 'is-unselected'}" draggable="true" data-index="${index}">
        <span class="drag-handle" title="拖动调整顺序"><i class="fas fa-grip-vertical"></i></span>
        <input class="file-select" type="checkbox" ${item.selected ? 'checked' : ''} aria-label="选择 ${escAttr(item.name || item.relative_path)}" />
        <span><span class="file-name"><span class="order-number">${order}.</span> ${escHtml(item.name || item.relative_path)}</span><span class="file-path">${escHtml(item.relative_path)}</span></span>
        <span class="file-size">${formatPageCount(item.page_count)}<br>${formatBytes(item.size)}</span>
        <span class="order-actions">
          <button class="icon-btn move-up" title="上移" ${index === 0 ? 'disabled' : ''}><i class="fas fa-chevron-up"></i></button>
          <button class="icon-btn move-down" title="下移" ${index === discoveredFiles.length - 1 ? 'disabled' : ''}><i class="fas fa-chevron-down"></i></button>
        </span>
      </div>`;
    }).join('');
    bindMergeFileEvents();
  } else {
    list.innerHTML = discoveredFiles.map((item) => `<label class="file-row ${item.relative_path === selectedRelativePath ? 'selected' : ''}">
      <input type="radio" name="sourcePdf" value="${escAttr(item.relative_path)}" ${item.relative_path === selectedRelativePath ? 'checked' : ''} />
      <span><span class="file-name">${escHtml(item.name || item.relative_path)}</span><span class="file-path">${escHtml(item.relative_path)}</span></span>
      <span class="file-size">${formatPageCount(item.page_count)}<br>${formatBytes(item.size)}</span>
    </label>`).join('');
    list.querySelectorAll('input[name="sourcePdf"]').forEach((input) => input.addEventListener('change', () => {
      selectedRelativePath = input.value;
      renderFiles();
      updateSuggestedNames();
    }));
  }
  updateSuggestedNames();
  updateSelectionSummary();
  updateSubmitState();
}

function bindMergeFileEvents() {
  document.querySelectorAll('.merge-row').forEach((row) => {
    const index = Number(row.dataset.index);
    row.querySelector('.file-select').addEventListener('change', (event) => {
      discoveredFiles[index].selected = event.target.checked;
      renderFiles();
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
      renderFiles();
    });
  });
}

function moveFile(index, offset) {
  const target = index + offset;
  if (target < 0 || target >= discoveredFiles.length) return;
  [discoveredFiles[index], discoveredFiles[target]] = [discoveredFiles[target], discoveredFiles[index]];
  renderFiles();
}

function setAllSelected(selected) {
  discoveredFiles.forEach((item) => { item.selected = selected; });
  renderFiles();
}

function updateSelectionSummary() {
  const selected = discoveredFiles.filter((item) => item.selected);
  const totalSize = selected.reduce((sum, item) => sum + Number(item.size || 0), 0);
  document.getElementById('selectedCount').textContent = selected.length.toLocaleString();
  document.getElementById('selectedSize').textContent = formatBytes(totalSize);
  document.getElementById('selectAllBtn').disabled = !discoveredFiles.length;
  document.getElementById('clearSelectionBtn').disabled = !discoveredFiles.length;
}

function setOperation(operation) {
  if (!SUPPORTED_OPERATIONS.includes(operation)) return;
  currentOperation = operation;
  document.querySelectorAll('.op-button').forEach((button) => button.classList.toggle('active', button.dataset.operation === operation));
  document.getElementById('mergeOptions').classList.toggle('is-hidden', operation !== 'merge');
  document.getElementById('splitOptions').classList.toggle('is-hidden', operation !== 'split');
  document.getElementById('compressOptions').classList.toggle('is-hidden', operation !== 'compress');
  document.getElementById('pageOptions').classList.toggle('is-hidden', !['extract', 'delete', 'rotate'].includes(operation));
  document.getElementById('mergeSelectionToolbar').classList.toggle('is-hidden', operation !== 'merge');
  document.getElementById('pageModeField').classList.toggle('is-hidden', operation === 'rotate');
  document.getElementById('angleField').classList.toggle('is-hidden', operation !== 'rotate');
  document.getElementById('pageSpec').placeholder = operation === 'rotate' ? '例如：1-3,5；全部页面可填写 all' : '例如：1-3,5,8';
  document.getElementById('pageSpecHint').textContent = operation === 'delete' ? '这里填写要删除的页面；至少保留一页。' : '逗号组合页码，短横线表示连续范围。';
  document.getElementById('submitBtn').innerHTML = `<i class="fas fa-play"></i> 开始${operationLabels[operation]}`;
  updatePageMode();
  renderFiles();
  if (discoveredFiles.length) {
    showMessage(
      operation === 'merge'
        ? `已找到 ${discoveredFiles.length} 个 PDF，按勾选顺序合并。`
        : `已找到 ${discoveredFiles.length} 个 PDF，请选择一个源文件。`,
      'success',
    );
  }
}

function updateSuggestedNames() {
  const selected = discoveredFiles.find((item) => item.relative_path === selectedRelativePath);
  const stem = (selected?.name || 'PDF').replace(/\.pdf$/i, '');
  document.getElementById('splitPrefix').value = `${stem}_拆分`;
  document.getElementById('compressFilename').value = `${stem}_压缩.pdf`;
  if (currentOperation !== 'merge') {
    document.getElementById('pageOutputFilename').value = `${stem}_${operationLabels[currentOperation] || '处理'}.pdf`;
  }
}

function updateSplitMode() {
  const custom = document.getElementById('splitMode').value === 'ranges';
  document.getElementById('pagesPerFileField').classList.toggle('is-hidden', custom);
  document.getElementById('pageGroupsField').classList.toggle('is-hidden', !custom);
}

function updateCompressionHint() {
  const mode = document.getElementById('compressionMode').value;
  document.getElementById('compressionHint').textContent = compressionModes[mode]?.description || '';
}

function updatePageMode() {
  const custom = currentOperation === 'rotate' || document.getElementById('pageSelectionMode').value === 'custom';
  document.getElementById('pageSpecField').classList.toggle('is-hidden', !custom);
}

function buildOptions() {
  if (currentOperation === 'split') return {
    split_mode: document.getElementById('splitMode').value,
    pages_per_file: Number(document.getElementById('pagesPerFile').value || 1),
    page_groups: document.getElementById('pageGroups').value.trim(),
    output_prefix: document.getElementById('splitPrefix').value.trim(),
  };
  if (currentOperation === 'compress') return {
    compression_mode: document.getElementById('compressionMode').value,
    output_filename: document.getElementById('compressFilename').value.trim(),
  };
  return {
    page_mode: currentOperation === 'rotate' ? 'custom' : document.getElementById('pageSelectionMode').value,
    page_spec: document.getElementById('pageSpec').value.trim(),
    output_filename: document.getElementById('pageOutputFilename').value.trim(),
    angle: Number(document.getElementById('rotateAngle').value || 90),
  };
}

async function submitTask() {
  const selected = discoveredFiles.filter((item) => item.selected);
  if (currentOperation === 'merge' && selected.length < 2) {
    showMessage('请至少选择 2 个 PDF 文件。', 'error');
    return;
  }
  if (currentOperation !== 'merge' && !selectedRelativePath) {
    showMessage('请先扫描并选择一个 PDF。', 'error');
    return;
  }
  const button = document.getElementById('submitBtn');
  button.disabled = true;
  clearResult();
  currentTaskId = '';
  document.getElementById('taskIdText').textContent = '';
  setStatus({ status: 'queued', progress: 0, message: '正在提交任务...' });
  currentTaskOperation = currentOperation;
  const isMerge = currentTaskOperation === 'merge';
  const endpoint = isMerge ? '/task/pdf-merge' : '/task/pdf-tools';
  const body = isMerge ? {
    directory_path: document.getElementById('directoryPath').value.trim(),
    relative_paths: selected.map((item) => item.relative_path),
    output_filename: document.getElementById('mergeFilename').value.trim() || '合并结果.pdf',
  } : {
    directory_path: document.getElementById('directoryPath').value.trim(),
    relative_path: selectedRelativePath,
    operation: currentTaskOperation,
    options: buildOptions(),
  };
  try {
    const response = await fetch(endpoint, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '提交失败');
    currentTaskId = payload.task_id;
    document.getElementById('taskIdText').textContent = `任务 ${currentTaskId.slice(0, 8)}`;
    startPolling();
    await loadStatus();
  } catch (error) {
    currentTaskOperation = '';
    setStatus({ status: 'failed', progress: 0, message: error.message || '提交失败' });
    showMessage(error.message || '提交失败。', 'error');
    updateSubmitState();
  }
}

function startPolling() {
  stopPolling();
  pollTimer = setInterval(loadStatus, POLL_MS);
  updateSubmitState();
}

function stopPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = null;
}

async function loadStatus() {
  if (!currentTaskId || !currentTaskOperation) return;
  const taskPath = currentTaskOperation === 'merge' ? 'pdf-merge' : 'pdf-tools';
  try {
    const response = await fetch(`/task/${taskPath}/status/${encodeURIComponent(currentTaskId)}`);
    if (!response.ok) return;
    const task = await response.json();
    setStatus(task);
    renderResult(task.result, task);
    if (['done', 'failed', 'cancelled'].includes(task.status)) {
      stopPolling();
      updateSubmitState();
    }
  } catch (error) { console.error('loadStatus', error); }
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
  const isMerge = currentTaskOperation === 'merge';
  const files = isMerge
    ? (Array.isArray(result.files) ? result.files : [])
    : (Array.isArray(result.output_files) ? result.output_files : []);
  document.getElementById('sourcePages').textContent = Number(isMerge ? result.total_pages : result.source_page_count || 0).toLocaleString();
  document.getElementById('outputFiles').textContent = Number(isMerge ? 1 : result.output_file_count || (result.output_pdf ? 1 : 0)).toLocaleString();
  document.getElementById('inputSize').textContent = formatBytes(isMerge ? result.input_total_size : result.input_size);
  document.getElementById('outputSize').textContent = formatBytes(result.output_size);
  document.getElementById('summaryText').textContent = result.summary_text || '';
  const links = [];
  if (result.output_pdf) {
    const label = isMerge ? '下载合并 PDF' : '下载处理结果';
    links.push(`<a class="action-link" href="${downloadUrl(task.task_id, result.output_pdf, result.output_filename || '处理结果.pdf')}" download><i class="fas fa-download"></i> ${label}</a>`);
  }
  if (result.archive_zip) links.push(`<a class="action-link" href="${downloadUrl(task.task_id, result.archive_zip, result.archive_filename || '拆分结果.zip')}" download><i class="fas fa-file-zipper"></i> 下载全部 ZIP</a>`);
  document.getElementById('downloadArea').innerHTML = links.join('');
  document.getElementById('resultRows').innerHTML = files.map((item, index) => `<tr><td>${isMerge ? `${index + 1}. ` : ''}${escHtml(item.relative_path || item.filename || '-')}</td><td>${Number(item.page_count || 0)}</td><td>${formatBytes(item.size)}</td></tr>`).join('');
  document.getElementById('resultTable').classList.toggle('is-hidden', !files.length);
}

function updateSubmitState() {
  const selectedCount = discoveredFiles.filter((item) => item.selected).length;
  const ready = currentOperation === 'merge' ? selectedCount >= 2 : Boolean(selectedRelativePath);
  document.getElementById('submitBtn').disabled = !ready || Boolean(pollTimer);
}

function clearResult() {
  document.getElementById('downloadArea').innerHTML = '';
  document.getElementById('summaryText').textContent = '';
  document.getElementById('resultTable').classList.add('is-hidden');
  document.getElementById('resultRows').innerHTML = '';
  ['sourcePages', 'outputFiles'].forEach((id) => { document.getElementById(id).textContent = '0'; });
  ['inputSize', 'outputSize'].forEach((id) => { document.getElementById(id).textContent = '0 B'; });
}

function resetPage() {
  stopPolling();
  currentTaskId = '';
  currentTaskOperation = '';
  discoveredFiles = [];
  selectedRelativePath = '';
  draggedIndex = null;
  document.getElementById('directoryPath').value = '';
  document.getElementById('recursiveInput').checked = true;
  document.getElementById('mergeFilename').value = '合并结果.pdf';
  document.getElementById('pageSelectionMode').value = 'custom';
  document.getElementById('taskIdText').textContent = '';
  document.getElementById('scanMessage').className = 'notice is-hidden';
  clearResult();
  renderFiles();
  updatePageMode();
  setStatus({ status: '', progress: 0, message: '等待提交' });
}

function showMessage(message, type) {
  const element = document.getElementById('scanMessage');
  element.textContent = message;
  element.className = `notice ${type || ''}`.trim();
}

function downloadUrl(taskId, path, name) {
  return `/task/${encodeURIComponent(taskId)}/download?file_path=${encodeURIComponent(path)}&download_name=${encodeURIComponent(name)}`;
}

function formatBytes(value) {
  let bytes = Number(value || 0);
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let unit = 0;
  while (bytes >= 1024 && unit < units.length - 1) { bytes /= 1024; unit += 1; }
  return `${bytes.toFixed(unit === 0 || bytes >= 100 ? 0 : bytes >= 10 ? 1 : 2)} ${units[unit]}`;
}

function formatPageCount(value) {
  const count = Number(value);
  return Number.isInteger(count) && count > 0 ? `${count.toLocaleString()} 页` : '页数未知';
}

function escHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char]));
}

function escAttr(value) { return escHtml(value).replace(/`/g, '&#096;'); }

document.addEventListener('DOMContentLoaded', init);
