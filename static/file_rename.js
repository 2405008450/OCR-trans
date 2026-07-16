const POLL_INTERVAL_MS = 1500;
const statusTextMap = {
  queued: '排队中',
  processing: '处理中',
  done: '已完成',
  failed: '失败',
  cancelled: '已取消',
};
const previewStatusMap = {
  ready: '将复制',
  unmatched: '未匹配',
  unchanged: '未变化',
  invalid: '名称无效',
  conflict: '路径冲突',
};

let discoveredFiles = [];
let scanCompleted = false;
let currentMode = 'cleanup';
let previewData = null;
let previewFingerprint = '';
let currentTaskId = '';
let pollTimer = null;
const expandedFolders = new Set();

function init() {
  document.getElementById('scanBtn').addEventListener('click', scanFiles);
  document.getElementById('resetBtn').addEventListener('click', resetPage);
  document.getElementById('exportListBtn').addEventListener('click', exportFileList);
  document.getElementById('previewBtn').addEventListener('click', generatePreview);
  document.getElementById('executeBtn').addEventListener('click', executeCopy);
  document.getElementById('fileSearch').addEventListener('input', renderFiles);
  document.getElementById('selectAllBtn').addEventListener('click', () => setVisibleSelection(true));
  document.getElementById('clearSelectionBtn').addEventListener('click', () => setVisibleSelection(false));
  document.getElementById('expandAllBtn').addEventListener('click', () => setAllFoldersExpanded(true));
  document.getElementById('collapseAllBtn').addEventListener('click', () => setAllFoldersExpanded(false));
  document.querySelectorAll('.mode-button').forEach((button) => {
    button.addEventListener('click', () => setMode(button.dataset.mode || 'cleanup'));
  });
  ['directoryPath'].forEach((id) => {
    document.getElementById(id).addEventListener('input', invalidateScan);
  });
  ['recursiveInput', 'hiddenInput'].forEach((id) => {
    document.getElementById(id).addEventListener('change', invalidateScan);
  });
  ['patternInput', 'replacementInput'].forEach((id) => {
    document.getElementById(id).addEventListener('input', invalidatePreview);
  });
  document.getElementById('ignoreCaseInput').addEventListener('change', invalidatePreview);
  [
    'cleanupLeadingNumberInput',
    'cleanupSeparatorSpaceInput',
    'cleanupSeparatorUnderscoreInput',
    'cleanupDatetimeInput',
    'cleanupDatetimeCompactInput',
    'cleanupDatetimeDottedInput',
    'cleanupTranslatedInput',
  ].forEach((id) => {
    document.getElementById(id).addEventListener('change', () => {
      updateCleanupControlAvailability();
      invalidatePreview();
    });
  });
  ['cleanupMaxDigitsInput', 'cleanupSuffixInput'].forEach((id) => {
    document.getElementById(id).addEventListener('input', invalidatePreview);
  });
  updateCleanupControlAvailability();
  loadConfig();
  renderFiles();
  updateActions();
}

async function loadConfig() {
  try {
    const response = await fetch('/task/file-rename/config');
    if (!response.ok) return;
    const config = await response.json();
    const roots = Array.isArray(config.allowed_roots) ? config.allowed_roots : [];
    document.getElementById('allowedRoots').innerHTML = roots.length
      ? roots.map((item) => {
          const scopeOnly = Boolean(item.scope_only);
          const icon = scopeOnly ? 'fa-circle-info' : (item.exists ? 'fa-circle-check' : 'fa-triangle-exclamation');
          const color = scopeOnly ? 'var(--cyan)' : (item.exists ? 'var(--green)' : 'var(--amber)');
          const base = scopeOnly ? `${item.path}（白名单前缀）` : item.path;
          const label = item.mount_path ? `${base} → ${item.mount_path}` : base;
          return `<div class="root-item"><i class="fas ${icon}" style="color:${color}"></i><span title="${escAttr(label)}">${escHtml(label)}</span></div>`;
        }).join('')
      : '<div class="root-item"><i class="fas fa-triangle-exclamation" style="color:var(--amber)"></i><span>未读取到允许根目录</span></div>';
  } catch (error) {
    console.error('loadConfig', error);
  }
}

async function scanFiles() {
  const directoryPath = document.getElementById('directoryPath').value.trim();
  if (!directoryPath) {
    showMessage('scanMessage', '请先填写共享目录路径。', 'error');
    return;
  }
  const button = document.getElementById('scanBtn');
  button.disabled = true;
  showMessage('scanMessage', '正在扫描目录，请稍候...', '');
  try {
    const response = await fetch('/task/file-rename/discover', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        directory_path: directoryPath,
        recursive: document.getElementById('recursiveInput').checked,
        include_hidden: document.getElementById('hiddenInput').checked,
      }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '扫描失败');
    discoveredFiles = (payload.files || []).map((item) => ({ ...item, selected: true }));
    expandedFolders.clear();
    allFolderPaths().forEach((path) => expandedFolders.add(path));
    scanCompleted = true;
    invalidatePreview();
    renderFiles();
    let message = `已找到并选中 ${discoveredFiles.length} 个文件，共 ${formatFileSize(payload.total_size || 0)}。`;
    if (payload.truncated) message += ` 扫描结果已按 ${payload.max_files} 个上限截断，请缩小目录范围。`;
    showMessage('scanMessage', message, payload.truncated ? 'warning' : 'success');
  } catch (error) {
    discoveredFiles = [];
    expandedFolders.clear();
    scanCompleted = false;
    invalidatePreview();
    renderFiles();
    showMessage('scanMessage', error.message || '扫描失败，请检查路径和权限。', 'error');
  } finally {
    button.disabled = false;
    updateActions();
  }
}

function renderFiles() {
  const list = document.getElementById('fileList');
  const visible = visibleFiles();
  if (!visible.length) {
    const message = scanCompleted
      ? (discoveredFiles.length ? '没有符合当前筛选条件的文件。' : '目录中没有可处理文件。')
      : '扫描目录后选择需要处理的文件。';
    list.innerHTML = `<div class="empty">${escHtml(message)}</div>`;
  } else {
    const tree = buildFileTree(visible);
    const folderStats = buildFolderSelectionStats();
    const fileIndexes = new Map(discoveredFiles.map((item, index) => [item.relative_path, index]));
    const searchActive = Boolean(document.getElementById('fileSearch').value.trim());
    list.innerHTML = renderTreeChildren(tree, 0, folderStats, fileIndexes, searchActive);

    list.querySelectorAll('input[data-file-index]').forEach((input) => {
      input.addEventListener('change', () => {
        const item = discoveredFiles[Number(input.dataset.fileIndex)];
        if (item) item.selected = input.checked;
        invalidatePreview();
        renderFiles();
      });
    });
    list.querySelectorAll('button[data-folder-toggle]').forEach((button) => {
      button.addEventListener('click', () => {
        const folderPath = button.dataset.folderToggle || '';
        if (expandedFolders.has(folderPath)) expandedFolders.delete(folderPath);
        else expandedFolders.add(folderPath);
        renderFiles();
      });
    });
    list.querySelectorAll('input[data-folder-select]').forEach((input) => {
      const folderPath = input.dataset.folderSelect || '';
      const stats = folderStats.get(folderPath) || { total: 0, selected: 0 };
      input.indeterminate = stats.selected > 0 && stats.selected < stats.total;
      input.addEventListener('change', () => {
        discoveredFiles.forEach((item) => {
          if (isFileInsideFolder(item.relative_path, folderPath)) item.selected = input.checked;
        });
        invalidatePreview();
        renderFiles();
      });
    });
  }
  updateSelectionSummary();
}

function buildFileTree(files) {
  const root = { path: '', name: '', folders: new Map(), files: [] };
  files.forEach((item) => {
    const parts = String(item.relative_path || '').split('/').filter(Boolean);
    let node = root;
    const folderParts = parts.slice(0, -1);
    folderParts.forEach((folderName, index) => {
      if (!node.folders.has(folderName)) {
        node.folders.set(folderName, {
          name: folderName,
          path: folderParts.slice(0, index + 1).join('/'),
          folders: new Map(),
          files: [],
        });
      }
      node = node.folders.get(folderName);
    });
    node.files.push(item);
  });
  return root;
}

function renderTreeChildren(node, depth, folderStats, fileIndexes, searchActive) {
  const folders = Array.from(node.folders.values()).sort((left, right) => naturalCompare(left.name, right.name));
  const files = [...node.files].sort((left, right) => naturalCompare(left.name || left.relative_path, right.name || right.relative_path));
  const folderHtml = folders.map((folder) => {
    const stats = folderStats.get(folder.path) || { total: 0, selected: 0 };
    const expanded = searchActive || expandedFolders.has(folder.path);
    const checked = stats.total > 0 && stats.selected === stats.total;
    const children = expanded
      ? renderTreeChildren(folder, depth + 1, folderStats, fileIndexes, searchActive)
      : '';
    return `<div class="tree-folder">
      <div class="folder-row" style="padding-left:${11 + depth * 22}px">
        <button type="button" class="tree-toggle" data-folder-toggle="${escAttr(folder.path)}" title="${expanded ? '折叠文件夹' : '展开文件夹'}"><i class="fas ${expanded ? 'fa-chevron-down' : 'fa-chevron-right'}"></i></button>
        <input type="checkbox" data-folder-select="${escAttr(folder.path)}" ${checked ? 'checked' : ''} title="选择或取消此文件夹内的全部文件" />
        <button type="button" class="folder-name" data-folder-toggle="${escAttr(folder.path)}" title="${expanded ? '折叠文件夹' : '展开文件夹'}"><i class="fas ${expanded ? 'fa-folder-open' : 'fa-folder'}"></i>${escHtml(folder.name)}</button>
        <span class="folder-meta">${stats.selected}/${stats.total}</span>
      </div>
      ${children}
    </div>`;
  }).join('');
  const fileHtml = files.map((item) => {
    const index = fileIndexes.get(item.relative_path);
    return `<label class="file-row" style="padding-left:${35 + depth * 22}px">
      <input type="checkbox" data-file-index="${index}" ${item.selected ? 'checked' : ''} />
      <span><span class="file-name"><i class="far fa-file" style="margin-right:7px;color:#94a3b8"></i>${escHtml(item.name || item.relative_path)}</span><span class="file-path">${escHtml(item.relative_path)}</span></span>
      <span class="file-size">${formatFileSize(item.size || 0)}</span>
    </label>`;
  }).join('');
  return `${folderHtml}${fileHtml}`;
}

function buildFolderSelectionStats() {
  const stats = new Map();
  discoveredFiles.forEach((item) => {
    const parts = String(item.relative_path || '').split('/').filter(Boolean);
    for (let depth = 1; depth < parts.length; depth += 1) {
      const folderPath = parts.slice(0, depth).join('/');
      const current = stats.get(folderPath) || { total: 0, selected: 0 };
      current.total += 1;
      if (item.selected) current.selected += 1;
      stats.set(folderPath, current);
    }
  });
  return stats;
}

function allFolderPaths() {
  const folders = new Set();
  discoveredFiles.forEach((item) => {
    const parts = String(item.relative_path || '').split('/').filter(Boolean);
    for (let depth = 1; depth < parts.length; depth += 1) {
      folders.add(parts.slice(0, depth).join('/'));
    }
  });
  return folders;
}

function setAllFoldersExpanded(expanded) {
  expandedFolders.clear();
  if (expanded) allFolderPaths().forEach((path) => expandedFolders.add(path));
  renderFiles();
}

function isFileInsideFolder(relativePath, folderPath) {
  return String(relativePath || '').startsWith(`${folderPath}/`);
}

function naturalCompare(left, right) {
  return String(left || '').localeCompare(String(right || ''), 'zh-CN', { numeric: true, sensitivity: 'base' });
}

function visibleFiles() {
  const keyword = document.getElementById('fileSearch').value.trim().toLocaleLowerCase('zh-CN');
  if (!keyword) return discoveredFiles;
  return discoveredFiles.filter((item) => String(item.relative_path || '').toLocaleLowerCase('zh-CN').includes(keyword));
}

function setVisibleSelection(selected) {
  visibleFiles().forEach((item) => { item.selected = selected; });
  invalidatePreview();
  renderFiles();
}

function updateSelectionSummary() {
  const selected = discoveredFiles.filter((item) => item.selected);
  const totalSize = selected.reduce((sum, item) => sum + Number(item.size || 0), 0);
  document.getElementById('selectedCount').textContent = selected.length.toLocaleString();
  document.getElementById('selectedSize').textContent = formatFileSize(totalSize);
  const hasFiles = discoveredFiles.length > 0;
  document.getElementById('selectAllBtn').disabled = !hasFiles;
  document.getElementById('clearSelectionBtn').disabled = !hasFiles;
  const hasFolders = allFolderPaths().size > 0;
  const searchActive = Boolean(document.getElementById('fileSearch').value.trim());
  document.getElementById('expandAllBtn').disabled = !hasFolders || searchActive;
  document.getElementById('collapseAllBtn').disabled = !hasFolders || searchActive;
  updateActions();
}

function setMode(mode) {
  currentMode = ['cleanup', 'numbering', 'regex'].includes(mode) ? mode : 'cleanup';
  document.querySelectorAll('.mode-button').forEach((button) => {
    button.classList.toggle('active', button.dataset.mode === currentMode);
  });
  document.getElementById('cleanupFields').classList.toggle('is-hidden', currentMode !== 'cleanup');
  document.getElementById('numberingFields').classList.toggle('is-hidden', currentMode !== 'numbering');
  document.getElementById('regexFields').classList.toggle('is-hidden', currentMode !== 'regex');
  invalidatePreview();
}

function updateCleanupControlAvailability() {
  const groups = [
    {
      master: 'cleanupLeadingNumberInput',
      controls: ['cleanupMaxDigitsInput', 'cleanupSeparatorSpaceInput', 'cleanupSeparatorUnderscoreInput'],
    },
    {
      master: 'cleanupDatetimeInput',
      controls: ['cleanupDatetimeCompactInput', 'cleanupDatetimeDottedInput'],
    },
    {
      master: 'cleanupTranslatedInput',
      controls: ['cleanupSuffixInput'],
    },
  ];
  groups.forEach((group) => {
    const master = document.getElementById(group.master);
    const enabled = master.checked;
    group.controls.forEach((id) => { document.getElementById(id).disabled = !enabled; });
    master.closest('.cleanup-rule')?.classList.toggle('is-disabled', !enabled);
  });
}

function buildRequestBody() {
  return {
    directory_path: document.getElementById('directoryPath').value.trim(),
    relative_paths: discoveredFiles.filter((item) => item.selected).map((item) => item.relative_path),
    mode: currentMode,
    recursive: document.getElementById('recursiveInput').checked,
    include_hidden: document.getElementById('hiddenInput').checked,
    regex_pattern: currentMode === 'regex' ? document.getElementById('patternInput').value : '',
    replacement: currentMode === 'regex' ? document.getElementById('replacementInput').value : '',
    ignore_case: currentMode === 'regex' && document.getElementById('ignoreCaseInput').checked,
    cleanup_remove_leading_number: document.getElementById('cleanupLeadingNumberInput').checked,
    cleanup_leading_number_max_digits: Math.max(1, Math.min(12, Number(document.getElementById('cleanupMaxDigitsInput').value) || 6)),
    cleanup_leading_number_space: document.getElementById('cleanupSeparatorSpaceInput').checked,
    cleanup_leading_number_underscore: document.getElementById('cleanupSeparatorUnderscoreInput').checked,
    cleanup_remove_datetime: document.getElementById('cleanupDatetimeInput').checked,
    cleanup_datetime_compact: document.getElementById('cleanupDatetimeCompactInput').checked,
    cleanup_datetime_dotted: document.getElementById('cleanupDatetimeDottedInput').checked,
    cleanup_remove_translated: document.getElementById('cleanupTranslatedInput').checked,
    cleanup_translated_suffix: document.getElementById('cleanupSuffixInput').value,
  };
}

async function generatePreview() {
  const body = buildRequestBody();
  if (!body.relative_paths.length) {
    showMessage('previewMessage', '请至少选择一个文件。', 'error');
    return;
  }
  const button = document.getElementById('previewBtn');
  button.disabled = true;
  showMessage('previewMessage', '正在生成改名预览...', '');
  try {
    const response = await fetch('/task/file-rename/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '预览失败');
    previewData = payload;
    previewFingerprint = JSON.stringify(body);
    renderPreview(payload);
    const problemCount = Number(payload.conflict_count || 0) + Number(payload.invalid_count || 0);
    if (problemCount) {
      showMessage('previewMessage', `预览完成，但有 ${problemCount} 个冲突或无效名称，解决后才能生成副本。`, 'error');
    } else if (!payload.process_count) {
      showMessage('previewMessage', '当前规则没有产生需要复制的改名文件。', 'warning');
    } else {
      showMessage('previewMessage', `预览完成：将复制 ${payload.process_count} 个文件，源文件不会修改。`, 'success');
    }
  } catch (error) {
    invalidatePreview(false);
    showMessage('previewMessage', error.message || '预览失败，请检查规则。', 'error');
  } finally {
    updateActions();
  }
}

function renderPreview(payload) {
  document.getElementById('selectedMetric').textContent = Number(payload.selected_count || 0).toLocaleString();
  document.getElementById('processMetric').textContent = Number(payload.process_count || 0).toLocaleString();
  document.getElementById('skipMetric').textContent = Number(payload.skipped_count || 0).toLocaleString();
  document.getElementById('conflictMetric').textContent = (Number(payload.conflict_count || 0) + Number(payload.invalid_count || 0)).toLocaleString();
  document.getElementById('sizeMetric').textContent = formatFileSize(payload.estimated_bytes || 0);
  document.getElementById('previewMode').textContent = payload.mode === 'cleanup'
    ? '常用清理'
    : payload.mode === 'regex'
      ? '高级正则'
      : `自动编号 · ${payload.number_width || 0} 位`;
  const operations = payload.operations || [];
  document.getElementById('previewRows').innerHTML = operations.length
    ? operations.map((item) => `<tr>
        <td><span class="status ${escAttr(item.status)}">${escHtml(previewStatusMap[item.status] || item.status)}</span></td>
        <td>${escHtml(item.source_relative_path)}</td>
        <td>${escHtml(item.target_relative_path || '-')}</td>
        <td>${escHtml(item.reason || '-')}</td>
      </tr>`).join('')
    : '<tr><td colspan="4"><div class="empty">没有预览项目。</div></td></tr>';
}

async function executeCopy() {
  const body = buildRequestBody();
  if (!previewData || previewFingerprint !== JSON.stringify(body)) {
    showMessage('previewMessage', '文件选择或规则已变化，请重新生成预览。', 'warning');
    invalidatePreview();
    return;
  }
  const count = Number(previewData.process_count || 0);
  const size = formatFileSize(previewData.estimated_bytes || 0);
  if (!window.confirm(`将在源目录内新建副本目录，复制并改名 ${count} 个文件（约 ${size}）。\n源文件不会修改，是否继续？`)) return;

  stopPolling();
  document.getElementById('executeBtn').disabled = true;
  setStatus({ status: 'queued', progress: 0, message: '正在提交副本任务...' });
  hideMessage('resultMessage');
  try {
    const response = await fetch('/task/file-rename', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || '任务提交失败');
    currentTaskId = payload.task_id;
    document.getElementById('taskIdText').textContent = currentTaskId ? `任务 ${currentTaskId.slice(0, 8)}` : '';
    startPolling();
    await loadStatus();
  } catch (error) {
    setStatus({ status: 'failed', progress: 0, message: error.message || '任务提交失败' });
    showMessage('resultMessage', error.message || '任务提交失败。', 'error');
    updateActions();
  }
}

function startPolling() {
  stopPolling();
  pollTimer = window.setInterval(loadStatus, POLL_INTERVAL_MS);
}

function stopPolling() {
  if (pollTimer) window.clearInterval(pollTimer);
  pollTimer = null;
}

async function loadStatus() {
  if (!currentTaskId) return;
  try {
    const response = await fetch(`/task/file-rename/status/${encodeURIComponent(currentTaskId)}`);
    const task = await response.json();
    if (!response.ok) throw new Error(task.detail || '读取任务状态失败');
    setStatus(task);
    if (task.status === 'done') {
      stopPolling();
      const result = task.result || {};
      const failureText = result.failed_count ? `，失败 ${result.failed_count} 个` : '';
      showMessage(
        'resultMessage',
        `副本已生成：${result.output_directory || '-'}。成功 ${result.copied_count || 0} 个${failureText}。`,
        result.failed_count ? 'warning' : 'success',
      );
      updateActions();
    } else if (task.status === 'failed' || task.status === 'cancelled') {
      stopPolling();
      showMessage('resultMessage', task.error || task.message || statusTextMap[task.status], 'error');
      updateActions();
    }
  } catch (error) {
    console.error('loadStatus', error);
  }
}

function setStatus(task) {
  const progress = Math.max(0, Math.min(100, Number(task.progress || 0)));
  const status = statusTextMap[task.status] || task.status || '等待提交';
  const detail = task.message || task.error || '';
  document.getElementById('progressFill').style.width = `${progress}%`;
  document.getElementById('progressText').textContent = `${progress}%`;
  document.getElementById('statusText').textContent = detail ? `${status}：${detail}` : status;
}

function invalidateScan() {
  discoveredFiles = [];
  expandedFolders.clear();
  scanCompleted = false;
  invalidatePreview();
  renderFiles();
  hideMessage('scanMessage');
}

function invalidatePreview(clearMessage = true) {
  previewData = null;
  previewFingerprint = '';
  renderPreview({});
  if (clearMessage) hideMessage('previewMessage');
  updateActions();
}

function updateActions() {
  const selectedCount = discoveredFiles.filter((item) => item.selected).length;
  const taskActive = Boolean(currentTaskId && pollTimer);
  document.getElementById('exportListBtn').disabled = !scanCompleted || !discoveredFiles.length;
  document.getElementById('previewBtn').disabled = !scanCompleted || !selectedCount || taskActive;
  const previewCurrent = previewData && previewFingerprint === JSON.stringify(buildRequestBody());
  const hasProblems = previewData && (Number(previewData.conflict_count || 0) + Number(previewData.invalid_count || 0) > 0);
  document.getElementById('executeBtn').disabled = !previewCurrent || !previewData.process_count || hasProblems || taskActive;
}

function exportFileList() {
  if (!scanCompleted || !discoveredFiles.length) {
    showMessage('scanMessage', '请先扫描目录，再导出文件名单。', 'warning');
    return;
  }
  const headers = ['序号', '文件名', '相对路径', '扩展名', '文件大小（字节）', '修改时间'];
  const rows = discoveredFiles.map((item, index) => [
    index + 1,
    item.name || '',
    item.relative_path || '',
    item.extension || '',
    Number(item.size || 0),
    item.modified_at || '',
  ]);
  const csv = [headers, ...rows]
    .map((row) => row.map(csvCell).join(','))
    .join('\r\n');
  const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `文件名单_${buildTimestamp()}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  showMessage('scanMessage', `已导出 ${discoveredFiles.length} 个文件的名单。`, 'success');
}

function csvCell(value) {
  let text = String(value ?? '');
  if (/^[=+\-@]/.test(text.trimStart())) text = `'${text}`;
  return `"${text.replace(/"/g, '""')}"`;
}

function buildTimestamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

function resetPage() {
  stopPolling();
  currentTaskId = '';
  discoveredFiles = [];
  expandedFolders.clear();
  scanCompleted = false;
  document.getElementById('directoryPath').value = '';
  document.getElementById('recursiveInput').checked = true;
  document.getElementById('hiddenInput').checked = false;
  document.getElementById('fileSearch').value = '';
  document.getElementById('ignoreCaseInput').checked = false;
  document.getElementById('cleanupLeadingNumberInput').checked = true;
  document.getElementById('cleanupMaxDigitsInput').value = '6';
  document.getElementById('cleanupSeparatorSpaceInput').checked = true;
  document.getElementById('cleanupSeparatorUnderscoreInput').checked = true;
  document.getElementById('cleanupDatetimeInput').checked = true;
  document.getElementById('cleanupDatetimeCompactInput').checked = true;
  document.getElementById('cleanupDatetimeDottedInput').checked = true;
  document.getElementById('cleanupTranslatedInput').checked = true;
  document.getElementById('cleanupSuffixInput').value = '_translated';
  updateCleanupControlAvailability();
  setMode('cleanup');
  hideMessage('scanMessage');
  hideMessage('resultMessage');
  document.getElementById('taskIdText').textContent = '';
  setStatus({ status: '', progress: 0, message: '等待提交' });
  invalidatePreview();
  renderFiles();
}

function showMessage(id, text, kind) {
  const element = document.getElementById(id);
  element.textContent = text;
  element.className = `notice${kind ? ` ${kind}` : ''}`;
}

function hideMessage(id) {
  const element = document.getElementById(id);
  element.textContent = '';
  element.className = 'notice is-hidden';
}

function formatFileSize(bytes) {
  const value = Number(bytes || 0);
  if (value < 1024) return `${value} B`;
  if (value < 1024 ** 2) return `${(value / 1024).toFixed(1)} KB`;
  if (value < 1024 ** 3) return `${(value / 1024 ** 2).toFixed(1)} MB`;
  return `${(value / 1024 ** 3).toFixed(2)} GB`;
}

function escHtml(value) {
  const div = document.createElement('div');
  div.textContent = value ?? '';
  return div.innerHTML;
}

function escAttr(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

init();
