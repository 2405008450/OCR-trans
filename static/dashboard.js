const POLL_INTERVAL = 2000;
const DETAIL_POLL_INTERVAL = 1500;
const DASHBOARD_TIME_ZONE = 'Asia/Shanghai';

let currentPage = 1;
let currentPageSize = 20;
let currentStatusFilter = '';
let currentTypeFilter = '';
let currentKeyword = '';
let listTimer = null;
let detailTimer = null;
let openTaskId = null;

const STATUS_BADGE = {
  queued: { cls: 'badge-queued', icon: 'fa-clock', text: '排队中' },
  running: { cls: 'badge-running', icon: 'fa-spinner fa-spin', text: '处理中' },
  done: { cls: 'badge-done', icon: 'fa-check', text: '已完成' },
  failed: { cls: 'badge-failed', icon: 'fa-xmark', text: '失败' },
  cancelled: { cls: 'badge-cancelled', icon: 'fa-ban', text: '已取消' },
};

function init() {
  document.getElementById('btnRefresh').addEventListener('click', () => loadAll());
  document.getElementById('filterType').addEventListener('change', (event) => {
    currentTypeFilter = event.target.value;
    currentPage = 1;
    loadList();
  });

  let searchDebounce = null;
  document.getElementById('searchInput').addEventListener('input', (event) => {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
      currentKeyword = event.target.value.trim();
      currentPage = 1;
      loadList();
    }, 400);
  });

  document.querySelectorAll('.stat-pill').forEach((pill) => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.stat-pill').forEach((item) => item.classList.remove('active'));
      pill.classList.add('active');
      currentStatusFilter = pill.dataset.filter;
      currentPage = 1;
      loadList();
    });
  });

  document.getElementById('drawerOverlay').addEventListener('click', closeDrawer);
  document.getElementById('drawerClose').addEventListener('click', closeDrawer);

  loadAll();
  startListPolling();
}

function startListPolling() {
  stopListPolling();
  listTimer = setInterval(() => loadAll(true), POLL_INTERVAL);
}

function stopListPolling() {
  if (listTimer) {
    clearInterval(listTimer);
    listTimer = null;
  }
}

async function loadAll(silent) {
  await Promise.all([loadStats(silent), loadList(silent)]);
}

async function loadStats(silent) {
  try {
    const response = await fetch('/task/dashboard/stats');
    if (!response.ok) return;
    const data = await response.json();
    document.getElementById('statTotal').textContent = data.total ?? 0;
    document.getElementById('statQueued').textContent = data.queued ?? 0;
    document.getElementById('statRunning').textContent = data.running ?? 0;
    document.getElementById('statDone').textContent = data.done ?? 0;
    document.getElementById('statFailed').textContent = data.failed ?? 0;
    document.getElementById('statCancelled').textContent = data.cancelled ?? 0;
  } catch (error) {
    if (!silent) console.error('loadStats', error);
  }
}

async function loadList(silent) {
  try {
    const params = new URLSearchParams({ page: currentPage, page_size: currentPageSize });
    if (currentStatusFilter) params.set('status', currentStatusFilter);
    if (currentTypeFilter) params.set('task_type', currentTypeFilter);
    if (currentKeyword) params.set('keyword', currentKeyword);

    const response = await fetch(`/task/list?${params.toString()}`);
    if (!response.ok) return;
    const data = await response.json();
    renderTable(data.items || []);
    renderPagination(data.total || 0, data.page || 1, data.page_size || 20);
  } catch (error) {
    if (!silent) console.error('loadList', error);
  }
}

function renderTable(items) {
  const body = document.getElementById('taskBody');
  const empty = document.getElementById('emptyState');
  if (!items.length) {
    body.innerHTML = '';
    empty.style.display = '';
    return;
  }
  empty.style.display = 'none';

  body.innerHTML = items.map((task) => {
    const badge = STATUS_BADGE[task.status] || STATUS_BADGE.queued;
    const progress = task.progress ?? 0;
    const createdAt = task.created_at ? formatTime(task.created_at) : '-';
    const fileName = escHtml(task.filename || '-');
    const shortName = fileName.length > 40 ? `${fileName.slice(0, 37)}...` : fileName;
    return `<tr data-id="${task.task_id}" onclick="openDetail('${task.task_id}')">
      <td>${escHtml(task.display_no || '-')}</td>
      <td>${escHtml(task.task_label || task.task_type)}</td>
      <td class="cell-filename" title="${fileName}">${shortName}</td>
      <td><span class="badge ${badge.cls}"><i class="fas ${badge.icon}"></i> ${badge.text}</span></td>
      <td><div class="mini-progress"><div class="mini-progress-fill" style="width:${progress}%"></div></div> <span style="font-size:12px;color:var(--muted)">${progress}%</span></td>
      <td class="cell-time">${createdAt}</td>
      <td>
        <button class="btn-detail" onclick="event.stopPropagation();openDetail('${task.task_id}')"><i class="fas fa-eye"></i> 详情</button>
        ${(task.status === 'queued' || task.status === 'running') && !task.cancel_requested ? `<button class="btn-cancel-table" onclick="event.stopPropagation();cancelTask('${task.task_id}')" title="取消任务"><i class="fas fa-ban"></i></button>` : ''}
        ${task.cancel_requested && task.status === 'running' ? `<span style="font-size:11px;color:#94a3b8;margin-left:4px">取消中...</span>` : ''}
      </td>
    </tr>`;
  }).join('');
}

function renderPagination(total, page, pageSize) {
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const container = document.getElementById('pagination');
  if (totalPages <= 1) {
    container.innerHTML = '';
    return;
  }

  let html = `<button ${page <= 1 ? 'disabled' : ''} onclick="goPage(${page - 1})"><i class="fas fa-chevron-left"></i></button>`;
  const start = Math.max(1, page - 2);
  const end = Math.min(totalPages, page + 2);
  if (start > 1) html += '<button onclick="goPage(1)">1</button>';
  if (start > 2) html += '<span>...</span>';
  for (let i = start; i <= end; i += 1) {
    html += `<button class="${i === page ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
  }
  if (end < totalPages - 1) html += '<span>...</span>';
  if (end < totalPages) html += `<button onclick="goPage(${totalPages})">${totalPages}</button>`;
  html += `<button ${page >= totalPages ? 'disabled' : ''} onclick="goPage(${page + 1})"><i class="fas fa-chevron-right"></i></button>`;
  container.innerHTML = html;
}

function goPage(page) {
  currentPage = page;
  loadList();
}

function openDetail(taskId) {
  openTaskId = taskId;
  document.getElementById('drawerOverlay').classList.add('open');
  document.getElementById('drawer').classList.add('open');
  loadDetail(taskId);
  startDetailPolling(taskId);
}

function closeDrawer() {
  openTaskId = null;
  stopDetailPolling();
  document.getElementById('drawerOverlay').classList.remove('open');
  document.getElementById('drawer').classList.remove('open');
}

function startDetailPolling(taskId) {
  stopDetailPolling();
  detailTimer = setInterval(() => {
    if (openTaskId !== taskId) {
      stopDetailPolling();
      return;
    }
    loadDetail(taskId, true);
  }, DETAIL_POLL_INTERVAL);
}

function stopDetailPolling() {
  if (detailTimer) {
    clearInterval(detailTimer);
    detailTimer = null;
  }
}

async function loadDetail(taskId, silent) {
  try {
    const response = await fetch(`/task/${taskId}/detail`);
    if (!response.ok) return;
    const data = await response.json();
    renderDetail(data);
    if (['done', 'failed', 'cancelled'].includes(data.status)) {
      stopDetailPolling();
    }
  } catch (error) {
    if (!silent) console.error('loadDetail', error);
  }
}

function renderDetail(task) {
  const badge = STATUS_BADGE[task.status] || STATUS_BADGE.queued;
  const progress = task.progress ?? 0;
  const createdAt = task.created_at ? formatTime(task.created_at) : '-';
  const startedAt = task.started_at ? formatTime(task.started_at) : '-';
  const finishedAt = task.finished_at ? formatTime(task.finished_at) : '-';
  const duration = computeDuration(task.started_at, task.finished_at);
  const eta = estimateFinishTime(task);

  const inputItems = normalizeInputFiles(task.input_files);
  const outputItems = Array.isArray(task.output_files) ? task.output_files : [];

  const inputHtml = inputItems.length ? `<div class="detail-section"><h3><i class="fas fa-file-import"></i> 输入文件</h3><div class="file-list">${inputItems.map((item) => {
    const name = item.name || (item.path || '').split('/').pop() || '输入文件';
    return `<div class="file-item"><span class="fi-name" title="${escHtml(item.path || '')}"><i class="fas fa-file-import" style="color:var(--amber);margin-right:6px"></i>${escHtml(name)}</span><button class="fi-dl" onclick="downloadFile('${task.task_id}','${escAttr(item.path || '')}','${escAttr(name)}')"><i class="fas fa-download"></i></button></div>`;
  }).join('')}</div></div>` : '';

  const outputHtml = outputItems.length ? `<div class="detail-section"><h3><i class="fas fa-file-export"></i> 输出文件</h3><div class="file-list">${outputItems.map((item) => {
    const displayName = item.name || (item.path || '').split('/').pop() || '输出文件';
    return `<div class="file-item"><span class="fi-name" title="${escHtml(item.path || '')}"><i class="fas fa-file-export" style="color:var(--green);margin-right:6px"></i>${escHtml(displayName)}</span><button class="fi-dl" onclick="downloadFile('${task.task_id}','${escAttr(item.path || '')}','${escAttr(displayName)}')"><i class="fas fa-download"></i></button></div>`;
  }).join('')}</div></div>` : '';

  const errorHtml = task.status === 'failed' && task.error ? `<div class="detail-section"><h3><i class="fas fa-exclamation-triangle"></i> 错误信息</h3><div class="detail-error">${escHtml(task.error)}</div></div>` : '';
  const logHtml = task.stream_log ? `<div class="detail-section"><h3><i class="fas fa-terminal"></i> 运行日志</h3><pre class="detail-log">${escHtml(task.stream_log)}</pre></div>` : '';

  document.getElementById('drawerContent').innerHTML = `
    <h2>${escHtml(task.display_no || '-')}</h2>
    <div class="drawer-subtitle">${escHtml(task.task_label || task.task_type || '-')} &middot; <span class="badge ${badge.cls}" style="font-size:12px"><i class="fas ${badge.icon}"></i> ${badge.text}</span></div>

    <div class="detail-section">
      <h3><i class="fas fa-chart-simple"></i> 进度</h3>
      <div style="display:flex;align-items:center;gap:12px">
        <div class="detail-progress-bar" style="flex:1"><div class="detail-progress-fill" style="width:${progress}%"></div></div>
        <span style="font-weight:700;min-width:40px">${progress}%</span>
      </div>
      <div style="color:var(--muted);font-size:13px;margin-top:6px">${escHtml(task.message || '')}</div>
      ${(task.status === 'queued' || task.status === 'running') && !task.cancel_requested ? `<button class="btn-cancel" style="margin-top:10px" onclick="cancelTask('${task.task_id}')"><i class="fas fa-ban"></i> 取消任务</button>` : ''}
      ${task.cancel_requested && task.status === 'running' ? `<div style="margin-top:10px;font-size:13px;color:#94a3b8"><i class="fas fa-spinner fa-spin"></i> 正在取消，等待当前步骤结束...</div>` : ''}
    </div>

    <div class="detail-section">
      <h3><i class="fas fa-info-circle"></i> 基本信息</h3>
      <div class="detail-grid">
        <div class="detail-item"><div class="dl">文件名</div><div class="dv">${escHtml(task.filename || '-')}</div></div>
        <div class="detail-item"><div class="dl">耗时</div><div class="dv">${duration}</div></div>
        <div class="detail-item"><div class="dl">提交时间</div><div class="dv">${createdAt}</div></div>
        <div class="detail-item"><div class="dl">开始时间</div><div class="dv">${startedAt}</div></div>
        <div class="detail-item"><div class="dl">预计完成时间</div><div class="dv">${eta}</div></div>
        <div class="detail-item"><div class="dl">完成时间</div><div class="dv">${finishedAt}</div></div>
      </div>
    </div>

    ${inputHtml}
    ${outputHtml}
    ${errorHtml}
    ${logHtml}
  `;
}

function normalizeInputFiles(inputFiles) {
  if (!inputFiles || typeof inputFiles !== 'object') return [];
  if (Array.isArray(inputFiles.files)) {
    return inputFiles.files
      .filter((item) => item && item.path)
      .map((item) => ({ path: item.path, name: item.original_filename || item.path.split('/').pop() }));
  }

  const results = [];
  const pathKeys = Object.keys(inputFiles).filter(
    (key) => (key.endsWith('_path') || key === 'input_path') && typeof inputFiles[key] === 'string' && inputFiles[key]
  );

  for (const key of pathKeys) {
    const filePath = inputFiles[key];
    const prefix = key.replace(/_path$/, '');
    const friendlyName =
      inputFiles[`${prefix}_filename`] ||
      inputFiles['original_filename'] && key === 'input_path' && inputFiles['original_filename'] ||
      filePath.split('/').pop();
    results.push({ path: filePath, name: friendlyName });
  }

  return results;
}

function downloadFile(taskId, filePath, friendlyName) {
  let url = `/task/${taskId}/download?file_path=${encodeURIComponent(filePath)}`;
  if (friendlyName) {
    url += `&download_name=${encodeURIComponent(friendlyName)}`;
  }
  const link = document.createElement('a');
  link.href = url;
  link.download = friendlyName || '';
  link.click();
}

async function cancelTask(taskId) {
  if (!confirm('确定要取消该任务吗？运行中的任务会在当前步骤结束后停止。')) {
    return;
  }
  try {
    const response = await fetch(`/task/${taskId}/cancel`, { method: 'POST' });
    const data = await response.json();
    if (!response.ok) {
      alert(data.detail || '取消失败');
      return;
    }
    loadAll();
    if (openTaskId === taskId) {
      loadDetail(taskId);
    }
  } catch (error) {
    console.error('cancelTask', error);
    alert('取消请求失败，请重试');
  }
}

function escHtml(value) {
  const div = document.createElement('div');
  div.textContent = value ?? '';
  return div.innerHTML;
}

function escAttr(value) {
  return String(value || '').replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

function formatTime(iso) {
  if (!iso) return '-';
  try {
    const date = parseServerTime(iso);
    if (Number.isNaN(date.getTime())) return iso;
    return formatDateInEast8(date, true);
  } catch {
    return iso;
  }
}

function computeDuration(start, end) {
  if (!start) return '-';
  const startTime = parseServerTime(start);
  const endTime = end ? parseServerTime(end) : new Date();
  if (Number.isNaN(startTime.getTime()) || Number.isNaN(endTime.getTime())) return '-';

  const diff = Math.max(0, Math.floor((endTime - startTime) / 1000));
  if (diff < 60) return `${diff}秒`;
  if (diff < 3600) return `${Math.floor(diff / 60)}分${diff % 60}秒`;
  return `${Math.floor(diff / 3600)}小时${Math.floor((diff % 3600) / 60)}分`;
}

function estimateFinishTime(task) {
  if (task.status === 'done' && task.finished_at) {
    return formatTimeToMinute(task.finished_at);
  }
  if (task.status !== 'running') {
    return '-';
  }

  const progress = Number(task.progress ?? 0);
  if (!Number.isFinite(progress) || progress <= 0 || progress >= 100 || !task.created_at) {
    return '-';
  }

  const createdAt = parseServerTime(task.created_at);
  if (Number.isNaN(createdAt.getTime())) {
    return '-';
  }

  const elapsedMs = Date.now() - createdAt.getTime();
  if (elapsedMs <= 0) {
    return '-';
  }

  const estimatedTotalMs = elapsedMs / (progress / 100);
  const estimatedFinishedAt = new Date(createdAt.getTime() + estimatedTotalMs);
  return formatDateInEast8(estimatedFinishedAt, false);
}

function formatTimeToMinute(iso) {
  if (!iso) return '-';
  try {
    const date = parseServerTime(iso);
    if (Number.isNaN(date.getTime())) return iso;
    return formatDateInEast8(date, false);
  } catch {
    return iso;
  }
}

function parseServerTime(iso) {
  if (!iso) return new Date(NaN);
  const normalized = /([zZ]|[+\-]\d{2}:\d{2})$/.test(iso) ? iso : `${iso}Z`;
  return new Date(normalized);
}

function formatDateInEast8(date, withSeconds) {
  const parts = new Intl.DateTimeFormat('zh-CN', {
    timeZone: DASHBOARD_TIME_ZONE,
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: withSeconds ? '2-digit' : undefined,
    hour12: false,
  }).formatToParts(date);

  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return withSeconds
    ? `${values.month}-${values.day} ${values.hour}:${values.minute}:${values.second}`
    : `${values.month}-${values.day} ${values.hour}:${values.minute}`;
}

init();

