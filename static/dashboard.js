const POLL_INTERVAL = 2000;
const DETAIL_POLL_INTERVAL = 1500;
const DASHBOARD_TIME_ZONE = 'Asia/Shanghai';
const INFERRED_BATCH_TYPES = new Set(['pdf2docx', 'doc_translate', 'msg_convert']);
const INFERRED_BATCH_WINDOW_MS = 2000;

let currentPage = 1;
let currentPageSize = 20;
let currentStatusFilter = '';
let currentTypeFilter = '';
let currentKeyword = '';
let currentFeedbackFilter = '';
let listTimer = null;
let detailTimer = null;
let openTaskId = null;
let feedbackTaskId = null;
let feedbackSubmitting = false;
const SENSITIVE_LOG_PATTERNS = [
  /\bopenrouter\b/i,
  /\bgoogle\/gemini-[\w.-]+\b/i,
  /\bGoogle gemini-3-flash-preview\b/i,
  /\bGoogle Gemini 2\.5 Flash\b/i,
  /\bGoogle Gemini 2\.5 Pro\b/i,
  /\[alignment-llm\].*route=/i,
  /\[alignment-llm\].*model=/i,
  /Gemini\s*路线/i,
  /^.*模型:.*gemini.*$/i,
];

const STATUS_BADGE = {
  queued: { cls: 'badge-queued', icon: 'fa-clock', text: '排队中' },
  running: { cls: 'badge-running', icon: 'fa-spinner fa-spin', text: '处理中' },
  done: { cls: 'badge-done', icon: 'fa-check', text: '已完成' },
  failed: { cls: 'badge-failed', icon: 'fa-xmark', text: '失败' },
  cancelled: { cls: 'badge-cancelled', icon: 'fa-ban', text: '已取消' },
};

const FEEDBACK_CATEGORY_TEXT = {
  processing_exception: '处理结果异常',
  format_issue: '格式或版式异常',
  accuracy_issue: '识别或翻译准确性异常',
  download_issue: '下载或文件异常',
  performance_issue: '处理过慢或卡住',
  exception: '处理结果异常',
  other: '其他问题',
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
      currentStatusFilter = pill.dataset.filter || '';
      currentFeedbackFilter = pill.dataset.feedback || '';
      currentPage = 1;
      loadList();
    });
  });

  document.getElementById('drawerOverlay').addEventListener('click', closeDrawer);
  document.getElementById('drawerClose').addEventListener('click', closeDrawer);
  document.getElementById('feedbackClose').addEventListener('click', closeFeedbackModal);
  document.getElementById('feedbackCancel').addEventListener('click', closeFeedbackModal);
  document.getElementById('feedbackModal').addEventListener('click', (event) => {
    if (event.target.id === 'feedbackModal') closeFeedbackModal();
  });
  document.getElementById('feedbackSubmit').addEventListener('click', submitFeedbackMark);

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
    document.getElementById('statFeedback').textContent = data.feedback_marked ?? 0;
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
    if (currentFeedbackFilter) params.set('feedback_marked', currentFeedbackFilter);

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
    const feedback = normalizeFeedback(task.feedback);
    const taskId = escJsString(task.task_id || '');
    const feedbackLabel = feedback.marked ? feedbackCategoryText(feedback.category) : '未标记';
    const feedbackCell = feedback.marked
      ? `<span class="badge badge-feedback" title="${escAttr(feedback.note || feedbackLabel)}"><i class="fas fa-flag"></i> 已标记</span>`
      : `<span class="badge badge-feedback-muted"><i class="far fa-flag"></i> 未标记</span>`;
    return `<tr data-id="${escAttr(task.task_id || '')}">
      <td>${escHtml(task.display_no || '-')}</td>
      <td>${escHtml(task.task_label || task.task_type)}</td>
      <td class="cell-filename" title="${fileName}">${shortName}</td>
      <td class="cell-ip">${escHtml(task.client_ip || '-')}</td>
      <td><span class="badge ${badge.cls}"><i class="fas ${badge.icon}"></i> ${badge.text}</span></td>
      <td>${feedbackCell}</td>
      <td><div class="mini-progress"><div class="mini-progress-fill" style="width:${progress}%"></div></div> <span style="font-size:12px;color:var(--muted)">${progress}%</span></td>
      <td class="cell-time">${createdAt}</td>
      <td><div class="action-cell">
        <button class="btn-detail" onclick="event.stopPropagation();openDetail('${taskId}')"><i class="fas fa-eye"></i> 详情</button>
        <button class="btn-feedback-table ${feedback.marked ? 'is-marked' : ''}" onclick="event.stopPropagation();openFeedbackModal('${taskId}')" title="${feedback.marked ? '更新异常标记' : '标记异常'}"><i class="fas fa-flag"></i></button>
        ${(task.status === 'queued' || task.status === 'running') && !task.cancel_requested ? `<button class="btn-cancel-table" onclick="event.stopPropagation();cancelTask('${taskId}')" title="取消任务"><i class="fas fa-ban"></i></button>` : ''}
        ${task.cancel_requested && task.status === 'running' ? `<span style="font-size:11px;color:#94a3b8;margin-left:4px">取消中...</span>` : ''}
      </div></td>
    </tr>`;
  }).join('');
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

  body.innerHTML = buildTableGroups(items).map((group) => {
    if (group.type === 'single') {
      return renderTaskRow(group.task);
    }
    return `${renderBatchRow(group)}${group.tasks.map((task) => renderTaskRow(task, true)).join('')}`;
  }).join('');
}

function buildTableGroups(items) {
  const groups = [];
  const batchMap = new Map();
  const inferredMap = new Map();

  items.forEach((task) => {
    const batch = normalizeBatch(task.batch);
    if (!batch.id) {
      const inferredKey = buildInferredBatchKey(task);
      if (!inferredKey) {
        groups.push({ type: 'single', task });
        return;
      }
      if (!inferredMap.has(inferredKey)) {
        const group = {
          type: 'batch',
          batch: buildInferredBatch(task, inferredKey, []),
          tasks: [],
        };
        inferredMap.set(inferredKey, group);
        groups.push(group);
      }
      inferredMap.get(inferredKey).tasks.push(task);
      return;
    }
    if (!batchMap.has(batch.id)) {
      const group = { type: 'batch', batch, tasks: [] };
      batchMap.set(batch.id, group);
      groups.push(group);
    }
    batchMap.get(batch.id).tasks.push(task);
  });
  groups.forEach((group) => {
    if (group.type === 'batch') {
      if (group.batch.inferred) {
        group.tasks.sort(compareTasksForInferredBatch);
        group.batch = buildInferredBatch(group.tasks[0], group.batch.id, group.tasks);
        return;
      }
      group.tasks.sort((left, right) => (normalizeBatch(left.batch).index || 0) - (normalizeBatch(right.batch).index || 0));
    }
  });
  return groups.flatMap((group) => {
    if (group.type !== 'batch' || !group.batch.inferred) {
      return [group];
    }
    return splitInferredBatchGroup(group);
  });
}

function normalizeBatch(batch) {
  if (!batch || typeof batch !== 'object' || !batch.id) {
    return { id: '', name: '', index: null, total: null, inferred: false };
  }
  return {
    id: String(batch.id || ''),
    name: batch.name || '',
    index: Number(batch.index || 0) || null,
    total: Number(batch.total || 0) || null,
    inferred: Boolean(batch.inferred),
  };
}

function buildInferredBatchKey(task) {
  const taskType = String(task.task_type || '').toLowerCase();
  if (!INFERRED_BATCH_TYPES.has(taskType)) {
    return '';
  }
  const createdAt = parseServerTime(task.created_at);
  if (Number.isNaN(createdAt.getTime())) {
    return '';
  }
  return `inferred:${taskType}:${Math.floor(createdAt.getTime() / 60000)}`;
}

function buildInferredBatch(task, inferredKey, tasks) {
  const taskLabel = task?.task_label || task?.task_type || '任务';
  const total = tasks.length || null;
  return {
    id: inferredKey,
    name: `${taskLabel} 批量任务（自动识别）`,
    index: null,
    total,
    inferred: true,
  };
}

function splitInferredBatchGroup(group) {
  const chunks = [];
  let current = [];

  group.tasks.forEach((task) => {
    const previous = current[current.length - 1];
    if (!previous || isInferredBatchNeighbor(previous, task)) {
      current.push(task);
      return;
    }
    chunks.push(current);
    current = [task];
  });
  if (current.length) {
    chunks.push(current);
  }

  return chunks.flatMap((tasks) => {
    if (tasks.length < 2) {
      return tasks.map((task) => ({ type: 'single', task }));
    }
    return [{
      type: 'batch',
      batch: buildInferredBatch(tasks[0], `${group.batch.id}:${displayNoSuffix(tasks[0]) || 'group'}`, tasks),
      tasks,
    }];
  });
}

function isInferredBatchNeighbor(left, right) {
  const leftNo = displayNoSuffix(left);
  const rightNo = displayNoSuffix(right);
  if (leftNo !== null && rightNo !== null && Math.abs(rightNo - leftNo) !== 1) {
    return false;
  }

  const leftCreated = parseServerTime(left.created_at);
  const rightCreated = parseServerTime(right.created_at);
  if (Number.isNaN(leftCreated.getTime()) || Number.isNaN(rightCreated.getTime())) {
    return false;
  }
  return Math.abs(rightCreated.getTime() - leftCreated.getTime()) <= INFERRED_BATCH_WINDOW_MS;
}

function compareTasksForInferredBatch(left, right) {
  const leftNo = displayNoSuffix(left);
  const rightNo = displayNoSuffix(right);
  if (leftNo !== null && rightNo !== null && leftNo !== rightNo) {
    return leftNo - rightNo;
  }
  return parseServerTime(left.created_at).getTime() - parseServerTime(right.created_at).getTime();
}

function displayNoSuffix(task) {
  const match = String(task?.display_no || '').match(/(\d+)$/);
  return match ? Number(match[1]) : null;
}

function renderBatchRow(group) {
  const batch = group.batch;
  const batchId = escJsString(batch.id);
  const batchName = batch.name || `批量任务 ${batch.id.slice(0, 8)}`;
  const visibleTotal = group.tasks.length;
  const total = batch.total || visibleTotal;
  const doneCount = group.tasks.filter((task) => task.status === 'done').length;
  const activeCount = group.tasks.filter((task) => ['queued', 'running'].includes(task.status) && !task.cancel_requested).length;
  const failedCount = group.tasks.filter((task) => task.status === 'failed').length;
  const cancelledCount = group.tasks.filter((task) => task.status === 'cancelled').length;
  const avgProgress = visibleTotal
    ? Math.round(group.tasks.reduce((sum, task) => sum + Number(task.progress || 0), 0) / visibleTotal)
    : 0;
  const statusText = [
    `${doneCount}/${total} 已完成`,
    activeCount ? `${activeCount} 处理中` : '',
    failedCount ? `${failedCount} 失败` : '',
    cancelledCount ? `${cancelledCount} 已暂停` : '',
  ].filter(Boolean).join(' · ');
  const firstCreated = group.tasks[0]?.created_at ? formatTime(group.tasks[0].created_at) : '-';
  const taskIds = escJsString(group.tasks.map((task) => task.task_id).filter(Boolean).join(','));
  const actionButtons = batch.inferred
    ? `<button class="btn-batch" onclick="event.stopPropagation();downloadTaskBatch('${taskIds}')"><i class="fas fa-box-archive"></i> 批量下载</button>
          ${activeCount ? `<button class="btn-batch btn-batch-warn" onclick="event.stopPropagation();cancelTaskBatch('${taskIds}')"><i class="fas fa-pause"></i> 暂停本批</button>` : ''}`
    : `<button class="btn-batch" onclick="event.stopPropagation();downloadBatchGroup('${batchId}')"><i class="fas fa-box-archive"></i> 批量下载</button>
          ${activeCount ? `<button class="btn-batch btn-batch-warn" onclick="event.stopPropagation();cancelBatchGroup('${batchId}')"><i class="fas fa-pause"></i> 暂停本批</button>` : ''}`;
  return `<tr class="batch-row" data-batch-id="${escAttr(batch.id)}">
    <td colspan="9">
      <div class="batch-header">
        <div class="batch-main">
          <div class="batch-title"><i class="fas fa-layer-group"></i> ${escHtml(batchName)}</div>
          <div class="batch-meta">${escHtml(statusText || '批量任务')} · 提交时间 ${escHtml(firstCreated)}</div>
        </div>
        <div class="batch-progress">
          <div class="mini-progress"><div class="mini-progress-fill" style="width:${avgProgress}%"></div></div>
          <span>${avgProgress}%</span>
        </div>
        <div class="batch-actions">
          ${actionButtons}
        </div>
      </div>
    </td>
  </tr>`;
}

function renderTaskRow(task, isBatchChild = false) {
  const badge = STATUS_BADGE[task.status] || STATUS_BADGE.queued;
  const progress = task.progress ?? 0;
  const createdAt = task.created_at ? formatTime(task.created_at) : '-';
  const fileName = escHtml(task.filename || '-');
  const shortName = fileName.length > 40 ? `${fileName.slice(0, 37)}...` : fileName;
  const feedback = normalizeFeedback(task.feedback);
  const taskId = escJsString(task.task_id || '');
  const batch = normalizeBatch(task.batch);
  const displayNo = isBatchChild && batch.index ? `${task.display_no || '-'} · #${batch.index}` : (task.display_no || '-');
  const feedbackLabel = feedback.marked ? feedbackCategoryText(feedback.category) : '未标记';
  const feedbackCell = feedback.marked
    ? `<span class="badge badge-feedback" title="${escAttr(feedback.note || feedbackLabel)}"><i class="fas fa-flag"></i> 已标记</span>`
    : `<span class="badge badge-feedback-muted"><i class="far fa-flag"></i> 未标记</span>`;
  return `<tr class="${isBatchChild ? 'batch-child-row' : ''}" data-id="${escAttr(task.task_id || '')}">
    <td>${escHtml(displayNo)}</td>
    <td>${escHtml(task.task_label || task.task_type)}</td>
    <td class="cell-filename" title="${fileName}">${shortName}</td>
    <td class="cell-ip">${escHtml(task.client_ip || '-')}</td>
    <td><span class="badge ${badge.cls}"><i class="fas ${badge.icon}"></i> ${badge.text}</span></td>
    <td>${feedbackCell}</td>
    <td><div class="mini-progress"><div class="mini-progress-fill" style="width:${progress}%"></div></div> <span style="font-size:12px;color:var(--muted)">${progress}%</span></td>
    <td class="cell-time">${createdAt}</td>
    <td><div class="action-cell">
      <button class="btn-detail" onclick="event.stopPropagation();openDetail('${taskId}')"><i class="fas fa-eye"></i> 详情</button>
      <button class="btn-feedback-table ${feedback.marked ? 'is-marked' : ''}" onclick="event.stopPropagation();openFeedbackModal('${taskId}')" title="${feedback.marked ? '更新异常标记' : '标记异常'}"><i class="fas fa-flag"></i></button>
      ${(task.status === 'queued' || task.status === 'running') && !task.cancel_requested ? `<button class="btn-cancel-table" onclick="event.stopPropagation();cancelTask('${taskId}')" title="取消任务"><i class="fas fa-ban"></i></button>` : ''}
      ${task.cancel_requested && task.status === 'running' ? `<span style="font-size:11px;color:#94a3b8;margin-left:4px">取消中...</span>` : ''}
    </div></td>
  </tr>`;
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
  document.body.classList.add('drawer-open');
  document.getElementById('drawerOverlay').classList.add('open');
  document.getElementById('drawer').classList.add('open');
  loadDetail(taskId);
  startDetailPolling(taskId);
}

function closeDrawer() {
  openTaskId = null;
  stopDetailPolling();
  document.body.classList.remove('drawer-open');
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
  const modelInfo = task.model_info && typeof task.model_info === 'object' ? task.model_info : null;
  const feedback = normalizeFeedback(task.feedback);
  const modelItemHtml = modelInfo
    ? `<div class="detail-item"><div class="dl">\u4f7f\u7528\u6a21\u578b</div><div class="dv">${escHtml(modelInfo.label || '-')}</div></div>`
    : '';

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

  const renderedLogHtml = buildSanitizedLogHtml(task.stream_log || '');
  const feedbackHtml = buildFeedbackDetailHtml(task, feedback);

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

    ${feedbackHtml}

    <div class="detail-section">
      <h3><i class="fas fa-info-circle"></i> 基本信息</h3>
      <div class="detail-grid">
        <div class="detail-item"><div class="dl">文件名</div><div class="dv">${escHtml(task.filename || '-')}</div></div>
        <div class="detail-item"><div class="dl">用户 IP</div><div class="dv">${escHtml(task.client_ip || '-')}</div></div>
        ${modelItemHtml}
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
    ${renderedLogHtml}
  `;
}

function normalizeFeedback(feedback) {
  if (!feedback || typeof feedback !== 'object') {
    return { marked: false, category: '', note: '', marked_at: null };
  }
  return {
    marked: Boolean(feedback.marked),
    category: feedback.category || '',
    note: feedback.note || '',
    marked_at: feedback.marked_at || null,
  };
}

function feedbackCategoryText(category) {
  return FEEDBACK_CATEGORY_TEXT[category] || category || '处理结果异常';
}

function buildFeedbackDetailHtml(task, feedback) {
  const taskId = escJsString(task.task_id || '');
  const markedAt = feedback.marked_at ? formatTime(feedback.marked_at) : '-';
  const stateBadge = feedback.marked
    ? `<span class="badge badge-feedback"><i class="fas fa-flag"></i> 已标记异常</span>`
    : `<span class="badge badge-feedback-muted"><i class="far fa-flag"></i> 未标记异常</span>`;
  const markedInfo = feedback.marked
    ? `<div class="feedback-meta">异常类型：${escHtml(feedbackCategoryText(feedback.category))}<br>标记时间：${escHtml(markedAt)}</div>
       ${feedback.note ? `<div class="feedback-note">${escHtml(feedback.note)}</div>` : ''}`
    : `<div class="feedback-meta">用户发现处理结果、下载文件或运行过程有异常时，可以在这里给开发人员留下反馈。</div>`;

  return `<div class="detail-section">
    <h3><i class="fas fa-flag"></i> 异常标记</h3>
    <div class="feedback-panel ${feedback.marked ? 'is-marked' : ''}">
      <div class="feedback-head">
        ${stateBadge}
        <div class="feedback-actions">
          <button class="btn-feedback" onclick="openFeedbackModal('${taskId}')"><i class="fas fa-flag"></i> ${feedback.marked ? '更新标记' : '标记异常'}</button>
          ${feedback.marked ? `<button class="btn-feedback-secondary" onclick="clearFeedbackMark('${taskId}')"><i class="fas fa-eraser"></i> 取消标记</button>` : ''}
        </div>
      </div>
      ${markedInfo}
    </div>
  </div>`;
}

async function openFeedbackModal(taskId) {
  feedbackTaskId = taskId;
  feedbackSubmitting = false;
  document.getElementById('feedbackCategory').value = 'processing_exception';
  document.getElementById('feedbackNote').value = '';
  document.getElementById('feedbackSubmit').disabled = false;
  document.getElementById('feedbackModal').classList.add('open');

  try {
    const response = await fetch(`/task/${encodeURIComponent(taskId)}/detail`);
    if (response.ok) {
      const task = await response.json();
      const feedback = normalizeFeedback(task.feedback);
      if (feedback.category && FEEDBACK_CATEGORY_TEXT[feedback.category]) {
        document.getElementById('feedbackCategory').value = feedback.category;
      }
      document.getElementById('feedbackNote').value = feedback.note || '';
    }
  } catch (error) {
    console.error('openFeedbackModal', error);
  }
  document.getElementById('feedbackNote').focus();
}

function closeFeedbackModal() {
  if (feedbackSubmitting) return;
  feedbackTaskId = null;
  document.getElementById('feedbackModal').classList.remove('open');
}

async function submitFeedbackMark() {
  if (!feedbackTaskId || feedbackSubmitting) return;

  const submitButton = document.getElementById('feedbackSubmit');
  const category = document.getElementById('feedbackCategory').value;
  const note = document.getElementById('feedbackNote').value.trim();
  feedbackSubmitting = true;
  submitButton.disabled = true;

  try {
    const response = await fetch(`/task/${encodeURIComponent(feedbackTaskId)}/feedback`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ marked: true, category, note }),
    });
    const data = await response.json();
    if (!response.ok) {
      alert(data.detail || '保存标记失败');
      return;
    }
    const savedTaskId = feedbackTaskId;
    feedbackTaskId = null;
    document.getElementById('feedbackModal').classList.remove('open');
    await loadAll();
    if (openTaskId === savedTaskId) {
      await loadDetail(savedTaskId);
    }
  } catch (error) {
    console.error('submitFeedbackMark', error);
    alert('保存标记失败，请重试');
  } finally {
    feedbackSubmitting = false;
    submitButton.disabled = false;
  }
}

async function clearFeedbackMark(taskId) {
  if (!confirm('确定要取消该任务的异常标记吗？')) {
    return;
  }
  try {
    const response = await fetch(`/task/${encodeURIComponent(taskId)}/feedback`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ marked: false }),
    });
    const data = await response.json();
    if (!response.ok) {
      alert(data.detail || '取消标记失败');
      return;
    }
    await loadAll();
    if (openTaskId === taskId) {
      await loadDetail(taskId);
    }
  } catch (error) {
    console.error('clearFeedbackMark', error);
    alert('取消标记失败，请重试');
  }
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

function sanitizeStreamLog(logText) {
  if (!logText) return '';
  return logText
    .split(/\r?\n/)
    .filter((line) => {
      const normalized = line.trim();
      if (!normalized) return true;
      return !SENSITIVE_LOG_PATTERNS.some((pattern) => pattern.test(normalized));
    })
    .join('\n')
    .trim();
}

function buildSanitizedLogHtml(logText) {
  const sanitized = sanitizeStreamLog(logText);
  if (!sanitized) return '';
  return `<div class="detail-section"><h3><i class="fas fa-terminal"></i> 运行日志</h3><pre class="detail-log">${escHtml(sanitized)}</pre></div>`;
}

async function downloadBatchGroup(batchId) {
  try {
    const response = await fetch(`/task/batch/${encodeURIComponent(batchId)}/download`);
    if (!response.ok) {
      let message = '批量下载失败，当前批次可能还没有已完成的输出文件。';
      try {
        const payload = await response.json();
        message = payload.detail || message;
      } catch (_) {}
      alert(message);
      return;
    }
    await saveDownloadResponse(response, buildFallbackArchiveName());
  } catch (error) {
    console.error('downloadBatchGroup', error);
    alert('批量下载失败，请稍后重试。');
  }
}

async function downloadTaskBatch(taskIdsValue) {
  const taskIds = parseTaskIds(taskIdsValue);
  if (!taskIds.length) {
    alert('没有可下载的批量任务。');
    return;
  }
  try {
    const response = await fetch('/task/batch-download', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ task_ids: taskIds }),
    });
    if (!response.ok) {
      let message = '批量下载失败，当前批次可能还没有已完成的输出文件。';
      try {
        const payload = await response.json();
        message = payload.detail || message;
      } catch (_) {}
      alert(message);
      return;
    }
    await saveDownloadResponse(response, buildFallbackArchiveName());
  } catch (error) {
    console.error('downloadTaskBatch', error);
    alert('批量下载失败，请稍后重试。');
  }
}

async function cancelBatchGroup(batchId) {
  if (!confirm('确定要暂停本批还在排队或处理中的任务吗？运行中的任务会在当前步骤结束后停止。')) {
    return;
  }
  try {
    const response = await fetch(`/task/batch/${encodeURIComponent(batchId)}/cancel`, { method: 'POST' });
    const data = await response.json();
    if (!response.ok) {
      alert(data.detail || '暂停本批失败');
      return;
    }
    await loadAll();
    if (openTaskId) {
      await loadDetail(openTaskId, true);
    }
  } catch (error) {
    console.error('cancelBatchGroup', error);
    alert('暂停本批请求失败，请稍后重试。');
  }
}

async function cancelTaskBatch(taskIdsValue) {
  const taskIds = parseTaskIds(taskIdsValue);
  if (!taskIds.length) {
    alert('没有可暂停的批量任务。');
    return;
  }
  if (!confirm('确定要暂停本批还在排队或处理中的任务吗？运行中的任务会在当前步骤结束后停止。')) {
    return;
  }
  try {
    const response = await fetch('/task/batch-cancel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ task_ids: taskIds }),
    });
    const data = await response.json();
    if (!response.ok) {
      alert(data.detail || '暂停本批失败');
      return;
    }
    await loadAll();
    if (openTaskId) {
      await loadDetail(openTaskId, true);
    }
  } catch (error) {
    console.error('cancelTaskBatch', error);
    alert('暂停本批请求失败，请稍后重试。');
  }
}

function parseTaskIds(taskIdsValue) {
  return String(taskIdsValue || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function buildFallbackArchiveName() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  const stamp = [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    '_',
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds()),
  ].join('');
  return `batch_download_${stamp}.zip`;
}

async function saveDownloadResponse(response, fallbackFilename) {
  const blob = await response.blob();
  const disposition = response.headers.get('content-disposition') || '';
  const filenameMatch = disposition.match(/filename\*=UTF-8''([^;]+)|filename="?([^"]+)"?/i);
  const filename = filenameMatch ? decodeURIComponent(filenameMatch[1] || filenameMatch[2] || fallbackFilename) : fallbackFilename;
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
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
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escJsString(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\r?\n/g, ' ');
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
