const POLL_INTERVAL = 2000;
const DETAIL_POLL_INTERVAL = 1500;

let currentPage = 1;
let currentPageSize = 20;
let currentStatusFilter = '';
let currentTypeFilter = '';
let currentKeyword = '';
let listTimer = null;
let detailTimer = null;
let openTaskId = null;

const STATUS_BADGE = {
  queued:    { cls: 'badge-queued',    icon: 'fa-clock',           text: '排队中' },
  running:   { cls: 'badge-running',   icon: 'fa-spinner fa-spin', text: '处理中' },
  done:      { cls: 'badge-done',      icon: 'fa-check',           text: '已完成' },
  failed:    { cls: 'badge-failed',    icon: 'fa-xmark',           text: '失败' },
  cancelled: { cls: 'badge-cancelled', icon: 'fa-ban',             text: '已取消' },
};


function init() {
  document.getElementById('btnRefresh').addEventListener('click', () => loadAll());
  document.getElementById('filterType').addEventListener('change', e => {
    currentTypeFilter = e.target.value;
    currentPage = 1;
    loadList();
  });

  let searchDebounce = null;
  document.getElementById('searchInput').addEventListener('input', e => {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
      currentKeyword = e.target.value.trim();
      currentPage = 1;
      loadList();
    }, 400);
  });

  document.querySelectorAll('.stat-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.stat-pill').forEach(p => p.classList.remove('active'));
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
function stopListPolling() { if (listTimer) { clearInterval(listTimer); listTimer = null; } }

async function loadAll(silent) {
  await Promise.all([loadStats(silent), loadList(silent)]);
}

async function loadStats(silent) {
  try {
    const resp = await fetch('/task/dashboard/stats');
    if (!resp.ok) return;
    const d = await resp.json();
    document.getElementById('statTotal').textContent = d.total ?? 0;
    document.getElementById('statQueued').textContent = d.queued ?? 0;
    document.getElementById('statRunning').textContent = d.running ?? 0;
    document.getElementById('statDone').textContent = d.done ?? 0;
    document.getElementById('statFailed').textContent = d.failed ?? 0;
    document.getElementById('statCancelled').textContent = d.cancelled ?? 0;
  } catch (e) { if (!silent) console.error('loadStats', e); }
}

async function loadList(silent) {
  try {
    const params = new URLSearchParams({ page: currentPage, page_size: currentPageSize });
    if (currentStatusFilter) params.set('status', currentStatusFilter);
    if (currentTypeFilter) params.set('task_type', currentTypeFilter);
    if (currentKeyword) params.set('keyword', currentKeyword);

    const resp = await fetch('/task/list?' + params);
    if (!resp.ok) return;
    const d = await resp.json();
    renderTable(d.items || []);
    renderPagination(d.total || 0, d.page || 1, d.page_size || 20);
  } catch (e) { if (!silent) console.error('loadList', e); }
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

  body.innerHTML = items.map(t => {
    const badge = STATUS_BADGE[t.status] || STATUS_BADGE.queued;
    const progress = t.progress ?? 0;
    const time = t.created_at ? formatTime(t.created_at) : '-';
    const fname = escHtml(t.filename || '-');
    const displayFname = fname.length > 40 ? fname.slice(0, 37) + '...' : fname;
    return `<tr data-id="${t.task_id}" onclick="openDetail('${t.task_id}')">
      <td>${escHtml(t.display_no || '-')}</td>
      <td>${escHtml(t.task_label || t.task_type)}</td>
      <td class="cell-filename" title="${fname}">${displayFname}</td>
      <td><span class="badge ${badge.cls}"><i class="fas ${badge.icon}"></i> ${badge.text}</span></td>
      <td><div class="mini-progress"><div class="mini-progress-fill" style="width:${progress}%"></div></div> <span style="font-size:12px;color:var(--muted)">${progress}%</span></td>
      <td class="cell-time">${time}</td>
      <td>
        <button class="btn-detail" onclick="event.stopPropagation();openDetail('${t.task_id}')"><i class="fas fa-eye"></i> 详情</button>
        ${(t.status === 'queued' || t.status === 'running') && !t.cancel_requested ? `<button class="btn-cancel-table" onclick="event.stopPropagation();cancelTask('${t.task_id}')" title="取消任务"><i class="fas fa-ban"></i></button>` : ''}
        ${t.cancel_requested && t.status === 'running' ? `<span style="font-size:11px;color:#94a3b8;margin-left:4px">取消中...</span>` : ''}
      </td>
    </tr>`;
  }).join('');
}

function renderPagination(total, page, pageSize) {
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const container = document.getElementById('pagination');
  if (totalPages <= 1) { container.innerHTML = ''; return; }

  let html = `<button ${page <= 1 ? 'disabled' : ''} onclick="goPage(${page - 1})"><i class="fas fa-chevron-left"></i></button>`;
  const start = Math.max(1, page - 2);
  const end = Math.min(totalPages, page + 2);
  if (start > 1) html += `<button onclick="goPage(1)">1</button>`;
  if (start > 2) html += `<span>...</span>`;
  for (let i = start; i <= end; i++) {
    html += `<button class="${i === page ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
  }
  if (end < totalPages - 1) html += `<span>...</span>`;
  if (end < totalPages) html += `<button onclick="goPage(${totalPages})">${totalPages}</button>`;
  html += `<button ${page >= totalPages ? 'disabled' : ''} onclick="goPage(${page + 1})"><i class="fas fa-chevron-right"></i></button>`;
  container.innerHTML = html;
}

function goPage(p) { currentPage = p; loadList(); }

// ── Detail drawer ──

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
    if (openTaskId !== taskId) { stopDetailPolling(); return; }
    loadDetail(taskId, true);
  }, DETAIL_POLL_INTERVAL);
}
function stopDetailPolling() { if (detailTimer) { clearInterval(detailTimer); detailTimer = null; } }

async function loadDetail(taskId, silent) {
  try {
    const resp = await fetch(`/task/${taskId}/detail`);
    if (!resp.ok) return;
    const d = await resp.json();
    renderDetail(d);
    if (d.status === 'done' || d.status === 'failed' || d.status === 'cancelled') stopDetailPolling();
  } catch (e) { if (!silent) console.error('loadDetail', e); }
}

function renderDetail(d) {
  const badge = STATUS_BADGE[d.status] || STATUS_BADGE.queued;
  const progress = d.progress ?? 0;

  const startedStr = d.started_at ? formatTime(d.started_at) : '-';
  const finishedStr = d.finished_at ? formatTime(d.finished_at) : '-';
  const createdStr = d.created_at ? formatTime(d.created_at) : '-';
  const duration = computeDuration(d.started_at, d.finished_at);

  let inputHtml = '';
  if (d.input_files && typeof d.input_files === 'object') {
    const entries = Object.entries(d.input_files).filter(([k, v]) => v && typeof v === 'string' && (k.endsWith('_path') || k === 'input_path'));
    if (entries.length) {
      inputHtml = entries.map(([k, v]) => {
        const name = v.split('/').pop() || v;
        return `<div class="file-item">
          <span class="fi-name" title="${escHtml(v)}"><i class="fas fa-file-import" style="color:var(--amber);margin-right:6px"></i>${escHtml(name)}</span>
          <button class="fi-dl" onclick="downloadFile('${d.task_id}','${escAttr(v)}')"><i class="fas fa-download"></i></button>
        </div>`;
      }).join('');
    }
  }

  let outputHtml = '';
  if (Array.isArray(d.output_files) && d.output_files.length) {
    outputHtml = d.output_files.map(f => {
      const displayName = f.name || f.path.split('/').pop();
      return `<div class="file-item">
        <span class="fi-name" title="${escHtml(f.path)}"><i class="fas fa-file-export" style="color:var(--green);margin-right:6px"></i>${escHtml(displayName)}</span>
        <button class="fi-dl" onclick="downloadFile('${d.task_id}','${escAttr(f.path)}','${escAttr(displayName)}')"><i class="fas fa-download"></i></button>
      </div>`;
    }).join('');
  }

  let errorHtml = '';
  if (d.status === 'failed' && d.error) {
    errorHtml = `<div class="detail-section"><h3><i class="fas fa-exclamation-triangle"></i> 错误信息</h3><div class="detail-error">${escHtml(d.error)}</div></div>`;
  }

  let logHtml = '';
  if (d.stream_log) {
    logHtml = `<div class="detail-section"><h3><i class="fas fa-terminal"></i> 运行日志</h3><pre class="detail-log">${escHtml(d.stream_log)}</pre></div>`;
  }

  document.getElementById('drawerContent').innerHTML = `
    <h2>${escHtml(d.display_no || '-')}</h2>
    <div class="drawer-subtitle">${escHtml(d.task_label || d.task_type)} &middot; <span class="badge ${badge.cls}" style="font-size:12px"><i class="fas ${badge.icon}"></i> ${badge.text}</span></div>

    <div class="detail-section">
      <h3><i class="fas fa-chart-simple"></i> 进度</h3>
      <div style="display:flex;align-items:center;gap:12px">
        <div class="detail-progress-bar" style="flex:1"><div class="detail-progress-fill" style="width:${progress}%"></div></div>
        <span style="font-weight:700;min-width:40px">${progress}%</span>
      </div>
      <div style="color:var(--muted);font-size:13px;margin-top:6px">${escHtml(d.message || '')}</div>
      ${(d.status === 'queued' || d.status === 'running') && !d.cancel_requested ? `<button class="btn-cancel" style="margin-top:10px" onclick="cancelTask('${d.task_id}')"><i class="fas fa-ban"></i> 取消任务</button>` : ''}
      ${d.cancel_requested && d.status === 'running' ? `<div style="margin-top:10px;font-size:13px;color:#94a3b8"><i class="fas fa-spinner fa-spin"></i> 正在取消，等待当前步骤完成...</div>` : ''}
    </div>

    <div class="detail-section">
      <h3><i class="fas fa-info-circle"></i> 基本信息</h3>
      <div class="detail-grid">
        <div class="detail-item"><div class="dl">文件名</div><div class="dv">${escHtml(d.filename || '-')}</div></div>
        <div class="detail-item"><div class="dl">耗时</div><div class="dv">${duration}</div></div>
        <div class="detail-item"><div class="dl">提交时间</div><div class="dv">${createdStr}</div></div>
        <div class="detail-item"><div class="dl">开始时间</div><div class="dv">${startedStr}</div></div>
        <div class="detail-item"><div class="dl">完成时间</div><div class="dv">${finishedStr}</div></div>
      </div>
    </div>

    ${inputHtml ? `<div class="detail-section"><h3><i class="fas fa-file-import"></i> 输入文件</h3><div class="file-list">${inputHtml}</div></div>` : ''}
    ${outputHtml ? `<div class="detail-section"><h3><i class="fas fa-file-export"></i> 输出文件</h3><div class="file-list">${outputHtml}</div></div>` : ''}
    ${errorHtml}
    ${logHtml}
  `;
}

function downloadFile(taskId, filePath, friendlyName) {
  let url = `/task/${taskId}/download?file_path=${encodeURIComponent(filePath)}`;
  if (friendlyName) url += `&download_name=${encodeURIComponent(friendlyName)}`;
  const a = document.createElement('a');
  a.href = url; a.download = friendlyName || ''; a.click();
}

async function cancelTask(taskId) {
  if (!confirm('确定要取消该任务吗？运行中的任务将在当前步骤完成后中止。')) return;
  try {
    const resp = await fetch(`/task/${taskId}/cancel`, { method: 'POST' });
    const d = await resp.json();
    if (!resp.ok) { alert(d.detail || '取消失败'); return; }
    loadAll();
    if (openTaskId === taskId) loadDetail(taskId);
  } catch (e) {
    console.error('cancelTask', e);
    alert('取消请求失败，请重试');
  }
}

// ── Helpers ──

function escHtml(s) {
  const d = document.createElement('div');
  d.textContent = s ?? '';
  return d.innerHTML;
}
function escAttr(s) { return (s || '').replace(/'/g, "\\'").replace(/"/g, '&quot;'); }

function formatTime(iso) {
  if (!iso) return '-';
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  } catch { return iso; }
}

function computeDuration(start, end) {
  if (!start) return '-';
  const s = new Date(start);
  const e = end ? new Date(end) : new Date();
  if (isNaN(s.getTime())) return '-';
  let diff = Math.max(0, Math.floor((e - s) / 1000));
  if (diff < 60) return `${diff}秒`;
  if (diff < 3600) return `${Math.floor(diff/60)}分${diff%60}秒`;
  return `${Math.floor(diff/3600)}时${Math.floor((diff%3600)/60)}分`;
}

init();
