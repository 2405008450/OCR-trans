const POLL_MS = 1600;
const FALLBACK_OCR_MODEL = 'google/gemini-3-flash-preview';
const FALLBACK_OCR_MODELS = {
  'google/gemini-3.1-flash-lite': { label: '极速版V2' },
  'google/gemini-3-flash-preview': { label: '快速版V2' },
  'google/gemini-3.5-flash': { label: '新模型' },
  'google/gemini-3.1-pro-preview': { label: '增强版V2' },
  'anthropic/claude-sonnet-5': { label: 'Claude Sonnet 5' },
};

let currentTaskId = '';
let pollTimer = null;
let defaultOcrModel = FALLBACK_OCR_MODEL;
let ocrModes = {};
let uploadMaxFileMb = 50;
let uploadExtensions = [];
let selectedUploadFile = null;

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
  document.querySelectorAll('input[name="ocrMode"]').forEach((input) => {
    input.addEventListener('change', updateOcrControls);
  });
  document.querySelectorAll('input[name="inputSource"]').forEach((input) => {
    input.addEventListener('change', updateInputSource);
  });
  document.getElementById('uploadFileInput').addEventListener('change', (event) => {
    setSelectedUploadFile(event.target.files?.[0] || null);
  });
  initUploadDropzone();
  initTooltips();
  updateInputSource();
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
  const cadSupport = config.cad_support && typeof config.cad_support === 'object'
    ? config.cad_support
    : {};
  const cadSupported = Array.isArray(cadSupport.supported_extensions)
    ? cadSupport.supported_extensions.join(' / ')
    : '';
  let cadHint = `CAD ${cad} 尚未配置解析工具`;
  if (cadSupport.oda_available) {
    const version = cadSupport.version ? ` ${cadSupport.version}` : '';
    cadHint = `CAD 支持：${cadSupported || cad}（ODA File Converter${version} 已就绪）`;
  } else if (cadSupport.direct_dxf) {
    cadHint = `CAD 支持：${cadSupported || '.dxf'}；DWG / DWS / DWT 需配置 ODA File Converter`;
  }
  if (cadSupport.unavailable_reason && !cadSupport.oda_available) {
    cadHint += `（${cadSupport.unavailable_reason}）`;
  }
  ocrModes = config.ocr_modes || {};
  uploadMaxFileMb = Math.max(1, Number(config.upload_max_file_mb || 50));
  uploadExtensions = Array.isArray(config.upload_extensions)
    ? config.upload_extensions.map((item) => String(item).toLowerCase())
    : [];
  const fileInput = document.getElementById('uploadFileInput');
  fileInput.accept = uploadExtensions.join(',');
  renderUploadSelection();
  const configuredModels = config.ocr_models && typeof config.ocr_models === 'object'
    ? config.ocr_models
    : {};
  const models = Object.keys(configuredModels).length ? configuredModels : FALLBACK_OCR_MODELS;
  defaultOcrModel = String(config.default_ocr_model || FALLBACK_OCR_MODEL);
  const modelSelect = document.getElementById('ocrModelSelect');
  modelSelect.innerHTML = Object.entries(models).map(([value, item]) => {
    const label = item?.label || value;
    return `<option value="${escAttr(value)}">${escHtml(label)}</option>`;
  }).join('');
  modelSelect.value = models[defaultOcrModel] ? defaultOcrModel : FALLBACK_OCR_MODEL;
  document.getElementById('formatHint').innerHTML = [
    `实际统计：${escHtml(countable)}`,
    `OCR 支持：PDF 与独立图片 ${escHtml(images)}`,
    escHtml(cadHint),
    `限制：最多 ${Number(config.max_files || 0).toLocaleString()} 个文件，单文件 ${Number(config.max_file_mb || 0)} MB`,
    `直接上传：单文件不超过 ${uploadMaxFileMb} MB`,
  ].join('<br>');
  updateOcrControls();
}

async function submitTask() {
  const inputSource = currentInputSource();
  const directoryPath = document.getElementById('directoryPath').value.trim();
  if (inputSource === 'path' && !directoryPath) {
    alert('请填写共享目录或文件路径');
    return;
  }
  if (inputSource === 'upload' && !selectedUploadFile) {
    alert('请选择需要统计的文件');
    return;
  }
  if (inputSource === 'upload' && !validateUploadFile(selectedUploadFile)) return;

  const submitBtn = document.getElementById('submitBtn');
  submitBtn.disabled = true;
  setStatus({ status: 'queued', progress: 0, message: '正在提交任务...' });

  const ocrMode = document.querySelector('input[name="ocrMode"]:checked')?.value || 'auto';
  const ocrModel = document.getElementById('ocrModelSelect').value || '';
  let endpoint = '/task/word-count';
  let requestOptions;
  if (inputSource === 'upload') {
    endpoint = '/task/word-count/upload';
    const formData = new FormData();
    formData.append('file', selectedUploadFile, selectedUploadFile.name);
    formData.append('ocr_mode', ocrMode);
    if (ocrModel) formData.append('ocr_model', ocrModel);
    requestOptions = { method: 'POST', body: formData };
  } else {
    const body = {
      directory_path: directoryPath,
      recursive: document.getElementById('recursiveInput').checked,
      include_hidden: document.getElementById('hiddenInput').checked,
      extensions: parseExtensions(document.getElementById('extensionsInput').value),
      ocr_mode: ocrMode,
      ocr_model: ocrModel || null,
    };
    requestOptions = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    };
  }

  try {
    const response = await fetch(endpoint, requestOptions);
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
  document.getElementById('cjkWords').textContent = formatNumber(cjkCandidateCount(summary));
  document.getElementById('latinWords').textContent = formatNumber(summary.total_billable_latin_count);
  document.getElementById('numberTokens').textContent = formatNumber(summary.total_number_token_count);
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
  if (result.ocr_text_archive) {
    links.push(`<a class="download-link secondary" href="${downloadUrl(task.task_id, result.ocr_text_archive, 'OCR识别文本.zip')}"><i class="fas fa-file-zipper"></i> 下载 OCR 文本</a>`);
  }
  area.innerHTML = links.join('');
  area.style.display = links.length ? 'flex' : 'none';
}

function renderFiles(files) {
  const tbody = document.getElementById('fileRows');
  if (!files.length) {
    tbody.innerHTML = '<tr><td colspan="8"><div class="empty">暂无文件明细</div></td></tr>';
    return;
  }
  const visible = files.slice(0, 200);
  tbody.innerHTML = visible.map((item) => {
    const status = item.status || '';
    const statusLabel = fileStatusText[status] || status || '-';
    const details = [];
    if (item.ocr_used) {
      details.push(`OCR ${formatNumber(item.ocr_page_count)} 页`);
      if (item.ocr_model) details.push(item.ocr_model);
      if (Array.isArray(item.ocr_failed_pages) && item.ocr_failed_pages.length) {
        details.push(`失败页 ${item.ocr_failed_pages.join(', ')}`);
      }
    }
    const primaryMessage = item.error || item.warning || item.message || '';
    if (primaryMessage) details.push(primaryMessage);
    const message = details.join(' · ');
    return `<tr>
      <td title="${escAttr(item.file_path || '')}">${escHtml(item.relative_path || item.filename || '-')}</td>
      <td><span class="badge ${escAttr(status)}">${escHtml(statusLabel)}</span></td>
      <td>${formatNumber(item.main_word_count)}</td>
      <td>${formatNumber(cjkCandidateCount(item))}</td>
      <td>${formatNumber(item.billable_latin_count)}</td>
      <td>${formatNumber(item.number_token_count)}</td>
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
  document.getElementById('inputSourcePath').checked = true;
  selectedUploadFile = null;
  document.getElementById('uploadFileInput').value = '';
  document.getElementById('ocrModeAuto').checked = true;
  document.getElementById('ocrModelSelect').value = defaultOcrModel;
  renderUploadSelection();
  updateInputSource();
  document.getElementById('taskIdText').textContent = '';
  document.getElementById('downloadArea').style.display = 'none';
  document.getElementById('fileRows').innerHTML = '<tr><td colspan="8"><div class="empty">暂无结果</div></td></tr>';
  document.getElementById('mainWords').textContent = '0';
  document.getElementById('extraWords').textContent = '0';
  document.getElementById('countedFiles').textContent = '0';
  document.getElementById('issueFiles').textContent = '0';
  document.getElementById('cjkWords').textContent = '0';
  document.getElementById('latinWords').textContent = '0';
  document.getElementById('numberTokens').textContent = '0';
  document.getElementById('submitBtn').disabled = false;
  setStatus({ status: '', progress: 0, message: '等待提交' });
}

function updateOcrControls() {
  const mode = document.querySelector('input[name="ocrMode"]:checked')?.value || 'auto';
  const modelSelect = document.getElementById('ocrModelSelect');
  if (modelSelect) modelSelect.disabled = mode === 'off';
  const hint = document.getElementById('ocrModeHint');
  if (!hint) return;
  const configured = mode === 'auto' && currentInputSource() === 'upload'
    ? '上传文件如需 OCR 将自动识别。'
    : ocrModes[mode]?.description;
  hint.textContent = configured || {
    auto: '单文件自动识别；目录任务需切换为“开启”。',
    on: '扫描 PDF 和独立图片将调用视觉模型识别。',
    off: '扫描 PDF 和图片仅标记为“需 OCR”。',
  }[mode];
}

function currentInputSource() {
  return document.querySelector('input[name="inputSource"]:checked')?.value || 'path';
}

function updateInputSource() {
  const uploadMode = currentInputSource() === 'upload';
  document.getElementById('pathInputFields').classList.toggle('is-hidden', uploadMode);
  document.getElementById('pathScanOptions').classList.toggle('is-hidden', uploadMode);
  document.getElementById('pathHints').classList.toggle('is-hidden', uploadMode);
  document.getElementById('uploadInputFields').classList.toggle('is-hidden', !uploadMode);
  document.getElementById('submitBtnText').textContent = uploadMode ? '上传并统计' : '开始统计';
  updateOcrControls();
}

function initUploadDropzone() {
  const dropzone = document.getElementById('uploadDropzone');
  ['dragenter', 'dragover'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add('drag-active');
    });
  });
  ['dragleave', 'drop'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove('drag-active');
    });
  });
  dropzone.addEventListener('drop', (event) => {
    setSelectedUploadFile(event.dataTransfer?.files?.[0] || null);
  });
}

function setSelectedUploadFile(file) {
  if (file && !validateUploadFile(file)) {
    selectedUploadFile = null;
    document.getElementById('uploadFileInput').value = '';
  } else {
    selectedUploadFile = file || null;
  }
  renderUploadSelection();
}

function validateUploadFile(file) {
  if (!file) return false;
  const extensionIndex = file.name.lastIndexOf('.');
  const extension = extensionIndex >= 0 ? file.name.slice(extensionIndex).toLowerCase() : '';
  if (uploadExtensions.length && !uploadExtensions.includes(extension)) {
    alert(`不支持的文件格式：${extension || '无扩展名'}`);
    return false;
  }
  if (file.size > uploadMaxFileMb * 1024 * 1024) {
    alert(`文件超过 ${uploadMaxFileMb} MB，请改用共享路径统计`);
    return false;
  }
  return true;
}

function renderUploadSelection() {
  const name = document.getElementById('uploadFileName');
  const meta = document.getElementById('uploadFileMeta');
  if (!name || !meta) return;
  name.textContent = selectedUploadFile?.name || '选择一个文件';
  meta.textContent = selectedUploadFile
    ? `${formatFileSize(selectedUploadFile.size)} · 最大 ${uploadMaxFileMb} MB`
    : `单文件最大 ${uploadMaxFileMb} MB`;
}

function formatFileSize(bytes) {
  const size = Math.max(0, Number(bytes || 0));
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function cjkCandidateCount(source) {
  const nested = (source && source.script_counts) || {};
  const han = numberValue(nested.han_count ?? source?.han_count ?? source?.total_han_count);
  const kana = numberValue(nested.kana_count ?? source?.kana_count ?? source?.total_kana_count);
  const hangul = numberValue(nested.hangul_count ?? source?.hangul_count ?? source?.total_hangul_count);
  const punct = numberValue(nested.cjk_punct_count ?? source?.cjk_punct_count ?? source?.total_cjk_punct_count);
  const scriptTotal = han + kana + hangul + punct;
  if (scriptTotal > 0) {
    return scriptTotal;
  }
  return Math.max(
    numberValue(source?.billable_chinese_count ?? source?.total_billable_chinese_count),
    numberValue(source?.billable_japanese_count ?? source?.total_billable_japanese_count),
    numberValue(source?.billable_korean_count ?? source?.total_billable_korean_count),
  );
}

function formatNumber(value) {
  return numberValue(value).toLocaleString();
}

function numberValue(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? number : 0;
}

function initTooltips() {
  let activeTarget = null;
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip-layer';
  document.body.appendChild(tooltip);

  const show = (target) => {
    const text = target?.getAttribute('data-tooltip') || '';
    if (!text) return;
    activeTarget = target;
    tooltip.textContent = text;
    tooltip.classList.add('visible');
    positionTooltip(target, tooltip);
  };

  const hide = (target) => {
    if (target && activeTarget && target !== activeTarget) return;
    activeTarget = null;
    tooltip.classList.remove('visible');
  };

  document.addEventListener('mouseover', (event) => {
    const target = event.target.closest?.('.hint-popover');
    if (target) show(target);
  });
  document.addEventListener('mouseout', (event) => {
    const target = event.target.closest?.('.hint-popover');
    if (target && !target.contains(event.relatedTarget)) hide(target);
  });
  document.addEventListener('focusin', (event) => {
    const target = event.target.closest?.('.hint-popover');
    if (target) show(target);
  });
  document.addEventListener('focusout', (event) => {
    const target = event.target.closest?.('.hint-popover');
    if (target) hide(target);
  });
  window.addEventListener('resize', () => {
    if (activeTarget) positionTooltip(activeTarget, tooltip);
  });
  window.addEventListener('scroll', () => {
    if (activeTarget) positionTooltip(activeTarget, tooltip);
  }, true);
}

function positionTooltip(target, tooltip) {
  const rect = target.getBoundingClientRect();
  const margin = 12;
  const maxLeft = window.innerWidth - tooltip.offsetWidth - margin;
  const left = Math.max(margin, Math.min(rect.left + rect.width / 2 - tooltip.offsetWidth / 2, maxLeft));
  const below = rect.bottom + 10;
  const above = rect.top - tooltip.offsetHeight - 10;
  const top = below + tooltip.offsetHeight + margin <= window.innerHeight
    ? below
    : Math.max(margin, above);
  tooltip.style.left = `${left}px`;
  tooltip.style.top = `${top}px`;
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
