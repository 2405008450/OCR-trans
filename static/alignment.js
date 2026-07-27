const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const sourceLangSelect = document.getElementById('sourceLang');
const targetLangSelect = document.getElementById('targetLang');
const modelSelect = document.getElementById('modelSelect');
let geminiRouteSelect = document.getElementById('geminiRouteSelect');
let geminiRouteDesc = document.getElementById('geminiRouteDesc');
const modelDesc = document.getElementById('modelDesc');
const modelIdDisplay = document.getElementById('modelIdDisplay');
const modelMaxOutput = document.getElementById('modelMaxOutput');
const enablePostSplit = document.getElementById('enablePostSplit');
const btnStart = document.getElementById('btnStart');
const btnReset = document.getElementById('btnReset');
const origFileLabel = document.getElementById('origFileLabel');
const transFileLabel = document.getElementById('transFileLabel');
const langHintText = document.getElementById('langHintText');
const languageRoute = document.getElementById('languageRoute');
const swapLanguagesBtn = document.getElementById('swapLanguages');
const sourceLangPicker = document.getElementById('sourceLangPicker');
const targetLangPicker = document.getElementById('targetLangPicker');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const resultSummary = document.getElementById('resultSummary');
const resultGrid = document.getElementById('resultGrid');

const progressBar = document.getElementById('progressBar');
const progressPercent = document.getElementById('progressPercent');
const progressDetails = document.getElementById('progressDetails');
const processingTitle = document.getElementById('processingTitle');
const processingText = document.getElementById('processingText');
const streamLogWrap = document.getElementById('streamLogWrap');
const streamLogEl = document.getElementById('streamLog');

const POLL_INTERVAL = 1500;
const ETA_TIME_ZONE = 'Asia/Shanghai';
let etaHint = null;
let isSubmitting = false;

function ensureEtaHint() {
    if (etaHint && etaHint.isConnected) return etaHint;
    const card = processingSection?.querySelector('.processing-card') || processingSection;
    if (!card) return null;
    etaHint = document.createElement('div');
    etaHint.className = 'eta-hint';
    etaHint.style.cssText = 'margin-top:10px;color:var(--text-secondary, var(--muted, #94a3b8));font-size:13px;';
    etaHint.textContent = '预计完成时间：计算中...';
    const anchor = typeof processingText !== 'undefined' && processingText ? processingText : null;
    if (anchor?.parentNode) {
        anchor.parentNode.insertBefore(etaHint, anchor.nextSibling);
    } else {
        card.appendChild(etaHint);
    }
    return etaHint;
}

function updateEtaHint(task) {
    const el = ensureEtaHint();
    if (!el) return;
    const text = buildEtaText(task);
    if (!text) {
        el.style.display = 'none';
        el.textContent = '';
        return;
    }
    el.style.display = 'block';
    el.textContent = text;
}

function buildEtaText(task) {
    if (!task) return '预计完成时间：计算中...';
    if (task.status === 'failed' || task.status === 'cancelled') return '';
    if (task.status === 'done' && task.finished_at) {
        return `预计完成时间：${formatEtaMinute(task.finished_at)}`;
    }
    if (task.status === 'queued') {
        return '预计完成时间：排队中，开始处理后计算';
    }

    const progress = Number(task.progress ?? 0);
    if (!Number.isFinite(progress) || progress <= 0 || progress >= 100 || !task.created_at) {
        return '预计完成时间：计算中...';
    }

    const createdAt = parseServerTime(task.created_at);
    if (Number.isNaN(createdAt.getTime())) {
        return '预计完成时间：计算中...';
    }

    const elapsedMs = Date.now() - createdAt.getTime();
    if (elapsedMs <= 0) {
        return '预计完成时间：计算中...';
    }

    const estimatedTotalMs = elapsedMs / (progress / 100);
    const estimatedFinishedAt = new Date(createdAt.getTime() + estimatedTotalMs);
    return `预计完成时间：${formatEtaDate(estimatedFinishedAt)}`;
}

function parseServerTime(iso) {
    if (!iso) return new Date(NaN);
    const normalized = /([zZ]|[+\-]\d{2}:\d{2})$/.test(iso) ? iso : `${iso}Z`;
    return new Date(normalized);
}

function formatEtaMinute(iso) {
    const date = parseServerTime(iso);
    if (Number.isNaN(date.getTime())) return '-';
    return formatEtaDate(date);
}

function formatEtaDate(date) {
    const parts = new Intl.DateTimeFormat('zh-CN', {
        timeZone: ETA_TIME_ZONE,
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
    }).formatToParts(date);
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return `${values.month}-${values.day} ${values.hour}:${values.minute}`;
}
const DEFAULT_MODEL_NAME = 'Google gemini-3-flash-preview';
let pollingTimer = null;
let configData = null;
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
const MODEL_DISPLAY_NAMES = {
    'Google Gemini 2.5 Flash': '快速版V1',
    'Google Gemini 2.5 Pro': '增强版V1',
    'Google gemini-3-flash-preview': '快速版V2',
    'Google: google/gemini-3.5-flash': '新模型',
    'Google: google/gemini-3.6-flash': 'Gemini 3.6 Flash',
    'Google: google/gemini-3.5-flash-lite': '极速版V3',
    'Google: google/gemini-3.1-pro-preview': '增强版V2',
    'DeepSeek-V4-Pro': 'DeepSeek-V4-Pro',
};
const LANGUAGE_ALIASES = {
    '中文': 'zh chinese mandarin 简体 繁体',
    '英语': 'en english',
    '西班牙语': 'es spanish espanol español',
    '葡语': 'pt portuguese portugues português 葡萄牙语',
    '日语': 'ja japanese 日本語',
    '俄语': 'ru russian русский',
    '韩语': 'ko korean 한국어',
    '阿语': 'ar arabic العربية 阿拉伯语',
    '法语': 'fr french francais français',
    '波兰语': 'pl polish polski',
    '意大利语': 'it italian italiano',
    '德语': 'de german deutsch',
};

(async function init() {
    ensureGeminiRouteSelect();
    try {
        const resp = await fetch('/task/alignment/config');
        if (resp.ok) {
            configData = await resp.json();
            populateSelects();
            populateThresholds();
        }
    } catch (e) {
        console.error('加载配置失败:', e);
        populateDefaults();
    }
})();

function ensureGeminiRouteSelect() {
    const routeGroup = document.getElementById('geminiRouteGroup');
    if (routeGroup) routeGroup.style.display = 'none';
    geminiRouteSelect = document.getElementById('geminiRouteSelect');
    geminiRouteDesc = document.getElementById('geminiRouteDesc');
}

function populateSelects() {
    const langs = configData?.languages || {};
    const models = configData?.models || {};

    sourceLangSelect.innerHTML = '';
    targetLangSelect.innerHTML = '';
    for (const [name, desc] of Object.entries(langs)) {
        sourceLangSelect.add(new Option(`${name} (${desc})`, name));
        targetLangSelect.add(new Option(`${name} (${desc})`, name));
    }
    sourceLangSelect.value = '中文';
    targetLangSelect.value = '英语';
    renderLanguagePickers();

    modelSelect.innerHTML = '';
    for (const name of Object.keys(models)) {
        modelSelect.add(new Option(getModelDisplayName(name), name));
    }
    if (models[DEFAULT_MODEL_NAME]) {
        modelSelect.value = DEFAULT_MODEL_NAME;
    }
    const routes = configData?.routes || {};
    const defaultRoute = configData?.default_route || "openrouter";
    Object.entries(routes).forEach(([value, info]) => {
        geminiRouteSelect.add(new Option(info.label || value, value));
    });
    geminiRouteSelect.value = routes[defaultRoute] ? defaultRoute : Object.keys(routes)[0];
    updateRouteInfo();
    updateModelInfo();
    updateLangLabels();
}

function populateThresholds() {
    const th = configData?.thresholds || {};
    const buf = configData?.buffer_chars || 2000;
    if (th[2]) document.getElementById('threshold2').value = th[2];
    if (th[3]) document.getElementById('threshold3').value = th[3];
    if (th[4]) document.getElementById('threshold4').value = th[4];
    if (th[5]) document.getElementById('threshold5').value = th[5];
    if (th[6]) document.getElementById('threshold6').value = th[6];
    if (th[7]) document.getElementById('threshold7').value = th[7];
    if (th[8]) document.getElementById('threshold8').value = th[8];
    document.getElementById('bufferChars').value = buf;
}

function populateDefaults() {
    const defaultLangs = ['中文', '英语', '日语', '韩语', '法语', '德语', '西班牙语', '俄语'];
    sourceLangSelect.innerHTML = '';
    targetLangSelect.innerHTML = '';
    for (const l of defaultLangs) {
        sourceLangSelect.add(new Option(l, l));
        targetLangSelect.add(new Option(l, l));
    }
    sourceLangSelect.value = '中文';
    targetLangSelect.value = '英语';
    renderLanguagePickers();

    modelSelect.innerHTML = '';
    modelSelect.add(new Option(getModelDisplayName('Google gemini-3-flash-preview'), 'Google gemini-3-flash-preview'));
    modelSelect.add(new Option(getModelDisplayName('Google: google/gemini-3.5-flash'), 'Google: google/gemini-3.5-flash'));
    modelSelect.add(new Option(getModelDisplayName('Google: google/gemini-3.6-flash'), 'Google: google/gemini-3.6-flash'));
    modelSelect.add(new Option(getModelDisplayName('Google: google/gemini-3.5-flash-lite'), 'Google: google/gemini-3.5-flash-lite'));
    modelSelect.add(new Option(getModelDisplayName('Google Gemini 2.5 Pro'), 'Google Gemini 2.5 Pro'));
    modelSelect.add(new Option(getModelDisplayName('DeepSeek-V4-Pro'), 'DeepSeek-V4-Pro'));
    modelSelect.value = DEFAULT_MODEL_NAME;
    geminiRouteSelect.innerHTML = '<option value="google">\u7ebf\u8def1</option><option value="openrouter">\u7ebf\u8def2</option>';
    geminiRouteSelect.value = "openrouter";
    updateRouteInfo();
    updateModelInfo();
    updateLangLabels();
}

function updateModelInfo() {
    const name = modelSelect.value;
    const info = configData?.models?.[name];
    if (info) {
        modelDesc.textContent = info.description || '';
        modelIdDisplay.textContent = getModelDisplayName(name);
        modelMaxOutput.textContent = info.max_output_display || (info.max_output ? `${info.max_output.toLocaleString()} tokens` : '-');
    } else {
        modelDesc.textContent = '';
        modelIdDisplay.textContent = '-';
        modelMaxOutput.textContent = '-';
    }
}

function updateLangLabels() {
    const src = sourceLangSelect.value;
    const tgt = targetLangSelect.value;
    origFileLabel.textContent = `原文文件 (${src}):`;
    transFileLabel.textContent = `译文文件 (${tgt}):`;

    const srcDesc = configData?.languages?.[src] || src;
    const tgtDesc = configData?.languages?.[tgt] || tgt;
    if (languageRoute) {
        languageRoute.innerHTML = `<i class="fas fa-arrow-right-arrow-left"></i><span>${escapeLanguageHtml(src)} <b>→</b> ${escapeLanguageHtml(tgt)}</span>`;
    }
    langHintText.textContent = `${srcDesc} → ${tgtDesc}；LLM 对齐后会按源语言再次检查并拆分多句原文键`;
    updateLanguagePickerTrigger(sourceLangPicker, sourceLangSelect);
    updateLanguagePickerTrigger(targetLangPicker, targetLangSelect);
}

function escapeLanguageHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function getLanguageDescription(name) {
    return configData?.languages?.[name] || name;
}

function getLanguageSearchText(name) {
    return `${name} ${getLanguageDescription(name)} ${LANGUAGE_ALIASES[name] || ''}`.toLocaleLowerCase();
}

function updateLanguagePickerTrigger(picker, select) {
    if (!picker || !select) return;
    const value = select.value;
    const strong = picker.querySelector('.language-picker-value strong');
    const detail = picker.querySelector('.language-picker-value span');
    if (strong) strong.textContent = value || '请选择语言';
    if (detail) detail.textContent = value ? getLanguageDescription(value) : '尚未选择';
}

function renderLanguagePickerOptions(picker, select, oppositeSelect) {
    if (!picker || !select) return;
    const container = picker.querySelector('.language-options');
    const search = picker.querySelector('input[type="search"]');
    if (!container) return;

    const query = (search?.value || '').trim().toLocaleLowerCase();
    const names = Array.from(select.options)
        .map((option) => option.value)
        .filter((name) => !query || getLanguageSearchText(name).includes(query));

    container.innerHTML = '';
    if (!names.length) {
        const empty = document.createElement('div');
        empty.className = 'language-empty-state';
        empty.innerHTML = '<i class="fas fa-magnifying-glass"></i> 没有找到匹配的语言';
        container.appendChild(empty);
        return;
    }

    names.forEach((name) => {
        const isSelected = name === select.value;
        const isOpposite = name === oppositeSelect?.value && !isSelected;
        const option = document.createElement('button');
        option.type = 'button';
        option.className = `language-option${isSelected ? ' selected' : ''}`;
        option.setAttribute('role', 'option');
        option.setAttribute('aria-selected', String(isSelected));
        option.disabled = isOpposite;
        option.title = isOpposite ? '原文和译文语言不能相同' : `${name} · ${getLanguageDescription(name)}`;

        const text = document.createElement('span');
        text.className = 'language-option-text';
        const title = document.createElement('strong');
        title.textContent = name;
        const desc = document.createElement('small');
        desc.textContent = isOpposite ? '已用于另一侧' : getLanguageDescription(name);
        text.append(title, desc);

        const status = document.createElement('i');
        status.className = `language-option-status fas ${isSelected ? 'fa-check' : 'fa-chevron-right'}`;
        option.append(text, status);
        option.addEventListener('click', () => {
            select.value = name;
            select.dispatchEvent(new Event('change', { bubbles: true }));
            closeLanguagePicker(picker);
        });
        container.appendChild(option);
    });
}

function closeLanguagePicker(picker) {
    if (!picker) return;
    const trigger = picker.querySelector('.language-picker-trigger');
    const menu = picker.querySelector('.language-picker-menu');
    picker.classList.remove('open');
    trigger?.setAttribute('aria-expanded', 'false');
    if (menu) menu.hidden = true;
}

function openLanguagePicker(picker, select, oppositeSelect) {
    [sourceLangPicker, targetLangPicker].forEach((item) => {
        if (item !== picker) closeLanguagePicker(item);
    });
    const trigger = picker?.querySelector('.language-picker-trigger');
    const menu = picker?.querySelector('.language-picker-menu');
    const search = picker?.querySelector('input[type="search"]');
    if (!picker || !trigger || !menu) return;
    picker.classList.add('open');
    trigger.setAttribute('aria-expanded', 'true');
    menu.hidden = false;
    if (search) search.value = '';
    renderLanguagePickerOptions(picker, select, oppositeSelect);
    window.requestAnimationFrame(() => search?.focus());
}

function setupLanguagePicker(picker, select, oppositeSelect) {
    if (!picker || !select) return;
    const trigger = picker.querySelector('.language-picker-trigger');
    const search = picker.querySelector('input[type="search"]');
    trigger?.addEventListener('click', () => {
        if (picker.classList.contains('open')) {
            closeLanguagePicker(picker);
        } else {
            openLanguagePicker(picker, select, oppositeSelect);
        }
    });
    search?.addEventListener('input', () => renderLanguagePickerOptions(picker, select, oppositeSelect));
    search?.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            closeLanguagePicker(picker);
            trigger?.focus();
        }
    });
}

function renderLanguagePickers() {
    updateLanguagePickerTrigger(sourceLangPicker, sourceLangSelect);
    updateLanguagePickerTrigger(targetLangPicker, targetLangSelect);
    if (sourceLangPicker?.classList.contains('open')) {
        renderLanguagePickerOptions(sourceLangPicker, sourceLangSelect, targetLangSelect);
    }
    if (targetLangPicker?.classList.contains('open')) {
        renderLanguagePickerOptions(targetLangPicker, targetLangSelect, sourceLangSelect);
    }
}

function updateRouteInfo() {
    const info = configData?.routes?.[geminiRouteSelect.value];
    geminiRouteDesc.textContent = info?.description || "";
}

function getModelDisplayName(name) {
    return MODEL_DISPLAY_NAMES[name] || name;
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

function renderStreamLog(logText) {
    if (!streamLogWrap || !streamLogEl) return;
    const sanitized = sanitizeStreamLog(logText);
    if (!sanitized) {
        streamLogWrap.style.display = 'none';
        streamLogEl.textContent = '';
        return;
    }
    streamLogWrap.style.display = 'block';
    streamLogEl.textContent = sanitized;
    streamLogEl.scrollTop = streamLogEl.scrollHeight;
}

modelSelect.addEventListener('change', updateModelInfo);
geminiRouteSelect?.addEventListener('change', updateRouteInfo);
sourceLangSelect.addEventListener('change', updateLangLabels);
targetLangSelect.addEventListener('change', updateLangLabels);
setupLanguagePicker(sourceLangPicker, sourceLangSelect, targetLangSelect);
setupLanguagePicker(targetLangPicker, targetLangSelect, sourceLangSelect);
swapLanguagesBtn?.addEventListener('click', () => {
    const previousSource = sourceLangSelect.value;
    sourceLangSelect.value = targetLangSelect.value;
    targetLangSelect.value = previousSource;
    updateLangLabels();
    renderLanguagePickers();
});
document.addEventListener('click', (event) => {
    [sourceLangPicker, targetLangPicker].forEach((picker) => {
        if (picker && !picker.contains(event.target)) closeLanguagePicker(picker);
    });
});
document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
        closeLanguagePicker(sourceLangPicker);
        closeLanguagePicker(targetLangPicker);
    }
});
btnStart.addEventListener('click', startAlignment);
btnReset.addEventListener('click', resetPage);

async function startAlignment() {
    if (isSubmitting) return;
    const origFile = originalFileInput.files[0];
    const transFile = translatedFileInput.files[0];

    if (!origFile || !transFile) {
        alert('请同时选择原文和译文文件');
        return;
    }

    const allowedExt = ['.docx', '.doc', '.pptx', '.xlsx', '.xls'];
    const origExt = origFile.name.substring(origFile.name.lastIndexOf('.')).toLowerCase();
    const transExt = transFile.name.substring(transFile.name.lastIndexOf('.')).toLowerCase();

    if (!allowedExt.includes(origExt)) {
        alert(`不支持的原文文件格式: ${origExt}\n支持: DOCX, DOC, PPTX, XLSX, XLS`);
        return;
    }
    if (!allowedExt.includes(transExt)) {
        alert(`不支持的译文文件格式: ${transExt}\n支持: DOCX, DOC, PPTX, XLSX, XLS`);
        return;
    }

    isSubmitting = true;
    btnStart.disabled = true;
    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    updateProgressUI(0, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('original_file', origFile);
        formData.append('translated_file', transFile);

        const params = new URLSearchParams({
            source_lang: sourceLangSelect.value,
            target_lang: targetLangSelect.value,
            model_name: modelSelect.value,
            gemini_route: geminiRouteSelect?.value || (configData?.default_route || 'openrouter'),
            enable_post_split: enablePostSplit.checked,
            threshold_2: document.getElementById('threshold2').value,
            threshold_3: document.getElementById('threshold3').value,
            threshold_4: document.getElementById('threshold4').value,
            threshold_5: document.getElementById('threshold5').value,
            threshold_6: document.getElementById('threshold6').value,
            threshold_7: document.getElementById('threshold7').value,
            threshold_8: document.getElementById('threshold8').value,
            buffer_chars: document.getElementById('bufferChars').value,
        });

        const resp = await fetch(`/task/alignment?${params}`, {
            method: 'POST',
            body: formData,
        });

        if (!resp.ok) {
            let msg = `请求失败: ${resp.status}`;
            try {
                const err = await resp.json();
                msg = err?.detail || msg;
            } catch (_) { }
            throw new Error(msg);
        }

        const data = await resp.json();
        if (data.status === 'ACCEPTED' && data.task_id) {
            updateProgressUI(5, '任务已提交，正在后台处理...');
            startPolling(data.task_id);
        }
    } catch (err) {
        alert(`提交失败: ${err.message}`);
        resetPage();
    } finally {
        if (uploadSection.style.display !== 'none') {
            isSubmitting = false;
            btnStart.disabled = false;
        }
    }
}

function updateProgressUI(progress, message, task = null) {
    progressBar.style.setProperty('--progress', `${progress}%`);
    progressPercent.textContent = `${progress}%`;
    processingTitle.textContent = message || '文档对齐处理中...';
    processingText.textContent = message || '正在处理...';
    updateEtaHint(task);
    progressDetails.innerHTML = `<div class="detail-item">${message}</div>`;
}

function startPolling(taskId) {
    pollStatus(taskId);
    pollingTimer = setInterval(() => pollStatus(taskId), POLL_INTERVAL);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
}

async function pollStatus(taskId) {
    try {
        const resp = await fetch(`/task/alignment/status/${taskId}`);
        if (!resp.ok) return;
        const status = await resp.json();

        updateProgressUI(status.progress || 0, status.message || '正在处理...', status);

        const logText = status.stream_log || '';
        renderStreamLog(logText);

        if (status.status === 'done') {
            if (!status.result || !status.result.output_excel) {
                stopPolling();
                processingTitle.textContent = '对齐失败';
                processingText.textContent = '任务已结束，但没有生成可下载的对齐结果文件。';
                return;
            }
            stopPolling();
            if (status.result && status.result.stream_log) {
                renderStreamLog(status.result.stream_log);
            }
            showResult(status.result);
        } else if (status.status === 'failed') {
            stopPolling();
            if (status.stream_log) {
                renderStreamLog(status.stream_log);
            }
            // 不调用 resetPage，保留实时输出便于排查
            processingTitle.textContent = '对齐失败';
            processingText.textContent = status.error || '未知错误';
            document.querySelector('.spinner')?.style && (document.querySelector('.spinner').style.display = 'none');
            // 添加"重新开始"按钮
            const retryBtn = document.createElement('button');
            retryBtn.className = 'btn-secondary';
            retryBtn.style.marginTop = '16px';
            retryBtn.innerHTML = '<i class="fas fa-rotate-right"></i> 重新开始';
            retryBtn.onclick = resetPage;
            const card = document.querySelector('.processing-card');
            if (card && !card.querySelector('.btn-secondary')) {
                card.appendChild(retryBtn);
            }
        }
    } catch (err) {
        console.error('轮询出错:', err);
    }
}

function showResult(result) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';

    const rowCount = result.row_count || 0;
    const fileType = (result.file_type || '').toUpperCase();
    const splitParts = result.split_parts || 1;

    resultSummary.innerHTML = `
        <div class="summary-card">
            <i class="fas fa-table"></i>
            <h3>${rowCount}</h3>
            <p>对齐行数</p>
        </div>
        <div class="summary-card">
            <i class="fas fa-file"></i>
            <h3>${fileType}</h3>
            <p>文件类型</p>
        </div>
        <div class="summary-card">
            <i class="fas fa-cut"></i>
            <h3>${splitParts}</h3>
            <p>分割份数</p>
        </div>
    `;

    let issuesHtml = '';
    if (result.issues && result.issues.length > 0) {
        issuesHtml = `
            <div class="issues-list">
                <h4><i class="fas fa-exclamation-triangle"></i> 质量检查警告 (${result.issues.length})</h4>
                <ul>${result.issues.map(i => `<li>${i}</li>`).join('')}</ul>
            </div>
        `;
    }

    // 中间文件下载区域
    let intermediateHtml = '';
    const files = result.intermediate_files || [];
    if (files.length > 0) {
        const fileItems = files.map(f => {
            const icon = f.type === 'excel' ? 'fa-file-excel' : f.type === 'word' ? 'fa-file-word' : 'fa-file';
            const cls = f.type === 'excel' ? 'download-btn-sm btn-excel' : f.type === 'word' ? 'download-btn-sm btn-word' : 'download-btn-sm';
            return `<a href="/${f.path}" download class="${cls}"><i class="fas ${icon}"></i> ${f.name}</a>`;
        }).join('');

        intermediateHtml = `
            <div class="intermediate-files">
                <div class="intermediate-header" onclick="this.parentElement.classList.toggle('open')">
                    <i class="fas fa-chevron-right"></i>
                    <span>中间处理文件 (${files.length})</span>
                </div>
                <div class="intermediate-list">
                    ${fileItems}
                </div>
            </div>
        `;
    }

    resultGrid.innerHTML = `
        <div class="result-item">
            <h3>输出文件</h3>
            <div class="download-links">
                <a href="/${result.output_excel}" download class="download-btn">
                    <i class="fas fa-file-excel"></i> 下载对齐结果 Excel
                </a>
            </div>
            ${issuesHtml}
            ${intermediateHtml}
        </div>
    `;
}

function resetPage() {
    isSubmitting = false;
    btnStart.disabled = false;
    originalFileInput.value = '';
    translatedFileInput.value = '';
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
    streamLogWrap.style.display = 'none';
    streamLogEl.textContent = '';
}

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}


