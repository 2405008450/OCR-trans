const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const sourceLangSelect = document.getElementById('sourceLang');
const targetLangSelect = document.getElementById('targetLang');
const modelSelect = document.getElementById('modelSelect');
const modelDesc = document.getElementById('modelDesc');
const modelIdDisplay = document.getElementById('modelIdDisplay');
const modelMaxOutput = document.getElementById('modelMaxOutput');
const enablePostSplit = document.getElementById('enablePostSplit');
const btnStart = document.getElementById('btnStart');
const btnReset = document.getElementById('btnReset');
const origFileLabel = document.getElementById('origFileLabel');
const transFileLabel = document.getElementById('transFileLabel');
const langHintText = document.getElementById('langHintText');

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
let pollingTimer = null;
let configData = null;

(async function init() {
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

    modelSelect.innerHTML = '';
    for (const name of Object.keys(models)) {
        modelSelect.add(new Option(name, name));
    }
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

    modelSelect.innerHTML = '';
    modelSelect.add(new Option('Google Gemini 2.5 Flash', 'Google Gemini 2.5 Flash'));
    modelSelect.add(new Option('Google Gemini 2.5 Pro', 'Google Gemini 2.5 Pro'));
}

function updateModelInfo() {
    const name = modelSelect.value;
    const info = configData?.models?.[name];
    if (info) {
        modelDesc.textContent = info.description || '';
        modelIdDisplay.textContent = info.id || '-';
        modelMaxOutput.textContent = info.max_output ? `${info.max_output.toLocaleString()} tokens` : '-';
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
    langHintText.textContent = `${srcDesc} → ${tgtDesc}`;
}

modelSelect.addEventListener('change', updateModelInfo);
sourceLangSelect.addEventListener('change', updateLangLabels);
targetLangSelect.addEventListener('change', updateLangLabels);
btnStart.addEventListener('click', startAlignment);
btnReset.addEventListener('click', resetPage);

async function startAlignment() {
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
            } catch (_) {}
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
    }
}

function updateProgressUI(progress, message) {
    progressBar.style.setProperty('--progress', `${progress}%`);
    progressPercent.textContent = `${progress}%`;
    processingTitle.textContent = message || '文档对齐处理中...';
    processingText.textContent = message || '正在处理...';
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

        updateProgressUI(status.progress || 0, status.message || '正在处理...');

        const logText = status.stream_log || '';
        if (logText) {
            streamLogWrap.style.display = 'block';
            streamLogEl.textContent = logText;
            streamLogEl.scrollTop = streamLogEl.scrollHeight;
        }

        if (status.status === 'done') {
            stopPolling();
            if (status.result && status.result.stream_log) {
                streamLogWrap.style.display = 'block';
                streamLogEl.textContent = status.result.stream_log;
                streamLogEl.scrollTop = streamLogEl.scrollHeight;
            }
            showResult(status.result);
        } else if (status.status === 'failed') {
            stopPolling();
            if (status.stream_log) {
                streamLogWrap.style.display = 'block';
                streamLogEl.textContent = status.stream_log;
                streamLogEl.scrollTop = streamLogEl.scrollHeight;
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

    resultGrid.innerHTML = `
        <div class="result-item">
            <h3>输出文件</h3>
            <div class="download-links">
                <a href="/${result.output_excel}" download class="download-btn">
                    <i class="fas fa-file-excel"></i> 下载对齐结果 Excel
                </a>
            </div>
            ${issuesHtml}
        </div>
    `;
}

function resetPage() {
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
