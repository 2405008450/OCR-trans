const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const sourceLangSelect = document.getElementById('sourceLang');
const targetLangSelect = document.getElementById('targetLang');
const modelSelect = document.getElementById('modelSelect');
const modelDesc = document.getElementById('modelDesc');
const enablePostSplit = document.getElementById('enablePostSplit');
const btnStart = document.getElementById('btnStart');
const btnReset = document.getElementById('btnReset');

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

const POLL_INTERVAL = 1500;
let pollingTimer = null;
let configData = null;

// 初始化：加载配置
(async function init() {
    try {
        const resp = await fetch('/task/alignment/config');
        if (resp.ok) {
            configData = await resp.json();
            populateSelects();
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
    for (const [name, desc] of Object.entries(models)) {
        modelSelect.add(new Option(name, name));
    }
    updateModelDesc();
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
    modelSelect.add(new Option('Google: Gemini 3 Pro Preview', 'Google: Gemini 3 Pro Preview'));
}

function updateModelDesc() {
    const name = modelSelect.value;
    const desc = configData?.models?.[name] || '';
    modelDesc.textContent = desc;
}

modelSelect.addEventListener('change', updateModelDesc);
btnStart.addEventListener('click', startAlignment);
btnReset.addEventListener('click', resetPage);

async function startAlignment() {
    const origFile = originalFileInput.files[0];
    const transFile = translatedFileInput.files[0];

    if (!origFile || !transFile) {
        alert('请同时选择原文和译文文件');
        return;
    }

    const allowedExt = ['.docx', '.pptx', '.xlsx', '.xls'];
    const origExt = origFile.name.substring(origFile.name.lastIndexOf('.')).toLowerCase();
    const transExt = transFile.name.substring(transFile.name.lastIndexOf('.')).toLowerCase();

    if (!allowedExt.includes(origExt)) {
        alert(`不支持的原文文件格式: ${origExt}\n支持: DOCX, PPTX, XLSX`);
        return;
    }
    if (!allowedExt.includes(transExt)) {
        alert(`不支持的译文文件格式: ${transExt}\n支持: DOCX, PPTX, XLSX`);
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

        if (status.status === 'done') {
            stopPolling();
            showResult(status.result);
        } else if (status.status === 'failed') {
            stopPolling();
            alert(`对齐失败: ${status.error || '未知错误'}`);
            resetPage();
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
}
