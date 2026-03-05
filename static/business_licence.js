'use strict';

// ─────────────────────────────────────────────
// DOM 引用
// ─────────────────────────────────────────────
const uploadArea        = document.getElementById('uploadArea');
const fileInput         = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview       = document.getElementById('filePreview');
const previewImage      = document.getElementById('previewImage');
const fileNameEl        = document.getElementById('fileName');
const btnRemove         = document.getElementById('btnRemove');
const btnStart          = document.getElementById('btnStart');
const sourceLangSel     = document.getElementById('sourceLang');
const targetLangSel     = document.getElementById('targetLang');
const configFileSel     = document.getElementById('configFile');

const uploadSection     = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection     = document.getElementById('resultSection');

const processingTitle   = document.getElementById('processingTitle');
const processingStatus  = document.getElementById('processingStatus');
const progressBar       = document.getElementById('progressBar');
const progressText      = document.getElementById('progressText');
const logPanel          = document.getElementById('logPanel');

const imageCompare      = document.getElementById('imageCompare');
const downloadLink      = document.getElementById('downloadLink');
const btnRestart        = document.getElementById('btnRestart');

// 印章验证弹窗
const verifyOverlay     = document.getElementById('verifyOverlay');
const verifyInfoBox     = document.getElementById('verifyInfoBox');
const btnConfirm        = document.getElementById('btnConfirm');
const btnCorrect        = document.getElementById('btnCorrect');
const btnSkip           = document.getElementById('btnSkip');
const correctionRow     = document.getElementById('correctionRow');
const correctionInput   = document.getElementById('correctionInput');
const btnSubmitCorrection = document.getElementById('btnSubmitCorrection');

// ─────────────────────────────────────────────
// 状态
// ─────────────────────────────────────────────
let selectedFile = null;
let currentTaskId = null;
let currentEventSource = null;
let uploadedInputUrl = null;  // 原图访问地址（用于对比展示）

// ─────────────────────────────────────────────
// 文件上传交互
// ─────────────────────────────────────────────
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
});
btnRemove.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFile();
});

function handleFile(file) {
    selectedFile = file;
    fileNameEl.textContent = file.name;
    previewImage.src = URL.createObjectURL(file);
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    btnStart.disabled = false;
}

function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    previewImage.src = '';
    uploadPlaceholder.style.display = 'flex';
    filePreview.style.display = 'none';
    btnStart.disabled = true;
}

// ─────────────────────────────────────────────
// 进度 / 日志更新
// ─────────────────────────────────────────────
function setProgress(pct, message) {
    progressBar.style.setProperty('--progress', pct + '%');
    progressText.textContent = pct + '%';
    if (message) processingStatus.textContent = message;
}

function appendLog(text) {
    const line = document.createElement('span');
    if (text.includes('错误') || text.includes('Error') || text.includes('error')) {
        line.className = 'log-error';
    } else if (text.startsWith('===') || text.startsWith('---')) {
        line.className = 'log-sep';
    } else if (text.includes('完成') || text.includes('✅') || text.includes('done')) {
        line.className = 'log-success';
    }
    line.textContent = text + '\n';
    logPanel.appendChild(line);
    logPanel.scrollTop = logPanel.scrollHeight;
}

// ─────────────────────────────────────────────
// 印章验证弹窗
// ─────────────────────────────────────────────
function showVerifyModal(regionInfo) {
    const bbox = Array.isArray(regionInfo.bbox) ? regionInfo.bbox.join(', ') : regionInfo.bbox;
    verifyInfoBox.textContent =
        `类型:    ${regionInfo.type_name}\n` +
        `文字类型: ${regionInfo.text_type}\n` +
        `位置:    (${bbox})\n` +
        `置信度:  ${(regionInfo.confidence * 100).toFixed(1)}%\n` +
        `\n识别内容:\n${regionInfo.text}\n` +
        `\n(${regionInfo.index} / ${regionInfo.total})`;
    correctionRow.classList.remove('active');
    correctionInput.value = regionInfo.text || '';
    verifyOverlay.classList.add('active');
    btnConfirm.focus();
}

function hideVerifyModal() {
    verifyOverlay.classList.remove('active');
    correctionRow.classList.remove('active');
}

async function sendVerify(action, text) {
    if (!currentTaskId) return;
    hideVerifyModal();
    try {
        await fetch(`/task/business-licence/verify/${currentTaskId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action, text: text || null }),
        });
    } catch (err) {
        console.error('verify request failed', err);
    }
}

btnConfirm.addEventListener('click', () => sendVerify('confirm', null));
btnSkip.addEventListener('click',    () => sendVerify('skip',    null));
btnCorrect.addEventListener('click', () => {
    correctionRow.classList.add('active');
    correctionInput.focus();
});
btnSubmitCorrection.addEventListener('click', () => {
    const text = correctionInput.value.trim();
    if (!text) return;
    sendVerify('correct', text);
});
correctionInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        const text = correctionInput.value.trim();
        if (text) sendVerify('correct', text);
    }
});

// 键盘快捷键（仅当弹窗打开时）
document.addEventListener('keydown', (e) => {
    if (!verifyOverlay.classList.contains('active')) return;
    if (correctionRow.classList.contains('active')) return; // 修正模式下不拦截
    if (e.key === 'y' || e.key === 'Y' || e.key === 'Enter') {
        sendVerify('confirm', null);
    } else if (e.key === 'n' || e.key === 'N') {
        correctionRow.classList.add('active');
        correctionInput.focus();
    } else if (e.key === 's' || e.key === 'S' || e.key === 'Escape') {
        sendVerify('skip', null);
    }
});

// ─────────────────────────────────────────────
// SSE 事件处理
// ─────────────────────────────────────────────
function handleSSEEvent(rawData) {
    let data;
    try {
        data = JSON.parse(rawData);
    } catch {
        return;
    }

    switch (data.type) {
        case 'progress':
            setProgress(data.progress || 0, data.message || '');
            break;
        case 'log':
            appendLog(data.text || '');
            break;
        case 'verification_request':
            showVerifyModal(data.region_info);
            processingStatus.textContent = '等待印章验证...';
            break;
        case 'done':
            handleDone(data);
            break;
        case 'error':
            handleError(data.message || '未知错误');
            break;
    }
}

function handleDone(data) {
    if (currentEventSource) {
        currentEventSource.close();
        currentEventSource = null;
    }
    setProgress(100, '翻译完成');

    const outputFilename = data.output_filename;
    const outputUrl = `/bl-outputs/${outputFilename}`;

    // 下载链接
    downloadLink.href = outputUrl;
    downloadLink.download = outputFilename;

    // 图片对比
    imageCompare.innerHTML = '';
    if (uploadedInputUrl) {
        imageCompare.innerHTML += `
            <div class="image-compare-item">
                <p>原图（翻译前）</p>
                <img src="${uploadedInputUrl}" alt="原图" onclick="window.open(this.src)">
            </div>`;
    }
    imageCompare.innerHTML += `
        <div class="image-compare-item">
            <p>翻译结果</p>
            <img src="${outputUrl}" alt="翻译结果" onclick="window.open(this.src)">
        </div>`;

    processingSection.style.display = 'none';
    resultSection.style.display = 'block';
}

function handleError(message) {
    if (currentEventSource) {
        currentEventSource.close();
        currentEventSource = null;
    }
    processingTitle.textContent = '处理失败';
    processingStatus.textContent = message;
    appendLog('❌ 错误: ' + message);
    setProgress(0, '处理失败');
    btnStart.disabled = false;
}

// ─────────────────────────────────────────────
// 提交任务
// ─────────────────────────────────────────────
btnStart.addEventListener('click', async () => {
    if (!selectedFile) return;

    // 保存原图预览 URL（用于结果对比）
    uploadedInputUrl = URL.createObjectURL(selectedFile);

    // 切换到处理区域
    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    resultSection.style.display = 'none';
    logPanel.innerHTML = '';
    setProgress(0, '上传文件...');
    btnStart.disabled = true;

    const formData = new FormData();
    formData.append('file', selectedFile);

    const params = new URLSearchParams({
        source_lang: sourceLangSel.value,
        target_lang: targetLangSel.value,
        config_file: configFileSel.value,
    });

    let taskId;
    try {
        const resp = await fetch(`/task/business-licence?${params}`, {
            method: 'POST',
            body: formData,
        });
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({ detail: resp.statusText }));
            throw new Error(err.detail || resp.statusText);
        }
        const json = await resp.json();
        taskId = json.task_id;
        currentTaskId = taskId;
    } catch (err) {
        handleError('提交失败: ' + err.message);
        uploadSection.style.display = 'block';
        processingSection.style.display = 'none';
        return;
    }

    setProgress(5, '连接处理流...');

    // 连接 SSE
    const es = new EventSource(`/task/business-licence/stream/${taskId}`);
    currentEventSource = es;

    es.onmessage = (e) => handleSSEEvent(e.data);
    es.onerror = () => {
        es.close();
        currentEventSource = null;
        // 若已完成则忽略断开
    };
});

// ─────────────────────────────────────────────
// 重新翻译
// ─────────────────────────────────────────────
btnRestart.addEventListener('click', () => {
    if (currentEventSource) {
        currentEventSource.close();
        currentEventSource = null;
    }
    currentTaskId = null;
    clearFile();
    resultSection.style.display = 'none';
    processingSection.style.display = 'none';
    uploadSection.style.display = 'block';
    logPanel.innerHTML = '';
    setProgress(0, '');
});
