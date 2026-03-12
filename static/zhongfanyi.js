const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const ruleFileInput = document.getElementById('ruleFile');
const useAiRuleCheckbox = document.getElementById('useAiRule');
const btnRun = document.getElementById('btnRun');
const btnReset = document.getElementById('btnReset');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

const progressBar = document.getElementById('progressBar');
const progressPercent = document.getElementById('progressPercent');
const progressDetails = document.getElementById('progressDetails');
const processingTitle = document.getElementById('processingTitle');
const processingText = document.getElementById('processingText');

const POLL_INTERVAL = 1500;
let pollingTimer = null;
let currentTaskId = null;

btnRun.addEventListener('click', runZhongfanyi);
btnReset.addEventListener('click', resetPage);

async function runZhongfanyi() {
    const originalFile = originalFileInput.files[0];
    const translatedFile = translatedFileInput.files[0];

    if (!originalFile || !translatedFile) {
        alert('请同时选择原文和译文文件（支持 .docx / .doc / .pdf）');
        return;
    }

    const useAiRule = useAiRuleCheckbox.checked;
    const ruleFile = ruleFileInput.files[0];
    if (useAiRule && !ruleFile) {
        alert('已勾选「使用 AI 生成规则」，请上传规则文件（pdf/docx/txt）');
        return;
    }

    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    updateProgressUI(0, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('original_file', originalFile);
        formData.append('translated_file', translatedFile);
        if (ruleFile) {
            formData.append('rule_file', ruleFile);
        }

        const url = '/task/zhongfanyi?use_ai_rule=' + (useAiRule ? 'true' : 'false');
        const resp = await fetch(url, {
            method: 'POST',
            body: formData,
        });

        if (!resp.ok) {
            let detailMsg = '';
            try {
                const errJson = await resp.json();
                const detail = errJson?.detail;
                if (typeof detail === 'string') detailMsg = detail;
                else if (detail && typeof detail === 'object')
                    detailMsg = detail.error || (Array.isArray(detail.traceback) ? detail.traceback.join('\n') : '');
            } catch (e) {}
            throw new Error(detailMsg || '请求失败: ' + resp.status);
        }

        const submitResp = await resp.json();
        if (submitResp.status === 'ACCEPTED' && submitResp.task_id) {
            updateProgressUI(5, '任务已提交，正在后台处理...');
            startPolling(submitResp.task_id);
        } else {
            if (submitResp.result) showResult(submitResp.result);
            else resetPage();
        }
    } catch (err) {
        alert('中翻译专检失败: ' + err.message);
        resetPage();
    }
}

function updateProgressUI(progress, message, details) {
    progressBar.style.setProperty('--progress', progress + '%');
    progressPercent.textContent = progress + '%';
    processingTitle.textContent = message || '中翻译专检处理中...';
    processingText.textContent = message || '正在处理...';
    if (details && details.length) {
        progressDetails.innerHTML = details.map(function (d) {
            return '<div class="detail-item">' + d + '</div>';
        }).join('');
    } else {
        progressDetails.innerHTML = '<div class="detail-item">' + (message || '') + '</div>';
    }
}

async function pollTaskStatus(taskId) {
    try {
        const resp = await fetch('/task/zhongfanyi/status/' + taskId);
        if (!resp.ok) return null;
        const status = await resp.json();
        updateProgressUI(
            status.progress || 0,
            status.message || '正在处理...',
            status.details || []
        );
        if (status.status === 'done') {
            stopPolling();
            if (status.result) showResult(status.result);
        } else if (status.status === 'failed') {
            stopPolling();
            alert('中翻译专检失败: ' + (status.error || '未知错误'));
            resetPage();
        }
        return status;
    } catch (err) {
        console.error('轮询状态出错:', err);
        return null;
    }
}

function startPolling(taskId) {
    currentTaskId = taskId;
    pollTaskStatus(taskId);
    pollingTimer = setInterval(function () { pollTaskStatus(taskId); }, POLL_INTERVAL);
}

function stopPolling() {
    if (pollingTimer) {
        clearInterval(pollingTimer);
        pollingTimer = null;
    }
    currentTaskId = null;
}

function showResult(data) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';

    var stats = data.stats || {};
    resultStats.innerHTML =
        '<div class="stat-card"><i class="fas fa-check"></i><h3>' + (stats.success ?? 0) + '</h3><p>成功</p></div>' +
        '<div class="stat-card"><i class="fas fa-times"></i><h3>' + (stats.failed ?? 0) + '</h3><p>失败</p></div>' +
        '<div class="stat-card"><i class="fas fa-forward"></i><h3>' + (stats.skipped ?? 0) + '</h3><p>跳过</p></div>';

    var reports = data.reports || {};
    var correctedPath = data.corrected_docx ? '/' + data.corrected_docx.replace(/^\/+/, '') : '';
    var links = '<a href="' + correctedPath + '" download class="download-btn"><i class="fas fa-file-word"></i> 下载修复后译文</a>';
    if (reports['正文_json']) links += '<a href="/' + reports['正文_json'] + '" download class="download-btn"><i class="fas fa-file-code"></i> 正文报告 JSON</a>';
    if (reports['页眉_json']) links += '<a href="/' + reports['页眉_json'] + '" download class="download-btn"><i class="fas fa-file-code"></i> 页眉报告 JSON</a>';
    if (reports['页脚_json']) links += '<a href="/' + reports['页脚_json'] + '" download class="download-btn"><i class="fas fa-file-code"></i> 页脚报告 JSON</a>';

    resultGrid.innerHTML = '<div class="result-item"><h3>输出文件</h3><div class="download-links">' + links + '</div></div>';
}

function resetPage() {
    originalFileInput.value = '';
    translatedFileInput.value = '';
    ruleFileInput.value = '';
    useAiRuleCheckbox.checked = false;
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
}
