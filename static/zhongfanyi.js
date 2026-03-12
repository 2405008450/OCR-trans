let originalFileInput, translatedFileInput, ruleFileInput, useAiRuleCheckbox, btnRun, btnReset;
let uploadSection, processingSection, resultSection, resultStats, resultGrid;
let progressBar, progressPercent, progressDetails, processingTitle, processingText;
let ruleEditorModal, ruleContentArea;

const POLL_INTERVAL = 1500;
let pollingTimer = null;
let currentTaskId = null;
// 本次会话编辑的规则内容，仅本任务使用，不写入磁盘
let sessionRuleContent = null;

function initElements() {
    originalFileInput = document.getElementById('originalFile');
    translatedFileInput = document.getElementById('translatedFile');
    ruleFileInput = document.getElementById('ruleFile');
    useAiRuleCheckbox = document.getElementById('useAiRule');
    btnRun = document.getElementById('btnRun');
    btnReset = document.getElementById('btnReset');
    uploadSection = document.getElementById('uploadSection');
    processingSection = document.getElementById('processingSection');
    resultSection = document.getElementById('resultSection');
    resultStats = document.getElementById('resultStats');
    resultGrid = document.getElementById('resultGrid');
    progressBar = document.getElementById('progressBar');
    progressPercent = document.getElementById('progressPercent');
    progressDetails = document.getElementById('progressDetails');
    processingTitle = document.getElementById('processingTitle');
    processingText = document.getElementById('processingText');
    ruleEditorModal = document.getElementById('ruleEditorModal');
    ruleContentArea = document.getElementById('ruleContent');

    if (btnRun) btnRun.addEventListener('click', runZhongfanyi);
    if (btnReset) btnReset.addEventListener('click', resetPage);

    var btnEditRule = document.getElementById('btnEditRule');
    if (btnEditRule) btnEditRule.addEventListener('click', openRuleEditor);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initElements);
} else {
    initElements();
}

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
        if (sessionRuleContent && sessionRuleContent.trim()) {
            formData.append('session_rule_content', sessionRuleContent.trim());
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
    sessionRuleContent = null;
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    stopPolling();
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
}

// 规则编辑器相关逻辑
function openRuleEditor() {
    if (!ruleEditorModal || !ruleContentArea) {
        alert('页面元素未就绪，请刷新后重试');
        return;
    }
    ruleEditorModal.style.display = 'flex';
    loadRuleContent();
}

function closeRuleEditor() {
    if (ruleEditorModal) ruleEditorModal.style.display = 'none';
}

async function loadRuleContent() {
    if (!ruleContentArea) return;
    var radio = document.querySelector('input[name="ruleType"]:checked');
    var ruleType = radio ? radio.value : 'custom';
    ruleContentArea.value = '加载中...';
    ruleContentArea.disabled = true;
    try {
        const resp = await fetch('/task/zhongfanyi/rule?rule_type=' + ruleType);
        if (!resp.ok) throw new Error('加载失败');
        const data = await resp.json();
        ruleContentArea.value = data.content;
    } catch (err) {
        ruleContentArea.value = '加载规则文件失败: ' + err.message;
    } finally {
        ruleContentArea.disabled = false;
    }
}

function saveRuleContent() {
    if (!ruleContentArea) return;
    var content = ruleContentArea.value;
    sessionRuleContent = content;
    alert('已保存为本次使用，专检时将采用此规则；不会修改磁盘上的规则文件。');
    closeRuleEditor();
}

// 点击模态框外部关闭
document.addEventListener('click', function(event) {
    if (ruleEditorModal && event.target === ruleEditorModal) {
        closeRuleEditor();
    }
});
