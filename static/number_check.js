const originalFileInput = document.getElementById('originalFile');
const translatedFileInput = document.getElementById('translatedFile');
const btnRunCheck = document.getElementById('btnRunCheck');
const btnReset = document.getElementById('btnReset');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

// 进度条相关元素
const progressBar = document.getElementById('progressBar');
const progressPercent = document.getElementById('progressPercent');
const progressDetails = document.getElementById('progressDetails');
const progressStep = document.getElementById('progressStep');
const processingTitle = document.getElementById('processingTitle');
const processingText = document.getElementById('processingText');

// 轮询间隔 (毫秒)
const POLL_INTERVAL = 1000;
let pollingTimer = null;
let currentTaskId = null;

btnRunCheck.addEventListener('click', runNumberCheck);
btnReset.addEventListener('click', resetPage);

async function runNumberCheck() {
    const originalFile = originalFileInput.files[0];
    const translatedFile = translatedFileInput.files[0];

    if (!originalFile || !translatedFile) {
        alert('请同时选择原文和译文 .docx 文件');
        return;
    }

    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';

    // 初始化进度条
    updateProgressUI(0, '正在提交任务...');

    try {
        const formData = new FormData();
        formData.append('original_file', originalFile);
        formData.append('translated_file', translatedFile);

        // 提交任务（立即返回task_id）
        const resp = await fetch('/task/number-check', {
            method: 'POST',
            body: formData,
        });

        if (!resp.ok) {
            let detailMsg = '';
            try {
                const errJson = await resp.json();
                const detail = errJson?.detail;
                if (typeof detail === 'string') {
                    detailMsg = detail;
                } else if (detail && typeof detail === 'object') {
                    const trace = Array.isArray(detail.traceback) ? detail.traceback.join('\n') : '';
                    detailMsg = `${detail.error || ''}\n${trace}`.trim();
                }
            } catch (e) {
                // ignore JSON parse error and fallback to status
            }
            throw new Error(detailMsg || `请求失败: ${resp.status}`);
        }

        const submitResp = await resp.json();

        if (submitResp.status === 'ACCEPTED' && submitResp.task_id) {
            // 开始轮询进度
            updateProgressUI(5, '任务已提交，正在后台处理...');
            startPolling(submitResp.task_id);
        } else {
            // 兼容旧版本的同步返回
            showResult(submitResp);
        }
    } catch (err) {
        alert(`数字专检失败: ${err.message}`);
        resetPage();
    }
}

// 更新进度条UI
function updateProgressUI(progress, message, details = []) {
    progressBar.style.setProperty('--progress', `${progress}%`);
    progressPercent.textContent = `${progress}%`;
    processingTitle.textContent = message || '数字专检处理中...';
    processingText.textContent = message || '正在处理...';

    // 更新详情
    if (details && details.length > 0) {
        progressDetails.innerHTML = details.map(d => `<div class="detail-item">${d}</div>`).join('');
    } else {
        progressDetails.innerHTML = `<div class="detail-item">${message}</div>`;
    }
}

// 轮询任务状态
async function pollTaskStatus(taskId) {
    try {
        const resp = await fetch(`/task/number-check/status/${taskId}`);
        if (!resp.ok) {
            console.error('获取任务状态失败');
            return null;
        }

        const status = await resp.json();

        // 更新进度条
        updateProgressUI(
            status.progress || 0,
            status.message || '正在处理...',
            status.details || []
        );

        // 检查任务是否完成
        if (status.status === 'done') {
            stopPolling();
            if (status.result) {
                showResult(status.result);
            }
        } else if (status.status === 'failed') {
            stopPolling();
            alert(`数字专检失败: ${status.error || '未知错误'}`);
            resetPage();
        }

        return status;
    } catch (err) {
        console.error('轮询任务状态出错:', err);
        return null;
    }
}

// 开始轮询
function startPolling(taskId) {
    currentTaskId = taskId;
    // 立即查询一次
    pollTaskStatus(taskId);
    // 设置定时轮询
    pollingTimer = setInterval(() => {
        pollTaskStatus(taskId);
    }, POLL_INTERVAL);
}

// 停止轮询
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

    const stats = data.stats || {};
    resultStats.innerHTML = `
        <div class="stat-card">
            <i class="fas fa-check"></i>
            <h3>${stats.success ?? 0}</h3>
            <p>成功</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-times"></i>
            <h3>${stats.failed ?? 0}</h3>
            <p>失败</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-forward"></i>
            <h3>${stats.skipped ?? 0}</h3>
            <p>跳过</p>
        </div>
    `;

    const reports = data.reports || {};
    resultGrid.innerHTML = `
        <div class="result-item">
            <h3>输出文件</h3>
            <div class="download-links">
                <a href="/${data.corrected_docx}" download class="download-btn">
                    <i class="fas fa-file-word"></i> 下载修复后译文
                </a>
                ${reports.body_json ? `<a href="/${reports.body_json}" download class="download-btn"><i class="fas fa-file-code"></i> 正文报告JSON</a>` : ''}
                ${reports.header_json ? `<a href="/${reports.header_json}" download class="download-btn"><i class="fas fa-file-code"></i> 页眉报告JSON</a>` : ''}
                ${reports.footer_json ? `<a href="/${reports.footer_json}" download class="download-btn"><i class="fas fa-file-code"></i> 页脚报告JSON</a>` : ''}
            </div>
        </div>
    `;
}

function resetPage() {
    originalFileInput.value = '';
    translatedFileInput.value = '';
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';

    // 停止轮询
    stopPolling();

    // 重置进度条
    progressBar.style.setProperty('--progress', '0%');
    progressPercent.textContent = '0%';
    progressDetails.innerHTML = '';
}
