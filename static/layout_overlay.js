(() => {
    'use strict';
    const byId = id => document.getElementById(id);
    const form = byId('overlay-form');
    const submit = byId('submit');
    let config;
    let timer;
    let taskId;
    let generation = 0;
    const terminal = new Set(['done', 'failed', 'cancelled']);

    function populate(id, options, selected) {
        const select = byId(id);
        select.replaceChildren();
        for (const [value, info] of Object.entries(options)) {
            select.add(new Option(typeof info === 'string' ? info : info.label || info.name || value, value));
        }
        select.value = selected;
    }

    async function request(url, options = {}) {
        const response = await fetch(url, {...options, signal: AbortSignal.timeout(30000)});
        const payload = await response.json();
        if (!response.ok) {
            const detail = payload.detail;
            throw new Error(typeof detail === 'string' ? detail : '请求失败，请检查参数与服务状态');
        }
        return payload;
    }

    function showMessage(message, error = false) {
        byId('task-card').hidden = false;
        byId('message').textContent = message;
        byId('message').classList.toggle('error', error);
    }

    function renderResult(snapshot) {
        const downloads = byId('downloads');
        downloads.replaceChildren();
        for (const file of snapshot.output_files || []) {
            const url = new URL(`/task/${encodeURIComponent(taskId)}/download`, location.origin);
            url.searchParams.set('file_path', file.path);
            url.searchParams.set('download_name', file.name);
            const anchor = document.createElement('a');
            anchor.href = url;
            anchor.textContent = file.name;
            downloads.append(anchor);
        }
        const result = snapshot.result || {};
        const report = result.qa_report || {};
        const labels = {passed: '质检通过', partial: '部分质检未完成，请人工复核', needs_review: '存在待复核问题', disabled: '未启用自动质检'};
        byId('qa-status').textContent = labels[report.status] || '';
        byId('warnings').replaceChildren();
        for (const warning of (report.warnings || []).slice(0, 30)) {
            const item = document.createElement('li');
            item.textContent = warning;
            byId('warnings').append(item);
        }
        if (report.issues?.length) {
            const item = document.createElement('li');
            item.textContent = `${report.issues.length} 个问题需复核，详情见质检报告。`;
            byId('warnings').append(item);
        }
    }

    async function poll(version) {
        if (version !== generation) return;
        try {
            const snapshot = await request(`/task/layout-overlay/status/${encodeURIComponent(taskId)}`);
            if (version !== generation) return;
            byId('progress').value = snapshot.progress || 0;
            byId('task-number').textContent = `任务：${snapshot.display_no || taskId}`;
            showMessage(snapshot.error || snapshot.message || snapshot.status, snapshot.status === 'failed');
            if (terminal.has(snapshot.status)) {
                submit.disabled = false;
                sessionStorage.removeItem('layout-overlay-task');
                if (snapshot.status === 'done') renderResult(snapshot);
                return;
            }
        } catch (error) {
            if (version !== generation) return;
            showMessage(`进度查询暂时失败，将继续重试。${error.message}`, true);
        }
        timer = setTimeout(() => poll(version), 2500);
    }

    form.addEventListener('submit', async event => {
        event.preventDefault();
        const file = byId('file').files[0];
        if (!config || !file) return;
        if (file.size > config.upload_max_mb * 1024 * 1024) {
            showMessage(`文件不能超过 ${config.upload_max_mb} MB`, true);
            return;
        }
        if (byId('source_lang').value === byId('target_lang').value) {
            showMessage('请选择不同的源语言和目标语言', true);
            return;
        }
        clearTimeout(timer);
        const version = ++generation;
        const data = new FormData(form);
        data.set('enable_qa', String(byId('enable_qa').checked));
        submit.disabled = true;
        byId('downloads').replaceChildren();
        byId('warnings').replaceChildren();
        byId('qa-status').textContent = '';
        byId('progress').value = 0;
        showMessage('正在上传原件…');
        try {
            const response = await request('/task/layout-overlay', {method: 'POST', body: data});
            taskId = response.task_id;
            sessionStorage.setItem('layout-overlay-task', taskId);
            poll(version);
        } catch (error) {
            submit.disabled = false;
            showMessage(error.message, true);
        }
    });

    async function initialize() {
        try {
            config = await request('/task/layout-overlay/config');
            populate('source_lang', config.languages, 'zh');
            populate('target_lang', config.languages, 'en');
            populate('ocr_provider', config.ocr_providers, config.default_ocr_provider);
            populate('translation_engine', config.translation_engines, config.default_translation_engine);
            populate('vision_model', config.models, config.default_model);
            populate('gemini_route', config.routes, config.default_route);
            byId('file').accept = config.allowed_extensions.join(',');
            byId('limits').textContent = `单个文件最多 ${config.upload_max_mb} MB、${config.max_pages} 页。`;
            submit.disabled = false;
            taskId = sessionStorage.getItem('layout-overlay-task');
            if (taskId) {
                submit.disabled = true;
                poll(++generation);
            }
        } catch (error) {
            showMessage(`配置加载失败：${error.message}`, true);
        }
    }
    initialize();
})();
