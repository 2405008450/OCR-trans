(() => {
    'use strict';
    const $ = (id) => document.getElementById(id);
    const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
    const bytes = (size) => `${(Number(size || 0) / 1048576).toFixed(2)} MB`;
    let config = null;
    let candidates = [];
    let uploads = [];
    let scanned = false;
    let scanning = false;
    let scanVersion = 0;
    let active = false;
    let taskId = null;
    let timer = null;
    const source = () => document.querySelector('input[name="inputSource"]:checked').value;
    const items = () => source() === 'path' ? candidates : uploads;
    const chosen = () => items().filter((item) => item.selected && item.selectable !== false);

    function notice(id, message, error = false) {
        $(id).textContent = message;
        $(id).hidden = !message;
        $(id).classList.toggle('error', error);
    }

    async function request(url, options = {}) {
        const response = await fetch(url, options);
        let payload;
        try { payload = await response.json(); } catch (_) { throw new Error(`服务响应异常（${response.status}）`); }
        if (!response.ok) throw new Error(typeof payload.detail === 'string' ? payload.detail : `请求失败（${response.status}）`);
        return payload;
    }

    function refresh() {
        const selected = chosen();
        const total = selected.reduce((sum, item) => sum + item.size, 0);
        const overLimit = config && total > config.max_total_mb * 1048576;
        $('selectionSummary').textContent = `已选择 ${selected.length} 个文件，合计 ${bytes(total)}${overLimit ? '，超过整批大小限制，请减少选择。' : ''}`;
        $('selectionSummary').classList.toggle('failure', Boolean(overLimit));
        $('submitBtn').disabled = active || scanning || !config?.ffmpeg_available || !selected.length ||
            selected.length > config.max_files || total > config.max_total_mb * 1048576 ||
            (source() === 'path' && !scanned);
        $('scanFilesBtn').disabled = scanning || active;
        $('inputControls').disabled = active || !config;
        $('submitBtn').textContent = active ? '正在处理…' : '开始转换';
    }

    function render() {
        $('candidateFileList').innerHTML = items().map((item, index) => `<label class="candidate-row">
            <input type="checkbox" data-index="${index}" ${item.selected ? 'checked' : ''} ${item.selectable === false ? 'disabled' : ''}>
            <span class="candidate-copy">${esc(item.relative_path || item.name)}<small>${bytes(item.size)}${item.reason ? ` · ${esc(item.reason)}` : ''}</small></span>
        </label>`).join('');
        refresh();
    }

    function invalidateScan() {
        scanVersion += 1;
        scanned = false;
        candidates = [];
        if (source() === 'path') { notice('selectionMessage', '请扫描路径并勾选需要转换的文件。'); render(); }
    }

    async function scan() {
        const path = $('directoryPath').value.trim();
        if (!path) { notice('selectionMessage', '请填写共享目录或文件路径。', true); return; }
        const version = ++scanVersion;
        scanning = true;
        scanned = false;
        candidates = [];
        render();
        notice('selectionMessage', '正在扫描共享路径…');
        try {
            const payload = await request('/task/audio-convert/discover', {
                method: 'POST', headers: {'Content-Type':'application/json'},
                body: JSON.stringify({directory_path: path, recursive: $('recursiveInput').checked}),
            });
            if (version !== scanVersion) return;
            candidates = (payload.files || []).map((item) => ({...item, selected: item.selectable !== false}));
            scanned = true;
            notice('selectionMessage', `找到 ${candidates.length} 个文件，可选 ${payload.selectable_count} 个。${payload.truncated ? `扫描结果已按 ${payload.max_files} 个上限截断，请缩小目录范围以处理其余文件。` : ''}`);
        } catch (error) {
            if (version === scanVersion) notice('selectionMessage', error.message, true);
        } finally { scanning = false; render(); }
    }

    function addUploads(files) {
        if (active || !config) return;
        const next = [...uploads];
        for (const file of Array.from(files || [])) {
            if (!file.name.toLowerCase().endsWith('.m4a')) { notice('selectionMessage', `仅支持 M4A 文件：${file.name}`, true); return; }
            if (!file.size || file.size > config.max_file_mb * 1048576) { notice('selectionMessage', `文件为空或超过 ${config.max_file_mb} MB：${file.name}`, true); return; }
            next.push({file, name:file.name, size:file.size, selected:true, selectable:true});
        }
        if (next.length > config.max_files || next.reduce((sum, item) => sum + item.size, 0) > config.max_total_mb * 1048576) {
            notice('selectionMessage', `每批最多 ${config.max_files} 个文件，总大小不能超过 ${config.max_total_mb} MB。`, true); return;
        }
        uploads = next;
        notice('selectionMessage', '已添加文件，可取消勾选不需要转换的项目。');
        render();
    }

    function download(path, label) {
        const query = new URLSearchParams({file_path:path, download_name:path.replaceAll('\\','/').split('/').pop()});
        return `<a class="download-link" href="/task/${encodeURIComponent(taskId)}/download?${esc(query.toString())}">${esc(label)}</a>`;
    }

    async function poll() {
        clearTimeout(timer);
        $('retryStatus').hidden = true;
        try {
            const task = await request(`/task/audio-convert/status/${encodeURIComponent(taskId)}`);
            const progress = Math.max(0, Math.min(100, Number(task.progress || 0)));
            $('progressFill').style.width = `${progress}%`;
            $('progressFill').parentElement.setAttribute('aria-valuenow', String(progress));
            const terminal = ['done','failed','cancelled'].includes(task.status);
            const label = {queued:'排队中', processing:'转换中', running:'转换中', done:'转换完成', failed:'转换失败', cancelled:'已取消'}[task.status] || '处理中';
            $('taskTitle').textContent = label;
            $('taskMessage').textContent = task.result?.summary_text || (task.status === 'queued' ? `排队中${task.queue_position ? `，当前位置 ${task.queue_position}` : ''}` : task.message || label);
            notice('taskError', task.error || '', true);
            $('archiveDownload').innerHTML = task.result?.archive_zip ? download(task.result.archive_zip, '下载全部 WAV（ZIP）') : '';
            $('resultList').innerHTML = (task.result?.files || []).map((item) => `<div class="result-row"><div>${esc(item.filename)}${item.error ? `<div class="failure">${esc(item.error)}</div>` : `<small class="muted"> · ${bytes(item.size)}</small>`}</div>${item.output_path ? download(item.output_path,'下载 WAV') : '<span class="failure">失败</span>'}</div>`).join('');
            if (terminal) { active = false; refresh(); }
            else timer = setTimeout(poll, 1800);
        } catch (error) {
            notice('taskError', `查询进度失败：${error.message}。任务仍可能在服务器处理中，可重新查询或打开工作台。`, true);
            $('retryStatus').hidden = false;
        }
    }

    async function submit() {
        refresh();
        if ($('submitBtn').disabled) return;
        const selected = chosen();
        const usePath = source() === 'path';
        let options;
        if (usePath) {
            options = {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({
                directory_path:$('directoryPath').value.trim(), recursive:$('recursiveInput').checked,
                relative_paths:selected.map((item) => item.relative_path),
            })};
        } else {
            const body = new FormData();
            selected.forEach((item) => body.append('files', item.file, item.name));
            options = {method:'POST', body};
        }
        active = true;
        refresh();
        notice('selectionMessage', '正在提交转换任务…');
        try {
            const payload = await request(usePath ? '/task/audio-convert' : '/task/audio-convert/upload', options);
            taskId = payload.task_id;
            const url = new URL(window.location.href);
            url.searchParams.set('task_id', taskId);
            history.replaceState(null, '', url);
            $('taskPanel').hidden = false;
            $('taskIdText').textContent = `任务 ${taskId}`;
            notice('selectionMessage', payload.deduped ? '已复用正在处理的相同任务。' : '任务已提交。');
            $('taskPanel').scrollIntoView({behavior:'smooth', block:'start'});
            await poll();
        } catch (error) { active = false; notice('selectionMessage', error.message, true); refresh(); }
    }

    $('directoryPath').addEventListener('input', invalidateScan);
    $('recursiveInput').addEventListener('change', invalidateScan);
    document.querySelectorAll('input[name="inputSource"]').forEach((input) => input.addEventListener('change', () => {
        $('pathPanel').hidden = source() !== 'path';
        $('uploadPanel').hidden = source() !== 'upload';
        notice('selectionMessage', ''); render();
    }));
    $('scanFilesBtn').addEventListener('click', scan);
    $('submitBtn').addEventListener('click', submit);
    $('retryStatus').addEventListener('click', poll);
    $('candidateFileList').addEventListener('change', (event) => {
        const index = event.target.dataset.index;
        if (index !== undefined) { items()[Number(index)].selected = event.target.checked; refresh(); }
    });
    $('selectAll').addEventListener('click', () => { items().forEach((item) => { item.selected = item.selectable !== false; }); render(); });
    $('clearSelection').addEventListener('click', () => {
        if (source() === 'upload') { uploads = []; $('fileInput').value = ''; }
        else candidates.forEach((item) => { item.selected = false; });
        render();
    });
    $('dropZone').addEventListener('click', () => { if (!active) $('fileInput').click(); });
    $('dropZone').addEventListener('keydown', (event) => { if (['Enter',' '].includes(event.key)) { event.preventDefault(); if (!active) $('fileInput').click(); } });
    $('fileInput').addEventListener('change', (event) => { addUploads(event.target.files); event.target.value = ''; });
    ['dragenter','dragover'].forEach((name) => $('dropZone').addEventListener(name, (event) => { event.preventDefault(); $('dropZone').classList.add('dragging'); }));
    ['dragleave','drop'].forEach((name) => $('dropZone').addEventListener(name, (event) => { event.preventDefault(); $('dropZone').classList.remove('dragging'); }));
    $('dropZone').addEventListener('drop', (event) => addUploads(event.dataTransfer.files));

    async function init() {
        try {
            config = await request('/task/audio-convert/config');
            notice('configNotice', config.ffmpeg_available ? '转换服务已就绪 · PCM 16 位 WAV' : config.unavailable_reason, !config.ffmpeg_available);
            $('limitHint').textContent = `每批最多 ${config.max_files} 个文件；单文件不超过 ${config.max_file_mb} MB，整批不超过 ${config.max_total_mb} MB。`;
            $('allowedRoots').textContent = (config.allowed_roots || []).map((item) => item.mount_path ? `${item.path} → ${item.mount_path}` : item.path).join('\n') || '请使用服务器已配置的共享路径。';
        } catch (error) { notice('configNotice', `读取配置失败：${error.message}，请刷新页面重试。`, true); }
        refresh();
        const saved = new URLSearchParams(window.location.search).get('task_id');
        if (saved) { taskId = saved; active = true; $('taskPanel').hidden = false; $('taskIdText').textContent = `任务 ${taskId}`; refresh(); await poll(); }
    }
    init();
})();
