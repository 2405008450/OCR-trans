// ========== 全局变量 ==========
let selectedFile = null;

// ========== DOM 元素 ==========
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const filePreview = document.getElementById('filePreview');
const previewImage = document.getElementById('previewImage');
const fileName = document.getElementById('fileName');
const btnRemove = document.getElementById('btnRemove');
const btnProcess = document.getElementById('btnProcess');
const btnNewTask = document.getElementById('btnNewTask');

const uploadSection = document.getElementById('uploadSection');
const processingSection = document.getElementById('processingSection');
const resultSection = document.getElementById('resultSection');

const docType = document.getElementById('docType');
const fromLang = document.getElementById('fromLang');
const toLang = document.getElementById('toLang');
const cardSide = document.getElementById('cardSide');
const enableVisualization = document.getElementById('enableVisualization');

// 结婚证专用选项
const idCardOptions = document.getElementById('idCardOptions');
const marriageCertOptions = document.getElementById('marriageCertOptions');
const enableMerge = document.getElementById('enableMerge');
const enableOverlapFix = document.getElementById('enableOverlapFix');
const enableColonFix = document.getElementById('enableColonFix');
const marriagePageTemplate = document.getElementById('marriagePageTemplate');
const fontSizeInput = document.getElementById('fontSize');

const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

function checkedOrDefault(el, defaultValue = false) {
    return el ? !!el.checked : defaultValue;
}

function valueOrDefault(el, defaultValue = '') {
    return el ? el.value : defaultValue;
}

function applyMarriageTemplate() {
    const template = valueOrDefault(marriagePageTemplate, 'page2');

    // 第一页(封面): merge=true, overlap=true, colon=false, 置信度0.8
    if (template === 'page1') {
        if (enableMerge) enableMerge.checked = true;
        if (enableOverlapFix) enableOverlapFix.checked = true;
        if (enableColonFix) enableColonFix.checked = false;
        return;
    }

    // 第二页: merge=false, overlap=true, colon=true
    if (template === 'page2') {
        if (enableMerge) enableMerge.checked = false;
        if (enableOverlapFix) enableOverlapFix.checked = true;
        if (enableColonFix) enableColonFix.checked = true;
        return;
    }

    // 第三页: merge=true, overlap=true, colon=false
    if (template === 'page3') {
        if (enableMerge) enableMerge.checked = true;
        if (enableOverlapFix) enableOverlapFix.checked = true;
        if (enableColonFix) enableColonFix.checked = false;
    }
}

// ========== 事件监听 ==========
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);
btnRemove.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFile();
});
btnProcess.addEventListener('click', processFile);
btnNewTask.addEventListener('click', resetApp);

// 证件类型切换：显示/隐藏对应选项
if (docType) {
    docType.addEventListener('change', handleDocTypeChange);
    handleDocTypeChange();
}
if (marriagePageTemplate) {
    marriagePageTemplate.addEventListener('change', applyMarriageTemplate);
}

function handleDocTypeChange() {
    const type = valueOrDefault(docType, 'id_card');
    if (!idCardOptions || !marriageCertOptions) return;
    if (type === 'marriage_cert') {
        idCardOptions.style.display = 'none';
        marriageCertOptions.style.display = 'block';
        applyMarriageTemplate();
    } else {
        idCardOptions.style.display = 'block';
        marriageCertOptions.style.display = 'none';
    }
}

// ========== 文件处理函数 ==========
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.style.borderColor = 'var(--primary-color)';
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.style.borderColor = 'var(--border-color)';
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFile(file) {
    // 验证文件类型
    const validTypes = ['image/jpeg', 'image/jpg', 'image/png', 'image/bmp', 'image/tiff'];
    if (!validTypes.includes(file.type) && !file.type.startsWith('image/')) {
        alert('不支持的文件类型！请上传图片文件。');
        return;
    }
    
    // 验证文件大小（50MB）
    if (file.size > 50 * 1024 * 1024) {
        alert('文件太大！请上传小于 50MB 的文件。');
        return;
    }
    
    selectedFile = file;
    fileName.textContent = file.name;
    
    // 显示预览
    const reader = new FileReader();
    reader.onload = (e) => {
        previewImage.src = e.target.result;
    };
    reader.readAsDataURL(file);
    
    uploadPlaceholder.style.display = 'none';
    filePreview.style.display = 'flex';
    btnProcess.disabled = false;
}

function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    uploadPlaceholder.style.display = 'block';
    filePreview.style.display = 'none';
    btnProcess.disabled = true;
}

// ========== 处理文件 ==========
async function processFile() {
    if (!selectedFile) return;
    
    // 切换到处理界面
    uploadSection.style.display = 'none';
    processingSection.style.display = 'block';
    
    // 模拟进度
    let progress = 0;
    const progressInterval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        updateProgress(progress);
    }, 500);
    
    try {
        // 构建表单数据
        const formData = new FormData();
        formData.append('file', selectedFile);
        
        // 构建URL参数（通用参数）
        const params = new URLSearchParams({
            from_lang: valueOrDefault(fromLang, 'zh'),
            to_lang: valueOrDefault(toLang, 'en'),
            enable_visualization: checkedOrDefault(enableVisualization, true),
            doc_type: valueOrDefault(docType, 'id_card'),
        });
        
        // 根据证件类型添加特定参数
        if (valueOrDefault(docType, 'id_card') === 'id_card') {
            params.append('card_side', valueOrDefault(cardSide, 'front'));
        } else if (valueOrDefault(docType, 'id_card') === 'marriage_cert') {
            params.append('marriage_page_template', valueOrDefault(marriagePageTemplate, 'page2'));
            params.append('enable_merge', checkedOrDefault(enableMerge, true));
            params.append('enable_overlap_fix', checkedOrDefault(enableOverlapFix, true));
            params.append('enable_colon_fix', checkedOrDefault(enableColonFix, false));
            const fs = parseInt(valueOrDefault(fontSizeInput, '18'));
            if (fs && fs >= 8 && fs <= 30) {
                params.append('font_size', fs);
            }
        }
        
        // 发送请求
        const response = await fetch(`/task/run?${params}`, {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`请求失败: ${response.status}`);
        }
        
        const result = await response.json();
        
        // 完成进度
        clearInterval(progressInterval);
        updateProgress(100);
        
        // 延迟显示结果
        setTimeout(() => {
            displayResult(result);
        }, 500);
        
    } catch (error) {
        clearInterval(progressInterval);
        console.error('处理失败:', error);
        alert(`处理失败: ${error.message}`);
        resetApp();
    }
}

function updateProgress(percent) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;
    
    const currentDocType = valueOrDefault(docType, 'id_card');
    const docLabel = currentDocType === 'marriage_cert' ? '结婚证' : '身份证';
    
    if (percent < 30) {
        processingStatus.textContent = `正在读取${docLabel}文件...`;
    } else if (percent < 60) {
        processingStatus.textContent = '正在进行OCR识别...';
    } else if (percent < 90) {
        processingStatus.textContent = '正在翻译文本...';
    } else {
        processingStatus.textContent = '正在生成结果...';
    }
}

// ========== 显示结果 ==========
function displayResult(data) {
    processingSection.style.display = 'none';
    resultSection.style.display = 'block';
    
    // 显示统计信息
    const currentDocType = valueOrDefault(docType, 'id_card');
    const docLabel = currentDocType === 'marriage_cert' ? '结婚证' : '身份证';
    
    const stats = `
        <div class="stat-card">
            <i class="fas fa-file-alt"></i>
            <h3>${data.filename}</h3>
            <p>文件名</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-stamp"></i>
            <h3>${docLabel}</h3>
            <p>证件类型</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-images"></i>
            <h3>${data.total_images}</h3>
            <p>处理图片数</p>
        </div>
        <div class="stat-card">
            <i class="fas fa-check-circle"></i>
            <h3>成功</h3>
            <p>处理状态</p>
        </div>
    `;
    resultStats.innerHTML = stats;
    
    // 显示每张图片的结果
    let gridHtml = '';
    data.results.forEach((item, index) => {
        gridHtml += `
            <div class="result-item">
                <h3>图片 ${index + 1}</h3>
                <div class="image-comparison">
                    ${item.corrected_image ? `
                        <div class="image-box">
                            <h4>矫正后的图片</h4>
                            <img src="/${item.corrected_image}" alt="矫正后" onclick="window.open('/${item.corrected_image}', '_blank')">
                        </div>
                    ` : ''}
                    ${item.visualization_image ? `
                        <div class="image-box">
                            <h4>OCR识别可视化</h4>
                            <img src="/${item.visualization_image}" alt="可视化" onclick="window.open('/${item.visualization_image}', '_blank')">
                        </div>
                    ` : ''}
                    <div class="image-box">
                        <h4>翻译后的图片</h4>
                        <img src="/${item.translated_image}" alt="翻译后" onclick="window.open('/${item.translated_image}', '_blank')">
                    </div>
                </div>
                <div class="download-links">
                    <a href="/${item.translated_image}" download class="download-btn">
                        <i class="fas fa-download"></i> 下载翻译图片
                    </a>
                    ${item.ocr_json ? `
                        <a href="/${item.ocr_json}" download class="download-btn">
                            <i class="fas fa-file-code"></i> 下载OCR数据
                        </a>
                    ` : ''}
                    ${item.translated_json ? `
                        <a href="/${item.translated_json}" download class="download-btn">
                            <i class="fas fa-language"></i> 下载翻译数据
                        </a>
                    ` : ''}
                </div>
            </div>
        `;
    });
    resultGrid.innerHTML = gridHtml;
}

// ========== 重置应用 ==========
function resetApp() {
    clearFile();
    uploadSection.style.display = 'block';
    processingSection.style.display = 'none';
    resultSection.style.display = 'none';
    progressFill.style.width = '0%';
    progressText.textContent = '0%';
}
