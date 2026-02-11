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

const fromLang = document.getElementById('fromLang');
const toLang = document.getElementById('toLang');
const enableCorrection = document.getElementById('enableCorrection');
const enableVisualization = document.getElementById('enableVisualization');

const processingStatus = document.getElementById('processingStatus');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultStats = document.getElementById('resultStats');
const resultGrid = document.getElementById('resultGrid');

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
    const validTypes = ['image/jpeg', 'image/jpg', 'image/png', 'image/bmp', 'application/pdf'];
    if (!validTypes.includes(file.type)) {
        alert('不支持的文件类型！请上传 JPG、PNG 或 PDF 文件。');
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
    if (file.type.startsWith('image/')) {
        const reader = new FileReader();
        reader.onload = (e) => {
            previewImage.src = e.target.result;
        };
        reader.readAsDataURL(file);
    } else {
        previewImage.src = 'https://via.placeholder.com/120?text=PDF';
    }
    
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
        
        // 构建URL参数
        const params = new URLSearchParams({
            from_lang: fromLang.value,
            to_lang: toLang.value,
            enable_correction: enableCorrection.checked,
            enable_visualization: enableVisualization.checked
        });
        
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
    
    if (percent < 30) {
        processingStatus.textContent = '正在读取文件...';
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
    const stats = `
        <div class="stat-card">
            <i class="fas fa-file-alt"></i>
            <h3>${data.filename}</h3>
            <p>文件名</p>
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

