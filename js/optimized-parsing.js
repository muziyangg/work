// 1. 引入Web Worker处理解析任务（避免阻塞主线程）
// 新建解析工作器文件 parser.worker.js
// 注意：Web Worker无法直接操作DOM，只能处理数据解析

// 主文件中创建工作器
function createParserWorker() {
    // 检查浏览器是否支持Web Worker
    if (!window.Worker) {
        console.warn('当前浏览器不支持Web Worker，将使用主线程解析');
        return null;
    }
    
    try {
        const worker = new Worker('/js/parser.worker.js');
        
        // 监听工作器消息
        worker.onmessage = (e) => {
            const { type, data, error } = e.data;
            
            if (error) {
                handleParsingError(error);
                return;
            }
            
            switch(type) {
                case 'xlsx-result':
                case 'pptx-result':
                    // 处理解析结果（在主线程更新DOM）
                    renderParsedContent(data.html);
                    break;
                case 'progress':
                    // 更新进度条
                    updateProgressIndicator(data.progress);
                    break;
            }
        };
        
        worker.onerror = (error) => {
            console.error('Worker error:', error);
            handleParsingError(`解析失败: ${error.message}`);
            worker.terminate();
        };
        
        return worker;
    } catch (err) {
        console.error('创建Worker失败:', err);
        return null;
    }
}

// 2. 分块处理和进度反馈
async function parseLargeFile(arrayBuffer, docType, fileSize) {
    const previewContainer = document.querySelector('#document-modal #document-preview');
    previewContainer.innerHTML = `
        <div class="parsing-progress">
            <div class="progress-bar" style="width: 0%"></div>
            <p class="progress-text">准备解析...</p>
        </div>
    `;
    
    // 创建工作器
    const worker = createParserWorker();
    
    if (worker) {
        // 使用Web Worker解析
        worker.postMessage({
            type: `parse-${docType}`,
            arrayBuffer,
            fileSize
        }, [arrayBuffer]); // 转移ArrayBuffer所有权以节省内存
    } else {
        // 降级到主线程解析，但使用分块处理
        try {
            let result;
            switch(docType) {
                case 'xlsx':
                    result = await parseXlsxInChunks(arrayBuffer, (progress) => {
                        updateProgressIndicator(progress);
                    });
                    break;
                case 'pptx':
                    result = await parsePptxInChunks(arrayBuffer, (progress) => {
                        updateProgressIndicator(progress);
                    });
                    break;
                default:
                    throw new Error(`不支持的文档类型: ${docType}`);
            }
            renderParsedContent(result.value);
        } catch (error) {
            handleParsingError(error.message);
        }
    }
}

// 3. 分块解析XLSX
async function parseXlsxInChunks(arrayBuffer, progressCallback) {
    try {
        progressCallback(10); // 开始解析
        
        // 先读取工作簿信息（不加载完整数据）
        const workbook = XLSX.read(arrayBuffer, { 
            type: 'array',
            cellStyles: false, // 禁用样式解析以提高速度
            cellHTML: false
        });
        
        progressCallback(20); // 工作簿信息读取完成
        
        let htmlContent = '<div class="xlsx-content">';
        htmlContent += `<h3>Excel文档: ${workbook.SheetNames.length} 个工作表</h3>`;
        
        const totalSheets = workbook.SheetNames.length;
        
        // 逐个解析工作表，使用requestIdleCallback分散负载
        for (let i = 0; i < totalSheets; i++) {
            const sheetName = workbook.SheetNames[i];
            
            // 计算当前进度
            const progress = 20 + Math.round((i / totalSheets) * 60);
            progressCallback(progress);
            
            // 等待空闲时间再处理，避免阻塞UI
            await new Promise(resolve => {
                requestIdleCallback(resolve, { timeout: 100 });
            });
            
            // 解析单个工作表
            const worksheet = workbook.Sheets[sheetName];
            const html = XLSX.utils.sheet_to_html(worksheet, {
                raw: true, // 不转换格式，提高速度
                header: false
            });
            
            htmlContent += `
                <div class="worksheet-container">
                    <h4>工作表 ${i + 1}: ${sheetName}</h4>
                    <div class="table-wrapper">${html}</div>
                </div>
                <hr>
            `;
        }
        
        htmlContent += '</div>';
        progressCallback(100); // 解析完成
        
        return { value: htmlContent };
    } catch (error) {
        console.error('XLSX解析错误:', error);
        throw error;
    }
}

// 4. 虚拟滚动处理大量内容
function renderParsedContent(html) {
    const previewContainer = document.querySelector('#document-modal #document-preview');
    
    // 检查内容大小，如果过大则使用虚拟滚动
    if (html.length > 500000) { // 500KB以上的内容
        enableVirtualScrolling(previewContainer, html);
    } else {
        // 普通渲染
        previewContainer.innerHTML = html;
        // 延迟加载图片（如果有）
        lazyLoadImages(previewContainer);
    }
}

// 虚拟滚动实现
function enableVirtualScrolling(container, html) {
    // 创建虚拟滚动容器
    container.innerHTML = `
        <div class="virtual-scroller">
            <div class="scroll-container"></div>
            <div class="virtual-placeholder"></div>
        </div>
    `;
    
    const scroller = container.querySelector('.virtual-scroller');
    const scrollContainer = container.querySelector('.scroll-container');
    const placeholder = container.querySelector('.virtual-placeholder');
    
    // 创建临时元素计算总高度
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;
    tempDiv.style.position = 'absolute';
    tempDiv.style.visibility = 'hidden';
    document.body.appendChild(tempDiv);
    
    // 设置占位符高度以启用滚动
    const totalHeight = tempDiv.offsetHeight;
    placeholder.style.height = `${totalHeight}px`;
    document.body.removeChild(tempDiv);
    
    // 可视区域高度
    const viewportHeight = container.clientHeight;
    // 每次渲染的缓冲区大小
    const buffer = 200;
    let currentStart = 0;
    
    // 渲染可见区域内容
    function renderVisibleContent() {
        const scrollTop = scroller.scrollTop;
        
        // 计算需要显示的内容范围
        const start = Math.max(0, Math.floor((scrollTop - buffer) / 10));
        const end = Math.ceil((scrollTop + viewportHeight + buffer) / 10);
        
        // 只在范围变化时重新渲染
        if (start !== currentStart) {
            currentStart = start;
            
            // 创建可见区域内容片段
            const fragment = document.createDocumentFragment();
            
            // 克隆可见部分的元素
            const allElements = tempDiv.children;
            const startIndex = Math.min(start, allElements.length);
            const endIndex = Math.min(end, allElements.length);
            
            for (let i = startIndex; i < endIndex; i++) {
                fragment.appendChild(allElements[i].cloneNode(true));
            }
            
            // 更新滚动容器
            scrollContainer.innerHTML = '';
            scrollContainer.appendChild(fragment);
            scrollContainer.style.transform = `translateY(${start * 10}px)`;
            
            // 延迟加载当前可见区域的图片
            lazyLoadImages(scrollContainer);
        }
        
        // 继续下一帧渲染
        requestAnimationFrame(renderVisibleContent);
    }
    
    // 开始渲染循环
    renderVisibleContent();
}

// 5. 图片延迟加载
function lazyLoadImages(container) {
    const images = container.querySelectorAll('img[data-src]');
    if (!images.length) return;
    
    // 使用IntersectionObserver延迟加载图片
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const img = entry.target;
                img.src = img.dataset.src;
                img.removeAttribute('data-src');
                observer.unobserve(img);
            }
        });
    }, {
        rootMargin: '200px 0px'
    });
    
    images.forEach(img => observer.observe(img));
}

// 6. 进度指示器更新
function updateProgressIndicator(percent) {
    const progressBar = document.querySelector('.progress-bar');
    const progressText = document.querySelector('.progress-text');
    
    if (progressBar && progressText) {
        progressBar.style.width = `${percent}%`;
        progressText.textContent = `正在解析: ${percent}%`;
    }
}

// 7. 错误处理
function handleParsingError(message) {
    const previewContainer = document.querySelector('#document-modal #document-preview');
    previewContainer.innerHTML = `<div class="error-message">
        <p>解析文档时发生错误</p>
        <p>错误: ${message}</p>
        <p>建议尝试下载文档查看</p>
    </div>`;
}

// 添加解析进度和虚拟滚动的CSS样式
const style = document.createElement('style');
style.textContent = `
/* 进度指示器样式 */
.parsing-progress {
    padding: 20px;
}

.progress-bar {
    height: 6px;
    background-color: #007bff;
    border-radius: 3px;
    transition: width 0.3s ease;
}

.progress-text {
    margin: 10px 0 0;
    color: #666;
    font-size: 0.9rem;
}

/* 虚拟滚动样式 */
.virtual-scroller {
    position: relative;
    height: 100%;
    overflow-y: auto;
}

.scroll-container {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
}

.virtual-placeholder {
    opacity: 0;
}

/* 限制表格渲染大小，避免过大表格导致的性能问题 */
.worksheet-container {
    max-width: 100%;
    overflow: hidden;
}

.table-wrapper {
    max-height: 600px; /* 限制表格高度 */
    overflow: auto;
}

/* 大型文档优化 */
.large-document .slide,
.large-document .worksheet-container {
    page-break-after: auto;
    -webkit-print-color-adjust: exact;
}
`;
document.head.appendChild(style);
