// 导入所需的库（注意：Web Worker 中通过 importScripts 加载外部库）
importScripts(
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/pptx-parser@0.3.1/dist/pptx-parser.min.js'
);

// 监听主线程发送的消息
self.onmessage = async (e) => {
    const { type, arrayBuffer, fileSize } = e.data;
    
    try {
        // 根据文档类型执行不同的解析操作
        switch(type) {
            case 'parse-xlsx':
                await parseXlsx(arrayBuffer, fileSize);
                break;
            case 'parse-pptx':
                await parsePptx(arrayBuffer, fileSize);
                break;
            default:
                throw new Error(`不支持的解析类型: ${type}`);
        }
    } catch (error) {
        // 向主线程发送错误信息
        self.postMessage({
            type: 'error',
            error: error.message || '解析过程中发生未知错误'
        });
    }
};

// 解析 XLSX 文件
async function parseXlsx(arrayBuffer, fileSize) {
    try {
        // 发送进度更新：10%
        self.postMessage({
            type: 'progress',
            data: { progress: 10 }
        });
        
        // 解析工作簿（禁用样式以提高性能）
        const workbook = XLSX.read(arrayBuffer, {
            type: 'array',
            cellStyles: false,
            cellHTML: false
        });
        
        // 发送进度更新：20%
        self.postMessage({
            type: 'progress',
            data: { progress: 20 }
        });
        
        let htmlContent = '<div class="xlsx-content">';
        htmlContent += `<h3>Excel文档: ${workbook.SheetNames.length} 个工作表</h3>`;
        
        const totalSheets = workbook.SheetNames.length;
        
        // 逐个解析工作表
        for (let i = 0; i < totalSheets; i++) {
            const sheetName = workbook.SheetNames[i];
            
            // 计算当前进度（20% ~ 80%）
            const progress = 20 + Math.round((i / totalSheets) * 60);
            self.postMessage({
                type: 'progress',
                data: { progress }
            });
            
            // 解析单个工作表为 HTML
            const worksheet = workbook.Sheets[sheetName];
            const html = XLSX.utils.sheet_to_html(worksheet, {
                raw: true,  // 不转换格式，提高速度
                header: false
            });
            
            htmlContent += `
                <div class="worksheet-container">
                    <h4>工作表 ${i + 1}: ${sheetName}</h4>
                    <div class="table-wrapper">${html}</div>
                </div>
                <hr>
            `;
            
            // 每处理一个工作表，短暂休眠以避免Worker过度占用CPU
            await new Promise(resolve => setTimeout(resolve, 10));
        }
        
        htmlContent += '</div>';
        
        // 发送完成信号：100%
        self.postMessage({
            type: 'progress',
            data: { progress: 100 }
        });
        
        // 发送解析结果
        self.postMessage({
            type: 'xlsx-result',
            data: { html: htmlContent }
        });
        
    } catch (error) {
        console.error('XLSX解析错误:', error);
        throw new Error(`Excel解析失败: ${error.message}`);
    }
}

// 解析 PPTX 文件
async function parsePptx(arrayBuffer, fileSize) {
    try {
        // 发送进度更新：10%
        self.postMessage({
            type: 'progress',
            data: { progress: 10 }
        });
        
        // 将 ArrayBuffer 转换为 Uint8Array
        const uint8Array = new Uint8Array(arrayBuffer);
        
        // 发送进度更新：30%
        self.postMessage({
            type: 'progress',
            data: { progress: 30 }
        });
        
        // 解析PPTX文件
        const presentation = await PptxParser.parse(uint8Array);
        
        // 发送进度更新：60%
        self.postMessage({
            type: 'progress',
            data: { progress: 60 }
        });
        
        let htmlContent = '<div class="pptx-content">';
        htmlContent += `<h3>PowerPoint文档: ${presentation.slides.length} 张幻灯片</h3>`;
        
        // 处理幻灯片内容
        presentation.slides.forEach((slide, index) => {
            htmlContent += `<div class="slide slide-${index + 1}">`;
            htmlContent += `<h4>幻灯片 ${index + 1}</h4>`;
            
            if (slide.text && slide.text.length > 0) {
                htmlContent += '<div class="slide-content">';
                
                // 处理文本内容（保留层级结构）
                slide.text.forEach(text => {
                    const tag = text.level === 0 ? 'h5' : 
                               text.level === 1 ? 'h6' : 'p';
                    htmlContent += `<${tag} class="slide-text level-${text.level}">${escapeHtml(text.content)}</${tag}>`;
                });
                
                htmlContent += '</div>';
            } else {
                htmlContent += '<p class="no-content">此幻灯片没有可提取的文本内容</p>';
            }
            
            htmlContent += '</div><hr>';
        });
        
        htmlContent += '</div>';
        
        // 发送完成信号：100%
        self.postMessage({
            type: 'progress',
            data: { progress: 100 }
        });
        
        // 发送解析结果
        self.postMessage({
            type: 'pptx-result',
            data: { html: htmlContent }
        });
        
    } catch (error) {
        console.error('PPTX解析错误:', error);
        throw new Error(`PowerPoint解析失败: ${error.message}`);
    }
}

// 辅助函数：转义HTML特殊字符
function escapeHtml(unsafe) {
    if (!unsafe) return '';
    return unsafe
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
    