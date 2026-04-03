/**
 * 实验室正常值范围检查系统 - 前端逻辑
 */

// 全局变量
let currentFile = null;
let analysisResult = null;

// DOM 元素
const fileInput = document.getElementById('file-input');
const uploadArea = document.getElementById('upload-area');
const fileInfo = document.getElementById('file-info');
const filenameDisplay = document.getElementById('filename');
const analyzeBtn = document.getElementById('analyze-btn');
const loadingSection = document.getElementById('loading');
const statsSection = document.getElementById('stats-section');
const exportSection = document.getElementById('export-section');
const resultsSection = document.getElementById('results-section');
const errorsSection = document.getElementById('errors-section');

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    initEventListeners();
});

// 事件监听
function initEventListeners() {
    // 文件选择
    fileInput.addEventListener('change', handleFileSelect);
    
    // 拖拽上传
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    
    // 分析按钮
    analyzeBtn.addEventListener('click', analyzeFile);
    
    // 导出按钮
    document.getElementById('export-csv').addEventListener('click', exportCSV);
    document.getElementById('export-html').addEventListener('click', exportHTML);
    document.getElementById('export-docx').addEventListener('click', exportDOCX);
    
    // 过滤器按钮
    document.querySelectorAll('[data-filter]').forEach(btn => {
        btn.addEventListener('click', handleFilter);
    });
}

// 处理文件选择
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        setFile(file);
    }
}

// 处理拖拽悬停
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.add('dragover');
}

// 处理拖拽离开
function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.remove('dragover');
}

// 处理文件拖放
function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        if (file.name.endsWith('.docx')) {
            setFile(file);
        } else {
            showAlert('请选择 .docx 格式的Word文档', 'warning');
        }
    }
}

// 设置当前文件
function setFile(file) {
    currentFile = file;
    filenameDisplay.textContent = file.name;
    fileInfo.classList.remove('d-none');
    analyzeBtn.disabled = false;
}

// 清除文件
function clearFile() {
    currentFile = null;
    fileInput.value = '';
    fileInfo.classList.add('d-none');
    analyzeBtn.disabled = true;
    hideResults();
}

// 隐藏结果区域
function hideResults() {
    statsSection.classList.add('d-none');
    exportSection.classList.add('d-none');
    resultsSection.classList.add('d-none');
    errorsSection.classList.add('d-none');
}

// 分析文件
async function analyzeFile() {
    if (!currentFile) {
        showAlert('请先选择文件', 'warning');
        return;
    }
    
    // 显示加载动画
    loadingSection.classList.remove('d-none');
    hideResults();
    analyzeBtn.disabled = true;
    
    const formData = new FormData();
    formData.append('file', currentFile);
    
    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            analysisResult = result;
            displayResults(result);
            showAlert('分析完成！', 'success');
        } else {
            showAlert(result.error || '分析失败', 'danger');
        }
    } catch (error) {
        showAlert('网络错误，请稍后重试', 'danger');
        console.error('Error:', error);
    } finally {
        loadingSection.classList.add('d-none');
        analyzeBtn.disabled = false;
    }
}

// 显示结果
function displayResults(result) {
    const { stats, records, errors } = result;
    
    // 显示统计信息
    document.getElementById('stat-records').textContent = stats.total_records;
    document.getElementById('stat-tables').textContent = stats.total_tables;
    document.getElementById('stat-errors').textContent = stats.error_count;
    document.getElementById('stat-warnings').textContent = stats.warning_count;
    
    statsSection.classList.remove('d-none');
    statsSection.classList.add('fade-in');
    
    // 显示导出按钮
    exportSection.classList.remove('d-none');
    exportSection.classList.add('fade-in');
    
    // 显示结果表格
    displayResultsTable(records, errors);
    resultsSection.classList.remove('d-none');
    resultsSection.classList.add('fade-in');
    
    // 显示错误详情
    if (errors.length > 0) {
        displayErrors(errors);
        errorsSection.classList.remove('d-none');
        errorsSection.classList.add('fade-in');
    }
}

// 显示结果表格
function displayResultsTable(records, errors) {
    const tbody = document.getElementById('results-tbody');
    tbody.innerHTML = '';
    
    // 创建错误查找字典
    const errorDict = {};
    errors.forEach(error => {
        if (error.row) {
            if (!errorDict[error.row]) {
                errorDict[error.row] = [];
            }
            errorDict[error.row].push(error);
        }
    });
    
    records.forEach(record => {
        const rowErrors = errorDict[record.row_index] || [];
        const hasError = rowErrors.some(e => e.severity === 'error');
        const hasWarning = rowErrors.some(e => e.severity === 'warning');
        
        let rowClass = '';
        let statusBadge = '';
        
        if (hasError) {
            rowClass = 'table-row-error';
            statusBadge = '<span class="badge bg-danger">错误</span>';
        } else if (hasWarning) {
            rowClass = 'table-row-warning';
            statusBadge = '<span class="badge bg-warning text-dark">警告</span>';
        } else {
            rowClass = 'table-row-success';
            statusBadge = '<span class="badge bg-success">通过</span>';
        }
        
        const ageRange = record.age_min || record.age_max 
            ? `${record.age_min || '-'} - ${record.age_max || '-'}` 
            : '-';
        const refRange = record.ref_min || record.ref_max 
            ? `${record.ref_min || '-'} - ${record.ref_max || '-'}` 
            : '-';
        
        const errorMessages = rowErrors.map(e => e.message).join('; ');
        
        const tr = document.createElement('tr');
        tr.className = rowClass;
        tr.dataset.rowIndex = record.row_index;
        tr.dataset.hasError = hasError;
        tr.dataset.hasWarning = hasWarning;
        
        tr.innerHTML = `
            <td>${record.row_index}</td>
            <td>${escapeHtml(record.item)}</td>
            <td>${escapeHtml(record.gender)}</td>
            <td>${escapeHtml(ageRange)}</td>
            <td>${escapeHtml(refRange)}</td>
            <td>${escapeHtml(record.unit)}</td>
            <td>
                ${statusBadge}
                ${errorMessages ? `<small class="d-block text-muted mt-1">${escapeHtml(errorMessages)}</small>` : ''}
            </td>
        `;
        
        tbody.appendChild(tr);
    });
}

// 显示错误详情
function displayErrors(errors) {
    const errorsList = document.getElementById('errors-list');
    errorsList.innerHTML = '';
    
    errors.forEach(error => {
        const item = document.createElement('div');
        item.className = `error-item ${error.severity}`;
        
        const icon = error.severity === 'error' ? 'bi-x-circle-fill' : 'bi-exclamation-triangle-fill';
        const title = error.severity === 'error' ? '错误' : '警告';
        
        item.innerHTML = `
            <div class="d-flex align-items-start">
                <i class="bi ${icon} me-2 ${error.severity === 'error' ? 'text-danger' : 'text-warning'}"></i>
                <div class="flex-grow-1">
                    <div class="error-title ${error.severity === 'error' ? 'text-danger' : 'text-warning'}">
                        ${title}: ${escapeHtml(error.item || '未知项目')}
                    </div>
                    <p class="error-message">${escapeHtml(error.message)}</p>
                    ${error.row ? `<div class="error-meta">行号: ${error.row}</div>` : ''}
                </div>
            </div>
        `;
        
        errorsList.appendChild(item);
    });
}

// 处理过滤器
function handleFilter(e) {
    const filter = e.target.dataset.filter;
    
    // 更新按钮状态
    document.querySelectorAll('[data-filter]').forEach(btn => {
        btn.classList.remove('active');
    });
    e.target.classList.add('active');
    
    // 过滤表格行
    const rows = document.querySelectorAll('#results-tbody tr');
    rows.forEach(row => {
        const hasError = row.dataset.hasError === 'true';
        const hasWarning = row.dataset.hasWarning === 'true';
        
        let show = false;
        switch (filter) {
            case 'all':
                show = true;
                break;
            case 'error':
                show = hasError;
                break;
            case 'warning':
                show = hasWarning;
                break;
        }
        
        row.style.display = show ? '' : 'none';
    });
}

// 导出CSV
async function exportCSV() {
    if (!analysisResult) return;
    
    try {
        const response = await fetch('/api/export/csv', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                records: analysisResult.records,
                errors: analysisResult.errors
            })
        });
        
        if (response.ok) {
            const blob = await response.blob();
            downloadFile(blob, 'lab_check_result.csv');
        } else {
            showAlert('导出失败', 'danger');
        }
    } catch (error) {
        showAlert('导出失败', 'danger');
        console.error('Error:', error);
    }
}

// 导出HTML
async function exportHTML() {
    if (!analysisResult) return;
    
    try {
        const response = await fetch('/api/export/html', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                records: analysisResult.records,
                errors: analysisResult.errors,
                stats: analysisResult.stats
            })
        });
        
        if (response.ok) {
            const blob = await response.blob();
            downloadFile(blob, 'lab_check_report.html');
        } else {
            showAlert('导出失败', 'danger');
        }
    } catch (error) {
        showAlert('导出失败', 'danger');
        console.error('Error:', error);
    }
}

// 导出DOCX
async function exportDOCX() {
    if (!analysisResult) return;
    
    try {
        const response = await fetch('/api/export/docx', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                filepath: analysisResult.filepath,
                records: analysisResult.records,
                errors: analysisResult.errors
            })
        });
        
        if (response.ok) {
            const blob = await response.blob();
            downloadFile(blob, 'lab_check_marked.docx');
        } else {
            const result = await response.json();
            showAlert(result.error || '导出失败', 'danger');
        }
    } catch (error) {
        showAlert('导出失败', 'danger');
        console.error('Error:', error);
    }
}

// 下载文件
function downloadFile(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

// 显示提示
function showAlert(message, type) {
    // 移除现有的提示
    const existingAlert = document.querySelector('.alert-floating');
    if (existingAlert) {
        existingAlert.remove();
    }
    
    // 创建新提示
    const alert = document.createElement('div');
    alert.className = `alert alert-${type} alert-dismissible fade show alert-floating`;
    alert.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 9999;
        min-width: 300px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    `;
    alert.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    document.body.appendChild(alert);
    
    // 自动关闭
    setTimeout(() => {
        alert.classList.remove('show');
        setTimeout(() => alert.remove(), 150);
    }, 3000);
}

// HTML转义
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
