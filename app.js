// Excel 자동화 앱 - Main Application Logic
// Global variables
let workbook = null;
let rawData = [];
let reportData = [];

// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadStatus = document.getElementById('uploadStatus');
const statsSection = document.getElementById('statsSection');
const previewSection = document.getElementById('previewSection');
const reportSection = document.getElementById('reportSection');
const rawdataPreview = document.getElementById('rawdataPreview');
const reportPreview = document.getElementById('reportPreview');
const downloadBtn = document.getElementById('downloadBtn');
const stats = document.getElementById('stats');

// Initialize event listeners
function init() {
    // Upload area click
    uploadArea.addEventListener('click', () => fileInput.click());
    
    // File input change
    fileInput.addEventListener('change', handleFileSelect);
    
    // Drag and drop
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    
    // Download button
    downloadBtn.addEventListener('click', downloadReport);
}

// Handle drag over
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.add('dragover');
}

// Handle drag leave
function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.remove('dragover');
}

// Handle drop
function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

// Handle file select
function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

// Handle file
function handleFile(file) {
    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                        'application/vnd.ms-excel'];
    const fileName = file.name.toLowerCase();
    
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        showStatus('error', '❌ Excel 파일만 업로드할 수 있습니다 (.xlsx, .xls)');
        return;
    }
    
    showStatus('info', '⏳ 파일을 읽는 중...');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            processExcelFile(e.target.result);
        } catch (error) {
            showStatus('error', '❌ 파일 처리 중 오류가 발생했습니다: ' + error.message);
            console.error(error);
        }
    };
    reader.onerror = function() {
        showStatus('error', '❌ 파일을 읽을 수 없습니다');
    };
    reader.readAsArrayBuffer(file);
}

// Process Excel file
function processExcelFile(data) {
    // Check if XLSX library is loaded
    if (typeof XLSX === 'undefined') {
        showStatus('error', '❌ Excel 라이브러리가 로드되지 않았습니다. lib/xlsx.full.min.js 파일을 확인해주세요.');
        return;
    }
    
    // Parse workbook
    workbook = XLSX.read(data, { type: 'array' });
    
    // Check if Rawdata sheet exists
    if (!workbook.SheetNames.includes('Rawdata')) {
        showStatus('error', '❌ "Rawdata" 시트를 찾을 수 없습니다. Excel 파일에 "Rawdata" 시트가 있어야 합니다.');
        return;
    }
    
    // Get Rawdata sheet
    const rawdataSheet = workbook.Sheets['Rawdata'];
    rawData = XLSX.utils.sheet_to_json(rawdataSheet, { defval: '' });
    
    if (rawData.length === 0) {
        showStatus('error', '❌ Rawdata 시트에 데이터가 없습니다.');
        return;
    }
    
    showStatus('success', '✅ 파일을 성공적으로 읽었습니다!');
    
    // Display data
    displayRawData();
    generateReport();
    displayReport();
    displayStats();
}

// Display raw data
function displayRawData() {
    if (rawData.length === 0) return;
    
    const headers = Object.keys(rawData[0]);
    let html = '<table><thead><tr>';
    
    // Headers
    headers.forEach(header => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // Rows (show first 100 rows)
    const displayData = rawData.slice(0, 100);
    displayData.forEach(row => {
        html += '<tr>';
        headers.forEach(header => {
            html += `<td>${row[header] || ''}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    
    if (rawData.length > 100) {
        html += `<p style="padding: 10px; text-align: center; color: #666;">
            처음 100개 행만 표시됩니다. 전체 ${rawData.length}개 행이 있습니다.
        </p>`;
    }
    
    rawdataPreview.innerHTML = html;
    previewSection.classList.remove('hidden');
}

// Generate report
function generateReport() {
    if (rawData.length === 0) return;
    
    const headers = Object.keys(rawData[0]);
    reportData = [];
    
    // Create summary statistics for each column
    headers.forEach(header => {
        const values = rawData.map(row => row[header]).filter(v => v !== '' && v !== null && v !== undefined);
        const numericValues = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
        
        const summary = {
            '컬럼명': header,
            '총 데이터 수': values.length,
            '고유값 수': new Set(values).size,
            '빈값 수': rawData.length - values.length
        };
        
        // Add numeric statistics if applicable
        if (numericValues.length > 0) {
            summary['최솟값'] = Math.min(...numericValues).toFixed(2);
            summary['최댓값'] = Math.max(...numericValues).toFixed(2);
            summary['평균'] = (numericValues.reduce((a, b) => a + b, 0) / numericValues.length).toFixed(2);
            summary['합계'] = numericValues.reduce((a, b) => a + b, 0).toFixed(2);
        }
        
        reportData.push(summary);
    });
    
    // Add frequency analysis for text columns
    const frequencyAnalysis = [];
    headers.forEach(header => {
        const values = rawData.map(row => row[header]).filter(v => v !== '' && v !== null && v !== undefined);
        const frequency = {};
        
        values.forEach(v => {
            const key = String(v);
            frequency[key] = (frequency[key] || 0) + 1;
        });
        
        // Get top 5 most frequent values
        const sorted = Object.entries(frequency)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 5);
        
        if (sorted.length > 0) {
            frequencyAnalysis.push({
                '컬럼명': header,
                '상위 빈도값': sorted.map(([val, count]) => `${val} (${count})`).join(', ')
            });
        }
    });
    
    // Combine reports
    reportData = [...reportData, ...frequencyAnalysis];
}

// Display report
function displayReport() {
    if (reportData.length === 0) return;
    
    const headers = Object.keys(reportData[0]);
    let html = '<table><thead><tr>';
    
    // Headers
    headers.forEach(header => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // Rows
    reportData.forEach(row => {
        html += '<tr>';
        headers.forEach(header => {
            html += `<td>${row[header] || ''}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    
    reportPreview.innerHTML = html;
    reportSection.classList.remove('hidden');
}

// Display statistics
function displayStats() {
    const totalRows = rawData.length;
    const totalColumns = Object.keys(rawData[0] || {}).length;
    
    // Calculate numeric columns
    const headers = Object.keys(rawData[0] || {});
    let numericColumns = 0;
    let textColumns = 0;
    
    headers.forEach(header => {
        const values = rawData.map(row => row[header]).filter(v => v !== '' && v !== null && v !== undefined);
        const numericValues = values.filter(v => !isNaN(parseFloat(v)));
        
        if (numericValues.length > values.length / 2) {
            numericColumns++;
        } else {
            textColumns++;
        }
    });
    
    stats.innerHTML = `
        <div class="stat-card">
            <h3>${totalRows}</h3>
            <p>총 데이터 행</p>
        </div>
        <div class="stat-card">
            <h3>${totalColumns}</h3>
            <p>총 컬럼 수</p>
        </div>
        <div class="stat-card">
            <h3>${numericColumns}</h3>
            <p>숫자 컬럼</p>
        </div>
        <div class="stat-card">
            <h3>${textColumns}</h3>
            <p>텍스트 컬럼</p>
        </div>
    `;
    
    statsSection.classList.remove('hidden');
}

// Download report
function downloadReport() {
    if (reportData.length === 0) {
        alert('다운로드할 보고서가 없습니다.');
        return;
    }
    
    // Create new workbook
    const wb = XLSX.utils.book_new();
    
    // Add Rawdata sheet
    const rawdataWS = XLSX.utils.json_to_sheet(rawData);
    XLSX.utils.book_append_sheet(wb, rawdataWS, 'Rawdata');
    
    // Add Report sheet
    const reportWS = XLSX.utils.json_to_sheet(reportData);
    XLSX.utils.book_append_sheet(wb, reportWS, 'Report');
    
    // Generate and download
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const filename = `Report_${timestamp}.xlsx`;
    XLSX.writeFile(wb, filename);
    
    showStatus('success', `✅ 보고서가 다운로드되었습니다: ${filename}`);
}

// Show status message
function showStatus(type, message) {
    uploadStatus.className = 'status ' + type;
    uploadStatus.textContent = message;
    uploadStatus.style.display = 'block';
    
    // Auto hide after 5 seconds
    setTimeout(() => {
        uploadStatus.style.display = 'none';
    }, 5000);
}

// Initialize app when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
