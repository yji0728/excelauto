// Excel 자동화 앱 - Main Application Logic

/**
 * Application Configuration
 */
const CONFIG = {
    REQUIRED_SHEET_NAME: 'Rawdata',
    VALID_FILE_TYPES: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                       'application/vnd.ms-excel'],
    VALID_FILE_EXTENSIONS: ['.xlsx', '.xls'],
    MAX_PREVIEW_ROWS: 100,
    TOP_FREQUENCY_COUNT: 5,
    STATUS_AUTO_HIDE_DELAY: 5000
};

/**
 * Application State
 */
const AppState = {
    workbook: null,
    rawData: [],
    reportData: [],
    
    reset() {
        this.workbook = null;
        this.rawData = [];
        this.reportData = [];
    }
};

/**
 * DOM Elements Cache
 */
const DOMElements = {
    init() {
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.uploadStatus = document.getElementById('uploadStatus');
        this.statsSection = document.getElementById('statsSection');
        this.previewSection = document.getElementById('previewSection');
        this.reportSection = document.getElementById('reportSection');
        this.rawdataPreview = document.getElementById('rawdataPreview');
        this.reportPreview = document.getElementById('reportPreview');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.stats = document.getElementById('stats');
    }
};

/**
 * File Validator Module
 */
const FileValidator = {
    isValidFileType(file) {
        const fileName = file.name.toLowerCase();
        return CONFIG.VALID_FILE_EXTENSIONS.some(ext => fileName.endsWith(ext));
    },
    
    validateFile(file) {
        if (!this.isValidFileType(file)) {
            throw new Error('Excel 파일만 업로드할 수 있습니다 (.xlsx, .xls)');
        }
        return true;
    }
};

/**
 * Event Handlers Module
 */
const EventHandlers = {
    init() {
        const { uploadArea, fileInput, downloadBtn } = DOMElements;
        
        // Upload area click
        uploadArea.addEventListener('click', () => fileInput.click());
        
        // File input change
        fileInput.addEventListener('change', this.handleFileSelect.bind(this));
        
        // Drag and drop
        uploadArea.addEventListener('dragover', this.handleDragOver.bind(this));
        uploadArea.addEventListener('dragleave', this.handleDragLeave.bind(this));
        uploadArea.addEventListener('drop', this.handleDrop.bind(this));
        
        // Download button
        downloadBtn.addEventListener('click', () => DataExporter.downloadReport());
    },
    
    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        DOMElements.uploadArea.classList.add('dragover');
    },
    
    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        DOMElements.uploadArea.classList.remove('dragover');
    },
    
    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        DOMElements.uploadArea.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.handleFile(files[0]);
        }
    },
    
    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.handleFile(files[0]);
        }
    },
    
    handleFile(file) {
        try {
            FileValidator.validateFile(file);
            StatusManager.show('info', '⏳ 파일을 읽는 중...');
            FileProcessor.readFile(file);
        } catch (error) {
            StatusManager.show('error', `❌ ${error.message}`);
        }
    }
};

/**
 * File Processor Module
 */
const FileProcessor = {
    readFile(file) {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                this.processExcelFile(e.target.result);
            } catch (error) {
                StatusManager.show('error', `❌ 파일 처리 중 오류가 발생했습니다: ${error.message}`);
                console.error(error);
            }
        };
        
        reader.onerror = () => {
            StatusManager.show('error', '❌ 파일을 읽을 수 없습니다');
        };
        
        reader.readAsArrayBuffer(file);
    },
    
    processExcelFile(data) {
        // Check if XLSX library is loaded
        if (typeof XLSX === 'undefined') {
            throw new Error('Excel 라이브러리가 로드되지 않았습니다. lib/xlsx.full.min.js 파일을 확인해주세요.');
        }
        
        // Parse workbook
        AppState.workbook = XLSX.read(data, { type: 'array' });
        
        // Validate and extract data
        this.validateAndExtractData();
        
        // Display results
        StatusManager.show('success', '✅ 파일을 성공적으로 읽었습니다!');
        UIRenderer.displayAll();
    },
    
    validateAndExtractData() {
        const { workbook } = AppState;
        
        // Check if Rawdata sheet exists
        if (!workbook.SheetNames.includes(CONFIG.REQUIRED_SHEET_NAME)) {
            throw new Error(`"${CONFIG.REQUIRED_SHEET_NAME}" 시트를 찾을 수 없습니다. Excel 파일에 "${CONFIG.REQUIRED_SHEET_NAME}" 시트가 있어야 합니다.`);
        }
        
        // Get Rawdata sheet
        const rawdataSheet = workbook.Sheets[CONFIG.REQUIRED_SHEET_NAME];
        AppState.rawData = XLSX.utils.sheet_to_json(rawdataSheet, { defval: '' });
        
        if (AppState.rawData.length === 0) {
            throw new Error('Rawdata 시트에 데이터가 없습니다.');
        }
    }
};

/**
 * Data Analyzer Module
 */
const DataAnalyzer = {
    generateReport() {
        if (AppState.rawData.length === 0) return;
        
        const headers = Object.keys(AppState.rawData[0]);
        AppState.reportData = [];
        
        // Generate column statistics
        const columnStats = this.generateColumnStatistics(headers);
        const frequencyAnalysis = this.generateFrequencyAnalysis(headers);
        
        // Combine reports
        AppState.reportData = [...columnStats, ...frequencyAnalysis];
    },
    
    generateColumnStatistics(headers) {
        return headers.map(header => {
            const values = this.getColumnValues(header);
            const numericValues = this.getNumericValues(values);
            
            const summary = {
                '컬럼명': header,
                '총 데이터 수': values.length,
                '고유값 수': new Set(values).size,
                '빈값 수': AppState.rawData.length - values.length
            };
            
            // Add numeric statistics if applicable
            if (numericValues.length > 0) {
                Object.assign(summary, this.calculateNumericStats(numericValues));
            }
            
            return summary;
        });
    },
    
    generateFrequencyAnalysis(headers) {
        return headers.map(header => {
            const values = this.getColumnValues(header);
            const frequency = this.calculateFrequency(values);
            
            // Get top N most frequent values
            const topValues = Object.entries(frequency)
                .sort((a, b) => b[1] - a[1])
                .slice(0, CONFIG.TOP_FREQUENCY_COUNT);
            
            return {
                '컬럼명': header,
                '상위 빈도값': topValues.map(([val, count]) => `${val} (${count})`).join(', ')
            };
        }).filter(item => item['상위 빈도값']);
    },
    
    getColumnValues(header) {
        return AppState.rawData
            .map(row => row[header])
            .filter(v => v !== '' && v !== null && v !== undefined);
    },
    
    getNumericValues(values) {
        return values
            .filter(v => !isNaN(parseFloat(v)))
            .map(v => parseFloat(v));
    },
    
    calculateNumericStats(numericValues) {
        const sum = numericValues.reduce((a, b) => a + b, 0);
        const avg = sum / numericValues.length;
        
        return {
            '최솟값': Math.min(...numericValues).toFixed(2),
            '최댓값': Math.max(...numericValues).toFixed(2),
            '평균': avg.toFixed(2),
            '합계': sum.toFixed(2)
        };
    },
    
    calculateFrequency(values) {
        const frequency = {};
        values.forEach(v => {
            const key = String(v);
            frequency[key] = (frequency[key] || 0) + 1;
        });
        return frequency;
    },
    
    calculateColumnTypes(headers) {
        let numericColumns = 0;
        let textColumns = 0;
        
        headers.forEach(header => {
            const values = this.getColumnValues(header);
            const numericValues = this.getNumericValues(values);
            
            if (numericValues.length > values.length / 2) {
                numericColumns++;
            } else {
                textColumns++;
            }
        });
        
        return { numericColumns, textColumns };
    }
};

/**
 * UI Renderer Module
 */
const UIRenderer = {
    displayAll() {
        this.displayRawData();
        DataAnalyzer.generateReport();
        this.displayReport();
        this.displayStats();
    },
    
    displayRawData() {
        if (AppState.rawData.length === 0) return;
        
        const headers = Object.keys(AppState.rawData[0]);
        const displayData = AppState.rawData.slice(0, CONFIG.MAX_PREVIEW_ROWS);
        
        let html = this.createTableHTML(headers, displayData);
        
        if (AppState.rawData.length > CONFIG.MAX_PREVIEW_ROWS) {
            html += this.createOverflowMessage(AppState.rawData.length);
        }
        
        DOMElements.rawdataPreview.innerHTML = html;
        DOMElements.previewSection.classList.remove('hidden');
    },
    
    displayReport() {
        if (AppState.reportData.length === 0) return;
        
        const headers = Object.keys(AppState.reportData[0]);
        const html = this.createTableHTML(headers, AppState.reportData);
        
        DOMElements.reportPreview.innerHTML = html;
        DOMElements.reportSection.classList.remove('hidden');
    },
    
    displayStats() {
        const totalRows = AppState.rawData.length;
        const headers = Object.keys(AppState.rawData[0] || {});
        const totalColumns = headers.length;
        
        const { numericColumns, textColumns } = DataAnalyzer.calculateColumnTypes(headers);
        
        DOMElements.stats.innerHTML = `
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
        
        DOMElements.statsSection.classList.remove('hidden');
    },
    
    createTableHTML(headers, data) {
        let html = '<table><thead><tr>';
        
        // Headers
        headers.forEach(header => {
            html += `<th>${this.escapeHtml(header)}</th>`;
        });
        html += '</tr></thead><tbody>';
        
        // Rows
        data.forEach(row => {
            html += '<tr>';
            headers.forEach(header => {
                html += `<td>${this.escapeHtml(row[header] || '')}</td>`;
            });
            html += '</tr>';
        });
        html += '</tbody></table>';
        
        return html;
    },
    
    createOverflowMessage(totalRows) {
        return `<p style="padding: 10px; text-align: center; color: #666;">
            처음 ${CONFIG.MAX_PREVIEW_ROWS}개 행만 표시됩니다. 전체 ${totalRows}개 행이 있습니다.
        </p>`;
    },
    
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = String(text);
        return div.innerHTML;
    }
};

/**
 * Data Exporter Module
 */
const DataExporter = {
    downloadReport() {
        if (AppState.reportData.length === 0) {
            alert('다운로드할 보고서가 없습니다.');
            return;
        }
        
        try {
            // Create new workbook
            const wb = XLSX.utils.book_new();
            
            // Add Rawdata sheet
            const rawdataWS = XLSX.utils.json_to_sheet(AppState.rawData);
            XLSX.utils.book_append_sheet(wb, rawdataWS, CONFIG.REQUIRED_SHEET_NAME);
            
            // Add Report sheet
            const reportWS = XLSX.utils.json_to_sheet(AppState.reportData);
            XLSX.utils.book_append_sheet(wb, reportWS, 'Report');
            
            // Generate and download
            const filename = this.generateFilename();
            XLSX.writeFile(wb, filename);
            
            StatusManager.show('success', `✅ 보고서가 다운로드되었습니다: ${filename}`);
        } catch (error) {
            StatusManager.show('error', `❌ 다운로드 중 오류가 발생했습니다: ${error.message}`);
            console.error(error);
        }
    },
    
    generateFilename() {
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        return `Report_${timestamp}.xlsx`;
    }
};

/**
 * Status Manager Module
 */
const StatusManager = {
    show(type, message) {
        const { uploadStatus } = DOMElements;
        uploadStatus.className = `status ${type}`;
        uploadStatus.textContent = message;
        uploadStatus.style.display = 'block';
        
        // Auto hide after delay
        this.scheduleHide();
    },
    
    scheduleHide() {
        if (this.hideTimeout) {
            clearTimeout(this.hideTimeout);
        }
        
        this.hideTimeout = setTimeout(() => {
            DOMElements.uploadStatus.style.display = 'none';
        }, CONFIG.STATUS_AUTO_HIDE_DELAY);
    },
    
    hideTimeout: null
};

/**
 * Application Initialization
 */
function initializeApp() {
    DOMElements.init();
    EventHandlers.init();
}

// Initialize app when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}
