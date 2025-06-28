// Global variables
let currentData = {
    case: [],
    talent: [],
    matching: []
};
let filteredData = {
    case: [],
    talent: [],
    matching: []
};
let currentPage = {
    case: 1,
    talent: 1,
    matching: 1
};
let recordsPerPage = 3;
let isAuthenticated = false;

// Column mappings for different tabs
const columnMappings = {
    case: {
        'Processed At': 0,
        'Email Timestamp': 1,
        'Email Subject': 2,
        'Project ID': 3,
        'Project Title': 4,
        'Project Description': 5,
        'Work Rate': 6,
        'Start Time': 7,
        'Contact Person': 8,
        'Contact Information': 9,
        'Project Type': 10,
        'Tech Requirements': 11,
        'Duration': 12,
        'Work Style': 13,
        'Client Company': 14,
        'Urgency': 15,
        'Special Requirements': 16
    },
    talent: {
        '分析日期': 0,
        'PDF 文件名': 1,
        'Resume ID': 2,
        '全名': 3,
        '职业头衔': 4,
        '工作经验年限': 5,
        '技术技能': 6,
        '证书': 7,
        '教育背景': 8,
        '先前雇主': 9,
        '主要项目': 10,
        '首选工作地点': 11,
        '语言能力': 12,
        '联系邮箱': 13,
        'LinkedIn 或个人网站': 14
    },
    matching: {
        'Match_Date': 0,
        'Talent_ID': 1,
        'Talent_Name': 2,
        'Match 1 Score': 3,
        'Match 1 Name': 4,
        'Match 2 Score': 5,
        'Match 2 Name': 6,
        'Match 3 Score': 7,
        'Match 3 Name': 8,
        'Match Details': 9
    }
};

// Predefined headers for each tab
const predefinedHeaders = {
    case: [
        'Processed At', 'Email Timestamp', 'Email Subject', 'Project ID',
        'Project Title', 'Project Description', 'Work Rate', 'Start Time',
        'Contact Person', 'Contact Information', 'Project Type',
        'Tech Requirements', 'Duration', 'Work Style', 'Client Company',
        'Urgency', 'Special Requirements'
    ],
    talent: [
        '分析日期', 'PDF 文件名', 'Resume ID', '全名', '职业头衔',
        '工作经验年限', '技术技能', '证书', '教育背景', '先前雇主',
        '主要项目', '首选工作地点', '语言能力', '联系邮箱', 'LinkedIn 或个人网站'
    ],
    matching: [
        'Match_Date', 'Talent_ID', 'Talent_Name',
        'Match 1 Score', 'Match 1 Name',
        'Match 2 Score', 'Match 2 Name',
        'Match 3 Score', 'Match 3 Name', 'Match Details'
    ]
};

// Sheet names configuration
const sheetNames = {
    case: 'projects',
    talent: 'resume_database',
    matching: 'matches'
};

// Load configuration from localStorage
function loadConfig() {
    const apiKey = localStorage.getItem('projectMgmtApiKey');
    const sheetId = localStorage.getItem('projectMgmtSheetId');

    if (apiKey) document.getElementById('apiKey').value = apiKey;
    if (sheetId) document.getElementById('sheetId').value = sheetId;
}

// Save configuration to localStorage
function saveConfig() {
    localStorage.setItem('projectMgmtApiKey', document.getElementById('apiKey').value);
    localStorage.setItem('projectMgmtSheetId', document.getElementById('sheetId').value);
}

// Auto-detect data range for a sheet
async function detectSheetRange(apiKey, sheetId, sheetName) {
    try {
        // First, get basic sheet info to determine dimensions
        const metadataUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?key=${apiKey}`;
        const metadataResponse = await fetch(metadataUrl);
        
        if (!metadataResponse.ok) {
            throw new Error(`无法获取表格信息: ${metadataResponse.status}`);
        }
        
        const metadata = await metadataResponse.json();
        const sheet = metadata.sheets.find(s => s.properties.title === sheetName);
        
        if (!sheet) {
            throw new Error(`未找到名为 "${sheetName}" 的工作表`);
        }
        
        const gridProperties = sheet.properties.gridProperties;
        const maxRows = Math.min(gridProperties.rowCount || 1000, 1000);
        const maxCols = Math.min(gridProperties.columnCount || 26, 26);
        
        // Convert column number to letter (A, B, C, ... Z)
        const getColumnLetter = (colNum) => {
            return String.fromCharCode(65 + colNum - 1);
        };
        
        const endColumn = getColumnLetter(maxCols);
        const range = `${sheetName}!A1:${endColumn}${maxRows}`;
        
        console.log(`检测到的数据范围: ${range}`);
        return range;
        
    } catch (error) {
        console.error(`检测工作表范围时出错:`, error);
        // Fallback to default range
        return `${sheetName}!A1:Z1000`;
    }
}

// Switch between tabs
function switchTab(tabName) {
    if (!isAuthenticated) return;

    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Remove active class from all tab buttons
    document.querySelectorAll('.tab-button').forEach(button => {
        button.classList.remove('active');
    });
    
    // Show selected tab content
    document.getElementById(tabName + 'Content').classList.add('active');
    document.getElementById(tabName + 'Tab').classList.add('active');
    
    // Load data if not already loaded
    if (currentData[tabName].length === 0) {
        loadData(tabName);
    } else {
        // Reset pagination when switching tabs
        currentPage[tabName] = 1;
        renderTable(tabName, filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName]);
    }
}

// Authenticate and load initial data
async function authenticateAndLoad() {
    const apiKey = document.getElementById('apiKey').value.trim();
    const sheetId = document.getElementById('sheetId').value.trim();

    if (!apiKey || !sheetId) {
        showError('case', '请输入Google API密钥和表格ID');
        return;
    }

    saveConfig();
    
    try {
        // Test authentication with case database first
        console.log('正在验证认证信息...');
        showLoading('case');
        
        // Detect range and load case data
        const caseRange = await detectSheetRange(apiKey, sheetId, sheetNames.case);
        await loadDataWithRange('case', apiKey, sheetId, caseRange);
        
        // If successful, show tabs and hide config
        isAuthenticated = true;
        document.getElementById('configSection').style.display = 'none';
        document.getElementById('tabNavigation').style.display = 'flex';
        
        console.log('认证成功，已启用所有标签页');
        
    } catch (error) {
        console.error('认证失败:', error);
        showError('case', '认证失败: ' + error.message);
    }
}

// Show error message
function showError(tabName, message) {
    const errorDiv = document.getElementById(tabName + 'Error');
    errorDiv.innerHTML = `<strong>错误:</strong> ${message}`;
    errorDiv.style.display = 'block';
    document.getElementById(tabName + 'Loading').style.display = 'none';
    document.getElementById(tabName + 'DataContainer').style.display = 'none';
    document.getElementById(tabName + 'Stats').style.display = 'none';
}

// Hide error message
function hideError(tabName) {
    document.getElementById(tabName + 'Error').style.display = 'none';
}

// Show loading state
function showLoading(tabName) {
    document.getElementById(tabName + 'Loading').style.display = 'block';
    document.getElementById(tabName + 'DataContainer').style.display = 'none';
    document.getElementById(tabName + 'Stats').style.display = 'none';
    hideError(tabName);
}

// Generate statistics for different tabs
function generateStats(tabName, data) {
    const statsContainer = document.getElementById(tabName + 'Stats');
    
    if (!data || data.length === 0) {
        statsContainer.style.display = 'none';
        return;
    }
    
    const hasHeader = data[0] && (
        data[0].join('').toLowerCase().includes('processed') ||
        data[0].join('').toLowerCase().includes('resume') ||
        data[0].join('').toLowerCase().includes('match')
    );
    const startIndex = hasHeader ? 1 : 0;
    const totalItems = data.length - startIndex;
    
    let statsHTML = '';
    
    if (tabName === 'case') {
        let processedCount = 0;
        let totalWorkRate = 0;
        let validRates = 0;
        
        for (let i = startIndex; i < data.length; i++) {
            const processedStatus = data[i][0] || '';
            if (processedStatus.includes('已') || processedStatus.toLowerCase().includes('yes')) {
                processedCount++;
            }
            
            const workRate = parseFloat(data[i][6]) || 0;
            if (workRate > 0) {
                totalWorkRate += workRate;
                validRates++;
            }
        }
        
        const avgWorkRate = validRates > 0 ? (totalWorkRate / validRates).toFixed(2) : 0;
        const completionRate = totalItems > 0 ? ((processedCount / totalItems) * 100).toFixed(1) : 0;

        statsHTML = `
            <div class="stat-card">
                <div class="number">${totalItems}</div>
                <div class="label">总项目数</div>
            </div>
            <div class="stat-card">
                <div class="number">${processedCount}</div>
                <div class="label">已处理项目</div>
            </div>
            <div class="stat-card">
                <div class="number">${completionRate}%</div>
                <div class="label">完成率</div>
            </div>
            <div class="stat-card">
                <div class="number">¥${avgWorkRate}</div>
                <div class="label">平均工作费率</div>
            </div>
        `;
    } else if (tabName === 'talent') {
        let expCounts = { '0-2': 0, '3-5': 0, '6-10': 0, '10+': 0 };
        let skillSet = new Set();
        
        for (let i = startIndex; i < data.length; i++) {
            const experience = data[i][5] || '';
            const skills = data[i][6] || '';
            
            const expNum = parseInt(experience) || 0;
            if (expNum <= 2) expCounts['0-2']++;
            else if (expNum <= 5) expCounts['3-5']++;
            else if (expNum <= 10) expCounts['6-10']++;
            else expCounts['10+']++;
            
            if (skills) {
                skills.split(/[,，、]/).forEach(skill => {
                    skillSet.add(skill.trim());
                });
            }
        }

        statsHTML = `
            <div class="stat-card">
                <div class="number">${totalItems}</div>
                <div class="label">总人才数</div>
            </div>
            <div class="stat-card">
                <div class="number">${skillSet.size}</div>
                <div class="label">技能种类</div>
            </div>
            <div class="stat-card">
                <div class="number">${expCounts['10+']}</div>
                <div class="label">资深人才 (10+年)</div>
            </div>
            <div class="stat-card">
                <div class="number">${expCounts['0-2']}</div>
                <div class="label">新手人才 (0-2年)</div>
            </div>
        `;
    } else if (tabName === 'matching') {
        let highMatches = 0;
        let mediumMatches = 0;
        let lowMatches = 0;
        let totalScore = 0;
        let validScores = 0;
        
        for (let i = startIndex; i < data.length; i++) {
            // Updated column indices - removed Match ID column
            const score1 = parseFloat(data[i][4]) || 0;  // Match 1 Score
            const score2 = parseFloat(data[i][7]) || 0;  // Match 2 Score
            const score3 = parseFloat(data[i][10]) || 0;  // Match 3 Score
            
            [score1, score2, score3].forEach(score => {
                if (score > 0) {
                    totalScore += score;
                    validScores++;
                    if (score >= 80) highMatches++;
                    else if (score >= 60) mediumMatches++;
                    else lowMatches++;
                }
            });
        }
        
        const avgScore = validScores > 0 ? (totalScore / validScores).toFixed(1) : 0;
    
        statsHTML = `
            <div class="stat-card">
                <div class="number">${totalItems}</div>
                <div class="label">总匹配记录</div>
            </div>
            <div class="stat-card">
                <div class="number">${highMatches}</div>
                <div class="label">高分匹配 (≥80)</div>
            </div>
            <div class="stat-card">
                <div class="number">${mediumMatches}</div>
                <div class="label">中等匹配 (60-79)</div>
            </div>
            <div class="stat-card">
                <div class="number">${avgScore}</div>
                <div class="label">平均匹配分数</div>
            </div>
        `;
    }
    
    statsContainer.innerHTML = statsHTML;
    statsContainer.style.display = 'grid';
}

// Format cell content with proper styling
function formatCellContent(value, columnIndex, tabName) {
    if (!value) return '';
    
    const cellValue = value.toString();
    
    if (tabName === 'case') {
        switch (columnIndex) {
            case 0: // Processed At
                const processed = cellValue.toLowerCase();
                if (processed.includes('已') || processed.includes('yes')) {
                    return `<span class="processed-status processed-yes">✅ ${cellValue}</span>`;
                } else {
                    return `<span class="processed-status processed-no">⏳ ${cellValue}</span>`;
                }
            case 1:
            case 7: // Timestamps
                return `<span class="timestamp">${cellValue}</span>`;
            case 2: // Email Subject
                return `<span class="email-subject">${cellValue}</span>`;
            case 3: // Project ID
                return `<span class="project-id">${cellValue}</span>`;
            case 4: // Project Title
                return `<span class="project-title">${cellValue}</span>`;
            case 5: // Project Description
                return `<span class="project-description">${cellValue}</span>`;
            case 6: // Work Rate
                return `<span class="work-rate">${cellValue}</span>`;
            case 8:
            case 9: // Contact info
                return `<span class="contact-info">${cellValue}</span>`;
            case 10: // Project Type
                return `<span class="project-type">${cellValue}</span>`;
            case 11: // Tech Requirements
                const skills = cellValue.split(/[,，、]/).slice(0, 5);
                return `<div class="skill-tags">${skills.map(skill => 
                    `<span class="skill-tag">${skill.trim()}</span>`
                ).join('')}</div>`;
            case 12: // Duration
                return `<span class="project-duration">${cellValue}</span>`;
            case 13: // Work Style
                return `<span class="work-style">${cellValue}</span>`;
            case 14: // Client Company
                return `<span class="client-company">${cellValue}</span>`;
            case 15: // Urgency
                const urgencyClass = cellValue.toLowerCase().includes('紧急') || cellValue.toLowerCase().includes('urgent') ? 'urgent' : 'normal';
                return `<span class="urgency ${urgencyClass}">${cellValue}</span>`;
            case 16: // Special Requirements
                return `<span class="special-requirements">${cellValue}</span>`;
        }
    } else if (tabName === 'talent') {
        switch (columnIndex) {
            case 0: // Analysis Date
                return `<span class="timestamp">${cellValue}</span>`;
            case 2: // Resume ID
                return `<span class="talent-id">${cellValue}</span>`;
            case 3: // Full Name
                return `<span class="project-title">${cellValue}</span>`;
            case 5: // Experience Years
                return `<span class="experience-years">${cellValue}年</span>`;
            case 6: // Skills
                const skills = cellValue.split(/[,，、]/).slice(0, 5);
                return `<div class="skill-tags">${skills.map(skill => 
                    `<span class="skill-tag">${skill.trim()}</span>`
                ).join('')}</div>`;
            case 13: // Email
                return `<span class="email-cell">${cellValue}</span>`;
            case 14: // LinkedIn
                if (cellValue.startsWith('http')) {
                    return `<a href="${cellValue}" target="_blank">${cellValue}</a>`;
                }
                return cellValue;
        }
   } else if (tabName === 'matching') {
        switch (columnIndex) {
            case 0: // Match_Date
                return `<span class="timestamp">${cellValue}</span>`;
            case 1: // Talent_ID
                return `<span class="talent-id">${cellValue}</span>`;
            case 2: // Talent_Name
                return `<span class="project-title">${cellValue}</span>`;
            case 3:
            case 5:
            case 7: // Match 1/2/3 Score (remapped indices)
                const score = parseFloat(cellValue) || 0;
                let scoreClass = 'low';
                if (score >= 80) scoreClass = 'high';
                else if (score >= 60) scoreClass = 'medium';
                return `<span class="match-score ${scoreClass}">${cellValue}</span>`;
            case 4:
            case 6:
            case 8: // Match 1/2/3 Name (remapped indices)
                return `<span class="project-title">${cellValue}</span>`;
            case 9: // Match Details (remapped indices)
                return formatMatchDetails(cellValue);
        }
    }
                    
    return cellValue;
}

// Format match details JSON into compact display
function formatMatchDetails(jsonString) {
    if (!jsonString || jsonString.trim() === '') {
        return '<div class="match-details-compact">📋 暂无匹配详情</div>';
    }
    
    try {
        const data = JSON.parse(jsonString);
        
        if (!data.matchedItems || !Array.isArray(data.matchedItems) || data.matchedItems.length === 0) {
            return '<div class="match-details-compact">📋 暂无匹配项目</div>';
        }
        
        // Improved: include reason in display
        let detailsHtml = '<div class="match-details-with-reason">';
        detailsHtml += '<div>🎯 匹配项目:</div>';
        
        data.matchedItems.slice(0, 3).forEach((match, index) => {
            if (match.name && match.score) {
                detailsHtml += `<div>• ${match.name} (${match.score}分)</div>`;
                if (match.reason) {
                    // Show reason
                    detailsHtml += `<div class="match-item-reason">💡 ${match.reason}</div>`;
                }
            }
        });
        
        detailsHtml += '</div>';
        return detailsHtml;
        
    } catch (error) {
        console.error('JSON解析错误:', error);
        return '<div class="match-details-compact match-error">⚠️ 数据解析失败</div>';
    }
}

// Render table for specific tab with pagination
function renderTable(tabName, data) {
    const table = document.getElementById(tabName + 'DataTable');
    
    if (!data || data.length === 0) {
        table.innerHTML = '<tr><td colspan="15">没有找到数据</td></tr>';
        updatePagination(tabName, 0, []);
        return;
    }

    // Determine if there's a header
    const hasHeader = data[0] && (
        data[0].join('').toLowerCase().includes('processed') ||
        data[0].join('').toLowerCase().includes('resume') ||
        data[0].join('').toLowerCase().includes('match')
    );
    
    const headerRow = hasHeader ? data[0] : null;
    const dataRows = hasHeader ? data.slice(1) : data;
    
    // Define hidden column indices (for matching table)
    const hiddenColumns = tabName === 'matching' ? [3, 6, 9] : []; // Match ID columns
    
    // Apply pagination
    const currentPageNum = currentPage[tabName];
    const pageSize = getPageSize(tabName);
    const startIndex = (currentPageNum - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const paginatedData = dataRows.slice(startIndex, endIndex);

    let html = '';
    
    // Use predefined headers (already removed Match ID)
    html += '<thead><tr>';
    predefinedHeaders[tabName].forEach((header) => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead>';

    // Create table body
    html += '<tbody>';
    
    if (paginatedData.length === 0) {
        const colspan = predefinedHeaders[tabName].length;
        html += `<tr><td colspan="${colspan}" style="text-align: center; padding: 40px; color: #6b7280;">当前页没有数据</td></tr>`;
    } else {
        for (let i = 0; i < paginatedData.length; i++) {
            const row = paginatedData[i];
            if (row && row.some(cell => cell !== undefined && cell !== '')) {
                html += '<tr>';
                
                // Filter and remap column indices
                let displayColumnIndex = 0;
                for (let j = 0; j < row.length && displayColumnIndex < predefinedHeaders[tabName].length; j++) {
                    // Skip hidden columns
                    if (hiddenColumns.includes(j)) {
                        continue;
                    }
                    
                    const cellValue = row[j] || '';
                    const formattedValue = formatCellContent(cellValue, displayColumnIndex, tabName);
                    html += `<td>${formattedValue}</td>`;
                    displayColumnIndex++;
                }
                html += '</tr>';
            }
        }
    }
    html += '</tbody>';

    table.innerHTML = html;
    
    // Update pagination controls
    updatePagination(tabName, dataRows.length, dataRows);
}

// Get current page size for a tab
function getPageSize(tabName) {
    const select = document.getElementById(tabName + 'PageSize');
    return parseInt(select.value) || 3;
}

// Update pagination controls and info
function updatePagination(tabName, totalRecords, dataRows) {
    const paginationContainer = document.getElementById(tabName + 'Pagination');
    const paginationInfo = document.getElementById(tabName + 'PaginationInfo');
    const pageInput = document.getElementById(tabName + 'PageInput');
    
    const pageSize = getPageSize(tabName);
    const totalPages = Math.ceil(totalRecords / pageSize) || 1;
    const currentPageNum = Math.min(currentPage[tabName], totalPages);
    
    // Update current page if it exceeds total pages
    currentPage[tabName] = currentPageNum;
    
    const startRecord = totalRecords > 0 ? (currentPageNum - 1) * pageSize + 1 : 0;
    const endRecord = Math.min(currentPageNum * pageSize, totalRecords);
    
    // Show pagination if there are records
    if (totalRecords > 0) {
        paginationContainer.style.display = 'flex';
        paginationInfo.textContent = `显示第 ${startRecord}-${endRecord} 条，共 ${totalRecords} 条记录`;
        pageInput.value = currentPageNum;
        pageInput.max = totalPages;
        
        // Update button states
        const firstBtn = document.getElementById(tabName + 'FirstBtn');
        const prevBtn = document.getElementById(tabName + 'PrevBtn');
        const nextBtn = document.getElementById(tabName + 'NextBtn');
        const lastBtn = document.getElementById(tabName + 'LastBtn');
        
        firstBtn.disabled = currentPageNum === 1;
        prevBtn.disabled = currentPageNum === 1;
        nextBtn.disabled = currentPageNum === totalPages;
        lastBtn.disabled = currentPageNum === totalPages;
        
    } else {
        paginationContainer.style.display = 'none';
    }
}

// Pagination functions
function goToPage(tabName, pageNum) {
    const dataToUse = filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName];
    const hasHeader = dataToUse[0] && (
        dataToUse[0].join('').toLowerCase().includes('processed') ||
        dataToUse[0].join('').toLowerCase().includes('resume') ||
        dataToUse[0].join('').toLowerCase().includes('match')
    );
    const dataRows = hasHeader ? dataToUse.slice(1) : dataToUse;
    const pageSize = getPageSize(tabName);
    const totalPages = Math.ceil(dataRows.length / pageSize) || 1;
    
    currentPage[tabName] = Math.max(1, Math.min(pageNum, totalPages));
    renderTable(tabName, dataToUse);
}

function previousPage(tabName) {
    if (currentPage[tabName] > 1) {
        goToPage(tabName, currentPage[tabName] - 1);
    }
}

function nextPage(tabName) {
    const dataToUse = filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName];
    const hasHeader = dataToUse[0] && (
        dataToUse[0].join('').toLowerCase().includes('processed') ||
        dataToUse[0].join('').toLowerCase().includes('resume') ||
        dataToUse[0].join('').toLowerCase().includes('match')
    );
    const dataRows = hasHeader ? dataToUse.slice(1) : dataToUse;
    const pageSize = getPageSize(tabName);
    const totalPages = Math.ceil(dataRows.length / pageSize) || 1;
    
    if (currentPage[tabName] < totalPages) {
        goToPage(tabName, currentPage[tabName] + 1);
    }
}

function goToLastPage(tabName) {
    const dataToUse = filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName];
    const hasHeader = dataToUse[0] && (
        dataToUse[0].join('').toLowerCase().includes('processed') ||
        dataToUse[0].join('').toLowerCase().includes('resume') ||
        dataToUse[0].join('').toLowerCase().includes('match')
    );
    const dataRows = hasHeader ? dataToUse.slice(1) : dataToUse;
    const pageSize = getPageSize(tabName);
    const totalPages = Math.ceil(dataRows.length / pageSize) || 1;
    
    goToPage(tabName, totalPages);
}

function goToPageInput(tabName) {
    const pageInput = document.getElementById(tabName + 'PageInput');
    const pageNum = parseInt(pageInput.value);
    if (pageNum && pageNum > 0) {
        goToPage(tabName, pageNum);
    }
}

function handlePageInputEnter(event, tabName) {
    if (event.key === 'Enter') {
        goToPageInput(tabName);
    }
}

function changePageSize(tabName) {
    currentPage[tabName] = 1; // Reset to first page when changing page size
    const dataToUse = filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName];
    renderTable(tabName, dataToUse);
}

// Apply filters for specific tab
function applyFilters(tabName) {
    if (!currentData[tabName] || currentData[tabName].length === 0) return;

    const data = currentData[tabName];
    const hasHeader = data[0] && (
        data[0].join('').toLowerCase().includes('processed') ||
        data[0].join('').toLowerCase().includes('resume') ||
        data[0].join('').toLowerCase().includes('match')
    );
    const startIndex = hasHeader ? 1 : 0;
    
    filteredData[tabName] = hasHeader ? [data[0]] : [];

    for (let i = startIndex; i < data.length; i++) {
        const row = data[i];
        let include = true;

        if (tabName === 'case') {
            const processedFilter = document.getElementById('caseProcessedFilter').value;
            const projectFilter = document.getElementById('caseProjectFilter').value.toLowerCase();
            const contactFilter = document.getElementById('caseContactFilter').value.toLowerCase();
            const dateFilter = document.getElementById('caseDateFilter').value;

            if (processedFilter) {
                const processedStatus = (row[0] || '').toString();
                if (processedFilter === '已处理') {
                    if (!processedStatus.includes('已') && !processedStatus.toLowerCase().includes('yes')) {
                        include = false;
                    }
                } else if (processedFilter === '未处理') {
                    if (processedStatus.includes('已') || processedStatus.toLowerCase().includes('yes')) {
                        include = false;
                    }
                }
            }

            if (projectFilter && include) {
                const projectId = (row[3] || '').toString().toLowerCase();
                if (!projectId.includes(projectFilter)) include = false;
            }

            if (contactFilter && include) {
                const contactPerson = (row[8] || '').toString().toLowerCase();
                if (!contactPerson.includes(contactFilter)) include = false;
            }

            if (dateFilter && include) {
                const startTime = row[7] || '';
                if (startTime && !startTime.includes(dateFilter)) include = false;
            }
        } else if (tabName === 'talent') {
            const skillFilter = document.getElementById('talentSkillFilter').value.toLowerCase();
            const locationFilter = document.getElementById('talentLocationFilter').value.toLowerCase();
            const experienceFilter = document.getElementById('talentExperienceFilter').value;
            const nameFilter = document.getElementById('talentNameFilter').value.toLowerCase();

            if (skillFilter && include) {
                const skills = (row[6] || '').toString().toLowerCase();
                if (!skills.includes(skillFilter)) include = false;
            }

            if (locationFilter && include) {
                const location = (row[11] || '').toString().toLowerCase();
                if (!location.includes(locationFilter)) include = false;
            }

            if (experienceFilter && include) {
                const experience = parseInt(row[5]) || 0;
                if (experienceFilter === '0-2' && (experience < 0 || experience > 2)) include = false;
                else if (experienceFilter === '3-5' && (experience < 3 || experience > 5)) include = false;
                else if (experienceFilter === '6-10' && (experience < 6 || experience > 10)) include = false;
                else if (experienceFilter === '10+' && experience <= 10) include = false;
            }

            if (nameFilter && include) {
                const name = (row[3] || '').toString().toLowerCase();
                if (!name.includes(nameFilter)) include = false;
            }
        } else if (tabName === 'matching') {
            const talentFilter = document.getElementById('matchingTalentFilter').value.toLowerCase();
            const scoreFilter = parseInt(document.getElementById('matchingScoreFilter').value) || 0;
            const dateFilter = document.getElementById('matchingDateFilter').value;
        
            if (talentFilter && include) {
                const talentId = (row[1] || '').toString().toLowerCase();
                if (!talentId.includes(talentFilter)) include = false;
            }
        
            if (scoreFilter && include) {
                // Updated column indices - removed Match ID columns
                const score1 = parseFloat(row[4]) || 0;  // Match 1 Score
                const score2 = parseFloat(row[7]) || 0;  // Match 2 Score
                const score3 = parseFloat(row[10]) || 0;  // Match 3 Score
                if (Math.max(score1, score2, score3) < scoreFilter) include = false;
            }
        
            if (dateFilter && include) {
                const matchDate = row[0] || '';
                if (matchDate && !matchDate.includes(dateFilter)) include = false;
            }
        }

        if (include) {
            filteredData[tabName].push(row);
        }
    }

    currentPage[tabName] = 1; // Reset to first page when applying filters
    renderTable(tabName, filteredData[tabName]);
    generateStats(tabName, filteredData[tabName]);
}

// Clear filters for specific tab
function clearFilters(tabName) {
    if (tabName === 'case') {
        document.getElementById('caseProcessedFilter').value = '';
        document.getElementById('caseProjectFilter').value = '';
        document.getElementById('caseContactFilter').value = '';
        document.getElementById('caseDateFilter').value = '';
    } else if (tabName === 'talent') {
        document.getElementById('talentSkillFilter').value = '';
        document.getElementById('talentLocationFilter').value = '';
        document.getElementById('talentExperienceFilter').value = '';
        document.getElementById('talentNameFilter').value = '';
    } else if (tabName === 'matching') {
        document.getElementById('matchingTalentFilter').value = '';
        document.getElementById('matchingScoreFilter').value = '';
        document.getElementById('matchingDateFilter').value = '';
    }
    
    filteredData[tabName] = [...currentData[tabName]];
    currentPage[tabName] = 1; // Reset to first page when clearing filters
    renderTable(tabName, currentData[tabName]);
    generateStats(tabName, currentData[tabName]);
}

// Export to CSV for specific tab
function exportToCSV(tabName) {
    const dataToExport = filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName];
    
    if (!dataToExport || dataToExport.length === 0) {
        alert('没有数据可以导出');
        return;
    }

    let csvContent = '';
    dataToExport.forEach(row => {
        const csvRow = row.map(cell => {
            const cleanCell = (cell || '').toString().replace(/<[^>]*>/g, '').replace(/"/g, '""');
            return `"${cleanCell}"`;
        }).join(',');
        csvContent += csvRow + '\n';
    });

    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${tabName}_数据_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
}

// Load data for specific tab with automatic range detection
async function loadData(tabName) {
    const apiKey = document.getElementById('apiKey').value.trim();
    const sheetId = document.getElementById('sheetId').value.trim();

    if (!apiKey || !sheetId) {
        showError(tabName, '请输入API密钥和表格ID');
        return;
    }

    showLoading(tabName);

    try {
        // Auto-detect range for the specific sheet
        const sheetRange = await detectSheetRange(apiKey, sheetId, sheetNames[tabName]);
        await loadDataWithRange(tabName, apiKey, sheetId, sheetRange);
        
    } catch (error) {
        console.error(`加载${tabName}数据时出错:`, error);
        let errorMessage = `加载${tabName}数据失败: ` + error.message;
        
        if (error.message.includes('403')) {
            errorMessage += '<br><br>可能的解决方案:<br>1. 检查API密钥是否正确<br>2. 确保已启用Google Sheets API<br>3. 检查API密钥的权限设置';
        } else if (error.message.includes('400')) {
            errorMessage += '<br><br>可能的解决方案:<br>1. 检查表格ID是否正确<br>2. 确保表格是公开可访问的<br>3. 检查工作表名称是否正确';
        } else if (error.message.includes('404')) {
            errorMessage += '<br><br>表格未找到，请检查表格ID是否正确';
        } else if (error.message.includes('未找到名为')) {
            errorMessage += `<br><br>请确保表格中包含名为 "${sheetNames[tabName]}" 的工作表`;
        }
        
        showError(tabName, errorMessage);
    }
}

// Load data with specific range
async function loadDataWithRange(tabName, apiKey, sheetId, sheetRange) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${sheetRange}?key=${apiKey}`;
    const response = await fetch(url);
    
    if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const result = await response.json();

    if (!result.values || result.values.length === 0) {
        throw new Error(`工作表 "${sheetNames[tabName]}" 中没有找到数据`);
    }

    // Filter out completely empty rows
    const filteredData = result.values.filter(row => 
        row && row.some(cell => cell !== undefined && cell !== null && cell.toString().trim() !== '')
    );

    if (filteredData.length === 0) {
        throw new Error(`工作表 "${sheetNames[tabName]}" 中没有找到有效数据`);
    }

    currentData[tabName] = filteredData;
    filteredData[tabName] = [...filteredData];
    
    generateStats(tabName, filteredData);
    renderTable(tabName, filteredData);

    document.getElementById(tabName + 'Loading').style.display = 'none';
    document.getElementById(tabName + 'DataContainer').style.display = 'block';
    
    console.log(`${tabName} 数据加载成功 (工作表: ${sheetNames[tabName]}):`, filteredData.length, '行数据');
}

// Page initialization
document.addEventListener('DOMContentLoaded', function() {
    loadConfig();
    
    // Add keyboard shortcuts
    document.addEventListener('keypress', function(e) {
        if (e.key === 'Enter' && !isAuthenticated) {
            authenticateAndLoad();
        }
    });

    // Add filter enter key events for each tab
    ['case', 'talent', 'matching'].forEach(tab => {
        const filterIds = {
            case: ['caseProcessedFilter', 'caseProjectFilter', 'caseContactFilter', 'caseDateFilter'],
            talent: ['talentSkillFilter', 'talentLocationFilter', 'talentExperienceFilter', 'talentNameFilter'],
            matching: ['matchingTalentFilter', 'matchingScoreFilter', 'matchingDateFilter']
        };
        
        filterIds[tab].forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.addEventListener('keypress', function(e) {
                    if (e.key === 'Enter') {
                        applyFilters(tab);
                    }
                });
            }
        });
    });
});