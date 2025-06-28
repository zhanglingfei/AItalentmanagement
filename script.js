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
let currentLanguage = 'zh';

// Language translations
const translations = {
    zh: {
        title: "📊 TERABOX AI人才案件匹配管理数据查看器",
        subtitle: "项目进度追踪与工作记录管理系统",
        apiKeyLabel: "Google API 密钥:",
        apiKeyPlaceholder: "输入您的Google Sheets API密钥",
        sheetIdLabel: "Google Sheets表格ID:",
        sheetIdPlaceholder: "包含所有数据的Google Sheets表格ID",
        loadBtn: "🔄 认证并加载数据",
        caseTab: "📋 案件数据库",
        talentTab: "👥 人才数据库",
        matchingTab: "🎯 匹配结果一览",
        caseLoadingText: "正在加载案件数据...",
        talentLoadingText: "正在加载人才数据...",
        matchingLoadingText: "正在加载匹配结果...",
        caseProcessedFilterLabel: "处理状态:",
        caseProjectFilterLabel: "项目ID:",
        caseContactFilterLabel: "联系人:",
        caseDateFilterLabel: "日期范围:",
        talentSkillFilterLabel: "技术技能:",
        talentLocationFilterLabel: "工作地点:",
        talentExperienceFilterLabel: "经验年限:",
        talentNameFilterLabel: "姓名:",
        matchingTalentFilterLabel: "人才ID:",
        matchingScoreFilterLabel: "最低匹配分数:",
        matchingDateFilterLabel: "匹配日期:",
        allOption: "全部",
        processedOption: "已处理",
        unprocessedOption: "未处理",
        exp02Option: "0-2年",
        exp35Option: "3-5年",
        exp610Option: "6-10年",
        exp10PlusOption: "10年以上",
        caseProjectFilterPlaceholder: "搜索项目ID...",
        caseContactFilterPlaceholder: "搜索联系人...",
        talentSkillFilterPlaceholder: "搜索技能...",
        talentLocationFilterPlaceholder: "搜索地点...",
        talentNameFilterPlaceholder: "搜索姓名...",
        matchingTalentFilterPlaceholder: "搜索人才ID...",
        matchingScoreFilterPlaceholder: "例如: 80",
        filterBtn: "🔍 筛选",
        clearBtn: "🔄 清除",
        caseTableTitle: "📈 案件数据表格",
        talentTableTitle: "👥 人才数据表格",
        matchingTableTitle: "🎯 匹配结果表格",
        exportBtn: "📥 导出CSV",
        refreshBtn: "🔄 刷新数据",
        paginationInfo: "显示第 {start}-{end} 条，共 {total} 条记录",
        perPageLabel: "每页:",
        pageLabel: "第",
        pageUnitLabel: "页",
        firstPageBtn: "首页",
        prevPageBtn: "上一页",
        nextPageBtn: "下一页",
        lastPageBtn: "末页",
        totalProjects: "总项目数",
        processedProjects: "已处理项目",
        completionRate: "完成率",
        avgWorkRate: "平均工作费率",
        totalTalents: "总人才数",
        skillTypes: "技能种类",
        seniorTalents: "资深人才 (10+年)",
        juniorTalents: "新手人才 (0-2年)",
        totalMatches: "总匹配记录",
        highMatches: "高分匹配 (≥80)",
        mediumMatches: "中等匹配 (60-79)",
        avgMatchScore: "平均匹配分数"
    },
    ja: {
        title: "📊 TERABOX AI人材案件マッチング管理データビューア",
        subtitle: "プロジェクト進捗追跡と作業記録管理システム",
        apiKeyLabel: "Google API キー:",
        apiKeyPlaceholder: "Google Sheets APIキーを入力してください",
        sheetIdLabel: "Google SheetsテーブルID:",
        sheetIdPlaceholder: "すべてのデータを含むGoogle SheetsテーブルID",
        loadBtn: "🔄 認証してデータを読み込む",
        caseTab: "📋 案件データベース",
        talentTab: "👥 人材データベース",
        matchingTab: "🎯 マッチング結果一覧",
        caseLoadingText: "案件データを読み込み中...",
        talentLoadingText: "人材データを読み込み中...",
        matchingLoadingText: "マッチング結果を読み込み中...",
        caseProcessedFilterLabel: "処理状況:",
        caseProjectFilterLabel: "プロジェクトID:",
        caseContactFilterLabel: "連絡先:",
        caseDateFilterLabel: "日付範囲:",
        talentSkillFilterLabel: "技術スキル:",
        talentLocationFilterLabel: "勤務地:",
        talentExperienceFilterLabel: "経験年数:",
        talentNameFilterLabel: "氏名:",
        matchingTalentFilterLabel: "人材ID:",
        matchingScoreFilterLabel: "最低マッチングスコア:",
        matchingDateFilterLabel: "マッチング日:",
        allOption: "すべて",
        processedOption: "処理済み",
        unprocessedOption: "未処理",
        exp02Option: "0-2年",
        exp35Option: "3-5年",
        exp610Option: "6-10年",
        exp10PlusOption: "10年以上",
        caseProjectFilterPlaceholder: "プロジェクトIDを検索...",
        caseContactFilterPlaceholder: "連絡先を検索...",
        talentSkillFilterPlaceholder: "スキルを検索...",
        talentLocationFilterPlaceholder: "場所を検索...",
        talentNameFilterPlaceholder: "氏名を検索...",
        matchingTalentFilterPlaceholder: "人材IDを検索...",
        matchingScoreFilterPlaceholder: "例: 80",
        filterBtn: "🔍 フィルター",
        clearBtn: "🔄 クリア",
        caseTableTitle: "📈 案件データテーブル",
        talentTableTitle: "👥 人材データテーブル",
        matchingTableTitle: "🎯 マッチング結果テーブル",
        exportBtn: "📥 CSVエクスポート",
        refreshBtn: "🔄 データを更新",
        paginationInfo: "{start}-{end} 件を表示、全 {total} 件",
        perPageLabel: "1ページあたり:",
        pageLabel: "",
        pageUnitLabel: "ページ目",
        firstPageBtn: "最初",
        prevPageBtn: "前へ",
        nextPageBtn: "次へ",
        lastPageBtn: "最後",
        totalProjects: "総プロジェクト数",
        processedProjects: "処理済みプロジェクト",
        completionRate: "完了率",
        avgWorkRate: "平均作業レート",
        totalTalents: "総人材数",
        skillTypes: "スキルの種類",
        seniorTalents: "シニア人材 (10+年)",
        juniorTalents: "ジュニア人材 (0-2年)",
        totalMatches: "総マッチング記録",
        highMatches: "高スコアマッチング (≥80)",
        mediumMatches: "中スコアマッチング (60-79)",
        avgMatchScore: "平均マッチングスコア"
    }
};

// Updated column mappings for different tabs
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
        '邮件接收时间': 1,
        'PDF 文件名': 2,
        'Resume ID': 3,
        '候选人全名': 4,
        '职业头衔': 5,
        '工作经验年限': 6,
        '技术技能': 7,
        '证书': 8,
        '教育背景': 9,
        '先前雇主': 10,
        '主要项目': 11,
        '专业领域': 12,
        '期望薪资(单价或单金)': 13,
        '可开始时间': 14,
        '工作地点偏好': 15,
        '工作方式偏好': 16,
        '语言能力': 17,
        '联系方式': 18,
        '自我介绍摘要': 19
    },
    matching: {
        'Match_Date': 0,
        'Talent_ID': 1,
        'Talent_Name': 2,
        'Match 1 ID': 3,
        'Match 1 Score': 4,
        'Match 1 Name': 5,
        'Match 2 ID': 6,
        'Match 2 Score': 7,
        'Match 2 Name': 8,
        'Match 3 ID': 9,
        'Match 3 Score': 10,
        'Match 3 Name': 11,
        'Match Details': 12
    }
};

// Updated predefined headers for each tab (display headers)
const predefinedHeaders = {
    case: [
        'Processed At', 'Email Timestamp', 'Email Subject', 'Project ID',
        'Project Title', 'Project Description', 'Work Rate', 'Start Time',
        'Contact Person', 'Contact Information', 'Project Type',
        'Tech Requirements', 'Duration', 'Work Style', 'Client Company',
        'Urgency', 'Special Requirements'
    ],
    talent: [
        '分析日期', '邮件接收时间', 'PDF 文件名', 'Resume ID', '候选人全名', '职业头衔',
        '工作经验年限', '技术技能', '证书', '教育背景', '先前雇主',
        '主要项目', '专业领域', '期望薪资(单价或单金)', '可开始时间', '工作地点偏好', 
        '工作方式偏好', '语言能力', '联系方式', '自我介绍摘要'
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

// Language switching functionality
function switchLanguage(lang) {
    currentLanguage = lang;
    localStorage.setItem('preferredLanguage', lang);
    
    // Update button states
    document.querySelectorAll('.lang-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById('lang' + lang.charAt(0).toUpperCase() + lang.slice(1)).classList.add('active');
    
    // Update all translatable elements
    updateLanguageContent();
}

function updateLanguageContent() {
    const elements = document.querySelectorAll('[data-key]');
    elements.forEach(element => {
        const key = element.getAttribute('data-key');
        if (translations[currentLanguage] && translations[currentLanguage][key]) {
            if (element.tagName === 'INPUT' && element.type !== 'button') {
                element.placeholder = translations[currentLanguage][key];
            } else {
                element.textContent = translations[currentLanguage][key];
            }
        }
    });
    
    // Update pagination info if visible
    ['case', 'talent', 'matching'].forEach(tabName => {
        const paginationInfo = document.getElementById(tabName + 'PaginationInfo');
        if (paginationInfo && paginationInfo.style.display !== 'none') {
            updatePaginationLanguage(tabName);
        }
    });
}

function updatePaginationLanguage(tabName) {
    const dataToUse = filteredData[tabName].length > 0 ? filteredData[tabName] : currentData[tabName];
    if (dataToUse.length === 0) return;
    
    const hasHeader = dataToUse[0] && (
        dataToUse[0].join('').toLowerCase().includes('processed') ||
        dataToUse[0].join('').toLowerCase().includes('resume') ||
        dataToUse[0].join('').toLowerCase().includes('match')
    );
    const dataRows = hasHeader ? dataToUse.slice(1) : dataToUse;
    const totalRecords = dataRows.length;
    const pageSize = getPageSize(tabName);
    const currentPageNum = currentPage[tabName];
    
    const startRecord = totalRecords > 0 ? (currentPageNum - 1) * pageSize + 1 : 0;
    const endRecord = Math.min(currentPageNum * pageSize, totalRecords);
    
    const paginationInfo = document.getElementById(tabName + 'PaginationInfo');
    const template = translations[currentLanguage].paginationInfo;
    paginationInfo.textContent = template
        .replace('{start}', startRecord)
        .replace('{end}', endRecord)
        .replace('{total}', totalRecords);
}

// Load configuration from localStorage
function loadConfig() {
    const apiKey = localStorage.getItem('projectMgmtApiKey');
    const sheetId = localStorage.getItem('projectMgmtSheetId');
    const savedLanguage = localStorage.getItem('preferredLanguage') || 'zh';

    if (apiKey) document.getElementById('apiKey').value = apiKey;
    if (sheetId) document.getElementById('sheetId').value = sheetId;
    
    // Set saved language
    switchLanguage(savedLanguage);
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
                <div class="label">${translations[currentLanguage].totalProjects}</div>
            </div>
            <div class="stat-card">
                <div class="number">${processedCount}</div>
                <div class="label">${translations[currentLanguage].processedProjects}</div>
            </div>
            <div class="stat-card">
                <div class="number">${completionRate}%</div>
                <div class="label">${translations[currentLanguage].completionRate}</div>
            </div>
            <div class="stat-card">
                <div class="number">¥${avgWorkRate}</div>
                <div class="label">${translations[currentLanguage].avgWorkRate}</div>
            </div>
        `;
    } else if (tabName === 'talent') {
        let expCounts = { '0-2': 0, '3-5': 0, '6-10': 0, '10+': 0 };
        let skillSet = new Set();
        
        for (let i = startIndex; i < data.length; i++) {
            const experience = data[i][6] || '';
            const skills = data[i][7] || '';
            
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
                <div class="label">${translations[currentLanguage].totalTalents}</div>
            </div>
            <div class="stat-card">
                <div class="number">${skillSet.size}</div>
                <div class="label">${translations[currentLanguage].skillTypes}</div>
            </div>
            <div class="stat-card">
                <div class="number">${expCounts['10+']}</div>
                <div class="label">${translations[currentLanguage].seniorTalents}</div>
            </div>
            <div class="stat-card">
                <div class="number">${expCounts['0-2']}</div>
                <div class="label">${translations[currentLanguage].juniorTalents}</div>
            </div>
        `;
    } else if (tabName === 'matching') {
        let highMatches = 0;
        let mediumMatches = 0;
        let lowMatches = 0;
        let totalScore = 0;
        let validScores = 0;
        
        for (let i = startIndex; i < data.length; i++) {
            const score1 = parseFloat(data[i][4]) || 0;  // Match 1 Score
            const score2 = parseFloat(data[i][7]) || 0;  // Match 2 Score
            const score3 = parseFloat(data[i][10]) || 0; // Match 3 Score
            
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
                <div class="label">${translations[currentLanguage].totalMatches}</div>
            </div>
            <div class="stat-card">
                <div class="number">${highMatches}</div>
                <div class="label">${translations[currentLanguage].highMatches}</div>
            </div>
            <div class="stat-card">
                <div class="number">${mediumMatches}</div>
                <div class="label">${translations[currentLanguage].mediumMatches}</div>
            </div>
            <div class="stat-card">
                <div class="number">${avgScore}</div>
                <div class="label">${translations[currentLanguage].avgMatchScore}</div>
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
            case 0:
            case 1: // Analysis Date, Email Timestamp
                return `<span class="timestamp">${cellValue}</span>`;
            case 3: // Resume ID
                return `<span class="talent-id">${cellValue}</span>`;
            case 4: // Full Name
                return `<span class="project-title">${cellValue}</span>`;
            case 6: // Experience Years
                return `<span class="experience-years">${cellValue}年</span>`;
            case 7: // Skills
                const skills = cellValue.split(/[,，、]/).slice(0, 5);
                return `<div class="skill-tags">${skills.map(skill => 
                    `<span class="skill-tag">${skill.trim()}</span>`
                ).join('')}</div>`;
            case 18: // Contact Info
                return `<span class="email-cell">${cellValue}</span>`;
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
            case 7: // Match 1/2/3 Score
                const score = parseFloat(cellValue) || 0;
                let scoreClass = 'low';
                if (score >= 80) scoreClass = 'high';
                else if (score >= 60) scoreClass = 'medium';
                return `<span class="match-score ${scoreClass}">${cellValue}</span>`;
            case 4:
            case 6:
            case 8: // Match 1/2/3 Name
                return `<span class="project-title">${cellValue}</span>`;
            case 9: // Match Details
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
    
    // Define hidden column indices for matching table (Match ID columns)
    const hiddenColumns = tabName === 'matching' ? [3, 6, 9] : []; 
    
    // Apply pagination
    const currentPageNum = currentPage[tabName];
    const pageSize = getPageSize(tabName);
    const startIndex = (currentPageNum - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const paginatedData = dataRows.slice(startIndex, endIndex);

    let html = '';
    
    // Use predefined headers
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
        
        // Update pagination info with current language
        const template = translations[currentLanguage].paginationInfo;
        paginationInfo.textContent = template
            .replace('{start}', startRecord)
            .replace('{end}', endRecord)
            .replace('{total}', totalRecords);
            
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

            if (contact