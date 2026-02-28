// ==================== CONFIGURATION ====================
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyuzXL4L5fP0s1N7uTg-cp8JZidkVVy6fXQncIvO83Cjfq3OEy4zNlmJEPeZcivQJMl/exec';
const PAGE_SIZE = 50;

// ==================== GLOBAL STATE ====================
let currentUser = null;
let currentMemberPage = 1;
let totalMembers = 0;
let currentMasulPage = 1;
let totalMasuls = 0;
let currentMemberSearch = '';
let currentMasulSearch = '';
let currentMemberFilters = {};
let currentMasulFilters = {};
let branchZoneMap = {};

// Store the last viewed member/masul data for printing
let lastViewedMember = null;
let lastViewedMasul = null;

// ==================== SURAT AL-ASR TYPING ANIMATION (with Arabic numerals) ====================
function typeSurahAsr() {
    const surahElement = document.getElementById('surahText');
    if (!surahElement) return;
    // Arabic text with Arabic-Indic numerals (١, ٢, ٣)
    const fullText = "وَٱلْعَصْرِ (١) إِنَّ ٱلْإِنسَـٰنَ لَفِى خُسْرٍ (٢) إِلَّا ٱلَّذِينَ ءَامَنُوا۟ وَعَمِلُوا۟ ٱلصَّـٰلِحَـٰتِ وَتَوَاصَوْا۟ بِٱلْحَقِّ وَتَوَاصَوْا۟ بِٱلصَّبْرِ (٣)";
    let index = 0;
    surahElement.innerHTML = '';
    function typeNext() {
        if (index < fullText.length) {
            surahElement.innerHTML += fullText.charAt(index);
            index++;
            setTimeout(typeNext, 50);
        } else {
            setTimeout(hidePreloader, 500);
        }
    }
    typeNext();
}

function hidePreloader() {
    const preloader = document.getElementById('surah-preloader');
    const pageContent = document.getElementById('page-content');
    if (preloader) {
        preloader.classList.add('fade-out');
        setTimeout(() => {
            preloader.style.display = 'none';
            if (pageContent) pageContent.style.display = 'block';
        }, 500);
    }
}

// ==================== LOADER ====================
function showLoader() { 
    const loader = document.getElementById('globalLoader');
    if (loader) loader.style.display = 'flex'; 
}
function hideLoader() { 
    const loader = document.getElementById('globalLoader');
    if (loader) loader.style.display = 'none'; 
}

// ==================== CUSTOM MODALS ====================
function showMessage(title, text) {
    document.getElementById('messageModalTitle').innerText = title;
    document.getElementById('messageModalText').innerText = text;
    document.getElementById('messageModal').style.display = 'block';
}
function closeMessageModal() {
    document.getElementById('messageModal').style.display = 'none';
}

function showConfirm(title, text) {
    return new Promise((resolve) => {
        document.getElementById('confirmModalTitle').innerText = title;
        document.getElementById('confirmModalText').innerText = text;
        document.getElementById('confirmModal').style.display = 'block';
        document.getElementById('confirmOkBtn').onclick = () => {
            closeConfirmModal();
            resolve(true);
        };
        document.getElementById('confirmCancelBtn').onclick = () => {
            closeConfirmModal();
            resolve(false);
        };
    });
}
function closeConfirmModal() {
    document.getElementById('confirmModal').style.display = 'none';
}

function showPrompt(title, text, defaultValue = '') {
    return new Promise((resolve) => {
        document.getElementById('promptModalTitle').innerText = title;
        document.getElementById('promptModalText').innerText = text;
        document.getElementById('promptInput').value = defaultValue;
        document.getElementById('promptModal').style.display = 'block';
        document.getElementById('promptOkBtn').onclick = () => {
            const val = document.getElementById('promptInput').value;
            closePromptModal();
            resolve(val);
        };
        document.getElementById('promptCancelBtn').onclick = () => {
            closePromptModal();
            resolve(null);
        };
    });
}
function closePromptModal() {
    document.getElementById('promptModal').style.display = 'none';
}

// ==================== API REQUEST ====================
async function apiRequest(action, data = {}, user = null) {
    showLoader();
    try {
        const payload = { action, ...data };
        if (user) payload.user = user;
        const formBody = new URLSearchParams();
        formBody.append('payload', JSON.stringify(payload));
        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: formBody.toString()
        });
        const result = await response.json();
        if (!result.success) {
            throw new Error(result.error || 'Unknown error');
        }
        return result;
    } finally {
        hideLoader();
    }
}

function showModal(modalId) {
    document.getElementById(modalId).style.display = 'block';
}

function hideModal(modalId) {
    document.getElementById(modalId).style.display = 'none';
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    });
}

// ==================== DEBOUNCE HELPER ====================
function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// ==================== SIDEBAR TOGGLE ====================
function initSidebar() {
    const sidebar = document.getElementById('sidebar');
    const toggleBtn = document.getElementById('toggleSidebar');
    if (!sidebar || !toggleBtn) return;
    toggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (window.innerWidth <= 768) {
            sidebar.classList.toggle('mobile-open');
            document.body.style.overflow = sidebar.classList.contains('mobile-open') ? 'hidden' : '';
        } else {
            sidebar.classList.toggle('collapsed');
        }
    });
    document.addEventListener('click', (e) => {
        if (window.innerWidth <= 768 && sidebar.classList.contains('mobile-open')) {
            if (!sidebar.contains(e.target) && !toggleBtn.contains(e.target)) {
                sidebar.classList.remove('mobile-open');
                document.body.style.overflow = '';
            }
        }
    });
    window.addEventListener('resize', () => {
        if (window.innerWidth > 768) {
            sidebar.classList.remove('mobile-open');
            document.body.style.overflow = '';
        } else {
            sidebar.classList.remove('collapsed');
        }
    });
}

// ==================== LOGIN & INIT ====================
document.addEventListener('DOMContentLoaded', () => {
    const style = document.createElement('style');
    style.innerHTML = `
        .modal-content .print-area { position: relative; }
        .modal-content .print-area::before {
            content: "";
            position: absolute;
            top: 0; left: 0; right: 0; bottom: 0;
            background-image: url('logo.png');
            background-repeat: no-repeat;
            background-position: center;
            background-size: 200px;
            opacity: 0.1;
            pointer-events: none;
            z-index: -1;
        }
        body::after {
            content: "";
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background-image: url('logo.png');
            background-repeat: no-repeat;
            background-position: center;
            background-size: 300px;
            opacity: 0.05;
            pointer-events: none;
            z-index: -1;
        }
        .sidebar {
            height: 100vh;
            overflow-y: auto;
            position: sticky;
            top: 0;
        }
    `;
    document.head.appendChild(style);

    if (window.location.pathname.includes('index.html') || window.location.pathname === '/') {
        typeSurahAsr();
    } else {
        hidePreloader();
    }

    const savedUser = sessionStorage.getItem('user');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        if (window.location.pathname.includes('dashboard.html')) {
            initSidebar();
            initializeDashboard();
        } else if (window.location.pathname.includes('registration.html')) {
            initializeRegistrationPage();
        }
    } else {
        if (!window.location.pathname.includes('index.html')) {
            window.location.href = 'index.html';
        }
    }

    const loginLink = document.getElementById('loginLink');
    if (loginLink) {
        loginLink.addEventListener('click', (e) => {
            e.preventDefault();
            showModal('loginModal');
        });
    }

    const loginModal = document.getElementById('loginModal');
    if (loginModal) {
        const span = loginModal.querySelector('.close');
        span.onclick = () => hideModal('loginModal');
        window.onclick = (event) => {
            if (event.target.classList.contains('modal-overlay') || event.target.classList.contains('modal')) {
                event.target.closest('.modal').style.display = 'none';
            }
        };
        document.getElementById('loginForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const role = document.getElementById('role').value;
            const code = document.getElementById('accessCode').value;
            try {
                const result = await apiRequest('login', { role, code });
                currentUser = result.user;
                sessionStorage.setItem('user', JSON.stringify(currentUser));
                window.location.href = 'dashboard.html';
            } catch (err) {
                showMessage('Login Failed', err.message);
            }
        });
    }

    const logoutLink = document.getElementById('logoutLink');
    if (logoutLink) {
        logoutLink.addEventListener('click', (e) => {
            e.preventDefault();
            sessionStorage.removeItem('user');
            window.location.href = 'index.html';
        });
    }
});

// ==================== DASHBOARD INIT ====================
async function initializeDashboard() {
    if (!currentUser) return;
    document.getElementById('roleDisplay').innerText = currentUser.role;

    if (currentUser.role !== 'Admin') {
        const adminSections = [
            'masulSection', 'zonesSection', 'branchesSection',
            'auditSection', 'configSection', 'exportSection'
        ];
        adminSections.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = 'none';
        });
    } else {
        document.querySelectorAll('.admin-only').forEach(el => el.style.display = 'block');
    }

    if (currentUser.role === 'Zonal Mas\'ul') {
        document.querySelectorAll('.zonal-only').forEach(el => el.style.display = 'block');
    } else if (currentUser.role === 'Branch Mas\'ul') {
        document.querySelectorAll('.branch-only').forEach(el => el.style.display = 'block');
    }

    setupNavigation();
    applyRoleBasedVisibility();
    showSection('membersSection');
    await loadDashboardStats();
    await loadMemberList(1, '');
    await loadFilterOptions();
    loadZonesForDropdowns();

    const memberListSearch = document.getElementById('memberListSearch');
    if (memberListSearch) {
        memberListSearch.addEventListener('input', debounce(applyMemberFilters, 300));
    }
    const masulSearch = document.getElementById('masulSearch');
    if (masulSearch) {
        masulSearch.addEventListener('input', debounce(applyMasulFilters, 300));
    }
}

function applyRoleBasedVisibility() {
    const role = currentUser.role;
    const zoneChartSection = document.getElementById('zoneStatsSection');
    const branchChartSection = document.getElementById('branchStatsSection');
    if (role === 'Branch Mas\'ul') {
        if (zoneChartSection) zoneChartSection.style.display = 'none';
        if (branchChartSection) branchChartSection.style.display = 'none';
    } else {
        if (zoneChartSection) zoneChartSection.style.display = 'block';
        if (branchChartSection) branchChartSection.style.display = 'block';
    }
}

function setupNavigation() {
    document.getElementById('navMembers').addEventListener('click', (e) => {
        e.preventDefault();
        showSection('membersSection');
        loadDashboardStats();
    });
    document.getElementById('navMemberList').addEventListener('click', (e) => {
        e.preventDefault();
        showSection('memberListSection');
        document.getElementById('memberListSearch').value = '';
        resetMemberFilters();
        loadMemberList(1, '');
    });
    const navMasulin = document.getElementById('navMasulin');
    if (navMasulin) {
        navMasulin.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('masulSection');
            document.getElementById('masulSearch').value = '';
            resetMasulFilters();
            loadMasuls(1, '');
        });
    }
    const navZones = document.getElementById('navZones');
    if (navZones) {
        navZones.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('zonesSection');
            loadZones();
        });
    }
    const navBranches = document.getElementById('navBranches');
    if (navBranches) {
        navBranches.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('branchesSection');
            loadBranches();
        });
    }
    const navAudit = document.getElementById('navAudit');
    if (navAudit) {
        navAudit.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('auditSection');
            loadAuditLog();
        });
    }
    const navConfig = document.getElementById('navConfig');
    if (navConfig) {
        navConfig.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('configSection');
            loadConfig();
        });
    }
    const navExport = document.getElementById('navExport');
    if (navExport) {
        navExport.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('exportSection');
        });
    }
    const navZoneStats = document.getElementById('navZoneStats');
    if (navZoneStats) {
        navZoneStats.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('zoneStatsSection');
            loadZoneStats();
        });
    }
    const navBranchStats = document.getElementById('navBranchStats');
    if (navBranchStats) {
        navBranchStats.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('branchStatsSection');
            loadBranchStats();
        });
    }
}

function showSection(sectionId) {
    const sections = [
        'membersSection', 'memberListSection', 'masulSection', 'zonesSection',
        'branchesSection', 'auditSection', 'configSection', 'exportSection',
        'zoneStatsSection', 'branchStatsSection'
    ];
    sections.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
    document.getElementById(sectionId).style.display = 'block';

    document.querySelectorAll('.sidebar-menu a').forEach(a => a.classList.remove('active'));
    const navMap = {
        membersSection: 'navMembers',
        memberListSection: 'navMemberList',
        masulSection: 'navMasulin',
        zonesSection: 'navZones',
        branchesSection: 'navBranches',
        auditSection: 'navAudit',
        configSection: 'navConfig',
        exportSection: 'navExport',
        zoneStatsSection: 'navZoneStats',
        branchStatsSection: 'navBranchStats'
    };
    const navId = navMap[sectionId];
    if (navId) {
        const navLink = document.getElementById(navId);
        if (navLink) navLink.classList.add('active');
    }
}

// ==================== FILTER OPTIONS ====================
async function loadFilterOptions() {
    try {
        const result = await apiRequest('getFilterOptions', {}, currentUser);
        populateSelect('filterMemberLevel', result.levels);
        populateSelect('filterMemberBranch', result.branches);
        populateSelect('filterMemberZone', result.zones);
        populateSelect('filterMasulRank', result.ranks);
        populateSelect('filterMasulBranch', result.branches);
        populateSelect('filterMasulZone', result.zones);
    } catch (err) {
        console.error('Failed to load filter options:', err);
    }
}

function populateSelect(selectId, options) {
    const select = document.getElementById(selectId);
    if (!select) return;
    const firstOption = select.options[0] ? select.options[0].cloneNode(true) : null;
    select.innerHTML = '';
    if (firstOption) select.appendChild(firstOption);
    options.forEach(val => {
        const option = document.createElement('option');
        option.value = val;
        option.textContent = val;
        select.appendChild(option);
    });
}

// ==================== MEMBERS LIST ====================
async function loadMemberList(page = 1, search = '', filters = {}) {
    try {
        currentMemberPage = page;
        currentMemberSearch = search;
        currentMemberFilters = filters;
        const result = await apiRequest('getMembers', { page, pageSize: PAGE_SIZE, search, filters }, currentUser);
        totalMembers = result.total;
        renderMemberListTable(result.members);
        renderMemberListPagination();
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load members: ' + err.message);
    }
}

function renderMemberListTable(members) {
    const tbody = document.querySelector('#memberListTable tbody');
    tbody.innerHTML = '';
    members.forEach(member => {
        const row = tbody.insertRow();
        row.insertCell().innerText = member.IntizarID;
        row.insertCell().innerText = member.RecruitmentID;
        row.insertCell().innerText = member.FullName;
        row.insertCell().innerText = member.Level;
        row.insertCell().innerText = member.Branch;
        const actions = row.insertCell();
        actions.innerHTML = `
            <button onclick="viewMember('${member.IntizarID}')">👁 View</button>
            <button onclick="viewMember('${member.IntizarID}')">🖨 Print</button>
            ${currentUser.role === 'Admin' || currentUser.role === 'Zonal Mas\'ul' ? `<button onclick="promoteMember('${member.IntizarID}')">⭐ Promote</button>` : ''}
            ${currentUser.role === 'Admin' ? `<button onclick="transferMember('${member.IntizarID}')">↗ Transfer</button>` : ''}
        `;
    });
}

function renderMemberListPagination() {
    const totalPages = Math.ceil(totalMembers / PAGE_SIZE);
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
        html += `<button class="page-btn ${i === currentMemberPage ? 'active' : ''}" onclick="loadMemberList(${i}, '${currentMemberSearch}', ${JSON.stringify(currentMemberFilters).replace(/"/g, '&quot;')})">${i}</button>`;
    }
    html += `<span> Total: ${totalMembers}</span>`;
    document.getElementById('memberListPagination').innerHTML = html;
}

function applyMemberFilters() {
    const filters = {
        level: document.getElementById('filterMemberLevel').value,
        gender: document.getElementById('filterMemberGender').value,
        branch: document.getElementById('filterMemberBranch').value,
        zone: document.getElementById('filterMemberZone').value
    };
    const search = document.getElementById('memberListSearch').value;
    loadMemberList(1, search, filters);
}

function resetMemberFilters() {
    document.getElementById('filterMemberLevel').value = '';
    document.getElementById('filterMemberGender').value = '';
    document.getElementById('filterMemberBranch').value = '';
    document.getElementById('filterMemberZone').value = '';
    applyMemberFilters();
}

// ==================== MASULS LIST ====================
async function loadMasuls(page = 1, search = '', filters = {}) {
    try {
        currentMasulPage = page;
        currentMasulSearch = search;
        currentMasulFilters = filters;
        const result = await apiRequest('getMasuls', { page, pageSize: PAGE_SIZE, search, filters }, currentUser);
        totalMasuls = result.total;
        renderMasulTable(result.masuls);
        renderMasulPagination();
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load masuls: ' + err.message);
    }
}

function renderMasulTable(masuls) {
    const tbody = document.querySelector('#masulTable tbody');
    tbody.innerHTML = '';
    masuls.forEach(masul => {
        const row = tbody.insertRow();
        row.insertCell().innerText = masul.IntizarID;
        row.insertCell().innerText = masul.MasulRecruitmentID;
        row.insertCell().innerText = masul.FullName;
        row.insertCell().innerText = masul.CurrentRank;
        row.insertCell().innerText = masul.Branch;
        const actions = row.insertCell();
        actions.innerHTML = `
            <button onclick="viewMasul('${masul.IntizarID}')">👁 View</button>
            <button onclick="viewMasul('${masul.IntizarID}')">🖨 Print</button>
            <button onclick="promoteMasul('${masul.IntizarID}')">⭐ Promote</button>
            <button onclick="transferMasul('${masul.IntizarID}')">↗ Transfer</button>
        `;
    });
}

function renderMasulPagination() {
    const totalPages = Math.ceil(totalMasuls / PAGE_SIZE);
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
        html += `<button class="page-btn ${i === currentMasulPage ? 'active' : ''}" onclick="loadMasuls(${i}, '${currentMasulSearch}', ${JSON.stringify(currentMasulFilters).replace(/"/g, '&quot;')})">${i}</button>`;
    }
    html += `<span> Total: ${totalMasuls}</span>`;
    document.getElementById('masulPagination').innerHTML = html;
}

function applyMasulFilters() {
    const filters = {
        rank: document.getElementById('filterMasulRank').value,
        gender: document.getElementById('filterMasulGender').value,
        branch: document.getElementById('filterMasulBranch').value,
        zone: document.getElementById('filterMasulZone').value
    };
    const search = document.getElementById('masulSearch').value;
    loadMasuls(1, search, filters);
}

function resetMasulFilters() {
    document.getElementById('filterMasulRank').value = '';
    document.getElementById('filterMasulGender').value = '';
    document.getElementById('filterMasulBranch').value = '';
    document.getElementById('filterMasulZone').value = '';
    applyMasulFilters();
}

// ==================== VIEW MEMBER (with image fallback and Print button) ====================
async function viewMember(intizarId) {
    try {
        const result = await apiRequest('getMember', { intizarId }, currentUser);
        const member = result.member;
        lastViewedMember = member; // store for printing

        let promotionList = '';
        try {
            const promHistory = JSON.parse(member.PromotionHistory || '[]');
            if (promHistory.length) {
                promotionList = '<ul>' + promHistory.map(entry => 
                    `<li>${new Date(entry.date).toLocaleDateString()}: ${entry.action || 'Promoted to ' + entry.level}</li>`
                ).join('') + '</ul>';
            } else {
                promotionList = '<p>No promotion history</p>';
            }
        } catch (e) {
            promotionList = '<p>Error parsing history</p>';
        }
        let transferList = '';
        try {
            const transHistory = JSON.parse(member.TransferHistory || '[]');
            if (transHistory.length) {
                transferList = '<ul>' + transHistory.map(entry => 
                    `<li>${new Date(entry.date).toLocaleDateString()}: from ${entry.fromBranch} to ${entry.toBranch}</li>`
                ).join('') + '</ul>';
            } else {
                transferList = '<p>No transfer history</p>';
            }
        } catch (e) {
            transferList = '<p>Error parsing transfers</p>';
        }
        const photoHtml = member.PhotoURL 
            ? `<img src="${member.PhotoURL}" alt="Passport" class="print-photo" onerror="tryAlternateImage(this, '${member.PhotoURL}')">` 
            : `<img src="logo.png" alt="Default" class="print-photo">`;

        const content = document.getElementById('viewContent');
        content.innerHTML = `
            <div class="print-area">
                <div class="print-header">
                    <img src="logo.png" alt="Logo" style="height:60px;">
                    <h2>INTIZARUL IMAMUL MUNTAZAR</h2>
                    <p>Member Biodata</p>
                </div>
                ${photoHtml}
                <p><strong>Intizar ID:</strong> ${member.IntizarID}</p>
                <p><strong>Recruitment ID:</strong> ${member.RecruitmentID}</p>
                <p><strong>Full Name:</strong> ${member.FullName}</p>
                <p><strong>Father's Name:</strong> ${member.FatherName}</p>
                <p><strong>Gender:</strong> ${member.Gender}</p>
                <p><strong>Date of Birth:</strong> ${member.DOB}</p>
                <p><strong>Place of Birth:</strong> ${member.PlaceOfBirth}</p>
                <p><strong>Phone:</strong> ${member.Phone}</p>
                <p><strong>Email:</strong> ${member.Email || '-'}</p>
                <p><strong>Address:</strong> ${member.Address}</p>
                <p><strong>State:</strong> ${member.State}</p>
                <p><strong>LGA:</strong> ${member.LGA}</p>
                <p><strong>Zone:</strong> ${member.Zone}</p>
                <p><strong>Branch:</strong> ${member.Branch}</p>
                <p><strong>Year:</strong> ${member.Year}</p>
                <p><strong>Level:</strong> ${member.Level}</p>
                <p><strong>Guardian Name:</strong> ${member.GuardianName}</p>
                <p><strong>Guardian Phone:</strong> ${member.GuardianPhone}</p>
                <p><strong>Guardian Address:</strong> ${member.GuardianAddress}</p>
                <p><strong>Promotion History:</strong> ${promotionList}</p>
                <p><strong>Transfer History:</strong> ${transferList}</p>
                <p><em>Generated on: ${new Date().toLocaleString()}</em></p>
                <div style="text-align: center; margin-top: 20px;">
                    <button onclick="printCurrentMember()" class="no-print">Print This Page</button>
                </div>
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== VIEW MASUL (with image fallback and Print button) ====================
async function viewMasul(intizarId) {
    try {
        const result = await apiRequest('getMasul', { intizarId }, currentUser);
        const masul = result.masul;
        lastViewedMasul = masul; // store for printing

        let promotionList = '';
        try {
            const promHistory = JSON.parse(masul.PromotionHistory || '[]');
            if (promHistory.length) {
                promotionList = '<ul>' + promHistory.map(entry => 
                    `<li>${new Date(entry.date).toLocaleDateString()}: ${entry.action || 'Promoted to ' + entry.rank}</li>`
                ).join('') + '</ul>';
            } else {
                promotionList = '<p>No promotion history</p>';
            }
        } catch (e) {
            promotionList = '<p>Error parsing history</p>';
        }
        const photoHtml = masul.PhotoURL 
            ? `<img src="${masul.PhotoURL}" alt="Passport" class="print-photo" onerror="tryAlternateImage(this, '${masul.PhotoURL}')">` 
            : `<img src="logo.png" alt="Default" class="print-photo">`;

        const content = document.getElementById('viewContent');
        content.innerHTML = `
            <div class="print-area">
                <div class="print-header">
                    <img src="logo.png" alt="Logo" style="height:60px;">
                    <h2>INTIZARUL IMAMUL MUNTAZAR</h2>
                    <p>Mas'ul Biodata</p>
                </div>
                ${photoHtml}
                <p><strong>Intizar ID:</strong> ${masul.IntizarID}</p>
                <p><strong>Mas'ul Recruitment ID:</strong> ${masul.MasulRecruitmentID}</p>
                <p><strong>Full Name:</strong> ${masul.FullName}</p>
                <p><strong>Father's Name:</strong> ${masul.FatherName}</p>
                <p><strong>Gender:</strong> ${masul.Gender}</p>
                <p><strong>Date of Birth:</strong> ${masul.DOB}</p>
                <p><strong>Place of Birth:</strong> ${masul.PlaceOfBirth}</p>
                <p><strong>Phone:</strong> ${masul.Phone}</p>
                <p><strong>Email:</strong> ${masul.Email || '-'}</p>
                <p><strong>Address:</strong> ${masul.Address}</p>
                <p><strong>State:</strong> ${masul.State}</p>
                <p><strong>LGA:</strong> ${masul.LGA}</p>
                <p><strong>Zone:</strong> ${masul.Zone}</p>
                <p><strong>Branch:</strong> ${masul.Branch}</p>
                <p><strong>Year:</strong> ${masul.Year}</p>
                <p><strong>Current Rank:</strong> ${masul.CurrentRank}</p>
                <p><strong>Source:</strong> ${masul.Source}</p>
                ${masul.OriginalMemberRecruitmentID ? `<p><strong>Original Member Recruitment ID:</strong> ${masul.OriginalMemberRecruitmentID}</p>` : ''}
                <p><strong>Promotion History:</strong> ${promotionList}</p>
                <p><em>Generated on: ${new Date().toLocaleString()}</em></p>
                <div style="text-align: center; margin-top: 20px;">
                    <button onclick="printCurrentMasul()" class="no-print">Print This Page</button>
                </div>
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== PRINT FUNCTIONS (use stored data, open new window, wait for images) ====================
function printCurrentMember() {
    if (!lastViewedMember) {
        showMessage('Error', 'No member data to print.');
        return;
    }
    const printContent = buildMemberPrintHTML(lastViewedMember);
    openPrintWindow(printContent, 'Member Biodata');
}

function printCurrentMasul() {
    if (!lastViewedMasul) {
        showMessage('Error', 'No masul data to print.');
        return;
    }
    const printContent = buildMasulPrintHTML(lastViewedMasul);
    openPrintWindow(printContent, 'Mas\'ul Biodata');
}

function openPrintWindow(content, title) {
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
            <head>
                <title>${title}</title>
                <link rel="stylesheet" href="style.css">
                <style>
                    @media print {
                        body { margin: 1cm; }
                        .print-header { text-align: center; }
                        .print-photo { max-width: 150px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
                        button, .no-print { display: none; }
                    }
                </style>
                <script>
                    function tryAlternateImage(img, originalUrl) {
                        var match = originalUrl.match(/[-\\w]{25,}/);
                        if (match) {
                            var fileId = match[0];
                            img.src = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';
                            img.onerror = function() {
                                img.src = 'logo.png';
                                img.onerror = null;
                            };
                        } else {
                            img.src = 'logo.png';
                            img.onerror = null;
                        }
                    }
                    window.onload = function() {
                        var images = document.images;
                        var loaded = 0;
                        if (images.length === 0) {
                            window.print();
                            window.onafterprint = function() { window.close(); };
                            return;
                        }
                        for (var i = 0; i < images.length; i++) {
                            var img = images[i];
                            if (img.complete) {
                                loaded++;
                            } else {
                                img.addEventListener('load', function() {
                                    loaded++;
                                    if (loaded === images.length) {
                                        window.print();
                                        window.onafterprint = function() { window.close(); };
                                    }
                                });
                                img.addEventListener('error', function() {
                                    loaded++;
                                    if (loaded === images.length) {
                                        window.print();
                                        window.onafterprint = function() { window.close(); };
                                    }
                                });
                            }
                        }
                        if (loaded === images.length) {
                            window.print();
                            window.onafterprint = function() { window.close(); };
                        }
                    };
                <\/script>
            </head>
            <body>
                <div class="print-area">${content}</div>
            </body>
        </html>
    `);
    printWindow.document.close();
}

// ==================== HELPERS FOR PRINT HTML ====================
function buildMemberPrintHTML(member) {
    let promotionList = '';
    try {
        const promHistory = JSON.parse(member.PromotionHistory || '[]');
        if (promHistory.length) {
            promotionList = '<ul>' + promHistory.map(entry => 
                `<li>${new Date(entry.date).toLocaleDateString()}: ${entry.action || 'Promoted to ' + entry.level}</li>`
            ).join('') + '</ul>';
        } else {
            promotionList = '<p>No promotion history</p>';
        }
    } catch (e) {
        promotionList = '<p>Error parsing history</p>';
    }
    let transferList = '';
    try {
        const transHistory = JSON.parse(member.TransferHistory || '[]');
        if (transHistory.length) {
            transferList = '<ul>' + transHistory.map(entry => 
                `<li>${new Date(entry.date).toLocaleDateString()}: from ${entry.fromBranch} to ${entry.toBranch}</li>`
            ).join('') + '</ul>';
        } else {
            transferList = '<p>No transfer history</p>';
        }
    } catch (e) {
        transferList = '<p>Error parsing transfers</p>';
    }
    const photoHtml = member.PhotoURL 
        ? `<img src="${member.PhotoURL}" alt="Passport" class="print-photo" onerror="tryAlternateImage(this, '${member.PhotoURL}')">` 
        : `<img src="logo.png" alt="Default" class="print-photo">`;

    return `
        <div class="print-header">
            <img src="logo.png" alt="Logo" style="height:60px;">
            <h2>INTIZARUL IMAMUL MUNTAZAR</h2>
            <p>Member Biodata</p>
        </div>
        ${photoHtml}
        <p><strong>Intizar ID:</strong> ${member.IntizarID}</p>
        <p><strong>Recruitment ID:</strong> ${member.RecruitmentID}</p>
        <p><strong>Full Name:</strong> ${member.FullName}</p>
        <p><strong>Father's Name:</strong> ${member.FatherName}</p>
        <p><strong>Gender:</strong> ${member.Gender}</p>
        <p><strong>Date of Birth:</strong> ${member.DOB}</p>
        <p><strong>Place of Birth:</strong> ${member.PlaceOfBirth}</p>
        <p><strong>Phone:</strong> ${member.Phone}</p>
        <p><strong>Email:</strong> ${member.Email || '-'}</p>
        <p><strong>Address:</strong> ${member.Address}</p>
        <p><strong>State:</strong> ${member.State}</p>
        <p><strong>LGA:</strong> ${member.LGA}</p>
        <p><strong>Zone:</strong> ${member.Zone}</p>
        <p><strong>Branch:</strong> ${member.Branch}</p>
        <p><strong>Year:</strong> ${member.Year}</p>
        <p><strong>Level:</strong> ${member.Level}</p>
        <p><strong>Guardian Name:</strong> ${member.GuardianName}</p>
        <p><strong>Guardian Phone:</strong> ${member.GuardianPhone}</p>
        <p><strong>Guardian Address:</strong> ${member.GuardianAddress}</p>
        <p><strong>Promotion History:</strong> ${promotionList}</p>
        <p><strong>Transfer History:</strong> ${transferList}</p>
        <p><em>Generated on: ${new Date().toLocaleString()}</em></p>
    `;
}

function buildMasulPrintHTML(masul) {
    let promotionList = '';
    try {
        const promHistory = JSON.parse(masul.PromotionHistory || '[]');
        if (promHistory.length) {
            promotionList = '<ul>' + promHistory.map(entry => 
                `<li>${new Date(entry.date).toLocaleDateString()}: ${entry.action || 'Promoted to ' + entry.rank}</li>`
            ).join('') + '</ul>';
        } else {
            promotionList = '<p>No promotion history</p>';
        }
    } catch (e) {
        promotionList = '<p>Error parsing history</p>';
    }
    const photoHtml = masul.PhotoURL 
        ? `<img src="${masul.PhotoURL}" alt="Passport" class="print-photo" onerror="tryAlternateImage(this, '${masul.PhotoURL}')">` 
        : `<img src="logo.png" alt="Default" class="print-photo">`;

    return `
        <div class="print-header">
            <img src="logo.png" alt="Logo" style="height:60px;">
            <h2>INTIZARUL IMAMUL MUNTAZAR</h2>
            <p>Mas'ul Biodata</p>
        </div>
        ${photoHtml}
        <p><strong>Intizar ID:</strong> ${masul.IntizarID}</p>
        <p><strong>Mas'ul Recruitment ID:</strong> ${masul.MasulRecruitmentID}</p>
        <p><strong>Full Name:</strong> ${masul.FullName}</p>
        <p><strong>Father's Name:</strong> ${masul.FatherName}</p>
        <p><strong>Gender:</strong> ${masul.Gender}</p>
        <p><strong>Date of Birth:</strong> ${masul.DOB}</p>
        <p><strong>Place of Birth:</strong> ${masul.PlaceOfBirth}</p>
        <p><strong>Phone:</strong> ${masul.Phone}</p>
        <p><strong>Email:</strong> ${masul.Email || '-'}</p>
        <p><strong>Address:</strong> ${masul.Address}</p>
        <p><strong>State:</strong> ${masul.State}</p>
        <p><strong>LGA:</strong> ${masul.LGA}</p>
        <p><strong>Zone:</strong> ${masul.Zone}</p>
        <p><strong>Branch:</strong> ${masul.Branch}</p>
        <p><strong>Year:</strong> ${masul.Year}</p>
        <p><strong>Current Rank:</strong> ${masul.CurrentRank}</p>
        <p><strong>Source:</strong> ${masul.Source}</p>
        ${masul.OriginalMemberRecruitmentID ? `<p><strong>Original Member Recruitment ID:</strong> ${masul.OriginalMemberRecruitmentID}</p>` : ''}
        <p><strong>Promotion History:</strong> ${promotionList}</p>
        <p><em>Generated on: ${new Date().toLocaleString()}</em></p>
    `;
}

// ==================== IMAGE FALLBACK FUNCTION ====================
function tryAlternateImage(img, originalUrl) {
    const match = originalUrl.match(/[-\w]{25,}/);
    if (match) {
        const fileId = match[0];
        img.src = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
        img.onerror = () => {
            img.src = 'logo.png';
            img.onerror = null;
        };
    } else {
        img.src = 'logo.png';
        img.onerror = null;
    }
}

// ==================== REGISTRATION PAGE ====================
function initializeRegistrationPage() {
    if (!currentUser) return;

    if (currentUser.role !== 'Admin') {
        document.querySelector('.role-selector').style.display = 'none';
        document.getElementById('masulFormContainer').style.display = 'none';
        document.getElementById('memberFormContainer').style.display = 'block';
    } else {
        document.getElementById('memberFormContainer').style.display = 'block';
        document.getElementById('masulFormContainer').style.display = 'none';
    }

    loadZonesForDropdowns();
    setDOBLimits();

    const masulGender = document.getElementById('masulGender');
    if (masulGender) {
        masulGender.addEventListener('change', function() {
            const gender = this.value;
            const rankSelect = document.getElementById('masulRank');
            const brotherRanks = ['Musa\'id', 'Areef', 'Muqaddam', 'Ra\'id', 'Raqeeb', 'Mulazim', 'Muhafiz', 'Ameed', 'Aqeeda', 'Qaid'];
            const sisterRanks = ['Musa\'ida', 'Areefa', 'Muqadama', 'Ra\'ida', 'Raqeeba', 'Mulazima', 'Muhafiza', 'Ameeda', 'Aqeeda', 'Qaida'];
            rankSelect.innerHTML = '<option value="">Select Rank</option>';
            if (gender === 'Brother') {
                brotherRanks.forEach(rank => {
                    rankSelect.innerHTML += `<option value="${rank}">${rank}</option>`;
                });
            } else if (gender === 'Sister') {
                sisterRanks.forEach(rank => {
                    rankSelect.innerHTML += `<option value="${rank}">${rank}</option>`;
                });
            }
        });
    }

    if (currentUser.role === 'Branch Mas\'ul') {
        const branchField = document.querySelector('select[name="branch"]');
        const zoneField = document.querySelector('select[name="zone"]');
        if (branchField && zoneField) {
            setTimeout(() => {
                const branchCode = currentUser.branchCode;
                const zoneName = branchZoneMap[branchCode];
                if (zoneName) {
                    for (let opt of zoneField.options) {
                        if (opt.value === zoneName) {
                            opt.selected = true;
                            zoneField.disabled = true;
                            break;
                        }
                    }
                    zoneField.dispatchEvent(new Event('change'));
                    setTimeout(() => {
                        for (let opt of branchField.options) {
                            if (opt.value === branchCode) {
                                opt.selected = true;
                                branchField.disabled = true;
                                break;
                            }
                        }
                    }, 500);
                }
            }, 1000);
        }
    }

    document.getElementById('memberForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        const photoFile = formData.get('photo');
        if (photoFile && photoFile.size > 0) {
            if (photoFile.size > 2 * 1024 * 1024) {
                showMessage('File Too Large', 'File size must be less than 2 MB');
                return;
            }
            data.photoBase64 = await fileToBase64(photoFile);
            data.photoName = photoFile.name;
        }
        try {
            const result = await apiRequest('registerMember', { data }, currentUser);
            showSuccessModal(result.intizarId, result.recruitmentId, data.zone, data.branch);
            e.target.reset();
            if (currentUser.role === 'Branch Mas\'ul') {
                document.querySelector('select[name="branch"]').disabled = false;
                document.querySelector('select[name="zone"]').disabled = false;
            }
        } catch (err) {
            showMessage('Registration Failed', err.message);
        }
    });

    document.getElementById('masulForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        const photoFile = formData.get('photo');
        if (photoFile && photoFile.size > 0) {
            if (photoFile.size > 2 * 1024 * 1024) {
                showMessage('File Too Large', 'File size must be less than 2 MB');
                return;
            }
            data.photoBase64 = await fileToBase64(photoFile);
            data.photoName = photoFile.name;
        }
        try {
            const result = await apiRequest('registerMasul', { data }, currentUser);
            showSuccessModal(result.intizarId, result.masulRecruitmentId, data.zone, data.branch);
            e.target.reset();
        } catch (err) {
            showMessage('Registration Failed', err.message);
        }
    });

    document.getElementById('source').addEventListener('change', function() {
        const field = document.getElementById('intizarIdField');
        field.style.display = this.value === 'xghalibun' ? 'block' : 'none';
        if (this.value !== 'xghalibun') {
            document.querySelector('input[name="intizarId"]').value = '';
        }
    });
}

function toggleRegistrationForm() {
    const role = document.getElementById('roleSelector').value;
    document.getElementById('memberFormContainer').style.display = role === 'member' ? 'block' : 'none';
    document.getElementById('masulFormContainer').style.display = role === 'masul' ? 'block' : 'none';
}

function showSuccessModal(intizarId, recruitmentId, zone, branch) {
    document.getElementById('generatedId').innerText = intizarId;
    document.getElementById('generatedRecruitmentId').innerText = recruitmentId;
    document.getElementById('generatedZone').innerText = zone;
    document.getElementById('generatedBranch').innerText = branch;
    document.getElementById('successModal').style.display = 'block';
}

function closeSuccessModal() {
    document.getElementById('successModal').style.display = 'none';
}

function setDOBLimits() {
    const today = new Date();
    const maxDateMember = new Date(today.getFullYear() - 7, today.getMonth(), today.getDate()).toISOString().split('T')[0];
    const maxDateMasul = new Date(today.getFullYear() - 18, today.getMonth(), today.getDate()).toISOString().split('T')[0];
    const memberDob = document.getElementById('memberDob');
    const masulDob = document.getElementById('masulDob');
    if (memberDob) memberDob.setAttribute('max', maxDateMember);
    if (masulDob) masulDob.setAttribute('max', maxDateMasul);
}

// ==================== ZONE/BRANCH DROPDOWNS ====================
async function loadZonesForDropdowns() {
    try {
        const result = await apiRequest('getZones', {}, currentUser);
        const zones = result.zones.filter(z => z.status === 'Active');
        const zoneSelects = document.querySelectorAll('select[name="zone"], #editBranchZone, #branchModal select[name="zoneName"]');
        zoneSelects.forEach(select => {
            if (!select) return;
            select.innerHTML = '<option value="">Select Zone</option>';
            zones.forEach(zone => {
                select.innerHTML += `<option value="${zone.zoneName}">${zone.zoneName}</option>`;
            });
        });

        branchZoneMap = {};
        for (let zone of zones) {
            const branchRes = await apiRequest('getBranches', { zone: zone.zoneName }, currentUser);
            branchRes.branches.forEach(b => {
                if (b.status === 'Active') {
                    branchZoneMap[b.branchCode] = zone.zoneName;
                }
            });
        }

        document.querySelectorAll('select[name="zone"]').forEach(select => {
            select.addEventListener('change', async function() {
                const zone = this.value;
                const branchSelect = this.closest('fieldset') ?
                    this.closest('fieldset').parentElement.querySelector('select[name="branch"]') :
                    document.querySelector('select[name="branch"]');
                if (!branchSelect) return;
                branchSelect.innerHTML = '<option value="">Select Branch</option>';
                if (!zone) return;
                try {
                    const result = await apiRequest('getBranches', { zone }, currentUser);
                    result.branches.filter(b => b.status === 'Active').forEach(branch => {
                        branchSelect.innerHTML += `<option value="${branch.branchCode}">${branch.branchName}</option>`;
                    });
                } catch (err) {
                    console.error(err);
                }
            });
        });
    } catch (err) {
        console.error(err);
    }
}

// ==================== ZONE/BRANCH ACTIONS ====================
function showAddZoneModal() { showModal('zoneModal'); }
function editZone(zoneId, zoneName) {
    document.getElementById('editZoneId').value = zoneId;
    document.getElementById('editZoneName').value = zoneName;
    showModal('editZoneModal');
}
document.addEventListener('DOMContentLoaded', () => {
    const zoneForm = document.getElementById('zoneForm');
    if (zoneForm) {
        zoneForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const zoneName = zoneForm.zoneName.value;
            try {
                await apiRequest('addZone', { zoneName }, currentUser);
                showMessage('Success', 'Zone added successfully');
                hideModal('zoneModal');
                zoneForm.reset();
                loadZones();
            } catch (err) {
                showMessage('Error', err.message);
            }
        });
    }
    const editZoneForm = document.getElementById('editZoneForm');
    if (editZoneForm) {
        editZoneForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const zoneId = document.getElementById('editZoneId').value;
            const newName = document.getElementById('editZoneName').value;
            try {
                await apiRequest('editZone', { zoneId, newName }, currentUser);
                showMessage('Success', 'Zone updated');
                hideModal('editZoneModal');
                loadZones();
            } catch (err) {
                showMessage('Error', err.message);
            }
        });
    }
    const branchForm = document.getElementById('branchForm');
    if (branchForm) {
        branchForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const branchName = branchForm.branchName.value;
            const zoneName = branchForm.zoneName.value;
            try {
                await apiRequest('addBranch', { branchName, zoneName }, currentUser);
                showMessage('Success', 'Branch added successfully');
                hideModal('branchModal');
                branchForm.reset();
                loadBranches();
            } catch (err) {
                showMessage('Error', err.message);
            }
        });
    }
    const editBranchForm = document.getElementById('editBranchForm');
    if (editBranchForm) {
        editBranchForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const branchCode = document.getElementById('editBranchCode').value;
            const newName = document.getElementById('editBranchName').value;
            const newZone = document.getElementById('editBranchZone').value;
            try {
                await apiRequest('editBranch', { branchCode, newName, newZone }, currentUser);
                showMessage('Success', 'Branch updated');
                hideModal('editBranchModal');
                loadBranches();
            } catch (err) {
                showMessage('Error', err.message);
            }
        });
    }
});
function showAddBranchModal() { showModal('branchModal'); }
function editBranch(branchCode, branchName, zone) {
    document.getElementById('editBranchCode').value = branchCode;
    document.getElementById('editBranchName').value = branchName;
    const zoneSelect = document.getElementById('editBranchZone');
    for (let opt of zoneSelect.options) {
        if (opt.value === zone) opt.selected = true;
    }
    showModal('editBranchModal');
}
async function disableZone(zoneId) {
    if (!(await showConfirm('Confirm', 'Disable this zone?'))) return;
    try {
        await apiRequest('disableZone', { zoneId }, currentUser);
        showMessage('Success', 'Zone disabled');
        loadZones();
    } catch (err) {
        showMessage('Error', err.message);
    }
}
async function enableZone(zoneId) {
    if (!(await showConfirm('Confirm', 'Enable this zone?'))) return;
    try {
        await apiRequest('enableZone', { zoneId }, currentUser);
        showMessage('Success', 'Zone enabled');
        loadZones();
    } catch (err) {
        showMessage('Error', err.message);
    }
}
async function disableBranch(branchCode) {
    if (!(await showConfirm('Confirm', 'Disable this branch?'))) return;
    try {
        await apiRequest('disableBranch', { branchCode }, currentUser);
        showMessage('Success', 'Branch disabled');
        loadBranches();
    } catch (err) {
        showMessage('Error', err.message);
    }
}
async function enableBranch(branchCode) {
    if (!(await showConfirm('Confirm', 'Enable this branch?'))) return;
    try {
        await apiRequest('enableBranch', { branchCode }, currentUser);
        showMessage('Success', 'Branch enabled');
        loadBranches();
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== LOAD ZONES ====================
async function loadZones() {
    try {
        const result = await apiRequest('getZones', {}, currentUser);
        const tbody = document.querySelector('#zonesTable tbody');
        tbody.innerHTML = '';
        result.zones.forEach(zone => {
            const row = tbody.insertRow();
            row.insertCell().innerText = zone.zoneName;
            row.insertCell().innerText = zone.status;
            const actions = row.insertCell();
            actions.innerHTML = `
                <button onclick="editZone('${zone.zoneId}', '${zone.zoneName}')">Edit</button>
                ${zone.status === 'Active'
                    ? `<button onclick="disableZone('${zone.zoneId}')">Disable</button>`
                    : `<button onclick="enableZone('${zone.zoneId}')">Enable</button>`}
            `;
        });
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load zones: ' + err.message);
    }
}

// ==================== LOAD BRANCHES ====================
async function loadBranches() {
    try {
        const result = await apiRequest('getBranches', {}, currentUser);
        const tbody = document.querySelector('#branchesTable tbody');
        tbody.innerHTML = '';
        result.branches.forEach(branch => {
            const row = tbody.insertRow();
            row.insertCell().innerText = branch.branchCode;
            row.insertCell().innerText = branch.branchName;
            row.insertCell().innerText = branch.zone;
            row.insertCell().innerText = branch.status;
            const actions = row.insertCell();
            actions.innerHTML = `
                <button onclick="editBranch('${branch.branchCode}', '${branch.branchName}', '${branch.zone}')">Edit</button>
                ${branch.status === 'Active'
                    ? `<button onclick="disableBranch('${branch.branchCode}')">Disable</button>`
                    : `<button onclick="enableBranch('${branch.branchCode}')">Enable</button>`}
            `;
        });
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load branches: ' + err.message);
    }
}

// ==================== AUDIT LOG ====================
async function loadAuditLog() {
    try {
        const result = await apiRequest('getAuditLog', {}, currentUser);
        const tbody = document.querySelector('#auditTable tbody');
        tbody.innerHTML = '';
        result.logs.forEach(log => {
            const row = tbody.insertRow();
            row.insertCell().innerText = new Date(log.timestamp).toLocaleString();
            row.insertCell().innerText = log.user;
            row.insertCell().innerText = log.action;
            row.insertCell().innerText = log.details;
        });
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load audit log: ' + err.message);
    }
}

// ==================== CONFIG ====================
async function loadConfig() {
    try {
        const adminCode = await apiRequest('getConfig', { key: 'admin_code' }, currentUser);
        document.getElementById('configAdminCode').value = adminCode.value || '';
        const prefix = await apiRequest('getConfig', { key: 'access_prefix' }, currentUser);
        document.getElementById('configPrefix').value = prefix.value || 'Muntazir@';
    } catch (err) {
        console.error(err);
    }
    document.getElementById('configForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const newAdminCode = document.getElementById('configAdminCode').value;
        const newPrefix = document.getElementById('configPrefix').value;
        try {
            if (newAdminCode) await apiRequest('updateConfig', { key: 'admin_code', value: newAdminCode }, currentUser);
            if (newPrefix) await apiRequest('updateConfig', { key: 'access_prefix', value: newPrefix }, currentUser);
            showMessage('Success', 'Configuration updated');
        } catch (err) {
            showMessage('Error', err.message);
        }
    });
}

// ==================== EXPORT ====================
async function exportData(type) {
    try {
        const result = await apiRequest('exportData', { type }, currentUser);
        const blob = new Blob([result.csv], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = result.filename;
        a.click();
        window.URL.revokeObjectURL(url);
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== PROMOTIONS ====================
async function promoteMember(intizarId) {
    if (!(await showConfirm('Confirm', 'Promote this member?'))) return;
    try {
        await apiRequest('promoteMember', { intizarId }, currentUser);
        showMessage('Success', 'Member promoted successfully');
        loadMemberList(currentMemberPage, currentMemberSearch, currentMemberFilters);
    } catch (err) {
        showMessage('Error', err.message);
    }
}
async function promoteMasul(intizarId) {
    if (!(await showConfirm('Confirm', 'Promote this Mas\'ul?'))) return;
    try {
        await apiRequest('promoteMasul', { intizarId }, currentUser);
        showMessage('Success', 'Mas\'ul promoted successfully');
        loadMasuls(currentMasulPage, currentMasulSearch, currentMasulFilters);
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== TRANSFERS ====================
async function transferMember(intizarId) {
    const newBranch = await showPrompt('Transfer Member', 'Enter new Branch Code:');
    if (!newBranch) return;
    try {
        await apiRequest('transferMember', { intizarId, newBranchCode: newBranch }, currentUser);
        showMessage('Success', 'Member transferred');
        loadMemberList(currentMemberPage, currentMemberSearch, currentMemberFilters);
    } catch (err) {
        showMessage('Error', err.message);
    }
}
async function transferMasul(intizarId) {
    const newBranch = await showPrompt('Transfer Mas\'ul', 'Enter new Branch Code:');
    if (!newBranch) return;
    try {
        await apiRequest('transferMasul', { intizarId, newBranchCode: newBranch }, currentUser);
        showMessage('Success', 'Mas\'ul transferred');
        loadMasuls(currentMasulPage, currentMasulSearch, currentMasulFilters);
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== DASHBOARD STATS & CHARTS (with graceful error handling) ====================
async function loadDashboardStats() {
    const statsContainer = document.querySelector('.stats-grid');
    try {
        const result = await apiRequest('getDashboardStats', {}, currentUser);
        const stats = result.stats;
        const setText = (id, value) => {
            const el = document.getElementById(id);
            if (el) el.innerText = value !== undefined && value !== null ? value : '0';
        };
        setText('statTotalCombined', stats.totalCombined);
        setText('statTotalMembers', stats.totalMembers);
        setText('statTotalMasuls', stats.totalMasuls);
        setText('statBrothers', stats.brothers);
        setText('statSisters', stats.sisters);
        setText('statBrothersMembers', stats.brothersMembers);
        setText('statSistersMembers', stats.sistersMembers);
        setText('statBrothersMasuls', stats.brothersMasuls);
        setText('statSistersMasuls', stats.sistersMasuls);
        setText('statBakiyatullah', stats.levelCounts.Bakiyatullah);
        setText('statAnsarullah', stats.levelCounts.Ansarullah);
        setText('statGhalibun', stats.levelCounts.Ghalibun);
        setText('statXGhalibun', stats.levelCounts['X-Ghalibun']);
        updateMembersChart(stats.levelCounts);
        const errorMsg = document.getElementById('statsError');
        if (errorMsg) errorMsg.remove();
    } catch (err) {
        console.error('Failed to load stats', err);
        let errorDiv = document.getElementById('statsError');
        if (!errorDiv) {
            errorDiv = document.createElement('div');
            errorDiv.id = 'statsError';
            errorDiv.style.color = 'red';
            errorDiv.style.textAlign = 'center';
            errorDiv.style.margin = '1rem 0';
            statsContainer.parentNode.insertBefore(errorDiv, statsContainer.nextSibling);
        }
        errorDiv.innerText = 'Failed to load statistics. Please refresh or try again later.';
    }
}

function updateMembersChart(levelCounts) {
    const canvas = document.getElementById('membersChart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (window.membersChart) window.membersChart.destroy();
    // Only create chart if there is data
    const labels = Object.keys(levelCounts);
    const data = Object.values(levelCounts);
    if (labels.length === 0 || data.every(v => v === 0)) {
        console.log('No data for members chart');
        return;
    }
    window.membersChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Number of Members',
                data: data,
                backgroundColor: ['#556B2F', '#C9A87C', '#556B2F', '#000000']
            }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });
}

// ==================== ZONE STATS ====================
async function loadZoneStats() {
    try {
        const result = await apiRequest('getZoneStats', {}, currentUser);
        const stats = result.stats || [];
        const tbody = document.querySelector('#zoneStatsTable tbody');
        if (tbody) {
            tbody.innerHTML = '';
            stats.forEach(zone => {
                const row = tbody.insertRow();
                row.insertCell().innerText = zone.zone;
                row.insertCell().innerText = zone.total;
                row.insertCell().innerText = zone.brothers;
                row.insertCell().innerText = zone.sisters;
            });
        }
        const canvas = document.getElementById('zoneChart');
        if (canvas && stats.length > 0) {
            const ctx = canvas.getContext('2d');
            if (window.zoneChart) window.zoneChart.destroy();
            window.zoneChart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: stats.map(z => z.zone),
                    datasets: [{
                        data: stats.map(z => z.total),
                        backgroundColor: ['#556B2F', '#C9A87C', '#2F4F2F', '#DAA520', '#6B8E23']
                    }]
                }
            });
        } else if (canvas) {
            console.log('No zone data for chart');
        }
    } catch (err) {
        console.error(err);
        const tbody = document.querySelector('#zoneStatsTable tbody');
        if (tbody) tbody.innerHTML = '<tr><td colspan="4" style="color:red;">Failed to load zone stats</td></tr>';
    }
}

// ==================== BRANCH STATS ====================
async function loadBranchStats() {
    try {
        const result = await apiRequest('getBranchStats', {}, currentUser);
        const stats = result.stats || [];
        const tbody = document.querySelector('#branchStatsTable tbody');
        if (tbody) {
            tbody.innerHTML = '';
            stats.forEach(b => {
                const row = tbody.insertRow();
                row.insertCell().innerText = b.branchCode;
                row.insertCell().innerText = b.branchName;
                row.insertCell().innerText = b.zone;
                row.insertCell().innerText = b.total;
                row.insertCell().innerText = b.brothers;
                row.insertCell().innerText = b.sisters;
            });
        }
        const canvas = document.getElementById('branchChart');
        if (canvas && stats.length > 0) {
            const ctx = canvas.getContext('2d');
            if (window.branchChart) window.branchChart.destroy();
            window.branchChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: stats.slice(0, 10).map(b => b.branchCode),
                    datasets: [{
                        label: 'Members per Branch',
                        data: stats.slice(0, 10).map(b => b.total),
                        backgroundColor: '#C9A87C'
                    }]
                }
            });
        } else if (canvas) {
            console.log('No branch data for chart');
        }
    } catch (err) {
        console.error(err);
        const tbody = document.querySelector('#branchStatsTable tbody');
        if (tbody) tbody.innerHTML = '<tr><td colspan="6" style="color:red;">Failed to load branch stats</td></tr>';
    }
}

// ==================== CLOSE MODALS ====================
document.querySelectorAll('.modal .close').forEach(span => {
    span.onclick = function() {
        this.closest('.modal').style.display = 'none';
    };
});
window.onclick = function(event) {
    if (event.target.classList.contains('modal-overlay') || event.target.classList.contains('modal')) {
        event.target.closest('.modal').style.display = 'none';
    }
};
