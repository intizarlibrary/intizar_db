// ==================== CONFIGURATION ====================
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx36cTxRsE8trVdKKXYs7X5GhDew6c94UzAegpUvLtV0FHPMedhjaYNxBjHkLR3AFha/exec'; // REPLACE WITH YOUR DEPLOYED URL
const PAGE_SIZE = 50; // number of rows per page

// ==================== GLOBAL STATE ====================
let currentUser = null;
let currentMemberPage = 1;
let totalMembers = 0;
let currentMasulPage = 1;
let totalMasuls = 0;

// Search state
let currentMemberSearch = '';
let currentMasulSearch = '';

// Branch -> Zone mapping (for Branch Mas'ul auto‑selection)
let branchZoneMap = {};

// ==================== SURAT AL-ASR TYPING ANIMATION ====================
function typeSurahAsr() {
    const surahElement = document.getElementById('surahText');
    if (!surahElement) return;
    const fullText = "وَٱلْعَصْرِ (1) إِنَّ ٱلْإِنسَـٰنَ لَفِى خُسْرٍ (2) إِلَّا ٱلَّذِينَ ءَامَنُوا۟ وَعَمِلُوا۟ ٱلصَّـٰلِحَـٰتِ وَتَوَاصَوْا۟ بِٱلْحَقِّ وَتَوَاصَوْا۟ بِٱلصَّبْرِ (3)";
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

function showSuccessMessage(message) {
    document.getElementById('successMessage').innerText = message;
    document.getElementById('successModal').style.display = 'block';
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
    const closeBtn = document.getElementById('closeSidebar');

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

    if (closeBtn) {
        closeBtn.addEventListener('click', () => {
            sidebar.classList.remove('mobile-open');
            document.body.style.overflow = '';
        });
    }

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
                alert('Login failed: ' + err.message);
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

    // Hide admin sections for non‑Admin users
    if (currentUser.role !== 'Admin') {
        const adminSections = [
            'masulSection', 'zonesSection', 'branchesSection',
            'auditSection', 'configSection', 'exportSection',
            'zoneStatsSection', 'branchStatsSection'
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
    showSection('membersSection');
    await loadDashboardStats();
    await loadMembers(1, ''); // start with empty search
    loadZonesForDropdowns();
    loadChart();

    // Live search for members
    const memberSearch = document.getElementById('memberSearch');
    if (memberSearch) {
        memberSearch.addEventListener('input', debounce(function() {
            loadMembers(1, this.value);
        }, 300));
    }

    // Live search for mas'ulin
    const masulSearch = document.getElementById('masulSearch');
    if (masulSearch) {
        masulSearch.addEventListener('input', debounce(function() {
            loadMasuls(1, this.value);
        }, 300));
    }
}

function setupNavigation() {
    document.getElementById('navMembers').addEventListener('click', (e) => {
        e.preventDefault();
        showSection('membersSection');
        // Reset search when navigating to Members
        document.getElementById('memberSearch').value = '';
        loadMembers(1, '');
        loadDashboardStats();
    });
    const navMasulin = document.getElementById('navMasulin');
    if (navMasulin) {
        navMasulin.addEventListener('click', (e) => {
            e.preventDefault();
            showSection('masulSection');
            // Reset search when navigating to Mas'ulin
            document.getElementById('masulSearch').value = '';
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
        'membersSection', 'masulSection', 'zonesSection',
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

// ==================== PAGINATION CONTROLS ====================
function renderMemberPagination() {
    const totalPages = Math.ceil(totalMembers / PAGE_SIZE);
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
        html += `<button class="page-btn ${i === currentMemberPage ? 'active' : ''}" onclick="loadMembers(${i}, '${currentMemberSearch}')">${i}</button>`;
    }
    html += `<span> Total: ${totalMembers}</span>`;
    document.getElementById('memberPagination').innerHTML = html;
}

function renderMasulPagination() {
    const totalPages = Math.ceil(totalMasuls / PAGE_SIZE);
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
        html += `<button class="page-btn ${i === currentMasulPage ? 'active' : ''}" onclick="loadMasuls(${i}, '${currentMasulSearch}')">${i}</button>`;
    }
    html += `<span> Total: ${totalMasuls}</span>`;
    document.getElementById('masulPagination').innerHTML = html;
}

// ==================== LOAD MEMBERS (with search) ====================
async function loadMembers(page = 1, search = '') {
    try {
        currentMemberPage = page;
        currentMemberSearch = search;
        const result = await apiRequest('getMembers', { page, pageSize: PAGE_SIZE, search }, currentUser);
        totalMembers = result.total;
        renderMemberTable(result.members);
        renderMemberPagination();
    } catch (err) {
        console.error(err);
        alert('Failed to load members: ' + err.message);
    }
}

function renderMemberTable(members) {
    const tbody = document.querySelector('#memberTable tbody');
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
            <button onclick="printMember('${member.IntizarID}')">🖨 Print</button>
            ${currentUser.role === 'Admin' || currentUser.role === 'Zonal Mas\'ul' ? `<button onclick="promoteMember('${member.IntizarID}')">⭐ Promote</button>` : ''}
            ${currentUser.role === 'Admin' ? `<button onclick="transferMember('${member.IntizarID}')">↗ Transfer</button>` : ''}
        `;
    });
}

// ==================== LOAD MASULS (with search) ====================
async function loadMasuls(page = 1, search = '') {
    try {
        currentMasulPage = page;
        currentMasulSearch = search;
        const result = await apiRequest('getMasuls', { page, pageSize: PAGE_SIZE, search }, currentUser);
        totalMasuls = result.total;
        renderMasulTable(result.masuls);
        renderMasulPagination();
    } catch (err) {
        console.error(err);
        alert('Failed to load masuls: ' + err.message);
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
            <button onclick="printMasul('${masul.IntizarID}')">🖨 Print</button>
            <button onclick="promoteMasul('${masul.IntizarID}')">⭐ Promote</button>
            <button onclick="transferMasul('${masul.IntizarID}')">↗ Transfer</button>
        `;
    });
}

// ==================== SEARCH FUNCTIONS ====================
function searchMembers() {
    const searchTerm = document.getElementById('memberSearch').value;
    loadMembers(1, searchTerm);
}

function clearMemberSearch() {
    document.getElementById('memberSearch').value = '';
    loadMembers(1, '');
}

function searchMasuls() {
    const searchTerm = document.getElementById('masulSearch').value;
    loadMasuls(1, searchTerm);
}

function clearMasulSearch() {
    document.getElementById('masulSearch').value = '';
    loadMasuls(1, '');
}

// ==================== VIEW MEMBER ====================
async function viewMember(intizarId) {
    try {
        const result = await apiRequest('getMember', { intizarId }, currentUser);
        const member = result.member;
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
        const content = document.getElementById('viewContent');
        content.innerHTML = `
            <div class="print-area">
                <div class="print-header">
                    <img src="logo.png" alt="Logo" style="height:60px;">
                    <h2>INTIZARUL IMAMUL MUNTAZAR</h2>
                    <p>Member Biodata</p>
                </div>
                ${member.PhotoURL ? `<img src="${member.PhotoURL}" alt="Passport" class="print-photo">` : ''}
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
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        alert(err.message);
    }
}

function printMember(intizarId) {
    viewMember(intizarId);
    setTimeout(() => {
        const printContents = document.getElementById('viewContent').innerHTML;
        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <html>
                <head>
                    <title>Member Biodata</title>
                    <link rel="stylesheet" href="style.css">
                    <style>
                        @media print { body { margin: 1cm; } .print-header { text-align: center; } .print-photo { max-width: 150px; } }
                    </style>
                </head>
                <body>${printContents}</body>
            </html>
        `);
        printWindow.document.close();
        printWindow.onload = () => {
            printWindow.print();
            printWindow.onafterprint = () => printWindow.close();
        };
    }, 100);
}

// ==================== VIEW MASUL ====================
async function viewMasul(intizarId) {
    try {
        const result = await apiRequest('getMasul', { intizarId }, currentUser);
        const masul = result.masul;
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
        const content = document.getElementById('viewContent');
        content.innerHTML = `
            <div class="print-area">
                <div class="print-header">
                    <img src="logo.png" alt="Logo" style="height:60px;">
                    <h2>INTIZARUL IMAMUL MUNTAZAR</h2>
                    <p>Mas'ul Biodata</p>
                </div>
                ${masul.PhotoURL ? `<img src="${masul.PhotoURL}" alt="Passport" class="print-photo">` : ''}
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
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        alert(err.message);
    }
}

function printMasul(intizarId) {
    viewMasul(intizarId);
    setTimeout(() => {
        const printContents = document.getElementById('viewContent').innerHTML;
        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <html>
                <head>
                    <title>Mas'ul Biodata</title>
                    <link rel="stylesheet" href="style.css">
                    <style>
                        @media print { body { margin: 1cm; } .print-header { text-align: center; } .print-photo { max-width: 150px; } }
                    </style>
                </head>
                <body>${printContents}</body>
            </html>
        `);
        printWindow.document.close();
        printWindow.onload = () => {
            printWindow.print();
            printWindow.onafterprint = () => printWindow.close();
        };
    }, 100);
}

// ==================== REGISTRATION PAGE ====================
function initializeRegistrationPage() {
    if (!currentUser) return;

    // Hide role selector and Mas'ul form for non‑Admin
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

    if (currentUser.role === 'Branch Mas\'ul') {
        const branchField = document.querySelector('select[name="branch"]');
        const zoneField = document.querySelector('select[name="zone"]');
        if (branchField && zoneField) {
            setTimeout(() => {
                const branchCode = currentUser.branchCode;
                const zoneName = branchZoneMap[branchCode];
                if (zoneName) {
                    // Select and disable zone
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
                alert('File size must be less than 2 MB');
                return;
            }
            data.photoBase64 = await fileToBase64(photoFile);
            data.photoName = photoFile.name;
        }
        try {
            const result = await apiRequest('registerMember', { data }, currentUser);
            // result contains { intizarId, recruitmentId, zone, branch }
            showSuccessModal(result.intizarId, result.recruitmentId, data.zone, data.branch);
            e.target.reset();
            if (currentUser.role === 'Branch Mas\'ul') {
                document.querySelector('select[name="branch"]').disabled = false;
                document.querySelector('select[name="zone"]').disabled = false;
            }
        } catch (err) {
            alert('Registration failed: ' + err.message);
        }
    });

    document.getElementById('masulForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        const photoFile = formData.get('photo');
        if (photoFile && photoFile.size > 0) {
            if (photoFile.size > 2 * 1024 * 1024) {
                alert('File size must be less than 2 MB');
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
            alert('Registration failed: ' + err.message);
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

// Form switching
function toggleRegistrationForm() {
    const role = document.getElementById('roleSelector').value;
    document.getElementById('memberFormContainer').style.display = role === 'member' ? 'block' : 'none';
    document.getElementById('masulFormContainer').style.display = role === 'masul' ? 'block' : 'none';
}

// Success modal for registration (now shows Recruitment ID)
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

// ==================== DATE OF BIRTH RESTRICTIONS ====================
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

        // Build branchZoneMap for Branch Mas'ul auto‑selection
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
                alert('Zone added successfully');
                hideModal('zoneModal');
                zoneForm.reset();
                loadZones();
            } catch (err) {
                alert(err.message);
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
                alert('Zone updated');
                hideModal('editZoneModal');
                loadZones();
            } catch (err) {
                alert(err.message);
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
                alert('Branch added successfully');
                hideModal('branchModal');
                branchForm.reset();
                loadBranches();
            } catch (err) {
                alert(err.message);
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
                alert('Branch updated');
                hideModal('editBranchModal');
                loadBranches();
            } catch (err) {
                alert(err.message);
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
    if (!confirm('Disable this zone?')) return;
    try {
        await apiRequest('disableZone', { zoneId }, currentUser);
        alert('Zone disabled');
        loadZones();
    } catch (err) {
        alert(err.message);
    }
}
async function enableZone(zoneId) {
    if (!confirm('Enable this zone?')) return;
    try {
        await apiRequest('enableZone', { zoneId }, currentUser);
        alert('Zone enabled');
        loadZones();
    } catch (err) {
        alert(err.message);
    }
}
async function disableBranch(branchCode) {
    if (!confirm('Disable this branch?')) return;
    try {
        await apiRequest('disableBranch', { branchCode }, currentUser);
        alert('Branch disabled');
        loadBranches();
    } catch (err) {
        alert(err.message);
    }
}
async function enableBranch(branchCode) {
    if (!confirm('Enable this branch?')) return;
    try {
        await apiRequest('enableBranch', { branchCode }, currentUser);
        alert('Branch enabled');
        loadBranches();
    } catch (err) {
        alert(err.message);
    }
}

// ==================== LOAD ZONES (for zones table) ====================
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
        alert('Failed to load zones: ' + err.message);
    }
}

// ==================== LOAD BRANCHES (for branches table) ====================
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
        alert('Failed to load branches: ' + err.message);
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
        alert('Failed to load audit log: ' + err.message);
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
            alert('Configuration updated');
        } catch (err) {
            alert(err.message);
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
        alert(err.message);
    }
}

// ==================== PROMOTIONS ====================
async function promoteMember(intizarId) {
    if (!confirm('Promote this member?')) return;
    try {
        await apiRequest('promoteMember', { intizarId }, currentUser);
        alert('Member promoted successfully');
        loadMembers(currentMemberPage, currentMemberSearch);
    } catch (err) {
        alert(err.message);
    }
}
async function promoteMasul(intizarId) {
    if (!confirm('Promote this Mas\'ul?')) return;
    try {
        await apiRequest('promoteMasul', { intizarId }, currentUser);
        alert('Mas\'ul promoted successfully');
        loadMasuls(currentMasulPage, currentMasulSearch);
    } catch (err) {
        alert(err.message);
    }
}

// ==================== TRANSFERS ====================
async function transferMember(intizarId) {
    const newBranch = prompt('Enter new Branch Code:');
    if (!newBranch) return;
    try {
        await apiRequest('transferMember', { intizarId, newBranchCode: newBranch }, currentUser);
        alert('Member transferred');
        loadMembers(currentMemberPage, currentMemberSearch);
    } catch (err) {
        alert(err.message);
    }
}
async function transferMasul(intizarId) {
    const newBranch = prompt('Enter new Branch Code:');
    if (!newBranch) return;
    try {
        await apiRequest('transferMasul', { intizarId, newBranchCode: newBranch }, currentUser);
        alert('Mas\'ul transferred');
        loadMasuls(currentMasulPage, currentMasulSearch);
    } catch (err) {
        alert(err.message);
    }
}

// ==================== DASHBOARD STATS & CHARTS ====================
async function loadDashboardStats() {
    try {
        const result = await apiRequest('getDashboardStats', {}, currentUser);
        const stats = result.stats;
        document.getElementById('statTotalCombined').innerText = stats.totalCombined;
        document.getElementById('statTotalMembers').innerText = stats.totalMembers;
        document.getElementById('statTotalMasuls').innerText = stats.totalMasuls;
        document.getElementById('statBrothers').innerText = stats.brothers;
        document.getElementById('statSisters').innerText = stats.sisters;
        document.getElementById('statBrothersMembers').innerText = stats.brothersMembers;
        document.getElementById('statSistersMembers').innerText = stats.sistersMembers;
        document.getElementById('statBrothersMasuls').innerText = stats.brothersMasuls;
        document.getElementById('statSistersMasuls').innerText = stats.sistersMasuls;
        document.getElementById('statBakiyatullah').innerText = stats.levelCounts.Bakiyatullah || 0;
        document.getElementById('statAnsarullah').innerText = stats.levelCounts.Ansarullah || 0;
        document.getElementById('statGhalibun').innerText = stats.levelCounts.Ghalibun || 0;
        document.getElementById('statXGhalibun').innerText = stats.levelCounts['X-Ghalibun'] || 0;
        updateMembersChart(stats.levelCounts);
    } catch (err) {
        console.error('Failed to load stats', err);
    }
}

function updateMembersChart(levelCounts) {
    const ctx = document.getElementById('membersChart').getContext('2d');
    if (window.membersChart) window.membersChart.destroy();
    window.membersChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(levelCounts),
            datasets: [{
                label: 'Number of Members',
                data: Object.values(levelCounts),
                backgroundColor: ['#556B2F', '#FFD700', '#556B2F', '#000000']
            }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });
}

async function loadZoneStats() {
    try {
        const result = await apiRequest('getZoneStats', {}, currentUser);
        const stats = result.stats;
        const tbody = document.querySelector('#zoneStatsTable tbody');
        tbody.innerHTML = '';
        stats.forEach(zone => {
            const row = tbody.insertRow();
            row.insertCell().innerText = zone.zone;
            row.insertCell().innerText = zone.total;
            row.insertCell().innerText = zone.brothers;
            row.insertCell().innerText = zone.sisters;
        });
        const ctx = document.getElementById('zoneChart').getContext('2d');
        if (window.zoneChart) window.zoneChart.destroy();
        window.zoneChart = new Chart(ctx, {
            type: 'pie',
            data: {
                labels: stats.map(z => z.zone),
                datasets: [{
                    data: stats.map(z => z.total),
                    backgroundColor: ['#556B2F', '#FFD700', '#2F4F2F', '#DAA520', '#6B8E23']
                }]
            }
        });
    } catch (err) {
        console.error(err);
        alert('Failed to load zone stats');
    }
}

async function loadBranchStats() {
    try {
        const result = await apiRequest('getBranchStats', {}, currentUser);
        const stats = result.stats;
        const tbody = document.querySelector('#branchStatsTable tbody');
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
        const ctx = document.getElementById('branchChart').getContext('2d');
        if (window.branchChart) window.branchChart.destroy();
        window.branchChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: stats.slice(0, 10).map(b => b.branchCode),
                datasets: [{
                    label: 'Members per Branch',
                    data: stats.slice(0, 10).map(b => b.total),
                    backgroundColor: '#556B2F'
                }]
            }
        });
    } catch (err) {
        console.error(err);
        alert('Failed to load branch stats');
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
