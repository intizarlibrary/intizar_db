/**
 * INTIZARUL IMAMUL MUNTAZAR – Frontend Logic & Application Engine (script.js)
 * Supports Google Apps Script live backend only.
 * Includes Graduate (Al-Mahdi) management, safe multi-field search,
 * collapsible sidebar navigation, and ID printing.
 */

// ==================== CONFIGURATION ====================
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxZPOn5xCpGlJrGjX92hrbCDeGqk3HqCfVhlTes9IbRJHUgIqBCU3dhsMaYJrWg7wcO4g/exec';
const PAGE_SIZE = 50;

// ==================== GLOBAL STATE ====================
let currentUser = JSON.parse(sessionStorage.getItem('iim_user')) || null;
let currentMembers = [];
let currentMasuls = [];
let currentGraduates = [];
let currentZones = [];
let currentBranches = []; // Cached branches for zone filtering
let memberSearchTerm = '';
let graduateSearchTerm = '';
let masulSearchTerm = '';
let currentMemberPage = 1;
let currentMasulPage = 1;
let currentMemberFilters = {};
let currentMasulFilters = {};

// ==================== LOADER ====================
let pendingRequests = 0;

function showLoader() {
    pendingRequests++;
    const loader = document.getElementById('globalLoader');
    if (loader) loader.style.display = 'flex';
}

function hideLoader() {
    pendingRequests = Math.max(0, pendingRequests - 1);
    if (pendingRequests === 0) {
        const loader = document.getElementById('globalLoader');
        if (loader) loader.style.display = 'none';
    }
}

// ==================== HELPER: Get Thumbnail URL from PhotoURL ====================
function getThumbnailUrl(photoUrl) {
    if (!photoUrl) return 'logo.png';
    const match = photoUrl.match(/[-\w]{25,}/);
    if (match) {
        return `https://drive.google.com/thumbnail?id=${match[0]}&sz=w1000`;
    }
    return photoUrl;
}

// ==================== SAFE NORMALIZATION ====================
function safeNormalize(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim().toLowerCase();
}

// ==================== CUSTOM MODALS ====================
function showMessage(title, text) {
    const modal = document.getElementById('messageModal');
    if (modal) {
        document.getElementById('messageModalTitle').innerText = title;
        document.getElementById('messageModalText').innerText = text;
        modal.style.display = 'block';
    } else {
        alert(`${title}: ${text}`);
    }
}

function closeMessageModal() {
    const modal = document.getElementById('messageModal');
    if (modal) modal.style.display = 'none';
}

function showConfirm(title, text) {
    return new Promise((resolve) => {
        const modal = document.getElementById('confirmModal');
        if (modal) {
            document.getElementById('confirmModalTitle').innerText = title;
            document.getElementById('confirmModalText').innerText = text;
            modal.style.display = 'block';
            document.getElementById('confirmOkBtn').onclick = () => {
                closeConfirmModal();
                resolve(true);
            };
            document.getElementById('confirmCancelBtn').onclick = () => {
                closeConfirmModal();
                resolve(false);
            };
        } else {
            resolve(confirm(`${title}\n\n${text}`));
        }
    });
}

function closeConfirmModal() {
    const modal = document.getElementById('confirmModal');
    if (modal) modal.style.display = 'none';
}

function showPrompt(title, text, defaultValue = '') {
    return new Promise((resolve) => {
        const modal = document.getElementById('promptModal');
        if (modal) {
            document.getElementById('promptModalTitle').innerText = title;
            document.getElementById('promptModalText').innerText = text;
            document.getElementById('promptInput').value = defaultValue;
            modal.style.display = 'block';
            document.getElementById('promptOkBtn').onclick = () => {
                const val = document.getElementById('promptInput').value;
                closePromptModal();
                resolve(val);
            };
            document.getElementById('promptCancelBtn').onclick = () => {
                closePromptModal();
                resolve(null);
            };
        } else {
            const val = prompt(`${title}\n\n${text}`, defaultValue);
            resolve(val);
        }
    });
}

function closePromptModal() {
    const modal = document.getElementById('promptModal');
    if (modal) modal.style.display = 'none';
}

function showModal(modalId) {
    const el = document.getElementById(modalId);
    if (el) el.style.display = 'block';
}

function hideModal(modalId) {
    const el = document.getElementById(modalId);
    if (el) el.style.display = 'none';
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    });
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
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Accept': 'application/json'
            },
            body: formBody.toString()
        });

        if (!response.ok) {
            throw new Error(`HTTP Error: ${response.status} - ${response.statusText}`);
        }

        const result = await response.json();

        if (!result.success) {
            throw new Error(result.error || 'Unknown error occurred');
        }

        return result;
    } catch (err) {
        console.error('API Request failed:', err);
        throw err;
    } finally {
        hideLoader();
    }
}

// ==================== SURAH PRELOADER ====================
function typeSurahAsr() {
    const surahElement = document.getElementById('surahText');
    if (!surahElement) return;
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

// ==================== SIDEBAR TOGGLE (FIXED) ====================
function initSidebar() {
    const sidebar = document.getElementById('sidebar');
    const toggleBtn = document.getElementById('sidebarToggleBtn');
    if (!sidebar || !toggleBtn) return;
    
    const isMobile = () => window.matchMedia('(max-width: 768px)').matches;

    toggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleSidebar();
    });
    
    document.addEventListener('click', (e) => {
        if (isMobile() && sidebar.classList.contains('mobile-open')) {
            if (!sidebar.contains(e.target) && !toggleBtn.contains(e.target)) {
                toggleSidebar(false);
            }
        }
    });
    
    window.addEventListener('resize', () => {
        if (!isMobile()) {
            sidebar.classList.remove('mobile-open');
            document.body.style.overflow = '';
            const overlay = document.getElementById('sidebarOverlay');
            if (overlay) overlay.classList.remove('show');
        } else {
            sidebar.classList.remove('collapsed');
        }
    });
}

// ==================== LOGIN & INIT ====================
document.addEventListener('DOMContentLoaded', () => {
    // Inject dynamic styles
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
        .id-card {
            max-width: 500px;
            margin: auto;
            border: 3px solid #155B2F;
            border-radius: 12px;
            padding: 20px;
            background: white;
        }
        .card-header {
            text-align: center;
            margin-bottom: 15px;
        }
        .card-logo {
            height: 70px;
        }
        .arabic-title {
            font-size: 2rem;
            color: #155B2F;
            margin: 5px 0;
            direction: rtl;
            font-family: 'Amiri', serif;
        }
        .ajami {
            font-size: 1.2rem;
            color: #C9A87C;
            margin-top: -5px;
            margin-bottom: 10px;
            font-family: 'Noto Naskh Arabic', serif;
        }
        .card-body {
            display: flex;
            gap: 20px;
            align-items: center;
            flex-wrap: wrap;
        }
        .card-photo {
            width: 130px;
            height: 150px;
            object-fit: cover;
            border-radius: 8px;
            border: 2px solid #C9A87C;
        }
        .card-details {
            flex: 1;
            min-width: 200px;
        }
        .card-details p {
            margin: 8px 0;
            font-size: 1rem;
        }
        .badge-graduate {
            background-color: #D1FAE5;
            color: #065F46;
            border: 1px solid #A7F3D0;
        }
        .badge-xghalibun {
            background-color: #FEF3C7;
            color: #92400E;
        }
        .badge-active {
            background-color: #E0E7FF;
            color: #3730A3;
        }
        .stat-card.graduate-stat-card {
            background: #1B4D3E;
            border-color: var(--gold);
        }
        .zone-branch-group {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            align-items: center;
        }
        .zone-branch-group select {
            flex: 1;
            min-width: 120px;
        }
    `;
    document.head.appendChild(style);

    // Preloader
    if (window.location.pathname.includes('index.html') || window.location.pathname === '/') {
        typeSurahAsr();
    } else {
        hidePreloader();
    }

    // Check authentication
    if (currentUser) {
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

    // Login link
    const loginLink = document.getElementById('loginLink');
    if (loginLink) {
        loginLink.addEventListener('click', (e) => {
            e.preventDefault();
            showModal('loginModal');
        });
    }

    // Login form
    const loginModal = document.getElementById('loginModal');
    if (loginModal) {
        const span = loginModal.querySelector('.close');
        if (span) span.onclick = () => hideModal('loginModal');
        window.onclick = (event) => {
            if (event.target.classList.contains('modal-overlay') || event.target.classList.contains('modal')) {
                event.target.closest('.modal').style.display = 'none';
            }
        };
        const loginForm = document.getElementById('loginForm');
        if (loginForm) {
            loginForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const role = document.getElementById('role').value;
                const code = document.getElementById('accessCode').value;
                try {
                    const result = await apiRequest('login', { role, code });
                    currentUser = result.user;
                    sessionStorage.setItem('iim_user', JSON.stringify(currentUser));
                    window.location.href = 'dashboard.html';
                } catch (err) {
                    showMessage('Login Failed', err.message);
                }
            });
        }
    }

    // Password toggle (eye icon)
    const togglePassBtn = document.getElementById('togglePasswordBtn');
    if (togglePassBtn) {
        togglePassBtn.addEventListener('click', function() {
            const input = document.getElementById('accessCode');
            if (input) {
                input.type = input.type === 'password' ? 'text' : 'password';
            }
        });
    }

    // Logout
    const logoutLink = document.getElementById('logoutLink');
    if (logoutLink) {
        logoutLink.addEventListener('click', (e) => {
            e.preventDefault();
            sessionStorage.removeItem('iim_user');
            currentUser = null;
            window.location.href = 'index.html';
        });
    }

    // Enter key support for search inputs
    document.querySelectorAll('.search-bar input').forEach(input => {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                const parent = input.closest('.search-bar');
                const btn = parent ? parent.querySelector('button') : null;
                if (btn) btn.click();
            }
        });
    });

    // Handle any inline onclick for logout in registration page
    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', (e) => {
            e.preventDefault();
            logout();
        });
    }
});

// ==================== DASHBOARD INIT ====================
async function initializeDashboard() {
    if (!currentUser) return;

    // Use userRoleBadge as per HTML
    const roleDisplay = document.getElementById('userRoleBadge');
    if (roleDisplay) roleDisplay.innerText = currentUser.role;

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

    // Show default section (overview)
    switchSection('overview');
    await loadDashboardStats();
    await loadMembersList(1, '');
    await loadFilterOptions();
    loadZonesForDropdowns();
}

// ==================== SECTION SWITCHING ====================
function switchSection(sectionId, e) {
    if (e) e.preventDefault();
    const map = {
        overview: 'overviewSection',
        members: 'membersSection',
        graduates: 'graduatesSection',
        masuls: 'masulsSection',
        zones: 'zonesSection',
        branches: 'branchesSection',
        audit: 'auditSection',
        config: 'configSection',
        export: 'exportSection'
    };
    const targetId = map[sectionId];
    if (!targetId) return;
    const target = document.getElementById(targetId);
    if (!target) return;

    // Hide all sections (use .dashboard-section to match HTML)
    const allSections = document.querySelectorAll('.dashboard-section');
    allSections.forEach(sec => sec.style.display = 'none');

    // Show target
    target.style.display = 'block';

    // Update active class on sidebar links
    document.querySelectorAll('.sidebar-menu a').forEach(a => a.classList.remove('active'));
    const link = document.querySelector(`.sidebar-menu a[data-section="${sectionId}"]`);
    if (link) link.classList.add('active');

    // Additional logic for specific sections
    if (sectionId === 'members') {
        loadMembersList(currentMemberPage, document.getElementById('memberSearchInput')?.value || '', currentMemberFilters);
    } else if (sectionId === 'masuls') {
        loadMasuls(currentMasulPage, document.getElementById('masulSearchInput')?.value || '', currentMasulFilters);
    } else if (sectionId === 'graduates') {
        loadGraduatesList();
    } else if (sectionId === 'zones') {
        loadZones();
    } else if (sectionId === 'branches') {
        loadBranches();
    } else if (sectionId === 'audit') {
        loadAuditLog();
    } else if (sectionId === 'config') {
        loadConfig();
    } else if (sectionId === 'overview') {
        loadDashboardStats();
    }
}

// ==================== FILTER OPTIONS ====================
async function loadFilterOptions() {
    try {
        const result = await apiRequest('getFilterOptions', {}, currentUser);
        populateSelect('filterLevel', result.levels, true);
        populateSelect('filterBranch', result.branches, true);
        populateSelect('filterZone', result.zones, true);
        populateSelect('filterMasulRank', result.ranks, true);
        populateSelect('filterMasulBranch', result.branches, true);
        populateSelect('filterMasulZone', result.zones, true);
        populateSelect('filterGender', result.genders || ['Brother', 'Sister'], true);
        populateSelect('filterMasulGender', result.genders || ['Brother', 'Sister'], true);
        populateSelect('gradFilterGender', result.genders || ['Brother', 'Sister'], true);
        populateSelect('gradFilterZone', result.zones, true);
        populateSelect('gradFilterBranch', result.branches, true);
    } catch (err) {
        console.error('Failed to load filter options:', err);
        showMessage('Error', 'Could not load filter options: ' + err.message);
    }
}

function populateSelect(selectId, options, keepAllOption = true) {
    const select = document.getElementById(selectId);
    if (!select) return;
    const currentValue = select.value;
    select.innerHTML = '';
    if (keepAllOption) {
        const allOption = document.createElement('option');
        allOption.value = '';
        allOption.textContent = 'All';
        select.appendChild(allOption);
    }
    if (options && Array.isArray(options)) {
        options.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            select.appendChild(option);
        });
    }
    if (currentValue && options && options.includes(currentValue)) {
        select.value = currentValue;
    }
}

// ==================== ZONE/BRANCH DROPDOWNS ====================
async function loadZonesForDropdowns() {
    try {
        const result = await apiRequest('getZones', {}, currentUser);
        const zones = result.zones.filter(z => z.status === 'Active');
        currentZones = zones;
        populateZoneSelects(zones);
        attachZoneChangeListeners();

        // Also populate branch dropdowns
        const branchResult = await apiRequest('getBranches', {}, currentUser);
        currentBranches = branchResult.branches.filter(b => b.status === 'Active');
        populateBranchSelects(currentBranches);

        // Trigger zone change for any pre-selected zones to populate branches
        document.querySelectorAll('select[name="zone"]').forEach(select => {
            if (select.value) {
                const event = new Event('change');
                select.dispatchEvent(event);
            }
        });
    } catch (err) {
        console.warn('Failed to load zones/branches:', err);
        showMessage('Notice', 'Could not load zones and branches. Please refresh.');
    }
}

function populateZoneSelects(zones) {
    const zoneSelects = document.querySelectorAll(
        'select[name="zone"], #editBranchZone, #branchModal select[name="zoneName"], ' +
        '#editMemberZone, #editMasulZone, #memZone, #masZone, #filterZone, #gradFilterZone'
    );
    zoneSelects.forEach(select => {
        if (!select) return;
        const currentValue = select.value;
        const isFilter = select.id === 'filterZone' || select.id === 'gradFilterZone';
        select.innerHTML = isFilter ? '<option value="">All Zones</option>' : '<option value="">Select Zone</option>';
        zones.forEach(zone => {
            select.innerHTML += `<option value="${zone.zoneName}">${zone.zoneName}</option>`;
        });
        if (currentValue) select.value = currentValue;
    });
}

function populateBranchSelects(branches) {
    const branchSelects = document.querySelectorAll(
        '#filterBranch, #gradFilterBranch, #filterMasulBranch, ' +
        '#editMemberBranch, #editMasulBranch, #memBranch, #masBranch'
    );
    branchSelects.forEach(select => {
        if (!select) return;
        const currentValue = select.value;
        const isFilter = select.id === 'filterBranch' || select.id === 'gradFilterBranch' || select.id === 'filterMasulBranch';
        select.innerHTML = isFilter ? '<option value="">All Branches</option>' : '<option value="">Select Branch</option>';
        branches.forEach(b => {
            select.innerHTML += `<option value="${b.branchCode}">${b.branchName} (${b.branchCode})</option>`;
        });
        if (currentValue) select.value = currentValue;
    });
}

function attachZoneChangeListeners() {
    document.querySelectorAll('select[name="zone"]').forEach(select => {
        select.removeEventListener('change', zoneChangeHandler);
        select.addEventListener('change', zoneChangeHandler);
    });
}

async function zoneChangeHandler(event) {
    const zone = event.target.value;
    // Find the branch select: look for .zone-branch-group or fieldset container
    let container = event.target.closest('.zone-branch-group');
    if (!container) container = event.target.closest('fieldset');
    if (!container) {
        // If no container, try to find a sibling branch select
        const branchSelect = event.target.closest('.zone-branch-group')?.querySelector('select[name="branch"]') ||
                           event.target.closest('fieldset')?.querySelector('select[name="branch"]') ||
                           event.target.parentElement?.querySelector('select[name="branch"]');
        if (branchSelect) {
            branchSelect.innerHTML = '<option value="">Select Branch</option>';
            if (!zone) return;
            const branches = currentBranches.filter(b => b.zone === zone);
            branches.forEach(b => {
                branchSelect.innerHTML += `<option value="${b.branchCode}">${b.branchName} (${b.branchCode})</option>`;
            });
        }
        return;
    }
    const branchSelect = container.querySelector('select[name="branch"]');
    if (!branchSelect) return;
    branchSelect.innerHTML = '<option value="">Select Branch</option>';
    if (!zone) return;

    // Use cached branches
    const branches = currentBranches.filter(b => b.zone === zone);
    branches.forEach(b => {
        branchSelect.innerHTML += `<option value="${b.branchCode}">${b.branchName} (${b.branchCode})</option>`;
    });
}

// ==================== REGISTRATION ZONE CHANGE HANDLER ====================
function handleRegZoneChange(prefix) {
    const zoneSelect = document.getElementById(prefix + 'Zone');
    const branchSelect = document.getElementById(prefix + 'Branch');
    if (!zoneSelect || !branchSelect) return;
    const zone = zoneSelect.value;
    branchSelect.innerHTML = '<option value="">Select Branch</option>';
    if (!zone) return;
    const branches = currentBranches.filter(b => b.zone === zone);
    branches.forEach(b => {
        branchSelect.innerHTML += `<option value="${b.branchCode}">${b.branchName} (${b.branchCode})</option>`;
    });
}

// ==================== MEMBERS LIST (FIXED COLUMNS) ====================
async function loadMembersList(page = 1, search = '', filters = {}) {
    currentMemberPage = page;
    currentMemberSearch = search;
    currentMemberFilters = filters;
    try {
        const result = await apiRequest('getMembers', { page, pageSize: PAGE_SIZE, search, filters }, currentUser);
        renderMemberListTable(result.members);
        renderMemberListPagination(result.total, page);
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load members: ' + err.message);
    }
}

function renderMemberListTable(members) {
    const tbody = document.querySelector('#membersTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (!members || members.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 2rem;">No members found.</td></tr>';
        return;
    }
    members.forEach(member => {
        const row = tbody.insertRow();
        row.insertCell().innerText = member.IntizarID || '';
        row.insertCell().innerText = member.RecruitmentID || '';
        row.insertCell().innerText = member.FullName || '';
        row.insertCell().innerText = member.FatherName || '';    // Fixed: Added missing column
        row.insertCell().innerText = member.Gender || '';        // Fixed: Added missing column
        row.insertCell().innerText = member.Level || '';
        row.insertCell().innerText = member.Zone || '';          // Fixed: Added missing column
        row.insertCell().innerText = member.Branch || '';
        const actions = row.insertCell();
        actions.innerHTML = `
            <button onclick="viewMember('${member.IntizarID}')">👁 View</button>
            ${currentUser && currentUser.role === 'Admin' ? `<button onclick="editMember('${member.IntizarID}')">✏️ Edit</button>` : ''}
            ${currentUser && (currentUser.role === 'Admin' || currentUser.role === 'Zonal Mas\'ul') ? `<button onclick="promoteMember('${member.IntizarID}')">⭐ Promote</button>` : ''}
            ${currentUser && currentUser.role === 'Admin' ? `<button onclick="transferMember('${member.IntizarID}')">↗ Transfer</button>` : ''}
        `;
    });
}

function renderMemberListPagination(total, page) {
    const container = document.getElementById('memberListPagination');
    if (!container) return;
    const totalPages = Math.ceil(total / PAGE_SIZE);
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
        html += `<button class="page-btn ${i === page ? 'active' : ''}" onclick="loadMembersList(${i}, '${currentMemberSearch || ''}', ${JSON.stringify(currentMemberFilters || {}).replace(/"/g, '&quot;')})">${i}</button>`;
    }
    html += `<span> Total: ${total}</span>`;
    container.innerHTML = html;
}

function applyMemberFilters() {
    const filters = {
        level: document.getElementById('filterLevel')?.value || '',
        gender: document.getElementById('filterGender')?.value || '',
        branch: document.getElementById('filterBranch')?.value || '',
        zone: document.getElementById('filterZone')?.value || ''
    };
    const search = document.getElementById('memberSearchInput')?.value || '';
    loadMembersList(1, search, filters);
}

function resetMemberFilters() {
    document.getElementById('filterLevel').value = '';
    document.getElementById('filterGender').value = '';
    document.getElementById('filterBranch').value = '';
    document.getElementById('filterZone').value = '';
    applyMemberFilters();
}

function searchMemberList() {
    const search = document.getElementById('memberSearchInput')?.value || '';
    loadMembersList(1, search, currentMemberFilters || {});
}

function clearMemberListSearch() {
    const el = document.getElementById('memberSearchInput');
    if (el) el.value = '';
    loadMembersList(1, '', {});
}

// ==================== GRADUATES LIST ====================
async function loadGraduatesList() {
    const search = document.getElementById('graduateSearchInput')?.value || '';
    const gender = document.getElementById('gradFilterGender')?.value || '';
    const zone = document.getElementById('gradFilterZone')?.value || '';
    const branch = document.getElementById('gradFilterBranch')?.value || '';

    try {
        const result = await apiRequest('getGraduates', {
            search,
            filters: { gender, zone, branch }
        }, currentUser);

        const tbody = document.getElementById('graduatesTableBody');
        if (!tbody) return;

        if (!result.members || result.members.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 2rem;">No Graduates found.</td></tr>';
            return;
        }

        currentGraduates = result.members;
        tbody.innerHTML = result.members.map(g => `
            <tr>
                <td><strong>${g.IntizarID || ''}</strong></td>
                <td>${g.RecruitmentID || ''}</td>
                <td><strong>${g.FullName || ''}</strong></td>
                <td>${g.FatherName || ''}</td>
                <td>${g.Gender || ''}</td>
                <td>${g.Zone || ''}</td>
                <td>${g.Branch || ''}</td>
                <td><span class="badge badge-graduate">Al-Mahdi Community</span></td>
                <td>
                    <button type="button" class="btn-sm" title="View ID Card" onclick="viewMember('${g.IntizarID}')">👁 View</button>
                    ${currentUser && (currentUser.role === 'Admin' || currentUser.role === 'Zonal Mas\'ul') ? `
                        <button type="button" class="btn-sm" style="background:#D97706;" onclick="proposeGraduateAsMasul('${g.IntizarID}')">🛡 Propose Mas'ul</button>
                    ` : ''}
                    ${currentUser && currentUser.role === 'Admin' ? `
                        <button type="button" class="btn-sm btn-secondary" onclick="editMember('${g.IntizarID}')">✏️ Edit</button>
                    ` : ''}
                </td>
            </tr>
        `).join('');

        const pagination = document.getElementById('graduatesPagination');
        if (pagination) {
            const total = result.total || result.members.length;
            pagination.innerHTML = `<span>Total Graduates: ${total}</span>`;
        }
    } catch (err) {
        console.error('Error loading graduates list:', err);
        showMessage('Error', 'Failed to load graduates: ' + err.message);
    }
}

function handleGraduateSearch() {
    clearTimeout(window.searchGraduateTimeout);
    window.searchGraduateTimeout = setTimeout(loadGraduatesList, 300);
}

function resetGraduateFilters() {
    ['graduateSearchInput', 'gradFilterGender', 'gradFilterZone', 'gradFilterBranch'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    loadGraduatesList();
}

// ==================== PROPOSE GRADUATE AS MAS'UL ====================
async function proposeGraduateAsMasul(intizarId) {
    try {
        const res = await apiRequest('getMember', { intizarId }, currentUser);
        if (!res.member) throw new Error('Member not found');
        const member = res.member;

        if (member.Level !== 'Graduate' && member.Status !== 'Graduate') {
            showMessage('Error', 'This member is not a Graduate. Only Graduates can be proposed as Mas\'ul.');
            return;
        }

        if (member.Status === 'Mas\'ul') {
            showMessage('Error', 'This member is already registered as a Mas\'ul.');
            return;
        }

        const gender = member.Gender;
        const brotherRanks = ['Musa\'id', 'Areef', 'Muqaddam', 'Ra\'id', 'Raqeeb', 'Mulazim', 'Muhafiz', 'Ameed', 'Aqeeda', 'Qaid'];
        const sisterRanks = ['Musa\'ida', 'Areefa', 'Muqadama', 'Ra\'ida', 'Raqeeba', 'Mulazima', 'Muhafiza', 'Ameeda', 'Aqeeda', 'Qaida'];
        const ranks = gender === 'Sister' ? sisterRanks : brotherRanks;

        const rank = await showPrompt('Select Initial Rank', `Enter the initial Mas'ul rank for ${member.FullName} (${gender}):\nValid ranks: ${ranks.join(', ')}`);
        if (!rank) return;

        if (!ranks.includes(rank)) {
            showMessage('Invalid Rank', `"${rank}" is not a valid rank for ${gender}.`);
            return;
        }

        const data = {
            source: 'Graduate',
            intizarId: member.IntizarID,
            fullName: member.FullName,
            fatherName: member.FatherName,
            gender: member.Gender,
            dob: member.DOB,
            placeOfBirth: member.PlaceOfBirth || '',
            phone: member.Phone,
            email: member.Email || '',
            address: member.Address,
            state: member.State,
            lga: member.LGA,
            zone: member.Zone,
            branch: member.Branch,
            year: new Date().getFullYear(),
            currentRank: rank,
            photoURL: member.PhotoURL || ''
        };

        const result = await apiRequest('registerMasul', { data }, currentUser);
        if (result.success) {
            showMessage('Success', `${member.FullName} has been proposed as Mas'ul with rank "${rank}".`);
            loadGraduatesList();
            if (document.getElementById('masulsSection')) loadMasuls(1, '', {});
            loadDashboardStats();
        }
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== MASULS LIST (FIXED COLUMNS) ====================
async function loadMasuls(page = 1, search = '', filters = {}) {
    currentMasulPage = page;
    currentMasulSearch = search;
    currentMasulFilters = filters;
    try {
        const result = await apiRequest('getMasuls', { page, pageSize: PAGE_SIZE, search, filters }, currentUser);
        renderMasulTable(result.masuls);
        renderMasulPagination(result.total, page);
    } catch (err) {
        console.error(err);
        showMessage('Error', 'Failed to load masuls: ' + err.message);
    }
}

function renderMasulTable(masuls) {
    const tbody = document.querySelector('#masulsTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (!masuls || masuls.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 2rem;">No Mas\'ulin found.</td></tr>';
        return;
    }
    masuls.forEach(masul => {
        const row = tbody.insertRow();
        row.insertCell().innerText = masul.IntizarID || '';
        row.insertCell().innerText = masul.MasulRecruitmentID || '';
        row.insertCell().innerText = masul.OriginalMemberRecruitmentID || ''; // Fixed: Added missing column
        row.insertCell().innerText = masul.FullName || '';
        row.insertCell().innerText = masul.CurrentRank || '';
        row.insertCell().innerText = masul.Source || '';                      // Fixed: Added missing column
        row.insertCell().innerText = masul.Zone || '';                        // Fixed: Added missing column
        row.insertCell().innerText = masul.Branch || '';
        const actions = row.insertCell();
        actions.innerHTML = `
            <button onclick="viewMasul('${masul.IntizarID}')">👁 View</button>
            ${currentUser && currentUser.role === 'Admin' ? `<button onclick="editMasul('${masul.IntizarID}')">✏️ Edit</button>` : ''}
            ${currentUser && (currentUser.role === 'Admin' || currentUser.role === 'Zonal Mas\'ul') ? `<button onclick="promoteMasul('${masul.IntizarID}')">⭐ Promote</button>` : ''}
            ${currentUser && currentUser.role === 'Admin' ? `<button onclick="transferMasul('${masul.IntizarID}')">↗ Transfer</button>` : ''}
        `;
    });
}

function renderMasulPagination(total, page) {
    const container = document.getElementById('masulPagination');
    if (!container) return;
    const totalPages = Math.ceil(total / PAGE_SIZE);
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
        html += `<button class="page-btn ${i === page ? 'active' : ''}" onclick="loadMasuls(${i}, '${currentMasulSearch || ''}', ${JSON.stringify(currentMasulFilters || {}).replace(/"/g, '&quot;')})">${i}</button>`;
    }
    html += `<span> Total: ${total}</span>`;
    container.innerHTML = html;
}

function applyMasulFilters() {
    const filters = {
        rank: document.getElementById('filterMasulRank')?.value || '',
        gender: document.getElementById('filterMasulGender')?.value || '',
        branch: document.getElementById('filterMasulBranch')?.value || '',
        zone: document.getElementById('filterMasulZone')?.value || ''
    };
    const search = document.getElementById('masulSearchInput')?.value || '';
    loadMasuls(1, search, filters);
}

function resetMasulFilters() {
    document.getElementById('filterMasulRank').value = '';
    document.getElementById('filterMasulGender').value = '';
    document.getElementById('filterMasulBranch').value = '';
    document.getElementById('filterMasulZone').value = '';
    applyMasulFilters();
}

function searchMasulList() {
    const search = document.getElementById('masulSearchInput')?.value || '';
    loadMasuls(1, search, currentMasulFilters || {});
}

function clearMasulListSearch() {
    const el = document.getElementById('masulSearchInput');
    if (el) el.value = '';
    loadMasuls(1, '', {});
}

// ==================== VIEW MEMBER ====================
async function viewMember(intizarId) {
    try {
        const result = await apiRequest('getMember', { intizarId }, currentUser);
        const member = result.member;
        lastViewedMember = member;

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

        const imgSrc = getThumbnailUrl(member.PhotoURL) || 'logo.png';
        const photoHtml = `<img src="${imgSrc}" alt="Passport" style="max-width:150px; border-radius:8px;" onerror="this.src='logo.png'; this.onerror=null;">`;

        const content = document.getElementById('viewContent');
        if (!content) return;
        content.innerHTML = `
            <div class="print-area">
                <div class="print-header">
                    <img src="logo.png" alt="Logo" style="height:60px;">
                    <h2 class="arabic-title">إنتظار ٱلإمام ٱلمنتظر</h2>
                    <p class="ajami">تربير رحي د غنغر جكى</p>
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
                    <button onclick="printCurrentMember()" class="no-print">🖨 Print ID Card</button>
                    <button onclick="screenshotCurrentMember()" class="no-print">📸 Screenshot</button>
                </div>
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== VIEW MASUL ====================
async function viewMasul(intizarId) {
    try {
        const result = await apiRequest('getMasul', { intizarId }, currentUser);
        const masul = result.masul;
        lastViewedMasul = masul;

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

        const imgSrc = getThumbnailUrl(masul.PhotoURL) || 'logo.png';
        const photoHtml = `<img src="${imgSrc}" alt="Passport" style="max-width:150px; border-radius:8px;" onerror="this.src='logo.png'; this.onerror=null;">`;

        const content = document.getElementById('viewContent');
        if (!content) return;
        content.innerHTML = `
            <div class="print-area">
                <div class="print-header">
                    <img src="logo.png" alt="Logo" style="height:60px;">
                    <h2 class="arabic-title">إنتظار ٱلإمام ٱلمنتظر</h2>
                    <p class="ajami">تربير رحي د غنغر جكى</p>
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
                    <button onclick="printCurrentMasul()" class="no-print">🖨 Print ID Card</button>
                    <button onclick="screenshotCurrentMasul()" class="no-print">📸 Screenshot</button>
                </div>
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== SIMPLIFIED ID CARD BUILDER ====================
function buildSimpleCard(person, type) {
    const logoAbsolute = new URL('logo.png', window.location.href).href;
    const imgSrc = getThumbnailUrl(person.PhotoURL) || logoAbsolute;
    const photoHtml = `<img src="${imgSrc}" alt="Photo" class="card-photo" crossorigin="anonymous" onerror="this.src='${logoAbsolute}'; this.onerror=null;">`;

    const idField = type === 'member' ? person.RecruitmentID : person.MasulRecruitmentID;

    return `
        <div class="id-card">
            <div class="card-header">
                <img src="${logoAbsolute}" alt="Logo" class="card-logo">
                <h2 class="arabic-title">إنتظار ٱلإمام ٱلمنتظر</h2>
                <p class="ajami">تربير رحي د غنغر جكى</p>
            </div>
            <div class="card-body">
                ${photoHtml}
                <div class="card-details">
                    <p><strong>Full Name:</strong> ${person.FullName}</p>
                    <p><strong>Intizar ID:</strong> ${person.IntizarID}</p>
                    <p><strong>Recruitment ID:</strong> ${idField}</p>
                    <p><strong>Zone:</strong> ${person.Zone}</p>
                    <p><strong>Branch:</strong> ${person.Branch}</p>
                    <p><strong>Type:</strong> ${type === 'member' ? 'Member' : 'Mas\'ul'}</p>
                </div>
            </div>
        </div>
    `;
}

// ==================== PRINT FUNCTIONS ====================
let lastViewedMember = null;
let lastViewedMasul = null;

function openPrintWindow(content, title) {
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        showMessage('Popup Blocked', 'Please allow popups to print.');
        return;
    }

    const logoAbsolute = new URL('logo.png', window.location.href).href;

    printWindow.document.write(`
        <html>
        <head>
            <title>${title}</title>
            <style>
                body { font-family: 'Segoe UI', sans-serif; margin: 1cm; }
                .id-card { max-width: 400px; margin: auto; border: 2px solid #155B2F; border-radius: 10px; padding: 20px; background: white; }
                .card-header { text-align: center; margin-bottom: 15px; }
                .card-logo { height: 70px; }
                .arabic-title { font-size: 1.8rem; color: #155B2F; margin: 5px 0; direction: rtl; font-family: 'Amiri', serif; }
                .ajami { font-size: 1.2rem; color: #C9A87C; margin-top: -5px; margin-bottom: 10px; font-family: 'Noto Naskh Arabic', serif; }
                .card-body { display: flex; gap: 20px; align-items: center; }
                .card-photo { width: 120px; height: 140px; object-fit: cover; border-radius: 8px; border: 2px solid #C9A87C; }
                .card-details { flex: 1; }
                .card-details p { margin: 8px 0; }
                @media print { button { display: none; } body { margin: 0.5cm; } }
            </style>
        </head>
        <body>
            <div class="id-card">${content}</div>
            <script>
                document.querySelectorAll('img').forEach(img => {
                    if (!img.src || img.src === '') {
                        img.src = '${logoAbsolute}';
                    }
                    img.onerror = function() {
                        this.src = '${logoAbsolute}';
                        this.onerror = null;
                    };
                });
                const images = document.querySelectorAll('img');
                let loaded = 0;
                images.forEach(img => {
                    if (img.complete) loaded++;
                    else img.onload = () => { loaded++; if (loaded === images.length) window.print(); };
                });
                if (loaded === images.length) window.print();
                window.onafterprint = function() { window.close(); };
            <\/script>
        </body>
        </html>
    `);
    printWindow.document.close();
}

function printCurrentMember() {
    if (!lastViewedMember) {
        showMessage('Error', 'No member data to print.');
        return;
    }
    const content = buildSimpleCard(lastViewedMember, 'member');
    openPrintWindow(content, 'Member ID Card');
}

function printCurrentMasul() {
    if (!lastViewedMasul) {
        showMessage('Error', 'No masul data to print.');
        return;
    }
    const content = buildSimpleCard(lastViewedMasul, 'masul');
    openPrintWindow(content, 'Mas\'ul ID Card');
}

// ==================== SCREENSHOT FUNCTIONS ====================
function captureElement(element, filename) {
    if (typeof html2canvas === 'undefined') {
        showMessage('Error', 'html2canvas library not loaded.');
        return;
    }
    const images = Array.from(element.getElementsByTagName('img'));
    const promises = images.map(img => {
        if (img.complete) return Promise.resolve();
        return new Promise(resolve => {
            img.addEventListener('load', resolve);
            img.addEventListener('error', () => {
                setTimeout(resolve, 50);
            });
        });
    });
    Promise.all(promises).then(() => {
        html2canvas(element, { scale: 2, useCORS: true, allowTaint: false }).then(canvas => {
            const link = document.createElement('a');
            link.download = filename;
            link.href = canvas.toDataURL('image/png');
            link.click();
        }).catch(err => {
            showMessage('Error', 'Screenshot failed: ' + err.message);
        });
    });
}

function screenshotCurrentMember() {
    if (!lastViewedMember) {
        showMessage('Error', 'No member data to capture.');
        return;
    }
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = buildSimpleCard(lastViewedMember, 'member');
    tempDiv.style.position = 'absolute';
    tempDiv.style.left = '-9999px';
    document.body.appendChild(tempDiv);
    const safeId = lastViewedMember.IntizarID.replace(/\//g, '-');
    captureElement(tempDiv, safeId + '.png').finally(() => document.body.removeChild(tempDiv));
}

function screenshotCurrentMasul() {
    if (!lastViewedMasul) {
        showMessage('Error', 'No masul data to capture.');
        return;
    }
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = buildSimpleCard(lastViewedMasul, 'masul');
    tempDiv.style.position = 'absolute';
    tempDiv.style.left = '-9999px';
    document.body.appendChild(tempDiv);
    const safeId = lastViewedMasul.IntizarID.replace(/\//g, '-');
    captureElement(tempDiv, safeId + '.png').finally(() => document.body.removeChild(tempDiv));
}

// ==================== EDIT FUNCTIONS ====================
async function editMember(intizarId) {
    try {
        showLoader();
        const result = await apiRequest('getMember', { intizarId }, currentUser);
        if (!result || !result.member) {
            throw new Error('No member data received from server');
        }
        const member = result.member;

        // Use the IDs from dashboard.html (editMember...)
        document.getElementById('editMemberIntizarId').value = member.IntizarID || '';
        document.getElementById('editMemberFullName').value = member.FullName || '';
        document.getElementById('editMemberFatherName').value = member.FatherName || '';
        document.getElementById('editMemberGender').value = member.Gender || 'Brother';
        document.getElementById('editMemberDob').value = formatDateForInput(member.DOB);
        document.getElementById('editMemberPlaceOfBirth').value = member.PlaceOfBirth || '';
        document.getElementById('editMemberPhone').value = member.Phone || '';
        document.getElementById('editMemberEmail').value = member.Email || '';
        document.getElementById('editMemberAddress').value = member.Address || '';
        document.getElementById('editMemberState').value = member.State || '';
        document.getElementById('editMemberLga').value = member.LGA || '';
        document.getElementById('editMemberYear').value = member.Year || '';
        document.getElementById('editMemberLevel').value = member.Level || 'Bakiyatullah';
        document.getElementById('editMemberGuardianName').value = member.GuardianName || '';
        document.getElementById('editMemberGuardianPhone').value = member.GuardianPhone || '';
        document.getElementById('editMemberGuardianAddress').value = member.GuardianAddress || '';

        await loadZonesForDropdowns();

        const zoneSelect = document.getElementById('editMemberZone');
        const branchSelect = document.getElementById('editMemberBranch');
        if (member.Zone) zoneSelect.value = member.Zone;

        // Trigger branch population for the zone
        zoneSelect.dispatchEvent(new Event('change'));
        setTimeout(() => {
            if (member.Branch) branchSelect.value = member.Branch;
        }, 200);

        showModal('editMemberModal');
    } catch (err) {
        console.error('Edit member error:', err);
        showMessage('Error', 'Failed to load member: ' + err.message);
    } finally {
        hideLoader();
    }
}

function closeEditMemberModal() {
    hideModal('editMemberModal');
}

async function editMasul(intizarId) {
    try {
        showLoader();
        const result = await apiRequest('getMasul', { intizarId }, currentUser);
        if (!result || !result.masul) {
            throw new Error('No masul data received from server');
        }
        const masul = result.masul;

        // Check if edit masul modal exists; if not, fallback to view
        if (!document.getElementById('editMasulModal')) {
            showMessage('Notice', 'Edit Mas\'ul functionality is not available in this interface. Please use the view option.');
            viewMasul(intizarId);
            hideLoader();
            return;
        }

        document.getElementById('editMasulIntizarId').value = masul.IntizarID || '';
        document.getElementById('editMasulFullName').value = masul.FullName || '';
        document.getElementById('editMasulFatherName').value = masul.FatherName || '';
        document.getElementById('editMasulGender').value = masul.Gender || 'Brother';
        document.getElementById('editMasulDob').value = formatDateForInput(masul.DOB);
        document.getElementById('editMasulPlaceOfBirth').value = masul.PlaceOfBirth || '';
        document.getElementById('editMasulPhone').value = masul.Phone || '';
        document.getElementById('editMasulEmail').value = masul.Email || '';
        document.getElementById('editMasulAddress').value = masul.Address || '';
        document.getElementById('editMasulState').value = masul.State || '';
        document.getElementById('editMasulLga').value = masul.LGA || '';
        document.getElementById('editMasulYear').value = masul.Year || '';

        await loadZonesForDropdowns();

        const zoneSelect = document.getElementById('editMasulZone');
        const branchSelect = document.getElementById('editMasulBranch');
        if (masul.Zone) zoneSelect.value = masul.Zone;

        zoneSelect.dispatchEvent(new Event('change'));
        setTimeout(() => {
            if (masul.Branch) branchSelect.value = masul.Branch;
        }, 200);

        updateMasulRankOptions(masul.Gender);
        document.getElementById('editMasulRank').value = masul.CurrentRank || '';

        showModal('editMasulModal');
    } catch (err) {
        console.error('Edit masul error:', err);
        showMessage('Error', 'Failed to load masul: ' + err.message);
    } finally {
        hideLoader();
    }
}

function closeEditMasulModal() {
    hideModal('editMasulModal');
}

function formatDateForInput(dateString) {
    if (!dateString) return '';
    const d = new Date(dateString);
    if (isNaN(d.getTime())) return dateString;
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function updateMasulRankOptions(gender) {
    const rankSelect = document.getElementById('editMasulRank');
    if (!rankSelect) return;
    const brotherRanks = ['Musa\'id','Areef','Muqaddam','Ra\'id','Raqeeb','Mulazim','Muhafiz','Ameed','Aqeeda','Qaid'];
    const sisterRanks = ['Musa\'ida','Areefa','Muqadama','Ra\'ida','Raqeeba','Mulazima','Muhafiza','Ameeda','Aqeeda','Qaida'];
    rankSelect.innerHTML = '<option value="">Select Rank</option>';
    const ranks = gender === 'Brother' ? brotherRanks : sisterRanks;
    ranks.forEach(rank => {
        rankSelect.innerHTML += `<option value="${rank}">${rank}</option>`;
    });
}

// ==================== DASHBOARD STATS & CHARTS ====================
async function loadDashboardStats() {
    const statsError = document.getElementById('statsError');
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
        setText('statBakiyatullah', stats.levelCounts?.Bakiyatullah || 0);
        setText('statAnsarullah', stats.levelCounts?.Ansarullah || 0);
        setText('statGhalibun', stats.levelCounts?.Ghalibun || 0);
        setText('statXGhalibun', stats.levelCounts?.['X-Ghalibun'] || 0);

        const gradEl = document.getElementById('statTotalGraduates');
        if (gradEl) gradEl.innerText = stats.levelCounts?.Graduate || 0;

        // Only update charts if canvas elements exist
        if (document.getElementById('levelChart')) {
            updateLevelChart(stats.levelCounts);
        }
        if (document.getElementById('zoneChart')) {
            updateZoneChart(stats.zoneCounts);
        }
        if (document.getElementById('branchChart')) {
            updateBranchChart(stats.branchCounts);
        }

        if (statsError) statsError.style.display = 'none';
    } catch (err) {
        console.error('Failed to load stats', err);
        if (statsError) {
            statsError.innerText = 'Failed to load statistics. Please refresh or try again later.';
            statsError.style.display = 'block';
        }
    }
}

function updateLevelChart(levelCounts) {
    const canvas = document.getElementById('levelChart');
    if (!canvas || typeof Chart === 'undefined') return;
    const existingChart = Chart.getChart(canvas);
    if (existingChart) existingChart.destroy();

    new Chart(canvas, {
        type: 'bar',
        data: {
            labels: Object.keys(levelCounts || {}),
            datasets: [{
                label: 'Members',
                data: Object.values(levelCounts || {}),
                backgroundColor: '#556B2F'
            }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });
}

function updateZoneChart(zoneCounts) {
    const canvas = document.getElementById('zoneChart');
    if (!canvas || typeof Chart === 'undefined') return;
    const existingChart = Chart.getChart(canvas);
    if (existingChart) existingChart.destroy();

    new Chart(canvas, {
        type: 'pie',
        data: {
            labels: Object.keys(zoneCounts || {}),
            datasets: [{
                data: Object.values(zoneCounts || {}),
                backgroundColor: ['#556B2F', '#C9A87C', '#2F4F2F', '#DAA520', '#6B8E23', '#8B4513', '#5F9EA0']
            }]
        },
        options: { responsive: true }
    });
}

function updateBranchChart(branchCounts) {
    const canvas = document.getElementById('branchChart');
    if (!canvas || typeof Chart === 'undefined') return;
    const existingChart = Chart.getChart(canvas);
    if (existingChart) existingChart.destroy();

    const sorted = Object.entries(branchCounts || {}).sort((a,b) => b[1] - a[1]).slice(0,10);
    new Chart(canvas, {
        type: 'bar',
        data: {
            labels: sorted.map(item => item[0]),
            datasets: [{
                label: 'Members',
                data: sorted.map(item => item[1]),
                backgroundColor: '#C9A87C'
            }]
        },
        options: { responsive: true, indexAxis: 'y', plugins: { legend: { display: false } } }
    });
}

// ==================== ZONE/BRANCH ACTIONS ====================
function showAddZoneModal() { showModal('zoneModal'); }
function showAddBranchModal() { showModal('branchModal'); }

function editZone(zoneId, zoneName) {
    document.getElementById('editZoneId').value = zoneId;
    document.getElementById('editZoneName').value = zoneName;
    showModal('editZoneModal');
}

function editBranch(branchCode, branchName, zone) {
    document.getElementById('editBranchCode').value = branchCode;
    document.getElementById('editBranchName').value = branchName;
    const zoneSelect = document.getElementById('editBranchZone');
    if (zoneSelect) {
        for (let opt of zoneSelect.options) {
            if (opt.value === zone) opt.selected = true;
        }
    }
    showModal('editBranchModal');
}

// Form event listeners
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

    // Edit Member Form - using HTML IDs (editMember...)
    const editMemberForm = document.getElementById('editMemberForm');
    if (editMemberForm) {
        editMemberForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const intizarId = document.getElementById('editMemberIntizarId').value;
            const data = {
                fullName: document.getElementById('editMemberFullName').value,
                fatherName: document.getElementById('editMemberFatherName').value,
                gender: document.getElementById('editMemberGender').value,
                dob: document.getElementById('editMemberDob').value,
                placeOfBirth: document.getElementById('editMemberPlaceOfBirth').value,
                phone: document.getElementById('editMemberPhone').value,
                email: document.getElementById('editMemberEmail').value,
                address: document.getElementById('editMemberAddress').value,
                state: document.getElementById('editMemberState').value,
                lga: document.getElementById('editMemberLga').value,
                zone: document.getElementById('editMemberZone').value,
                branch: document.getElementById('editMemberBranch').value,
                year: document.getElementById('editMemberYear').value,
                level: document.getElementById('editMemberLevel').value,
                guardianName: document.getElementById('editMemberGuardianName').value,
                guardianPhone: document.getElementById('editMemberGuardianPhone').value,
                guardianAddress: document.getElementById('editMemberGuardianAddress').value
            };
            try {
                await apiRequest('updateMember', { intizarId, data }, currentUser);
                showMessage('Success', 'Member updated successfully');
                closeEditMemberModal();
                loadMembersList(currentMemberPage || 1, currentMemberSearch || '', currentMemberFilters || {});
            } catch (err) {
                showMessage('Error', err.message);
            }
        });
    }

    // Edit Masul Form
    const editMasulForm = document.getElementById('editMasulForm');
    if (editMasulForm) {
        editMasulForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const intizarId = document.getElementById('editMasulIntizarId').value;
            const data = {
                fullName: document.getElementById('editMasulFullName').value,
                fatherName: document.getElementById('editMasulFatherName').value,
                gender: document.getElementById('editMasulGender').value,
                dob: document.getElementById('editMasulDob').value,
                placeOfBirth: document.getElementById('editMasulPlaceOfBirth').value,
                phone: document.getElementById('editMasulPhone').value,
                email: document.getElementById('editMasulEmail').value,
                address: document.getElementById('editMasulAddress').value,
                state: document.getElementById('editMasulState').value,
                lga: document.getElementById('editMasulLga').value,
                zone: document.getElementById('editMasulZone').value,
                branch: document.getElementById('editMasulBranch').value,
                year: document.getElementById('editMasulYear').value,
                currentRank: document.getElementById('editMasulRank').value
            };
            try {
                await apiRequest('updateMasul', { intizarId, data }, currentUser);
                showMessage('Success', 'Mas\'ul updated successfully');
                closeEditMasulModal();
                loadMasuls(currentMasulPage || 1, currentMasulSearch || '', currentMasulFilters || {});
            } catch (err) {
                showMessage('Error', err.message);
            }
        });
    }

    const editMasulGender = document.getElementById('editMasulGender');
    if (editMasulGender) {
        editMasulGender.addEventListener('change', function() {
            updateMasulRankOptions(this.value);
        });
    }
});

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

// ==================== LOAD ZONES (FIXED COLUMNS) ====================
async function loadZones() {
    try {
        const result = await apiRequest('getZones', {}, currentUser);
        const tbody = document.querySelector('#zonesTableBody');
        if (!tbody) return;
        tbody.innerHTML = '';
        if (!result.zones || result.zones.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 2rem;">No zones found.</td></tr>';
            return;
        }
        result.zones.forEach(zone => {
            const row = tbody.insertRow();
            row.insertCell().innerText = zone.zoneId || '';        // Fixed: Added missing column
            row.insertCell().innerText = zone.zoneName || '';
            row.insertCell().innerText = zone.status || '';
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
        const tbody = document.querySelector('#branchesTableBody');
        if (!tbody) return;
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
        const tbody = document.querySelector('#auditTableBody');
        if (!tbody) return;
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
}

// Save system config (called from button)
window.saveSystemConfig = async function() {
    const newAdminCode = document.getElementById('configAdminCode').value;
    const newPrefix = document.getElementById('configPrefix').value;
    try {
        if (newAdminCode) await apiRequest('updateConfig', { key: 'admin_code', value: newAdminCode }, currentUser);
        if (newPrefix) await apiRequest('updateConfig', { key: 'access_prefix', value: newPrefix }, currentUser);
        showMessage('Success', 'Configuration updated');
    } catch (err) {
        showMessage('Error', err.message);
    }
};

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
    if (!(await showConfirm('Confirm', 'Promote this member to the next level?'))) return;
    try {
        const result = await apiRequest('promoteMember', { intizarId }, currentUser);
        showMessage('Success', `Member promoted to ${result.newLevel}`);
        loadMembersList(currentMemberPage || 1, currentMemberSearch || '', currentMemberFilters || {});
        loadDashboardStats();
        loadGraduatesList();
    } catch (err) {
        showMessage('Error', err.message);
    }
}

async function promoteMasul(intizarId) {
    if (!(await showConfirm('Confirm', 'Promote this Mas\'ul to the next rank?'))) return;
    try {
        const result = await apiRequest('promoteMasul', { intizarId }, currentUser);
        showMessage('Success', `Mas'ul promoted to ${result.newRank}`);
        loadMasuls(currentMasulPage || 1, currentMasulSearch || '', currentMasulFilters || {});
        loadDashboardStats();
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
        loadMembersList(currentMemberPage || 1, currentMemberSearch || '', currentMemberFilters || {});
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
        loadMasuls(currentMasulPage || 1, currentMasulSearch || '', currentMasulFilters || {});
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== REGISTRATION PAGE (FIXED IDS) ====================
async function initializeRegistrationPage() {
    if (!currentUser) return;

    // Hide Mas'ul registration tab for non-admin users (frontend)
    const masulTab = document.querySelector('.tab-btn[data-tab="masul"]');
    const masulContainer = document.getElementById('masulRegistrationForm');
    if (currentUser.role !== 'Admin') {
        if (masulTab) masulTab.style.display = 'none';
        if (masulContainer) masulContainer.style.display = 'none';
        // Show member form
        const memberContainer = document.getElementById('memberRegistrationForm');
        if (memberContainer) memberContainer.style.display = 'block';
        // Ensure the member tab is active
        const memberTab = document.querySelector('.tab-btn[data-tab="member"]');
        if (memberTab) memberTab.classList.add('active');
    } else {
        // Admin: show both, default to member
        if (masulTab) masulTab.style.display = 'inline-flex'; // or block
        // Ensure default active tab is member
        const activeTab = document.querySelector('.tab-btn.active');
        if (!activeTab || activeTab.dataset.tab !== 'member') {
            const memberTab = document.querySelector('.tab-btn[data-tab="member"]');
            if (memberTab) memberTab.classList.add('active');
            if (masulTab) masulTab.classList.remove('active');
        }
        // Show/hide forms based on active tab
        toggleRegistrationForm();
    }

    await loadZonesForDropdowns();
    setDOBLimits();

    // Masul rank dropdown by gender (FIXED: Corrected 'masGender' and 'masRank' IDs)
    const masulGender = document.getElementById('masGender');
    if (masulGender) {
        masulGender.addEventListener('change', function() {
            const gender = this.value;
            const rankSelect = document.getElementById('masRank');
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

    // Branch Mas'ul lock
    if (currentUser.role === 'Branch Mas\'ul') {
        const branchField = document.getElementById('memBranch');
        const zoneField = document.getElementById('memZone');
        if (branchField && zoneField) {
            const branchCode = currentUser.branchCode;
            try {
                const branches = await apiRequest('getBranches', {}, currentUser);
                const branch = branches.branches.find(b => b.branchCode === branchCode);
                if (branch) {
                    zoneField.value = branch.zone;
                    // Dispatch change to populate branch dropdown
                    zoneField.dispatchEvent(new Event('change'));
                    setTimeout(() => {
                        branchField.value = branchCode;
                        zoneField.disabled = true;
                        branchField.disabled = true;
                    }, 300);
                }
            } catch (err) {
                console.warn('Could not lock branch field:', err);
            }
        }
    }

    // Attach form submission listeners for memberForm and masulForm
    const memberForm = document.getElementById('memberForm');
    if (memberForm && !memberForm.hasAttribute('data-listener')) {
        memberForm.setAttribute('data-listener', 'true');
        memberForm.addEventListener('submit', handleMemberRegistrationSubmit);
    }

    const masulForm = document.getElementById('masulForm');
    if (masulForm && !masulForm.hasAttribute('data-listener')) {
        masulForm.setAttribute('data-listener', 'true');
        masulForm.addEventListener('submit', handleMasulRegistrationSubmit);
    }
}

// Registration handlers
async function handleMemberRegistrationSubmit(e) {
    e.preventDefault();
    const form = document.getElementById('memberForm');
    if (!form) {
        showMessage('Error', 'Member registration form not found.');
        return;
    }
    const formData = new FormData(form);
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
    pendingMemberData = data;
    showRegistrationConfirm(data, 'member');
}

async function handleMasulRegistrationSubmit(e) {
    e.preventDefault();
    const form = document.getElementById('masulForm');
    if (!form) {
        showMessage('Error', 'Masul registration form not found.');
        return;
    }
    const formData = new FormData(form);
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
    pendingMasulData = data;
    showRegistrationConfirm(data, 'masul');
}

let pendingMemberData = null;
let pendingMasulData = null;

function showRegistrationConfirm(data, type) {
    const modal = document.getElementById('registrationConfirmModal');
    const content = document.getElementById('registrationConfirmContent');
    if (!modal || !content) {
        // Fallback: show confirmation via confirm
        const msg = `Confirm ${type} registration:\nName: ${data.fullName}\nFather: ${data.fatherName}\nGender: ${data.gender}\nZone: ${data.zone}\nBranch: ${data.branch}`;
        if (confirm(msg)) {
            submitConfirmedRegistration();
        }
        return;
    }
    let html = `<p><strong>Name:</strong> ${data.fullName}</p>
                <p><strong>Father's Name:</strong> ${data.fatherName}</p>
                <p><strong>Gender:</strong> ${data.gender}</p>
                <p><strong>Date of Birth:</strong> ${data.dob}</p>
                <p><strong>Phone:</strong> ${data.phone}</p>
                <p><strong>Email:</strong> ${data.email || '-'}</p>
                <p><strong>Zone:</strong> ${data.zone}</p>
                <p><strong>Branch:</strong> ${data.branch}</p>`;
    if (type === 'member') {
        html += `<p><strong>Entry Level:</strong> ${data.entryLevel}</p>`;
    } else {
        html += `<p><strong>Current Rank:</strong> ${data.currentRank}</p>
                 <p><strong>Source:</strong> ${data.source}</p>`;
    }
    content.innerHTML = html;
    modal.style.display = 'block';
}

function closeRegistrationConfirmModal() {
    const modal = document.getElementById('registrationConfirmModal');
    if (modal) modal.style.display = 'none';
    pendingMemberData = null;
    pendingMasulData = null;
}

async function submitConfirmedRegistration() {
    if (pendingMemberData) {
        try {
            const result = await apiRequest('registerMember', { data: pendingMemberData }, currentUser);
            showSuccessModal(pendingMemberData.fullName, result.intizarId, result.recruitmentId, pendingMemberData.zone, pendingMemberData.branch);
            const form = document.getElementById('memberForm');
            if (form) form.reset();
            pendingMemberData = null;
            closeRegistrationConfirmModal();
        } catch (err) {
            showMessage('Registration Failed', err.message);
        }
    } else if (pendingMasulData) {
        try {
            const result = await apiRequest('registerMasul', { data: pendingMasulData }, currentUser);
            showSuccessModal(pendingMasulData.fullName, result.intizarId, result.masulRecruitmentId, pendingMasulData.zone, pendingMasulData.branch);
            const form = document.getElementById('masulForm');
            if (form) form.reset();
            pendingMasulData = null;
            closeRegistrationConfirmModal();
        } catch (err) {
            showMessage('Registration Failed', err.message);
        }
    }
}

function toggleRegistrationForm() {
    const activeTab = document.querySelector('.tab-btn.active');
    const tab = activeTab ? activeTab.dataset.tab : 'member';
    const memberContainer = document.getElementById('memberRegistrationForm');
    const masulContainer = document.getElementById('masulRegistrationForm');
    if (tab === 'member') {
        if (memberContainer) memberContainer.style.display = 'block';
        if (masulContainer) masulContainer.style.display = 'none';
    } else {
        if (memberContainer) memberContainer.style.display = 'none';
        if (masulContainer) masulContainer.style.display = 'block';
    }
}

function switchRegistrationTab(tab) {
    // Update active tab
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.tab === tab) btn.classList.add('active');
    });
    toggleRegistrationForm();
}

// ==================== FIXED: SUCCESS MODAL ====================
function showSuccessModal(name, intizarId, recruitmentId, zone, branch) {
    document.getElementById('sucName').innerText = name;
    document.getElementById('sucIntizarId').innerText = intizarId;
    document.getElementById('sucRecruitmentId').innerText = recruitmentId;
    document.getElementById('sucZone').innerText = zone;
    document.getElementById('sucBranch').innerText = branch;
    document.getElementById('successModal').style.display = 'block';
}

function closeSuccessModal() {
    document.getElementById('successModal').style.display = 'none';
}

function setDOBLimits() {
    const today = new Date();
    const maxDateMember = new Date(today.getFullYear() - 7, today.getMonth(), today.getDate()).toISOString().split('T')[0];
    const maxDateMasul = new Date(today.getFullYear() - 18, today.getMonth(), today.getDate()).toISOString().split('T')[0];
    const memberDob = document.getElementById('memDOB');
    const masulDob = document.getElementById('masDOB');
    if (memberDob) memberDob.setAttribute('max', maxDateMember);
    if (masulDob) masulDob.setAttribute('max', maxDateMasul);
}

// ==================== SIDEBAR TOGGLE GLOBAL (FIXED) ====================
function toggleSidebar(forceState) {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    const overlay = document.getElementById('sidebarOverlay');
    const isMobile = window.matchMedia('(max-width: 768px)').matches; // Fixed: Using matchMedia

    if (isMobile) {
        if (typeof forceState === 'boolean') {
            if (forceState) {
                sidebar.classList.add('mobile-open');
                if (overlay) overlay.classList.add('show');
            } else {
                sidebar.classList.remove('mobile-open');
                if (overlay) overlay.classList.remove('show');
            }
        } else {
            sidebar.classList.toggle('mobile-open');
            if (overlay) overlay.classList.toggle('show');
        }
        document.body.style.overflow = sidebar.classList.contains('mobile-open') ? 'hidden' : '';
    } else {
        if (typeof forceState === 'boolean') {
            if (forceState) sidebar.classList.remove('collapsed');
            else sidebar.classList.add('collapsed');
        } else {
            sidebar.classList.toggle('collapsed');
        }
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

// ==================== EXPOSE GLOBAL FUNCTIONS ====================
window.viewMember = viewMember;
window.viewMasul = viewMasul;
window.editMember = editMember;
window.editMasul = editMasul;
window.promoteMember = promoteMember;
window.promoteMasul = promoteMasul;
window.transferMember = transferMember;
window.transferMasul = transferMasul;
window.printCurrentMember = printCurrentMember;
window.printCurrentMasul = printCurrentMasul;
window.screenshotCurrentMember = screenshotCurrentMember;
window.screenshotCurrentMasul = screenshotCurrentMasul;
window.closeEditMemberModal = closeEditMemberModal;
window.closeEditMasulModal = closeEditMasulModal;
window.closeIdCardModal = function() { hideModal('idCardModal'); };
window.openSpreadsheet = async function() {
    try {
        const result = await apiRequest('getSpreadsheetUrl', {}, currentUser);
        window.open(result.url, '_blank');
    } catch (err) {
        showMessage('Error', err.message);
    }
};
window.exportData = exportData;
window.loadMembersList = loadMembersList;
window.loadMasuls = loadMasuls;
window.loadGraduatesList = loadGraduatesList;
window.loadZones = loadZones;
window.loadBranches = loadBranches;
window.loadAuditLog = loadAuditLog;
window.showAddZoneModal = showAddZoneModal;
window.showAddBranchModal = showAddBranchModal;
window.editZone = editZone;
window.editBranch = editBranch;
window.disableZone = disableZone;
window.enableZone = enableZone;
window.disableBranch = disableBranch;
window.enableBranch = enableBranch;
window.applyMemberFilters = applyMemberFilters;
window.resetMemberFilters = resetMemberFilters;
window.searchMemberList = searchMemberList;
window.clearMemberListSearch = clearMemberListSearch;
window.applyMasulFilters = applyMasulFilters;
window.resetMasulFilters = resetMasulFilters;
window.searchMasulList = searchMasulList;
window.clearMasulListSearch = clearMasulListSearch;
window.handleGraduateSearch = handleGraduateSearch;
window.resetGraduateFilters = resetGraduateFilters;
window.proposeGraduateAsMasul = proposeGraduateAsMasul;
window.submitConfirmedRegistration = submitConfirmedRegistration;
window.closeRegistrationConfirmModal = closeRegistrationConfirmModal;
window.closeSuccessModal = closeSuccessModal;
window.toggleRegistrationForm = toggleRegistrationForm;
window.switchRegistrationTab = switchRegistrationTab;
window.showMessage = showMessage;
window.closeMessageModal = closeMessageModal;
window.showConfirm = showConfirm;
window.closeConfirmModal = closeConfirmModal;
window.showPrompt = showPrompt;
window.closePromptModal = closePromptModal;
window.switchSection = switchSection;
window.toggleSidebar = toggleSidebar;
window.handleRegZoneChange = handleRegZoneChange;
window.saveSystemConfig = saveSystemConfig;
window.logout = function() {
    sessionStorage.removeItem('iim_user');
    currentUser = null;
    window.location.href = 'index.html';
};
// Alias for openLiveGoogleSheet
window.openLiveGoogleSheet = window.openSpreadsheet;
// Alias for printIdCardDirectly and screenshotIdCard
window.printIdCardDirectly = function() {
    if (lastViewedMember) printCurrentMember();
    else if (lastViewedMasul) printCurrentMasul();
    else showMessage('No card', 'Please view a member or masul first.');
};
window.screenshotIdCard = function() {
    if (lastViewedMember) screenshotCurrentMember();
    else if (lastViewedMasul) screenshotCurrentMasul();
    else showMessage('No card', 'Please view a member or masul first.');
};
