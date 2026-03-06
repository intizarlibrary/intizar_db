// ==================== CONFIGURATION ====================
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxhvZBmOP4GosnhdihTfvqoJRz5SnOTXx6C4IFcKnQr5l2OuOtjBVKbrPD1acgq9yg/exec';
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
let branchNameToCode = {};
let lastViewedMember = null;
let lastViewedMasul = null;

// ==================== CACHE CONSTANTS (30 minutes) ====================
const CACHE_FILTER_OPTIONS = 'filterOptions';
const CACHE_ZONES = 'zones';
const CACHE_BRANCH_MAP = 'branchMap';
const CACHE_BRANCH_NAME_TO_CODE = 'branchNameToCode';
const CACHE_TIMESTAMP = 'cacheTimestamp';
const CACHE_DURATION = 30 * 60 * 1000; // 30 minutes

// Hardcoded fallback data (adjust to your actual zones/branches)
const HARDCODED_ZONES = ['SOKOTO ZONE', 'KADUNA ZONE', 'ABUJA ZONE', 'ZARIA ZONE', 'KANO ZONE', 'BAUCHI ZONE', 'MALUMFASHI ZONE', 'NIGER ZONE', 'QUM ZONE'];
const HARDCODED_BRANCH_MAP = {
    'SOKOTO ZONE': ['Sokoto', 'Mafara', 'Yaure', 'Ilela', 'Zuru', 'Yabo'],
    'KADUNA ZONE': ['Kaduna', 'Jaji', 'Mjos'],
    'ABUJA ZONE': ['Maraba', 'Lafia', 'Keffi/Doma', 'Minna', 'Suleja'],
    'ZARIA ZONE': ['Zaria', 'Danja', 'D/Wai', 'Kudan', 'Soba'],
    'KANO ZONE': ['Kano', 'Kazaure', 'Potiskum', 'Gashuwa'],
    'BAUCHI ZONE': ['Bauchi', 'Gombe', 'Azare', 'Jos'],
    'MALUMFASHI ZONE': ['Malumfashi', 'Bakori', 'Katsina'],
    'NIGER ZONE': ['Niyame', 'Maradi'],
    'QUM ZONE': ['Qum']
};

// ==================== LOADER WITH REQUEST COUNTER ====================
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
            sessionStorage.clear(); // Clear all cached data
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
    showSection('membersSection');
    await loadDashboardStats();
    await loadMemberList(1, '');
    await loadFilterOptions();
    loadZonesForDropdowns();

    // Remove input listeners – search on button only
    // Search buttons should call searchMemberList and searchMasulList
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
}

function showSection(sectionId) {
    const sections = [
        'membersSection', 'memberListSection', 'masulSection', 'zonesSection',
        'branchesSection', 'auditSection', 'configSection', 'exportSection'
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
        exportSection: 'navExport'
    };
    const navId = navMap[sectionId];
    if (navId) {
        const navLink = document.getElementById(navId);
        if (navLink) navLink.classList.add('active');
    }
}

// ==================== FILTER OPTIONS (cached) ====================
async function loadFilterOptions(forceRefresh = false) {
    if (!forceRefresh) {
        const cached = sessionStorage.getItem(CACHE_FILTER_OPTIONS);
        const timestamp = sessionStorage.getItem(CACHE_TIMESTAMP);
        if (cached && timestamp && (Date.now() - parseInt(timestamp)) < CACHE_DURATION) {
            const result = JSON.parse(cached);
            populateSelect('filterMemberLevel', result.levels, true);
            populateSelect('filterMemberBranch', result.branches, true);
            populateSelect('filterMemberZone', result.zones, true);
            populateSelect('filterMasulRank', result.ranks.Brother, true); // will be updated by gender change
            populateSelect('filterMasulBranch', result.branches, true);
            populateSelect('filterMasulZone', result.zones, true);
            return;
        }
    }

    try {
        const result = await apiRequest('getFilterOptions', {}, currentUser);
        sessionStorage.setItem(CACHE_FILTER_OPTIONS, JSON.stringify(result));
        sessionStorage.setItem(CACHE_TIMESTAMP, Date.now().toString());

        populateSelect('filterMemberLevel', result.levels, true);
        populateSelect('filterMemberBranch', result.branches, true);
        populateSelect('filterMemberZone', result.zones, true);
        populateSelect('filterMasulRank', result.ranks.Brother, true);
        populateSelect('filterMasulBranch', result.branches, true);
        populateSelect('filterMasulZone', result.zones, true);
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
    options.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        select.appendChild(option);
    });
    if (currentValue && options.includes(currentValue)) {
        select.value = currentValue;
    }
}

// ==================== ZONE/BRANCH DROPDOWNS (cached) ====================
async function loadZonesForDropdowns(forceRefresh = false) {
    let zones, branches;

    if (!forceRefresh) {
        const cachedZones = sessionStorage.getItem(CACHE_ZONES);
        const cachedBranchMap = sessionStorage.getItem(CACHE_BRANCH_MAP);
        const cachedNameToCode = sessionStorage.getItem(CACHE_BRANCH_NAME_TO_CODE);
        const timestamp = sessionStorage.getItem(CACHE_TIMESTAMP);
        if (cachedZones && cachedBranchMap && cachedNameToCode && timestamp && 
            (Date.now() - parseInt(timestamp)) < CACHE_DURATION) {
            zones = JSON.parse(cachedZones);
            branchZoneMap = JSON.parse(cachedBranchMap);
            branchNameToCode = JSON.parse(cachedNameToCode);
            populateZoneSelects(zones);
            attachZoneChangeListeners();
            return;
        }
    }

    try {
        const result = await apiRequest('getZones', {}, currentUser);
        zones = result.zones.filter(z => z.status === 'Active');
        sessionStorage.setItem(CACHE_ZONES, JSON.stringify(zones));
        
        branchZoneMap = {};
        branchNameToCode = {};
        const branchMap = {};
        const nameToCode = {};
        
        for (let zone of zones) {
            const branchRes = await apiRequest('getBranches', { zone: zone.zoneName }, currentUser);
            branchRes.branches.forEach(b => {
                if (b.status === 'Active') {
                    branchMap[b.branchCode] = zone.zoneName;
                    nameToCode[b.branchName] = b.branchCode;
                }
            });
        }
        
        branchZoneMap = branchMap;
        branchNameToCode = nameToCode;
        sessionStorage.setItem(CACHE_BRANCH_MAP, JSON.stringify(branchMap));
        sessionStorage.setItem(CACHE_BRANCH_NAME_TO_CODE, JSON.stringify(nameToCode));
        sessionStorage.setItem(CACHE_TIMESTAMP, Date.now().toString());

        populateZoneSelects(zones);
        attachZoneChangeListeners();

    } catch (err) {
        console.warn('Using hardcoded zones/branches due to API error:', err);
        // Use hardcoded fallback
        zones = HARDCODED_ZONES.map(z => ({ zoneName: z }));
        branchZoneMap = {};
        branchNameToCode = {};
        for (let zone in HARDCODED_BRANCH_MAP) {
            HARDCODED_BRANCH_MAP[zone].forEach(branchName => {
                // Create a simple code: first 2 letters + number? For simplicity use branchName as code
                const code = branchName.replace(/\s+/g, '').substring(0,3).toUpperCase();
                branchZoneMap[code] = zone;
                branchNameToCode[branchName] = code;
            });
        }
        populateZoneSelects(zones);
        attachZoneChangeListeners();
        showMessage('Notice', 'Using default zone/branch lists.');
    }
}

function populateZoneSelects(zones) {
    const zoneSelects = document.querySelectorAll('select[name="zone"], #editBranchZone, #branchModal select[name="zoneName"], #editMemberZone, #editMasulZone');
    zoneSelects.forEach(select => {
        if (!select) return;
        const currentValue = select.value;
        select.innerHTML = '<option value="">Select Zone</option>';
        zones.forEach(zone => {
            select.innerHTML += `<option value="${zone.zoneName}">${zone.zoneName}</option>`;
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
    const branchSelect = event.target.closest('fieldset') ?
        event.target.closest('fieldset').parentElement.querySelector('select[name="branch"]') :
        document.querySelector('select[name="branch"]');
    if (!branchSelect) return;
    branchSelect.innerHTML = '<option value="">Select Branch</option>';
    if (!zone) return;

    // Use cached branchMap to get branches for zone
    const branchesForZone = Object.entries(branchZoneMap)
        .filter(([code, z]) => z === zone)
        .map(([code]) => {
            const name = Object.keys(branchNameToCode).find(key => branchNameToCode[key] === code);
            return { code, name };
        });

    branchesForZone.forEach(b => {
        branchSelect.innerHTML += `<option value="${b.code}">${b.name}</option>`;
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
            ${currentUser.role === 'Admin' ? `<button onclick="editMember('${member.IntizarID}')">✏️ Edit</button>` : ''}
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
    const branchDropdown = document.getElementById('filterMemberBranch').value;
    const branchManual = document.getElementById('filterMemberBranchManual')?.value || '';
    const zoneDropdown = document.getElementById('filterMemberZone').value;
    const zoneManual = document.getElementById('filterMemberZoneManual')?.value || '';

    const branchCode = branchManual || branchDropdown; // manual overrides dropdown
    const zone = zoneManual || zoneDropdown;

    const filters = {
        level: document.getElementById('filterMemberLevel').value,
        gender: document.getElementById('filterMemberGender').value,
        branch: branchCode,
        zone: zone
    };
    const search = document.getElementById('memberListSearch').value;
    loadMemberList(1, search, filters);
}

function resetMemberFilters() {
    document.getElementById('filterMemberLevel').value = '';
    document.getElementById('filterMemberGender').value = '';
    document.getElementById('filterMemberBranch').value = '';
    document.getElementById('filterMemberZone').value = '';
    document.getElementById('filterMemberBranchManual').value = '';
    document.getElementById('filterMemberZoneManual').value = '';
    applyMemberFilters();
}

// Search on button click
function searchMemberList() {
    currentMemberSearch = document.getElementById('memberListSearch').value;
    currentMemberPage = 1;
    loadMemberList(currentMemberPage, currentMemberSearch, currentMemberFilters);
}

function clearMemberListSearch() {
    document.getElementById('memberListSearch').value = '';
    currentMemberSearch = '';
    currentMemberPage = 1;
    loadMemberList(currentMemberPage, currentMemberSearch, currentMemberFilters);
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
            ${currentUser.role === 'Admin' ? `<button onclick="editMasul('${masul.IntizarID}')">✏️ Edit</button>` : ''}
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
    const branchDropdown = document.getElementById('filterMasulBranch').value;
    const branchManual = document.getElementById('filterMasulBranchManual')?.value || '';
    const zoneDropdown = document.getElementById('filterMasulZone').value;
    const zoneManual = document.getElementById('filterMasulZoneManual')?.value || '';

    const branchCode = branchManual || branchDropdown;
    const zone = zoneManual || zoneDropdown;

    const filters = {
        rank: document.getElementById('filterMasulRank').value,
        gender: document.getElementById('filterMasulGender').value,
        branch: branchCode,
        zone: zone
    };
    const search = document.getElementById('masulSearch').value;
    loadMasuls(1, search, filters);
}

function resetMasulFilters() {
    document.getElementById('filterMasulRank').value = '';
    document.getElementById('filterMasulGender').value = '';
    document.getElementById('filterMasulBranch').value = '';
    document.getElementById('filterMasulZone').value = '';
    document.getElementById('filterMasulBranchManual').value = '';
    document.getElementById('filterMasulZoneManual').value = '';
    applyMasulFilters();
}

function searchMasulList() {
    currentMasulSearch = document.getElementById('masulSearch').value;
    currentMasulPage = 1;
    loadMasuls(currentMasulPage, currentMasulSearch, currentMasulFilters);
}

function clearMasulListSearch() {
    document.getElementById('masulSearch').value = '';
    currentMasulSearch = '';
    currentMasulPage = 1;
    loadMasuls(currentMasulPage, currentMasulSearch, currentMasulFilters);
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
                    <button onclick="printCurrentMember()" class="no-print">Print ID Card</button>
                    <button onclick="screenshotCurrentMember()" class="no-print">Screenshot</button>
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
                    <button onclick="printCurrentMasul()" class="no-print">Print ID Card</button>
                    <button onclick="screenshotCurrentMasul()" class="no-print">Screenshot</button>
                </div>
            </div>
        `;
        showModal('viewModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

// ==================== PRINT FUNCTIONS ====================
function openPrintWindow(content, title) {
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        showMessage('Popup Blocked', 'Please allow popups to print.');
        return;
    }

    const logoUrl = new URL('logo.png', window.location.href).href;

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
                    @media print {
                        button { display: none; }
                        body { margin: 0.5cm; }
                    }
                </style>
                <script>
                    function whenAllImagesLoaded(doc) {
                        const images = Array.from(doc.images);
                        if (images.length === 0) return Promise.resolve();
                        return Promise.all(images.map(img => {
                            if (img.complete) return Promise.resolve();
                            return new Promise(resolve => {
                                img.addEventListener('load', resolve);
                                img.addEventListener('error', resolve);
                            });
                        }));
                    }
                    window.onload = function() {
                        whenAllImagesLoaded(document).then(() => {
                            setTimeout(() => {
                                window.print();
                                window.onafterprint = function() { window.close(); };
                            }, 100);
                        });
                    };
                <\/script>
            </head>
            <body>
                <div class="id-card">${content}</div>
                <script>
                    document.querySelectorAll('.card-logo').forEach(img => img.src = '${logoUrl}');
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
function captureElement(element) {
    const images = Array.from(element.getElementsByTagName('img'));
    const promises = images.map(img => {
        if (img.complete) return Promise.resolve();
        return new Promise(resolve => {
            img.addEventListener('load', resolve);
            img.addEventListener('error', () => {
                // Allow fallback to load
                setTimeout(resolve, 50);
            });
        });
    });
    Promise.all(promises).then(() => {
        html2canvas(element, { scale: 2, useCORS: true, allowTaint: false }).then(canvas => {
            const link = document.createElement('a');
            link.download = 'member-card.png';
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
    captureElement(tempDiv).finally(() => document.body.removeChild(tempDiv));
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
    captureElement(tempDiv).finally(() => document.body.removeChild(tempDiv));
}

// ==================== EDIT FUNCTIONS (ADMIN ONLY) ====================
async function editMember(intizarId) {
    try {
        const result = await apiRequest('getMember', { intizarId }, currentUser);
        const member = result.member;

        document.getElementById('editMemberIntizarId').value = member.IntizarID;
        document.getElementById('editMemberFullName').value = member.FullName;
        document.getElementById('editMemberFatherName').value = member.FatherName;
        document.getElementById('editMemberGender').value = member.Gender;
        document.getElementById('editMemberDob').value = member.DOB;
        document.getElementById('editMemberPlaceOfBirth').value = member.PlaceOfBirth || '';
        document.getElementById('editMemberPhone').value = member.Phone;
        document.getElementById('editMemberEmail').value = member.Email || '';
        document.getElementById('editMemberAddress').value = member.Address;
        document.getElementById('editMemberState').value = member.State;
        document.getElementById('editMemberLga').value = member.LGA;
        document.getElementById('editMemberYear').value = member.Year;
        document.getElementById('editMemberLevel').value = member.Level;
        document.getElementById('editMemberGuardianName').value = member.GuardianName;
        document.getElementById('editMemberGuardianPhone').value = member.GuardianPhone;
        document.getElementById('editMemberGuardianAddress').value = member.GuardianAddress;

        await loadZonesForDropdowns(false);
        const zoneSelect = document.getElementById('editMemberZone');
        const branchSelect = document.getElementById('editMemberBranch');
        zoneSelect.value = member.Zone;
        zoneSelect.dispatchEvent(new Event('change', { bubbles: true }));
        // Wait for branch options to populate
        setTimeout(() => {
            branchSelect.value = member.Branch;
        }, 1000);

        showModal('editMemberModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

function closeEditMemberModal() {
    document.getElementById('editMemberModal').style.display = 'none';
}

async function editMasul(intizarId) {
    try {
        const result = await apiRequest('getMasul', { intizarId }, currentUser);
        const masul = result.masul;

        document.getElementById('editMasulIntizarId').value = masul.IntizarID;
        document.getElementById('editMasulFullName').value = masul.FullName;
        document.getElementById('editMasulFatherName').value = masul.FatherName;
        document.getElementById('editMasulGender').value = masul.Gender;
        document.getElementById('editMasulDob').value = masul.DOB;
        document.getElementById('editMasulPlaceOfBirth').value = masul.PlaceOfBirth || '';
        document.getElementById('editMasulPhone').value = masul.Phone;
        document.getElementById('editMasulEmail').value = masul.Email || '';
        document.getElementById('editMasulAddress').value = masul.Address;
        document.getElementById('editMasulState').value = masul.State;
        document.getElementById('editMasulLga').value = masul.LGA;
        document.getElementById('editMasulYear').value = masul.Year;
        document.getElementById('editMasulRank').value = masul.CurrentRank;

        await loadZonesForDropdowns(false);
        const zoneSelect = document.getElementById('editMasulZone');
        const branchSelect = document.getElementById('editMasulBranch');
        zoneSelect.value = masul.Zone;
        zoneSelect.dispatchEvent(new Event('change', { bubbles: true }));
        setTimeout(() => {
            branchSelect.value = masul.Branch;
        }, 1000);

        updateMasulRankOptions(masul.Gender);
        showModal('editMasulModal');
    } catch (err) {
        showMessage('Error', err.message);
    }
}

function updateMasulRankOptions(gender) {
    const rankSelect = document.getElementById('editMasulRank');
    const brotherRanks = ['Musa\'id','Areef','Muqaddam','Ra\'id','Raqeeb','Mulazim','Muhafiz','Ameed','Aqeeda','Qaid'];
    const sisterRanks = ['Musa\'ida','Areefa','Muqadama','Ra\'ida','Raqeeba','Mulazima','Muhafiza','Ameeda','Aqeeda','Qaida'];
    rankSelect.innerHTML = '<option value="">Select Rank</option>';
    const ranks = gender === 'Brother' ? brotherRanks : sisterRanks;
    ranks.forEach(rank => {
        rankSelect.innerHTML += `<option value="${rank}">${rank}</option>`;
    });
}

function closeEditMasulModal() {
    document.getElementById('editMasulModal').style.display = 'none';
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
        setText('statBakiyatullah', stats.levelCounts.Bakiyatullah);
        setText('statAnsarullah', stats.levelCounts.Ansarullah);
        setText('statGhalibun', stats.levelCounts.Ghalibun);
        setText('statXGhalibun', stats.levelCounts['X-Ghalibun']);

        updateLevelChart(stats.levelCounts);
        updateZoneChart(stats.zoneCounts);
        updateBranchChart(stats.branchCounts);

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
    if (!canvas) return;
    const existingChart = Chart.getChart(canvas);
    if (existingChart) existingChart.destroy();

    new Chart(canvas, {
        type: 'bar',
        data: {
            labels: Object.keys(levelCounts),
            datasets: [{
                label: 'Members',
                data: Object.values(levelCounts),
                backgroundColor: '#556B2F'
            }]
        },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });
}

function updateZoneChart(zoneCounts) {
    const canvas = document.getElementById('zoneChart');
    if (!canvas) return;
    const existingChart = Chart.getChart(canvas);
    if (existingChart) existingChart.destroy();

    new Chart(canvas, {
        type: 'pie',
        data: {
            labels: Object.keys(zoneCounts),
            datasets: [{
                data: Object.values(zoneCounts),
                backgroundColor: ['#556B2F', '#C9A87C', '#2F4F2F', '#DAA520', '#6B8E23', '#8B4513', '#5F9EA0']
            }]
        },
        options: { responsive: true }
    });
}

function updateBranchChart(branchCounts) {
    const canvas = document.getElementById('branchChart');
    if (!canvas) return;
    const existingChart = Chart.getChart(canvas);
    if (existingChart) existingChart.destroy();

    const sorted = Object.entries(branchCounts).sort((a,b) => b[1] - a[1]).slice(0,10);
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

// ==================== OPEN SPREADSHEET ====================
async function openSpreadsheet() {
    try {
        const result = await apiRequest('getSpreadsheetUrl', {}, currentUser);
        window.open(result.url, '_blank');
    } catch (err) {
        showMessage('Error', err.message);
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
            showSuccessModal(data.fullName, result.intizarId, result.recruitmentId, data.zone, data.branch);
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
            showSuccessModal(data.fullName, result.intizarId, result.masulRecruitmentId, data.zone, data.branch);
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

function showSuccessModal(name, intizarId, recruitmentId, zone, branch) {
    document.getElementById('generatedName').innerText = name;
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

// ==================== EDIT FORM EVENT LISTENERS ====================
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('editMemberForm')?.addEventListener('submit', async (e) => {
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
            loadMemberList(currentMemberPage, currentMemberSearch, currentMemberFilters);
        } catch (err) {
            showMessage('Error', err.message);
        }
    });

    document.getElementById('editMasulForm')?.addEventListener('submit', async (e) => {
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
            loadMasuls(currentMasulPage, currentMasulSearch, currentMasulFilters);
        } catch (err) {
            showMessage('Error', err.message);
        }
    });

    document.getElementById('editMasulGender')?.addEventListener('change', function() {
        updateMasulRankOptions(this.value);
    });
});

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
