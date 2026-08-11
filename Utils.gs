/**
 * INTIZARUL IMAMUL MUNTAZAR – Backend Utility & Business Logic (Utils.gs)
 * Handles Google Sheets operations, normalization, ID generation,
 * permissions, audit logging, search, statistics, and workflows.
 */

const SPREADSHEET_NAME = 'Intizarul Imam Muntazar Database';

// ==================== SPREADSHEET MANAGEMENT ====================
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // ID is invalid – create new below
    }
  }
  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

// ==================== SAFE NORMALIZATION ====================
function safeNormalize(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase();
}

// ==================== SHEET INITIALIZATION ====================
function ensureSheetsExist() {
  const ss = getSpreadsheet();

  const sheets = {
    Members: [
      'IntizarID', 'RecruitmentID', 'FullName', 'FatherName', 'Gender', 'DOB',
      'PlaceOfBirth', 'Phone', 'Email', 'Address', 'State', 'LGA', 'Zone',
      'Branch', 'Year', 'Level', 'PhotoURL', 'PromotionHistory', 'TransferHistory',
      'GuardianName', 'GuardianPhone', 'GuardianAddress', 'Status'
    ],
    Masuls: [
      'IntizarID', 'MasulRecruitmentID', 'FullName', 'FatherName', 'Gender', 'DOB',
      'PlaceOfBirth', 'Phone', 'Email', 'Address', 'State', 'LGA', 'Zone',
      'Branch', 'Year', 'CurrentRank', 'PhotoURL', 'Source', 'PromotionHistory',
      'OriginalMemberRecruitmentID'
    ],
    Zones: ['ZoneID', 'ZoneName', 'Status'],
    Branches: ['BranchCode', 'BranchName', 'Zone', 'Status'],
    Config: ['Key', 'Value'],
    AuditLog: ['Timestamp', 'User', 'Action', 'Details'],
    BranchCounters: ['BranchCode', 'LastSerial'],
    PromotionHistory: ['Timestamp', 'IntizarID', 'Type', 'From', 'To', 'By', 'Details'],
    TransferHistory:  ['Timestamp', 'IntizarID', 'Type', 'FromBranch', 'ToBranch', 'Zone', 'By']
  };

  for (let name in sheets) {
    if (!ss.getSheetByName(name)) {
      const sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
    }
  }

  // Preload default zones if empty
  const zoneSheet = ss.getSheetByName('Zones');
  if (zoneSheet.getLastRow() <= 1) {
    const zones = [
      'SOKOTO ZONE', 'KADUNA ZONE', 'ABUJA ZONE', 'ZARIA ZONE', 'KANO ZONE',
      'BAUCHI ZONE', 'MALUMFASHI ZONE', 'NIGER ZONE', 'QUM ZONE'
    ];
    zones.forEach(zone => zoneSheet.appendRow([Utilities.getUuid(), zone, 'Active']));
  }

  // Preload default branches if empty
  const branchSheet = ss.getSheetByName('Branches');
  if (branchSheet.getLastRow() <= 1) {
    const branches = [
      ['SK', 'Sokoto', 'SOKOTO ZONE', 'Active'],
      ['MFR', 'Mafara', 'SOKOTO ZONE', 'Active'],
      ['YR', 'Yaure', 'SOKOTO ZONE', 'Active'],
      ['IL', 'Ilela', 'SOKOTO ZONE', 'Active'],
      ['ZR', 'Zuru', 'SOKOTO ZONE', 'Active'],
      ['YB', 'Yabo', 'SOKOTO ZONE', 'Active'],
      ['KD', 'Kaduna', 'KADUNA ZONE', 'Active'],
      ['JJ', 'Jaji', 'KADUNA ZONE', 'Active'],
      ['MJ', 'Mjos', 'KADUNA ZONE', 'Active'],
      ['MRB', 'Maraba', 'ABUJA ZONE', 'Active'],
      ['LF', 'Lafia', 'ABUJA ZONE', 'Active'],
      ['KF', 'Keffi/Doma', 'ABUJA ZONE', 'Active'],
      ['MN', 'Minna', 'ABUJA ZONE', 'Active'],
      ['SLJ', 'Suleja', 'ABUJA ZONE', 'Active'],
      ['ZAR', 'Zaria', 'ZARIA ZONE', 'Active'],
      ['DJ', 'Danja', 'ZARIA ZONE', 'Active'],
      ['DW', 'D/Wai', 'ZARIA ZONE', 'Active'],
      ['KUD', 'Kudan', 'ZARIA ZONE', 'Active'],
      ['SOB', 'Soba', 'ZARIA ZONE', 'Active'],
      ['KN', 'Kano', 'KANO ZONE', 'Active'],
      ['KZ', 'Kazaure', 'KANO ZONE', 'Active'],
      ['PT', 'Potiskum', 'KANO ZONE', 'Active'],
      ['GSW', 'Gashuwa', 'KANO ZONE', 'Active'],
      ['BAU', 'Bauchi', 'BAUCHI ZONE', 'Active'],
      ['GM', 'Gombe', 'BAUCHI ZONE', 'Active'],
      ['AZ', 'Azare', 'BAUCHI ZONE', 'Active'],
      ['JS', 'Jos', 'BAUCHI ZONE', 'Active'],
      ['MLF', 'Malumfashi', 'MALUMFASHI ZONE', 'Active'],
      ['BK', 'Bakori', 'MALUMFASHI ZONE', 'Active'],
      ['KT', 'Katsina', 'MALUMFASHI ZONE', 'Active'],
      ['NY', 'Niyame', 'NIGER ZONE', 'Active'],
      ['MRD', 'Maradi', 'NIGER ZONE', 'Active'],
      ['QM', 'Qum', 'QUM ZONE', 'Active']
    ];
    branches.forEach(b => branchSheet.appendRow(b));
  }

  // Default configuration keys
  if (!getConfigValue('access_prefix')) setConfig('access_prefix', 'Muntazir@');
  if (!getConfigValue('admin_code')) setConfig('admin_code', 'Muntazir@Global');
  if (!getConfigValue('global_intizar')) setConfig('global_intizar', '0');
  if (!getConfigValue('global_masul_serial')) setConfig('global_masul_serial', '0');
}

// ==================== CONFIG HELPERS ====================
function getConfigValue(key) {
  const sheet = getSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function setConfig(key, value) {
  const sheet = getSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ==================== AUDIT LOG ====================
function logAudit(user, action, details) {
  const sheet = getSpreadsheet().getSheetByName('AuditLog');
  sheet.appendRow([new Date(), user, action, details]);
}

// ==================== MEMBER RECRUITMENT ID ====================
function nextMemberRecruitmentId(branchCode, recruitmentYear) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Members');
    const data = sheet.getDataRange().getValues();
    const year = recruitmentYear.toString().slice(-2);

    let serial = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][13] === branchCode && data[i][22] === 'Active') {
        serial++;
      }
    }
    serial++;
    const padded = String(serial).padStart(4, '0');
    return `INT/${branchCode}/${year}/${padded}`;
  } finally {
    lock.releaseLock();
  }
}

// ==================== MAS'UL RECRUITMENT ID ====================
function nextMasulRecruitmentId(branchCode, recruitmentYear) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Masuls');
    const data = sheet.getDataRange().getValues();
    const year = recruitmentYear.toString().slice(-2);

    let serial = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) serial++;
    }
    serial++;
    const padded = String(serial).padStart(5, '0');
    return `IIM/${branchCode}/${year}/${padded}`;
  } finally {
    lock.releaseLock();
  }
}

// ==================== VALIDATION HELPERS ====================
function isValidBranchCode(branchCode) {
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === branchCode && data[i][3] === 'Active') return true;
  }
  return false;
}

function getBranchZone(branchCode) {
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === branchCode) return data[i][2];
  }
  return null;
}

function isValidZone(zoneName) {
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === zoneName && data[i][2] === 'Active') return true;
  }
  return false;
}

function getZoneId(zoneName) {
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === zoneName) return data[i][0];
  }
  return null;
}

function calculateAge(dobString) {
  const dob = new Date(dobString);
  const today = new Date();
  let age = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
  return age;
}

// ==================== LOGIN ====================
function login(role, code) {
  const prefix = getConfigValue('access_prefix') || 'Muntazir@';

  if (role === 'Admin') {
    const adminCode = getConfigValue('admin_code') || 'Muntazir@Global';
    if (code !== adminCode) throw new Error('Invalid admin code');
    logAudit('Admin', 'LOGIN_SUCCESS', 'Admin login');
    return { success: true, user: { role } };
  }

  if (role === 'Zonal Mas\'ul') {
    if (!code.startsWith(prefix)) throw new Error('Invalid code format');
    const zoneName = code.substring(prefix.length);
    if (!isValidZone(zoneName)) throw new Error('Invalid or inactive zone');
    logAudit('Zonal Mas\'ul:' + zoneName, 'LOGIN_SUCCESS', 'Zonal login');
    return { success: true, user: { role, zone: zoneName, zoneId: getZoneId(zoneName) } };
  }

  if (role === 'Branch Mas\'ul') {
    if (!code.startsWith(prefix)) throw new Error('Invalid code format');
    const branchName = code.substring(prefix.length);
    const sheet = getSpreadsheet().getSheetByName('Branches');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === branchName && data[i][3] === 'Active') {
        logAudit('Branch Mas\'ul:' + branchName, 'LOGIN_SUCCESS', 'Branch login');
        return {
          success: true,
          user: { role, branch: branchName, branchCode: data[i][0], zone: data[i][2] }
        };
      }
    }
    throw new Error('Invalid or inactive branch');
  }

  throw new Error('Invalid role');
}

// ==================== MEMBER REGISTRATION ====================
function registerMember(data, user) {
  if (user.role === 'Branch Mas\'ul' && data.branch !== user.branchCode) {
    throw new Error('You can only register members in your own branch');
  }
  if (!['Admin', 'Branch Mas\'ul'].includes(user.role)) {
    throw new Error('Insufficient permissions');
  }

  const required = [
    'fullName', 'fatherName', 'gender', 'dob', 'phone', 'address', 'state', 'lga',
    'zone', 'branch', 'year', 'entryLevel',
    'guardianName', 'guardianPhone', 'guardianAddress'
  ];
  for (let f of required) {
    if (!data[f]) throw new Error(`Missing required field: ${f}`);
  }

  const age = calculateAge(data.dob);
  if (age < 7) throw new Error('Member must be at least 7 years old');

  const branchZone = getBranchZone(data.branch);
  if (branchZone !== data.zone) throw new Error('Branch does not belong to selected zone');

  const allowedLevels = ['Bakiyatullah', 'Ansarullah', 'Ghalibun'];
  if (!allowedLevels.includes(data.entryLevel)) {
    throw new Error('Entry level must be Bakiyatullah, Ansarullah, or Ghalibun');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Members');
    const existingRows = sheet.getDataRange().getValues();
    const dataRows = existingRows.slice(1);

    // Safe duplicate check
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (safeNormalize(row[2]) === safeNormalize(data.fullName) &&
          safeNormalize(row[3]) === safeNormalize(data.fatherName) &&
          safeNormalize(row[4]) === safeNormalize(data.gender) &&
          safeNormalize(row[5]) === safeNormalize(data.dob) &&
          safeNormalize(row[7]) === safeNormalize(data.phone) &&
          safeNormalize(row[13]) === safeNormalize(data.branch)) {
        throw new Error('Duplicate registration detected. This person is already registered with Intizar ID: ' + row[0]);
      }
    }

    let currentIntizar = parseInt(getConfigValue('global_intizar') || '0');
    let nextIntizar = currentIntizar + 1;
    const intizarId = 'MTZR/' + nextIntizar.toString().padStart(5, '0');
    
    const recruitmentId = nextMemberRecruitmentId(data.branch, data.year);
    const photoURL = data.photoURL || '';

    const row = [
      intizarId,
      recruitmentId,
      data.fullName,
      data.fatherName,
      data.gender,
      data.dob,
      data.placeOfBirth || '',
      data.phone,
      data.email || '',
      data.address,
      data.state,
      data.lga,
      data.zone,
      data.branch,
      data.year,
      data.entryLevel,
      photoURL,
      JSON.stringify([{ date: new Date(), level: data.entryLevel, action: 'Registered' }]),
      '[]',
      data.guardianName,
      data.guardianPhone,
      data.guardianAddress,
      'Active'
    ];

    sheet.appendRow(row);
    setConfig('global_intizar', nextIntizar.toString());

    logAudit(user.role + ':' + (user.branch || user.zone || 'Admin'), 'MEMBER_REGISTERED',
      `Intizar ID: ${intizarId}, Name: ${data.fullName}`);

    return { success: true, intizarId, recruitmentId };
  } finally {
    lock.releaseLock();
  }
}

// ==================== MAS'UL REGISTRATION (Supports 3 Sources: proposed, X-Ghalibun, Graduate) ====================
function registerMasul(data, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can register Mas\'ul');

  const required = [
    'fullName', 'fatherName', 'gender', 'dob', 'phone', 'address', 'state', 'lga',
    'zone', 'branch', 'year', 'currentRank', 'source'
  ];
  for (let f of required) {
    if (!data[f]) throw new Error(`Missing required field: ${f}`);
  }

  if (calculateAge(data.dob) < 18) throw new Error('Mas\'ul must be at least 18 years old');

  const branchZone = getBranchZone(data.branch);
  if (branchZone !== data.zone) throw new Error('Branch does not belong to selected zone');

  let intizarId = '';
  let originalMemberRecruitmentId = '';
  let memberRowIndex = -1;

  // Safe normalized source checking
  const normSource = safeNormalize(data.source);

  // Handles internal candidate sources (X-Ghalibun or Graduate)
  if (normSource === 'x-ghalibun' || normSource === 'xghalibun' || normSource === 'graduate') {
    if (!data.intizarId) throw new Error(`Intizar ID required for candidate from ${data.source}`);
    
    const memberSheet = getSpreadsheet().getSheetByName('Members');
    const memberData = memberSheet.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < memberData.length; i++) {
      if (safeNormalize(memberData[i][0]) === safeNormalize(data.intizarId)) {
        found = true;
        intizarId = memberData[i][0]; // PRESERVE PERMANENT MTZR ID EXACTLY
        originalMemberRecruitmentId = memberData[i][1]; // PRESERVE INT/... MEMBER ID
        memberRowIndex = i + 1;
        break;
      }
    }
    if (!found) throw new Error('Member with Intizar ID ' + data.intizarId + ' not found in Members registry');

    // Mark member as 'Mas\'ul' in Members sheet (Archived from active member list)
    const headers = memberData[0];
    const statusColIdx = headers.indexOf('Status');
    if (statusColIdx !== -1) {
      memberSheet.getRange(memberRowIndex, statusColIdx + 1).setValue('Mas\'ul');
    }

    // Log in PromotionHistory
    const promoHistorySheet = getSpreadsheet().getSheetByName('PromotionHistory');
    promoHistorySheet.appendRow([
      new Date(), intizarId, 'Member',
      data.source, 'Mas\'ul',
      user.role,
      `Archived as Mas'ul (Registered from ${data.source})`
    ]);

  } else if (normSource === 'proposed' || normSource.includes('proposed')) {
    // External candidate pathway: Generate new permanent MTZR ID
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      let currentIntizar = parseInt(getConfigValue('global_intizar') || '0');
      let nextIntizar = currentIntizar + 1;
      intizarId = 'MTZR/' + nextIntizar.toString().padStart(5, '0');
      setConfig('global_intizar', nextIntizar.toString());
    } finally {
      lock.releaseLock();
    }
  } else {
    throw new Error('Invalid source specified: ' + data.source);
  }

  // Generate new Mas'ul Recruitment ID (IIM/BRANCH/YEAR/XXXXX)
  const masulRecruitmentId = nextMasulRecruitmentId(data.branch, data.year);
  const photoURL = data.photoURL || '';

  const row = [
    intizarId,                      // Permanent MTZR ID
    masulRecruitmentId,            // New IIM/... Mas'ul Recruitment ID
    data.fullName,
    data.fatherName,
    data.gender,
    data.dob,
    data.placeOfBirth || '',
    data.phone,
    data.email || '',
    data.address,
    data.state,
    data.lga,
    data.zone,
    data.branch,
    data.year,
    data.currentRank,
    photoURL,
    data.source,                    // 'proposed', 'X-Ghalibun', or 'Graduate'
    JSON.stringify([{ date: new Date(), rank: data.currentRank, action: 'Registered' }]),
    originalMemberRecruitmentId     // Preserved original INT/... Member Recruitment ID
  ];

  const sheet = getSpreadsheet().getSheetByName('Masuls');
  sheet.appendRow(row);

  logAudit('Admin', 'MASUL_REGISTERED',
    `Intizar ID: ${intizarId}, Masul ID: ${masulRecruitmentId}, Name: ${data.fullName}, Source: ${data.source}`);

  return { success: true, intizarId, masulRecruitmentId, originalMemberRecruitmentId };
}

// ==================== GET MEMBERS (Excludes Archived) ====================
function getMembers(user, page = 1, pageSize = 50, search = '', filters = {}) {
  const sheet = getSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let allRows = data.slice(1);

  // Role permissions
  if (user.role === 'Zonal Mas\'ul') {
    allRows = allRows.filter(row => row[12] === user.zone);
  } else if (user.role === 'Branch Mas\'ul') {
    allRows = allRows.filter(row => row[13] === user.branchCode);
  }

  // Exclude archived members who transitioned to Mas'ul
  allRows = allRows.filter(row => row[22] !== 'Mas\'ul');

  // Multi-field safe search
  if (search && search.trim() !== '') {
    const term = safeNormalize(search);
    allRows = allRows.filter(row => {
      return safeNormalize(row[0]).includes(term) || // Intizar ID
             safeNormalize(row[1]).includes(term) || // Recruitment ID
             safeNormalize(row[2]).includes(term) || // Full Name
             safeNormalize(row[3]).includes(term) || // Father Name
             safeNormalize(row[7]).includes(term) || // Phone
             safeNormalize(row[12]).includes(term) || // Zone
             safeNormalize(row[13]).includes(term) || // Branch
             safeNormalize(row[15]).includes(term);   // Level
    });
  }

  // Filters
  if (filters.level) {
    allRows = allRows.filter(row => row[15] === filters.level);
  }
  if (filters.gender) {
    allRows = allRows.filter(row => row[4] === filters.gender);
  }
  if (filters.branch) {
    allRows = allRows.filter(row => row[13] === filters.branch);
  }
  if (filters.zone) {
    allRows = allRows.filter(row => row[12] === filters.zone);
  }

  const total = allRows.length;
  const start = (page - 1) * pageSize;
  const paginatedRows = allRows.slice(start, start + pageSize);

  const members = paginatedRows.map(row => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    return obj;
  });

  return { success: true, members, total, page, pageSize };
}

// ==================== GET GRADUATES (Al-Mahdi Community) ====================
function getGraduates(user, page = 1, pageSize = 50, search = '', filters = {}) {
  // Uses existing getMembers logic filtered specifically for Level = 'Graduate'
  const gradFilters = Object.assign({}, filters, { level: 'Graduate' });
  return getMembers(user, page, pageSize, search, gradFilters);
}

// ==================== PROMOTION WORKFLOW ====================
function promoteMember(intizarId, user) {
  const sheet = getSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1, member = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === intizarId) {
      rowIndex = i + 1;
      member = data[i];
      break;
    }
  }
  if (!member) throw new Error('Member not found');
  if (member[22] && member[22] !== 'Active') {
    throw new Error('Member is archived and cannot be promoted');
  }

  if (user.role === 'Zonal Mas\'ul' && member[12] !== user.zone) {
    throw new Error('You can only promote members in your own zone');
  }

  // Level Progression: Bakiyatullah -> Ansarullah -> Ghalibun -> X-Ghalibun -> Graduate
  const levelOrder = ['Bakiyatullah', 'Ansarullah', 'Ghalibun', 'X-Ghalibun', 'Graduate'];
  const currentLevel = member[15];
  const idx = levelOrder.indexOf(currentLevel);

  if (idx === -1) throw new Error('Invalid current level');
  if (idx === levelOrder.length - 1) throw new Error('Already at highest stage (Graduate - Al-Mahdi Community)');

  const newLevel = levelOrder[idx + 1];

  let history = [];
  try { history = JSON.parse(member[17] || '[]'); } catch (e) { history = []; }
  history.push({ date: new Date(), from: currentLevel, to: newLevel, by: user.role });

  sheet.getRange(rowIndex, 16).setValue(newLevel);
  sheet.getRange(rowIndex, 18).setValue(JSON.stringify(history));

  const promoHist = getSpreadsheet().getSheetByName('PromotionHistory');
  promoHist.appendRow([
    new Date(), intizarId, 'Member',
    currentLevel, newLevel,
    user.role,
    `Promoted from ${currentLevel} to ${newLevel}`
  ]);

  logAudit(user.role + ':' + (user.zone || user.branch || 'Admin'), 'MEMBER_PROMOTED',
    `Intizar ID: ${intizarId}, from ${currentLevel} to ${newLevel}`);

  return { success: true, newLevel };
}

// Direct promotion to Graduate for X-Ghalibun
function promoteToGraduate(intizarId, user) {
  const sheet = getSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1, member = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === intizarId) {
      rowIndex = i + 1;
      member = data[i];
      break;
    }
  }
  if (!member) throw new Error('Member not found');
  if (member[15] !== 'X-Ghalibun') {
    throw new Error('Only members at X-Ghalibun level can progress to Graduate (Al-Mahdi Community)');
  }

  return promoteMember(intizarId, user);
}

// Promote Mas'ul Rank
function promoteMasul(intizarId, user) {
  const sheet = getSpreadsheet().getSheetByName('Masuls');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1, masul = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === intizarId) {
      rowIndex = i + 1;
      masul = data[i];
      break;
    }
  }
  if (!masul) throw new Error('Mas\'ul not found');

  if (user.role === 'Zonal Mas\'ul' && masul[12] !== user.zone) {
    throw new Error('You can only promote Mas\'ulin in your own zone');
  }

  const gender = masul[4];
  const currentRank = masul[15];

  const brotherRanks = ['Musa\'id', 'Areef', 'Muqaddam', 'Ra\'id', 'Raqeeb', 'Mulazim', 'Muhafiz', 'Ameed', 'Aqeeda', 'Qaid'];
  const sisterRanks = ['Musa\'ida', 'Areefa', 'Muqadama', 'Ra\'ida', 'Raqeeba', 'Mulazima', 'Muhafiza', 'Ameeda', 'Aqeeda', 'Qaida'];

  const ranks = (gender === 'Sister') ? sisterRanks : brotherRanks;
  let idx = ranks.indexOf(currentRank);

  if (idx === -1) {
    idx = 0;
  }

  if (idx >= ranks.length - 1) {
    throw new Error(`Mas'ul is already at the highest rank (${ranks[ranks.length - 1]})`);
  }

  const newRank = ranks[idx + 1];

  let history = [];
  try { history = JSON.parse(masul[18] || '[]'); } catch (e) { history = []; }
  history.push({ date: new Date(), from: currentRank, to: newRank, by: user.role });

  sheet.getRange(rowIndex, 16).setValue(newRank);
  sheet.getRange(rowIndex, 19).setValue(JSON.stringify(history));

  const promoHist = getSpreadsheet().getSheetByName('PromotionHistory');
  promoHist.appendRow([
    new Date(), intizarId, 'Masul',
    currentRank, newRank,
    user.role,
    `Mas'ul rank promoted from ${currentRank} to ${newRank}`
  ]);

  logAudit(user.role + ':' + (user.zone || user.branch || 'Admin'), 'MASUL_PROMOTED',
    `Intizar ID: ${intizarId}, from ${currentRank} to ${newRank}`);

  return { success: true, newRank };
}

// ==================== EXPORT CSV DATA ====================
function exportData(type, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can export data');
  let sheetName = 'Members';
  if (type === 'masuls') sheetName = 'Masuls';
  
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  let data = sheet.getDataRange().getValues();
  if (type === 'graduates') {
    const headers = data[0];
    const gradRows = data.slice(1).filter(r => r[15] === 'Graduate' && r[22] === 'Active');
    data = [headers, ...gradRows];
  }

  if (data.length <= 1) throw new Error('No data available to export');

  const csvRows = data.map(row => 
    row.map(cell => {
      const str = String(cell);
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  );

  const csv = csvRows.join('\n');
  const filename = `${sheetName}_${new Date().toISOString().replace(/[:.]/g, '-')}.csv`;

  logAudit('Admin', 'EXPORT_' + type.toUpperCase(), `Exported ${sheetName} CSV (${data.length - 1} records)`);
  return { success: true, csv, filename };
}
