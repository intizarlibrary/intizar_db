// ==================== SPREADSHEET MANAGEMENT ====================
const SPREADSHEET_NAME = 'Intizarul Imam Muntazar Database';

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // ID is invalid – will create a new one below
    }
  }
  // Create a new spreadsheet and store its ID
  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

// ==================== SHEET INITIALIZATION ====================
function ensureSheetsExist() {
  const ss = getSpreadsheet();

  const sheets = {
    Members: [
      'IntizarID', 'RecruitmentID', 'FullName', 'FatherName', 'Gender', 'DOB',
      'PlaceOfBirth', 'Phone', 'Email', 'Address', 'State', 'LGA', 'Zone',
      'Branch', 'Year', 'Level', 'PhotoURL', 'PromotionHistory', 'TransferHistory',
      'GuardianName', 'GuardianPhone', 'GuardianAddress'
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
    BranchCounters: ['BranchCode', 'LastSerial']
  };

  for (let name in sheets) {
    if (!ss.getSheetByName(name)) {
      const sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
    }
  }

  // Preload zones
  const zoneSheet = ss.getSheetByName('Zones');
  if (zoneSheet.getLastRow() <= 1) {
    const zones = [
      'SOKOTO ZONE', 'KADUNA ZONE', 'ABUJA ZONE', 'ZARIA ZONE', 'KANO ZONE',
      'BAUCHI ZONE', 'MALUMFASHI ZONE', 'NIGER ZONE', 'QUM ZONE'
    ];
    zones.forEach(zone => zoneSheet.appendRow([Utilities.getUuid(), zone, 'Active']));
  }

  // Preload branches
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

  // Default config
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

// ==================== ID GENERATION (atomic) ====================
function nextIntizarId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let curr = parseInt(getConfigValue('global_intizar') || '0');
    curr++;
    setConfig('global_intizar', curr.toString());
    return 'MTZR/' + curr.toString().padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Generates a member recruitment ID using the branch code and the year from the registration form.
 * @param {string} branchCode - The branch code.
 * @param {string|number} recruitmentYear - The year entered during registration (e.g., "2025").
 * @returns {string} Formatted recruitment ID.
 */
function nextMemberRecruitmentId(branchCode, recruitmentYear) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Members');
    const data = sheet.getDataRange().getValues();
    const year = recruitmentYear.toString().slice(-2);

    // Count all members belonging to this branch (ignore year)
    let serial = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][13] === branchCode) { // Column 13 = Branch
        serial++;
      }
    }
    serial++; // next number
    const padded = String(serial).padStart(3, '0');
    return `INT/${branchCode}/${year}/${padded}`;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Generates a Mas'ul recruitment ID using the branch code and the year from the registration form.
 * @param {string} branchCode - The branch code.
 * @param {string|number} recruitmentYear - The year entered during registration.
 * @returns {string} Formatted recruitment ID.
 */
function nextMasulRecruitmentId(branchCode, recruitmentYear) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet().getSheetByName('Masuls');
    const data = sheet.getDataRange().getValues();
    const year = recruitmentYear.toString().slice(-2);

    // Count ALL mas'uls (global)
    let serial = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) serial++; // any row with IntizarID
    }
    serial++;
    const padded = String(serial).padStart(3, '0');
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

// ==================== PHOTO UPLOAD TO DRIVE ====================
function savePhotoToDrive(base64Data, fileName) {
  let mimeType = 'image/jpeg'; // default
  try {
    const folderName = 'Intizarul_Photos';
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const extension = fileName.split('.').pop().toLowerCase();
    if (extension === 'png') {
      mimeType = 'image/png';
    } else if (extension === 'gif') {
      mimeType = 'image/gif';
    } else if (extension === 'jpg' || extension === 'jpeg') {
      mimeType = 'image/jpeg';
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Return direct image URL instead of viewer URL
    const fileId = file.getId();
    return `https://drive.google.com/uc?export=view&id=${fileId}`;
  } catch (e) {
    console.warn('Drive upload failed, using base64: ' + e.toString());
    return 'data:' + mimeType + ';base64,' + base64Data;
  }
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

  const intizarId = nextIntizarId();
  const recruitmentId = nextMemberRecruitmentId(data.branch, data.year);

  let photoURL = '';
  if (data.photoBase64) {
    if (!data.photoName || !data.photoName.match(/\.(jpg|jpeg|png|gif)$/i)) {
      throw new Error('Only image files (JPG, PNG, GIF) are allowed');
    }
    photoURL = savePhotoToDrive(data.photoBase64, data.photoName);
  }

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
    data.guardianAddress
  ];

  const sheet = getSpreadsheet().getSheetByName('Members');
  sheet.appendRow(row);

  logAudit(user.role + ':' + (user.branch || user.zone || 'Admin'), 'MEMBER_REGISTERED',
    `Intizar ID: ${intizarId}, Name: ${data.fullName}`);

  return { success: true, intizarId, recruitmentId };
}

function calculateAge(dobString) {
  const dob = new Date(dobString);
  const today = new Date();
  let age = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
  return age;
}

// ==================== MAS'UL REGISTRATION ====================
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
  if (data.source === 'X-Ghalibun') {
    if (!data.intizarId) throw new Error('Intizar ID required for X-Ghalibun');
    const memberSheet = getSpreadsheet().getSheetByName('Members');
    const memberData = memberSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < memberData.length; i++) {
      if (memberData[i][0] === data.intizarId) {
        if (memberData[i][15] !== 'X-Ghalibun') throw new Error('Member is not X-Ghalibun');
        found = true;
        intizarId = data.intizarId;
        originalMemberRecruitmentId = memberData[i][1];
        break;
      }
    }
    if (!found) throw new Error('Member not found');
  } else if (data.source === 'proposed') {
    intizarId = nextIntizarId();
  } else {
    throw new Error('Invalid source');
  }

  const brotherRanks = ['Musa\'id', 'Areef', 'Muqaddam', 'Ra\'id', 'Raqeeb', 'Mulazim', 'Muhafiz', 'Ameed', 'Aqeeda', 'Qaid'];
  const sisterRanks = ['Musa\'ida', 'Areefa', 'Muqadama', 'Ra\'ida', 'Raqeeba', 'Mulazima', 'Muhafiza', 'Ameeda', 'Aqeeda', 'Qaida'];
  const allowedEntryBrother = ['Musa\'id', 'Areef', 'Muqaddam'];
  const allowedEntrySister = ['Musa\'ida', 'Areefa', 'Muqadama'];

  if (data.gender === 'Brother') {
    if (!allowedEntryBrother.includes(data.currentRank)) throw new Error('Entry rank for Brother must be Musa\'id, Areef, or Muqaddam');
  } else if (data.gender === 'Sister') {
    if (!allowedEntrySister.includes(data.currentRank)) throw new Error('Entry rank for Sister must be Musa\'ida, Areefa, or Muqadama');
  } else {
    throw new Error('Invalid gender');
  }

  const masulRecruitmentId = nextMasulRecruitmentId(data.branch, data.year);

  let photoURL = '';
  if (data.photoBase64) {
    if (!data.photoName || !data.photoName.match(/\.(jpg|jpeg|png|gif)$/i)) {
      throw new Error('Only image files (JPG, PNG, GIF) are allowed');
    }
    photoURL = savePhotoToDrive(data.photoBase64, data.photoName);
  }

  const row = [
    intizarId,
    masulRecruitmentId,
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
    data.source,
    JSON.stringify([{ date: new Date(), rank: data.currentRank, action: 'Registered' }]),
    originalMemberRecruitmentId
  ];

  const sheet = getSpreadsheet().getSheetByName('Masuls');
  sheet.appendRow(row);

  logAudit('Admin', 'MASUL_REGISTERED',
    `Intizar ID: ${intizarId}, Name: ${data.fullName}, Source: ${data.source}`);

  return { success: true, intizarId, masulRecruitmentId, originalMemberRecruitmentId };
}

// ==================== GET MEMBERS (with search & filters) ====================
function getMembers(user, page = 1, pageSize = 50, search = '', filters = {}) {
  const sheet = getSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let allRows = data.slice(1); // exclude headers

  // Filter by role
  if (user.role === 'Zonal Mas\'ul') {
    allRows = allRows.filter(row => row[12] === user.zone);
  } else if (user.role === 'Branch Mas\'ul') {
    allRows = allRows.filter(row => row[13] === user.branchCode);
  }

  // Apply search filter
  if (search && search.trim() !== '') {
    const term = search.trim().toLowerCase();
    allRows = allRows.filter(row => {
      return (row[2] && row[2].toLowerCase().includes(term)) ||
             (row[0] && row[0].toLowerCase().includes(term)) ||
             (row[1] && row[1].toLowerCase().includes(term));
    });
  }

  // Apply advanced filters
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
  const end = start + pageSize;
  const paginatedRows = allRows.slice(start, end);

  const members = paginatedRows.map(row => {
    const member = {};
    headers.forEach((h, idx) => { member[h] = row[idx]; });
    return member;
  });

  return { success: true, members, total, page, pageSize };
}

// ==================== GET MASULS (with search & filters) ====================
function getMasuls(user, page = 1, pageSize = 50, search = '', filters = {}) {
  if (user.role !== 'Admin') throw new Error('Only Admin can view Mas\'ul list');

  const sheet = getSpreadsheet().getSheetByName('Masuls');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let allRows = data.slice(1).filter(row => row[0]); // exclude empty

  // Apply search filter
  if (search && search.trim() !== '') {
    const term = search.trim().toLowerCase();
    allRows = allRows.filter(row => {
      return (row[2] && row[2].toLowerCase().includes(term)) || // FullName
             (row[0] && row[0].toLowerCase().includes(term)) || // IntizarID
             (row[1] && row[1].toLowerCase().includes(term));   // MasulRecruitmentID
    });
  }

  // Apply advanced filters
  if (filters.rank) {
    allRows = allRows.filter(row => row[15] === filters.rank);
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
  const end = start + pageSize;
  const paginatedRows = allRows.slice(start, end);

  const masuls = paginatedRows.map(row => {
    const masul = {};
    headers.forEach((h, idx) => { masul[h] = row[idx]; });
    return masul;
  });

  return { success: true, masuls, total, page, pageSize };
}

// ==================== GET ZONES ====================
function getZones(user) {
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues();
  const zones = [];
  for (let i = 1; i < data.length; i++) {
    zones.push({ zoneId: data[i][0], zoneName: data[i][1], status: data[i][2] });
  }
  return { success: true, zones };
}

// ==================== GET BRANCHES ====================
function getBranches(user, zoneFilter) {
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues();
  const branches = [];

  if (user && user.role === 'Zonal Mas\'ul') {
    zoneFilter = user.zone;
  }

  for (let i = 1; i < data.length; i++) {
    if (!zoneFilter || data[i][2] === zoneFilter) {
      branches.push({
        branchCode: data[i][0],
        branchName: data[i][1],
        zone: data[i][2],
        status: data[i][3]
      });
    }
  }
  return { success: true, branches };
}

// ==================== PROMOTIONS ====================
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

  if (user.role === 'Zonal Mas\'ul' && member[12] !== user.zone) {
    throw new Error('You can only promote members in your own zone');
  }

  const levelOrder = ['Bakiyatullah', 'Ansarullah', 'Ghalibun', 'X-Ghalibun'];
  const currentLevel = member[15];
  const idx = levelOrder.indexOf(currentLevel);
  if (idx === -1) throw new Error('Invalid current level');
  if (idx === levelOrder.length - 1) throw new Error('Already at highest level');

  const newLevel = levelOrder[idx + 1];

  let history = [];
  try { history = JSON.parse(member[17] || '[]'); } catch (e) { history = []; }
  history.push({ date: new Date(), from: currentLevel, to: newLevel, by: user.role });

  sheet.getRange(rowIndex, 16).setValue(newLevel);
  sheet.getRange(rowIndex, 18).setValue(JSON.stringify(history));

  logAudit(user.role + ':' + (user.zone || user.branch || 'Admin'), 'MEMBER_PROMOTED',
    `Intizar ID: ${intizarId}, from ${currentLevel} to ${newLevel}`);

  return { success: true, newLevel };
}

function promoteMasul(intizarId, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can promote Mas\'ul');

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

  const brotherRanks = ['Musa\'id', 'Areef', 'Muqaddam', 'Ra\'id', 'Raqeeb', 'Mulazim', 'Muhafiz', 'Ameed', 'Aqeeda', 'Qaid'];
  const sisterRanks = ['Musa\'ida', 'Areefa', 'Muqadama', 'Ra\'ida', 'Raqeeba', 'Mulazima', 'Muhafiza', 'Ameeda', 'Aqeeda', 'Qaida'];
  const gender = masul[4];
  const rankOrder = gender === 'Brother' ? brotherRanks : sisterRanks;

  const currentRank = masul[15];
  const idx = rankOrder.indexOf(currentRank);
  if (idx === -1) throw new Error('Invalid current rank');
  if (idx === rankOrder.length - 1) throw new Error('Already at highest rank');

  const newRank = rankOrder[idx + 1];

  let history = [];
  try { history = JSON.parse(masul[18] || '[]'); } catch (e) { history = []; }
  history.push({ date: new Date(), from: currentRank, to: newRank, by: 'Admin' });

  sheet.getRange(rowIndex, 16).setValue(newRank);
  sheet.getRange(rowIndex, 19).setValue(JSON.stringify(history));

  logAudit('Admin', 'MASUL_PROMOTED', `Intizar ID: ${intizarId}, from ${currentRank} to ${newRank}`);

  return { success: true, newRank };
}

// ==================== TRANSFERS ====================
function transferMember(intizarId, newBranchCode, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can transfer members');
  if (!isValidBranchCode(newBranchCode)) throw new Error('Invalid branch code');

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

  const oldBranch = member[13];
  const newZone = getBranchZone(newBranchCode);

  let transferHistory = [];
  try { transferHistory = JSON.parse(member[18] || '[]'); } catch (e) { transferHistory = []; }
  transferHistory.push({ date: new Date(), fromBranch: oldBranch, toBranch: newBranchCode, by: 'Admin' });

  sheet.getRange(rowIndex, 14).setValue(newBranchCode);
  sheet.getRange(rowIndex, 13).setValue(newZone);
  sheet.getRange(rowIndex, 19).setValue(JSON.stringify(transferHistory));

  logAudit('Admin', 'MEMBER_TRANSFERRED', `Intizar ID: ${intizarId} to ${newBranchCode}`);

  return { success: true, newBranch: newBranchCode, newZone };
}

function transferMasul(intizarId, newBranchCode, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can transfer Mas\'ul');
  if (!isValidBranchCode(newBranchCode)) throw new Error('Invalid branch code');

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

  const oldBranch = masul[13];
  const newZone = getBranchZone(newBranchCode);

  sheet.getRange(rowIndex, 14).setValue(newBranchCode);
  sheet.getRange(rowIndex, 13).setValue(newZone);

  logAudit('Admin', 'MASUL_TRANSFERRED', `Intizar ID: ${intizarId} to ${newBranchCode}`);

  return { success: true, newBranch: newBranchCode, newZone };
}

// ==================== ZONE MANAGEMENT ====================
function addZone(zoneName, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can add zones');
  const sheet = getSpreadsheet().getSheetByName('Zones');
  sheet.appendRow([Utilities.getUuid(), zoneName, 'Active']);
  logAudit('Admin', 'ZONE_ADDED', zoneName);
  return { success: true };
}

function editZone(zoneId, newName, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can edit zones');
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === zoneId) {
      sheet.getRange(i + 1, 2).setValue(newName);
      logAudit('Admin', 'ZONE_EDITED', `Zone ${zoneId} renamed to ${newName}`);
      return { success: true };
    }
  }
  throw new Error('Zone not found');
}

function disableZone(zoneId, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can disable zones');
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === zoneId) {
      sheet.getRange(i + 1, 3).setValue('Disabled');
      logAudit('Admin', 'ZONE_DISABLED', `Zone ${zoneId}`);
      return { success: true };
    }
  }
  throw new Error('Zone not found');
}

function enableZone(zoneId, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can enable zones');
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === zoneId) {
      sheet.getRange(i + 1, 3).setValue('Active');
      logAudit('Admin', 'ZONE_ENABLED', `Zone ${zoneId}`);
      return { success: true };
    }
  }
  throw new Error('Zone not found');
}

// ==================== BRANCH MANAGEMENT ====================
function addBranch(branchName, zoneName, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can add branches');
  if (!isValidZone(zoneName)) throw new Error('Zone is invalid or inactive');

  const base = branchName.replace(/[^A-Za-z]/g, '').substring(0, 2).toUpperCase();
  let code = base;
  let counter = 1;
  const branchSheet = getSpreadsheet().getSheetByName('Branches');
  const existingCodes = branchSheet.getRange('A:A').getValues().flat();
  while (existingCodes.includes(code)) {
    code = base + counter;
    counter++;
  }

  branchSheet.appendRow([code, branchName, zoneName, 'Active']);
  logAudit('Admin', 'BRANCH_ADDED', `${branchName} (${code}) in ${zoneName}`);
  return { success: true, branchCode: code };
}

function editBranch(branchCode, newName, newZone, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can edit branches');
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === branchCode) {
      if (newName) sheet.getRange(i + 1, 2).setValue(newName);
      if (newZone) {
        if (!isValidZone(newZone)) throw new Error('Invalid zone');
        sheet.getRange(i + 1, 3).setValue(newZone);
      }
      logAudit('Admin', 'BRANCH_EDITED', `Branch ${branchCode} updated`);
      return { success: true };
    }
  }
  throw new Error('Branch not found');
}

function disableBranch(branchCode, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can disable branches');
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === branchCode) {
      sheet.getRange(i + 1, 4).setValue('Disabled');
      logAudit('Admin', 'BRANCH_DISABLED', `Branch ${branchCode}`);
      return { success: true };
    }
  }
  throw new Error('Branch not found');
}

function enableBranch(branchCode, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can enable branches');
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === branchCode) {
      sheet.getRange(i + 1, 4).setValue('Active');
      logAudit('Admin', 'BRANCH_ENABLED', `Branch ${branchCode}`);
      return { success: true };
    }
  }
  throw new Error('Branch not found');
}

// ==================== AUDIT LOG ====================
function getAuditLog(user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can view audit log');
  const sheet = getSpreadsheet().getSheetByName('AuditLog');
  const data = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    logs.push({
      timestamp: data[i][0],
      user: data[i][1],
      action: data[i][2],
      details: data[i][3]
    });
  }
  return { success: true, logs };
}

// ==================== CONFIG ====================
function getConfig(key, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can view config');
  return { success: true, value: getConfigValue(key) };
}

function updateConfig(key, value, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can edit config');
  setConfig(key, value);
  logAudit('Admin', 'CONFIG_UPDATED', `${key} = ${value}`);
  return { success: true };
}

// ==================== EXPORT DATA ====================
function exportData(type, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can export data');
  const sheetName = type === 'members' ? 'Members' : 'Masuls';
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const csv = data.map(row => row.join(',')).join('\n');
  logAudit('Admin', 'EXPORT_' + type.toUpperCase(), `Exported ${sheetName} data`);
  return { success: true, csv, filename: sheetName + '_' + new Date().toISOString() + '.csv' };
}

// ==================== GET SINGLE MEMBER ====================
function getMember(intizarId, user) {
  const sheet = getSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === intizarId) {
      if (user.role === 'Zonal Mas\'ul' && data[i][12] !== user.zone) {
        throw new Error('Access denied');
      }
      if (user.role === 'Branch Mas\'ul' && data[i][13] !== user.branchCode) {
        throw new Error('Access denied');
      }
      const member = {};
      headers.forEach((h, idx) => { member[h] = data[i][idx]; });
      return { success: true, member };
    }
  }
  throw new Error('Member not found');
}

// ==================== GET SINGLE MASUL ====================
function getMasul(intizarId, user) {
  if (user.role !== 'Admin') throw new Error('Only Admin can view Mas\'ul details');
  const sheet = getSpreadsheet().getSheetByName('Masuls');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === intizarId) {
      const masul = {};
      headers.forEach((h, idx) => { masul[h] = data[i][idx]; });
      return { success: true, masul };
    }
  }
  throw new Error('Mas\'ul not found');
}

// ==================== DASHBOARD STATISTICS ====================
function getDashboardStats(user) {
  const membersSheet = getSpreadsheet().getSheetByName('Members');
  const masulsSheet = getSpreadsheet().getSheetByName('Masuls');

  let membersData = membersSheet.getDataRange().getValues().slice(1);
  let masulsData = masulsSheet.getDataRange().getValues().slice(1);

  // Apply role‑based filtering
  if (user.role === 'Zonal Mas\'ul') {
    membersData = membersData.filter(row => row[12] === user.zone);
    masulsData = masulsData.filter(row => row[12] === user.zone);
  } else if (user.role === 'Branch Mas\'ul') {
    membersData = membersData.filter(row => row[13] === user.branchCode);
    masulsData = masulsData.filter(row => row[13] === user.branchCode);
  }

  const totalMembers = membersData.length;
  const totalMasuls = masulsData.length;
  const totalCombined = totalMembers + totalMasuls;

  // Gender breakdown (all)
  const brothers = membersData.filter(row => row[4] === 'Brother').length +
                   masulsData.filter(row => row[4] === 'Brother').length;
  const sisters = membersData.filter(row => row[4] === 'Sister').length +
                  masulsData.filter(row => row[4] === 'Sister').length;

  // Gender breakdown by role
  const brothersMembers = membersData.filter(row => row[4] === 'Brother').length;
  const sistersMembers = membersData.filter(row => row[4] === 'Sister').length;
  const brothersMasuls = masulsData.filter(row => row[4] === 'Brother').length;
  const sistersMasuls = masulsData.filter(row => row[4] === 'Sister').length;

  // Member levels
  const levelCounts = {
    Bakiyatullah: 0,
    Ansarullah: 0,
    Ghalibun: 0,
    'X-Ghalibun': 0
  };
  membersData.forEach(row => {
    const level = row[15];
    if (levelCounts.hasOwnProperty(level)) levelCounts[level]++;
  });

  return {
    success: true,
    stats: {
      totalCombined,
      totalMembers,
      totalMasuls,
      brothers,
      sisters,
      brothersMembers,
      sistersMembers,
      brothersMasuls,
      sistersMasuls,
      levelCounts
    }
  };
}

function getZoneStats(user) {
  if (user.role !== 'Admin') throw new Error('Permission denied');
  const zonesSheet = getSpreadsheet().getSheetByName('Zones');
  const membersSheet = getSpreadsheet().getSheetByName('Members');
  const zones = zonesSheet.getDataRange().getValues().slice(1);
  const members = membersSheet.getDataRange().getValues().slice(1);

  const stats = zones.map(zone => {
    const zoneName = zone[1];
    const zoneMembers = members.filter(m => m[12] === zoneName);
    return {
      zone: zoneName,
      total: zoneMembers.length,
      brothers: zoneMembers.filter(m => m[4] === 'Brother').length,
      sisters: zoneMembers.filter(m => m[4] === 'Sister').length
    };
  });
  // Include zones with zero members
  return { success: true, stats };
}

function getBranchStats(user) {
  if (user.role !== 'Admin') throw new Error('Permission denied');
  const branchesSheet = getSpreadsheet().getSheetByName('Branches');
  const membersSheet = getSpreadsheet().getSheetByName('Members');
  const branches = branchesSheet.getDataRange().getValues().slice(1);
  const members = membersSheet.getDataRange().getValues().slice(1);

  const stats = branches.map(branch => {
    const branchCode = branch[0];
    const branchName = branch[1];
    const zone = branch[2];
    const branchMembers = members.filter(m => m[13] === branchCode);
    return {
      branchCode,
      branchName,
      zone,
      total: branchMembers.length,
      brothers: branchMembers.filter(m => m[4] === 'Brother').length,
      sisters: branchMembers.filter(m => m[4] === 'Sister').length
    };
  });
  // Include branches with zero members
  return { success: true, stats };
}

// ==================== FILTER OPTIONS (distinct values) ====================
function getDistinctBranches() {
  const sheet = getSpreadsheet().getSheetByName('Branches');
  const data = sheet.getDataRange().getValues().slice(1);
  const branches = [...new Set(data.map(row => row[1]))]; // branch names
  return branches;
}

function getDistinctZones() {
  const sheet = getSpreadsheet().getSheetByName('Zones');
  const data = sheet.getDataRange().getValues().slice(1);
  return data.map(row => row[1]).filter((v,i,a) => a.indexOf(v) === i);
}

function getDistinctLevels() {
  const sheet = getSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues().slice(1);
  return [...new Set(data.map(row => row[15]))]; // Level column
}

function getDistinctGenders() {
  return ['Brother', 'Sister'];
}

function getDistinctRanks() {
  const sheet = getSpreadsheet().getSheetByName('Masuls');
  const data = sheet.getDataRange().getValues().slice(1);
  return [...new Set(data.map(row => row[15]))]; // CurrentRank column
}

// ==================== MANUAL INITIALIZATION FUNCTIONS ====================
function initializeSheets() {
  ensureSheetsExist();
  console.log('Sheets initialized.');
}

function initializeDriveFolder() {
  const folderName = 'Intizarul_Photos';
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    DriveApp.createFolder(folderName);
    console.log('Folder created: ' + folderName);
  } else {
    console.log('Folder already exists.');
  }
}

function initializeConfig() {
  setConfig('admin_code', 'Muntazir@Global');
  setConfig('access_prefix', 'Muntazir@');
  setConfig('global_intizar', '0');
  setConfig('global_masul_serial', '0');
  console.log('Config initialized.');
}

function fullSetup() {
  initializeSheets();
  initializeDriveFolder();
  initializeConfig();
  console.log('System ready.');
}
