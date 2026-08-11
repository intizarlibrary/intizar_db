/**
 * INTIZARUL IMAMUL MUNTAZAR – Backend Dispatcher (Code.gs)
 * Entry point for doGet and doPost HTTPS requests from the web frontend.
 * Routes JSON payloads to functions in Utils.gs.
 */

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    ensureSheetsExist();

    let params = {};
    if (e && e.parameter && e.parameter.payload) {
      params = JSON.parse(e.parameter.payload);
    } else if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      throw new Error('No payload received. Please send payload field in request.');
    }

    const action = params.action;
    const user = params.user || null;
    let result;

    switch (action) {
      case 'ping':
        result = { success: true, message: 'System online', timestamp: new Date().toISOString() };
        break;

      case 'login':
        result = login(params.role, params.code);
        break;

      case 'registerMember':
        result = registerMember(params.data, user);
        break;

      case 'registerMasul':
        result = registerMasul(params.data, user);
        break;

      case 'getMembers':
        result = getMembers(user, params.page, params.pageSize, params.search, params.filters);
        break;

      case 'getGraduates':
        result = getGraduates(user, params.page, params.pageSize, params.search, params.filters);
        break;

      case 'getMasuls':
        result = getMasuls(user, params.page, params.pageSize, params.search, params.filters);
        break;

      case 'getZones':
        result = getZones(user);
        break;

      case 'getBranches':
        result = getBranches(user, params.zone);
        break;

      case 'promoteMember':
        result = promoteMember(params.intizarId, user);
        break;

      case 'promoteToGraduate':
        result = promoteToGraduate(params.intizarId, user);
        break;

      case 'promoteMasul':
        result = promoteMasul(params.intizarId, user);
        break;

      case 'transferMember':
        result = transferMember(params.intizarId, params.newBranchCode, user);
        break;

      case 'transferMasul':
        result = transferMasul(params.intizarId, params.newBranchCode, user);
        break;

      case 'addZone':
        result = addZone(params.zoneName, user);
        break;

      case 'editZone':
        result = editZone(params.zoneId, params.newName, user);
        break;

      case 'disableZone':
        result = disableZone(params.zoneId, user);
        break;

      case 'enableZone':
        result = enableZone(params.zoneId, user);
        break;

      case 'addBranch':
        result = addBranch(params.branchName, params.zoneName, user);
        break;

      case 'editBranch':
        result = editBranch(params.branchCode, params.newName, params.newZone, user);
        break;

      case 'disableBranch':
        result = disableBranch(params.branchCode, user);
        break;

      case 'enableBranch':
        result = enableBranch(params.branchCode, user);
        break;

      case 'getAuditLog':
        result = getAuditLog(user);
        break;

      case 'getConfig':
        result = getConfig(params.key, user);
        break;

      case 'updateConfig':
        result = updateConfig(params.key, params.value, user);
        break;

      case 'exportData':
        result = exportData(params.type, user);
        break;

      case 'getMember':
        result = getMember(params.intizarId, user);
        break;

      case 'getMasul':
        result = getMasul(params.intizarId, user);
        break;

      case 'updateMember':
        result = updateMember(params.intizarId, params.data, user);
        break;

      case 'updateMasul':
        result = updateMasul(params.intizarId, params.data, user);
        break;

      case 'getDashboardStats':
        result = getDashboardStats(user);
        break;

      case 'getZoneStats':
        result = getZoneStats(user);
        break;

      case 'getBranchStats':
        result = getBranchStats(user);
        break;

      case 'getFilterOptions':
        result = {
          success: true,
          branches: getDistinctBranches(),
          zones: getDistinctZones(),
          levels: getDistinctLevels(),
          genders: getDistinctGenders(),
          ranks: getDistinctRanks()
        };
        break;

      case 'getSpreadsheetUrl':
        result = getSpreadsheetUrl(user);
        break;

      default:
        throw new Error('Unknown action: ' + action);
    }

    return createJsonOutput(result);

  } catch (err) {
    logAudit('SYSTEM', 'ERROR', err.toString());
    return createJsonOutput({ success: false, error: err.toString() });
  }
}

function createJsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
