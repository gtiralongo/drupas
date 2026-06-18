// ==================== ADMIN AUTH ====================

function requireAdminCode(code) {
  ensureSheets();
  var stored = getConfig('admin_code');
  return stored && stored === String(code);
}

function getAdminEmails(code) {
  if (!requireAdminCode(code)) return [];
  var raw = getConfig('admin_emails');
  if (!raw) return [];
  return JSON.parse(raw);
}

function addAdminEmail(email, code) {
  if (!requireAdminCode(code)) return false;
  var raw = getConfig('admin_emails');
  var list = raw ? JSON.parse(raw) : [];
  if (list.indexOf(email) === -1) {
    list.push(email);
    setConfig('admin_emails', JSON.stringify(list));
  }
  return true;
}

function removeAdminEmail(email, code) {
  if (!requireAdminCode(code)) return false;
  var raw = getConfig('admin_emails');
  var list = raw ? JSON.parse(raw) : [];
  var idx = list.indexOf(email);
  if (idx !== -1) {
    list.splice(idx, 1);
    setConfig('admin_emails', JSON.stringify(list));
  }
  return true;
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function verifyAdminCode(code) {
  return requireAdminCode(code);
}

function updateAdminCode(oldCode, newCode) {
  if (!requireAdminCode(oldCode)) return false;
  setConfig('admin_code', String(newCode));
  return true;
}



// ==================== STORE CONFIG ====================

function getStoreConfig() {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Config');
  var data = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  return config;
}

function updateStoreConfig(data, code) {
  if (!requireAdminCode(code)) return false;
  for (var key in data) {
    setConfig(key, data[key]);
  }
  return true;
}
