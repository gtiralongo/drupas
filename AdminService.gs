// ==================== ADMIN AUTH ====================

function parseAdminEmails_(raw) {
  try {
    var parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed;
  } catch(e) {}
  return [];
}

function hashPassword_(password) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return digest.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function checkAdminEmail(email) {
  ensureSheets();
  if (!email) return false;
  var raw = getConfig('admin_emails');
  if (!raw) return false;
  var list = parseAdminEmails_(raw);
  if (list.length === 0) return true;
  var normalized = email.toLowerCase().trim();
  for (var i = 0; i < list.length; i++) {
    if (list[i].toLowerCase().trim() === normalized) return true;
  }
  return false;
}

function checkAdminLogin(email, password) {
  if (!checkAdminEmail(email)) return false;
  var storedHash = getConfig('admin_password');
  if (!storedHash) return true;
  if (!password) return false;
  return hashPassword_(password) === storedHash;
}

function setAdminPassword(password, email) {
  if (!checkAdminEmail(email)) return { success: false, message: 'No autorizado' };
  if (!password || password.length < 4) return { success: false, message: 'Mínimo 4 caracteres' };
  setConfig('admin_password', hashPassword_(password));
  return { success: true, message: 'Contraseña guardada' };
}

function getAdminEmails(email) {
  if (!checkAdminEmail(email)) return [];
  var raw = getConfig('admin_emails');
  if (!raw) return [];
  return parseAdminEmails_(raw);
}

function addAdminEmail(newEmail, email) {
  if (!checkAdminEmail(email)) return { success: false, message: 'No autorizado' };
  var raw = getConfig('admin_emails');
  var list = raw ? parseAdminEmails_(raw) : [];
  var normalized = newEmail.toLowerCase().trim();
  if (list.indexOf(normalized) === -1) {
    list.push(normalized);
    setConfig('admin_emails', JSON.stringify(list));
    return { success: true, message: 'Admin agregado' };
  }
  return { success: false, message: 'El email ya es admin' };
}

function removeAdminEmail(targetEmail, email) {
  if (!checkAdminEmail(email)) return { success: false, message: 'No autorizado' };
  var raw = getConfig('admin_emails');
  var list = raw ? parseAdminEmails_(raw) : [];
  var normalized = targetEmail.toLowerCase().trim();
  var idx = list.indexOf(normalized);
  if (idx !== -1) {
    list.splice(idx, 1);
    setConfig('admin_emails', JSON.stringify(list));
    return { success: true, message: 'Admin eliminado' };
  }
  return { success: false, message: 'Email no encontrado' };
}

function hasAdminPassword() {
  var storedHash = getConfig('admin_password');
  return !!storedHash;
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

function updateStoreConfig(data, email) {
  if (!checkAdminEmail(email)) return false;
  for (var key in data) {
    setConfig(key, data[key]);
  }
  return true;
}