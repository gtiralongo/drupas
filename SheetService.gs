// ==================== CONFIG ====================

function getConfig(key) {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Config');
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function setConfig(key, value) {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Config');
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ==================== PRODUCTOS ====================

function getProducts() {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var products = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'Si') {
      products.push({
        id: Number(data[i][0]),
        nombre: data[i][1] || '',
        descripcion: data[i][2] || '',
        precio: Number(data[i][3]) || 0,
        categoria: data[i][4] || '',
        stock: Number(data[i][5]) || -1,
        imagenURL: data[i][6] || ''
      });
    }
  }
  return products;
}

function getAllProducts() {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var products = [];
  for (var i = 1; i < data.length; i++) {
    products.push({
      id: Number(data[i][0]),
      nombre: data[i][1] || '',
      descripcion: data[i][2] || '',
      precio: Number(data[i][3]) || 0,
      categoria: data[i][4] || '',
      stock: Number(data[i][5]) || -1,
      imagenURL: data[i][6] || '',
      activo: data[i][7] || 'Si'
    });
  }
  return products;
}

function addProduct(data, code) {
  if (!requireAdminCode(code)) return false;
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var lastRow = sheet.getLastRow();
  var newId = lastRow;
  sheet.appendRow([
    newId,
    data.nombre,
    data.descripcion,
    Number(data.precio),
    data.categoria,
    Number(data.stock) || -1,
    data.imagenURL || '',
    'Si'
  ]);
  return newId;
}

function updateProduct(id, data, code) {
  if (!requireAdminCode(code)) return false;
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(id)) {
      var row = i + 1;
      if (data.nombre !== undefined) sheet.getRange(row, 2).setValue(data.nombre);
      if (data.descripcion !== undefined) sheet.getRange(row, 3).setValue(data.descripcion);
      if (data.precio !== undefined) sheet.getRange(row, 4).setValue(Number(data.precio));
      if (data.categoria !== undefined) sheet.getRange(row, 5).setValue(data.categoria);
      if (data.stock !== undefined) sheet.getRange(row, 6).setValue(Number(data.stock));
      if (data.imagenURL !== undefined) sheet.getRange(row, 7).setValue(data.imagenURL);
      if (data.activo !== undefined) sheet.getRange(row, 8).setValue(data.activo);
      return true;
    }
  }
  return false;
}

function toggleProductActive(id, code) {
  if (!requireAdminCode(code)) return false;
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(id)) {
      var current = values[i][7] === 'Si' ? 'No' : 'Si';
      sheet.getRange(i + 1, 8).setValue(current);
      return current;
    }
  }
  return null;
}

function deleteProduct(id, code) {
  if (!requireAdminCode(code)) return false;
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// ==================== PEDIDOS ====================

function createOrder(data) {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
  var lastRow = sheet.getLastRow();

  var date = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  var orderId = 'PED-' + String(lastRow).padStart(4, '0');

  var total = 0;
  var productos = [];
  if (data.productos && Array.isArray(data.productos)) {
    for (var i = 0; i < data.productos.length; i++) {
      var p = data.productos[i];
      total += p.precio * p.cantidad;
      productos.push({ id: p.id, nombre: p.nombre, cantidad: p.cantidad, precio: p.precio });
    }
  }

  sheet.appendRow([
    orderId,
    date,
    data.nombre || '',
    data.email || '',
    data.telefono || '',
    data.direccion || '',
    JSON.stringify(productos),
    total,
    'Pendiente',
    data.notas || ''
  ]);

  return { orderId: orderId, total: total, fecha: date };
}

function getOrders(code) {
  if (!requireAdminCode(code)) return [];
  ensureSheets();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Pedidos');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  var orders = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var itemsRaw = row[6] || '[]';
    if (typeof itemsRaw !== 'string') itemsRaw = '[]';
    orders.push({
      id: String(row[0] || ''),
      fecha: String(row[1] || ''),
      nombre: String(row[2] || ''),
      email: String(row[3] || ''),
      telefono: String(row[4] || ''),
      direccion: String(row[5] || ''),
      productos: itemsRaw,
      total: Number(row[7]) || 0,
      estado: String(row[8] || 'Pendiente'),
      notas: String(row[9] || '')
    });
  }
  orders.reverse();
  return orders;
}

function updateOrderStatus(orderId, estado, code) {
  if (!requireAdminCode(code)) return false;
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === orderId) {
      sheet.getRange(i + 1, 9).setValue(estado);
      return true;
    }
  }
  return false;
}
