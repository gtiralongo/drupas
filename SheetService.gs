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

function productFromRow_(row, includeAll) {
  var p = {
    id: Number(row[0]),
    nombre: row[1] || '',
    descripcion: row[2] || '',
    precio: Number(row[3]) || 0,
    costo: Number(row[8]) || 0,
    categoria: row[4] || '',
    stock: Number(row[5]) || -1,
    imagenURL: row[6] || ''
  };
  if (includeAll) p.activo = row[7] || 'Si';
  return p;
}

function getProducts() {
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var products = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'Si') products.push(productFromRow_(data[i], false));
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
    products.push(productFromRow_(data[i], true));
  }
  return products;
}

function addProduct(data, email) {
  if (!checkAdminEmail(email)) return false;
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
    'Si',
    Number(data.costo) || 0
  ]);
  return newId;
}

function updateProduct(id, data, email) {
  if (!checkAdminEmail(email)) return false;
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
      if (data.costo !== undefined) sheet.getRange(row, 9).setValue(Number(data.costo));
      return true;
    }
  }
  return false;
}

function toggleProductActive(id, email) {
  if (!checkAdminEmail(email)) return false;
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

function deleteProduct(id, email) {
  if (!checkAdminEmail(email)) return false;
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

function orderFromRow_(row) {
  var itemsRaw = row[6] || '[]';
  if (typeof itemsRaw !== 'string') itemsRaw = '[]';
  return {
    id: String(row[0] || ''),
    fecha: String(row[1] || ''),
    nombre: String(row[2] || ''),
    email: String(row[3] || ''),
    telefono: String(row[4] || ''),
    direccion: String(row[5] || ''),
    productos: itemsRaw,
    total: Number(row[7]) || 0,
    estado: String(row[8] || 'Pendiente'),
    notas: String(row[9] || ''),
    archivado: row[10] === 'Si'
  };
}

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
    data.notas || '',
    'No'
  ]);

  return { orderId: orderId, total: total, fecha: date };
}

function getOrders(email) {
  if (!checkAdminEmail(email)) return [];
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var orders = [];
  for (var i = 0; i < data.length; i++) {
    var o = orderFromRow_(data[i]);
    if (!o.archivado) orders.push(o);
  }
  orders.reverse();
  return orders;
}

function getAllOrders(email) {
  if (!checkAdminEmail(email)) return [];
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var orders = [];
  for (var i = 0; i < data.length; i++) {
    orders.push(orderFromRow_(data[i]));
  }
  orders.reverse();
  return orders;
}

function updateOrderStatus(orderId, estado, email) {
  if (!checkAdminEmail(email)) return false;
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

function archiveOrder(orderId, email) {
  if (!checkAdminEmail(email)) return false;
  ensureSheets();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === orderId) {
      sheet.getRange(i + 1, 11).setValue('Si');
      return true;
    }
  }
  return false;
}

// ==================== COSTOS ====================

function getCostSummary(email) {
  try {
    if (!checkAdminEmail(email)) return [];
    ensureSheets();
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var pSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Productos');
    var pData = pSheet.getDataRange().getValues();
    var costMap = {};
    for (var i = 1; i < pData.length; i++) {
      costMap[pData[i][0]] = Number(pData[i][8]) || 0;
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var summary = {};

    for (var i = 0; i < data.length; i++) {
      var o = orderFromRow_(data[i]);
      if (o.archivado) continue;
      var items;
      try { items = JSON.parse(o.productos || '[]'); } catch(e) { items = []; }
      for (var j = 0; j < items.length; j++) {
        var item = items[j];
        var pid = item.id;
        if (!summary[pid]) {
          summary[pid] = {
            id: pid,
            nombre: item.nombre || 'Producto',
            cantidad: 0,
            ingresoTotal: 0,
            costoTotal: 0
          };
        }
        summary[pid].cantidad += item.cantidad || 1;
        summary[pid].ingresoTotal += (item.precio || 0) * (item.cantidad || 1);
        summary[pid].costoTotal += (costMap[pid] || 0) * (item.cantidad || 1);
      }
    }

    var result = [];
    for (var key in summary) {
      summary[key].margen = summary[key].ingresoTotal - summary[key].costoTotal;
      result.push(summary[key]);
    }
    result.sort(function(a, b) { return b.cantidad - a.cantidad; });
    return result;
  } catch(e) {
    Logger.log('Error en getCostSummary: ' + e.message + ' | ' + e.stack);
    throw e;
  }
}