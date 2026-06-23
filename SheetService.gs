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

function sendOrderEmail(email, orderId) {
  if (!email) return;
  try {
    var config = getStoreConfig();
    var storeName = config.store_name || 'Finca las Drupas';
    var bankName = config.bank_name || 'Banco Nación';
    var bankALIAS = config.bank_ALIAS || '1234567890123456789012';
    var bankHolder = config.bank_holder || 'Finca las Drupas';
    var deliveryInfo = config.delivery_info || 'Te vamos a contactar por WhatsApp para coordinar la entrega.';
    var whatsapp = config.whatsapp || '';

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
    var data = sheet.getDataRange().getValues();
    var productos = [];
    var total = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === orderId) {
        productos = JSON.parse(data[i][6] || '[]');
        total = data[i][7];
        break;
      }
    }

    var itemsHtml = '';
    for (var j = 0; j < productos.length; j++) {
      var item = productos[j];
      itemsHtml += '<tr><td style="padding:8px 12px;border-bottom:1px solid #e0d6c8;font-size:14px">' + item.nombre + ' × ' + item.cantidad + '</td><td style="padding:8px 12px;border-bottom:1px solid #e0d6c8;font-size:14px;text-align:right">$' + Number(item.precio * item.cantidad).toLocaleString('es-AR') + '</td></tr>';
    }

    var whatsappHtml = whatsapp ? '<p style="font-size:14px;color:#666;margin:12px 0 0">Ante cualquier duda, escribinos por <a href="https://wa.me/' + whatsapp + '" style="color:#6B8F5E;font-weight:600;text-decoration:none">WhatsApp</a></p>' : '';

    var htmlBody = '<div style="max-width:560px;margin:0 auto;font-family:Helvetica,Arial,sans-serif;background:#F5EFE2;padding:32px 20px">' +
      '<div style="background:#2D2926;padding:24px;border-radius:12px 12px 0 0;text-align:center">' +
      '<h1 style="color:#A8C69B;font-family:Georgia,serif;font-size:22px;margin:0;letter-spacing:-0.3px">' + escHtml(storeName) + '</h1>' +
      '</div>' +
      '<div style="background:#fff;padding:24px;border-radius:0 0 12px 12px">' +
      '<div style="text-align:center;margin:0 0 20px">' +
      '<div style="width:56px;height:56px;border-radius:50%;background:#4D7C3F;margin:0 auto 12px;display:flex;align-items:center;justify-content:center">' +
      '<svg width="24" height="24" viewBox="0 0 24 24" style="stroke:#fff;stroke-width:3;fill:none"><polyline points="20 6 9 17 4 12"/></svg>' +
      '</div>' +
      '<h2 style="font-size:20px;color:#2D2926;margin:0 0 4px">¡Pedido Confirmado!</h2>' +
      '<p style="font-size:13px;color:#999;margin:0">Tu número de pedido es:</p>' +
      '<p style="font-size:28px;font-weight:700;color:#6B8F5E;margin:4px 0 16px;letter-spacing:2px;font-family:Georgia,serif">' + orderId + '</p>' +
      '</div>' +
      '<div style="background:#FFF8E7;border:1.5px solid #E8C87A;border-radius:10px;padding:16px;margin:0 0 20px;text-align:center">' +
      '<p style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:1px;margin:0 0 8px;font-weight:600">Datos de pago</p>' +
      '<p style="font-size:14px;color:#2D2926;margin:0;font-weight:600">' + escHtml(bankName) + '</p>' +
      '<p style="font-size:13px;color:#666;margin:4px 0">ALIAS: ' + escHtml(bankALIAS) + '</p>' +
      '<p style="font-size:13px;color:#666;margin:0">Titular: ' + escHtml(bankHolder) + '</p>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;margin:0 0 16px">' +
      '<thead><tr><th style="padding:8px 12px;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:0.5px;text-align:left;border-bottom:2px solid #e0d6c8">Producto</th><th style="padding:8px 12px;font-size:11px;color:#999;text-transform:uppercase;letter-spacing:0.5px;text-align:right;border-bottom:2px solid #e0d6c8">Subtotal</th></tr></thead>' +
      '<tbody>' + itemsHtml + '</tbody>' +
      '<tfoot><tr><td style="padding:10px 12px;font-size:16px;font-weight:700;border-top:2px solid #2D2926">Total</td><td style="padding:10px 12px;font-size:16px;font-weight:700;text-align:right;border-top:2px solid #2D2926;color:#6B8F5E">$' + Number(total).toLocaleString('es-AR') + '</td></tr></tfoot>' +
      '</table>' +
      '<div style="background:#F5EFE2;border-radius:8px;padding:14px;text-align:center;font-size:13px;color:#666;line-height:1.5">' + escHtml(deliveryInfo) + '</div>' +
      whatsappHtml +
      '</div>' +
      '<p style="text-align:center;font-size:11px;color:#ccc;margin:16px 0 0">' + escHtml(storeName) + ' — Gracias por tu compra</p>' +
      '</div>';

    GmailApp.sendEmail({
      to: email,
      subject: 'Pedido confirmado - ' + storeName + ' - ' + orderId,
      htmlBody: htmlBody,
      name: storeName
    });
  } catch(e) {
    Logger.log('Error al enviar email confirmación: ' + e.message);
  }
}

function createOrder(data) {
  ensureSheets();

  var enabled = getConfig('orders_enabled');
  if (enabled === 'No') {
    return { error: 'Los pedidos están cerrados temporalmente.' };
  }

  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Pedidos');
  var lastRow = sheet.getLastRow();

  var date = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  var orderId = 'P-' + String(lastRow).padStart(4, '0');

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
    '',
    JSON.stringify(productos),
    total,
    'Pendiente',
    data.notas || '',
    'No'
  ]);

  return { orderId: orderId, total: total, fecha: date };
}

function escHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
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
