const SHEET_ID = '158EgsOYAIZG9qJ2WZuFGRNJipv7_06o2IEb1Rdq66zY';

function doGet(e) {
  var page = e.parameter.page || 'store';

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Panel Admin - Finca las Drupas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Pedidos - Finca las Drupas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  ensureSheet_('Productos', ['ID', 'Nombre', 'Descripcion', 'Precio', 'Categoria', 'Stock', 'ImagenURL', 'Activo', 'Costo']);
  ensureSheet_('Pedidos', ['ID', 'Fecha', 'ClienteNombre', 'ClienteEmail', 'ClienteTel', 'Direccion', 'Productos', 'Total', 'Estado', 'Notas', 'Archivado']);
  ensureSheet_('Config', ['Clave', 'Valor']);

  var config = ss.getSheetByName('Config');
  if (config.getLastRow() <= 1) {
    config.appendRow(['store_name', 'Finca las Drupas']);
    config.appendRow(['store_tagline', 'Productos regionales']);
    config.appendRow(['admin_emails', '[]']);
    config.appendRow(['admin_code', 'admin123']);
    config.appendRow(['primary_color', '#6B8F5E']);
    config.appendRow(['secondary_color', '#D4A574']);
    config.appendRow(['whatsapp', '']);
    config.appendRow(['delivery_info', 'Consultá por delivery a tu zona']);
    config.appendRow(['bank_name', 'Banco Nación']);
    config.appendRow(['bank_cbu', '1234567890123456789012']);
    config.appendRow(['bank_holder', 'Finca las Drupas']);
  }

  ensureConfigKey_(config, 'bank_name', 'Banco Nación');
  ensureConfigKey_(config, 'bank_cbu', '1234567890123456789012');
  ensureConfigKey_(config, 'bank_holder', 'Finca las Drupas');
}

function ensureConfigKey_(config, key, defaultValue) {
  var data = config.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return;
  }
  config.appendRow([key, defaultValue]);
}

function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    if (name === 'Productos') {
      sheet.appendRow([1, 'Aceite de Oliva Extra Virgen 500ml', 'Aceite de oliva extra virgen, cosecha temprana. Producido en nuestra finca.', 12000, 'Aceite', -1, '', 'Si', 5000]);
      sheet.appendRow([2, 'Aceitunas Verdes x500g', 'Aceitunas verdes en salmuera, seleccionadas a mano.', 5000, 'Aceitunas', -1, '', 'Si', 2000]);
      sheet.appendRow([3, 'Miel Pura de Campo x500g', 'Miel cruda sin procesar, cosecha directa de nuestras colmenas.', 7000, 'Miel', -1, '', 'Si', 3000]);
      sheet.appendRow([4, 'Maní tostado x250g', 'Maní tostado artesanal, salado justo a punto.', 3500, 'Maní', -1, '', 'Si', 1500]);
    }
    return;
  }
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headerRow.indexOf(headers[i]) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(headers[i]);
    }
  }
}

function requestMailPermission() {
  MailApp.getRemainingDailyQuota();
  return true;
}