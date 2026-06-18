const SHEET_ID = '158EgsOYAIZG9qJ2WZuFGRNJipv7_06o2IEb1Rdq66zY';

function doGet(e) {
  var page = e.parameter.page || 'store';

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Panel Admin - Vértice Gin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Pedidos - Vértice Gin')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  ensureSheet_(ss, 'Productos', ['ID', 'Nombre', 'Descripcion', 'Precio', 'Categoria', 'Stock', 'ImagenURL', 'Activo']);
  ensureSheet_(ss, 'Pedidos', ['ID', 'Fecha', 'ClienteNombre', 'ClienteEmail', 'ClienteTel', 'Direccion', 'Productos', 'Total', 'Estado', 'Notas']);
  ensureSheet_(ss, 'Config', ['Clave', 'Valor']);

  var config = ss.getSheetByName('Config');
  if (config.getLastRow() <= 1) {
    config.appendRow(['store_name', 'Vértice Gin de Pueblo']);
    config.appendRow(['store_tagline', 'Destilamos pueblo']);
    config.appendRow(['admin_emails', '[]']);
    config.appendRow(['admin_code', 'admin123']);
    config.appendRow(['primary_color', '#8B6914']);
    config.appendRow(['secondary_color', '#C9A84C']);
    config.appendRow(['whatsapp', '']);
    config.appendRow(['delivery_info', 'Consultá por delivery a tu zona']);
  }
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    if (name === 'Productos') {
      sheet.appendRow([1, 'Vértice Gin 750ml', 'Gin artesanal destilado en alambique de cobre. 42% ABV. Botánicos cosechados en la plaza de Pasteur.', 18000, 'Gin', -1, '', 'Si']);
      sheet.appendRow([2, 'Vértice Gin 375ml', 'Media botella. Ideal para regalar o probar nuestro gin.', 10000, 'Gin', -1, '', 'Si']);
      sheet.appendRow([3, 'Vértice Gin Mini 200ml', 'Formato degustación. Llevate la experiencia Vértice a donde vayas.', 5500, 'Gin', -1, '', 'Si']);
    }
  }
}
