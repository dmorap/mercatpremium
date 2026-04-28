// =============================================================
// MercatPremium - Gestión de Artículos de Segunda Mano
// =============================================================

var CONFIG = {
  SPREADSHEET_NAME: 'MercatPremium - Inventario',
  ROOT_FOLDER_NAME:  'MercatPremium',
  ARTICLES_FOLDER:   'Articulos',
  SHEET_NAME:        'Inventario'
};

// ─── PUNTO DE ENTRADA DE LA WEB APP ───────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('MercatPremium')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── CONFIGURACIÓN INICIAL ────────────────────────────────────
// Busca o crea la carpeta raíz, el spreadsheet y la carpeta de artículos.
// Devuelve sus IDs para que el resto de funciones los usen.
function getOrCreateSetup() {
  // Carpeta raíz
  var rootFolder;
  var rootFolders = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  if (rootFolders.hasNext()) {
    rootFolder = rootFolders.next();
  } else {
    rootFolder = DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  }

  // Hoja de cálculo
  var spreadsheet;
  var files = rootFolder.getFilesByName(CONFIG.SPREADSHEET_NAME);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    spreadsheet = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
    DriveApp.getFileById(spreadsheet.getId()).moveTo(rootFolder);
    initializeSheet(spreadsheet);
  }

  // Carpeta de artículos
  var articlesFolder;
  var articlesFolders = rootFolder.getFoldersByName(CONFIG.ARTICLES_FOLDER);
  if (articlesFolders.hasNext()) {
    articlesFolder = articlesFolders.next();
  } else {
    articlesFolder = rootFolder.createFolder(CONFIG.ARTICLES_FOLDER);
  }

  return {
    spreadsheetId:  spreadsheet.getId(),
    rootFolderId:   rootFolder.getId(),
    articlesFolderId: articlesFolder.getId(),
    rootFolderUrl:  rootFolder.getUrl(),
    spreadsheetUrl: spreadsheet.getUrl()
  };
}

// Crea las cabeceras y el formato inicial del sheet
function initializeSheet(spreadsheet) {
  var sheet = spreadsheet.getActiveSheet();
  sheet.setName(CONFIG.SHEET_NAME);

  var headers = [
    'ID', 'Nombre', 'Descripcion', 'Categoria', 'Estado_Articulo',
    'Precio', 'Fecha_Entrada', 'Carpeta_ID', 'Foto_Principal_ID',
    'Estado_Venta', 'Plataforma', 'Fecha_Venta', 'Notas'
  ];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Ajustar anchos de columna
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
}

// ─── LEER INVENTARIO ─────────────────────────────────────────
function getInventory() {
  try {
    var setup = getOrCreateSetup();
    var spreadsheet = SpreadsheetApp.openById(setup.spreadsheetId);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { items: [], setup: setup };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    var tz   = Session.getScriptTimeZone();

    var items = data
      .filter(function(row) { return row[0] !== ''; })
      .map(function(row) {
        return {
          id:             row[0],
          nombre:         row[1],
          descripcion:    row[2],
          categoria:      row[3],
          estadoArticulo: row[4],
          precio:         row[5],
          fechaEntrada:   row[6] ? Utilities.formatDate(new Date(row[6]), tz, 'dd/MM/yyyy') : '',
          carpetaId:      row[7],
          fotoPrincipalId: row[8],
          estadoVenta:    row[9] || 'Disponible',
          plataforma:     row[10],
          fechaVenta:     row[11] ? Utilities.formatDate(new Date(row[11]), tz, 'dd/MM/yyyy') : '',
          notas:          row[12]
        };
      });

    return { items: items, setup: setup };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── AÑADIR ARTÍCULO ──────────────────────────────────────────
function addItem(data) {
  try {
    var setup = getOrCreateSetup();
    var spreadsheet = SpreadsheetApp.openById(setup.spreadsheetId);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

    var id = 'ART-' + new Date().getTime();

    // Crear carpeta para este artículo
    var articlesFolder = DriveApp.getFolderById(setup.articlesFolderId);
    var itemFolder = articlesFolder.createFolder(id + ' - ' + data.nombre);

    var newRow = [
      id,
      data.nombre,
      data.descripcion,
      data.categoria,
      data.estadoArticulo,
      parseFloat(data.precio) || 0,
      new Date(),
      itemFolder.getId(),
      '',
      'Disponible',
      '',
      '',
      data.notas || ''
    ];

    sheet.appendRow(newRow);

    return {
      success:   true,
      id:        id,
      folderId:  itemFolder.getId(),
      folderUrl: itemFolder.getUrl()
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── ACTUALIZAR ARTÍCULO ──────────────────────────────────────
function updateItem(data) {
  try {
    var setup = getOrCreateSetup();
    var spreadsheet = SpreadsheetApp.openById(setup.spreadsheetId);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return r[0]; });
    var rowIndex = ids.indexOf(data.id) + 2;
    if (rowIndex < 2) return { error: 'Artículo no encontrado' };

    sheet.getRange(rowIndex, 2).setValue(data.nombre);
    sheet.getRange(rowIndex, 3).setValue(data.descripcion);
    sheet.getRange(rowIndex, 4).setValue(data.categoria);
    sheet.getRange(rowIndex, 5).setValue(data.estadoArticulo);
    sheet.getRange(rowIndex, 6).setValue(parseFloat(data.precio) || 0);
    sheet.getRange(rowIndex, 10).setValue(data.estadoVenta);
    sheet.getRange(rowIndex, 11).setValue(data.plataforma || '');
    sheet.getRange(rowIndex, 13).setValue(data.notas || '');

    if (data.estadoVenta === 'Vendido') {
      sheet.getRange(rowIndex, 12).setValue(new Date());
    }

    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── ELIMINAR ARTÍCULO ────────────────────────────────────────
function deleteItem(itemId) {
  try {
    var setup = getOrCreateSetup();
    var spreadsheet = SpreadsheetApp.openById(setup.spreadsheetId);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return r[0]; });
    var rowIndex = ids.indexOf(itemId) + 2;
    if (rowIndex < 2) return { error: 'Artículo no encontrado' };

    // Mover carpeta a la papelera
    var folderId = sheet.getRange(rowIndex, 8).getValue();
    if (folderId) {
      try { DriveApp.getFolderById(folderId).setTrashed(true); } catch (e2) {}
    }

    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── FOTOS ────────────────────────────────────────────────────
function getItemPhotos(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files  = folder.getFiles();
    var photos = [];

    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType().indexOf('image/') === 0) {
        photos.push({
          id:      file.getId(),
          name:    file.getName(),
          thumb:   'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400',
          fullUrl: file.getUrl()
        });
      }
    }

    return { photos: photos };
  } catch (e) {
    return { error: e.toString() };
  }
}

// Establece la foto principal de un artículo
function setMainPhoto(itemId, photoId) {
  try {
    var setup = getOrCreateSetup();
    var spreadsheet = SpreadsheetApp.openById(setup.spreadsheetId);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return r[0]; });
    var rowIndex = ids.indexOf(itemId) + 2;
    if (rowIndex < 2) return { error: 'Artículo no encontrado' };

    sheet.getRange(rowIndex, 9).setValue(photoId);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── GENERAR ANUNCIO ──────────────────────────────────────────
function generateListing(itemId, platform) {
  try {
    var result = getInventory();
    var item   = result.items.filter(function(i) { return i.id === itemId; })[0];
    if (!item) return { error: 'Artículo no encontrado' };

    var text = '';
    var cat  = item.categoria.toLowerCase().replace(/\s+/g, '');

    if (platform === 'wallapop') {
      text  = '📦 ' + item.nombre + '\n\n';
      text += item.descripcion + '\n\n';
      text += '✅ Estado: ' + item.estadoArticulo + '\n';
      text += '💶 Precio: ' + item.precio + '€\n\n';
      text += '#segundamano #' + cat + ' #wallapop';

    } else if (platform === 'vinted') {
      text  = item.nombre + '\n\n';
      text += item.descripcion + '\n\n';
      text += 'Estado: ' + item.estadoArticulo + '\n';
      text += 'Precio: ' + item.precio + '€\n\n';
      text += '📍 Consultas por chat de Vinted';

    } else if (platform === 'facebook') {
      text  = '🛒 SE VENDE: ' + item.nombre + '\n\n';
      text += '📝 ' + item.descripcion + '\n\n';
      text += '✅ Estado: ' + item.estadoArticulo + '\n';
      text += '💶 Precio: ' + item.precio + '€\n\n';
      text += '📩 Contactar por privado\n';
      text += '#segundamano #vendo #' + cat;
    }

    return { success: true, text: text, item: item };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── ESTADÍSTICAS ─────────────────────────────────────────────
function getStats() {
  try {
    var result = getInventory();
    if (result.error) return result;

    var items      = result.items;
    var disponibles = items.filter(function(i) { return i.estadoVenta === 'Disponible'; });
    var vendidos    = items.filter(function(i) { return i.estadoVenta === 'Vendido'; });
    var reservados  = items.filter(function(i) { return i.estadoVenta === 'Reservado'; });

    var ingresoTotal = vendidos.reduce(function(s, i) { return s + (parseFloat(i.precio) || 0); }, 0);
    var valorStock   = disponibles.reduce(function(s, i) { return s + (parseFloat(i.precio) || 0); }, 0);

    return {
      total:          items.length,
      disponibles:    disponibles.length,
      vendidos:       vendidos.length,
      reservados:     reservados.length,
      ingresoTotal:   ingresoTotal.toFixed(2),
      valorStock:     valorStock.toFixed(2)
    };
  } catch (e) {
    return { error: e.toString() };
  }
}
