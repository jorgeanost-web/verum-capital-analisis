/**
 * VERUM CAPITAL — Google Apps Script v2
 * ─────────────────────────────────────────────────────────────────
 * Endpoints disponibles (action):
 *   init            → Crea carpeta + Google Doc de notas
 *   upload_file     → Sube archivo a la carpeta
 *   save_analysis   → Guarda JSON del análisis + actualiza índice
 *   list_analyses   → Lista todas las oportunidades del índice
 *   get_analysis    → Lee el JSON de un análisis concreto
 *   update_status   → Cambia el estado de una oportunidad
 *
 * Para redesplegar tras modificar:
 *   Implementar → Gestionar implementaciones → editar → Nueva versión
 * ─────────────────────────────────────────────────────────────────
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'init')            return initFolder(data);
    if (data.action === 'upload_file')     return uploadFile(data);
    if (data.action === 'save_analysis')   return saveAnalysis(data);
    if (data.action === 'list_analyses')   return listAnalyses();
    if (data.action === 'get_analysis')    return getAnalysis(data);
    if (data.action === 'update_status')   return updateStatus(data);

    return jsonResponse({ error: 'Acción no reconocida: ' + data.action });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── ÍNDICE CENTRAL (Google Sheet) ────────────────────────────────

function getOrCreateIndexSheet() {
  var PROP_KEY = 'vc_index_sheet_id';
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty(PROP_KEY);

  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId).getActiveSheet();
    } catch (e) {
      // La hoja fue eliminada — la recreamos
    }
  }

  var ss = SpreadsheetApp.create('Verum Capital — Índice de Oportunidades');
  var sheet = ss.getActiveSheet();
  sheet.setName('Oportunidades');
  sheet.appendRow([
    'id', 'nombre', 'direccion', 'tipologia', 'estrategia',
    'superficie', 'precio', 'fecha', 'analista', 'estado',
    'folderId', 'folderUrl', 'docUrl', 'createdAt', 'updatedAt'
  ]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 15)
    .setBackground('#3A3A36')
    .setFontColor('#C9A96E')
    .setFontWeight('bold');

  props.setProperty(PROP_KEY, ss.getId());
  return sheet;
}

// ─── CREAR CARPETA + GOOGLE DOC DE NOTAS ──────────────────────────

function initFolder(data) {
  var folder = DriveApp.getRootFolder().createFolder(data.folderName);

  var doc = DocumentApp.create(data.docTitle || ('Notas — ' + data.folderName));
  var docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);

  var body = doc.getBody();

  body.appendParagraph('VERUM CAPITAL — NOTAS INTERNAS')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(data.folderName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('DATOS DEL ACTIVO')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);

  var tabla = [
    ['Tipo de activo', data.assetType || '—'],
    ['Estrategia', data.strategy || '—'],
    ['Dirección / Zona', data.address || '—'],
    ['Superficie', data.surface ? data.surface + ' m²' : '—'],
    ['Fecha de análisis', data.date || '—'],
    ['Analista', data.analyst || '—'],
    ['Contacto', data.contactName || '—'],
    ['Tipo de contacto', data.contactType || '—'],
  ];

  tabla.forEach(function(row) {
    body.appendParagraph(row[0] + ': ' + row[1]);
  });

  body.appendParagraph('');
  body.appendParagraph('NOTAS INTERNAS')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);

  if (data.internalNotes && data.internalNotes.trim()) {
    body.appendParagraph(data.internalNotes);
  } else {
    body.appendParagraph('(Sin notas internas)').setItalic(true);
  }

  body.appendParagraph('');
  body.appendParagraph('Documento generado automáticamente por el sistema de análisis de Verum Capital.')
    .setItalic(true);

  doc.saveAndClose();

  return jsonResponse({
    folderId: folder.getId(),
    folderUrl: folder.getUrl(),
    docUrl: doc.getUrl(),
  });
}

// ─── SUBIR ARCHIVO ────────────────────────────────────────────────

function uploadFile(data) {
  var folder = DriveApp.getFolderById(data.folderId);
  var decoded = Utilities.base64Decode(data.base64Data);
  var blob = Utilities.newBlob(decoded, data.mimeType || 'application/octet-stream', data.fileName);
  var file = folder.createFile(blob);

  return jsonResponse({
    ok: true,
    fileId: file.getId(),
    fileName: file.getName(),
  });
}

// ─── GUARDAR ANÁLISIS COMPLETO ────────────────────────────────────

function saveAnalysis(data) {
  // 1. Guardar JSON del estado completo en la carpeta de Drive
  try {
    var folder = DriveApp.getFolderById(data.folderId);
    var existing = folder.getFilesByName('analisis.json');
    while (existing.hasNext()) { existing.next().setTrashed(true); }

    var jsonBlob = Utilities.newBlob(
      JSON.stringify(data.analysisState),
      'application/json',
      'analisis.json'
    );
    folder.createFile(jsonBlob);
  } catch (e) {
    return jsonResponse({ error: 'Error guardando JSON: ' + e.message });
  }

  // 2. Actualizar índice en la hoja de cálculo
  try {
    var sheet = getOrCreateIndexSheet();
    var rows = sheet.getDataRange().getValues();
    var id = data.id;
    var existingRow = -1;
    var createdAt = new Date().toISOString();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        existingRow = i + 1;
        createdAt = rows[i][13] || createdAt;
        break;
      }
    }

    var now = new Date().toISOString();
    var rowData = [
      id, data.nombre || '', data.direccion || '', data.tipologia || '',
      data.estrategia || '', data.superficie || '', data.precio || '',
      data.fecha || '', data.analista || '', data.estado || 'En análisis',
      data.folderId || '', data.folderUrl || '', data.docUrl || '',
      createdAt, now
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  } catch (e) {
    return jsonResponse({ error: 'Error actualizando índice: ' + e.message });
  }

  return jsonResponse({ ok: true });
}

// ─── LISTAR ANÁLISIS ──────────────────────────────────────────────

function listAnalyses() {
  try {
    var sheet = getOrCreateIndexSheet();
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return jsonResponse({ analyses: [] });

    var headers = values[0];
    var analyses = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (!row[0]) continue;
      var obj = {};
      headers.forEach(function(h, j) { obj[h] = row[j]; });
      analyses.push(obj);
    }
    analyses.reverse();
    return jsonResponse({ analyses: analyses });
  } catch (e) {
    return jsonResponse({ analyses: [], error: e.message });
  }
}

// ─── OBTENER ANÁLISIS CONCRETO ────────────────────────────────────

function getAnalysis(data) {
  try {
    var folder = DriveApp.getFolderById(data.folderId);
    var files = folder.getFilesByName('analisis.json');
    if (!files.hasNext()) {
      return jsonResponse({ error: 'No se encontró analisis.json en esta carpeta.' });
    }
    var content = files.next().getBlob().getDataAsString();
    return jsonResponse({ analysisState: JSON.parse(content) });
  } catch (e) {
    return jsonResponse({ error: e.message });
  }
}

// ─── ACTUALIZAR ESTADO ────────────────────────────────────────────

function updateStatus(data) {
  try {
    var sheet = getOrCreateIndexSheet();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === data.id) {
        sheet.getRange(i + 1, 10).setValue(data.estado);
        sheet.getRange(i + 1, 15).setValue(new Date().toISOString());
        return jsonResponse({ ok: true });
      }
    }
    return jsonResponse({ error: 'Análisis no encontrado.' });
  } catch (e) {
    return jsonResponse({ error: e.message });
  }
}

// ─── HELPER ───────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
