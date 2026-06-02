/**
 * VERUM CAPITAL — Google Apps Script
 * ─────────────────────────────────────────────────────────────────
 * Cómo desplegar (una sola vez):
 *
 * 1. Abre script.google.com → "Nuevo proyecto"
 * 2. Borra el contenido por defecto y pega este código completo
 * 3. Guardar (Ctrl+S)
 * 4. Clic en "Implementar" → "Nueva implementación"
 * 5. Tipo: "Aplicación web"
 *    - Descripción: Verum Capital Análisis
 *    - Ejecutar como: Yo (tu cuenta de Google)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Clic en "Implementar" → Autoriza los permisos que pida
 * 7. Copia la URL que aparece (termina en /exec)
 * 8. Pégala en el campo "URL del Google Apps Script" en la landing
 *
 * Si modificas el script, ve a "Implementar" → "Gestionar implementaciones"
 * → edita la implementación existente → "Nueva versión" → Implementar.
 * ─────────────────────────────────────────────────────────────────
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'init') {
      return initFolder(data);
    }

    if (data.action === 'upload_file') {
      return uploadFile(data);
    }

    return jsonResponse({ error: 'Acción no reconocida: ' + data.action });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── CREAR CARPETA + GOOGLE DOC DE NOTAS ──────────────────────────

function initFolder(data) {
  // Crear carpeta en la raíz de Mi Drive
  var folder = DriveApp.getRootFolder().createFolder(data.folderName);

  // Crear Google Doc con notas internas
  var doc = DocumentApp.create(data.docTitle || ('Notas — ' + data.folderName));

  // Mover el doc a la carpeta (los nuevos docs se crean en la raíz)
  var docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);

  // Contenido del documento
  var body = doc.getBody();

  // Cabecera
  body.appendParagraph('VERUM CAPITAL — NOTAS INTERNAS')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph(data.folderName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Ficha de datos
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

  // Notas internas
  body.appendParagraph('');
  body.appendParagraph('NOTAS INTERNAS')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);

  if (data.internalNotes && data.internalNotes.trim()) {
    body.appendParagraph(data.internalNotes);
  } else {
    body.appendParagraph('(Sin notas internas)')
      .setItalic(true);
  }

  // Pie
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

// ─── SUBIR ARCHIVO A LA CARPETA ───────────────────────────────────

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

// ─── HELPER ───────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
