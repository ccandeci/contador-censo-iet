// Contador de censo IET: cada censista tiene su propia hoja ("Nombre - Puesto N").
const SHEET_ID = '1X4--GyDXaH0upTz3lT_xbFH3My2RgsuTVUfuESzlTT0';
const ENCABEZADOS = ['Puesto', 'Fecha', 'Hora', 'Tipo de vehículo', 'Sentido', 'Censista', 'Clima'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Contador de censo · IET')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function nombreHoja(p) {
  const limpiar = t => String(t || '').replace(/[\[\]\*\?:\/\\]/g, '').replace(/\s+/g, ' ').trim();
  const nombre = (limpiar(p.censista) || 'Sin nombre') + ' - Puesto ' + (limpiar(p.puesto) || 's/n');
  return nombre.slice(0, 100);
}

function guardar(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const nombre = nombreHoja(p);
    const sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const filas = (p.recs || []).map(r => [p.puesto, p.fecha, r[0], r[1], r[2], p.censista, p.clima]);
    if (filas.length) {
      const ini = sh.getLastRow() + 1;
      const rango = sh.getRange(ini, 1, filas.length, ENCABEZADOS.length);
      rango.setNumberFormat('@'); // texto: evita que Sheets cambie fechas y horas
      rango.setValues(filas);
    }
    return filas.length;
  } finally {
    lock.releaseLock();
  }
}
