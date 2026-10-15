/**
 * COPIA DE "Planilla pagos 2.0" A LA HOJA "base"
 *
 * importRange()        → la corre el trigger cada 10 minutos. Si hay una generación corriendo,
 *                        saltea esa vuelta y copia en la próxima.
 * conBaseFresca_(fn)   → la usan los 4 ítems de menú de Generar/Actualizar: toma el candado,
 *                        refresca base y recién ahí genera. Mientras tanto el trigger no puede
 *                        pisar base.
 *
 * Escribe primero y limpia después lo que sobra: base nunca queda vacía a mitad de camino.
 * Si el origen viene vacío, no toca base y lanza error.
 *
 * Este archivo REEMPLAZA al anterior (el que tenía SOURCE_LAST_COLUMN = 41 y la función onEdit).
 * onEdit se eliminó: como trigger simple no puede abrir otra planilla y fallaba en cada edición.
 */

const SOURCE_SS_ID = '18lSReMfnMoFHTY3TYxrExeeHlbaLdajkxNzR5x4p-Lk';
const SOURCE_SHEET_NAME = 'Planilla pagos 2.0';
const DEST_SHEET_NAME = 'base';

function importRange() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('importRange: hay otro proceso corriendo, se saltea esta vuelta.');
    return;
  }
  try {
    importarBase_();
  } finally {
    lock.releaseLock();
  }
}

/** Copia sin tomar el candado: solo la llaman importRange y conBaseFresca_, que ya lo tienen. */
function importarBase_() {
  const origen = SpreadsheetApp.openById(SOURCE_SS_ID).getSheetByName(SOURCE_SHEET_NAME);
  if (!origen) throw new Error('No encuentro la hoja "' + SOURCE_SHEET_NAME + '" en el origen.');

  const filas = origen.getLastRow();
  const cols = origen.getLastColumn();
  if (filas < 2 || cols < 1) {
    throw new Error('El origen vino vacío (' + filas + ' filas). No se tocó base.');
  }
  const valores = origen.getRange(1, 1, filas, cols).getValues();

  const destino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEST_SHEET_NAME);
  if (!destino) throw new Error('No encuentro la hoja "' + DEST_SHEET_NAME + '" en la Taskboard.');
  if (destino.getMaxRows() < filas) {
    destino.insertRowsAfter(destino.getMaxRows(), filas - destino.getMaxRows());
  }
  if (destino.getMaxColumns() < cols) {
    destino.insertColumnsAfter(destino.getMaxColumns(), cols - destino.getMaxColumns());
  }

  const ultFila = destino.getLastRow();
  const ultCol = destino.getLastColumn();

  destino.getRange(1, 1, filas, cols).setValues(valores);

  // Recién ahora se borra lo que sobró de la copia anterior.
  if (ultFila > filas) {
    destino.getRange(filas + 1, 1, ultFila - filas, Math.max(ultCol, cols)).clearContent();
  }
  if (ultCol > cols) {
    destino.getRange(1, cols + 1, filas, ultCol - cols).clearContent();
  }
  SpreadsheetApp.flush();
}

/** Toma el candado, refresca base y corre la función del menú. */
function conBaseFresca_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    SpreadsheetApp.getUi().alert('Hay otro proceso del script corriendo. Probá de nuevo en un minuto.');
    return;
  }
  try {
    importarBase_();
    fn();
  } finally {
    lock.releaseLock();
  }
}

function menuGenerarHojaDia()     { conBaseFresca_(generarHojaPagosDelDia); }
function menuActualizarHojaDia()  { conBaseFresca_(actualizarHojaPagosDelDia); }
// conMontoAPagar_ (CargaGalicia.gs) repone la columna MONTO A PAGAR después de regenerar.
function menuGenerarAgrupada()    { conBaseFresca_(() => conMontoAPagar_(generarHojaAgrupadaDelDia)); }
function menuActualizarAgrupada() { conBaseFresca_(() => conMontoAPagar_(actualizarHojaAgrupadaDelDia)); }
