/**
 * DESCARGA DE FACTURAS A Y C QUE TODAVÍA NO SE CARGARON EN EL BANCO
 *
 * descargarFacturasPendientes() → ítem de menú. Junta en un ZIP las facturas de los grupos A y C
 * de la agrupada de HOY que tienen el ESTADO vacío, lo guarda en PDF_CARPETA_ID y abre una ventana
 * para bajarlo.
 *
 * QUÉ ENTRA
 *   - Tipo A o C, con IDs (no bandas ni resumen) y ESTADO vacío.
 *   - NO entran las filas "↻ EN BANCO": nacen con ESTADO vacío pero ya estaban cargadas en el
 *     banco en la agrupada anterior (agrupAnterior_). Es el mismo criterio que CargaGalicia.gs.
 *   - Una factura por link de la columna LINKS. Un link repetido en dos grupos baja una sola vez.
 *
 * NOMBRE DE CADA ARCHIVO EN EL ZIP: "<TAREA> - <BENEFICIARIO>.<ext>", para reconocerlo sin abrirlo.
 * La tarea sale de la hoja del día (INVOICE/COMPROBANTE → TASK). Si el link no aparece ahí, se usan
 * los IDs del grupo. Si dos archivos quedan con el mismo nombre, se numeran.
 *
 * Los links que no son de Drive o que no se pueden abrir no cortan la descarga: se listan al final.
 *
 * DEPENDE de Pagoslb.gs (formatearNombreHoja_, indicesPorHeader_, extraerIdDeDrive_,
 * TIPOS_VALIDABLES), Hojaagrupada.gs (AGRUP_PREFIJO, AGRUP_HEADERS, AGRUP_COL,
 * agrupAnterior_, PDF_CARPETA_ID) y Cancelaciones.gs (estadoCancelaciones_).
 */

/** Por encima de este tamaño la ventana no ofrece la descarga directa, solo el link de Drive. */
const FP_MAX_BYTES_DESCARGA_DIRECTA = 8 * 1024 * 1024;

function descargarFacturasPendientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nombreDia = formatearNombreHoja_(new Date());
  const hoja = ss.getSheetByName(AGRUP_PREFIJO + nombreDia);

  if (!hoja || hoja.getLastRow() < 2) {
    ui.alert('No existe la hoja "' + AGRUP_PREFIJO + nombreDia + '" o está vacía. Generá primero la agrupada.');
    return;
  }

  const taskPorLink = fpTaskPorLink_(ss.getSheetByName(nombreDia));
  const anterior = agrupAnterior_(ss, nombreDia, estadoCancelaciones_(ss));

  const facturas = [];      // { link, nombre }
  const vistos = {};
  let grupos = 0, enBanco = 0, sinLink = 0;

  hoja.getRange(2, 1, hoja.getLastRow() - 1, AGRUP_HEADERS.length).getValues().forEach(v => {
    const tipo = String(v[AGRUP_COL.TIPO - 1] || '').trim();
    const ids = String(v[AGRUP_COL.IDS - 1] || '').split(',').map(s => s.trim()).filter(String);
    if (ids.length === 0 || TIPOS_VALIDABLES.indexOf(tipo) === -1) return;
    if (String(v[AGRUP_COL.ESTADO - 1] || '').trim() !== '') return;
    if (ids.some(t => anterior.arrastre[t])) { enBanco++; return; }

    const links = String(v[AGRUP_COL.LINKS - 1] || '').split('\n').map(s => s.trim()).filter(String);
    if (links.length === 0) { sinLink++; return; }
    grupos++;

    const benef = String(v[AGRUP_COL.BENEF - 1] || '').trim();
    links.forEach(l => {
      if (vistos[l]) return;
      vistos[l] = true;
      facturas.push({ link: l, nombre: (taskPorLink[l] || ids.join(',')) + (benef ? ' - ' + benef : '') });
    });
  });

  const salteadas =
    (enBanco ? '\n' + enBanco + ' grupo(s) "↻ EN BANCO" no se bajan: ya estaban cargados en el banco.' : '') +
    (sinLink ? '\n' + sinLink + ' grupo(s) sin link de factura.' : '');

  if (facturas.length === 0) {
    ui.alert('No hay facturas A o C pendientes de cargar en el banco (ESTADO vacío).' + salteadas);
    return;
  }

  const blobs = [];
  const fallidas = [];
  const usados = {};
  facturas.forEach(f => {
    const id = extraerIdDeDrive_(f.link);
    if (!id) { fallidas.push(f.nombre + ': el link no es de Drive'); return; }
    try {
      const blob = DriveApp.getFileById(id).getBlob();
      const ext = (String(blob.getName()).match(/\.[A-Za-z0-9]{2,5}$/) || ['.pdf'])[0];
      let nombre = fpNombreArchivo_(f.nombre);
      if (usados[nombre]) nombre += ' (' + (++usados[nombre]) + ')';
      else usados[nombre] = 1;
      blobs.push(blob.setName(nombre + ext));
    } catch (e) {
      fallidas.push(f.nombre + ': ' + e.message);
    }
  });

  if (blobs.length === 0) {
    ui.alert('No se pudo abrir ninguna factura:\n' + fallidas.join('\n'));
    return;
  }

  const nombreZip = 'Facturas pendientes ' + nombreDia + '.zip';
  const zip = Utilities.zip(blobs, nombreZip);
  const carpeta = DriveApp.getFolderById(PDF_CARPETA_ID);
  const previos = carpeta.getFilesByName(nombreZip);
  while (previos.hasNext()) previos.next().setTrashed(true);
  const archivo = carpeta.createFile(zip);

  const bytes = zip.getBytes();
  const directa = bytes.length <= FP_MAX_BYTES_DESCARGA_DIRECTA
    ? Utilities.base64Encode(bytes) : null;

  ui.showModalDialog(
    HtmlService.createHtmlOutput(fpHtmlDescarga_(nombreZip, blobs.length, grupos, archivo.getUrl(),
                                                 directa, fallidas, salteadas))
      .setWidth(520).setHeight(260 + Math.min(fallidas.length, 8) * 18),
    'Facturas pendientes'
  );
}

/** Link de factura → TASK, desde INVOICE/COMPROBANTE de la hoja del día. */
function fpTaskPorLink_(hojaDia) {
  const res = {};
  if (!hojaDia || hojaDia.getLastRow() < 2) return res;
  const idx = indicesPorHeader_(hojaDia);
  const iTask = idx['TASK'], iInv = idx['INVOICE/COMPROBANTE'];
  if (iTask < 0 || iInv < 0) return res;
  hojaDia.getRange(2, 1, hojaDia.getLastRow() - 1, hojaDia.getLastColumn()).getValues().forEach(v => {
    const t = String(v[iTask] == null ? '' : v[iTask]).trim();
    const l = String(v[iInv] == null ? '' : v[iInv]).trim();
    if (t !== '' && l !== '' && !res[l]) res[l] = t;
  });
  return res;
}

/** Nombre de archivo sin caracteres que Windows rechaza, recortado. */
function fpNombreArchivo_(s) {
  const limpio = String(s).replace(/[\\/:*?"<>|\n\r]+/g, '-').trim();
  return limpio.length > 120 ? limpio.substring(0, 120) : (limpio || 'factura');
}

function fpHtmlDescarga_(nombre, cantidad, grupos, url, base64, fallidas, salteadas) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const boton = base64
    ? '<a class="btn" download="' + esc(nombre) + '" href="data:application/zip;base64,' + base64 + '">⬇ Descargar</a> '
    : '';
  return '<style>body{font-family:Arial,sans-serif;font-size:13px;margin:0;padding:10px;line-height:1.6}' +
    '.btn{background:#0a0a8c;color:#fff;padding:5px 12px;border-radius:4px;text-decoration:none;font-weight:bold}' +
    '.err{background:#fce8e6;border-left:3px solid #cc0000;padding:6px 8px;margin-top:10px;font-size:12px}</style>' +
    '<div><b>' + esc(nombre) + '</b><br>' + cantidad + ' factura(s) de ' + grupos + ' grupo(s) A/C sin estado.</div>' +
    '<div style="margin-top:10px">' + boton + '<a href="' + esc(url) + '" target="_blank">Abrir en Drive</a></div>' +
    (base64 ? '' : '<div style="color:#666;font-size:12px">El ZIP es grande para bajarlo desde la ventana: ' +
                   'abrilo en Drive y apretá Descargar.</div>') +
    (salteadas ? '<div style="color:#666;font-size:12px;margin-top:8px">' + esc(salteadas).replace(/\n/g, '<br>') + '</div>' : '') +
    (fallidas.length ? '<div class="err"><b>No se pudieron bajar (' + fallidas.length + '):</b><br>' +
                       fallidas.map(esc).join('<br>') + '</div>' : '');
}
