/**
 * ARCHIVO DE CARGA A GALICIA OFFICE BANKING
 *
 * Toma de la agrupada de HOY los grupos A y C con ESTADO = CARGADO EN BANCO y arma la planilla de
 * importación de transferencias de Galicia ("Importá transferencias desde un archivo Excel"),
 * respetando la plantilla del banco tal cual: mismas hojas, validaciones, comentarios y formato.
 * Solo se escriben las columnas A (CBU/CVU/Alias/Nro cuenta), B (Monto) y C (Concepto = Varios);
 * Descripción, Email y Mensaje quedan vacíos.
 *
 *   generarArchivoGalicia()  → ítem de menú. Junta las filas, muestra qué entra y qué no, pide
 *                              confirmación, guarda el .xlsx en PDF_CARPETA_ID y abre una ventana
 *                              para descargarlo.
 *   conMontoAPagar_(fn)      → envuelve la generación de la agrupada para que la columna
 *                              MONTO A PAGAR sobreviva a la regeneración (ver abajo).
 *
 * ── MONTO QUE SE TRANSFIERE ──────────────────────────────────────────────────────────────────
 *   C → MONTO de la agrupada (el de la factura).
 *   A → MONTO A PAGAR: el neto de retenciones, que cargás a mano en esa columna. Va a la derecha
 *       de todo (después de APROB. PAYPAL), FUERA de AGRUP_HEADERS, igual que la columna de PayPal,
 *       para no cambiar nada de lo que ya leen los demás archivos.
 *       Se preserva al regenerar por el CONJUNTO EXACTO de IDs de la fila: si cambia la composición
 *       del grupo, el neto ya no corresponde a esas facturas y se pierde a propósito.
 *       ⚠️ Solo se preserva si la agrupada se regenera desde el menú (menuGenerarAgrupada /
 *       menuActualizarAgrupada). Si corrés generarHojaAgrupadaDelDia a mano desde el editor, la
 *       columna desaparece.
 *
 * ── QUÉ FILAS NO ENTRAN (y se listan con el motivo) ──────────────────────────────────────────
 *   - A sin MONTO A PAGAR, o con MONTO A PAGAR mayor que la factura (la retención resta, no suma).
 *   - Sin monto de factura verificado, o moneda distinta de ARS.
 *   - CBU/alias que la plantilla de Galicia rechazaría (CBU que no tiene 22 dígitos, alias fuera
 *     de 6–20 caracteres), o un grupo con dos cuentas distintas en la hoja del día.
 *   - Tareas que ya estaban CARGADO EN BANCO en la agrupada anterior ("↻ EN BANCO"): ya están en
 *     el banco y volver a subirlas es pagar dos veces. Es el mismo resguardo que tenía
 *     "Marcar todo como CARGADO EN BANCO".
 *   - Tareas canceladas 2 veces o más (Cancelaciones.gs).
 *
 * ── DOBLE CARGA EN EL MISMO DÍA ──────────────────────────────────────────────────────────────
 * Cada archivo generado deja anotado (Document Properties) qué filas salieron. Si generás otro
 * archivo el mismo día, esas filas se ofrecen aparte y por defecto NO se vuelven a incluir.
 *
 * ── LÍMITE DEL BANCO ─────────────────────────────────────────────────────────────────────────
 * Galicia acepta hasta 30 pagos por archivo. Con más, se arman varios archivos (_1, _2, …).
 *
 * ── CÓMO SE ARMA EL ARCHIVO ──────────────────────────────────────────────────────────────────
 * La plantilla del banco va embebida en base64 (GALICIA_PLANTILLA_BASE64, al final). Es un .xlsx
 * (el banco la publica con extensión .xls, pero adentro es xlsx). Se descomprime, se escriben las
 * celdas vacías de la hoja "Formulario" (xl/worksheets/sheet1.xml) y se vuelve a comprimir.
 * No pasa por Google Sheets: una conversión ida y vuelta podría perder validaciones o comentarios.
 *
 * DEPENDE de Pagoslb.gs (formatearNombreHoja_, aNumero_, redondear2_, formatearImporte_,
 * indicesPorHeader_, TIPOS_VALIDABLES, EST_CARGADO),
 * Hojaagrupada.gs (AGRUP_PREFIJO, AGRUP_HEADERS, AGRUP_COL, agrupAnterior_, PDF_CARPETA_ID),
 * Cancelaciones.gs (estadoCancelaciones_) y PaypalAprobaciones.gs (ampliarFiltro_).
 */

const CG_HEADER_PAGAR = 'MONTO A PAGAR';
const CG_TIPO_CON_RETENCION = 'A - Transf. Banc. Local (Arg)';
const CG_CONCEPTO = 'Varios';
const CG_MONEDA = 'ARS';
const CG_MAX_POR_ARCHIVO = 30;            // límite de Galicia por archivo
const CG_MONTO_MAX = 999999999999.99;     // 12 dígitos enteros y 2 decimales
const CG_HOJA_FORMULARIO = 'xl/worksheets/sheet1.xml';
const CG_PROP_PREFIJO = 'CG_EXPORTADAS ';
const CG_MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/* ══════════════════ COLUMNA MONTO A PAGAR ══════════════════ */

/** Columna (1-based) de MONTO A PAGAR en la hoja, o -1. */
function cgColumnaPagar_(hoja) {
  const ancho = Math.max(hoja.getLastColumn(), 1);
  const hdr = hoja.getRange(1, 1, 1, ancho).getValues()[0]
    .map(h => String(h == null ? '' : h).trim().toUpperCase());
  const i = hdr.indexOf(CG_HEADER_PAGAR);
  return i < 0 ? -1 : i + 1;
}

/** Clave de preservación: el conjunto ordenado de IDs de la fila. '' si es banda o resumen. */
function cgClaveIds_(valor) {
  return String(valor == null ? '' : valor).split(',').map(s => s.trim()).filter(String).sort().join(',');
}

/** IDs → MONTO A PAGAR de la agrupada que existe ahora. Vacío si no hay hoja o columna. */
function cgLeerMontosAPagar_(hoja) {
  const res = {};
  if (!hoja || hoja.getLastRow() < 2) return res;
  const col = cgColumnaPagar_(hoja);
  if (col < 0) return res;
  const n = hoja.getLastRow() - 1;
  const ids = hoja.getRange(2, AGRUP_COL.IDS, n, 1).getValues();
  const vals = hoja.getRange(2, col, n, 1).getValues();
  ids.forEach((v, i) => {
    const k = cgClaveIds_(v[0]);
    const m = aNumero_(vals[i][0]);
    if (k !== '' && m !== null) res[k] = m;
  });
  return res;
}

/** Corre la generación de la agrupada y después repone la columna MONTO A PAGAR. */
function conMontoAPagar_(fn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = AGRUP_PREFIJO + formatearNombreHoja_(new Date());
  const antes = cgLeerMontosAPagar_(ss.getSheetByName(nombre));
  fn();
  const hoja = ss.getSheetByName(nombre);
  if (hoja) cgEscribirColumnaPagar_(hoja, antes);
}

/**
 * Crea (si hace falta) y llena la columna MONTO A PAGAR. Va como mínimo en la columna 17: la 16 es
 * de APROB. PAYPAL, que se escribe siempre en AGRUP_HEADERS.length + 1 aunque todavía no exista.
 * Las filas tipo A quedan en blanco para cargar; el resto toma el fondo de su fila.
 */
function cgEscribirColumnaPagar_(hoja, valores) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return;

  let col = cgColumnaPagar_(hoja);
  if (col < 0) col = Math.max(hoja.getLastColumn() + 1, AGRUP_HEADERS.length + 2);
  if (hoja.getMaxColumns() < col) {
    hoja.insertColumnsAfter(hoja.getMaxColumns(), col - hoja.getMaxColumns());
  }

  hoja.getRange(1, col).setValue(CG_HEADER_PAGAR)
    .setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#434343').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNote('Solo tipo A: lo que se transfiere, neto de retenciones.\n' +
             'Lo usa "Galicia: generar archivo de carga". Las tipo C transfieren el MONTO.');

  const n = ultima - 1;
  const ids = hoja.getRange(2, AGRUP_COL.IDS, n, 1).getValues();
  const tipos = hoja.getRange(2, AGRUP_COL.TIPO, n, 1).getValues();
  const fondos = hoja.getRange(2, AGRUP_COL.COMPROBANTE, n, 1).getBackgrounds();
  const salida = ids.map((v, i) => {
    const k = cgClaveIds_(v[0]);
    if (k !== '' && String(tipos[i][0]).trim() === CG_TIPO_CON_RETENCION) fondos[i][0] = '#ffffff';
    return [k !== '' && Object.prototype.hasOwnProperty.call(valores, k) ? valores[k] : ''];
  });

  hoja.getRange(2, col, n, 1).setValues(salida).setBackgrounds(fondos)
    .setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  hoja.getRange(1, col, ultima, 1)
    .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  hoja.setColumnWidth(col, 120);
  ampliarFiltro_(hoja, col);
}

/* ══════════════════ GENERACIÓN DEL ARCHIVO ══════════════════ */

function generarArchivoGalicia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nombreDia = formatearNombreHoja_(new Date());
  const hoja = ss.getSheetByName(AGRUP_PREFIJO + nombreDia);
  const hojaDia = ss.getSheetByName(nombreDia);

  if (!hoja || hoja.getLastRow() < 2) {
    ui.alert('No existe la hoja "' + AGRUP_PREFIJO + nombreDia + '" o está vacía. Generá primero la agrupada.');
    return;
  }
  if (!hojaDia) {
    ui.alert('No existe la hoja del día "' + nombreDia + '": de ahí sale el CBU de cada tarea.');
    return;
  }

  const res = cgFilasParaElBanco_(ss, hoja, hojaDia, nombreDia);

  const prop = PropertiesService.getDocumentProperties();
  const claveProp = CG_PROP_PREFIJO + nombreDia;
  const exportadas = JSON.parse(prop.getProperty(claveProp) || '[]');
  const repetidas = res.ok.filter(f => exportadas.indexOf(f.clave) !== -1);
  let filas = res.ok;

  const textoExcluidas = res.excluidas.length === 0 ? '' :
    '\n\n❌ NO ENTRAN (' + res.excluidas.length + '):\n' +
    res.excluidas.map(e => '• ' + e.benef + ' (' + e.ids + ') — ' + e.motivo).join('\n');

  if (filas.length === 0) {
    ui.alert('No hay ninguna fila A o C en "' + EST_CARGADO + '" que se pueda cargar al banco.' + textoExcluidas);
    return;
  }

  if (repetidas.length > 0) {
    const r = ui.alert(
      '⚠️ YA SALIERON EN UN ARCHIVO DE HOY\n\n' +
      repetidas.map(f => '• ' + f.benef + ' — ' + formatearImporte_(f.monto) + ' (' + f.ids + ')').join('\n') +
      '\n\nSi ya las subiste al banco, incluirlas de nuevo es pagarlas dos veces.\n\n' +
      'SÍ = incluirlas igual · NO = dejarlas afuera · CANCELAR = no generar nada',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (r === ui.Button.CANCEL || r === ui.Button.CLOSE) return;
    if (r === ui.Button.NO) filas = filas.filter(f => repetidas.indexOf(f) === -1);
    if (filas.length === 0) { ui.alert('No queda ninguna fila para cargar.'); return; }
  }

  const total = redondear2_(filas.reduce((s, f) => s + f.monto, 0));
  const partes = Math.ceil(filas.length / CG_MAX_POR_ARCHIVO);
  const obs = filas.filter(f => f.obs);
  const resp = ui.alert(
    filas.length + ' transferencia(s) · ARS ' + formatearImporte_(total) +
    (partes > 1 ? ' · ' + partes + ' archivos (Galicia acepta ' + CG_MAX_POR_ARCHIVO + ' por archivo)' : '') + '\n\n' +
    filas.map(f => '• ' + f.tipo.charAt(0) + ' · ' + f.benef + ' — ' + formatearImporte_(f.monto) + ' → ' + f.dato).join('\n') +
    (obs.length ? '\n\n⚠️ OBSERVACIONES:\n' + obs.map(f => '• ' + f.benef + ' — ' + f.obs).join('\n') : '') +
    textoExcluidas + '\n\n¿Generar el archivo?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const carpeta = DriveApp.getFolderById(PDF_CARPETA_ID);
  const archivos = [];
  for (let p = 0; p < partes; p++) {
    const lote = filas.slice(p * CG_MAX_POR_ARCHIVO, (p + 1) * CG_MAX_POR_ARCHIVO);
    const nombre = 'Galicia ' + nombreDia + (partes > 1 ? '_' + (p + 1) : '') + '.xlsx';
    const blob = cgArmarXlsx_(lote, nombre);

    const previos = carpeta.getFilesByName(nombre);
    while (previos.hasNext()) previos.next().setTrashed(true);
    const archivo = carpeta.createFile(blob);

    archivos.push({
      nombre: nombre,
      cantidad: lote.length,
      total: redondear2_(lote.reduce((s, f) => s + f.monto, 0)),
      url: archivo.getUrl(),
      data: Utilities.base64Encode(blob.getBytes())
    });
  }

  // Se anota recién después de generar: si algo falla antes, no queda marcado como exportado.
  filas.forEach(f => { if (exportadas.indexOf(f.clave) === -1) exportadas.push(f.clave); });
  cgLimpiarPropsViejas_(prop, claveProp);
  prop.setProperty(claveProp, JSON.stringify(exportadas));

  ui.showModalDialog(
    HtmlService.createHtmlOutput(cgHtmlDescarga_(archivos)).setWidth(520).setHeight(160 + archivos.length * 70),
    'Archivo de carga a Galicia'
  );
}

/**
 * Lee la agrupada y separa lo que va al banco de lo que no.
 * ok:        [{ clave, tipo, benef, ids, dato, monto, obs }]
 * excluidas: [{ benef, ids, motivo }]
 */
function cgFilasParaElBanco_(ss, hoja, hojaDia, nombreDia) {
  const datoPorTask = cgDatosDePago_(hojaDia);
  const canc = estadoCancelaciones_(ss);
  const anterior = agrupAnterior_(ss, nombreDia, canc);
  const colPagar = cgColumnaPagar_(hoja);

  const ancho = Math.max(hoja.getLastColumn(), AGRUP_HEADERS.length);
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, ancho).getValues();
  const ok = [], excluidas = [];

  valores.forEach(v => {
    const tipo = String(v[AGRUP_COL.TIPO - 1] || '').trim();
    const ids = String(v[AGRUP_COL.IDS - 1] || '').split(',').map(s => s.trim()).filter(String);
    if (ids.length === 0) return;                                          // banda o resumen
    if (TIPOS_VALIDABLES.indexOf(tipo) === -1) return;
    if (String(v[AGRUP_COL.ESTADO - 1] || '').trim().toUpperCase() !== EST_CARGADO) return;

    const benef = String(v[AGRUP_COL.BENEF - 1] || '(sin beneficiario)');
    const idsTxt = ids.join(', ');
    const fuera = motivo => excluidas.push({ benef: benef, ids: idsTxt, motivo: motivo });

    const enBanco = ids.filter(t => anterior.arrastre[t]);
    if (enBanco.length) {
      fuera('ya estaba CARGADO EN BANCO en la agrupada del ' + anterior.arrastre[enBanco[0]].fecha +
            ': está en el banco desde ese día. Si hay que volver a subirla, cargala a mano.');
      return;
    }
    if (ids.some(t => canc[t] && canc[t].veces >= 2)) {
      fuera('cancelada 2 veces o más: no se repaga, hay que avisar en AppSheet');
      return;
    }

    const moneda = String(v[AGRUP_COL.MONEDA - 1] || '').trim().toUpperCase();
    if (moneda !== CG_MONEDA) { fuera('moneda ' + (moneda || 'vacía') + ', Galicia transfiere en ' + CG_MONEDA); return; }

    const montoFactura = aNumero_(v[AGRUP_COL.MONTO - 1]);
    if (montoFactura === null || montoFactura <= 0) { fuera('sin monto de factura verificado'); return; }

    let monto = montoFactura;
    if (tipo === CG_TIPO_CON_RETENCION) {
      const neto = colPagar > 0 ? aNumero_(v[colPagar - 1]) : null;
      if (neto === null) { fuera('tipo A sin ' + CG_HEADER_PAGAR + ' (el neto de retenciones)'); return; }
      if (neto <= 0) { fuera(CG_HEADER_PAGAR + ' tiene que ser mayor que cero'); return; }
      if (redondear2_(neto - montoFactura) > 0) {
        fuera(CG_HEADER_PAGAR + ' ' + formatearImporte_(neto) + ' es mayor que la factura ' +
              formatearImporte_(montoFactura) + ': la retención resta, no suma');
        return;
      }
      monto = neto;
    }
    monto = redondear2_(monto);
    if (monto > CG_MONTO_MAX) { fuera('monto mayor al máximo de Galicia (12 dígitos enteros)'); return; }

    const datos = [];
    ids.forEach(t => {
      const d = datoPorTask[t];
      if (d && datos.indexOf(d) === -1) datos.push(d);
    });
    if (datos.length === 0) { fuera('no encuentro el CBU de ' + idsTxt + ' en la hoja del día'); return; }
    if (datos.length > 1) { fuera('el grupo tiene más de una cuenta: ' + datos.join(' / ')); return; }
    const cuenta = cgCuentaGalicia_(datos[0]);
    if (!cuenta.ok) { fuera(cuenta.motivo); return; }

    ok.push({
      clave: cgClaveIds_(ids.join(',')),
      tipo: tipo,
      benef: benef,
      ids: idsTxt,
      dato: cuenta.dato,
      monto: monto,
      obs: cuenta.obs || ''
    });
  });

  return { ok: ok, excluidas: excluidas };
}

/** TASK → CBU / DATO DE PAGO tal como está en la hoja del día. */
function cgDatosDePago_(hojaDia) {
  const res = {};
  if (hojaDia.getLastRow() < 2) return res;
  const idx = indicesPorHeader_(hojaDia);
  const iTask = idx['TASK'], iDato = idx['CBU / DATO DE PAGO'];
  if (iTask < 0 || iDato < 0) return res;
  hojaDia.getRange(2, 1, hojaDia.getLastRow() - 1, hojaDia.getLastColumn()).getValues().forEach(v => {
    const t = String(v[iTask] == null ? '' : v[iTask]).trim();
    const d = String(v[iDato] == null ? '' : v[iDato]).trim();
    if (t !== '' && d !== '') res[t] = d;
  });
  return res;
}

/**
 * Normaliza el dato de pago a lo que acepta la columna A de la plantilla:
 *   CBU/CVU      22 dígitos (se sacan espacios y guiones);
 *   Alias        6 a 20 caracteres: letras, números, punto y guión. No se corrige nada: un alias
 *                con espacios no entra, porque sacarle los espacios podría mandar la plata a otro;
 *   Nro cuenta   cuenta Galicia, solo dígitos. Se acepta de 6 a 19 dígitos con una observación.
 * 20, 21 o más de 22 dígitos es casi seguro un CBU mal copiado: no entra.
 */
function cgCuentaGalicia_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (/^[\d\s-]+$/.test(s)) {
    const d = s.replace(/[\s-]/g, '');
    if (d.length === 22) return { ok: true, dato: d };
    if (d.length >= 20) return { ok: false, motivo: 'CBU de ' + d.length + ' dígitos (tiene que tener 22): ' + d };
    if (d.length >= 6) {
      return { ok: true, dato: d, obs: 'el dato ' + d + ' no es un CBU de 22 dígitos: se carga como número de cuenta Galicia' };
    }
    return { ok: false, motivo: 'dato de pago "' + s + '" demasiado corto' };
  }
  if (/^[A-Za-z0-9.-]{6,20}$/.test(s)) return { ok: true, dato: s };
  return { ok: false, motivo: 'dato de pago "' + s + '" no es un CBU, CVU ni alias válido para Galicia' };
}

/* ══════════════════ ARMADO DEL XLSX ══════════════════ */

/** Plantilla de Galicia con las filas escritas en la hoja Formulario, como blob .xlsx. */
function cgArmarXlsx_(filas, nombre) {
  const plantilla = Utilities.newBlob(Utilities.base64Decode(GALICIA_PLANTILLA_BASE64), 'application/zip', 'plantilla.zip');
  const partes = Utilities.unzip(plantilla);
  let encontrada = false;
  const nuevas = partes.map(b => {
    if (b.getName() !== CG_HOJA_FORMULARIO) return b;
    encontrada = true;
    return Utilities.newBlob(cgLlenarFormulario_(b.getDataAsString('UTF-8'), filas), 'application/xml', CG_HOJA_FORMULARIO);
  });
  if (!encontrada) throw new Error('La plantilla embebida no tiene ' + CG_HOJA_FORMULARIO + '.');
  return Utilities.zip(nuevas, nombre).setContentType(CG_MIME_XLSX).setName(nombre);
}

/**
 * Escribe A (dato), B (monto) y C (concepto) desde la fila 2. Solo reemplaza celdas que en la
 * plantilla están vacías; si no encuentra una, corta: es preferible a un archivo a medio armar.
 */
function cgLlenarFormulario_(xml, filas) {
  filas.forEach((f, i) => {
    const r = i + 2;
    xml = cgPonerCelda_(xml, 'A' + r, 'inlineStr', '<is><t>' + cgEscaparXml_(f.dato) + '</t></is>');
    xml = cgPonerCelda_(xml, 'B' + r, null, '<v>' + f.monto.toFixed(2) + '</v>');
    xml = cgPonerCelda_(xml, 'C' + r, 'inlineStr', '<is><t>' + CG_CONCEPTO + '</t></is>');
  });
  return xml;
}

function cgPonerCelda_(xml, ref, tipo, contenido) {
  const re = new RegExp('<c r="' + ref + '"( s="\\d+")?\\s*/>');
  if (!re.test(xml)) throw new Error('La plantilla no tiene la celda ' + ref + ' vacía.');
  return xml.replace(re, (m, estilo) =>
    '<c r="' + ref + '"' + (estilo || '') + (tipo ? ' t="' + tipo + '"' : '') + '>' + contenido + '</c>');
}

function cgEscaparXml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Borra las marcas de exportación de otros días: solo importa la de hoy. */
function cgLimpiarPropsViejas_(prop, claveHoy) {
  Object.keys(prop.getProperties()).forEach(k => {
    if (k.indexOf(CG_PROP_PREFIJO) === 0 && k !== claveHoy) prop.deleteProperty(k);
  });
}

/**
 * Ventana de descarga. El botón baja el archivo directo desde el navegador; el link de Drive queda
 * como alternativa si el navegador bloquea la descarga desde la ventana.
 */
function cgHtmlDescarga_(archivos) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const filas = archivos.map(a =>
    '<div class="a"><b>' + esc(a.nombre) + '</b> · ' + a.cantidad + ' transferencia(s) · ARS ' +
    esc(formatearImporte_(a.total)) + '<br>' +
    '<a class="btn" download="' + esc(a.nombre) + '" href="data:' + CG_MIME_XLSX + ';base64,' + a.data + '">⬇ Descargar</a>' +
    ' <a href="' + esc(a.url) + '" target="_blank">Abrir en Drive</a></div>'
  ).join('');
  return '<style>body{font-family:Arial,sans-serif;font-size:13px;margin:0;padding:10px}' +
    '.a{margin-bottom:12px;line-height:1.9}' +
    '.btn{background:#0a0a8c;color:#fff;padding:5px 12px;border-radius:4px;text-decoration:none;font-weight:bold}</style>' +
    filas + '<div style="color:#666;font-size:11px">Subilo en Galicia Office: Transferencias → Importá desde un archivo Excel.</div>';
}

/**
 * Plantilla_Transferencias_Galicia tal como la publica el banco (xlsx en base64).
 * Si Galicia cambia la plantilla, reemplazar esta cadena por la nueva:
 *   en una terminal: base64 -w0 Plantilla_Transferencias_Galicia.xls
 */
const GALICIA_PLANTILLA_BASE64 =
  'UEsDBBQABgAIAAAAIQDvQuZEuwEAABsIAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADMldtqGzEQhu8LfYdFt8UrJ4VSite5aNrLNpD0ARRp' +
  'vCu8OqAZO/bbdyTboQQfWNaQ3uxB2vm/XxrtzOxu4/pqDQlt8I24qaeiAq+Dsb5txJ+nn5OvokJS3qg+eGjEFlDczT9+mD1tI2DF' +
  '0R4b0RHFb1Ki7sAprEMEzzOLkJwifk2tjEovVQvydjr9InXwBJ4mlDXEfHYPC7Xqqfqx4eGdk2frRfV9911GNULF2FutiI3KtTdv' +
  'IJOwWFgNJuiVY+kaYwJlsAMg19cxWSamRyDihaGQR5nRt2+Y1mXPefx4RIIeh9nc70PNkWUp2NmIn3izThDWPDNmHzj+PqkXXvUJ' +
  'QEafBuyN/eYTkqyB6kEl+qUcp0NuevkS0vI5hGV9XmRotkrWaqesP2zMGX75GGW53VzZSF5fER7o4/Y/8fH5nXwQlwGQ5To+JUXm' +
  'QgKQtj3gtY9hEb1E7lQC80hcYNqrG/hX+4IPHVyue3jtX+CgewFvdjUG5f5hvI290DkuF/uHFCJyP0kwPPmHUpyjJ5GFIJGF12J8' +
  'rOa8ErkZjT5tkLudATOUrVdIwY3G72SOwGVp7fO/AAAA//8DAFBLAwQUAAYACAAAACEAE16+ZQIBAADfAgAACwAIAl9yZWxzLy5y' +
  'ZWxzIKIEAiigAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKyS' +
  'TUsDMRCG74L/Icy9O9sqItJsL0XoTWT9ATGZ/WA3mZCkuv33RkF0obYeepyvd555mfVmsqN4oxB7dhKWRQmCnGbTu1bCS/24uAcR' +
  'k3JGjexIwoEibKrrq/UzjSrlodj1Poqs4qKELiX/gBh1R1bFgj25XGk4WJVyGFr0Sg+qJVyV5R2G3xpQzTTFzkgIO3MDoj74vPm8' +
  'NjdNr2nLem/JpSMrkKZEzpBZ+JDZQurzNaJWoaUkwbB+yumIyvsiYwMeJ1r9n+jva9FSUkYlhZoDneb57DgFtLykRXMTf9yZRnzn' +
  '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' +
  'ABMAAAB4bC9tZWRpYS9pbWFnZTEucG5niVBORw0KGgoAAAANSUhEUgAAA0EAAADrCAMAAACYTG3+AAADAFBMVEX/eBz/qWn/gjD/' +
  'pWn/4s3/1bj/rnX/sXn/zan/48z/59b/chL/1Lb/iz//dRv/bAD/fir/jD//toH/sn7/9vD/z6//cxf/kTv/ZAD/gin/iT3/wZb/' +
  'rXH/7OD/5dL/dg7/n1//cwf/eCD/6tv/7uP/3cL/qXD/rnj/ZgD/nVX/9e//pGT/mkz/xqD/kEX/uYb/oVr/+PT/y6T/0LD/lkX/' +
  'sHf/uor/////+/j/bgX/lU3/0rT/6tz/pGH/wJH/xJv/bg7/gi7/8Ob/4Mv//Pr/aQf/fSX/t4T/yKL/aQL/2r//2Lz/kUD/bAn/' +
  'hSn/lVD/kUj/lUn/8un/6Nj/ol//y6b/v5L/hzf/9O3/mkf/fRT/hy3/w5r/n2D/ijj/aQD/3sf/n1v/vIr/8uj/28P/vI3/0rH/' +
  '38j/mlj/mlD/rG7/agD/jjr/1LT/oWH/vJD/kEf/eiT/fRz/ijH/xpz/v5D/t37/7+P/iSv/mlL/fSH/3MT/dhX/iTb/nVL/lkL/' +
  '4Mj/mFD/y6j/z6z/p2v/jTH/dhL/t4b/eBn/0K7/fBr/soD/omT/hjj/hDD/mVX/bxH/giX/hDT/0rD/n1j/awr/iTz/qGT/k0P/' +
  'xJf/hyX/q23/6dr/nFj/nFv/yaT/4cv/sHv/lkj/iTX/yZ3/izb/uov/q2//pV3/xZv/lEv/sn3/2sD/2b//jDf/xZ7/bgD/2L//' +
  'giP/48//9u//eAn/qW3/n1z/xp//+PP/vY//7eD/qWz/7N//lkr/xp7/xpv/giL/eAf/+/f/vI7/exr/+vf/8ef/mFL/8uf/2b7/' +
  'tYT/vI//vIT/w5z//Pn//fv/bAP//v3/0bP/s37/ZQD/giD/izT/+vb/hCL/59f/x57/8OX/x5//3cb/s3r/vIP/ZQH/+fX/snz/' +
  'vY3/9/H/2sH/nFz/17b/iz3/awv/vo//7N7/kkf/t3v/7eH/7uH//v7/28H/8+v/sX7/ZwP/v4z/y6f/5ND/qG7/vZL/693/qG3/' +
  'ehf/4Mf/27+KUbLFAAAkrElEQVR42u2dCZyNVR/Hv8gYI0x2IlvLSAihUqnXSJEte0P2XZaQbSyNXbxGkRnLMLIzGWVphKK8hddS' +
  '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' +
  'AAAAeGwvY29tbWVudHMxLnhtbMyV3W6bMBTH7yv1HSzuXb4cwFGSCgxMmdqmalNpty44iTXAyHaitNMeZs+yF5sJSVe1msTNqkqA' +
  '7MPhf45/nGNPLvd1BXZMKi6aqeVeOBZgTSFK3qyn1sMyh5EFlKZNSSvRsKn1xJR1OTs/mxSirlmjFTACjZpaG63bsW2rYsNqqi5E' +
  'yxrzZiVkTbWZyrWtWsloqTaM6bqyPccJ7JryxuoVxnUxRKSm8vu2hSZ6SzV/5BXXTwctC9TFeL5uhKSPlUl0L0/Ce/lOuOaFFEqs' +
  '9IURssVqxQv2Lj8X2ZLteEfGmk3oVm+EVKfB7OuWNsDEowUXIK52VLLnid17zU4D437kdMWVfpkAyVZTK3Yt0PvPy6llwKsNbdlx' +
  'vJfjLTfmH0GEEclSAnOEMoj8fASjhBCYENfJnSSJsev/NPlptjcBpLluzUM9gx2tpha27C5qJSTgTcn2zEhGbmeUuWh077SkG1HT' +
  'zriiNa+eeqvXGeyDmp5BQGjdCrBIruZf4uXibr44P4MgP/xeAUoGClaVdAyW2bfl4VXWaMlAAJ6A54GCSlpoJpnqvrrnDWCq7cgp' +
  '0HCw3hrETE1sswDbJG/3a7GP6N5gSwZh890QoTBAkJBkBJHj+zAiJIJB6AXI3CQK8WfARnpsNw/X2d3iHnR4rn//2vNaANcDJV9z' +
  'bSCZRmOmYDuYL7aSFbym1WBuZBA35OQGkedAF3uGWxz4EOM0g24YeUmCRhmJPxO3Q7kNLJx0WL+NoiDFOIQjMkIQuYkPk8CPoU9i' +
  'H0dpGrte9mGFMwaLWzJf3MRXb7rtWDYv3faqZv722kAu2SAurkMQzrGphDRIISJhCKMgINA3sIg3yjAOog9sqP+PJR+ExYnSLCN5' +
  'AvPENA1ysAcTnHrQwzHKHD9IiZd+WizmDH61Nf+7XE47cX+InWZq9gcAAP//AwBQSwMEFAAGAAgAAAAhAHkXSYNiAQAAdAQAACcA' +
  'AAB4bC9wcmludGVyU2V0dGluZ3MvcHJpbnRlclNldHRpbmdzMi5iaW7sUctKw0AUPUlE1I39gC6Ce6Gtwdpl2yQlJTEhSbXbtB3L' +
  'QDoJaSr4+g//wv9w6dov8APc6p1QxYWI4kpwLnfOYe45d14eOKYokGFJeYYSOsYIEBGatDLFCgswiKpySkpOjBHqkEPZ0DYfcatp' +
  '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';
