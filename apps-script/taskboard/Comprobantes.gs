/**
 * ARMADO DE COMPROBANTES FINALES DE PAGO
 *
 * Dos formas de hacer lo mismo:
 *
 *   abrirPanelComprobante()      Panel lateral, una fila por vez. Para un comprobante suelto.
 *   abrirPanelComprobantesLote() Ventana grande. Cargás todos los archivos de una (capturas de
 *                                Wise/PIX, comprobantes de Galicia y retenciones en PDF), los vas
 *                                viendo uno por uno y los asignás a su fila. Al final genera todos
 *                                los PDF.
 *
 * El PDF final queda así:
 *   1) ORDEN DE PAGO  (la que ya genera crearPdfDeGrupo_: logo, tabla, links)
 *   2) COMPROBANTE    (la captura del pago, o el PDF de Galicia)
 *   3) DOCUMENTACIÓN  (orden de pago del sistema y/o certificados de retención; solo desde el
 *                      panel de lote. Las que retienen, solo en tipo A)
 *
 * Se guarda en la carpeta de Drive que mira AppSheet y el link queda en la columna COMPROBANTE.
 *
 * ── COLUMNAS ─────────────────────────────────────────────────────────────────────────────────
 * La agrupada se lee por nombre de encabezado (colsAgrup_, en Hojaagrupada.gs), no por posición:
 * los paneles se usan también sobre agrupadas de días anteriores, que pueden tener otro diseño.
 *
 * ── CÓMO SE ASIGNA UNA CAPTURA A SU FILA (panel de lote) ─────────────────────────────────────
 * El OCR de Drive NO se usa acá: su cuota es la misma que necesita validarSeleccion para leer las
 * facturas, y en una prueba real murió en la segunda imagen.
 * En cambio, cada captura se lee con OCR EN EL NAVEGADOR (Tesseract.js, cargado desde jsdelivr).
 * Corre en tu computadora y no gasta ninguna cuota de Google. Con ese texto el panel busca la fila:
 *   1) si aparece el código de una tarea (ej. MER-LIN-10-297)          → verificado;
 *   2) si no, si coinciden el dato de pago (mail, chave PIX, cuenta) o el nombre del beneficiario
 *      Y el monto, en UNA sola fila, sin una moneda que contradiga    → verificado;
 *   3) cualquier otra cosa (solo monto, varias filas, nada)           → sin sugerencia: se elige
 *      con clic y queda SIN VERIFICAR, como antes.
 * La sugerencia se confirma con Enter o con el botón verde. Tipear el código sigue funcionando
 * como alternativa. Si Tesseract no carga, el panel lo dice y se sigue a mano.
 *
 * ── POR QUÉ SE GENERA DE A UNO, DESDE EL NAVEGADOR ───────────────────────────────────────────
 * Cada PDF son 4 o 5 segundos entre crear el Doc, insertar la imagen, exportar y guardar. Con 39
 * operaciones son más de 3 minutos, y Apps Script corta la ejecución a los 6. Por eso el panel
 * llama al servidor UNA VEZ POR OPERACIÓN: cada ejecución dura segundos y el límite nunca se toca.
 * Si se corta o cerrás la ventana, lo ya generado queda hecho y con su link puesto.
 *
 * ── COMPROBANTES DE GALICIA Y RETENCIONES EN PDF ─────────────────────────────────────────────
 * Apps Script no tiene librería para concatenar PDFs, así que se unen EN EL NAVEGADOR, dentro del
 * panel de lote, con pdf-lib y pdf.js cargados desde cdnjs:
 *   - pdf.js lee el texto del PDF (sin OCR de Drive, sin gastar su cuota) y dibuja la 1ª página;
 *   - el servidor genera la orden de pago y la devuelve en base64 (comprobanteOrdenPdf);
 *   - pdf-lib une orden → Galicia → retenciones y el servidor guarda el resultado con el nombre
 *     de los códigos, en la carpeta de AppSheet (comprobanteGuardarFinal).
 * El panel de a una fila sigue aceptando solo imágenes.
 *
 * DEPENDE de Hojaagrupada.gs (AGRUP_PREFIJO, colsAgrup_, valAgrup_, crearPdfDeGrupo_,
 * blobOrdenDePago_, guardarPdfEnCarpeta_, nombreArchivoPdf_, PDF_FUENTE, PDF_CARPETA_ID) y de
 * Pagoslb.gs (aNumero_, formatearImporte_, indicesPorHeader_, extraerIdDeDrive_,
 * formatearNombreHoja_, normalizarNombre_, TIPOS_VALIDABLES, EST_*).
 */

/** Carpeta de Drive que mira AppSheet: ahí va el PDF final. */
const COMPROB_CARPETA_ID = '1rbTV514bc4vb1hVurllKxRUVJrl-8Qwg';

/**
 * Carpeta "Payments_receipt_reciever" (dueño process.private@): la que carga el comprobante en AS.
 * Cada PDF final se COPIA acá apenas se genera. Se copia y no se mueve porque el link
 * "Ver comprobante" apunta al archivo de COMPROB_CARPETA_ID, y el proceso de AS saca los archivos
 * de esta carpeta al procesarlos: si se moviera, el link quedaría roto.
 */
const COMPROB_CARPETA_AS_ID = '1Hgbn-05JRaCjkcHQdHVhO0sxFKhkN4DN';

/** Tipos de archivo aceptados. Un PDF se rechaza: no se puede concatenar (ver encabezado). */
const COMPROB_MIMES_OK = ['image/png', 'image/jpeg', 'image/jpg'];

/** El único tipo que puede llevar retenciones: factura A. */
const COMPROB_TIPO_CON_RETENCION = 'A - Transf. Banc. Local (Arg)';

/** Área útil de una página carta con márgenes de 1 pulgada, en puntos. */
const COMPROB_ANCHO_PAGINA = 468;
const COMPROB_ALTO_PAGINA = 620;

/* ══════════════════ NÚCLEO COMPARTIDO ══════════════════ */

/**
 * Arma el PDF final de una fila y le pone el link en la columna COMPROBANTE.
 * archivo = { data: base64 sin el prefijo "data:", mimeType: string, nombre: string }
 * Devuelve siempre { ok: bool, ... } para que el panel muestre el motivo sin romperse.
 */
function comprobanteArmarEnHoja_(nombreHoja, fila, archivo) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, motivo: 'No encuentro la hoja "' + nombreHoja + '".' };

    if (COMPROB_MIMES_OK.indexOf(String(archivo.mimeType).toLowerCase()) === -1) {
      return {
        ok: false,
        motivo: 'Solo puedo insertar imágenes (PNG o JPG); este archivo es ' + archivo.mimeType +
                '. Si es el comprobante de Galicia o una retención en PDF, todavía no está ' +
                'implementado: Apps Script no puede concatenar PDFs.'
      };
    }

    const op = operacionParaPdf_(hoja, fila);
    if (!op.ok) return op;

    const blob = Utilities.newBlob(
      Utilities.base64Decode(archivo.data),
      archivo.mimeType,
      String(archivo.nombre || 'comprobante')
    );

    const url = crearPdfDeGrupo_(op.datos, DriveApp.getFolderById(COMPROB_CARPETA_ID), blob);
    marcarComprobanteArmado_(hoja, fila, url);
    copiarComprobanteAAS_(hoja, fila, url);

    return { ok: true, url: url, id: extraerIdDeDrive_(url), nombre: nombreArchivoPdf_(op.datos.ids) + '.pdf' };

  } catch (err) {
    return { ok: false, motivo: 'Error armando el PDF: ' + err.message };
  }
}

/**
 * Lee la fila de la agrupada y arma los datos de la orden de pago. Es la misma validación para
 * las capturas y para los PDF de Galicia: sin IDs no es una operación, y sin monto verificado no
 * se arma nada (para A y C el monto tiene que venir de la factura).
 */
function operacionParaPdf_(hoja, fila) {
  const C = colsAgrup_(hoja);
  if (C.IDS < 0) return { ok: false, motivo: 'La hoja "' + hoja.getName() + '" no tiene la columna TRANSACCIONES (IDs).' };
  const v = hoja.getRange(fila, 1, 1, C.ANCHO).getValues()[0];
  const ids = String(valAgrup_(v, C, 'IDS') || '').trim();
  if (ids === '') return { ok: false, motivo: 'La fila ' + fila + ' ya no es una operación.' };

  if (aNumero_(valAgrup_(v, C, 'MONTO')) === null) {
    return { ok: false, motivo: 'Esa fila no tiene monto verificado. Si es de tipo A o C, leé la ' +
                                'factura con "Validar beneficiario y monto" y regenerá la agrupada.' };
  }

  return {
    ok: true,
    datos: {
      tipo: valAgrup_(v, C, 'TIPO'),
      beneficiario: valAgrup_(v, C, 'BENEF'),
      ids: ids,
      monto: aNumero_(valAgrup_(v, C, 'MONTO')),
      moneda: valAgrup_(v, C, 'MONEDA'),
      links: String(valAgrup_(v, C, 'LINKS') || '').split('\n').map(s => s.trim()).filter(String),
      fecha: hoja.getName().replace(AGRUP_PREFIJO, '')
    }
  };
}

/**
 * Pone el link en COMPROBANTE y pasa el estado. El comprobante armado cierra el circuito. No se
 * pisan RECHAZADO ni CANCELADO, porque ahí tener un comprobante es una contradicción que conviene
 * que se vea.
 */
function marcarComprobanteArmado_(hoja, fila, url) {
  const C = colsAgrup_(hoja);
  if (C.COMPROBANTE < 0) throw new Error('la hoja no tiene la columna COMPROBANTE');
  hoja.getRange(fila, C.COMPROBANTE).setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText('Ver comprobante').setLinkUrl(url).build()
  );
  if (C.ESTADO < 0) return;
  const celdaEstado = hoja.getRange(fila, C.ESTADO);
  const estadoActual = String(celdaEstado.getValue() || '').trim().toUpperCase();
  if (estadoActual !== EST_RECHAZADO && estadoActual !== EST_CANCELADO) {
    celdaEstado.setValue(EST_COMPROBANTE);
  }
}

/**
 * Copia el PDF final a la carpeta de AS con el mismo nombre (los IDs separados por coma).
 * Si en esa carpeta todavía hay uno con el mismo nombre (AS aún no lo procesó), se manda a la
 * papelera antes, igual que hace guardarPdfEnCarpeta_ en la otra carpeta. Si AS ya lo procesó,
 * no hay forma de saberlo desde acá: regenerar un comprobante lo vuelve a mandar a AS.
 * Si la copia falla, el comprobante igual queda armado y la falla se ve como nota en la celda
 * COMPROBANTE: hay que subirlo a mano.
 */
function copiarComprobanteAAS_(hoja, fila, url) {
  const C = colsAgrup_(hoja);
  if (C.COMPROBANTE < 0) return;
  const celda = hoja.getRange(fila, C.COMPROBANTE);
  try {
    const archivo = DriveApp.getFileById(extraerIdDeDrive_(url));
    const nombre = archivo.getName();
    const carpetaAS = DriveApp.getFolderById(COMPROB_CARPETA_AS_ID);
    const previos = carpetaAS.getFilesByName(nombre);
    while (previos.hasNext()) previos.next().setTrashed(true);
    archivo.makeCopy(nombre, carpetaAS);
    celda.clearNote();
  } catch (e) {
    celda.setNote('⚠️ No se copió a la carpeta de AS: ' + e.message + '\nSubilo a mano.');
  }
}

/**
 * PDF DE GALICIA, paso 1 de 2. Devuelve la orden de pago en base64 para que el panel la una en el
 * navegador con el comprobante de Galicia y las retenciones.
 */
function comprobanteOrdenPdf(nombreHoja, fila) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, motivo: 'No encuentro la hoja "' + nombreHoja + '".' };
    const op = operacionParaPdf_(hoja, fila);
    if (!op.ok) return op;
    return { ok: true, data: Utilities.base64Encode(blobOrdenDePago_(op.datos).getBytes()) };
  } catch (err) {
    return { ok: false, motivo: 'Error armando la orden de pago: ' + err.message };
  }
}

/**
 * PDF DE GALICIA, paso 2 de 2. Guarda el PDF ya unido en la carpeta de AppSheet, con el nombre de
 * los códigos, y marca la fila.
 * Antes de guardar confirma que la fila siga teniendo los mismos IDs con los que se armó la orden:
 * si regeneraste la agrupada con el panel abierto, las filas pueden haberse corrido, y el PDF
 * quedaría en la fila de otra operación.
 */
function comprobanteGuardarFinal(nombreHoja, fila, pdfBase64, idsEsperados) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, motivo: 'No encuentro la hoja "' + nombreHoja + '".' };
    const op = operacionParaPdf_(hoja, fila);
    if (!op.ok) return op;
    if (op.datos.ids !== String(idsEsperados)) {
      return { ok: false, motivo: 'La fila ' + fila + ' cambió mientras armabas (ahora tiene ' + op.datos.ids +
                                  '). ¿Regeneraste la agrupada? Cerrá el panel y volvé a abrirlo. No se guardó nada.' };
    }
    const nombre = nombreArchivoPdf_(op.datos.ids);
    const blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), MimeType.PDF, nombre + '.pdf');
    const url = guardarPdfEnCarpeta_(blob, nombre, DriveApp.getFolderById(COMPROB_CARPETA_ID));
    marcarComprobanteArmado_(hoja, fila, url);
    copiarComprobanteAAS_(hoja, fila, url);
    return { ok: true, url: url, id: extraerIdDeDrive_(url), nombre: nombre + '.pdf' };
  } catch (err) {
    return { ok: false, motivo: 'Error guardando el PDF: ' + err.message };
  }
}

/**
 * Inserta la captura del pago al final del documento, en su propia página.
 * La escala se ajusta primero al ancho útil y, si con eso la altura se pasa de la página, se
 * reescala por altura. Sin esto una captura larga se corta.
 */
function insertarComprobanteEnDoc_(body, blob) {
  body.appendPageBreak();

  const rotulo = body.appendParagraph('COMPROBANTE DE PAGO');
  rotulo.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
    [DocumentApp.Attribute.FONT_SIZE]: 11,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#0a0a8c'
  });

  const img = body.appendImage(blob);
  const w = img.getWidth();
  const h = img.getHeight();
  if (w > 0 && h > 0) {
    let escala = COMPROB_ANCHO_PAGINA / w;
    if (h * escala > COMPROB_ALTO_PAGINA) escala = COMPROB_ALTO_PAGINA / h;
    img.setWidth(Math.round(w * escala)).setHeight(Math.round(h * escala));
  }
}

/** Pasa una fila de la agrupada (leída con colsAgrup_) al formato que muestran los dos paneles. */
function filaOperacion_(fila, v, C) {
  const ids = String(valAgrup_(v, C, 'IDS') || '').trim();
  const monto = aNumero_(valAgrup_(v, C, 'MONTO'));
  return {
    fila: fila,
    tipo: String(valAgrup_(v, C, 'TIPO') || ''),
    beneficiario: String(valAgrup_(v, C, 'BENEF') || '(sin beneficiario)'),
    ids: ids,
    cantidad: ids.split(',').filter(String).length,
    monto: monto,
    moneda: String(valAgrup_(v, C, 'MONEDA') || '').trim().toUpperCase(),
    montoTexto: (monto === null ? '—' : formatearImporte_(monto)) + ' ' + String(valAgrup_(v, C, 'MONEDA') || ''),
    estado: String(valAgrup_(v, C, 'ESTADO') || '(sin estado)'),
    yaTiene: String(valAgrup_(v, C, 'COMPROBANTE') || '').trim() !== ''
  };
}

/* ══════════════════ PANEL DE A UNA FILA ══════════════════ */

function abrirPanelComprobante() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(panelComprobanteHtml_())
      .setTitle('Comprobante de pago')
      .setWidth(320)
  );
}

/** Datos de la fila activa, para que el panel te los muestre antes de armar nada. */
function comprobanteDatosFilaActiva() {
  const hoja = SpreadsheetApp.getActiveSheet();

  if (hoja.getName().indexOf(AGRUP_PREFIJO) !== 0) {
    return { ok: false, motivo: 'Parate sobre una hoja "' + AGRUP_PREFIJO + 'DD-MM-YYYY".\nEstás en "' + hoja.getName() + '".' };
  }

  const fila = hoja.getActiveRange().getRow();
  if (fila < 2) return { ok: false, motivo: 'Esa es la fila de encabezado.' };

  const C = colsAgrup_(hoja);
  const v = hoja.getRange(fila, 1, 1, C.ANCHO).getValues()[0];
  if (String(valAgrup_(v, C, 'IDS') || '').trim() === '') {
    return { ok: false, motivo: 'Esa fila no es una operación: es la fila de resumen o una banda de tipo.' };
  }

  const d = filaOperacion_(fila, v, C);
  d.ok = true;
  return d;
}

function comprobanteArmar(fila, archivo) {
  return comprobanteArmarEnHoja_(SpreadsheetApp.getActiveSheet().getName(), fila, archivo);
}

/* ══════════════════ PANEL DE LOTE ══════════════════ */

function abrirPanelComprobantesLote() {
  const hoja = SpreadsheetApp.getActiveSheet();
  if (hoja.getName().indexOf(AGRUP_PREFIJO) !== 0) {
    SpreadsheetApp.getUi().alert(
      'Parate sobre una hoja "' + AGRUP_PREFIJO + 'DD-MM-YYYY" antes de abrir el panel.\n' +
      'Estás en "' + hoja.getName() + '".'
    );
    return;
  }
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(panelLoteHtml_()).setWidth(1000).setHeight(640),
    'Armar comprobantes en lote'
  );
}

/** Todas las operaciones de la agrupada activa, para la lista del panel de lote. */
function comprobantesLoteOperaciones() {
  const hoja = SpreadsheetApp.getActiveSheet();
  if (hoja.getName().indexOf(AGRUP_PREFIJO) !== 0) {
    return { ok: false, motivo: 'La hoja activa no es una agrupada.' };
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { ok: false, motivo: 'La hoja está vacía.' };

  const C = colsAgrup_(hoja);
  if (C.IDS < 0) return { ok: false, motivo: 'La hoja no tiene la columna TRANSACCIONES (IDs).' };
  const valores = hoja.getRange(2, 1, ultimaFila - 1, C.ANCHO).getValues();
  const nombreDia = hoja.getName().replace(AGRUP_PREFIJO, '');
  const cbuPorTask = cbusDeHojaDelDia_(nombreDia);
  const pagoPorTask = datosDePagoDeHojaDelDia_(nombreDia);
  const ops = [];
  valores.forEach((v, i) => {
    if (String(valAgrup_(v, C, 'IDS') || '').trim() === '') return;   // banda o fila de resumen
    const op = filaOperacion_(i + 2, v, C);
    // CBU de cada tarea del grupo, para que el panel reconozca el comprobante de Galicia por su
    // cuenta de destino. Solo viaja al panel: la agrupada sigue sin mostrar el CBU.
    op.cbus = [];
    // Dato de pago de cada tarea (mail, chave PIX, cuenta), para reconocer la captura por OCR.
    op.pagos = [];
    op.ids.split(',').forEach(t => {
      const c = cbuPorTask[t.trim()];
      if (c && op.cbus.indexOf(c) === -1) op.cbus.push(c);
      (pagoPorTask[t.trim()] || []).forEach(p => { if (op.pagos.indexOf(p) === -1) op.pagos.push(p); });
    });
    ops.push(op);
  });

  return { ok: true, hoja: hoja.getName(), operaciones: ops,
           tiposGalicia: TIPOS_VALIDABLES, tipoRetencion: COMPROB_TIPO_CON_RETENCION };
}

/**
 * TASK → CBU (22 dígitos) desde la columna "CBU / DATO DE PAGO" de la hoja del día. Los alias y
 * los datos que no son un CBU de 22 dígitos se ignoran: para esas filas el panel compara por nombre.
 */
function cbusDeHojaDelDia_(nombreDia) {
  const res = {};
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreDia);
  if (!hoja || hoja.getLastRow() < 2) return res;
  const idx = indicesPorHeader_(hoja);
  const iTask = idx['TASK'];
  const iDato = idx['CBU / DATO DE PAGO'];
  if (iTask < 0 || iDato < 0) return res;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues().forEach(v => {
    const task = String(v[iTask] == null ? '' : v[iTask]).trim();
    const cbu = String(v[iDato] == null ? '' : v[iDato]).replace(/\D/g, '');
    if (task !== '' && cbu.length === 22) res[task] = cbu;
  });
  return res;
}

/**
 * TASK → [dato de pago] desde la hoja del día. Toma "CBU / DATO DE PAGO" y, si es un mail, también
 * el que figure en DATOS ADICIONALES (en PIX ahí viene el mail cuando la chave es un CPF/CNPJ).
 * Solo lo usa el panel de lote para reconocer capturas; no se escribe en ningún lado.
 */
function datosDePagoDeHojaDelDia_(nombreDia) {
  const res = {};
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreDia);
  if (!hoja || hoja.getLastRow() < 2) return res;
  const idx = indicesPorHeader_(hoja);
  const iTask = idx['TASK'];
  const iDato = idx['CBU / DATO DE PAGO'];
  const iAdic = idx['DATOS ADICIONALES'];
  if (iTask < 0 || iDato < 0) return res;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues().forEach(v => {
    const task = String(v[iTask] == null ? '' : v[iTask]).trim();
    if (task === '') return;
    const lista = [];
    const dato = String(v[iDato] == null ? '' : v[iDato]).trim();
    if (dato !== '') lista.push(dato);
    if (iAdic >= 0) {
      (String(v[iAdic] == null ? '' : v[iAdic]).match(/[\w.+-]+@[\w-]+(\.[\w-]+)+/g) || [])
        .forEach(m => { if (lista.indexOf(m) === -1) lista.push(m); });
    }
    res[task] = lista;
  });
  return res;
}

/** Una operación por llamada: así ninguna ejecución se acerca al límite de 6 minutos. */
function comprobantesLoteArmarUno(nombreHoja, fila, archivo) {
  return comprobanteArmarEnHoja_(nombreHoja, fila, archivo);
}

/**
 * Junta los PDF recién generados en un ZIP y devuelve un link de descarga directa.
 *
 * El ZIP se guarda en PDF_CARPETA_ID (la carpeta de las órdenes de pago), NO en la que mira
 * AppSheet: un .zip ahí adentro podría confundir a AppSheet, que espera comprobantes sueltos.
 * Tampoco va a la raíz de Drive: DriveApp.getRootFolder() puede pedir un permiso que el script
 * todavía no tiene, y una autorización nueva pedida desde una ventana modal falla en silencio.
 * Usando una carpeta por ID se reusa exactamente el mismo acceso que ya funciona.
 * Se pisa el del mismo nombre en cada corrida, así no se acumulan.
 *
 * Se devuelve la URL NORMAL de Drive, no uc?export=download. Esa segunda redirige a
 * drive.usercontent.google.com y falla con 403 en dos casos habituales: cuando hay varias cuentas
 * de Google abiertas y el link se abre con el authuser equivocado, y cuando Google intercepta la
 * descarga de un comprimido con la advertencia de "no se pudo analizar en busca de virus", que
 * necesita un token de confirmación que ese link no lleva.
 * La URL normal maneja las dos cosas desde la interfaz de Drive: es un clic más y no se rompe.
 */
function comprobantesLoteZip(ids, nombreDia) {
  try {
    if (!ids || !ids.length) return { ok: false, motivo: 'No hay archivos generados para bajar.' };

    const blobs = [];
    const fallados = [];
    ids.forEach(id => {
      try {
        blobs.push(DriveApp.getFileById(id).getBlob());
      } catch (e) {
        fallados.push(id);
      }
    });

    if (blobs.length === 0) {
      return { ok: false, motivo: 'No se pudo leer ninguno de los archivos desde Drive.' };
    }

    const nombre = 'comprobantes ' + String(nombreDia || '') + '.zip';
    const zip = Utilities.zip(blobs, nombre);

    const carpeta = DriveApp.getFolderById(PDF_CARPETA_ID);
    const previos = carpeta.getFilesByName(nombre);
    while (previos.hasNext()) previos.next().setTrashed(true);

    const archivo = carpeta.createFile(zip);

    return {
      ok: true,
      nombre: nombre,
      cantidad: blobs.length,
      fallados: fallados.length,
      url: archivo.getUrl(),
      carpeta: carpeta.getUrl()
    };

  } catch (err) {
    return { ok: false, motivo: 'Error armando el ZIP: ' + err.message };
  }
}

/**
 * DIAGNÓSTICO. Correr A MANO desde el editor de Apps Script (botón Run, eligiendo esta función).
 *
 * Sirve para dos cosas:
 *   1) Si el script necesita una autorización nueva, corriéndolo desde el editor Google te muestra
 *      el cartel de permisos. Desde una ventana modal no puede mostrarlo y la llamada muere en
 *      silencio, que es exactamente el síntoma del botón que queda en "Armando el ZIP…".
 *   2) Recorre la columna COMPROBANTE de la agrupada de hoy, arma el ZIP con lo que encuentre y
 *      deja el resultado en Ver > Registros de ejecución.
 */
function diagnosticoZipComprobantes() {
  const nombreDia = formatearNombreHoja_(new Date());
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AGRUP_PREFIJO + nombreDia);
  if (!hoja) { Logger.log('No existe la hoja ' + AGRUP_PREFIJO + nombreDia); return; }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) { Logger.log('La agrupada está vacía.'); return; }

  const C = colsAgrup_(hoja);
  if (C.COMPROBANTE < 0) { Logger.log('La agrupada no tiene la columna COMPROBANTE.'); return; }
  const enlaces = hoja.getRange(2, C.COMPROBANTE, ultimaFila - 1, 1).getRichTextValues();
  const ids = [];
  enlaces.forEach(r => {
    const url = r[0].getLinkUrl();
    if (!url) return;
    const id = extraerIdDeDrive_(url);
    Logger.log('link: ' + url + '   →   id: ' + id);
    if (id) ids.push(id);
  });

  Logger.log('IDs encontrados: ' + ids.length);
  if (ids.length === 0) { Logger.log('Nada para comprimir.'); return; }

  const r = comprobantesLoteZip(ids, nombreDia);
  Logger.log('Resultado: ' + JSON.stringify(r));
}

/* ══════════════════ HTML DEL PANEL DE A UNA ══════════════════ */

function panelComprobanteHtml_() {
  return `
<style>
  body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; margin: 0; padding: 12px; color: #222; }
  .caja { border: 1px solid #ddd; border-radius: 6px; padding: 10px; margin-bottom: 12px; background: #fafafa; }
  .caja .ben { font-weight: bold; font-size: 14px; margin-bottom: 4px; }
  .caja .monto { font-size: 16px; font-weight: bold; color: #0a0a8c; margin: 6px 0; }
  .caja .meta { color: #666; font-size: 11px; line-height: 1.5; }
  .aviso { background: #fff4d6; border-left: 3px solid #e69138; padding: 8px; font-size: 12px; margin-bottom: 12px; }
  .error { background: #fce8e6; border-left: 3px solid #cc0000; padding: 8px; font-size: 12px; white-space: pre-wrap; }
  .ok { background: #e6f4ea; border-left: 3px solid #137333; padding: 8px; font-size: 12px; }
  button { width: 100%; padding: 9px; font-size: 13px; border-radius: 4px; border: none; cursor: pointer; margin-bottom: 8px; }
  #go { background: #0a0a8c; color: #fff; font-weight: bold; }
  #go:disabled { background: #bbb; cursor: default; }
  #refrescar { background: #eee; color: #333; }
  input[type=file] { width: 100%; margin-bottom: 10px; font-size: 12px; }
</style>

<button id="refrescar">↻ Leer la fila donde estoy</button>
<div id="info"></div>
<input type="file" id="archivo" accept="image/png,image/jpeg">
<button id="go" disabled>Armar comprobante</button>
<div id="msg"></div>

<script>
  var filaActual = null;
  function $(id) { return document.getElementById(id); }

  function cargar() {
    $('info').innerHTML = '<div class="caja"><div class="meta">Leyendo…</div></div>';
    $('go').disabled = true;
    $('msg').innerHTML = '';
    google.script.run.withSuccessHandler(pintar)
      .withFailureHandler(function (e) { $('info').innerHTML = '<div class="error">' + e.message + '</div>'; })
      .comprobanteDatosFilaActiva();
  }

  function pintar(d) {
    if (!d.ok) {
      filaActual = null;
      $('info').innerHTML = '<div class="error">' + d.motivo + '</div>';
      $('go').disabled = true;
      return;
    }
    filaActual = d.fila;
    var aviso = d.yaTiene
      ? '<div class="aviso">Esta fila ya tiene un comprobante armado. Si seguís, se reemplaza.</div>'
      : '';
    $('info').innerHTML = aviso +
      '<div class="caja">' +
        '<div class="ben">' + d.beneficiario + '</div>' +
        '<div class="monto">' + d.montoTexto + '</div>' +
        '<div class="meta">' + d.tipo + '<br>' +
        d.cantidad + ' transacción(es): ' + d.ids + '<br>' +
        'Estado: ' + d.estado + '<br>Fila ' + d.fila + '</div>' +
      '</div>';
    $('go').disabled = !$('archivo').files.length;
  }

  $('refrescar').onclick = cargar;
  $('archivo').onchange = function () { $('go').disabled = !(filaActual && this.files.length); };

  $('go').onclick = function () {
    var f = $('archivo').files[0];
    if (!f || !filaActual) return;
    $('go').disabled = true;
    $('msg').innerHTML = '<div class="caja"><div class="meta">Armando el PDF…</div></div>';

    var lector = new FileReader();
    lector.onload = function () {
      google.script.run
        .withSuccessHandler(function (r) {
          if (r.ok) {
            $('msg').innerHTML = '<div class="ok">Listo: <b>' + r.nombre + '</b><br>' +
              '<a href="' + r.url + '" target="_blank">Abrir en Drive</a></div>';
            $('archivo').value = '';
          } else {
            $('msg').innerHTML = '<div class="error">' + r.motivo + '</div>';
            $('go').disabled = false;
          }
        })
        .withFailureHandler(function (e) {
          $('msg').innerHTML = '<div class="error">' + e.message + '</div>';
          $('go').disabled = false;
        })
        .comprobanteArmar(filaActual, { data: lector.result.split(',')[1], mimeType: f.type, nombre: f.name });
    };
    lector.readAsDataURL(f);
  };

  cargar();
</script>
`;
}

/* ══════════════════ HTML DEL PANEL DE LOTE ══════════════════ */

/**
 * pdf-lib y pdf.js se cargan desde cdnjs. pdf.worker.min.js va como <script> y no como Worker:
 * con el worker cargado en la página, pdf.js corre en el hilo principal y no intenta crear un Web
 * Worker desde otro dominio, que dentro de la ventana de Apps Script no está garantizado.
 * Para PDF de una o dos páginas el hilo principal sobra.
 *
 * Tesseract.js (OCR de las capturas) se carga desde jsdelivr, versión fija. Crea su propio Worker
 * a partir de un blob, así que no depende del dominio. La primera vez baja el motor y el idioma
 * inglés (unos MB; después queda en caché del navegador). Se usa inglés y no español a propósito:
 * lo que importa leer son mails, montos y códigos, y ese modelo es más liviano.
 */
function panelLoteHtml_() {
  return `
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js"></script>
<style>
  * { box-sizing: border-box; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; margin: 0; padding: 0; color: #222; }
  #barra { padding: 10px 14px; background: #0a0a8c; color: #fff; display: flex; align-items: center; gap: 14px; }
  #barra b { font-size: 15px; }
  #barra .cont { font-size: 12px; opacity: .9; }
  #barra .ocr { font-size: 11px; opacity: .85; margin-left: auto; }
  #cuerpo { display: flex; height: 480px; }
  #izq { width: 52%; padding: 12px; overflow: auto; border-right: 1px solid #ddd; background: #f7f7fa; }
  #der { width: 48%; padding: 12px; overflow: auto; }
  #preview { width: 100%; border: 1px solid #ccc; background: #fff; }
  #sinPreview { display: none; padding: 30px; text-align: center; color: #777; border: 1px dashed #bbb; background: #fff; }
  #nomArchivo { font-size: 12px; color: #555; margin: 8px 0; word-break: break-all; }
  #pie { padding: 10px 14px; border-top: 1px solid #ddd; display: flex; gap: 8px; align-items: center; background: #fafafa; }
  button { padding: 8px 14px; font-size: 13px; border-radius: 4px; border: 1px solid #ccc; background: #fff; cursor: pointer; }
  button.primario { background: #0a0a8c; color: #fff; border-color: #0a0a8c; font-weight: bold; }
  button:disabled { opacity: .45; cursor: default; }
  #aceptar { width: 100%; margin-bottom: 8px; background: #137333; color: #fff; border-color: #137333; font-weight: bold; text-align: left; }
  .op { border: 1px solid #ddd; border-radius: 5px; padding: 8px 10px; margin-bottom: 6px; cursor: pointer; background: #fff; }
  .op:hover { border-color: #0a0a8c; background: #f0f0ff; }
  .op .ben { font-weight: bold; }
  .op .monto { color: #0a0a8c; font-weight: bold; }
  .op .meta { color: #777; font-size: 11px; margin-top: 2px; }
  .op.usada { opacity: .4; cursor: default; background: #eee; }
  .op.usada:hover { border-color: #ddd; background: #eee; }
  .op.tiene { border-left: 4px solid #e69138; }
  .op.sug { border: 2px solid #137333; }
  #buscar { width: 100%; padding: 7px; margin-bottom: 8px; border: 1px solid #ccc; border-radius: 4px; }
  #errorCarga { display: none; }
  #resumen { padding: 14px; overflow: auto; height: 480px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th, td { border: 1px solid #ddd; padding: 5px 7px; text-align: left; vertical-align: top; }
  th { background: #eee; }
  .est-ok { color: #137333; font-weight: bold; }
  .est-err { color: #cc0000; font-weight: bold; }
  .est-pend { color: #888; }
  .meta { color: #777; font-size: 11px; }
  .rol { display: inline-block; font-size: 10px; padding: 1px 5px; border-radius: 3px; background: #e8e8f5; color: #0a0a8c; }
  .vacio { color: #777; padding: 20px; text-align: center; }
  .datosPdf { font-size: 12px; padding: 7px 9px; margin-bottom: 8px; background: #fff; border: 1px solid #ddd; border-radius: 4px; }
  .sinSug { color: #cc0000; }
  .rechazo { font-size: 12px; padding: 7px 9px; margin-bottom: 8px; background: #fce8e6; border-left: 3px solid #cc0000; }
  #avisoZip { margin: 0 14px 10px 14px; padding: 9px 11px; font-size: 12px; border-radius: 4px; }
  #avisoZip.info { background: #e6f4ea; border-left: 3px solid #137333; }
  #avisoZip.error { background: #fce8e6; border-left: 3px solid #cc0000; }
  #codigo { width: 100%; padding: 8px; margin-bottom: 6px; border: 1px solid #999; border-radius: 4px; font-size: 13px; }
  #msgCodigo { font-size: 12px; margin-bottom: 8px; padding: 6px 8px; border-radius: 4px; display: none; }
  #msgCodigo.info { display: block; background: #e6f4ea; border-left: 3px solid #137333; }
  #msgCodigo.error { display: block; background: #fce8e6; border-left: 3px solid #cc0000; }
</style>

<div id="barra">
  <b>Armar comprobantes en lote</b>
  <span class="cont" id="contador"></span>
  <span class="ocr" id="estadoOcr"></span>
</div>

<div id="cuerpo">
  <div id="izq">
    <div id="zonaInicio">
      <p>Elegí todos los archivos juntos: capturas de Wise/PIX/PayPal/Payoneer (PNG o JPG), comprobantes de Galicia y retenciones (PDF). Después los vas a ver uno por uno.</p>
      <p>Captura: el panel la lee y, si reconoce el código de tarea o el dato de pago + el monto, te sugiere la fila; confirmás con Enter. PDF de Galicia o retención: el panel lee el PDF y te sugiere la fila.</p>
      <input type="file" id="archivos" multiple accept="image/png,image/jpeg,application/pdf">
    </div>
    <div id="zonaImg" style="display:none">
      <div id="infoPdf"></div>
      <button id="aceptar" style="display:none"></button>
      <input type="text" id="codigo" placeholder="Alternativa: tipeá el código de la tarea y apretá Enter">
      <div id="msgCodigo"></div>
      <img id="preview">
      <div id="sinPreview">Sin vista previa</div>
      <div id="nomArchivo"></div>
    </div>
  </div>
  <div id="der">
    <input type="text" id="buscar" placeholder="Filtrar por beneficiario o código…">
    <div id="errorCarga" class="rechazo"></div>
    <div id="lista"><div class="vacio">Cargando operaciones de la agrupada…</div></div>
  </div>
</div>

<div id="resumen" style="display:none"></div>
<div id="avisoZip" style="display:none"></div>

<div id="pie">
  <button id="anterior">← Anterior</button>
  <button id="saltear">Saltear →</button>
  <span style="flex:1"></span>
  <button id="verResumen">Ver resumen y generar</button>
  <button id="generar" class="primario" style="display:none">Generar los PDF</button>
  <button id="zip" style="display:none">⬇ Descargar todos en un ZIP</button>
  <button id="volver" style="display:none">← Volver a asignar</button>
</div>

<script>
  var HOJA = null, OPS = [], ARCH = [], idx = 0, generando = false, RESULT = {};
  var TIPOS_GALICIA = [], TIPO_RETENCION = '';
  var OPS_CARGADAS = false;   // false hasta que el servidor respondió con la lista
  var PDF_OK = !!(window.PDFLib && window.pdfjsLib && window.pdfjsWorker);
  var OCR_OK = !!window.Tesseract;
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function fmt(n) { return n === null || n === undefined ? '—' : Number(n).toFixed(2).replace('.', ','); }

  // Misma función que usa la validación de facturas (Pagoslb.gs), inyectada tal cual: así el
  // panel y el servidor comparan nombres exactamente igual.
  ${normalizarNombre_.toString()}

  function mismoNombre(a, b) {
    var x = normalizarNombre_(a), y = normalizarNombre_(b);
    if (!x.size || x.size !== y.size) return false;
    var ok = true;
    x.forEach(function (p) { if (!y.has(p)) ok = false; });
    return ok;
  }
  function esTipoArg(o) { return TIPOS_GALICIA.indexOf(o.tipo) !== -1; }
  function esTipoA(o) { return o.tipo === TIPO_RETENCION; }
  /*
   * Monto de la fila como número, o null. NO se compara contra null: una fila sin monto puede
   * llegar del servidor con el monto undefined (sin la propiedad), y "undefined === null" es
   * false. Así se colaba una fila sin monto y hasta decía "Galicia es igual a la fila".
   */
  function montoDe(o) { return (typeof o.monto === 'number' && !isNaN(o.monto)) ? o.monto : null; }
  function opDeFila(fila) { return OPS.filter(function (o) { return o.fila === fila; })[0]; }
  function esPrincipal(a) { return a && (a.rol === 'captura' || a.rol === 'galicia'); }

  /*
   * Si la carga de operaciones falla, el motivo va a #errorCarga, que pintarLista no toca. Antes
   * iba a la misma caja que la lista y se borraba en cuanto elegías los archivos, dejando un
   * "Nada coincide con el filtro" que no decía qué había pasado.
   */
  function errorDeCarga(texto) {
    $('errorCarga').innerHTML = texto;
    $('errorCarga').style.display = 'block';
  }

  google.script.run.withSuccessHandler(function (r) {
    OPS_CARGADAS = true;
    if (!r || !r.ok) {
      errorDeCarga('No se pudieron cargar las operaciones: ' + esc(r ? r.motivo : 'respuesta vacía del servidor') +
                   '. Cerrá el panel, parate en la agrupada de hoy y volvé a abrirlo.');
      pintarLista();
      return;
    }
    HOJA = r.hoja; OPS = r.operaciones;
    // Con un Comprobantes.gs viejo en el servidor, las filas llegarían sin los datos nuevos.
    if (OPS.length && !OPS.every(function (o) { return Array.isArray(o.cbus) && Array.isArray(o.pagos); })) {
      OPS = [];
      errorDeCarga('El Apps Script está corriendo una versión vieja de Comprobantes.gs (la lista no trae ' +
        'los datos nuevos). Reemplazá el archivo entero, guardá y volvé a abrir el panel.');
      pintarLista();
      return;
    }
    TIPOS_GALICIA = r.tiposGalicia || []; TIPO_RETENCION = r.tipoRetencion || '';
    if (ARCH.length) mostrar(); else pintarLista();
    actualizarContador();
  }).withFailureHandler(function (e) {
    OPS_CARGADAS = true;
    errorDeCarga('Error del servidor al cargar las operaciones: ' + esc(e.message) +
                 '. Si en Apps Script → Ejecuciones no figura comprobantesLoteOperaciones, corré ' +
                 'cualquier función a mano desde el editor: puede haber un permiso pendiente.');
    pintarLista();
  }).comprobantesLoteOperaciones();

  /* ── Carga de archivos ─────────────────────────────────────────────────────────────── */

  $('archivos').onchange = function () {
    var files = Array.prototype.slice.call(this.files);
    if (!files.length) return;
    ARCH = new Array(files.length);
    RESULT = {};
    var leidos = 0;
    function listo() {
      if (++leidos === files.length) { idx = 0; mostrar(); leerCapturas(); }
    }
    files.forEach(function (f, i) {
      var lector = new FileReader();
      lector.onload = function () {
        var a = { nombre: f.name, tipo: f.type, dataUrl: lector.result, fila: null, verificado: false,
                  rol: 'captura', datos: null, preview: lector.result, sugerida: null, motivoSug: '', motivoRechazo: '',
                  ocr: null, ocrEstado: '' };
        ARCH[i] = a;
        if (f.type === 'application/pdf') {
          analizarPdf(a).then(listo, function (e) {
            a.rol = 'rechazado'; a.preview = null;
            a.motivoRechazo = 'No pude leer el PDF: ' + e.message;
            listo();
          });
        } else {
          a.ocrEstado = OCR_OK ? 'pendiente' : 'sin OCR';
          listo();
        }
      };
      lector.readAsDataURL(f);
    });
  };

  /* ── OCR de las capturas, en el navegador ──────────────────────────────────────────── */

  /*
   * Una sola instancia de Tesseract y de a una captura por vez, en segundo plano: mientras tanto
   * podés ir asignando. Cuando termina la captura que tenés en pantalla, la pantalla se actualiza.
   */
  function leerCapturas() {
    var cola = ARCH.filter(function (a) { return a && a.rol === 'captura' && a.ocrEstado === 'pendiente'; });
    if (!cola.length) return;
    if (!OCR_OK) {
      $('estadoOcr').textContent = 'OCR no disponible: seguí con código o clic';
      return;
    }
    var hechas = 0;
    $('estadoOcr').textContent = 'Cargando OCR…';
    Tesseract.createWorker('eng').then(function (worker) {
      function siguiente(n) {
        if (n >= cola.length) {
          $('estadoOcr').textContent = 'OCR: ' + hechas + ' captura(s) leídas';
          worker.terminate();
          return;
        }
        $('estadoOcr').textContent = 'OCR: leyendo ' + (n + 1) + ' de ' + cola.length + '…';
        var a = cola[n];
        worker.recognize(a.dataUrl).then(function (res) {
          a.ocr = res.data.text || '';
          a.ocrEstado = 'ok';
          hechas++;
        }, function (e) {
          a.ocrEstado = 'error';
          a.motivoSug = 'el OCR falló (' + (e && e.message ? e.message : e) + ')';
        }).then(function () {
          if (ARCH[idx] === a && !a.fila) mostrar();
          siguiente(n + 1);
        });
      }
      siguiente(0);
    }, function (e) {
      cola.forEach(function (a) { a.ocrEstado = 'error'; a.motivoSug = 'no se pudo cargar el OCR'; });
      $('estadoOcr').textContent = 'OCR no disponible: seguí con código o clic';
      if (ARCH[idx] && ARCH[idx].rol === 'captura') mostrar();
      msgCodigo('No se pudo cargar el OCR del navegador (' + esc(e && e.message ? e.message : e) +
                '). Asigná tipeando el código o con clic.', true);
    });
  }

  /** Número a partir de "1.234,56", "1,234.56", "60.00" o "900". null si no es número. */
  function numeroDe(s) {
    s = String(s).replace(/[^\\d.,]/g, '');
    if (!/\\d/.test(s)) return null;
    var ult = Math.max(s.lastIndexOf('.'), s.lastIndexOf(','));
    if (ult >= 0) {
      var dec = s.length - ult - 1;
      if (dec === 1 || dec === 2) s = s.substring(0, ult).replace(/[.,]/g, '') + '.' + s.substring(ult + 1);
      else s = s.replace(/[.,]/g, '');
    }
    var n = Number(s);
    return isNaN(n) ? null : n;
  }

  var MONEDAS = ['USD', 'BRL', 'EUR', 'ARS', 'MXN', 'COP', 'CLP', 'GBP', 'PEN', 'UYU'];
  var SIMBOLOS = { BRL: ['R$'], EUR: ['€'], USD: ['US$', 'U$S'] };

  /** Lo que sirve del texto OCR: mails, números largos, importes, monedas y palabras. */
  function datosDeTexto(texto) {
    var t = String(texto || '');
    var up = t.toUpperCase().replace(/[–—]/g, '-');
    var d = {
      up: up.replace(/\\s*-\\s*/g, '-'),
      minus: t.toLowerCase(),
      largos: {}, importes: [], monedas: {}, tokens: normalizarNombre_(t)
    };
    (t.match(/\\d[\\d.\\/-]*\\d/g) || []).forEach(function (x) {
      var dig = x.replace(/\\D/g, '');
      if (dig.length >= 6) d.largos[dig] = true;
    });
    (t.match(/\\d[\\d.,]*/g) || []).forEach(function (x) {
      var n = numeroDe(x);
      if (n !== null && d.importes.indexOf(n) === -1) d.importes.push(n);
    });
    MONEDAS.forEach(function (m) { if (new RegExp('(^|[^A-Z])' + m + '([^A-Z]|$)').test(up)) d.monedas[m] = true; });
    Object.keys(SIMBOLOS).forEach(function (m) {
      SIMBOLOS[m].forEach(function (s) { if (up.indexOf(s) !== -1) d.monedas[m] = true; });
    });
    return d;
  }

  function escRe(s) { return String(s).replace(/[.*+?^$(){}|[\\]\\\\]/g, '\\\\$&'); }

  /** true si algún código de tarea de la fila aparece entero en el texto (MER-LIN-10-29 no matchea MER-LIN-10-297). */
  function tieneCodigo(d, o) {
    return o.ids.split(',').some(function (id) {
      var c = normCodigo(id);
      return c !== '' && new RegExp('(^|[^A-Z0-9])' + escRe(c) + '(?![0-9])').test(d.up);
    });
  }

  /** Qué identifica a la fila en el texto: su mail, su chave/cuenta o el nombre completo. '' si nada. */
  function identidadEn(d, o) {
    var datos = (o.pagos || []).concat([o.beneficiario]);
    for (var i = 0; i < datos.length; i++) {
      var x = String(datos[i] || '').trim();
      if (x === '') continue;
      if (x.indexOf('@') !== -1) {
        var mail = x.toLowerCase().replace(/\\s+/g, '');
        if (new RegExp('(^|[^a-z0-9._%+-])' + escRe(mail) + '(?![a-z0-9])').test(d.minus)) return 'el mail';
        continue;
      }
      var dig = x.replace(/\\D/g, '');
      if (dig.length >= 6 && d.largos[dig]) return 'la cuenta / chave';
    }
    var tk = normalizarNombre_(o.beneficiario);
    if (tk.size >= 2) {
      var todos = true;
      tk.forEach(function (p) { if (!d.tokens.has(p)) todos = false; });
      if (todos) return 'el nombre';
    }
    return '';
  }

  function montoEn(d, o) {
    var m = montoDe(o);
    if (m === null) return false;
    return d.importes.some(function (n) { return Math.abs(n - m) <= 0.01; });
  }

  /** false solo si la captura muestra monedas y ninguna es la de la fila. */
  function monedaCompatible(d, o) {
    var vistas = Object.keys(d.monedas);
    if (!vistas.length || !o.moneda) return true;
    return !!d.monedas[o.moneda];
  }

  /*
   * Sugerencia para una captura. Solo es "verificada" si hay UNA fila que cumple, y el criterio es
   * código de tarea, o identidad (mail / cuenta / nombre) + monto. Con solo el monto no se sugiere:
   * dos proveedores pueden cobrar lo mismo el mismo día.
   */
  function sugerirCaptura(a) {
    a.sugerida = null;
    if (a.ocrEstado === 'pendiente') { a.motivoSug = 'leyendo la captura…'; return; }
    if (a.ocrEstado !== 'ok') { if (!a.motivoSug) a.motivoSug = 'sin OCR'; return; }
    var d = datosDeTexto(a.ocr);

    var porCodigo = OPS.filter(function (o) { return tieneCodigo(d, o); });
    if (porCodigo.length === 1) { a.sugerida = porCodigo[0].fila; a.motivoSug = 'el código de tarea está en la captura'; return; }
    if (porCodigo.length > 1) { a.motivoSug = 'aparecen códigos de ' + porCodigo.length + ' filas: elegí con clic'; return; }

    var fuertes = [], soloIdentidad = [];
    OPS.forEach(function (o) {
      var id = identidadEn(d, o);
      if (!id) return;
      if (montoEn(d, o) && monedaCompatible(d, o)) fuertes.push({ o: o, id: id });
      else soloIdentidad.push({ o: o, id: id });
    });
    if (fuertes.length === 1) {
      a.sugerida = fuertes[0].o.fila;
      a.motivoSug = 'coinciden ' + fuertes[0].id + ' y el monto';
      return;
    }
    if (fuertes.length > 1) {
      a.motivoSug = fuertes.length + ' filas coinciden en dato de pago y monto: elegí con clic';
    } else if (soloIdentidad.length) {
      a.motivoSug = 'coincide ' + soloIdentidad[0].id + ' con ' +
        soloIdentidad.map(function (x) { return x.o.beneficiario + ' (' + x.o.montoTexto + ')'; }).join(', ') +
        ', pero no el monto o la moneda: revisá y elegí con clic';
    } else {
      a.motivoSug = 'no encontré en la captura ni un código de tarea ni un dato de pago de la agrupada';
    }
  }

  function bytesDe(dataUrl) {
    var bin = atob(dataUrl.split(',')[1]);
    var u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }

  /*
   * Un PDF se abre dos veces al cargarlo: con pdf-lib, para confirmar desde ya que se va a poder
   * unir (un PDF protegido falla acá y no al final), y con pdf.js, para leer el texto y dibujar
   * la primera página. Todo en el navegador: no se usa el OCR de Drive ni su cuota.
   */
  function analizarPdf(a) {
    if (!PDF_OK) {
      a.rol = 'rechazado'; a.preview = null;
      a.motivoRechazo = 'No se cargaron las librerías de PDF (pdf-lib / pdf.js). Revisá la conexión y reabrí el panel.';
      return Promise.resolve();
    }
    var bytes = bytesDe(a.dataUrl);
    return PDFLib.PDFDocument.load(bytes).then(function () {
      return pdfjsLib.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
    }, function (e) {
      throw new Error('no se puede abrir para unirlo (' + e.message + ')');
    }).then(function (doc) {
      var pedidos = [];
      for (var p = 1; p <= doc.numPages; p++) {
        pedidos.push(doc.getPage(p).then(function (pg) { return pg.getTextContent(); }));
      }
      return Promise.all(pedidos).then(function (contenidos) {
        var texto = contenidos.map(function (c) {
          return c.items.map(function (it) { return it.str + (it.hasEOL ? '\\n' : ' '); }).join('');
        }).join('\\n');
        clasificar(a, texto);
        return dibujarPrimeraPagina(doc).then(function (png) { a.preview = png; },
                                              function () { a.preview = null; });
      });
    });
  }

  function dibujarPrimeraPagina(doc) {
    return doc.getPage(1).then(function (pg) {
      var vp = pg.getViewport({ scale: 900 / pg.getViewport({ scale: 1 }).width });
      var c = document.createElement('canvas');
      c.width = vp.width; c.height = vp.height;
      return pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise
        .then(function () { return c.toDataURL('image/png'); });
    });
  }

  function importeAR(s) { return Number(String(s).replace(/\\./g, '').replace(',', '.')); }

  /*
   * Total retenido del archivo. En "Retenciones Realizadas" los importes van entre paréntesis, uno
   * por régimen; el certificado los repite en "Total Retenido - <régimen>". Se toma el mayor de
   * los dos para no contar dos veces cuando el archivo trae la orden y el certificado juntos.
   * 0 significa que el archivo no retiene nada: es la orden de pago del sistema y nada más.
   */
  function totalRetenido(texto) {
    var bloque = seccion(texto, /Retenciones\\s+Realizadas/i, /Intereses\\s*\\/|Valores\\s+Entregados|Total\\s+a\\s+Pagar/i);
    var enParentesis = 0;
    (bloque.match(/\\(([\\d.]+,\\d{2})\\)/g) || []).forEach(function (x) {
      enParentesis += importeAR(x.replace(/[()]/g, ''));
    });
    var enCertificado = 0;
    (texto.match(/Total\\s+Retenido[^\\n]*?([\\d.]+,\\d{2})/gi) || []).forEach(function (l) {
      enCertificado += importeAR(l.match(/([\\d.]+,\\d{2})\\s*$/)[1]);
    });
    return Math.round(Math.max(enParentesis, enCertificado) * 100) / 100;
  }

  function seccion(t, desde, hasta) {
    var i = t.search(desde);
    if (i < 0) return '';
    var r = t.substring(i);
    var j = r.search(hasta);
    return j > 0 ? r.substring(0, j) : r;
  }

  /*
   * Galicia Office Banking: monto en "Datos del pago"; CUIT (11 dígitos), razón social y cuenta
   * de destino (22 dígitos) en "Datos del destinatario".
   * Retención: el proveedor es la línea que sigue a "Datos del Proveedor", y su CUIT el primero
   * con rótulo "CUIT:" (el de Eleven viene como "C.U.I.T." y no matchea).
   */
  function clasificar(a, texto) {
    if (/Office Banking/i.test(texto) && /Datos del destinatario/i.test(texto)) {
      a.rol = 'galicia';
      var pago = seccion(texto, /Datos del pago/i, /Datos del pagador/i);
      var dest = seccion(texto, /Datos del destinatario/i, /COMPROBANTE DE LA OPERACI/i);
      var m = pago.match(/\\$\\s*([\\d.]+,\\d{2})/);
      var cuit = dest.match(/\\b(\\d{11})\\b/);
      var cbu = dest.match(/\\b(\\d{22})\\b/);
      var nombre = null;
      if (cuit) {
        dest.split('\\n').forEach(function (l) {
          var k = l.indexOf(cuit[1]);
          if (nombre === null && k !== -1) nombre = l.substring(k + 11).trim() || null;
        });
      }
      a.datos = { monto: m ? importeAR(m[1]) : null, cuit: cuit ? cuit[1] : null,
                  cbu: cbu ? cbu[1] : null, nombre: nombre };
    } else if (/Certificado de Retenci/i.test(texto) ||
               (/Datos del Proveedor/i.test(texto) && /Comprobantes Pagados/i.test(texto))) {
      // Documentación contable de Eleven: la orden de pago del sistema, el certificado de
      // retención, o las dos cosas en un mismo archivo. Van después del comprobante de Galicia.
      a.rol = 'contable';
      var lineas = texto.split('\\n').map(function (l) { return l.trim(); });
      var i = -1;
      for (var k = 0; k < lineas.length; k++) if (/^Datos del Proveedor/i.test(lineas[k])) { i = k; break; }
      var c = texto.match(/CUIT:\\s*(\\d{2}-?\\d{8}-?\\d)/);
      var tp = texto.match(/Total\\s+a\\s+Pagar\\s*:?\\s*\\$?\\s*([\\d.]+,\\d{2})/i);
      a.datos = { nombre: (i >= 0 && lineas[i + 1]) ? lineas[i + 1] : null,
                  cuit: c ? c[1].replace(/\\D/g, '') : null,
                  totalAPagar: tp ? importeAR(tp[1]) : null,
                  retenido: totalRetenido(texto) };
      a.subtipo = a.datos.retenido ? 'con retención' : 'orden del sistema';
    } else {
      a.rol = 'rechazado';
      a.motivoRechazo = 'No lo reconozco: no es un comprobante de Galicia Office Banking ni documentación ' +
                        'contable de Eleven (orden de pago del sistema o certificado de retención).';
    }
  }

  /*
   * Fila sugerida para un PDF. Galicia: primero por CBU de destino contra el CBU de la hoja del
   * día; si no hay (en la hoja hay alias), por nombre. Retención: por nombre, solo filas tipo A.
   * Solo sugiere cuando hay UNA coincidencia; si hay varias o ninguna, se elige con clic.
   * Las capturas van por sugerirCaptura.
   */
  function sugerir(a) {
    if (a.rol === 'captura') { sugerirCaptura(a); return; }
    a.sugerida = null; a.motivoSug = '';
    if (!a.datos || !OPS.length) return;
    var cands;
    if (a.rol === 'galicia') {
      cands = OPS.filter(esTipoArg);
      if (a.datos.cbu) {
        var porCbu = cands.filter(function (o) { return (o.cbus || []).indexOf(a.datos.cbu) !== -1; });
        if (porCbu.length === 1) { a.sugerida = porCbu[0].fila; a.motivoSug = 'coincide el CBU'; return; }
        if (porCbu.length > 1) { a.motivoSug = 'el CBU está en ' + porCbu.length + ' filas: elegí con clic'; return; }
      }
    } else if (a.rol === 'contable') {
      cands = a.datos.retenido ? OPS.filter(esTipoA) : OPS.filter(esTipoArg);
    } else {
      return;
    }
    if (!a.datos.nombre) { a.motivoSug = 'no pude leer el nombre en el PDF'; return; }
    var porNombre = cands.filter(function (o) { return mismoNombre(o.beneficiario, a.datos.nombre); });
    if (porNombre.length === 1) { a.sugerida = porNombre[0].fila; a.motivoSug = 'coincide el nombre'; }
    else a.motivoSug = porNombre.length
      ? 'el nombre coincide con ' + porNombre.length + ' filas: elegí con clic'
      : 'ninguna fila coincide por CBU ni por nombre';
  }

  /* ── Pantalla de asignación ────────────────────────────────────────────────────────── */

  function infoPdfHtml(a) {
    if (a.rol === 'rechazado') return '<div class="rechazo">✖ ' + esc(a.motivoRechazo) + ' No se puede asignar.</div>';
    if (a.rol === 'captura') {
      return a.sugerida ? '' : '<div class="datosPdf"><b>Captura</b> · <span class="sinSug">Sin sugerencia: ' +
        esc(a.motivoSug) + '</span></div>';
    }
    var d = a.datos || {};
    var txt = a.rol === 'galicia'
      ? '<b>Comprobante Galicia</b> · ' + esc(d.nombre || '¿nombre?') + ' · CUIT ' + esc(d.cuit || '?') +
        ' · CBU ' + esc(d.cbu || '?') + ' · <b>$ ' + fmt(d.monto) + '</b>'
      : '<b>' + esc(a.subtipo === 'con retención' ? 'Orden del sistema con retención' : 'Orden de pago del sistema') +
        '</b> · ' + esc(d.nombre || '¿nombre?') + ' · CUIT ' + esc(d.cuit || '?') +
        (d.totalAPagar !== null ? ' · Total a Pagar <b>$ ' + fmt(d.totalAPagar) + '</b>' : '') +
        (d.retenido ? ' · retenido $ ' + fmt(d.retenido) : '');
    if (!a.sugerida) txt += '<br><span class="sinSug">Sin sugerencia: ' + esc(a.motivoSug) + '</span>';
    return '<div class="datosPdf">' + txt + '</div>';
  }

  function mostrar() {
    if (!ARCH.length) return;
    $('zonaInicio').style.display = 'none';
    $('zonaImg').style.display = 'block';
    var a = ARCH[idx];
    sugerir(a);
    if (a.preview) { $('preview').src = a.preview; $('preview').style.display = ''; $('sinPreview').style.display = 'none'; }
    else { $('preview').style.display = 'none'; $('sinPreview').style.display = 'block'; }
    $('codigo').style.display = a.rol === 'captura' ? '' : 'none';
    $('infoPdf').innerHTML = infoPdfHtml(a);
    var o = a.sugerida ? opDeFila(a.sugerida) : null;
    var bloqueoSug = o ? motivoBloqueo(a, o) : '';
    if (bloqueoSug) {
      // La fila es la correcta pero no se puede usar: se dice cuál es y qué falta, sin botón.
      $('infoPdf').innerHTML += '<div class="rechazo">La fila que corresponde es <b>' + esc(o.beneficiario) +
        '</b> (' + esc(o.ids) + ', ' + esc(a.motivoSug) + '), pero no se puede asignar: ' + esc(bloqueoSug) + '.' +
        (montoDe(o) === null ? '<br>Corré "Validar beneficiario y monto" sobre ' + esc(o.ids) + ' en la hoja del día, ' +
          'después "Actualizar hoja agrupada del día", y volvé a abrir este panel.' : '') + '</div>';
      o = null;
    }
    $('aceptar').style.display = o ? '' : 'none';
    if (o) $('aceptar').textContent = '✓ Enter: asignar a ' + o.beneficiario + ' — ' + o.montoTexto + '  (' + a.motivoSug + ')' +
      (o.yaTiene ? '  · YA TIENE COMPROBANTE, se reemplaza' : '');
    $('nomArchivo').textContent = a.nombre + (a.fila
      ? '   →   asignado a la fila ' + a.fila + (a.verificado ? '  (✓ verificado)' : '  (⚠ por clic, sin verificar)')
      : '');
    actualizarContador();
    pintarLista();
    // Con sugerencia, el foco va al botón verde: Enter confirma. Sin sugerencia, al código.
    if (o) $('aceptar').focus(); else if (a.rol === 'captura') $('codigo').focus();
  }

  function actualizarContador() {
    var asignados = ARCH.filter(function (a) { return a && a.fila; }).length;
    $('contador').textContent = ARCH.length
      ? ('Archivo ' + (idx + 1) + ' de ' + ARCH.length + '   ·   ' + asignados + ' asignados   ·   ' + (ARCH.length - asignados) + ' sin asignar')
      : (OPS.length + ' operaciones en la hoja');
  }

  /** Índice del archivo principal (captura o Galicia) asignado a una fila, o -1. */
  function principalDe(fila) {
    for (var i = 0; i < ARCH.length; i++) if (esPrincipal(ARCH[i]) && ARCH[i].fila === fila) return i;
    return -1;
  }

  /** Por qué el archivo actual NO puede ir a esa fila. Vacío si puede. */
  function motivoBloqueo(a, o) {
    if (!a || !o) return '';
    if (a.rol === 'rechazado') return 'archivo rechazado';
    // Misma regla que el servidor (operacionParaPdf_): sin monto verificado no se arma nada. Se
    // corta acá para que te enteres al asignar y no al generar.
    if (montoDe(o) === null) return 'sin monto verificado: falta leer la factura';
    if (a.rol === 'contable' && a.datos.retenido && !esTipoA(o)) return 'solo el tipo A lleva retenciones';
    if (a.rol === 'contable' && !esTipoArg(o)) return 'la documentación contable solo va en A o C';
    if (a.rol === 'galicia' && !esTipoArg(o)) return 'un comprobante de Galicia solo va en A o C';
    if (esPrincipal(a)) {
      var p = principalDe(o.fila);
      if (p >= 0 && p !== idx) return 'ya tiene ' + ARCH[p].nombre;
    }
    return '';
  }

  function pintarLista() {
    if (!OPS.length) {
      // Sin operaciones no hay nada que filtrar: se dice por qué está vacía.
      $('lista').innerHTML = '<div class="vacio">' + (OPS_CARGADAS
        ? 'No hay operaciones para mostrar.'
        : 'Cargando operaciones de la agrupada…') + '</div>';
      return;
    }
    var q = $('buscar').value.toLowerCase();
    var a = ARCH.length ? ARCH[idx] : null;
    var html = '';
    OPS.forEach(function (o) {
      if (q && (o.beneficiario + ' ' + o.ids).toLowerCase().indexOf(q) === -1) return;
      var bloqueo = motivoBloqueo(a, o);
      var sug = a && a.sugerida === o.fila;
      html += '<div class="op' + (bloqueo ? ' usada' : '') + (o.yaTiene ? ' tiene' : '') + (sug ? ' sug' : '') +
              '" data-fila="' + o.fila + '" data-bloqueo="' + (bloqueo ? '1' : '') + '">' +
        '<div><span class="ben">' + esc(o.beneficiario) + '</span> — <span class="monto">' + esc(o.montoTexto) + '</span></div>' +
        '<div class="meta">' + esc(o.tipo) + ' · ' + esc(o.ids) +
        // Sin CBU en la hoja del día, un PDF de Galicia solo puede cruzarse por nombre, y Galicia
        // recorta la razón social ("Giuliano Peccetto" contra "PECCETTO GIULIANO TOMAS"). Se avisa
        // para que el clic sobre esa fila sea una decisión y no una sorpresa.
        (esTipoArg(o) && !(o.cbus || []).length ? ' · sin CBU en la hoja del día' : '') +
        (o.yaTiene ? ' · <b>ya tiene comprobante</b>' : '') +
        (bloqueo ? ' · ' + esc(bloqueo) : '') + '</div>' +
      '</div>';
    });
    $('lista').innerHTML = html || '<div class="vacio">Nada coincide con el filtro.</div>';

    Array.prototype.forEach.call($('lista').children, function (el) {
      if (el.dataset.bloqueo === '1') return;
      el.onclick = function () {
        if (!ARCH.length) return;
        asignar(parseInt(el.dataset.fila, 10), false);
      };
    });
  }

  function msgCodigo(html, esError) {
    $('msgCodigo').className = esError ? 'error' : 'info';
    $('msgCodigo').innerHTML = html;
  }

  function asignar(fila, verificado) {
    var a = ARCH[idx];
    var op = opDeFila(fila);
    var bloqueo = motivoBloqueo(a, op);
    if (bloqueo) { msgCodigo('No se asignó: ' + esc(bloqueo) + '.', true); return; }
    a.fila = fila;
    a.verificado = verificado;
    $('codigo').value = '';
    $('buscar').value = '';   // el filtro no se arrastra al archivo siguiente
    if (idx < ARCH.length - 1) idx++;
    mostrar();
    msgCodigo((verificado ? '✓ ' : '⚠ SIN VERIFICAR: ') + esc(a.nombre) + ' → ' +
              esc(op.beneficiario) + ' · ' + esc(op.montoTexto) + ' · ' + esc(op.ids), !verificado);
  }

  $('aceptar').onclick = function () {
    var a = ARCH[idx];
    if (a && a.sugerida) asignar(a.sugerida, true);
  };

  /*
   * VERIFICACIÓN POR CÓDIGO TIPEADO (alternativa al OCR)
   * Solo cuenta como verificado si coincide EXACTO con uno de los IDs de una única fila. El clic
   * sigue funcionando para capturas sin código, pero queda SIN VERIFICAR.
   */
  function normCodigo(s) {
    return String(s || '').toUpperCase().trim().replace(/[\\s_]+/g, '-').replace(/-+/g, '-');
  }

  function filasConCodigo(txt) {
    var n = normCodigo(txt);
    return OPS.filter(function (o) {
      return o.ids.split(',').some(function (id) { return normCodigo(id) === n; });
    });
  }

  $('codigo').onkeydown = function (ev) {
    if (ev.key !== 'Enter') return;
    if (!ARCH.length) return;
    var txt = this.value;
    if (!normCodigo(txt)) return;

    var hits = filasConCodigo(txt);
    if (!hits.length) {
      msgCodigo('"' + esc(normCodigo(txt)) + '" no está en ' + esc(HOJA) + '. O está mal tipeado, ' +
                'o esa operación no está en la agrupada de hoy. No se asignó nada.', true);
      return;
    }
    if (hits.length > 1) {
      $('buscar').value = normCodigo(txt);
      pintarLista();
      msgCodigo('Ese código aparece en ' + hits.length + ' filas. Elegí con clic: va a quedar SIN VERIFICAR.', true);
      return;
    }
    var usadaPor = principalDe(hits[0].fila);
    if (usadaPor >= 0 && usadaPor !== idx) {
      msgCodigo('Esa fila ya está asignada a ' + esc(ARCH[usadaPor].nombre) + '. Dos archivos para el ' +
                'mismo código: revisá cuál corresponde. No se asignó nada.', true);
      return;
    }
    asignar(hits[0].fila, true);
  };

  $('buscar').oninput = pintarLista;
  $('anterior').onclick = function () { if (idx > 0) { idx--; $('msgCodigo').className = ''; mostrar(); } };
  $('saltear').onclick = function () { if (idx < ARCH.length - 1) { idx++; $('msgCodigo').className = ''; mostrar(); } };

  /* ── Resumen: una fila por operación ───────────────────────────────────────────────── */

  /*
   * Arma una entrada por fila de la agrupada con su archivo principal y su documentación, y la
   * evalúa. "bloqueo" impide generar esa fila. "obs" son observaciones: no impiden, pero obligan
   * a apretar Generar dos veces.
   */
  function armarGrupos() {
    var G = {}, sueltos = [];
    ARCH.forEach(function (a, i) {
      if (!a) return;
      if (!a.fila) { sueltos.push(i); return; }
      var g = G[a.fila] || (G[a.fila] = { fila: a.fila, principal: -1, contables: [] });
      if (a.rol === 'contable') g.contables.push(i); else g.principal = i;
    });
    var grupos = Object.keys(G).map(function (k) { return G[k]; })
      .sort(function (x, y) { return x.fila - y.fila; });
    grupos.forEach(evaluar);
    return { grupos: grupos, sueltos: sueltos };
  }

  function evaluar(g) {
    var o = opDeFila(g.fila);
    g.op = o; g.bloqueo = ''; g.obs = []; g.info = '';
    if (g.principal < 0) {
      g.bloqueo = 'falta el comprobante de pago (Galicia o captura)'; return;
    }
    var p = ARCH[g.principal];
    if (g.contables.length && p.rol !== 'galicia') {
      g.bloqueo = 'la documentación contable solo se une a un comprobante de Galicia, no a una captura'; return;
    }
    [g.principal].concat(g.contables).forEach(function (i) {
      if (!ARCH[i].verificado) g.obs.push(ARCH[i].nombre + ': asignado por clic, sin verificar');
    });
    if (p.rol !== 'galicia') return;

    var mg = p.datos ? p.datos.monto : null;      // lo que salió del banco
    var mf = montoDe(o);                          // lo que dice la fila (para A y C, la factura)
    var retenido = 0, totalAPagar = null;
    g.contables.forEach(function (i) {
      var d = ARCH[i].datos;
      retenido += d.retenido || 0;
      if (d.totalAPagar !== null && totalAPagar === null) totalAPagar = d.totalAPagar;
      if (d.cuit && p.datos.cuit && d.cuit !== p.datos.cuit) {
        g.obs.push(ARCH[i].nombre + ': CUIT ' + d.cuit + ' distinto del destinatario de Galicia ' + p.datos.cuit);
      }
    });
    retenido = Math.round(retenido * 100) / 100;

    if (mg === null || mf === null) {
      g.obs.push('no pude comparar el monto de Galicia con el de la fila');
      return;
    }

    // El "Total a Pagar" de la orden del sistema es exactamente lo que tiene que haber salido del
    // banco. Cuando está, es el chequeo más fuerte que hay: los tres números tienen que cerrar.
    if (totalAPagar !== null && Math.abs(totalAPagar - mg) > 0.01) {
      g.obs.push('Total a Pagar $ ' + fmt(totalAPagar) + ' ≠ Galicia $ ' + fmt(mg) +
                 ': el comprobante y la orden del sistema no son del mismo pago');
    }

    var dif = Math.round((mf - mg) * 100) / 100;
    if (dif < -0.01) {
      g.obs.push('Galicia $ ' + fmt(mg) + ' es MAYOR que la fila ' + fmt(mf));
    } else if (dif > 0.01) {
      if (!esTipoA(o)) {
        g.obs.push('Galicia $ ' + fmt(mg) + ' ≠ fila ' + fmt(mf) + ' (tipo C no lleva retención)');
      } else if (!retenido) {
        g.obs.push('Galicia es $ ' + fmt(dif) + ' menor que la fila y no adjuntaste ninguna retención');
      } else if (Math.abs(dif - retenido) > 0.01) {
        g.obs.push('fila − Galicia = ' + fmt(dif) + ', pero lo retenido es ' + fmt(retenido) +
                   ': faltan retenciones o alguna no corresponde');
      } else {
        g.info = 'fila ' + fmt(mf) + ' − Galicia ' + fmt(mg) + ' = ' + fmt(dif) + ', igual a lo retenido';
      }
    } else if (retenido) {
      g.obs.push('Galicia es igual a la fila, pero la documentación retiene ' + fmt(retenido));
    }

    var conRet = g.contables.filter(function (i) { return ARCH[i].datos.retenido; }).length;
    if (conRet > o.cantidad) g.obs.push(conRet + ' retenciones para ' + o.cantidad + ' ID(s)');
  }

  var ULTIMO = null;   // último resumen armado: lo usa Generar

  $('verResumen').onclick = function () {
    ULTIMO = armarGrupos();
    var filas = '';
    ULTIMO.grupos.forEach(function (g) {
      var res = RESULT[g.fila];
      var archivos = [g.principal].concat(g.contables).filter(function (i) { return i >= 0; })
        .map(function (i) { return '<span class="rol">' + (ARCH[i].rol === 'contable' ? ARCH[i].subtipo : ARCH[i].rol) + '</span> ' + esc(ARCH[i].nombre); }).join('<br>');
      var chequeos = g.bloqueo
        ? '<span class="est-err">✖ ' + esc(g.bloqueo) + '</span>'
        : (g.obs.length ? g.obs.map(function (x) { return '<span class="est-err">⚠ ' + esc(x) + '</span>'; }).join('<br>')
                        : '<span class="est-ok">✓ sin observaciones</span>') +
          (g.info ? '<br><span class="meta">' + esc(g.info) + '</span>' : '');
      var estado = res ? (res.estado === 'ok'
          ? '<a href="' + res.url + '" target="_blank">' + esc(res.detalle) + '</a>'
          : esc(res.detalle)) : (g.bloqueo ? 'no se genera' : 'pendiente');
      var clase = res ? (res.estado === 'ok' ? 'est-ok' : 'est-err') : 'est-pend';
      filas += '<tr><td><b>' + esc(g.op.beneficiario) + '</b><br>' + esc(g.op.montoTexto) +
        '<br><span class="meta">' + esc(g.op.ids) + '</span></td>' +
        '<td>' + archivos + '</td><td>' + chequeos + '</td>' +
        '<td class="' + clase + '" id="est' + g.fila + '">' + estado + '</td></tr>';
    });
    var sueltos = ULTIMO.sueltos.map(function (i) {
      var a = ARCH[i];
      return '<li>' + esc(a.nombre) + (a.rol === 'rechazado' ? ' — ' + esc(a.motivoRechazo) : ' — sin asignar') + '</li>';
    }).join('');
    $('resumen').innerHTML =
      '<table><tr><th>Operación</th><th>Archivos (en el orden en que se unen)</th><th>Chequeos</th><th>Estado</th></tr>' +
      filas + '</table>' +
      (sueltos ? '<p><b>Archivos que no van a ningún PDF:</b></p><ul>' + sueltos + '</ul>' : '');
    $('cuerpo').style.display = 'none';
    $('resumen').style.display = 'block';
    $('verResumen').style.display = 'none';
    $('anterior').style.display = 'none';
    $('saltear').style.display = 'none';
    $('generar').style.display = '';
    $('volver').style.display = '';
    $('generar').disabled = !pendientesDeGenerar().length;
  };

  function pendientesDeGenerar() {
    if (!ULTIMO) return [];
    return ULTIMO.grupos.filter(function (g) {
      return !g.bloqueo && !(RESULT[g.fila] && RESULT[g.fila].estado === 'ok');
    });
  }

  $('volver').onclick = function () {
    if (generando) return;
    if ($('generar').dataset.forzar) {
      delete $('generar').dataset.forzar;
      $('generar').textContent = 'Generar los PDF';
      $('avisoZip').style.display = 'none';
    }
    $('cuerpo').style.display = 'flex';
    $('resumen').style.display = 'none';
    $('verResumen').style.display = '';
    $('anterior').style.display = '';
    $('saltear').style.display = '';
    $('generar').style.display = 'none';
    $('volver').style.display = 'none';
  };

  $('generar').onclick = function () {
    var cola = pendientesDeGenerar();
    if (!cola.length) return;

    // Primer clic avisa si hay observaciones; el segundo genera.
    var conObs = cola.filter(function (g) { return g.obs.length; }).length;
    if (conObs && !$('generar').dataset.forzar) {
      $('generar').dataset.forzar = '1';
      $('generar').textContent = 'Generar igual (' + conObs + ' con observaciones)';
      avisoZip('<b>' + conObs + ' operación(es) con observaciones</b> (en rojo en la columna Chequeos). ' +
               'Revisalas contra los archivos. Si están bien, apretá de nuevo el botón.', true);
      return;
    }
    $('avisoZip').style.display = 'none';

    generando = true;
    $('generar').disabled = true;
    $('volver').disabled = true;
    procesar(cola, 0);
  };

  /** Orden de pago (la genera el servidor) → comprobante de Galicia → documentación contable. */
  function unirPdfs(ordenB64, archivos) {
    var out;
    return PDFLib.PDFDocument.create().then(function (d) {
      out = d;
      var fuentes = [ordenB64].concat(archivos.map(function (a) { return bytesDe(a.dataUrl); }));
      return fuentes.reduce(function (prom, src) {
        return prom
          .then(function () { return PDFLib.PDFDocument.load(src); })
          .then(function (doc) { return out.copyPages(doc, doc.getPageIndices()); })
          .then(function (pags) { pags.forEach(function (pg) { out.addPage(pg); }); });
      }, Promise.resolve());
    }).then(function () { return out.saveAsBase64(); });
  }

  // Una operación a la vez: cada llamada al servidor dura segundos y nunca toca el límite de 6
  // minutos. Lo que ya se generó queda hecho aunque se corte a la mitad.
  function procesar(cola, n) {
    if (n >= cola.length) {
      generando = false;
      $('volver').disabled = false;
      $('generar').textContent = 'Listo';
      if (Object.keys(RESULT).some(function (k) { return RESULT[k].id; })) $('zip').style.display = '';
      return;
    }
    var g = cola[n];
    var celda = $('est' + g.fila);
    celda.textContent = 'generando…';
    celda.className = 'est-pend';

    function fin(r) {
      if (r.ok) {
        RESULT[g.fila] = { estado: 'ok', detalle: r.nombre, id: r.id, url: r.url };
        celda.className = 'est-ok';
        celda.innerHTML = '<a href="' + r.url + '" target="_blank">' + esc(r.nombre) + '</a>';
      } else {
        RESULT[g.fila] = { estado: 'error', detalle: r.motivo };
        celda.className = 'est-err';
        celda.textContent = r.motivo;
      }
      procesar(cola, n + 1);
    }
    function falla(e) { fin({ ok: false, motivo: e.message }); }

    var p = ARCH[g.principal];
    if (p.rol === 'captura') {
      google.script.run.withSuccessHandler(fin).withFailureHandler(falla)
        .comprobantesLoteArmarUno(HOJA, g.fila, { data: p.dataUrl.split(',')[1], mimeType: p.tipo, nombre: p.nombre });
      return;
    }
    google.script.run
      .withSuccessHandler(function (r) {
        if (!r.ok) { fin(r); return; }
        var archivos = [p].concat(g.contables.map(function (i) { return ARCH[i]; }));
        unirPdfs(r.data, archivos).then(function (b64) {
          google.script.run.withSuccessHandler(fin).withFailureHandler(falla)
            .comprobanteGuardarFinal(HOJA, g.fila, b64, g.op.ids);
        }, function (e) { fin({ ok: false, motivo: 'Error uniendo los PDF: ' + e.message }); });
      })
      .withFailureHandler(falla)
      .comprobanteOrdenPdf(HOJA, g.fila);
  }

  $('zip').onclick = function () {
    var ids = Object.keys(RESULT).map(function (k) { return RESULT[k].id; }).filter(Boolean);
    if (!ids.length) { avisoZip('No hay ningún PDF generado para comprimir.', true); return; }

    $('zip').disabled = true;
    $('zip').textContent = 'Armando el ZIP…';
    avisoZip('Comprimiendo ' + ids.length + ' archivo(s)…', false);

    // Si en 90 segundos no volvió ni por éxito ni por error, la llamada se perdió. El caso típico
    // es que el script necesite una autorización nueva: Google no puede pedirla desde una ventana
    // modal y la llamada muere sin avisar. Sin este timeout el botón queda colgado para siempre.
    var respondio = false;
    var reloj = setTimeout(function () {
      if (respondio) return;
      $('zip').disabled = false;
      $('zip').textContent = '⬇ Descargar todos en un ZIP';
      avisoZip('El servidor no respondió en 90 segundos. Abrí Apps Script → Ejecuciones y fijate ' +
               'si figura comprobantesLoteZip. Si no figura, corré a mano la función ' +
               'diagnosticoZipComprobantes desde el editor: eso dispara el pedido de permisos que ' +
               'una ventana modal no puede mostrar.', true);
    }, 90000);

    google.script.run
      .withSuccessHandler(function (r) {
        respondio = true; clearTimeout(reloj);
        $('zip').disabled = false;
        if (!r.ok) { $('zip').textContent = '⬇ Descargar todos en un ZIP'; avisoZip(r.motivo, true); return; }
        $('zip').textContent = '⬇ ' + r.nombre;
        avisoZip('<b>' + esc(r.nombre) + '</b> — ' + r.cantidad + ' archivo(s)' +
                 (r.fallados ? ', ' + r.fallados + ' no se pudieron leer' : '') + '.<br>' +
                 'Se abre en Drive: apretá <b>Descargar</b> arriba a la derecha. ' +
                 'Un .zip no tiene vista previa, así que vas a ver la pantalla en gris; es normal.' +
                 '<br><a href="' + r.url + '" target="_blank">Abrir el ZIP</a> · ' +
                 '<a href="' + r.carpeta + '" target="_blank">Abrir la carpeta</a>', false);
        // URL normal de Drive: la de descarga directa da 403 con varias cuentas abiertas y con la
        // advertencia de virus que Google pone a los comprimidos.
        window.open(r.url, '_blank');
      })
      .withFailureHandler(function (e) {
        respondio = true; clearTimeout(reloj);
        $('zip').disabled = false;
        $('zip').textContent = '⬇ Descargar todos en un ZIP';
        avisoZip('Error del servidor: ' + e.message, true);
      })
      .comprobantesLoteZip(ids, HOJA.replace('AGRUP ', ''));
  };

  function avisoZip(html, esError) {
    var d = $('avisoZip');
    d.style.display = 'block';
    d.className = esError ? 'error' : 'info';
    d.innerHTML = html;
  }
</script>
`;
}
