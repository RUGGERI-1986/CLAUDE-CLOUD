/**
 * Genera/actualiza una hoja con el nombre del día de la fecha (DD-MM-YYYY),
 * separando y filtrando los pagos de la hoja "base" por tipo de transacción.
 *
 * Filtro: status_name = "Solicitado" o "Solicitado - Agrupación"
 * Columnas de salida (15):
 * TIPO | TASK | CBU / DATO DE PAGO | DATOS ADICIONALES | BENEFICIARIO | MONTO | MONTO FACTURA |
 * MONEDA | CHECK | COMENTARIO | COMENTARIO ANTERIOR | SUBGRUPO | INVOICE/COMPROBANTE | ESTADO | NUEVO
 *
 * Las posiciones de columna NO se escriben más como números sueltos: están todas en el objeto COL.
 * Si mañana se agrega otra columna, se cambia HEADERS y COL y nada más.
 *
 * Clave única de fila para preservar CHECK/COMENTARIO/SUBGRUPO entre actualizaciones: TASK + MONTO.
 *
 * ── QUÉ HAY DE NUEVO ────────────────────────────────────────────────────────────────────────
 * 1) COMENTARIO ANTERIOR (col J): al generar o actualizar, se recorren hasta las últimas
 *    MAX_HOJAS_ANTERIORES hojas "AGRUP DD-MM-YYYY" anteriores a hoy (de la más nueva a la más
 *    vieja) y se trae el comentario más reciente que tenga esa tarea, prefijado con la fecha de la
 *    hoja de la que salió: "01-09 | falta comprobante". Como una fila agrupada cubre varias
 *    tareas, su comentario se reparte a todas las que figuran en su columna de IDs.
 *    OJO: el comentario de la columna I de la hoja del día NO se arrastra a ningún lado. El
 *    arrastre entre días sale exclusivamente de la hoja agrupada.
 *    La marca "PAYPAL: ..." (PaypalAprobaciones.gs) no se arrastra: se saca antes.
 *
 * 2) YA PAGADO (se escribe en ESTADO): la evidencia de pago es ESTADO = "TRANSFERIDO" en una hoja
 *    agrupada anterior. Ese es el ÚNICO valor que cuenta como pago hecho. El CHECK NO cuenta:
 *    significa "revisé esta fila", y hay filas tildadas con problemas de método de pago, montos
 *    que no coinciden o falta de aprobación, que no se transfirieron.
 *    Las filas ya transferidas quedan marcadas "⚠️ PAGADO EL dd-MM - FALTA COMPROBANTE" con fondo
 *    naranja en la celda ESTADO. La columna ESTADO es filtrable, así que se pueden aislar todas
 *    esas filas de una.
 *    La marca se hereda de la hoja del día anterior, así que sobrevive a LimpiezaHojas: cuando se
 *    borra la hoja agrupada donde pusiste TRANSFERIDO, la evidencia sigue viva en la cadena.
 *    Excepción: si la tarea se canceló (ESTADO = CANCELADO en una agrupada anterior, o cancelación
 *    de Wise confirmada), la evidencia con fecha igual o anterior a la cancelación no cuenta y la
 *    tarea vuelve (ver registrarPagado_ y Cancelaciones.gs).
 *
 * 8) CANCELADAS: la columna COMENTARIO arranca con "CANCELADO dd-MM · ..." en las tareas canceladas
 *    en un día anterior. Se repagan una sola vez; si se cancelan de nuevo, la marca dice
 *    "NO REPAGAR · AVISAR EN APPSHEET" y salta un popup. Detalle en Cancelaciones.gs.
 *
 * 9) MONTO FACTURA DEL DÍA ANTERIOR (A y C): si una tarea estaba en la agrupada anterior en
 *    CARGADO EN BANCO o COMPROBANTE CARGADO con monto, su MONTO FACTURA se trae de la hoja del día
 *    de esa fecha y no hace falta volver a pasar el lector. Detalle en montosFacturaAnteriores_.
 *
 * 3) SUBGRUPO (col K): columna tuya, texto libre, vacía por defecto. Dos filas del mismo
 *    beneficiario con etiquetas distintas van a filas distintas de la hoja agrupada. Es la forma
 *    manual de forzar una separación que la agrupación automática no hace. Se preserva al
 *    actualizar la hoja del día, pero NO al generarla de cero.
 *
 * 4) Payoneer entra en ORDEN_TIPOS: antes las filas de ese tipo se descartaban en silencio.
 *
 * 5) Para Transf. Banc. Internac. / Local (Br) / Local (Mx) / Wise, DATOS ADICIONALES trae la
 *    columna concatenated_summary de "base" tal cual, sin parsear.
 *
 * 6) EXCLUSIONES: las filas de Criptomoneda (TIPOS_EXCLUIDOS_HOJA_DIA) y las de
 *    Transf. Banc. Internac. cuyo BENEFICIARIO contenga "Backlinks Global LLC"
 *    (EXCLUIR_BENEFICIARIO_POR_TIPO) no se traen a la hoja del día. Como la hoja agrupada se
 *    alimenta de la hoja del día, tampoco aparecen ahí.
 *
 * 7) ALERTA DE MONTO ALTO: toda fila en USD que supere UMBRAL_MONTO_ALERTA se pinta de rojo, deja
 *    una nota en la celda MONEDA y se lista en un popup al terminar de generar o actualizar. La
 *    hipótesis es moneda mal cargada (BRL o ARS marcados como USD). NO es una validación: un pago
 *    grande legítimo en USD también salta, y no hay forma de distinguirlo mirando solo el monto.
 *
 * CUOTA DE DRIVE
 * Cada factura en PDF/imagen se convierte con el OCR de Drive, que tiene límite de llamadas por
 * usuario y por minuto. Por eso las dos validaciones se resuelven con una única lectura por
 * archivo, cacheada durante la corrida, y ante un error de cuota se reintenta una vez tras pausa.
 *
 * PAGOS AGRUPADOS
 * Una factura puede cubrir varias filas de pago: una sola lleva el link y las demás quedan con
 * status "Solicitado - Agrupación" sin link. La base no identifica al grupo, así que la asociación
 * se infiere por mismo CBU / DATO DE PAGO + mismo BENEFICIARIO dentro de la hoja del día.
 */

const BASE_SHEET_NAME = 'base';
const STATUS_PERMITIDOS = ['Solicitado', 'Solicitado - Agrupación'];
const STATUS_AGRUPACION = 'Solicitado - Agrupación';

const ORDEN_TIPOS = [
  'A - Transf. Banc. Local (Arg)',
  'C - Transf. Banc. Local (Arg)',
  'Criptomoneda',
  'Payoneer',
  'Paypal',
  'PIX',
  'Plataforma externa',
  'Transf. Banc. Internac.',
  'Transf. Banc. Local (Br)',
  'Transf. Banc. Local (Mx)',
  'Transf. Banc. Local (Usa)',
  'Wise'
];

const HEADERS = [
  'TIPO', 'TASK', 'CBU / DATO DE PAGO', 'DATOS ADICIONALES', 'BENEFICIARIO', 'MONTO',
  'MONTO FACTURA', 'MONEDA', 'CHECK', 'COMENTARIO', 'COMENTARIO ANTERIOR', 'SUBGRUPO',
  'INVOICE/COMPROBANTE', 'ESTADO', 'NUEVO'
];

/** Número de columna (1-based) de cada campo. Para índice de array: COL.X - 1. */
const COL = {
  TIPO: 1, TASK: 2, DATO: 3, ADIC: 4, BENEF: 5, MONTO: 6, MONTO_FAC: 7, MONEDA: 8,
  CHECK: 9, COMENT: 10, COMENT_ANT: 11, SUBGRUPO: 12, INVOICE: 13, ESTADO: 14, NUEVO: 15
};

/**
 * MONTO      lo que bajó de AppSheet (columna total_amount de "base").
 * MONTO_FAC  el importe leído de la factura por validarSeleccion: el mismo "comparable" que ya
 *            usaba la validación, o sea Importe Total menos Importe Otros Tributos.
 *
 * Para los tipos A y C (TIPOS_VALIDABLES) el monto que manda es el de la FACTURA, no el de base.
 * La hoja agrupada suma MONTO_FAC en vez de MONTO para esos tipos. Se guarda en la hoja y no se
 * relee en cada corrida porque el OCR de Drive tiene una cuota mínima por minuto: releer once
 * facturas al generar la agrupada la haría fallar casi siempre.
 * Se preserva al actualizar la hoja del día, igual que CHECK, COMENTARIO y SUBGRUPO.
 */

/**
 * Color por tipo. Todos son pasteles de la misma intensidad: la banda de tipo usa el color tal
 * cual y las filas de datos lo aclaran al 60% (aclararColor_). Un color oscuro acá rompe las dos
 * cosas: la banda queda ilegible con letra negra y las filas salen saturadas al lado del resto.
 * Usa va en arena: es el único tono que no comparte familia con ningún otro tipo ni con las
 * alertas (rojo de monto alto, naranja de pagado/repetido, amarillo de revisión, magenta de Wise).
 */
const COLOR_POR_TIPO = {
  'A - Transf. Banc. Local (Arg)': '#d9ead3',
  'C - Transf. Banc. Local (Arg)': '#b6d7a8',
  'Criptomoneda': '#fce5cd',
  'Payoneer': '#ffd966',
  'Paypal': '#cfe2f3',
  'PIX': '#d9d2e9',
  'Plataforma externa': '#f4cccc',
  'Transf. Banc. Internac.': '#fff2cc',
  'Transf. Banc. Local (Br)': '#c9daf8',
  'Transf. Banc. Local (Mx)': '#ead1dc',
  'Transf. Banc. Local (Usa)': '#d5c4a1',
  'Wise': '#d0e0e3'
};

/**
 * Alerta de monto alto: una fila en USD por encima del umbral probablemente sea una moneda mal
 * cargada (BRL o ARS marcados como USD). Se pinta la fila entera y se deja la nota en MONEDA.
 * La nota va en MONEDA y no en MONTO porque validarSeleccion pisa la nota de la celda MONTO.
 */
const UMBRAL_MONTO_ALERTA = 200;
const MONEDA_MONTO_ALERTA = 'USD';
const COLOR_MONTO_ALERTA = '#e06666';

/** Tipos en los que DATOS ADICIONALES se llena con la columna D de "base" (summary) sin parsear. */
const TIPOS_SUMMARY_CRUDO = [
  'Transf. Banc. Internac.',
  'Transf. Banc. Local (Br)',
  'Transf. Banc. Local (Mx)',
  'Transf. Banc. Local (Usa)',
  'Wise'
];

/**
 * Tipos de transacción que NO se traen a la hoja del día. Como la hoja agrupada se alimenta de la
 * del día, tampoco aparecen ahí. Se excluyen acá y no sacándolos de ORDEN_TIPOS para que quede
 * explícito: un tipo ausente de ORDEN_TIPOS se descarta en silencio y no te enterás.
 */
const TIPOS_EXCLUIDOS_HOJA_DIA = ['Criptomoneda'];

/**
 * Filas que NO se traen a la hoja del día (y por lo tanto tampoco a la agrupada, que se alimenta
 * de ella). Se excluye por tipo de transacción + texto contenido en el BENEFICIARIO.
 * La comparación ignora mayúsculas, tildes y puntuación, pero NO reagrupa siglas: "L.L.C." queda
 * como "L L C" y no matchea contra "LLC". En la base los 28 casos vienen como "Backlinks Global
 * LLC" exacto, así que alcanza; si aparece otra variante, agregarla a la lista.
 */
const EXCLUIR_BENEFICIARIO_POR_TIPO = {
  'Transf. Banc. Internac.': ['Backlinks Global LLC']
};

function filaExcluida_(tipo, beneficiario) {
  const lista = EXCLUIR_BENEFICIARIO_POR_TIPO[tipo];
  if (!lista) return false;
  const texto = textoPlano_(beneficiario);
  if (texto === '') return false;
  return lista.some(x => texto.indexOf(textoPlano_(x)) !== -1);
}

/** Texto en mayúsculas, sin tildes ni puntuación, para comparar nombres por contención. */
function textoPlano_(valor) {
  return String(valor == null ? '' : valor)
    .toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

/* ══════════════════ ESTADOS DE LA HOJA AGRUPADA ══════════════════ */

/**
 * Vocabulario cerrado de la columna ESTADO de la hoja agrupada. Se define acá y no en
 * Hojaagrupada.gs porque historialDiasAnteriores_ necesita TRANSFERIDO, y las dependencias entre
 * archivos ya van en ese sentido (Hojaagrupada depende de este archivo, nunca al revés).
 *
 * SOLO "TRANSFERIDO" cuenta como pago hecho. Ningún otro valor saca la fila de la agrupada.
 * El CHECK no interviene: significa "revisé esta fila", no "la transferencia salió".
 */
const EST_CARGADO     = 'CARGADO EN BANCO';
const EST_TRANSFERIDO = 'TRANSFERIDO';
const EST_COMPROBANTE = 'COMPROBANTE CARGADO';
const EST_PROBLEMA    = 'CON PROBLEMA';
const EST_RECHAZADO   = 'RECHAZADO';
const EST_CANCELADO   = 'CANCELADO';

const ESTADOS_AGRUP = [EST_CARGADO, EST_TRANSFERIDO, EST_COMPROBANTE,
                       EST_PROBLEMA, EST_RECHAZADO, EST_CANCELADO];

/**
 * Estados que valen como "la plata salió". COMPROBANTE CARGADO es posterior a TRANSFERIDO en el
 * circuito (transferís, sacás el comprobante, lo armás y lo subís), así que también cuenta: si no,
 * poner el comprobante haría que la fila dejara de estar marcada como pagada y volviera mañana.
 */
const ESTADOS_PAGADO = [EST_TRANSFERIDO, EST_COMPROBANTE];

/* ══════════════════ HISTORIAL DE DÍAS ANTERIORES ══════════════════ */

/** Cuántas hojas anteriores se miran hacia atrás (vale para las del día y para las agrupadas). */
const MAX_HOJAS_ANTERIORES = 10;
const RE_NOMBRE_HOJA_DIA = /^\d{2}-\d{2}-\d{4}$/;
const RE_NOMBRE_HOJA_AGRUP = /^AGRUP (\d{2}-\d{2}-\d{4})$/;
const MARCA_PAGADO_PREFIJO = '⚠️ PAGADO EL ';
const MARCA_PAGADO_SUFIJO = ' - FALTA COMPROBANTE';
const MARCA_RETIRADO = 'RETIRADO DE LA BASE';
const COLOR_YA_PAGADO = '#f9cb9c';

/**
 * Alto fijo de las filas de la hoja del día, en píxeles (21 = alto estándar de Sheets).
 * El summary crudo de DATOS ADICIONALES trae saltos de línea reales, y ante eso Sheets agranda
 * la fila aunque la celda esté en CLIP. Se fuerza el alto al final de escribirHojaDelDia_.
 */
const ALTO_FILA_DIA = 21;

function generarHojaPagosDelDia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = formatearNombreHoja_(new Date());

  const filas = leerFilasDeBase_(ss);
  // Antes del historial: una cancelación anula la evidencia de pago anterior (registrarPagado_).
  sincronizarCancelaciones_(ss, nombreHoja);
  const hist = historialDiasAnteriores_(ss, nombreHoja);
  const comentarios = comentariosAgrupAnteriores_(ss, nombreHoja);
  const pagadas = aplicarHistorial_(filas, hist, comentarios);
  const dobles = aplicarCancelacionesHojaDia_(filas, estadoCancelaciones_(ss));
  const traidos = aplicarMontosAnteriores_(filas, montosFacturaAnteriores_(ss, nombreHoja));
  const salida = agruparPorTipo_(filas);

  let hoja = ss.getSheetByName(nombreHoja);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(nombreHoja);

  escribirHojaDelDia_(hoja, salida);
  marcarMontosTraidos_(hoja, salida, traidos);
  repintarCanceladasWise_();   // wise canceladas.gs: el magenta va antes de los popups, que bloquean
  avisarPagadas_(ss, pagadas, Object.keys(traidos).length);
  avisarCanceladasDosVeces_(dobles);
  avisarMontosAltos_(salida);
}

function actualizarHojaPagosDelDia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = formatearNombreHoja_(new Date());
  const hojaExistente = ss.getSheetByName(nombreHoja);

  if (!hojaExistente) {
    generarHojaPagosDelDia();
    return;
  }

  const previos = leerEstadoPrevio_(hojaExistente);

  const filasFrescas = leerFilasDeBase_(ss);
  const clavesFrescas = {};
  filasFrescas.forEach(f => {
    const clave = claveFila_(f[COL.TASK - 1], f[COL.MONTO - 1]);
    clavesFrescas[clave] = true;
    const prev = previos[clave];
    if (prev) {
      f[COL.CHECK - 1] = prev.check;
      f[COL.COMENT - 1] = prev.comentario;
      f[COL.SUBGRUPO - 1] = prev.subgrupo;
      f[COL.MONTO_FAC - 1] = prev.montoFactura;
    } else {
      f[COL.NUEVO - 1] = '⭐ NUEVO';
    }
  });

  // Antes del historial: una cancelación anula la evidencia de pago anterior (registrarPagado_).
  sincronizarCancelaciones_(ss, nombreHoja);
  const hist = historialDiasAnteriores_(ss, nombreHoja);
  const comentarios = comentariosAgrupAnteriores_(ss, nombreHoja);
  const pagadas = aplicarHistorial_(filasFrescas, hist, comentarios);
  const dobles = aplicarCancelacionesHojaDia_(filasFrescas, estadoCancelaciones_(ss));
  // Después de restaurar MONTO FACTURA de la hoja de hoy: lo que ya tenías no se pisa.
  const traidos = aplicarMontosAnteriores_(filasFrescas, montosFacturaAnteriores_(ss, nombreHoja));

  const filasRetiradas = [];
  Object.keys(previos).forEach(clave => {
    if (!clavesFrescas[clave]) {
      const p = previos[clave];
      if (p.check || (p.comentario && String(p.comentario).trim() !== '')) {
        filasRetiradas.push(p.filaCompleta);
      }
    }
  });

  const salida = agruparPorTipo_(filasFrescas, filasRetiradas);

  ss.deleteSheet(hojaExistente);
  const hoja = ss.insertSheet(nombreHoja);
  escribirHojaDelDia_(hoja, salida);
  marcarMontosTraidos_(hoja, salida, traidos);
  repintarCanceladasWise_();   // wise canceladas.gs: el magenta va antes de los popups, que bloquean
  avisarPagadas_(ss, pagadas, Object.keys(traidos).length);
  avisarCanceladasDosVeces_(dobles);
  avisarMontosAltos_(salida);
}

/** Un solo toast: dos toasts seguidos se pisan y el primero no se llega a leer. */
function avisarPagadas_(ss, cantidad, montosTraidos) {
  const partes = [];
  if (cantidad > 0) {
    partes.push(cantidad + ' fila(s) figuran como TRANSFERIDO en un día anterior: quedan marcadas como PAGADO en la columna ESTADO. Solo falta subir el comprobante.');
  }
  if (montosTraidos > 0) {
    partes.push(montosTraidos + ' MONTO FACTURA traído(s) de la hoja del día anterior (en cursiva, con nota): no hace falta volver a pasar el lector.');
  }
  if (partes.length) ss.toast(partes.join('  ·  '), 'Pagos LB', 10);
}

function claveFila_(task, monto) {
  return String(task) + '|' + String(monto);
}

/**
 * Devuelve, para cada header conocido, su posición (0-based) en la hoja indicada, o -1 si no está.
 * Se busca por NOMBRE y no por posición fija para que las hojas viejas (12 o 13 columnas, sin
 * COMENTARIO ANTERIOR o sin SUBGRUPO) se sigan leyendo bien.
 */
function indicesPorHeader_(hoja) {
  const ancho = Math.max(hoja.getLastColumn(), 1);
  const fila = hoja.getRange(1, 1, 1, ancho).getValues()[0]
    .map(h => String(h == null ? '' : h).trim().toUpperCase());
  const idx = {};
  HEADERS.forEach(nombre => { idx[nombre] = fila.indexOf(nombre.toUpperCase()); });
  return idx;
}

function parsearNombreHoja_(nombre) {
  const m = String(nombre).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

/**
 * Anota una tarea como pagada. Precedencia:
 *   - si la tarea está cancelada con fecha igual o posterior a esta evidencia, la evidencia NO
 *     cuenta: la plata no salió. Dos fuentes:
 *       · cancelación de Wise confirmada (pagoAnuladoPorWise_, en wise canceladas.gs);
 *       · ESTADO = CANCELADO en una agrupada anterior (pagoAnuladoPorCancelacion_, en
 *         Cancelaciones.gs).
 *     Aplica también a las marcas heredadas, porque llevan la fecha original de la transferencia;
 *   - una marca heredada nunca se pisa: es la que trae la fecha original de la transferencia;
 *   - una marca heredada sí pisa a una fecha deducida de TRANSFERIDO o de un CHECK;
 *   - entre dos fechas del mismo rango gana la primera encontrada, y como las hojas se recorren de
 *     la más nueva a la más vieja, esa es la más reciente.
 */
function registrarPagado_(indice, clave, fecha, heredado) {
  if (pagoAnuladoPorWise_(clave, fecha) || pagoAnuladoPorCancelacion_(clave, fecha)) return;
  const previo = indice[clave];
  if (previo && (previo.heredado || !heredado)) return;
  indice[clave] = { fecha: fecha, heredado: !!heredado };
}

/** Extrae el "dd-MM" de una marca "⚠️ PAGADO EL 02-09 - FALTA COMPROBANTE". null si no lo es. */
function fechaDeMarcaPagado_(estado) {
  const m = String(estado == null ? '' : estado).match(/PAGADO EL (\d{2}-\d{2})/);
  return m ? m[1] : null;
}

/**
 * Recorre las hojas "AGRUP DD-MM-YYYY" anteriores a hoy y marca como pagadas las tareas de toda
 * fila cuyo ESTADO sea TRANSFERIDO o COMPROBANTE CARGADO. Una fila agrupada cubre varias tareas,
 * así que el estado se
 * reparte a todas las que figuran en su columna de IDs.
 * Escribe sobre el índice que recibe; no devuelve nada.
 * Si la hoja agrupada no tiene columna ESTADO (las generadas antes de este cambio), se saltea.
 */
function transferidasEnAgrupAnteriores_(ss, fechaHoy, indicePorTask) {
  const anteriores = hojasAnterioresA_(ss, fechaHoy, h => parsearNombreHojaAgrup_(h.getName()));

  anteriores.forEach(o => {
    const ultimaFila = o.hoja.getLastRow();
    if (ultimaFila < 2) return;

    const ancho = Math.max(o.hoja.getLastColumn(), 1);
    const hdr = o.hoja.getRange(1, 1, 1, ancho).getValues()[0]
      .map(h => String(h == null ? '' : h).trim().toUpperCase());

    const iIds = hdr.indexOf('TRANSACCIONES (IDS)');
    const iEstado = hdr.indexOf('ESTADO');
    if (iIds < 0 || iEstado < 0) return;

    const etiqueta = Utilities.formatDate(o.fecha, ss.getSpreadsheetTimeZone(), 'dd-MM');
    const valores = o.hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues();

    valores.forEach(v => {
      const est = String(v[iEstado] == null ? '' : v[iEstado]).trim().toUpperCase();
      if (ESTADOS_PAGADO.indexOf(est) === -1) return;
      String(v[iIds] || '').split(',').forEach(id => {
        const task = id.trim();
        if (task === '') return;
        registrarPagado_(indicePorTask, task, etiqueta, false);
      });
    });
  });
}

/**
 * Devuelve { porClave, porTask } con las tareas ya pagadas y en qué fecha.
 *   porClave[TASK|MONTO] → CHECK tildado en una hoja del día anterior (compatibilidad con hojas
 *                          viejas; el CHECK ya no es la fuente de verdad del pago).
 *   porTask[TASK]        → ESTADO = TRANSFERIDO en una hoja agrupada anterior, o marca
 *                          "PAGADO EL" heredada de la columna ESTADO de una hoja del día anterior.
 *
 * porTask usa TASK sola y no TASK + MONTO porque la hoja agrupada guarda el monto SUMADO del
 * grupo, no el de cada tarea. Es el mismo criterio que ya usa comentariosAgrupAnteriores_.porTask.
 *
 * "Día anterior" es la última hoja existente con fecha menor a hoy, no la fecha calendario de
 * ayer: un lunes trae lo del viernes.
 * Los comentarios NO salen de acá: salen de comentariosAgrupAnteriores_.
 */
function historialDiasAnteriores_(ss, nombreHojaHoy) {
  const hoy = parsearNombreHoja_(nombreHojaHoy);
  const hist = { porClave: {}, porTask: {} };
  if (!hoy) return hist;

  const anteriores = hojasAnterioresA_(ss, hoy, h => parsearNombreHoja_(h.getName()));

  anteriores.forEach(o => {
    const ultimaFila = o.hoja.getLastRow();
    if (ultimaFila < 2) return;

    const idx = indicesPorHeader_(o.hoja);
    if (idx['TASK'] < 0 || idx['MONTO'] < 0) return;

    const etiqueta = Utilities.formatDate(o.fecha, ss.getSpreadsheetTimeZone(), 'dd-MM');
    const ancho = o.hoja.getLastColumn();
    const valores = o.hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues();

    valores.forEach(v => {
      const task = v[idx['TASK']];
      if (task === '' || task == null) return;

      // Marca ya escrita en un día anterior: trae la fecha original y tiene prioridad.
      if (idx['ESTADO'] >= 0) {
        const original = fechaDeMarcaPagado_(v[idx['ESTADO']]);
        if (original) registrarPagado_(hist.porTask, String(task), original, true);
      }

      if (idx['CHECK'] < 0 || v[idx['CHECK']] !== true) return;
      registrarPagado_(hist.porClave, claveFila_(task, v[idx['MONTO']]), etiqueta, false);
    });
  });

  transferidasEnAgrupAnteriores_(ss, hoy, hist.porTask);

  return hist;
}

/**
 * Lista las hojas anteriores a una fecha, de la más nueva a la más vieja, hasta el tope.
 * fechaDeHoja recibe la hoja y devuelve su fecha, o null si el nombre no corresponde.
 */
function hojasAnterioresA_(ss, fechaCorte, fechaDeHoja) {
  return ss.getSheets()
    .map(h => ({ hoja: h, fecha: fechaDeHoja(h) }))
    .filter(o => o.fecha && o.fecha.getTime() < fechaCorte.getTime())
    .sort((a, b) => b.fecha.getTime() - a.fecha.getTime())
    .slice(0, MAX_HOJAS_ANTERIORES);
}

/** Fecha de una hoja "AGRUP DD-MM-YYYY", o null si el nombre no es ese. */
function parsearNombreHojaAgrup_(nombre) {
  const m = String(nombre).match(RE_NOMBRE_HOJA_AGRUP);
  return m ? parsearNombreHoja_(m[1]) : null;
}

/**
 * Recorre las hojas "AGRUP DD-MM-YYYY" anteriores a hoy, de la más nueva a la más vieja, y arma
 * dos índices con el comentario más reciente que encuentre:
 *   porTask[TASK]              → para la hoja del día. La fila agrupada cubre varias tareas, así
 *                                que su comentario se reparte a todas las que figuran en IDs.
 *   porGrupo[tipo|nombre|mon]  → para la hoja agrupada de hoy, que se compara grupo contra grupo.
 *
 * La clave del día es TASK sola, no TASK+MONTO: la hoja agrupada guarda el monto SUMADO del grupo,
 * no el de cada tarea, así que no hay con qué armar la clave compuesta.
 *
 * La marca "PAYPAL: ..." que escribe PaypalAprobaciones.gs al final del comentario NO se arrastra:
 * cambia con cada estado de PayPal y COMENTARIO ANTERIOR es parte de la clave de grupo.
 *
 * Devuelve { porTask:{}, porGrupo:{} }. Si no hay ninguna hoja agrupada anterior, ambos vacíos.
 */
function comentariosAgrupAnteriores_(ss, nombreHojaHoy) {
  const hoy = parsearNombreHoja_(nombreHojaHoy);
  const res = { porTask: {}, porGrupo: {} };
  if (!hoy) return res;

  const anteriores = hojasAnterioresA_(ss, hoy, h => parsearNombreHojaAgrup_(h.getName()));

  anteriores.forEach(o => {
    const ultimaFila = o.hoja.getLastRow();
    if (ultimaFila < 2) return;

    const ancho = Math.max(o.hoja.getLastColumn(), 1);
    const hdr = o.hoja.getRange(1, 1, 1, ancho).getValues()[0]
      .map(h => String(h == null ? '' : h).trim().toUpperCase());

    const iTipo = hdr.indexOf('TIPO');
    const iBenef = hdr.indexOf('BENEFICIARIO');
    const iIds = hdr.indexOf('TRANSACCIONES (IDS)');
    const iMoneda = hdr.indexOf('MONEDA');
    const iComent = hdr.indexOf('COMENTARIO');
    if (iComent < 0 || iIds < 0) return; // hoja agrupada vieja, sin columna COMENTARIO

    const etiqueta = Utilities.formatDate(o.fecha, ss.getSpreadsheetTimeZone(), 'dd-MM');
    const valores = o.hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues();

    valores.forEach(v => {
      const texto = quitarMarcaPaypal_(v[iComent]);
      if (texto === '') return;
      const registro = { fecha: etiqueta, texto: texto };

      String(v[iIds] || '').split(',').forEach(id => {
        const task = id.trim();
        if (task === '') return;
        if (!res.porTask[task]) res.porTask[task] = registro;
      });

      if (iTipo >= 0) {
        const clave = claveGrupoAgrup_(v[iTipo], iBenef >= 0 ? v[iBenef] : '', iMoneda >= 0 ? v[iMoneda] : '');
        if (!res.porGrupo[clave]) res.porGrupo[clave] = registro;
      }
    });
  });

  return res;
}

/** Clave de un grupo de la hoja agrupada: tipo + beneficiario normalizado + moneda. */
function claveGrupoAgrup_(tipo, beneficiario, moneda) {
  return String(tipo) + '|' +
         Array.from(normalizarNombre_(beneficiario)).sort().join(' ') + '|' +
         String(moneda == null ? '' : moneda).trim().toUpperCase();
}

/** Escribe COMENTARIO ANTERIOR y la marca de ya pagado. Devuelve cuántas filas quedaron marcadas. */
function aplicarHistorial_(filas, hist, comentarios) {
  let pagadas = 0;
  filas.forEach(f => {
    const com = comentarios.porTask[String(f[COL.TASK - 1])];
    if (com) {
      f[COL.COMENT_ANT - 1] = com.fecha + ' | ' + com.texto;
    }

    const porTask = hist.porTask[String(f[COL.TASK - 1])];
    const porClave = hist.porClave[claveFila_(f[COL.TASK - 1], f[COL.MONTO - 1])];
    // La heredada gana siempre: conserva la fecha real de la transferencia.
    const h = (porTask && porTask.heredado) ? porTask : (porClave || porTask);

    if (h) {
      f[COL.ESTADO - 1] = MARCA_PAGADO_PREFIJO + h.fecha + MARCA_PAGADO_SUFIJO;
      pagadas++;
    }
  });
  return pagadas;
}

/**
 * Lee la hoja del día que ya existe para preservar CHECK, COMENTARIO y SUBGRUPO. Se mapea por
 * nombre de header, así que también funciona sobre una hoja generada con una versión anterior.
 */
function leerEstadoPrevio_(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return {};

  const idx = indicesPorHeader_(hoja);
  if (idx['TASK'] < 0 || idx['MONTO'] < 0) return {};

  const ancho = hoja.getLastColumn();
  const valores = hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues();
  const previos = {};

  valores.forEach(v => {
    const task = v[idx['TASK']];
    if (task === '' || task == null) return;

    const filaPreservada = [];
    HEADERS.forEach(nombre => {
      const i = idx[nombre];
      filaPreservada.push(i >= 0 ? v[i] : '');
    });
    filaPreservada[COL.ESTADO - 1] = MARCA_RETIRADO;
    filaPreservada[COL.NUEVO - 1] = '';

    previos[claveFila_(task, v[idx['MONTO']])] = {
      check: idx['CHECK'] >= 0 ? v[idx['CHECK']] : false,
      comentario: idx['COMENTARIO'] >= 0 ? v[idx['COMENTARIO']] : '',
      subgrupo: idx['SUBGRUPO'] >= 0 ? v[idx['SUBGRUPO']] : '',
      montoFactura: idx['MONTO FACTURA'] >= 0 ? v[idx['MONTO FACTURA']] : '',
      filaCompleta: filaPreservada
    };
  });

  return previos;
}

function leerFilasDeBase_(ss) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) {
    throw new Error('No se encontró la hoja "' + BASE_SHEET_NAME + '"');
  }

  const data = baseSheet.getDataRange().getValues();
  const headerRow = data[0];
  const col = mapearColumnas_(headerRow);

  const filas = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tipo = row[col.method_type];
    const status = row[col.status_name];

    if (!tipo) continue;
    if (TIPOS_EXCLUIDOS_HOJA_DIA.indexOf(tipo) !== -1) continue;
    if (STATUS_PERMITIDOS.indexOf(status) === -1) continue;

    const summary = row[col.concatenated_summary] || '';
    const campos = parsearSummary_(summary);
    const beneficiario = campos['Beneficiary'] || '';

    if (filaExcluida_(tipo, beneficiario)) continue;

    const extraido = extraerDatoPago_(tipo, campos);

    // Para estos tipos el parseo por etiqueta venía incompleto: se pega el summary crudo.
    const adicional = TIPOS_SUMMARY_CRUDO.indexOf(tipo) !== -1
      ? String(summary)
      : extraido.adicional;

    filas.push([
      tipo,                            // TIPO
      row[col.task_name],              // TASK
      extraido.dato,                   // CBU / DATO DE PAGO
      adicional,                       // DATOS ADICIONALES
      beneficiario,                    // BENEFICIARIO
      row[col.total_amount],           // MONTO
      '',                              // MONTO FACTURA (la escribe validarSeleccion)
      row[col.currency] || '',         // MONEDA
      false,                           // CHECK
      '',                              // COMENTARIO
      '',                              // COMENTARIO ANTERIOR
      '',                              // SUBGRUPO
      row[col.payment_invoice] || '',  // INVOICE/COMPROBANTE
      '',                              // ESTADO
      ''                               // NUEVO
    ]);
  }
  return filas;
}

function agruparPorTipo_(filas, filasRetiradas) {
  filasRetiradas = filasRetiradas || [];

  const filasPorTipo = {};
  ORDEN_TIPOS.forEach(t => filasPorTipo[t] = []);

  filas.forEach(f => {
    if (filasPorTipo[f[COL.TIPO - 1]]) filasPorTipo[f[COL.TIPO - 1]].push(f);
  });
  filasRetiradas.forEach(f => {
    if (filasPorTipo[f[COL.TIPO - 1]]) filasPorTipo[f[COL.TIPO - 1]].push(f);
  });

  const salida = [];
  ORDEN_TIPOS.forEach(tipo => {
    const filasTipo = filasPorTipo[tipo];
    if (filasTipo.length === 0) return;

    const encabezado = [];
    for (let i = 0; i < HEADERS.length; i++) encabezado.push('');
    encabezado[COL.TIPO - 1] = tipo;
    salida.push(encabezado);

    filasTipo.forEach(f => salida.push(f));
  });
  return salida;
}

function esEncabezadoDeGrupo_(fila) {
  return fila[COL.TASK - 1] === '' && fila[COL.TIPO - 1] !== '';
}

function escribirHojaDelDia_(hoja, salida) {
  const headerRange = hoja.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#434343')
    .setHorizontalAlignment('center');
  hoja.setFrozenRows(1);

  if (salida.length > 0) {
    // Formato TEXTO en C y D ANTES de escribir: si no, Sheets convierte a número las claves
    // PIX / CBU / cuentas que parecen numéricas y les come el cero inicial.
    hoja.getRange(2, COL.DATO, salida.length, 2).setNumberFormat('@');

    hoja.getRange(2, 1, salida.length, HEADERS.length).setValues(salida);

    hoja.getRange(2, COL.MONTO, salida.length, 2).setNumberFormat('#,##0.00');

    // El summary crudo tiene hasta 13 líneas: con CLIP la fila mantiene su alto normal y el texto
    // completo se ve al hacer clic en la celda.
    hoja.getRange(2, COL.ADIC, salida.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    hoja.getRange(2, COL.COMENT_ANT, salida.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    salida.forEach((fila, i) => {
      if (esEncabezadoDeGrupo_(fila)) return;
      const filaSheet = i + 2;

      hoja.getRange(filaSheet, COL.CHECK).insertCheckboxes();

      const estado = String(fila[COL.ESTADO - 1] || '');
      if (estado === MARCA_RETIRADO) {
        hoja.getRange(filaSheet, COL.ESTADO).setFontColor('#cc0000').setFontWeight('bold');
      } else if (estado.indexOf(MARCA_PAGADO_PREFIJO) === 0) {
        hoja.getRange(filaSheet, COL.ESTADO)
          .setBackground(COLOR_YA_PAGADO)
          .setFontColor('#783f04')
          .setFontWeight('bold');
      }

      if (fila[COL.NUEVO - 1] === '⭐ NUEVO') {
        hoja.getRange(filaSheet, COL.NUEVO).setFontColor('#e69138').setFontWeight('bold');
      }

      // Marca de Cancelaciones.gs: en negrita, y en rojo si ya no se puede repagar.
      const coment = String(fila[COL.COMENT - 1] || '');
      if (coment.indexOf('CANCELADO') === 0) {
        const celda = hoja.getRange(filaSheet, COL.COMENT).setFontWeight('bold');
        if (coment.indexOf('AVISAR EN APPSHEET') !== -1) celda.setFontColor('#cc0000');
      }
    });

    hoja.getRange(1, 1, salida.length + 1, HEADERS.length)
      .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

    resaltarFilasDeGrupo_(hoja, salida);

    hoja.getRange(1, 1, salida.length + 1, HEADERS.length).createFilter();
  }

  hoja.autoResizeColumns(1, HEADERS.length);
  // Anchos fijos después del autoajuste: si no, DATOS ADICIONALES se estira con el summary entero.
  hoja.setColumnWidth(COL.ADIC, 260);
  hoja.setColumnWidth(COL.COMENT, 220);
  hoja.setColumnWidth(COL.COMENT_ANT, 240);
  hoja.setColumnWidth(COL.MONTO_FAC, 110);
  hoja.setColumnWidth(COL.SUBGRUPO, 110);
  hoja.setColumnWidth(COL.ESTADO, 230);

  // Alto uniforme: sin esto, cada fila con summary multilínea se estira a 13 renglones.
  if (salida.length > 0) {
    try {
      hoja.setRowHeightsForced(2, salida.length, ALTO_FILA_DIA);
    } catch (e) {
      hoja.setRowHeights(2, salida.length, ALTO_FILA_DIA);
    }
  }

  recortarGrilla_(hoja, salida.length + 1, HEADERS.length);
}

function mapearColumnas_(headerRow) {
  const nombres = {
    task_name: ['task_name', 'TA'],
    method_type: ['method_type'],
    concatenated_summary: ['concatenated_summary'],
    total_amount: ['total_amount'],
    currency: ['currency'],
    status_name: ['status_name'],
    payment_invoice: ['payment_invoice']
  };
  const col = {};
  Object.keys(nombres).forEach(key => {
    const idx = headerRow.findIndex(h => nombres[key].indexOf(h) !== -1);
    if (idx === -1) {
      throw new Error('No se encontró la columna "' + key + '" en la hoja "' + BASE_SHEET_NAME + '". Revisar headers.');
    }
    col[key] = idx;
  });
  return col;
}

function parsearSummary_(summary) {
  const campos = {};
  String(summary).split('\n').forEach(linea => {
    const idx = linea.indexOf(':');
    if (idx === -1) return;
    const label = linea.substring(0, idx).trim();
    const valor = linea.substring(idx + 1).trim();
    campos[label] = valor;
  });
  return campos;
}

function extraerDatoPago_(tipo, campos) {
  switch (tipo) {
    case 'A - Transf. Banc. Local (Arg)':
    case 'C - Transf. Banc. Local (Arg)':
      return { dato: campos['CBU/ALIAS'] || '', adicional: '' };

    case 'PIX':
      return { dato: campos['Chave PIX'] || '', adicional: campos['Email'] || '' };

    case 'Criptomoneda':
      return { dato: campos['TRC-20'] || '', adicional: '' };

    case 'Transf. Banc. Internac.':
      return {
        dato: campos['IBAN'] || '',
        adicional: [campos['SWIFT'] ? 'SWIFT: ' + campos['SWIFT'] : '', campos['Institution'] ? campos['Institution'] : '']
          .filter(String).join(' / ')
      };

    case 'Transf. Banc. Local (Br)':
      return {
        dato: campos['N° de Cuenta'] || '',
        adicional: campos['Agencia'] ? 'Ag. ' + campos['Agencia'] : ''
      };

    case 'Transf. Banc. Local (Mx)':
      return {
        dato: campos['CLABE'] || '',
        adicional: campos['Institution'] ? campos['Institution'] : ''
      };
    case 'Transf. Banc. Local (Usa)':
      return {
        dato: campos['CLABE'] || '',
        adicional: campos['Institution'] ? campos['Institution'] : ''
      };


    case 'Wise':
    case 'Paypal':
    case 'Payoneer':
      return { dato: campos['Email'] || '', adicional: '' };

    case 'Plataforma externa':
      return { dato: campos['Email'] || campos['URL'] || '', adicional: '' };

    default:
      return { dato: '', adicional: '' };
  }
}

function formatearNombreHoja_(fecha) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(fecha, tz, 'dd-MM-yyyy');
}

function resaltarFilasDeGrupo_(hoja, salida) {
  let colorActual = null;
  salida.forEach((fila, i) => {
    const filaSheet = i + 2;

    if (esEncabezadoDeGrupo_(fila)) {
      colorActual = COLOR_POR_TIPO[fila[COL.TIPO - 1]] || '#d9d9d9';
      hoja.getRange(filaSheet, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground(colorActual)
        .setFontColor('#000000');
    } else if (colorActual) {
      // El rojo de monto alto pisa el color por tipo en toda la fila.
      const alerta = montoAltoSospechoso_(fila[COL.MONTO - 1], fila[COL.MONEDA - 1]);
      hoja.getRange(filaSheet, 1, 1, HEADERS.length)
        .setBackground(alerta ? COLOR_MONTO_ALERTA : aclararColor_(colorActual, 0.6));

      if (alerta) {
        hoja.getRange(filaSheet, COL.MONEDA).setNote(
          '⚠️ ' + formatearImporte_(aNumero_(fila[COL.MONTO - 1])) + ' ' + MONEDA_MONTO_ALERTA +
          ' supera el umbral de ' + UMBRAL_MONTO_ALERTA + '.\n' +
          'Verificar que la moneda sea correcta y no BRL o ARS cargado como ' + MONEDA_MONTO_ALERTA + '.'
        );
      }

      const estado = String(fila[COL.ESTADO - 1] || '');
      if (estado.indexOf(MARCA_PAGADO_PREFIJO) === 0) {
        hoja.getRange(filaSheet, COL.ESTADO).setBackground(COLOR_YA_PAGADO);
      }
    }
  });
}

/**
 * Borra las filas y columnas sobrantes de una hoja recién generada. Sheets crea 1000x26 por
 * defecto y esas celdas vacías cuentan para el límite de 10 millones de celdas del archivo.
 * Se dejan como mínimo 2 filas: no se puede tener una hoja con una sola fila y esa fila congelada.
 */
function recortarGrilla_(hoja, filasUsadas, columnasUsadas) {
  const filas = Math.max(filasUsadas, 2);
  const sobranFilas = hoja.getMaxRows() - filas;
  if (sobranFilas > 0) hoja.deleteRows(filas + 1, sobranFilas);

  const sobranColumnas = hoja.getMaxColumns() - columnasUsadas;
  if (sobranColumnas > 0) hoja.deleteColumns(columnasUsadas + 1, sobranColumnas);
}

function aclararColor_(hex, factor) {
  const r = parseInt(hex.substring(1, 3), 16);
  const g = parseInt(hex.substring(3, 5), 16);
  const b = parseInt(hex.substring(5, 7), 16);
  const mix = c => Math.round(c + (255 - c) * factor);
  const toHex = c => c.toString(16).padStart(2, '0');
  return '#' + toHex(mix(r)) + toHex(mix(g)) + toHex(mix(b));
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Pagos LB')
    .addItem('Generar hoja del día', 'menuGenerarHojaDia')
    .addItem('Actualizar hoja del día', 'menuActualizarHojaDia')
    .addItem('Validar beneficiario y monto (filas seleccionadas)', 'validarSeleccion')
    .addItem('[DIAGNÓSTICO] Ver texto OCR de fila', 'diagnosticoOcrFila')
    .addToUi();

  ui.createMenu('Agrupado')
    .addItem('Generar hoja agrupada del día', 'menuGenerarAgrupada')
    .addItem('Actualizar hoja agrupada del día', 'menuActualizarAgrupada')
    .addSeparator()
    .addItem('Galicia: generar archivo de carga al banco', 'generarArchivoGalicia')
    .addItem('Generar hoja de APROBACIÓN', 'generarHojaAprobacion')
    .addItem('PayPal: revisar aprobaciones y confirmar abonados', 'revisarAprobacionesPaypal')
    .addSeparator()
    .addItem('Armar comprobante de pago (panel)', 'abrirPanelComprobante')
    .addItem('Armar comprobantes EN LOTE', 'abrirPanelComprobantesLote')
    .addToUi();
}

function tacharFilaSeleccionada() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const filaActiva = hoja.getActiveRange().getRow();

  if (filaActiva === 1) {
    SpreadsheetApp.getUi().alert('No se puede tachar la fila de encabezado.');
    return;
  }

  const primeraCelda = hoja.getRange(filaActiva, COL.TASK).getValue();
  if (primeraCelda === '') {
    SpreadsheetApp.getUi().alert('Esa fila es un encabezado de grupo, no un pago. Seleccioná una fila de datos.');
    return;
  }

  const rango = hoja.getRange(filaActiva, 1, 1, HEADERS.length);
  const yaTachada = rango.getFontLine() === 'line-through';
  rango.setFontLine(yaTachada ? 'none' : 'line-through');
}

/* ══════════════════ MONTO FACTURA DEL DÍA ANTERIOR ══════════════════ */

/** Principio de la nota que deja un monto traído. Se usa también para encadenar días. */
const NOTA_MONTO_TRAIDO = '↻ MONTO DE FACTURA TRAÍDO';

/**
 * Estados de la agrupada anterior que valen como "acepté este monto". Es una función y no una
 * constante por lo mismo que estadosQueSeparan_: EST_* tiene que existir cuando se evalúa.
 */
function estadosQueArrastranMonto_() {
  return [EST_CARGADO, EST_COMPROBANTE];
}

/** Clave de una factura de una tarea: TASK + link de INVOICE/COMPROBANTE. */
function claveFactura_(task, link) {
  return String(task == null ? '' : task).trim() + '|' + String(link == null ? '' : link).trim();
}

/**
 * MONTO FACTURA a traer, indexado por TASK|LINK.
 *
 * De dónde sale: la última agrupada anterior a hoy y la hoja del día de ESA misma fecha (la
 * agrupada se armó con ella). El monto se toma por tarea de la hoja del día, no de la agrupada:
 * la agrupada guarda el monto SUMADO del grupo y no hay forma de repartirlo entre las tareas.
 *
 * Una tarea se trae si es de tipo A o C, tiene link de factura y MONTO FACTURA numérico en la hoja
 * del día anterior, y además:
 *   - en la agrupada anterior su grupo estaba en CARGADO EN BANCO o COMPROBANTE CARGADO, con
 *     monto: ese estado es tu aceptación del monto; o
 *   - en la hoja del día anterior el monto ya había sido traído (tiene la nota NOTA_MONTO_TRAIDO).
 *     Sin esto la cadena se corta al segundo día: las filas "↻ EN BANCO" nacen sin estado y
 *     "Marcar todo como CARGADO EN BANCO" no las marca, así que la agrupada de hoy no las vería
 *     como aceptadas.
 * Las filas satélite ("Solicitado - Agrupación") no tienen link ni MONTO FACTURA: no se traen ni
 * hace falta, porque la agrupada suma por factura.
 */
function montosFacturaAnteriores_(ss, nombreHojaHoy) {
  const res = {};
  const hoy = parsearNombreHoja_(nombreHojaHoy);
  if (!hoy) return res;

  const agrups = hojasAnterioresA_(ss, hoy, h => parsearNombreHojaAgrup_(h.getName()));
  if (agrups.length === 0) return res;

  const ag = agrups[0].hoja;
  const nombreDiaAnt = ag.getName().replace(/^AGRUP /, '');
  const etiqueta = nombreDiaAnt.substring(0, 5);
  const hojaDiaAnt = ss.getSheetByName(nombreDiaAnt);
  if (!hojaDiaAnt || hojaDiaAnt.getLastRow() < 2) return res;

  // 1) Tareas cuyo grupo quedó aceptado en la agrupada anterior.
  const aceptadas = {};
  if (ag.getLastRow() >= 2) {
    const ancho = Math.max(ag.getLastColumn(), 1);
    const hdr = ag.getRange(1, 1, 1, ancho).getValues()[0]
      .map(h => String(h == null ? '' : h).trim().toUpperCase());
    const iTipo = hdr.indexOf('TIPO');
    const iIds = hdr.indexOf('TRANSACCIONES (IDS)');
    const iEstado = hdr.indexOf('ESTADO');
    const iMonto = hdr.indexOf('MONTO');
    if (iTipo >= 0 && iIds >= 0 && iEstado >= 0 && iMonto >= 0) {
      const estados = estadosQueArrastranMonto_();
      ag.getRange(2, 1, ag.getLastRow() - 1, ancho).getValues().forEach(v => {
        if (TIPOS_VALIDABLES.indexOf(String(v[iTipo]).trim()) === -1) return;
        const est = String(v[iEstado] == null ? '' : v[iEstado]).trim().toUpperCase();
        if (estados.indexOf(est) === -1) return;
        const m = aNumero_(v[iMonto]);
        if (m === null || m <= 0) return;
        String(v[iIds] || '').split(',').map(x => x.trim()).filter(String)
          .forEach(t => { aceptadas[t] = est; });
      });
    }
  }

  // 2) MONTO FACTURA por tarea en la hoja del día de esa fecha.
  const idx = indicesPorHeader_(hojaDiaAnt);
  const iTask = idx['TASK'];
  const iMF = idx['MONTO FACTURA'];
  const iInv = idx['INVOICE/COMPROBANTE'];
  const iTipoD = idx['TIPO'];
  if (iTask < 0 || iMF < 0 || iInv < 0) return res;

  const n = hojaDiaAnt.getLastRow() - 1;
  const valores = hojaDiaAnt.getRange(2, 1, n, hojaDiaAnt.getLastColumn()).getValues();
  const notas = hojaDiaAnt.getRange(2, iMF + 1, n, 1).getNotes();

  valores.forEach((v, i) => {
    const task = String(v[iTask] == null ? '' : v[iTask]).trim();
    if (task === '') return;
    if (iTipoD >= 0 && TIPOS_VALIDABLES.indexOf(String(v[iTipoD]).trim()) === -1) return;
    const link = String(v[iInv] == null ? '' : v[iInv]).trim();
    if (link === '') return;
    const mf = aNumero_(v[iMF]);
    if (mf === null || mf <= 0) return;

    const est = aceptadas[task];
    const yaTraido = String(notas[i][0] || '').indexOf(NOTA_MONTO_TRAIDO) === 0;
    if (!est && !yaTraido) return;

    res[claveFactura_(task, link)] = {
      monto: mf,
      fecha: etiqueta,
      motivo: est ? 'estado ' + est : 'ya venía traído de un día anterior'
    };
  });

  return res;
}

/**
 * Completa MONTO FACTURA en las filas A y C que tienen la MISMA tarea y el MISMO link de factura
 * que en el día anterior. Si el link cambió, es otra factura y no se trae nada.
 * Nunca pisa un MONTO FACTURA que ya esté en la fila (de una actualización de hoy o de un
 * validarSeleccion). Si el que está coincide con el traído, igual se lo marca como traído para que
 * la cadena siga mañana.
 * Devuelve TASK|LINK → info de las filas marcadas, para marcarMontosTraidos_.
 */
function aplicarMontosAnteriores_(filas, anteriores) {
  const traidos = {};
  filas.forEach(f => {
    if (TIPOS_VALIDABLES.indexOf(f[COL.TIPO - 1]) === -1) return;
    const k = claveFactura_(f[COL.TASK - 1], f[COL.INVOICE - 1]);
    const a = anteriores[k];
    if (!a) return;
    const actual = aNumero_(f[COL.MONTO_FAC - 1]);
    if (actual === null) {
      f[COL.MONTO_FAC - 1] = a.monto;
      traidos[k] = a;
    } else if (redondear2_(actual) === redondear2_(a.monto)) {
      traidos[k] = a;
    }
  });
  return traidos;
}

/** Nota y cursiva en MONTO FACTURA de las filas cuyo monto se trajo del día anterior. */
function marcarMontosTraidos_(hoja, salida, traidos) {
  if (Object.keys(traidos).length === 0) return;
  salida.forEach((fila, i) => {
    if (esEncabezadoDeGrupo_(fila)) return;
    if (fila[COL.ESTADO - 1] === MARCA_RETIRADO) return;
    const a = traidos[claveFactura_(fila[COL.TASK - 1], fila[COL.INVOICE - 1])];
    if (!a) return;
    hoja.getRange(i + 2, COL.MONTO_FAC)
      .setNote(NOTA_MONTO_TRAIDO + ' de la hoja del ' + a.fecha + ' (' + a.motivo + ').\n' +
               'No se volvió a leer la factura: es el mismo link. Si querés releerla, corré ' +
               '"Validar beneficiario y monto" sobre esta fila.')
      .setFontStyle('italic');
  });
}

/* ══════════════════ LECTURA / OCR DE FACTURAS ══════════════════ */

const PAUSA_REINTENTO_OCR_MS = 4000;

function obtenerTextoFactura_(id) {
  try {
    return leerTextoDeDrive_(id);
  } catch (err) {
    if (!esErrorDeCuotaOcr_(err)) throw err;
    Utilities.sleep(PAUSA_REINTENTO_OCR_MS);
    return leerTextoDeDrive_(id);
  }
}

function esErrorDeCuotaOcr_(err) {
  const msg = String(err && err.message ? err.message : err).toLowerCase();
  return msg.indexOf('rate limit') !== -1 || msg.indexOf('quota') !== -1;
}

function leerTextoDeDrive_(id) {
  let docTemporal = null;
  try {
    const archivoOriginal = DriveApp.getFileById(id);
    const mimeType = archivoOriginal.getMimeType();

    if (mimeType === MimeType.GOOGLE_DOCS) {
      return DocumentApp.openById(id).getBody().getText();
    }

    const blob = archivoOriginal.getBlob().setContentType(mimeType);
    const recurso = { title: 'OCR_temp_' + id };
    docTemporal = Drive.Files.insert(recurso, blob, { ocr: true, ocrLanguage: 'es' });
    return DocumentApp.openById(docTemporal.id).getBody().getText();
  } finally {
    if (docTemporal) {
      try { DriveApp.getFileById(docTemporal.id).setTrashed(true); } catch (e) { /* no bloquea el resultado */ }
    }
  }
}

function obtenerTextoFacturaCacheado_(id, cache) {
  if (cache && Object.prototype.hasOwnProperty.call(cache, id)) {
    const guardado = cache[id];
    if (guardado.error) throw guardado.error;
    return guardado.texto;
  }

  try {
    const texto = obtenerTextoFactura_(id);
    if (cache) cache[id] = { texto: texto };
    return texto;
  } catch (err) {
    if (cache) cache[id] = { error: err };
    throw err;
  }
}

/** Extrae el ID de archivo de una URL de Drive tipo .../file/d/ID/view o .../d/ID/edit. */
function extraerIdDeDrive_(url) {
  const match = String(url).match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function diagnosticoOcrFila() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const fila = hoja.getActiveRange().getRow();
  const linkFactura = hoja.getRange(fila, COL.INVOICE).getValue();

  Logger.log('Valor crudo de la celda de factura, fila ' + fila + ': [' + linkFactura + ']');

  if (!linkFactura) {
    SpreadsheetApp.getUi().alert('Esa fila no tiene link en INVOICE/COMPROBANTE.');
    return;
  }

  const id = extraerIdDeDrive_(linkFactura);
  if (!id) {
    SpreadsheetApp.getUi().alert('No se pudo extraer el ID de Drive del link.\n\nValor de la celda: ' + linkFactura);
    return;
  }

  try {
    const texto = obtenerTextoFactura_(id);
    const cabecera = recortarBloqueEmisor_(texto);
    const extraido = parsearRazonSocial_(cabecera);
    const imp = parsearImporteFactura_(texto);

    Logger.log('--- TEXTO OCR CRUDO (fila ' + fila + ') ---');
    Logger.log(texto);
    Logger.log('--- CABECERA (bloque emisor, hasta el corte) ---');
    Logger.log(cabecera);
    Logger.log('--- RAZÓN SOCIAL EXTRAÍDA ---');
    Logger.log(extraido.ok ? extraido.razonSocial : 'NO SE PUDO: ' + extraido.motivo);
    Logger.log('--- IMPORTE EXTRAÍDO ---');
    Logger.log(imp.ok ? ('Rótulo "' + imp.rotulo + '": ' + imp.total + ' − Otros Tributos ' + imp.otrosTributos + ' = ' + imp.comparable) : 'NO SE PUDO: ' + imp.motivo);

    // El bloque de totales SIEMPRE está al final del texto. Mostrar los primeros 900 caracteres
    // solo enseña la cabecera del emisor, que es justo lo que no hace falta para diagnosticar un
    // importe que no se pudo leer.
    const lineas = texto.split('\n');
    const relevantes = [];
    lineas.forEach((l, i) => {
      if (/total|subtotal|importe|neto|iva|\$/i.test(l)) {
        relevantes.push(String(i).padStart(3, ' ') + ': ' + l.trim());
      }
    });
    Logger.log('--- LÍNEAS CON TOTAL / IMPORTE / IVA / $ ---');
    Logger.log(relevantes.join('\n'));

    const cola = texto.length > 1200 ? texto.substring(texto.length - 1200) : texto;

    SpreadsheetApp.getUi().alert(
      'Razón Social detectada: ' + (extraido.ok ? extraido.razonSocial : '(sin rótulo) — se validará por presencia del nombre en la cabecera') +
      '\nImporte detectado: ' + (imp.ok ? formatearImporte_(imp.comparable) + ' (rótulo: ' + imp.rotulo + ')' : '(ninguno) — ' + imp.motivo) +
      '\n\n=== LÍNEAS CON TOTAL / IMPORTE / IVA / $ ===\n' +
      (relevantes.length ? relevantes.join('\n') : '(ninguna)') +
      '\n\n=== ÚLTIMOS 1200 CARACTERES DEL OCR ===\n' + cola
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error corriendo OCR: ' + err.message);
  }
}

/* ══════════════════ CONTEXTO DE AGRUPACIÓN ══════════════════ */

const TIPOS_VALIDABLES = ['A - Transf. Banc. Local (Arg)', 'C - Transf. Banc. Local (Arg)'];

function claveAgrupacion_(cbu, beneficiario) {
  const c = String(cbu == null ? '' : cbu).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const b = Array.from(normalizarNombre_(beneficiario)).sort().join(' ');
  return c + '|' + b;
}

function contextoAgrupacion_(hoja, statusPorTask) {
  const ultimaFila = hoja.getLastRow();
  const ctx = { satelites: {}, conFactura: {} };
  if (ultimaFila < 2) return ctx;

  const valores = hoja.getRange(2, 1, ultimaFila - 1, HEADERS.length).getValues();

  valores.forEach(v => {
    const tipo = v[COL.TIPO - 1];
    const task = v[COL.TASK - 1];
    const cbu = v[COL.DATO - 1];
    const beneficiario = v[COL.BENEF - 1];
    const monto = v[COL.MONTO - 1];
    const link = v[COL.INVOICE - 1];

    if (task === '' || tipo === '') return;
    if (TIPOS_VALIDABLES.indexOf(tipo) === -1) return;

    const clave = claveAgrupacion_(cbu, beneficiario);
    const tieneFactura = link && String(link).trim() !== '';

    if (tieneFactura) {
      if (!ctx.conFactura[clave]) ctx.conFactura[clave] = [];
      ctx.conFactura[clave].push({ task: String(task), link: link });
      return;
    }

    if (statusPorTask[String(task)] === STATUS_AGRUPACION) {
      const n = aNumero_(monto);
      if (n === null) return;
      if (!ctx.satelites[clave]) ctx.satelites[clave] = [];
      ctx.satelites[clave].push({ task: String(task), monto: n });
    }
  });

  return ctx;
}

function anclasDeGrupo_(task, clave, ctx) {
  return (ctx.conFactura[clave] || []).filter(a => a.task !== String(task));
}

/* ══════════════════ VALIDACIÓN (BENEFICIARIO + MONTO) ══════════════════ */

const COLOR_VALID_MISMATCH = '#f4cccc'; // rojo claro: no coincide
const COLOR_VALID_REVISION = '#ffe599'; // amarillo: revisión manual

function validarSeleccion() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const rango = hoja.getActiveRange();
  const filaInicio = rango.getRow();
  const cantFilas = rango.getNumRows();

  const statusPorTask = mapaStatusPorTask_();
  const ctx = contextoAgrupacion_(hoja, statusPorTask);
  const cache = {};

  let benOk = 0, benMismatch = 0, benRevision = 0;
  let monOk = 0, monMismatch = 0, monRevision = 0;
  let saltadas = 0;
  let facturasEscritas = 0;

  for (let i = 0; i < cantFilas; i++) {
    const fila = filaInicio + i;
    if (fila === 1) continue;

    const v = hoja.getRange(fila, 1, 1, HEADERS.length).getValues()[0];
    const tipo = v[COL.TIPO - 1];
    const task = v[COL.TASK - 1];
    const cbu = v[COL.DATO - 1];
    const beneficiario = v[COL.BENEF - 1];
    const monto = v[COL.MONTO - 1];
    const linkFactura = v[COL.INVOICE - 1];

    if (task === '') { saltadas++; continue; }
    if (TIPOS_VALIDABLES.indexOf(tipo) === -1) { saltadas++; continue; }

    const clave = claveAgrupacion_(cbu, beneficiario);
    const statusEnBase = statusPorTask[String(task)];
    // Si la fila está marcada por monto alto, el fondo "normal" al que se vuelve es el rojo de la
    // alerta: si no, un OK de validación le borraría la marca a media fila.
    const fondoNormal = montoAltoSospechoso_(monto, v[COL.MONEDA - 1])
      ? COLOR_MONTO_ALERTA
      : colorFondoDeFila_(tipo);

    const resBen = resultadoValidacion_(task, beneficiario, linkFactura, statusEnBase, clave, ctx, cache);
    const celdaBen = hoja.getRange(fila, COL.BENEF);
    celdaBen.setNote(resBen.texto);
    if (resBen.estado === 'OK') { celdaBen.setBackground(fondoNormal); benOk++; }
    else if (resBen.estado === 'MISMATCH') { celdaBen.setBackground(COLOR_VALID_MISMATCH); benMismatch++; }
    else { celdaBen.setBackground(COLOR_VALID_REVISION); benRevision++; }

    const resMon = resultadoValidacionMonto_(task, monto, linkFactura, statusEnBase, clave, ctx, cache);
    const celdaMon = hoja.getRange(fila, COL.MONTO);
    celdaMon.setNote(resMon.texto);
    if (resMon.estado === 'OK') { celdaMon.setBackground(fondoNormal); monOk++; }
    else if (resMon.estado === 'MISMATCH') { celdaMon.setBackground(COLOR_VALID_MISMATCH); monMismatch++; }
    else { celdaMon.setBackground(COLOR_VALID_REVISION); monRevision++; }

    // El importe leído de la factura queda guardado en la hoja: es el monto que después usa la
    // hoja agrupada para los tipos A y C. No se relee al generar la agrupada porque la cuota de
    // OCR de Drive no lo aguanta. Si esta vez no se pudo leer, NO se borra lo que ya había: un
    // fallo transitorio del OCR no debería hacer perder un importe que ya estaba verificado.
    // Una lectura nueva reemplaza al monto traído del día anterior: se borra su nota y la cursiva,
    // así mañana no se arrastra como si lo hubieras aceptado con un estado.
    if (typeof resMon.comparable === 'number') {
      hoja.getRange(fila, COL.MONTO_FAC).setValue(resMon.comparable).setNote('').setFontStyle('normal');
      facturasEscritas++;
    }
  }

  ui.alert(
    'Validación completa.\n\n' +
    'BENEFICIARIO\n' +
    '  ✅ Coinciden: ' + benOk + '\n' +
    '  ❌ No coinciden: ' + benMismatch + '\n' +
    '  ⚠️ Revisión manual: ' + benRevision + '\n\n' +
    'MONTO\n' +
    '  ✅ Coinciden: ' + monOk + '\n' +
    '  ❌ No coinciden: ' + monMismatch + '\n' +
    '  ⚠️ Revisión manual: ' + monRevision + '\n\n' +
    'Filas salteadas (tipo no aplicable): ' + saltadas + '\n\n' +
    'MONTO FACTURA escrito en ' + facturasEscritas + ' fila(s).\n' +
    'Ese es el importe que la hoja agrupada va a usar para los tipos A y C.\n\n' +
    'Pasá el mouse sobre cada celda para ver el detalle.'
  );
}

function colorFondoDeFila_(tipo) {
  return aclararColor_(COLOR_POR_TIPO[tipo] || '#d9d9d9', 0.6);
}

function resultadoValidacion_(task, beneficiario, linkFactura, statusEnBase, clave, ctx, cache) {
  if (!beneficiario || String(beneficiario).trim() === '') {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — BENEFICIARIO vacío en la hoja' };
  }

  let link = linkFactura;
  let origen = '';

  if (!link || String(link).trim() === '') {
    if (statusEnBase !== STATUS_AGRUPACION) {
      return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — la fila no tiene link de factura' };
    }

    const anclas = anclasDeGrupo_(task, clave, ctx);
    if (anclas.length === 0) {
      return {
        estado: 'REVISION',
        texto: 'NO SE PUDO, REVISION MANUAL — pago agrupado sin factura propia, y en esta hoja no ' +
               'hay ninguna fila con factura para el mismo CBU y beneficiario.'
      };
    }
    if (anclas.length > 1) {
      return {
        estado: 'REVISION',
        texto: 'NO SE PUDO, REVISION MANUAL — pago agrupado sin factura propia, y hay más de una ' +
               'fila con factura para el mismo CBU y beneficiario: ' +
               anclas.map(a => a.task).join(', ') + '. No se puede saber cuál corresponde.'
      };
    }

    link = anclas[0].link;
    origen = ' (pago agrupado: se usó la factura de ' + anclas[0].task + ')';
  }

  const an = analizarEmisorDeFactura_(link, cache);
  if (!an.ok) {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — ' + an.motivo + origen };
  }

  if (an.razonSocial) {
    if (mismoConjuntoDePalabras_(beneficiario, an.razonSocial)) {
      return { estado: 'OK', texto: '✅ COINCIDE' + origen + ' — factura: ' + an.razonSocial };
    }
    return {
      estado: 'MISMATCH',
      texto: '❌ NO COINCIDE' + origen + ' — factura: ' + an.razonSocial + '\nhoja: ' + beneficiario
    };
  }

  const tokens = normalizarNombre_(beneficiario);
  if (tokens.size < 2) {
    return {
      estado: 'REVISION',
      texto: 'NO SE PUDO, REVISION MANUAL — la factura no tiene el rótulo "Razón Social:" y ' +
             'el BENEFICIARIO es demasiado corto para buscarlo por presencia' + origen
    };
  }

  if (nombreApareceEn_(tokens, an.cabecera)) {
    return {
      estado: 'OK',
      texto: '✅ COINCIDE (verificado por presencia)' + origen + ' — la factura no tiene el rótulo ' +
             '"Razón Social:", pero el nombre del beneficiario aparece en la cabecera del emisor.' +
             '\nhoja: ' + beneficiario
    };
  }

  return {
    estado: 'REVISION',
    texto: 'NO SE PUDO, REVISION MANUAL — la factura no tiene el rótulo "Razón Social:" y tampoco ' +
           'aparece el nombre del beneficiario en la cabecera del emisor.' + origen +
           '\nhoja: ' + beneficiario
  };
}

function analizarEmisorDeFactura_(urlDrive, cache) {
  const id = extraerIdDeDrive_(urlDrive);
  if (!id) return { ok: false, motivo: 'no se pudo extraer el ID de Drive del link' };

  let texto;
  try {
    texto = obtenerTextoFacturaCacheado_(id, cache);
  } catch (err) {
    return { ok: false, motivo: 'error leyendo la factura: ' + err.message };
  }

  if (!texto || texto.trim() === '') {
    return { ok: false, motivo: 'la factura no devolvió texto (OCR vacío)' };
  }

  const cabecera = recortarBloqueEmisor_(texto);
  const parseo = parsearRazonSocial_(cabecera);

  return { ok: true, cabecera: cabecera, razonSocial: parseo.ok ? parseo.razonSocial : null };
}

const MARCADORES_RECEPTOR = [
  /Apellido\s*y\s*Nombre/i,
  /(^|\n)\s*Cliente\s*:/i,
  /(^|\n)\s*Se[ñn]or(es)?\s*:/i,
  /(^|\n)\s*Sr\s*\(?e?s?\)?\s*\.?\s*:/i,
  /Raz[oó]n\s*Social\s*del\s*Cliente/i,
  /Condici[oó]n\s+de\s+venta/i,
  /(^|\n)\s*Cond\.?\s*Vta\.?\s*:/i
];

function recortarBloqueEmisor_(texto) {
  let corte = -1;
  MARCADORES_RECEPTOR.forEach(re => {
    const pos = texto.search(re);
    if (pos > 0 && (corte < 0 || pos < corte)) corte = pos;
  });
  return corte > 0 ? texto.substring(0, corte) : texto;
}

const ETIQUETAS_POSTERIORES = /\s+(Domicilio|Direcci[oó]n|Tel|Tel[eé]fono|Cel|Fax|C\.?U\.?I\.?T|CUIL|IVA|Ing\.?\s*Brutos|Ingresos\s+Brutos|Condici[oó]n|Situaci[oó]n|Fecha|Localidad|Provincia|E-?mail|Correo|Cliente|Punto\s+de\s+Venta|Comp\.?\s*Nro)\b[^\n]*$/i;

function parsearRazonSocial_(cabecera) {
  const m = cabecera.match(/Raz[oó]n\s*Social\s*:[ \t]*([^\n]*)\n?[ \t]*([^\n]*)/i);
  if (!m) return { ok: false, motivo: 'no se encontró "Razón Social:" en la cabecera de la factura' };

  let valor = (m[1] || '').trim();
  if (valor === '') valor = (m[2] || '').trim();
  if (valor === '') return { ok: false, motivo: '"Razón Social:" apareció sin valor legible' };

  valor = valor.replace(ETIQUETAS_POSTERIORES, '').trim();

  const idx = valor.indexOf(':');
  if (idx >= 0) {
    valor = valor.substring(0, idx).replace(/\s+\S+$/, '').trim();
  }

  if (valor === '') return { ok: false, motivo: 'después de "Razón Social:" vino otro rótulo, no un nombre' };

  return { ok: true, razonSocial: valor };
}

/* ══════════════════ COMPARACIÓN DE NOMBRES ══════════════════ */

function mismoConjuntoDePalabras_(nombreA, nombreB) {
  const a = normalizarNombre_(nombreA);
  const b = normalizarNombre_(nombreB);
  if (a.size === 0 || b.size === 0) return false;
  if (a.size !== b.size) return false;

  for (const palabra of a) {
    if (!b.has(palabra)) return false;
  }
  return true;
}

function nombreApareceEn_(tokensNombre, texto) {
  const tokensTexto = Array.from(normalizarNombre_(texto));
  const setTexto = new Set(tokensTexto);

  for (const palabra of tokensNombre) {
    if (setTexto.has(palabra)) continue;
    if (palabra.length >= 4 && tokensTexto.some(t => t.indexOf(palabra) !== -1)) continue;
    return false;
  }
  return true;
}

function normalizarNombre_(nombre) {
  let s = String(nombre == null ? '' : nombre)
    .toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');

  s = s.replace(/&/g, ' Y ');

  s = s.replace(/SOCIEDAD\s+DE\s+RESPONSABILIDAD\s+LIMITADA/g, ' SRL ');
  s = s.replace(/SOCIEDAD\s+POR\s+ACCIONES\s+SIMPLIFICADA/g, ' SAS ');
  s = s.replace(/SOCIEDAD\s+ANONIMA/g, ' SA ');
  s = s.replace(/\bLIMITADA\b/g, ' LTDA ');
  s = s.replace(/\bLTDA\b\s*\.?/g, ' LTDA ');
  s = s.replace(/\bLTD\b\s*\.?/g, ' LTDA ');

  s = s.replace(/\bS\s*\.?\s*A\s*\.?\s*S\s*\.?(?![A-Z0-9])/g, ' SAS ');
  s = s.replace(/\bS\s*\.?\s*C\s*\.?\s*A\s*\.?(?![A-Z0-9])/g, ' SCA ');
  s = s.replace(/\bS\s*\.?\s*R\s*\.?\s*L\s*\.?(?![A-Z0-9])/g, ' SRL ');
  s = s.replace(/\bS\s*\.?\s*A\s*\.?(?![A-Z0-9])/g, ' SA ');
  s = s.replace(/\bS\s*\.?\s*H\s*\.?(?![A-Z0-9])/g, ' SH ');

  s = s.replace(/[^A-Z0-9]+/g, ' ').trim();
  return new Set(s.split(/\s+/).filter(p => p.length > 0));
}

/* ══════════════════ ALERTA DE MONTO ALTO ══════════════════ */

/** true si la fila está en la moneda vigilada y supera el umbral. */
function montoAltoSospechoso_(monto, moneda) {
  if (String(moneda == null ? '' : moneda).trim().toUpperCase() !== MONEDA_MONTO_ALERTA) return false;
  const n = aNumero_(monto);
  return n !== null && n > UMBRAL_MONTO_ALERTA;
}

/** Popup con el detalle de las filas que superan el umbral. No hace nada si no hay ninguna. */
function avisarMontosAltos_(salida) {
  const altos = salida.filter(f =>
    !esEncabezadoDeGrupo_(f) && montoAltoSospechoso_(f[COL.MONTO - 1], f[COL.MONEDA - 1]));

  if (altos.length === 0) return;

  const detalle = altos.map(f =>
    '• ' + f[COL.TASK - 1] + ' — ' + f[COL.BENEF - 1] + ' — ' +
    formatearImporte_(aNumero_(f[COL.MONTO - 1])) + ' ' + MONEDA_MONTO_ALERTA
  ).join('\n');

  SpreadsheetApp.getUi().alert(
    '⚠️ REVISAR MONEDA\n\n' +
    altos.length + ' fila(s) en ' + MONEDA_MONTO_ALERTA + ' superan los ' + UMBRAL_MONTO_ALERTA +
    '. Puede ser una moneda mal cargada (BRL o ARS marcados como ' + MONEDA_MONTO_ALERTA + '):\n\n' +
    detalle + '\n\n' +
    'Están pintadas en rojo en la hoja. Si el monto es correcto, ignorá la marca.'
  );
}

/* ══════════════════ VALIDACIÓN DE MONTO ══════════════════ */

function resultadoValidacionMonto_(task, monto, linkFactura, statusEnBase, clave, ctx, cache) {
  const montoHoja = aNumero_(monto);
  if (montoHoja === null) {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — MONTO vacío o no numérico en la hoja' };
  }

  const tieneFactura = linkFactura && String(linkFactura).trim() !== '';

  if (!tieneFactura && statusEnBase === STATUS_AGRUPACION) {
    const anclas = anclasDeGrupo_(task, clave, ctx);
    return {
      estado: 'REVISION',
      texto: 'NO SE PUDO, REVISION MANUAL — pago agrupado: esta fila no lleva factura propia y su ' +
             'monto individual no figura en ninguna factura.' +
             (anclas.length > 0
               ? '\nEl total se verifica en la fila que sí tiene factura: ' + anclas.map(a => a.task).join(', ')
               : '\nNo se encontró en esta hoja ninguna fila con factura para el mismo CBU y beneficiario.')
    };
  }

  if (!tieneFactura) {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — la fila no tiene link de factura' };
  }

  const id = extraerIdDeDrive_(linkFactura);
  if (!id) {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — no se pudo extraer el ID de Drive del link' };
  }

  let texto;
  try {
    texto = obtenerTextoFacturaCacheado_(id, cache);
  } catch (err) {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — error leyendo la factura: ' + err.message };
  }
  if (!texto || texto.trim() === '') {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — la factura no devolvió texto (OCR vacío)' };
  }

  const imp = parsearImporteFactura_(texto);
  if (!imp.ok) {
    return { estado: 'REVISION', texto: 'NO SE PUDO, REVISION MANUAL — ' + imp.motivo };
  }

  const solo = redondear2_(montoHoja);
  const factura = imp.comparable;
  const cabecera =
    'factura: ' + formatearImporte_(factura) +
    ' (rótulo: ' + imp.rotulo + ')' +
    (imp.otrosTributos !== 0 ? '\nmenos Otros Tributos: ' + formatearImporte_(imp.otrosTributos) : '');

  if (solo === factura) {
    return { estado: 'OK', comparable: factura,
             texto: '✅ COINCIDE — hoja: ' + formatearImporte_(solo) + '\n' + cabecera };
  }

  const satelites = (ctx.satelites[clave] || []).filter(s => s.task !== String(task));
  if (satelites.length > 0) {
    let suma = solo;
    satelites.forEach(s => { suma = redondear2_(suma + s.monto); });

    const tareas = [String(task)].concat(satelites.map(s => s.task)).join(', ');

    if (suma === factura) {
      return {
        estado: 'OK',
        comparable: factura,
        texto: '✅ COINCIDE COMO PAGO AGRUPADO — la factura cubre esta fila más ' +
               satelites.length + ' fila(s) más.\n' +
               'suma: ' + formatearImporte_(solo) + ' + ' +
               satelites.map(s => formatearImporte_(s.monto)).join(' + ') + ' = ' + formatearImporte_(suma) + '\n' +
               'tareas sumadas: ' + tareas + '\n' + cabecera
      };
    }

    return {
      estado: 'MISMATCH',
      comparable: factura,
      texto: '❌ NO COINCIDE — ' + cabecera +
             '\nhoja (fila sola): ' + formatearImporte_(solo) +
             '\nhoja (con agrupación ' + tareas + '): ' + formatearImporte_(suma) +
             '\ndiferencia contra la suma: ' + formatearImporte_(redondear2_(suma - factura))
    };
  }

  // Si la diferencia es exactamente una alícuota de IVA, lo más probable es que se haya leído el
  // subtotal y no el total. No se devuelve comparable: escribir ese número en MONTO FACTURA sería
  // guardar un importe que sabemos mal, y la orden de pago saldría por el neto.
  const alicuota = ratioEsAlicuotaIva_(solo, factura);
  if (alicuota !== null) {
    return {
      estado: 'MISMATCH',
      texto: '❌ NO COINCIDE Y HUELE A SUBTOTAL — hoja ÷ factura = 1 + ' + alicuota + '%, ' +
             'que es exactamente la alícuota de IVA.\n' +
             'Lo más probable es que se haya leído el subtotal en vez del total.\n' +
             cabecera + '\nhoja: ' + formatearImporte_(solo) +
             '\n\nNO se guardó nada en MONTO FACTURA. Revisá la factura a mano.'
    };
  }

  return {
    estado: 'MISMATCH',
    comparable: factura,
    texto: '❌ NO COINCIDE — hoja: ' + formatearImporte_(solo) + '\n' + cabecera +
           '\ndiferencia: ' + formatearImporte_(redondear2_(solo - factura))
  };
}

/**
 * Rótulos de total. El importe tiene que estar en la MISMA línea que el rótulo: por eso [ \t]* y
 * no \s*. Con \s* el regex cruzaba saltos de línea y en las facturas con los totales apilados
 *     Subtotal / IVA 2,5 % / Total:
 *     $ 195.121,95   $ 4.878,05   $ 200.000,00
 * agarraba el PRIMER número que venía después del rótulo, que es el neto de IVA. Eso hacía que
 * "Total" devolviera 195.121,95 en una factura de 200.000,00: cuatro de nueve filas de un día real
 * salieron con el subtotal en vez del total.
 * Cuando ningún rótulo matchea en su línea, se prueba totalPorLineaDeImportes_.
 */
const ROTULOS_TOTAL = [
  { nombre: 'Importe Total', re: /Importe\s+Total\s*:?[ \t]*\$?[ \t]*([\d.,]+)/i },
  { nombre: 'Total a Pagar', re: /Total\s+a\s+Pagar\s*:?[ \t]*\$?[ \t]*([\d.,]+)/i },
  { nombre: 'TOTAL', re: /(?:^|\n)[ \t]*TOTAL[ \t]*:?[ \t]*\$?[ \t]*([\d.,]+)/i }
];

/** Ancla del bloque de totales de ARCA/AFIP. */
const RE_TOTAL_ANCLA = /(?:Importe\s+Total|Total\s+a\s+Pagar)[ \t]*:?[ \t]*\$?/i;

/**
 * Bloque de totales con los rótulos separados de los importes. Es el layout estándar del PDF de
 * ARCA/AFIP, y el OCR lo devuelve así:
 *
 *     Subtotal: $
 *     Importe Otros Tributos: $ Importe Total: $
 *     76500,00 0,00
 *     76500,00
 *
 * Los rótulos vienen primero, todos juntos, y los importes después EN EL MISMO ORDEN, repartidos
 * en las líneas que haga falta. Ni el regex por rótulo (no hay número en la línea) ni
 * totalPorLineaDeImportes_ (los importes están en dos líneas distintas) lo agarran.
 *
 * Se toman los importes que siguen al rótulo de total y se prueba, para k entre 3 y 10, si los
 * primeros k-1 suman el k-ésimo. Esa verificación aritmética es lo único que valida el resultado:
 * si ninguna k cierra, se devuelve null y la fila va a revisión.
 *
 * Cantidad de importes según el tipo de factura de ARCA:
 *   Factura C: 3 (Subtotal, Otros Tributos, Total).
 *   Factura A: 8 (Neto Gravado, IVA 27%, IVA 21%, IVA 10,5%, IVA 5%, IVA 2,5%, Otros Tributos,
 *              Total). Con el tope anterior de 6 el total de las A nunca entraba en el rango y
 *              todas las filas tipo A quedaban sin MONTO FACTURA.
 * Con la secuencia de la A [N, 0, 0, 0, 0, I, 0, T], ningún k menor a 8 cierra la suma, así que
 * ampliar el tope no produce falsos positivos en ninguno de los dos layouts.
 *
 * "Importe Otros Tributos" es siempre el rótulo inmediatamente anterior a "Importe Total" en este
 * layout, así que su importe es el anteúltimo de la serie. Solo se usa si ese rótulo aparece de
 * verdad en los 200 caracteres previos al ancla.
 */
function totalPorBloqueDeEtiquetas_(texto) {
  const m = texto.match(RE_TOTAL_ANCLA);
  if (!m) return null;

  const inicioAncla = texto.indexOf(m[0]);
  const cola = texto.substring(inicioAncla + m[0].length, inicioAncla + m[0].length + 400);
  const nums = (cola.match(/[\d][\d.,]*/g) || []).map(x => aNumero_(x)).filter(n => n !== null);

  for (let k = 3; k <= 10 && k <= nums.length; k++) {
    let suma = 0;
    for (let j = 0; j < k - 1; j++) suma = redondear2_(suma + nums[j]);
    const ultimo = redondear2_(nums[k - 1]);
    if (Math.abs(suma - ultimo) > 0.02) continue;

    const antes = texto.substring(Math.max(0, inicioAncla - 200), inicioAncla);
    const hayOtros = /Importe\s+Otros\s+Tributos/i.test(antes);
    return { total: ultimo, otros: hayOtros ? redondear2_(nums[k - 2]) : null };
  }
  return null;
}

/**
 * Rescate para el layout de totales apilados: una línea con tres o más importes donde el ÚLTIMO
 * es la suma de los anteriores. Lo que identifica al total es esa comprobación aritmética, no la
 * posición: 195.121,95 + 4.878,05 = 200.000,00 confirma cuál de los tres es el total.
 * Se recorre de abajo hacia arriba porque los totales están al pie. Si ninguna línea cierra, se
 * devuelve null y la fila va a revisión: es preferible a escribir un importe dudoso.
 */
function totalPorLineaDeImportes_(texto) {
  const lineas = String(texto).split('\n');
  for (let i = lineas.length - 1; i >= 0; i--) {
    const nums = (lineas[i].match(/[\d][\d.,]*/g) || [])
      .map(x => aNumero_(x))
      .filter(n => n !== null && n > 0);
    if (nums.length < 3) continue;

    const ultimo = redondear2_(nums[nums.length - 1]);
    let suma = 0;
    for (let j = 0; j < nums.length - 1; j++) suma = redondear2_(suma + nums[j]);
    if (Math.abs(suma - ultimo) <= 0.02) return ultimo;
  }
  return null;
}

/** Alícuotas de IVA vigentes, para detectar cuando se leyó el neto en vez del total. */
const ALICUOTAS_IVA = [0.025, 0.05, 0.105, 0.21, 0.27];

/**
 * Si hoja ÷ factura da exactamente 1 + una alícuota de IVA, es casi seguro que se leyó el subtotal
 * y no el total. Devuelve la alícuota en porcentaje, o null.
 */
function ratioEsAlicuotaIva_(montoHoja, montoFactura) {
  if (!montoFactura || montoFactura <= 0) return null;
  const r = montoHoja / montoFactura;
  for (let i = 0; i < ALICUOTAS_IVA.length; i++) {
    if (Math.abs(r - (1 + ALICUOTAS_IVA[i])) < 0.0005) {
      return Math.round(ALICUOTAS_IVA[i] * 1000) / 10;
    }
  }
  return null;
}

function parsearImporteFactura_(texto) {
  let total = null, rotulo = null;

  for (let i = 0; i < ROTULOS_TOTAL.length && total === null; i++) {
    total = primerImporte_(texto, ROTULOS_TOTAL[i].re);
    if (total !== null) rotulo = ROTULOS_TOTAL[i].nombre;
  }

  // Otros Tributos leído del bloque, cuando el bloque es el que resolvió el total.
  let otrosDelBloque = null;

  if (total === null) {
    const bloque = totalPorBloqueDeEtiquetas_(texto);
    if (bloque) {
      total = bloque.total;
      otrosDelBloque = bloque.otros;
      rotulo = 'bloque de totales (verificado: los importes previos suman el total)';
    }
  }

  if (total === null) {
    total = totalPorLineaDeImportes_(texto);
    if (total !== null) rotulo = 'línea de importes (verificado: los previos suman el último)';
  }

  if (total === null) {
    return { ok: false, motivo: 'no se encontró ningún rótulo de total en la factura' };
  }

  const otros = primerImporte_(texto, /Importe\s+Otros\s+Tributos\s*:?[ \t]*\$?[ \t]*([\d.,]+)/i);
  const otrosNum = (otros !== null) ? otros : (otrosDelBloque !== null ? otrosDelBloque : 0);
  return {
    ok: true,
    rotulo: rotulo,
    total: redondear2_(total),
    otrosTributos: redondear2_(otrosNum),
    comparable: redondear2_(total - otrosNum)
  };
}

function primerImporte_(texto, regex) {
  const m = texto.match(regex);
  return m ? aNumero_(m[1]) : null;
}

function aNumero_(valor) {
  if (typeof valor === 'number') return isNaN(valor) ? null : valor;

  let s = String(valor == null ? '' : valor).replace(/[^\d.,-]/g, '');
  if (s === '' || !/\d/.test(s)) return null;

  const ultSep = Math.max(s.lastIndexOf('.'), s.lastIndexOf(','));
  if (ultSep >= 0) {
    const decimales = s.length - ultSep - 1;
    if (decimales === 1 || decimales === 2) {
      s = s.substring(0, ultSep).replace(/[.,]/g, '') + '.' + s.substring(ultSep + 1);
    } else {
      s = s.replace(/[.,]/g, '');
    }
  }

  const n = Number(s);
  return isNaN(n) ? null : n;
}

function redondear2_(n) {
  return Math.round(n * 100) / 100;
}

function formatearImporte_(n) {
  return n.toFixed(2).replace('.', ',');
}

function mapaStatusPorTask_() {
  try {
    const baseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BASE_SHEET_NAME);
    if (!baseSheet) return {};
    const data = baseSheet.getDataRange().getValues();
    const col = mapearColumnas_(data[0]);
    const mapa = {};
    for (let i = 1; i < data.length; i++) {
      const t = data[i][col.task_name];
      if (t) mapa[String(t)] = data[i][col.status_name];
    }
    return mapa;
  } catch (e) {
    return {};
  }
}
