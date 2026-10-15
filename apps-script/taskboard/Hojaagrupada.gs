/**
 * HOJA AGRUPADA POR BENEFICIARIO (hoja satélite del día) + HOJA DE APROBACIÓN
 *
 * Genera una hoja "AGRUP DD-MM-YYYY" a partir de la hoja del día "DD-MM-YYYY", con una fila por
 * grupo, los montos sumados, los IDs de tarea unificados y los links de comprobante uno debajo
 * del otro en la misma celda.
 *
 * Correrla de nuevo sobre un día ya generado NO es destructivo: funciona como "actualizar".
 * Vuelve a leer la hoja del día (montos, IDs y links frescos) y reescribe CHECK, ESTADO y
 * COMENTARIO tal como los tenías.
 *
 * Columnas (16):
 * TIPO | BENEFICIARIO | SUBGRUPO | N° TX | TRANSACCIONES (IDs) | MONTO | MONEDA | CHECK |
 * ESTADO | ORIGEN | COMENTARIO | COMENTARIO ANTERIOR | MONTO A PAGAR | LINKS | PDF | COMPROBANTE
 *
 * ── QUÉ ES CADA COSA ────────────────────────────────────────────────────────────────────────
 *
 * CHECK        Tuyo. Significa "revisé esta fila", nada más. NO significa que la transferencia
 *              salió, y por eso ya no arrastra nada al día siguiente.
 *
 * ESTADO       Desplegable cerrado (ESTADOS_AGRUP, definido en Pagoslb.gs). Nace VACÍO: el script
 *              no afirma nada por vos. Es el dato que sí tiene consecuencias:
 *                (vacío)          → todavía no lo cargaste en el banco. No llega a APROBACIÓN.
 *                CARGADO EN BANCO → lo cargaste, espera autorización. Aparece en APROBACIÓN y,
 *                                   si es A o C, en el archivo de carga a Galicia (CargaGalicia.gs).
 *                TRANSFERIDO      → la plata salió. Mañana la fila queda marcada "PAGADO EL" en
 *                                   la hoja del día y sale de la agrupada.
 *                COMPROBANTE CARGADO → se puso solo, porque la columna COMPROBANTE ya tiene el PDF
 *                                   final armado. Cuenta como pagado igual que TRANSFERIDO.
 *                                   No pisa RECHAZADO ni CANCELADO.
 *                CON PROBLEMA / RECHAZADO / CANCELADO → no sacan la fila de ningún lado, pero
 *                                   tampoco llegan a APROBACIÓN.
 *
 * ORIGEN       Lo pone el script, no se edita:
 *                ⭐ NUEVO          → el grupo no estaba en la corrida anterior de HOY.
 *                ↻ DE AYER        → alguna de sus tareas ya estaba en la agrupada anterior.
 *                ↻ EN BANCO dd-MM → sus tareas estaban en CARGADO EN BANCO (o COMPROBANTE
 *                                   CARGADO) en la agrupada anterior. Va en fila aparte: ver
 *                                   SEPARACIÓN POR ESTADO ANTERIOR más abajo.
 *
 * SUBGRUPO     Viene de la columna SUBGRUPO de la hoja del día. Dos filas del mismo beneficiario
 *              con etiquetas distintas van a filas distintas acá. Es tu forma de forzar una
 *              separación a mano.
 *              En A y C muestra además el CBU / DATO DE PAGO de la hoja del día, delante de la
 *              etiqueta: "0070...123 | etiqueta". Dos CBU distintos del mismo beneficiario van en
 *              filas distintas (AGRUP_SEPARAR_POR_CBU).
 *
 * COMENTARIO   Tuyo, texto libre. En las filas Paypal, PaypalAprobaciones.gs agrega al final
 *              "| PAYPAL: <estado>" y lo reescribe en cada corrida.
 *
 * COMENTARIO   Lo pone el script: el comentario que ese grupo tuvo en una agrupada anterior.
 * ANTERIOR     Sale de la hoja del día (columna COMENTARIO ANTERIOR), no de una comparación
 *              grupo contra grupo. Así queda consistente con la partición: si dos tareas del
 *              mismo beneficiario traen comentarios distintos, van a filas distintas y cada una
 *              muestra el suyo.
 *
 * MONTO A      Tuyo, solo tipo A: el neto de retenciones que se transfiere. Lo usa y lo preserva
 * PAGAR        CargaGalicia.gs (conMontoAPagar_). Esta hoja la crea vacía.
 *
 * ── SEPARACIÓN POR ESTADO ANTERIOR ──────────────────────────────────────────────────────────
 * Si una tarea estaba en la agrupada anterior en un estado de estadosQueSeparan_() (CARGADO EN
 * BANCO o COMPROBANTE CARGADO), hoy va en su propia fila y NO se suma con operaciones nuevas del
 * mismo beneficiario. Esa plata ya estaba cargada en el banco: sumarle una tarea nueva arma un
 * monto que no corresponde a ninguna transferencia real.
 * Las tareas que ayer compartían fila siguen juntas hoy: la separación respeta la fila de ayer
 * (la clave es el conjunto de IDs de esa fila, no su número de fila, así que ordenar o filtrar la
 * agrupada de ayer no la rompe).
 * En la práctica solo pasa con CARGADO EN BANCO: COMPROBANTE CARGADO y TRANSFERIDO cuentan como
 * pago, esas tareas quedan "PAGADO EL" en la hoja del día y no llegan a la agrupada.
 * Si la transferencia de ayer salió, lo que corresponde es ponerle TRANSFERIDO: mañana deja de
 * aparecer.
 * Estas filas nacen con ESTADO vacío y "Marcar todo como CARGADO EN BANCO" NO las marca: ya
 * están en el banco, y marcarlas las mandaría de nuevo a APROBACIÓN, donde se podrían autorizar
 * dos veces. Si una tarea se canceló después de la carga en el banco, la separación no aplica:
 * esa carga ya no existe.
 *
 * ── CANCELADAS ──────────────────────────────────────────────────────────────────────────────
 * Las tareas canceladas en un día anterior (Cancelaciones.gs) llevan "CANCELADO dd-MM · ..." al
 * principio de COMENTARIO y van en fila aparte del resto del beneficiario. Las canceladas 2 veces o
 * más dicen "NO REPAGAR · AVISAR EN APPSHEET", van en rojo, y "Marcar todo como CARGADO EN BANCO"
 * no las marca.
 *
 * ── DECISIONES TOMADAS ──────────────────────────────────────────────────────────────────────
 *  - NO se convierten monedas. Dos monedas del mismo beneficiario = dos filas.
 *  - NO se filtra por CHECK ni se excluyen las filas "RETIRADO DE LA BASE".
 *  - SÍ se excluyen las filas marcadas "PAGADO EL" en ESTADO de la hoja del día: ya se
 *    transfirieron un día anterior. Para incluirlas, poner AGRUP_EXCLUIR_YA_PAGADOS en false.
 *  - En A y C el CBU / dato de pago aparece en SUBGRUPO y dos cuentas distintas del mismo
 *    beneficiario son dos filas. En el resto de los tipos el dato de pago NO aparece y dos cuentas
 *    del mismo beneficiario se suman en una sola fila.
 *  - Paypal y Payoneer entran, aunque su summary no trae "Beneficiary": esas filas se agrupan y
 *    se muestran por su dato de pago (el mail).
 *  - Criptomoneda no aparece: se excluye antes, al armar la hoja del día.
 *  - Cualquier beneficiario que CONTENGA "Backlinks Global LLC" se excluye por completo.
 *  - CHECK, ESTADO y COMENTARIO se preservan al regenerar por TIPO + BENEFICIARIO + MONEDA +
 *    COMENTARIO ANTERIOR + SUBGRUPO + separación por estado anterior + cancelación + CBU (A y C).
 *    La composición del grupo NO entra en la clave.
 *    ⚠️ CONSECUENCIA ASUMIDA: si cambia el comentario arrastrado o la etiqueta de SUBGRUPO, la
 *    clave cambia y esa fila pierde el tilde, el estado y el comentario que le hayas puesto.
 *    ⚠️ BORDE: dos grupos sin BENEFICIARIO y sin dato de pago comparten clave y uno pisa al otro.
 *  - Si COMENTARIO ANTERIOR contiene "COMPROBANTE CARGADO" y viene del día anterior, la fila se
 *    pinta de verde. Si viene de más atrás, significa que el comprobante nunca se cargó de
 *    verdad: la fila se pinta de naranja, deja nota y salta un popup al generar.
 *
 * DEPENDE de Pagoslb.gs: formatearNombreHoja_, normalizarNombre_, aNumero_, redondear2_,
 * formatearImporte_, colorFondoDeFila_, textoPlano_, indicesPorHeader_, claveGrupoAgrup_,
 * hojasAnterioresA_, parsearNombreHoja_, parsearNombreHojaAgrup_, recortarGrilla_, ORDEN_TIPOS,
 * HEADERS, COL, MARCA_PAGADO_PREFIJO, EST_* y ESTADOS_AGRUP.
 */

const AGRUP_PREFIJO = 'AGRUP ';

const AGRUP_HEADERS = [
  'TIPO', 'BENEFICIARIO', 'SUBGRUPO', 'N° TX', 'TRANSACCIONES (IDs)', 'MONTO', 'MONEDA',
  'CHECK', 'ESTADO', 'ORIGEN', 'COMENTARIO', 'COMENTARIO ANTERIOR', 'MONTO A PAGAR', 'LINKS', 'PDF',
  'COMPROBANTE'
];

/** Número de columna (1-based) de cada campo de la hoja agrupada. */
const AGRUP_COL = {
  TIPO: 1, BENEF: 2, SUBGRUPO: 3, NTX: 4, IDS: 5, MONTO: 6, MONEDA: 7,
  CHECK: 8, ESTADO: 9, ORIGEN: 10, COMENT: 11, COMENT_ANT: 12, PAGAR: 13, LINKS: 14, PDF: 15,
  COMPROBANTE: 16
};

/**
 * Tipos que no entran en la hoja agrupada. Vacío: Paypal ya entra, porque las filas sin
 * BENEFICIARIO se agrupan y se muestran por su dato de pago (el mail).
 * Criptomoneda no figura acá porque se excluye antes, en la hoja del día.
 */
const AGRUP_TIPOS_EXCLUIDOS = [];
const AGRUP_BENEFICIARIOS_EXCLUIDOS = ['Backlinks Global LLC'];
const AGRUP_EXCLUIR_YA_PAGADOS = true;

/** Texto que, si aparece en COMENTARIO ANTERIOR, marca la fila como comprobante ya cargado. */
const AGRUP_TEXTO_COMPROBANTE = 'COMPROBANTE CARGADO';
const AGRUP_COLOR_COMPROBANTE = '#0ce847'; // verde: comprobante cargado ayer, todo en orden
const AGRUP_COLOR_REPETIDO    = '#f6b26b'; // naranja: el mismo comentario aparece por 2° día

const AGRUP_MARCA_NUEVO   = '⭐ NUEVO';
const AGRUP_MARCA_DE_AYER = '↻ DE AYER';

/** Hoja que ve la persona que autoriza los pagos. Nombre fijo: el link no cambia nunca. */
const APROB_SHEET_NAME = 'APROBACIÓN';
const APROB_HEADERS = ['TIPO', 'TRANSACCIONES', 'BENEFICIARIO', 'MONTO', 'MONEDA', 'LINK FACTURA'];

/** Carpeta de Drive donde se guardan los PDF. */
const PDF_CARPETA_ID = '10ghHh2xyOKFh4nPAC1f35bd2GmjCwvZ-';

/**
 * ID del archivo del logo de Eleven en Drive (PNG o JPG). Si queda vacío o el archivo no se puede
 * leer, se usa el logo embebido en base64 y, si tampoco, el PDF sale sin logo.
 */
const PDF_LOGO_FILE_ID = '';

const PDF_FUENTE = 'Calibri';
const PDF_LARGO_MAX_NOMBRE = 180;

/**
 * Estados de la agrupada ANTERIOR que obligan a poner sus tareas en fila aparte hoy.
 * Es una función y no una constante porque EST_* se definen en Pagoslb.gs: una constante de
 * nivel superior que los use depende del orden en que Apps Script carga los archivos.
 * Si querés sumar CON PROBLEMA (para que una operación trabada no arrastre a una nueva del mismo
 * beneficiario), agregala acá.
 */
function estadosQueSeparan_() {
  return [EST_CARGADO, EST_COMPROBANTE];
}

/**
 * Para los tipos A y C (TIPOS_VALIDABLES), dos CBU distintos del mismo beneficiario van en filas
 * distintas. Una fila de estos tipos es UNA transferencia de Galicia, y una transferencia va a una
 * sola cuenta: sumar dos cuentas arma un monto que no corresponde a nada, y el panel de lote no
 * podría reconocer el comprobante por CBU.
 * Con false, se suman como antes y la columna SUBGRUPO muestra todos los CBU del grupo.
 * ⚠️ Un alias y el CBU de la MISMA cuenta se ven como cuentas distintas: no hay forma de saber que
 *    son la misma sin consultar el banco. Si pasa, se ven dos filas y las juntás a mano.
 */
const AGRUP_SEPARAR_POR_CBU = true;

/** Tipos cuya columna SUBGRUPO muestra el CBU / DATO DE PAGO de la hoja del día. */
function muestraCbu_(tipo) {
  return TIPOS_VALIDABLES.indexOf(tipo) !== -1;
}

/** CBU o alias normalizado para comparar: mayúsculas, sin espacios, puntos ni guiones. */
function claveCbu_(tipo, dato) {
  if (!AGRUP_SEPARAR_POR_CBU || !muestraCbu_(tipo)) return '';
  return String(dato == null ? '' : dato).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Texto de la columna SUBGRUPO de la agrupada. En A y C: el CBU / dato de pago, y si además
 * pusiste una etiqueta en SUBGRUPO de la hoja del día, va después de " | ". En el resto: la
 * etiqueta sola, como siempre.
 */
function textoSubgrupo_(g) {
  if (!muestraCbu_(g.tipo)) return g.subgrupo;
  return g.datos.concat(g.subgrupo ? [g.subgrupo] : []).join(' | ');
}

/** Texto corto para la columna ORIGEN. */
function etiquetaEstadoAnterior_(estado, fecha) {
  const corto = estado === EST_CARGADO ? 'EN BANCO'
              : estado === EST_COMPROBANTE ? 'COMPROBANTE'
              : estado;
  return '↻ ' + corto + ' ' + fecha;
}

/* ══════════════════ GENERACIÓN DE LA HOJA ══════════════════ */

/** Ítem de menú "Actualizar hoja agrupada del día". Es la misma función: regenerar ES actualizar. */
function actualizarHojaAgrupadaDelDia() {
  generarHojaAgrupadaDelDia();
}

function generarHojaAgrupadaDelDia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreDia = formatearNombreHoja_(new Date());
  const hojaDia = ss.getSheetByName(nombreDia);

  if (!hojaDia) {
    SpreadsheetApp.getUi().alert('No existe la hoja del día "' + nombreDia + '". Generala primero.');
    return;
  }

  // Cancelaciones y agrupada anterior se leen ANTES de agrupar: definen qué tareas van aparte.
  sincronizarCancelaciones_(ss, nombreDia);
  const canc = estadoCancelaciones_(ss);
  const anterior = agrupAnterior_(ss, nombreDia, canc);

  const filas = leerFilasDeHojaDelDia_(hojaDia);
  // porTask: etiqueta de SUBGRUPO y CBU de cada tarea. leerEstadoPrevioAgrup_ los necesita porque
  // la columna SUBGRUPO de la agrupada ya no muestra solo la etiqueta (en A y C lleva el CBU).
  const porTask = {};
  filas.forEach(f => {
    f.arrastre = anterior.arrastre[f.task] || null;
    f.canc = canc[f.task] || null;
    f.marcaCanc = marcaCancelacion_(f.canc);
    f.cbuClave = claveCbu_(f.tipo, f.dato);
    porTask[f.task] = { subgrupo: f.subgrupo, cbuClave: f.cbuClave };
  });

  const grupos = agruparPorBeneficiario_(filas);
  if (grupos.length === 0) {
    SpreadsheetApp.getUi().alert('No hay filas para agrupar en la hoja "' + nombreDia + '".');
    return;
  }

  const nombreAgrup = AGRUP_PREFIJO + nombreDia;
  const existente = ss.getSheetByName(nombreAgrup);

  const previos = existente ? leerEstadoPrevioAgrup_(existente, anterior, canc, porTask) : null;

  grupos.forEach(g => {
    const clave = claveGrupoAgrupExt_(g.tipo, g.beneficiario, g.moneda, g.comentarioAnterior,
                                      g.subgrupo, separadorGrupo_(g.arrastre, g.marcaCanc, g.cbuClave));

    const prev = previos ? previos[clave] : null;
    if (prev) {
      g.check = prev.check;
      g.estado = prev.estado;
      g.comentario = prev.comentario;
      g.comprobante = prev.comprobante;
    } else {
      // Fila nueva: nace SIN estado. Pasa a CARGADO EN BANCO recién cuando corrés a mano
      // "Marcar todo como cargado en banco", después de haberlos cargado de verdad.
      // Así la hoja de APROBACIÓN refleja algo que afirmaste vos, no algo que asumió el script.
      g.estado = '';
      if (previos) g.origen = AGRUP_MARCA_NUEVO;
    }

    // Si la fila ya tiene el comprobante final armado, el estado pasa solo a COMPROBANTE CARGADO:
    // el archivo en Drive es prueba de que el circuito terminó. No se pisan RECHAZADO ni CANCELADO,
    // porque ahí un comprobante presente es una contradicción que hay que ver, no tapar.
    if (g.comprobante &&
        g.estado !== EST_RECHAZADO && g.estado !== EST_CANCELADO) {
      g.estado = EST_COMPROBANTE;
    }

    // Todas las tareas del grupo tienen la misma marca: la marca es parte de la clave de grupo.
    if (g.marcaCanc) g.comentario = ponerMarcaCancelado_(g.comentario, g.marcaCanc);

    // El estado de ayer manda sobre NUEVO: es lo que explica por qué la fila está separada.
    // DE AYER se decide por tareas y no por beneficiario: una operación nueva de un beneficiario
    // que ayer también cobraba no es "de ayer".
    if (g.arrastre) {
      g.origen = g.arrastre.etiqueta;
    } else if (g.origen === '' && g.ids.some(t => anterior.ids[t])) {
      g.origen = AGRUP_MARCA_DE_AYER;
    }
    g.repetido = comentarioRepetido_(g.comentarioAnterior, anterior.fecha);
  });

  // Montos de A/C cargados a mano en la agrupada: al regenerar se pisan con MONTO FACTURA de la
  // hoja del día. Se avisa ANTES de borrar nada; con NO, la agrupada queda exactamente como estaba.
  if (existente) {
    const pisados = montosQueSePisan_(existente, grupos);
    if (pisados.length > 0 && !confirmarMontosPisados_(pisados)) return;
  }

  if (existente) ss.deleteSheet(existente);

  escribirHojaAgrupada_(ss.insertSheet(nombreAgrup), grupos, nombreDia);
  // PaypalAprobaciones.gs: solo lee Gmail, no manda nada. Si falla, la agrupada queda igual.
  try { refrescarColumnaPaypal_(ss.getSheetByName(nombreAgrup)); }
  catch (e) { ss.toast('No se pudo leer Gmail para el estado de PayPal: ' + e.message, 'PayPal', 10); }
  repintarCanceladasWise_();   // WiseCanceladas.gs: el magenta va antes de los popups, que bloquean
  avisarSeparadosPorEstadoAnterior_(ss, grupos);
  avisarMontosSinFactura_(grupos);
  avisarComprobantesRepetidos_(grupos);
}

/**
 * Grupos A/C cuyo MONTO actual en la agrupada no coincide con el que va a calcular el script
 * (vacío o distinto). Casi siempre es un monto tipeado a mano en la agrupada, que no vive en
 * MONTO FACTURA de la hoja del día y por eso se pierde al regenerar.
 * Se compara por conjunto de IDs. Si la fila vieja no tiene un grupo nuevo con los mismos IDs
 * (cambió la composición), se toman los grupos nuevos donde quedaron sus tareas: si alguno sale
 * sin monto, se avisa igual. Se avisa de más antes que de menos.
 */
function montosQueSePisan_(hoja, grupos) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];

  const nuevoPorIds = {};
  const nuevoPorTask = {};
  grupos.forEach(g => {
    nuevoPorIds[g.ids.slice().sort().join(',')] = g;
    g.ids.forEach(t => { nuevoPorTask[t] = g; });
  });

  const res = [];
  hoja.getRange(2, 1, ultima - 1, AGRUP_HEADERS.length).getValues().forEach(v => {
    if (TIPOS_VALIDABLES.indexOf(String(v[AGRUP_COL.TIPO - 1]).trim()) === -1) return;
    const ids = String(v[AGRUP_COL.IDS - 1] || '').split(',').map(s => s.trim()).filter(String);
    if (ids.length === 0) return;                         // banda o resumen
    const viejo = aNumero_(v[AGRUP_COL.MONTO - 1]);
    if (viejo === null) return;                           // no había nada que perder

    let nuevo;
    const exacto = nuevoPorIds[ids.slice().sort().join(',')];
    if (exacto) {
      nuevo = exacto.monto;
    } else {
      const destinos = ids.map(t => nuevoPorTask[t]).filter(Boolean);
      if (destinos.length === 0) return;                  // las tareas ya no están hoy
      nuevo = destinos.some(g => g.monto === null) ? null : 'recompuesto';
    }
    if (nuevo === 'recompuesto') return;
    if (nuevo !== null && redondear2_(Math.abs(nuevo - viejo)) <= 0.01) return;

    res.push({
      benef: String(v[AGRUP_COL.BENEF - 1] || '(sin beneficiario)'),
      ids: ids,
      moneda: String(v[AGRUP_COL.MONEDA - 1] || ''),
      viejo: viejo,
      nuevo: nuevo
    });
  });
  return res;
}

/** Popup SÍ/NO con los montos que se van a pisar. true = seguir con la actualización. */
function confirmarMontosPisados_(pisados) {
  const ui = SpreadsheetApp.getUi();
  const detalle = pisados.map(p =>
    '• ' + p.benef + ' (' + p.ids.join(', ') + ')\n' +
    '   en la agrupada: ' + formatearImporte_(p.viejo) + ' ' + p.moneda +
    '  →  queda: ' + (p.nuevo === null ? 'VACÍO' : formatearImporte_(p.nuevo) + ' ' + p.moneda)
  ).join('\n');
  const resp = ui.alert(
    '⚠️ SE VAN A PISAR MONTOS DE LA AGRUPADA\n\n' +
    pisados.length + ' grupo(s) A/C tienen en la agrupada un monto que no sale de MONTO FACTURA ' +
    'de la hoja del día (probablemente cargado a mano):\n\n' + detalle + '\n\n' +
    'Para no perderlos: NO, cargalos en MONTO FACTURA de la hoja del día y volvé a actualizar.\n\n' +
    '¿Actualizar igual y pisarlos?',
    ui.ButtonSet.YES_NO
  );
  return resp === ui.Button.YES;
}

/**
 * Datos de la última hoja agrupada anterior a hoy:
 *   fecha     → "dd-MM" de esa hoja (para detectar comentarios que se repiten más de un día);
 *   ids       → set de todas las tareas que tenía (para marcar "↻ DE AYER");
 *   arrastre  → TASK → { clave, estado, fecha, etiqueta } para las tareas cuyo grupo estaba en un
 *               estado de estadosQueSeparan_(). La clave es el conjunto ordenado de IDs de la fila
 *               de ayer: así las tareas que ayer iban juntas siguen juntas, y ordenar o filtrar la
 *               hoja de ayer no cambia nada.
 *               Si la tarea se canceló el mismo día de esa agrupada o después (canc, de
 *               Cancelaciones.gs), no hay arrastre: la carga en el banco ya no existe.
 */
function agrupAnterior_(ss, nombreDia, canc) {
  const res = { fecha: null, ids: {}, arrastre: {} };
  const hoy = parsearNombreHoja_(nombreDia);
  if (!hoy) return res;
  canc = canc || {};

  const anteriores = hojasAnterioresA_(ss, hoy, h => parsearNombreHojaAgrup_(h.getName()));
  if (anteriores.length === 0) return res;

  const o = anteriores[0];
  res.fecha = Utilities.formatDate(o.fecha, ss.getSpreadsheetTimeZone(), 'dd-MM');
  const ymd = ymdDeNombreAgrup_(o.hoja.getName());

  const ultimaFila = o.hoja.getLastRow();
  if (ultimaFila < 2) return res;

  const ancho = Math.max(o.hoja.getLastColumn(), 1);
  const hdr = o.hoja.getRange(1, 1, 1, ancho).getValues()[0]
    .map(h => String(h == null ? '' : h).trim().toUpperCase());

  const iIds = hdr.indexOf('TRANSACCIONES (IDS)');
  const iEstado = hdr.indexOf('ESTADO');
  if (iIds < 0) return res;

  const separan = estadosQueSeparan_();

  o.hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues().forEach(v => {
    const ids = String(v[iIds] == null ? '' : v[iIds]).split(',').map(s => s.trim()).filter(String);
    if (ids.length === 0) return;   // banda o fila de resumen
    ids.forEach(t => { res.ids[t] = true; });

    const est = iEstado >= 0 ? String(v[iEstado] == null ? '' : v[iEstado]).trim().toUpperCase() : '';
    if (separan.indexOf(est) === -1) return;

    const info = {
      clave: 'ANT ' + res.fecha + ' ' + ids.slice().sort().join(','),
      estado: est,
      fecha: res.fecha,
      etiqueta: etiquetaEstadoAnterior_(est, res.fecha)
    };
    ids.forEach(t => {
      if (res.arrastre[t]) return;
      if (canc[t] && ymd && canc[t].ultima >= ymd) return;   // cancelada después de cargarla
      res.arrastre[t] = info;
    });
  });

  return res;
}

/**
 * Clave de preservación. A la clave base (tipo + beneficiario normalizado + moneda) se le suman
 * los discriminantes que parten grupos: el comentario arrastrado, la etiqueta de SUBGRUPO, el CBU
 * (en A y C) y la
 * separación por estado anterior. Sin ellos, dos filas del mismo beneficiario compartirían clave
 * y una le pisaría el estado a la otra al regenerar.
 */
function separadorGrupo_(arrastre, marcaCanc, cbuClave) {
  return (arrastre ? arrastre.clave : '') + '#' + (marcaCanc || '') + '#' + (cbuClave || '');
}

function claveGrupoAgrupExt_(tipo, beneficiario, moneda, comentarioAnterior, subgrupo, separador) {
  return claveGrupoAgrup_(tipo, beneficiario, moneda) + '|' +
         String(comentarioAnterior == null ? '' : comentarioAnterior).trim() + '|' +
         String(subgrupo == null ? '' : subgrupo).trim().toUpperCase() + '|' +
         String(separador == null ? '' : separador);
}

/**
 * Lee la agrupada de hoy que ya existe para preservar CHECK, ESTADO, COMENTARIO y COMPROBANTE.
 * El separador (estado anterior + marca de cancelación + CBU) y la etiqueta de SUBGRUPO no se leen
 * de las celdas: se recalculan desde los IDs de cada fila con los mismos mapas (anterior.arrastre,
 * canc y porTask) que usa la agrupación, así la clave coincide.
 */
function leerEstadoPrevioAgrup_(hoja, anterior, canc, porTask) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return {};

  const ancho = Math.max(hoja.getLastColumn(), 1);
  const hdr = hoja.getRange(1, 1, 1, ancho).getValues()[0]
    .map(h => String(h == null ? '' : h).trim().toUpperCase());

  const iTipo = hdr.indexOf('TIPO');
  const iBenef = hdr.indexOf('BENEFICIARIO');
  const iIds = hdr.indexOf('TRANSACCIONES (IDS)');
  const iMoneda = hdr.indexOf('MONEDA');
  const iCheck = hdr.indexOf('CHECK');
  const iEstado = hdr.indexOf('ESTADO');
  const iComent = hdr.indexOf('COMENTARIO');
  const iComentAnt = hdr.indexOf('COMENTARIO ANTERIOR');
  const iComprob = hdr.indexOf('COMPROBANTE');

  if (iTipo < 0 || iIds < 0) return {};

  const arrastre = (anterior && anterior.arrastre) || {};
  canc = canc || {};
  porTask = porTask || {};

  const rango = hoja.getRange(2, 1, ultimaFila - 1, ancho);
  const valores = rango.getValues();
  // getValues() sobre la celda del comprobante devuelve el texto visible ("Ver comprobante"), no
  // la URL. El enlace real sale del rich text, y sin esto se perdería en cada regeneración.
  const enlaces = iComprob >= 0 ? rango.getRichTextValues() : null;
  const previos = {};

  valores.forEach((v, i) => {
    // Una fila sin IDs es una banda de tipo o la fila de resumen, no un grupo real.
    const ids = String(v[iIds] == null ? '' : v[iIds]).split(',').map(s => s.trim()).filter(String);
    if (ids.length === 0) return;

    let infoArrastre = null;
    for (let k = 0; k < ids.length; k++) {
      if (arrastre[ids[k]]) { infoArrastre = arrastre[ids[k]]; break; }
    }
    // Todas las tareas de un grupo comparten marca de cancelación, etiqueta de SUBGRUPO y CBU
    // (son parte de la clave de grupo): alcanza con la primera. La etiqueta y el CBU salen de la
    // hoja del día y no de la celda SUBGRUPO, que en A y C muestra el CBU y no la etiqueta.
    const deTask = porTask[ids[0]] || { subgrupo: '', cbuClave: '' };
    const separador = separadorGrupo_(infoArrastre, marcaCancelacion_(canc[ids[0]]), deTask.cbuClave);

    const clave = claveGrupoAgrupExt_(
      v[iTipo],
      iBenef >= 0 ? v[iBenef] : '',
      iMoneda >= 0 ? v[iMoneda] : '',
      iComentAnt >= 0 ? v[iComentAnt] : '',
      deTask.subgrupo,
      separador
    );

    previos[clave] = {
      check: iCheck >= 0 ? v[iCheck] === true : false,
      estado: iEstado >= 0 ? String(v[iEstado] || '') : '',
      comentario: iComent >= 0 ? v[iComent] : '',
      comprobante: enlaces ? (enlaces[i][iComprob].getLinkUrl() || '') : ''
    };
  });

  return previos;
}

/**
 * Lee las filas de datos de la hoja del día, salteando encabezado y bandas de tipo.
 * Se indexa por nombre de header y no por posición fija, así una hoja generada con una versión
 * anterior (12 o 13 columnas) se sigue leyendo bien.
 */
function leerFilasDeHojaDelDia_(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];

  const idx = indicesPorHeader_(hoja);
  if (idx['TIPO'] < 0 || idx['TASK'] < 0) return [];

  const ancho = Math.max(hoja.getLastColumn(), 1);
  const valores = hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues();
  const leer = (v, nombre) => (idx[nombre] >= 0 ? v[idx[nombre]] : '');

  const filas = [];
  valores.forEach(v => {
    const tipo = leer(v, 'TIPO');
    const task = leer(v, 'TASK');
    if (task === '' || tipo === '') return;                       // banda de tipo o fila vacía
    if (AGRUP_TIPOS_EXCLUIDOS.indexOf(tipo) !== -1) return;
    if (beneficiarioExcluido_(leer(v, 'BENEFICIARIO'))) return;
    if (AGRUP_EXCLUIR_YA_PAGADOS &&
        String(leer(v, 'ESTADO') || '').indexOf(MARCA_PAGADO_PREFIJO) === 0) return;

    filas.push({
      tipo: tipo,
      task: String(task),
      dato: String(leer(v, 'CBU / DATO DE PAGO') || '').trim(),
      beneficiario: String(leer(v, 'BENEFICIARIO') || '').trim(),
      monto: aNumero_(leer(v, 'MONTO')),
      montoFactura: aNumero_(leer(v, 'MONTO FACTURA')),
      moneda: String(leer(v, 'MONEDA') || '').trim().toUpperCase(),
      comentarioAnterior: String(leer(v, 'COMENTARIO ANTERIOR') || '').trim(),
      subgrupo: String(leer(v, 'SUBGRUPO') || '').trim(),
      link: String(leer(v, 'INVOICE/COMPROBANTE') || '').trim(),
      // Los completa generarHojaAgrupadaDelDia con la agrupada anterior y las cancelaciones.
      arrastre: null,
      canc: null,
      marcaCanc: '',
      cbuClave: ''
    });
  });

  return filas;
}

/**
 * Arma un grupo por TIPO + BENEFICIARIO normalizado + MONEDA + COMENTARIO ANTERIOR + SUBGRUPO +
 * separación por estado anterior + marca de cancelación + CBU (solo A y C).
 *
 * Los cinco últimos son los discriminantes que parten un beneficiario en varias filas: si de
 * cuatro tareas del mismo proveedor una trae un comentario arrastrado, ayer ya estaba cargada en el
 * banco, se canceló o va a otra cuenta, no se suma con las otras.
 *
 * Si la fila no trae BENEFICIARIO (Payoneer, Paypal), se usa el dato de pago como identidad: es lo
 * único que hay, y agrupa bien dos pagos al mismo mail. Si tampoco hay dato de pago, la fila queda
 * sola en su propio grupo.
 *
 * Orden de salida: por tipo (según ORDEN_TIPOS), después por moneda, y dentro de cada moneda por
 * monto descendente. La moneda entra antes que el monto porque si no se mezclarían 1.000 BRL con
 * 100 USD en la misma escala y el orden no querría decir nada.
 */
function agruparPorBeneficiario_(filas) {
  const mapa = {};

  filas.forEach(f => {
    let beneficiario = f.beneficiario;
    if (beneficiario === '' && f.dato !== '') beneficiario = f.dato;

    const nombreClave = beneficiario === ''
      ? '#SIN-BENEFICIARIO#' + f.task
      : Array.from(normalizarNombre_(beneficiario)).sort().join(' ');

    const clave = f.tipo + '|' + nombreClave + '|' + f.moneda + '|' +
                  f.comentarioAnterior + '|' + f.subgrupo.toUpperCase() + '|' +
                  (f.arrastre ? f.arrastre.clave : '') + '|' + (f.marcaCanc || '') + '|' +
                  (f.cbuClave || '');

    if (!mapa[clave]) {
      mapa[clave] = {
        tipo: f.tipo,
        beneficiario: beneficiario,
        moneda: f.moneda,
        subgrupo: f.subgrupo,
        comentarioAnterior: f.comentarioAnterior,
        arrastre: f.arrastre || null,
        canc: f.canc || null,
        marcaCanc: f.marcaCanc || '',
        cbuClave: f.cbuClave || '',
        datos: [],
        ids: [], monto: 0, montoBase: 0, montoFactura: 0, links: [],
        conLink: 0, porFactura: {},
        check: false, estado: '', comentario: '', comprobante: '', origen: '', repetido: false
      };
    }
    const g = mapa[clave];
    g.ids.push(f.task);
    if (f.dato !== '' && g.datos.indexOf(f.dato) === -1) g.datos.push(f.dato);
    if (f.monto !== null) g.montoBase = redondear2_(g.montoBase + f.monto);
    if (f.link !== '' && g.links.indexOf(f.link) === -1) g.links.push(f.link);

    // Se indexa por LINK, no por fila: lo que se suma son facturas distintas.
    // Dos casos reales que obligan a esto:
    //   - una factura que cubre doce filas (las satélite no traen link) → se cuenta una vez;
    //   - dos pagos reales de la misma tarea contra la MISMA factura → también se cuenta una vez.
    // Sumando por fila, el segundo caso contaba la factura dos veces.
    if (f.link !== '') {
      g.conLink++;
      if (!g.porFactura[f.link]) g.porFactura[f.link] = { monto: null, tasks: [] };
      g.porFactura[f.link].tasks.push(f.task);
      if (f.montoFactura !== null) g.porFactura[f.link].monto = f.montoFactura;
    }
  });

  // El monto oficial: base para todos los tipos, factura para A y C.
  Object.keys(mapa).forEach(k => {
    const g = mapa[k];
    if (TIPOS_VALIDABLES.indexOf(g.tipo) === -1) {
      g.monto = g.montoBase;
      g.montoIncompleto = '';
      return;
    }
    const links = Object.keys(g.porFactura);
    if (links.length === 0) {
      g.monto = null;
      g.montoIncompleto = 'ninguna de sus ' + g.ids.length + ' fila(s) tiene factura adjunta';
      return;
    }

    const faltan = [];
    let suma = 0;
    links.forEach(l => {
      const e = g.porFactura[l];
      if (e.monto === null) faltan.push(e.tasks.join(' y '));
      else suma = redondear2_(suma + e.monto);
      if (e.tasks.length > 1) g.facturaCompartida = true;
    });

    g.facturas = links.length;
    if (faltan.length > 0) {
      g.monto = null;
      g.montoIncompleto = 'falta leer la factura de ' + faltan.join(', ') +
                          ' (corré "Validar beneficiario y monto" sobre esas filas)';
    } else {
      g.montoFactura = suma;
      g.monto = suma;
      g.montoIncompleto = '';
    }
  });

  return Object.keys(mapa)
    .map(k => mapa[k])
    .sort((a, b) => {
      const dt = ORDEN_TIPOS.indexOf(a.tipo) - ORDEN_TIPOS.indexOf(b.tipo);
      if (dt !== 0) return dt;
      if (a.moneda !== b.moneda) return a.moneda.localeCompare(b.moneda);
      return (b.monto === null ? -1 : b.monto) - (a.monto === null ? -1 : a.monto);
    });
}

/** true si el beneficiario contiene alguno de los nombres excluidos (sin tildes ni mayúsculas). */
function beneficiarioExcluido_(beneficiario) {
  const texto = textoPlano_(beneficiario);
  if (texto === '') return false;
  return AGRUP_BENEFICIARIOS_EXCLUIDOS.some(x => texto.indexOf(textoPlano_(x)) !== -1);
}

/** true si el comentario contiene "COMPROBANTE CARGADO" (ignora mayúsculas, tildes y puntuación). */
function comprobanteYaCargado_(comentario) {
  const texto = textoPlano_(comentario);
  if (texto === '') return false;
  return texto.indexOf(textoPlano_(AGRUP_TEXTO_COMPROBANTE)) !== -1;
}

/**
 * true si el comentario arrastrado dice "COMPROBANTE CARGADO" pero NO salió de la agrupada del día
 * anterior, sino de una más vieja. Significa que el comentario sobrevivió un ciclo entero: el
 * comprobante nunca se cargó de verdad, o la fila nunca se marcó TRANSFERIDO.
 */
function comentarioRepetido_(comentarioAnterior, fechaAgrupAnterior) {
  if (!fechaAgrupAnterior) return false;
  if (!comprobanteYaCargado_(comentarioAnterior)) return false;
  const m = String(comentarioAnterior).match(/^(\d{2}-\d{2})\s*\|/);
  if (!m) return false;
  return m[1] !== fechaAgrupAnterior;
}

/** "USD 1.200,00 · BRL 3.400,00" a partir de una lista de grupos. */
function resumenPorMoneda_(grupos) {
  const totales = {};
  grupos.forEach(g => {
    if (g.monto === null) return;   // grupo sin monto verificado: no suma
    const m = g.moneda || '(sin moneda)';
    totales[m] = redondear2_((totales[m] || 0) + g.monto);
  });
  return Object.keys(totales).sort()
    .map(m => m + ' ' + formatearImporte_(totales[m]))
    .join(' · ');
}

/**
 * resumenPorMoneda_ + cuántos grupos quedaron afuera por no tener monto verificado.
 * Sin esto, una banda con todos los grupos sin MONTO FACTURA quedaba "5 grupo(s) · " y nada más.
 */
function textoMontos_(grupos) {
  const montos = resumenPorMoneda_(grupos);
  const sinMonto = grupos.filter(g => g.monto === null).length;
  const extra = sinMonto ? sinMonto + ' sin monto factura' : '';
  return [montos, extra].filter(String).join(' · ') || 'sin montos';
}

/** Letra de columna a partir del número 1-based. 1 → A, 27 → AA. */
function letraColumna_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function escribirHojaAgrupada_(hoja, grupos, nombreDia) {
  hoja.getRange(1, 1, 1, AGRUP_HEADERS.length)
    .setValues([AGRUP_HEADERS])
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#434343')
    .setHorizontalAlignment('center');

  // ── Armado de las filas: resumen, después banda de tipo + sus grupos ──────────────────────
  const salida = [];
  const meta = [];   // por fila: { grupo: obj|null }

  // Pagado = TRANSFERIDO o COMPROBANTE CARGADO (los dos cuentan como pago en registrarPagado_).
  const pendientes = grupos.filter(g => g.estado !== EST_TRANSFERIDO && g.estado !== EST_COMPROBANTE);
  const resumen = new Array(AGRUP_HEADERS.length).fill('');
  resumen[AGRUP_COL.TIPO - 1] = nombreDia;
  resumen[AGRUP_COL.BENEF - 1] =
    grupos.length + ' grupos · ' + (grupos.length - pendientes.length) + ' pagados · ' +
    'Pendiente: ' + (pendientes.length ? textoMontos_(pendientes) : 'nada');
  salida.push(resumen);
  meta.push({ grupo: null });

  ORDEN_TIPOS.forEach(tipo => {
    const delTipo = grupos.filter(g => g.tipo === tipo);
    if (delTipo.length === 0) return;

    const banda = new Array(AGRUP_HEADERS.length).fill('');
    banda[AGRUP_COL.TIPO - 1] = tipo;
    banda[AGRUP_COL.BENEF - 1] = delTipo.length + ' grupo(s) · ' + textoMontos_(delTipo);
    salida.push(banda);
    meta.push({ grupo: null });

    delTipo.forEach(g => {
      const fila = new Array(AGRUP_HEADERS.length).fill('');
      fila[AGRUP_COL.TIPO - 1] = g.tipo;
      fila[AGRUP_COL.BENEF - 1] = g.beneficiario;
      fila[AGRUP_COL.SUBGRUPO - 1] = textoSubgrupo_(g);
      fila[AGRUP_COL.NTX - 1] = g.ids.length;
      fila[AGRUP_COL.IDS - 1] = g.ids.join(',');
      fila[AGRUP_COL.MONTO - 1] = (g.monto === null ? '' : g.monto);
      fila[AGRUP_COL.MONEDA - 1] = g.moneda;
      fila[AGRUP_COL.CHECK - 1] = g.check === true;
      fila[AGRUP_COL.ESTADO - 1] = g.estado || '';
      fila[AGRUP_COL.ORIGEN - 1] = g.origen || '';
      fila[AGRUP_COL.COMENT - 1] = g.comentario || '';
      fila[AGRUP_COL.COMENT_ANT - 1] = g.comentarioAnterior || '';
      fila[AGRUP_COL.PAGAR - 1] = '';   // la llena conMontoAPagar_ (CargaGalicia.gs)
      fila[AGRUP_COL.LINKS - 1] = g.links.join('\n');
      fila[AGRUP_COL.PDF - 1] = '';
      fila[AGRUP_COL.COMPROBANTE - 1] = '';   // el enlace se repone más abajo, como rich text
      salida.push(fila);
      meta.push({ grupo: g });
    });
  });

  const total = salida.length;

  // ── Escritura ────────────────────────────────────────────────────────────────────────────
  hoja.getRange(2, AGRUP_COL.IDS, total, 1).setNumberFormat('@');
  // SUBGRUPO en TEXTO antes de escribir: un CBU son 22 dígitos, y como número Sheets lo redondea
  // a 15 dígitos significativos y lo muestra en notación científica. El CBU quedaría inservible.
  hoja.getRange(2, AGRUP_COL.SUBGRUPO, total, 1).setNumberFormat('@');
  // La fila de resumen lleva la fecha "07-09-2026" en la columna TIPO: sin formato TEXTO, Sheets
  // la convierte en fecha y la muestra como 7 de septiembre en otro formato.
  hoja.getRange(2, AGRUP_COL.TIPO, 1, 1).setNumberFormat('@');
  hoja.getRange(2, 1, total, AGRUP_HEADERS.length).setValues(salida);
  hoja.getRange(2, AGRUP_COL.MONTO, total, 1).setNumberFormat('#,##0.00');
  hoja.getRange(2, AGRUP_COL.PAGAR, total, 1).setNumberFormat('#,##0.00');
  hoja.getRange(2, AGRUP_COL.NTX, total, 1).setHorizontalAlignment('center');
  hoja.getRange(2, AGRUP_COL.IDS, total, 1).setWrap(true).setVerticalAlignment('top');
  hoja.getRange(2, AGRUP_COL.COMENT, total, 2).setWrap(true).setVerticalAlignment('top');
  hoja.getRange(2, AGRUP_COL.LINKS, total, 2).setWrap(true).setVerticalAlignment('top');

  // Casillas y desplegable solo en las filas de datos: en las bandas quedan en null.
  const reglaCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(ESTADOS_AGRUP, true)
    .setAllowInvalid(false)
    .setHelpText('Elegí un estado de la lista. Solo TRANSFERIDO saca la fila mañana.')
    .build();

  hoja.getRange(2, AGRUP_COL.CHECK, total, 1)
    .setDataValidations(meta.map(m => [m.grupo ? reglaCheck : null]));
  hoja.getRange(2, AGRUP_COL.ESTADO, total, 1)
    .setDataValidations(meta.map(m => [m.grupo ? reglaEstado : null]));

  // ── Colores ──────────────────────────────────────────────────────────────────────────────
  // Prioridad: naranja de comentario repetido > verde de comprobante cargado > color por tipo.
  const fondos = meta.map((m, i) => {
    const anchoFila = AGRUP_HEADERS.length;
    if (!m.grupo) {
      const color = i === 0 ? '#666666' : (COLOR_POR_TIPO[salida[i][AGRUP_COL.TIPO - 1]] || '#d9d9d9');
      return new Array(anchoFila).fill(color);
    }
    let color;
    if (m.grupo.repetido) color = AGRUP_COLOR_REPETIDO;
    else if (comprobanteYaCargado_(m.grupo.comentarioAnterior)) color = AGRUP_COLOR_COMPROBANTE;
    else color = colorFondoDeFila_(m.grupo.tipo);
    return new Array(anchoFila).fill(color);
  });
  hoja.getRange(2, 1, total, AGRUP_HEADERS.length).setBackgrounds(fondos);

  // Las dos filas de resumen (la general y las bandas) van en negrita.
  meta.forEach((m, i) => {
    if (m.grupo) return;
    const rango = hoja.getRange(i + 2, 1, 1, AGRUP_HEADERS.length).setFontWeight('bold');
    if (i === 0) rango.setFontColor('#ffffff');
  });

  // Reponer el enlace del comprobante ya armado. Va después de setValues porque el rich text se
  // pisa con cualquier escritura de valores sobre la misma celda.
  meta.forEach((m, i) => {
    if (!m.grupo || !m.grupo.comprobante) return;
    hoja.getRange(i + 2, AGRUP_COL.COMPROBANTE).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('Ver comprobante').setLinkUrl(m.grupo.comprobante).build()
    );
  });

  // Monto: nota cuando no se pudo tomar de la factura, o cuando difiere de lo que dijo AppSheet.
  // La diferencia no se corrige ni se oculta: manda la factura, pero queda registrada.
  meta.forEach((m, i) => {
    if (!m.grupo) return;
    const g = m.grupo;
    if (g.montoIncompleto) {
      hoja.getRange(i + 2, AGRUP_COL.MONTO)
        .setNote('⚠️ SIN MONTO VERIFICADO\n' + g.montoIncompleto +
                 '\n\nAppSheet dice ' + formatearImporte_(g.montoBase) + ' ' + g.moneda +
                 ', pero para ' + g.tipo + ' el monto que vale es el de la factura.' +
                 '\nEste grupo no llega a la hoja de APROBACIÓN ni se le puede armar el PDF.')
        .setBackground(AGRUP_COLOR_REPETIDO);
    } else if (TIPOS_VALIDABLES.indexOf(g.tipo) !== -1 &&
               redondear2_(Math.abs(g.monto - g.montoBase)) > 0.01) {
      hoja.getRange(i + 2, AGRUP_COL.MONTO)
        .setNote('Monto tomado de la factura: ' + formatearImporte_(g.monto) +
                 ' (' + g.facturas + ' factura/s distinta/s)' +
                 '\nAppSheet (base) decía: ' + formatearImporte_(g.montoBase) +
                 '\nDiferencia: ' + formatearImporte_(redondear2_(g.monto - g.montoBase)) +
                 (g.facturaCompartida
                   ? '\n\n⚠️ Hay una factura que cubre más de una fila de pago, así que se contó ' +
                     'UNA sola vez. Si esas filas son pagos separados que suman más que la factura, ' +
                     'este total está corto: verificalo antes de transferir.'
                   : ''));
    }
  });

  // Nota en las filas con comentario repetido.
  meta.forEach((m, i) => {
    if (!m.grupo || !m.grupo.repetido) return;
    hoja.getRange(i + 2, AGRUP_COL.COMENT_ANT).setNote(
      '⚠️ Este comentario viene de una agrupada anterior a la de ayer, o sea que ya apareció más ' +
      'de un día.\nSi el comprobante se hubiera cargado de verdad, la fila no debería seguir acá. ' +
      'Revisá si la transferencia salió y si el comprobante está subido a AppSheet.'
    );
  });

  // Comentario de cancelación: negrita, y rojo si ya no se puede repagar.
  meta.forEach((m, i) => {
    if (!m.grupo || !m.grupo.canc) return;
    const celda = hoja.getRange(i + 2, AGRUP_COL.COMENT).setFontWeight('bold');
    if (m.grupo.canc.veces >= 2) celda.setFontColor('#cc0000');
  });

  // Nota en las filas separadas por el estado que tenían en la agrupada anterior.
  meta.forEach((m, i) => {
    if (!m.grupo || !m.grupo.arrastre) return;
    const a = m.grupo.arrastre;
    hoja.getRange(i + 2, AGRUP_COL.ORIGEN).setNote(
      'Estas tareas estaban en "' + a.estado + '" en la agrupada del ' + a.fecha + '.\n' +
      'Van en fila aparte para que no se sumen a operaciones nuevas del mismo beneficiario: esa ' +
      'plata ya estaba cargada en el banco.\n\n' +
      'Si la transferencia salió, poné TRANSFERIDO: mañana deja de aparecer.'
    );
  });

  // ── Formato condicional: las filas revisadas se atenúan solas ────────────────────────────
  // Solo fuente, sin fondo: si pintara el fondo taparía el verde y el naranja, que son la señal
  // importante. La regla se reevalúa sola al tildar, sin regenerar la hoja.
  const rangoDatos = hoja.getRange(2, 1, total, AGRUP_HEADERS.length);
  const reglaTildado = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + letraColumna_(AGRUP_COL.CHECK) + '2=TRUE')
    .setFontColor('#999999')
    .setStrikethrough(true)
    .setRanges([rangoDatos])
    .build();
  hoja.setConditionalFormatRules([reglaTildado]);

  // ── Bordes, filtro, congelado ────────────────────────────────────────────────────────────
  hoja.getRange(1, 1, total + 1, AGRUP_HEADERS.length)
    .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(1, 1, total + 1, AGRUP_HEADERS.length).createFilter();

  hoja.setFrozenRows(2);      // encabezado + fila de resumen
  hoja.setFrozenColumns(2);   // TIPO + BENEFICIARIO siempre visibles

  // ── Anchos ───────────────────────────────────────────────────────────────────────────────
  hoja.autoResizeColumns(1, AGRUP_HEADERS.length);
  hoja.setColumnWidth(AGRUP_COL.TIPO, 180);
  hoja.setColumnWidth(AGRUP_COL.BENEF, 240);
  hoja.setColumnWidth(AGRUP_COL.SUBGRUPO, 190);
  hoja.setColumnWidth(AGRUP_COL.NTX, 55);
  hoja.setColumnWidth(AGRUP_COL.IDS, 230);
  hoja.setColumnWidth(AGRUP_COL.CHECK, 60);
  hoja.setColumnWidth(AGRUP_COL.ESTADO, 150);
  hoja.setColumnWidth(AGRUP_COL.ORIGEN, 140);
  hoja.setColumnWidth(AGRUP_COL.COMENT, 240);
  hoja.setColumnWidth(AGRUP_COL.COMENT_ANT, 240);
  hoja.setColumnWidth(AGRUP_COL.PAGAR, 120);
  hoja.setColumnWidth(AGRUP_COL.LINKS, 300);
  hoja.setColumnWidth(AGRUP_COL.PDF, 100);
  hoja.setColumnWidth(AGRUP_COL.COMPROBANTE, 130);

  recortarGrilla_(hoja, total + 1, AGRUP_HEADERS.length);

  // LINKS y PDF colapsados: se abren con el "+" de arriba cuando los necesitás.
  // Va al final, después de recortarGrilla_, porque el agrupado se pierde si se borran columnas.
  try {
    hoja.getRange(1, AGRUP_COL.LINKS, 1, 2).shiftColumnGroupDepth(1);
    hoja.getColumnGroup(AGRUP_COL.LINKS, 1).collapse();
  } catch (e) {
    hoja.hideColumns(AGRUP_COL.LINKS, 2);
  }
}

/**
 * Toast (no bloquea) con los beneficiarios que quedaron partidos en dos filas: una con lo que ayer
 * estaba cargado en el banco y otra con lo nuevo. No avisa por las filas separadas que están solas,
 * porque ahí no se evitó ninguna suma.
 */
function avisarSeparadosPorEstadoAnterior_(ss, grupos) {
  const porBase = {};
  grupos.forEach(g => {
    const k = claveGrupoAgrup_(g.tipo, g.beneficiario, g.moneda);
    if (!porBase[k]) porBase[k] = { nombre: g.beneficiario, conArrastre: false, sinArrastre: false };
    if (g.arrastre) porBase[k].conArrastre = true;
    else porBase[k].sinArrastre = true;
  });

  const partidos = Object.keys(porBase)
    .map(k => porBase[k])
    .filter(x => x.conArrastre && x.sinArrastre)
    .map(x => x.nombre);

  if (partidos.length === 0) return;
  ss.toast(partidos.length + ' beneficiario(s) quedaron en dos filas porque ayer ya tenían una ' +
           'operación cargada en el banco: ' + partidos.join(', ') + '. Ver columna ORIGEN.',
           'Agrupado', 10);
}

/**
 * Popup con los grupos A y C que quedaron sin monto porque no se pudo tomar de la factura.
 * Esos grupos no llegan a APROBACIÓN ni se les puede armar el PDF, así que tienen que aparecer
 * en algún lado: si desaparecieran en silencio, un pago se podría perder.
 */
function avisarMontosSinFactura_(grupos) {
  const sinMonto = grupos.filter(g => g.montoIncompleto);
  if (sinMonto.length === 0) return;

  const detalle = sinMonto.map(g =>
    '• ' + g.beneficiario + ' (' + g.tipo + ')\n' +
    '   AppSheet decía ' + formatearImporte_(g.montoBase) + ' ' + g.moneda + '\n' +
    '   ' + g.montoIncompleto
  ).join('\n');

  SpreadsheetApp.getUi().alert(
    '⚠️ GRUPOS SIN MONTO DE FACTURA\n\n' +
    sinMonto.length + ' grupo(s) de tipo A o C quedaron sin monto. Para esos tipos el monto sale ' +
    'de la factura, no de AppSheet:\n\n' + detalle + '\n\n' +
    'Están en naranja en la columna MONTO. No aparecen en APROBACIÓN ni se les puede armar el PDF ' +
    'hasta que se lea la factura.'
  );
}

/** Popup con los grupos cuyo comentario de comprobante se repite. No hace nada si no hay ninguno. */
function avisarComprobantesRepetidos_(grupos) {
  const repetidos = grupos.filter(g => g.repetido);
  if (repetidos.length === 0) return;

  const detalle = repetidos.map(g =>
    '• ' + g.beneficiario + ' — ' + (g.monto === null ? 'sin monto' : formatearImporte_(g.monto)) + ' ' + g.moneda +
    ' — ' + g.comentarioAnterior
  ).join('\n');

  SpreadsheetApp.getUi().alert(
    '⚠️ COMPROBANTE QUE SE REPITE\n\n' +
    repetidos.length + ' grupo(s) arrastran un "' + AGRUP_TEXTO_COMPROBANTE + '" de hace más de ' +
    'un día. Si el comprobante estuviera cargado, esas filas ya no deberían aparecer:\n\n' +
    detalle + '\n\n' +
    'Están pintadas de naranja. Revisá si la transferencia salió y si el comprobante está en AppSheet.'
  );
}

/* ══════════════════ MARCADO EN BLOQUE ══════════════════ */

/**
 * Pone CARGADO EN BANCO en toda fila de la agrupada de hoy que tenga el ESTADO vacío.
 * NO pisa ningún estado ya cargado: si una fila dice TRANSFERIDO, CON PROBLEMA, RECHAZADO o
 * CANCELADO, se deja como está.
 * Tampoco marca:
 *   - las filas "↻ EN BANCO": ya estaban cargadas en el banco en la agrupada anterior;
 *   - las tareas canceladas 2 veces o más (Cancelaciones.gs).
 * Las canceladas una sola vez SÍ se marcan: se repagan una vez.
 *
 * Es el paso que habilita la hoja de APROBACIÓN. Se corre cuando terminaste de cargar los pagos
 * en el banco de verdad, no al generar la agrupada: lo que ve quien autoriza tiene que ser algo
 * que vos afirmaste, no algo que el script asumió por vos.
 * (Ya no está en el menú: CargaGalicia.gs aplica los mismos resguardos al armar el archivo.)
 */
function marcarTodoCargadoEnBanco() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nombreDia = formatearNombreHoja_(new Date());
  const hoja = ss.getSheetByName(AGRUP_PREFIJO + nombreDia);

  if (!hoja) {
    ui.alert('No existe la hoja "' + AGRUP_PREFIJO + nombreDia + '". Generá primero la agrupada.');
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    ui.alert('La hoja agrupada está vacía.');
    return;
  }

  const ids = hoja.getRange(2, AGRUP_COL.IDS, ultimaFila - 1, 1).getValues();
  const rangoEstado = hoja.getRange(2, AGRUP_COL.ESTADO, ultimaFila - 1, 1);
  const estados = rangoEstado.getValues();

  // Mismos mapas que usó la generación de la agrupada, para saltear lo mismo que ella separó.
  const canc = estadoCancelaciones_(ss);
  const anterior = agrupAnterior_(ss, nombreDia, canc);

  const aMarcar = [];
  let yaTenian = 0;
  let enBanco = 0;
  let noRepagar = 0;

  for (let i = 0; i < estados.length; i++) {
    const lista = String(ids[i][0] || '').split(',').map(s => s.trim()).filter(String);
    if (lista.length === 0) continue;                                 // banda o fila de resumen
    if (String(estados[i][0] || '').trim() !== '') { yaTenian++; continue; }
    // Cancelada 2 veces o más: no se repaga, se avisa en AppSheet.
    if (lista.some(t => canc[t] && canc[t].veces >= 2)) { noRepagar++; continue; }
    // Ya estaba cargada en el banco en la agrupada anterior: marcarla la mandaría otra vez a
    // APROBACIÓN y se podría autorizar dos veces. Si salió, se le pone TRANSFERIDO a mano.
    if (lista.some(t => anterior.arrastre[t])) { enBanco++; continue; }
    aMarcar.push(i);
  }

  const salteadas =
    (enBanco ? enBanco + ' fila(s) ya estaban cargadas en el banco en la agrupada anterior ' +
               '(ORIGEN "↻ EN BANCO"): no se tocan. Si salieron, poneles TRANSFERIDO.\n' : '') +
    (noRepagar ? noRepagar + ' fila(s) se cancelaron 2 veces o más: no se tocan. Hay que avisar ' +
                 'en AppSheet.\n' : '');

  if (aMarcar.length === 0) {
    ui.alert('No hay ninguna fila para marcar.\n\n' + salteadas);
    return;
  }

  const resp = ui.alert(
    aMarcar.length + ' fila(s) van a pasar a "' + EST_CARGADO + '".\n' +
    yaTenian + ' fila(s) ya tienen un estado y no se tocan.\n' + salteadas + '\n' +
    'Estas filas van a aparecer en la hoja de APROBACIÓN. Corré esto solo si ya las cargaste ' +
    'en el banco.\n\n¿Confirmás?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  aMarcar.forEach(i => { estados[i][0] = EST_CARGADO; });
  rangoEstado.setValues(estados);

  ui.alert(aMarcar.length + ' fila(s) quedaron en "' + EST_CARGADO + '".\n\n' +
           'Ya podés generar la hoja de APROBACIÓN.');
}

/* ══════════════════ HOJA DE APROBACIÓN ══════════════════ */

/**
 * Arma la hoja que consulta quien autoriza los pagos. Nombre fijo "APROBACIÓN", así el link que le
 * pasás una vez le sirve siempre. Se reescribe entera en cada corrida.
 * Trae solo los grupos de la agrupada de hoy cuyo ESTADO sea CARGADO EN BANCO.
 * Queda protegida con aviso: si alguien intenta editarla, Sheets le avisa antes.
 */
function generarHojaAprobacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nombreDia = formatearNombreHoja_(new Date());
  const hojaAgrup = ss.getSheetByName(AGRUP_PREFIJO + nombreDia);

  if (!hojaAgrup) {
    ui.alert('No existe la hoja "' + AGRUP_PREFIJO + nombreDia + '". Generá primero la agrupada.');
    return;
  }

  const ultimaFila = hojaAgrup.getLastRow();
  const filas = [];
  const incompletos = [];

  if (ultimaFila >= 2) {
    hojaAgrup.getRange(2, 1, ultimaFila - 1, AGRUP_HEADERS.length).getValues().forEach(v => {
      if (String(v[AGRUP_COL.IDS - 1] || '').trim() === '') return;
      if (String(v[AGRUP_COL.ESTADO - 1] || '').trim().toUpperCase() !== EST_CARGADO) return;

      const moneda = String(v[AGRUP_COL.MONEDA - 1] || '').trim();
      const monto = aNumero_(v[AGRUP_COL.MONTO - 1]);

      // Sin moneda o sin monto no hay nada que autorizar: no se sabe cuánto se transfiere ni en
      // qué. La fila no entra, pero se lista al final para que no desaparezca en silencio.
      if (moneda === '' || monto === null || monto === 0) {
        incompletos.push(
          '• ' + String(v[AGRUP_COL.BENEF - 1] || '(sin beneficiario)') +
          ' — ' + String(v[AGRUP_COL.IDS - 1]) +
          ' — ' + (monto === null || monto === 0 ? 'SIN MONTO' : formatearImporte_(monto)) +
          ' ' + (moneda === '' ? 'SIN MONEDA' : moneda)
        );
        return;
      }

      filas.push([
        v[AGRUP_COL.TIPO - 1],
        String(v[AGRUP_COL.IDS - 1] || '').split(',').join('\n'),
        v[AGRUP_COL.BENEF - 1],
        monto,
        moneda,
        String(v[AGRUP_COL.LINKS - 1] || '')
      ]);
    });
  }

  const avisoIncompletos = incompletos.length === 0 ? '' :
    '\n\n⚠️ ' + incompletos.length + ' fila(s) quedaron AFUERA por no tener moneda o monto:\n' +
    incompletos.join('\n') + '\n\nCorregilas en la hoja del día y volvé a generar.';

  if (filas.length === 0) {
    ui.alert('No hay ningún grupo en estado "' + EST_CARGADO + '" con moneda y monto cargados. ' +
             'No hay nada que autorizar.' + avisoIncompletos);
    return;
  }

  const previa = ss.getSheetByName(APROB_SHEET_NAME);
  if (previa) ss.deleteSheet(previa);
  const hoja = ss.insertSheet(APROB_SHEET_NAME, 0);

  const totales = {};
  filas.forEach(f => {
    const m = String(f[4] || '(sin moneda)');
    totales[m] = redondear2_((totales[m] || 0) + (f[3] === null ? 0 : f[3]));
  });
  const totalTexto = Object.keys(totales).sort()
    .map(m => m + ' ' + formatearImporte_(totales[m])).join('   ·   ');

  const ancho = APROB_HEADERS.length;

  hoja.getRange(1, 1, 1, ancho).merge()
    .setValue('PAGOS A AUTORIZAR — ' + nombreDia)
    .setFontSize(16).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#0a0a8c').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 38);

  hoja.getRange(2, 1, 1, ancho).merge()
    .setValue(filas.length + ' transferencias   ·   ' + totalTexto)
    .setFontSize(12).setFontWeight('bold').setFontColor('#0a0a8c')
    .setBackground('#e8e8f5').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 30);

  hoja.getRange(3, 1, 1, ancho)
    .setValues([APROB_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#434343').setHorizontalAlignment('center');

  hoja.getRange(4, 1, filas.length, ancho).setValues(filas);
  hoja.getRange(4, 4, filas.length, 1).setNumberFormat('#,##0.00');
  hoja.getRange(4, 2, filas.length, 1).setWrap(true).setVerticalAlignment('top')
    .setFontSize(9).setFontColor('#666666');
  hoja.getRange(4, 6, filas.length, 1).setWrap(true).setVerticalAlignment('top').setFontSize(9);
  hoja.getRange(4, 1, filas.length, ancho).setVerticalAlignment('middle');

  // Bandas alternadas, más legible que el color por tipo para alguien que no trabaja la hoja.
  for (let i = 0; i < filas.length; i++) {
    if (i % 2 === 1) hoja.getRange(4 + i, 1, 1, ancho).setBackground('#f6f6f9');
  }

  hoja.getRange(3, 1, filas.length + 1, ancho)
    .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  hoja.setFrozenRows(3);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 150);
  hoja.setColumnWidth(3, 260);
  hoja.setColumnWidth(4, 110);
  hoja.setColumnWidth(5, 75);
  hoja.setColumnWidth(6, 300);

  recortarGrilla_(hoja, filas.length + 3, ancho);

  hoja.protect()
    .setDescription('Hoja de aprobación: la genera el script, no editar a mano')
    .setWarningOnly(true);

  ss.setActiveSheet(hoja);
  ui.alert(
    'Hoja "' + APROB_SHEET_NAME + '" lista.\n\n' +
    filas.length + ' transferencias a autorizar.\n' + totalTexto + '\n\n' +
    'El nombre de la hoja no cambia nunca, así que el link que ya pasaste sigue sirviendo.' +
    avisoIncompletos
  );
}

/* ══════════════════ PDF POR FILA ══════════════════ */

function generarPdfsDeSeleccion() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (hoja.getName().indexOf(AGRUP_PREFIJO) !== 0) {
    ui.alert('Esta función corre sobre la hoja agrupada ("' + AGRUP_PREFIJO + 'DD-MM-YYYY"). Estás parado en "' + hoja.getName() + '".');
    return;
  }

  const carpeta = DriveApp.getFolderById(PDF_CARPETA_ID);
  const rango = hoja.getActiveRange();
  const filaInicio = rango.getRow();
  const cantFilas = rango.getNumRows();

  let generados = 0, salteadas = 0;
  const errores = [];

  for (let i = 0; i < cantFilas; i++) {
    const fila = filaInicio + i;
    if (fila === 1) continue;

    const v = hoja.getRange(fila, 1, 1, AGRUP_HEADERS.length).getValues()[0];
    // Sin IDs no es un grupo: es la fila de resumen o una banda de tipo.
    if (String(v[AGRUP_COL.IDS - 1] || '').trim() === '') { salteadas++; continue; }
    // Sin monto verificado no hay orden de pago posible: para A y C el monto sale de la factura.
    if (aNumero_(v[AGRUP_COL.MONTO - 1]) === null) {
      errores.push('Fila ' + fila + ': sin monto verificado. Leé la factura con "Validar ' +
                   'beneficiario y monto" y regenerá la agrupada.');
      continue;
    }

    const datos = {
      tipo: v[AGRUP_COL.TIPO - 1],
      beneficiario: v[AGRUP_COL.BENEF - 1],
      ids: String(v[AGRUP_COL.IDS - 1] || ''),
      monto: aNumero_(v[AGRUP_COL.MONTO - 1]),
      moneda: v[AGRUP_COL.MONEDA - 1],
      links: String(v[AGRUP_COL.LINKS - 1] || '').split('\n').map(s => s.trim()).filter(String),
      fecha: hoja.getName().replace(AGRUP_PREFIJO, '')
    };

    try {
      const url = crearPdfDeGrupo_(datos, carpeta);
      hoja.getRange(fila, AGRUP_COL.PDF).setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText('Ver PDF').setLinkUrl(url).build()
      );
      generados++;
    } catch (err) {
      errores.push('Fila ' + fila + ': ' + err.message);
    }
  }

  ui.alert(
    'PDF generados: ' + generados + '\n' +
    'Filas salteadas: ' + salteadas +
    (errores.length ? '\n\nErrores:\n' + errores.join('\n') : '')
  );
}

/**
 * Arma la ORDEN DE PAGO como Google Doc y la devuelve como blob PDF. No toca ninguna carpeta.
 * Si recibe imagenBlob, la captura del pago se agrega en su propia página al final.
 * Separada de la escritura en Drive para que Comprobantes.gs pueda tomar este blob y unirlo con
 * el comprobante de Galicia y las retenciones antes de guardar nada.
 */
function blobOrdenDePago_(datos, imagenBlob) {
  const nombre = nombreArchivoPdf_(datos.ids);
  const doc = DocumentApp.create('TMP_PDF_' + nombre);

  try {
    const body = doc.getBody();
    body.clear();
    body.setAttributes({ [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE });

    insertarLogo_(body);

    const titulo = body.appendParagraph('ORDEN DE PAGO');
    titulo.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
      [DocumentApp.Attribute.FONT_SIZE]: 20,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#0a0a8c'
    });

    const fecha = body.appendParagraph(datos.fecha);
    fecha.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#666666'
    });

    const tabla = body.appendTable([
      ['BENEFICIARIO', String(datos.beneficiario || '—')],
      ['TIPO DE TRANSACCIÓN', String(datos.tipo || '—')],
      ['MONTO TOTAL', (datos.monto === null ? '—' : formatearImporte_(datos.monto)) + ' ' + String(datos.moneda || '')],
      ['TRANSACCIONES', datos.ids.split(',').join('\n')]
    ]);
    formatearTablaPdf_(tabla);

    if (datos.links.length > 0) {
      const rotulo = body.appendParagraph('COMPROBANTES');
      rotulo.setAttributes({
        [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
        [DocumentApp.Attribute.FONT_SIZE]: 11,
        [DocumentApp.Attribute.BOLD]: true
      });
      datos.links.forEach(link => {
        const p = body.appendParagraph(link);
        p.setAttributes({
          [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
          [DocumentApp.Attribute.FONT_SIZE]: 9,
          [DocumentApp.Attribute.LINK_URL]: link
        });
      });
    }

    // El comprobante bancario va en su propia página, después de la orden de pago.
    if (imagenBlob) insertarComprobanteEnDoc_(body, imagenBlob);

    doc.saveAndClose();

    return DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(nombre + '.pdf');
  } finally {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) { /* no bloquea */ }
  }
}

/** Guarda un blob PDF en la carpeta, pisando cualquier archivo previo con el mismo nombre. */
function guardarPdfEnCarpeta_(blob, nombre, carpeta) {
  blob.setName(nombre + '.pdf');
  const previos = carpeta.getFilesByName(nombre + '.pdf');
  while (previos.hasNext()) previos.next().setTrashed(true);
  return carpeta.createFile(blob).getUrl();
}

/** Orden de pago sola, guardada en carpeta. Es lo que usa "Generar PDF de filas seleccionadas". */
function crearPdfDeGrupo_(datos, carpeta, imagenBlob) {
  const nombre = nombreArchivoPdf_(datos.ids);
  return guardarPdfEnCarpeta_(blobOrdenDePago_(datos, imagenBlob), nombre, carpeta);
}

function insertarLogo_(body) {
  try {
    const blob = PDF_LOGO_FILE_ID
      ? DriveApp.getFileById(PDF_LOGO_FILE_ID).getBlob()
      : logoBlob_();
    const img = body.appendImage(blob);
    const ratio = img.getHeight() / img.getWidth();
    img.setWidth(120).setHeight(Math.round(120 * ratio));
    img.getParent().asParagraph().setSpacingAfter(18);
  } catch (e) { /* sin logo, el PDF sale igual */ }
}

function formatearTablaPdf_(tabla) {
  tabla.setBorderColor('#cccccc');
  for (let r = 0; r < tabla.getNumRows(); r++) {
    const filaT = tabla.getRow(r);
    filaT.getCell(0).setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
      [DocumentApp.Attribute.FONT_SIZE]: 9,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#666666'
    });
    filaT.getCell(1).setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: PDF_FUENTE,
      [DocumentApp.Attribute.FONT_SIZE]: 11,
      [DocumentApp.Attribute.BOLD]: false,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000'
    });
  }
  tabla.setColumnWidth(0, 150);
}

/** Los IDs separados por coma son el nombre del archivo. Se recorta si se pasa del límite. */
function nombreArchivoPdf_(ids) {
  const limpio = String(ids).replace(/[\\/:*?"<>|]/g, '-').trim();
  if (limpio === '') return 'SIN-ID';
  return limpio.length > PDF_LARGO_MAX_NOMBRE ? limpio.substring(0, PDF_LARGO_MAX_NOMBRE) + '...' : limpio;
}
