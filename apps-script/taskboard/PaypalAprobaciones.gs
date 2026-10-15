/**
 * APROBACIONES PAYPAL (Dai / Cami) + RESPUESTA "ABONADO"
 *
 * Los pagos por PayPal solo se abonan con aprobación de Dai (Daiana Solarz) o Cami (Camila De
 * Angelis), que llega por mail en el hilo "[Pagos LB] - <ID> - Aprobación Paypal". Este archivo:
 *
 *   1) Escribe el estado de cada tarea Paypal AL FINAL de la columna COMENTARIO de la agrupada del
 *      día, como "PAYPAL: <estado>" (ver MARCA EN COMENTARIO). Se llama sola al generar o
 *      actualizar la agrupada (SOLO LEE Gmail, no manda nada) y también desde el menú.
 *   2) Desde el menú "PayPal: revisar aprobaciones y confirmar abonados", responde A TODOS en el
 *      hilo "Buenas, abonado. Saludos." + firma, en cada tarea Paypal que ya tiene comprobante en
 *      la agrupada y aprobación de Dai o Cami. Antes de mandar muestra la lista y pide confirmación.
 *
 * ESTADOS (por tarea):
 *   SIN HILO DE APROBACIÓN   no hay ningún mail con el ID y "Aprobación" en el asunto.
 *   PENDIENTE                hay hilo, pero Dai y Cami no respondieron.
 *   REVISAR RESPUESTA        la última respuesta de Dai o Cami NO es una aprobación (una pregunta,
 *                            "Podemos avanzar!", etc.). El texto queda en la nota de la celda.
 *   LISTO PARA ABONAR        la última respuesta de Dai o Cami es una aprobación.
 *   A ENVIAR "ABONADO"       listo + la fila tiene comprobante + todavía no respondiste "abonado".
 *   YA ABONADO               hay un mensaje tuyo con "abonado" posterior a la aprobación.
 *   ⚠️ PAGADO SIN APROBACIÓN  la fila tiene comprobante pero no hay aprobación. No se manda nada.
 *   ERROR GMAIL              falló la lectura de Gmail. La agrupada se genera igual.
 *
 * QUÉ CUENTA COMO APROBACIÓN
 *   Solo el ÚLTIMO mensaje de Dai o Cami del hilo, y solo el texto que escribieron ellas: se corta
 *   en "El ... escribió:" / "On ... wrote:", porque la cita trae el "Pago aprobado" de María, que
 *   no habilita el pago. Es aprobación si contiene "OK PARA ABONAR", o si una oración arranca con
 *   "Aprobado" / "Pago(s) aprobado(s)", en ambos casos sin signo de pregunta después.
 *   Si después de aprobar una de ellas pregunta algo, gana la pregunta: queda REVISAR.
 *
 * MARCA EN COMENTARIO
 *   La marca va SIEMPRE al final, después de " | ": "lo que escribiste | PAYPAL: Aprobado". Al final
 *   y no al principio porque el principio es de la marca de cancelación (Cancelaciones.gs). Se
 *   reescribe en cada corrida, así que no queda vieja ni se duplica. Si escribís vos un texto que
 *   termine en "PAYPAL: ...", el script lo trata como suyo y lo reemplaza.
 *   No se arrastra al día siguiente: comentariosAgrupAnteriores_ (Pagoslb.gs) la saca con
 *   quitarMarcaPaypal_ antes de armar COMENTARIO ANTERIOR. Si no la sacara, COMENTARIO ANTERIOR
 *   cambiaría con cada estado de PayPal y, como es parte de la clave de grupo, la fila perdería su
 *   CHECK y su ESTADO al regenerar.
 *   El detalle (quién respondió y qué) va en la nota de la celda COMENTARIO; el fondo de la celda
 *   toma el color del estado.
 *
 * DEPENDE de Hojaagrupada.gs (AGRUP_PREFIJO, colsAgrup_, valAgrup_) y de Pagoslb.gs
 * (formatearNombreHoja_).
 */

const PP_TIPO = 'Paypal';
const PP_APROBADORAS = ['daiana.solarz@agenciaeleven.com', 'camila.deangelis@agenciaeleven.com'];
const PP_MI_MAIL = 'lucas.leone@agenciaeleven.com';
const PP_MARCA = 'PAYPAL: ';
const PP_TEXTO_ABONADO = 'Buenas, abonado. Saludos.';
const PP_FIRMA = 'Lucas Agustín Leone\nAdministrative Analyst Jr.\neleven.agency';
const PP_DIAS_BUSQUEDA = 120;

/** La marca de PayPal: el último tramo del comentario que arranca con "PAYPAL: ". */
const RE_MARCA_PAYPAL = /(?:^|\s*\|\s*)PAYPAL: [^|]*$/;

const PP_SIN_HILO   = 'SIN HILO DE APROBACIÓN';
const PP_PENDIENTE  = 'PENDIENTE';
const PP_REVISAR    = 'REVISAR RESPUESTA';
const PP_LISTO      = 'LISTO PARA ABONAR';
const PP_A_ENVIAR   = 'A ENVIAR "ABONADO"';
const PP_ABONADO    = 'YA ABONADO';
const PP_SIN_APROB  = '⚠️ PAGADO SIN APROBACIÓN';
const PP_ERROR      = 'ERROR GMAIL';

/**
 * Estilo de cada estado. El orden es la gravedad: una fila con varios IDs toma el peor.
 * etiqueta = lo que se ve en el comentario (corto). El estado completo y el texto del mail van en
 * la nota. fondo = color de la celda COMENTARIO.
 */
const PP_GRAVEDAD = [
  { estado: PP_SIN_APROB, etiqueta: '⚠ Pagado sin aprob.', fondo: '#f4cccc' },
  { estado: PP_ERROR,     etiqueta: 'Error Gmail',         fondo: '#f4cccc' },
  { estado: PP_REVISAR,   etiqueta: 'Revisar respuesta',   fondo: '#fff2cc' },
  { estado: PP_SIN_HILO,  etiqueta: 'Sin hilo',            fondo: '#fff2cc' },
  { estado: PP_PENDIENTE, etiqueta: 'Pendiente',           fondo: '#f3f3f3' },
  { estado: PP_A_ENVIAR,  etiqueta: 'Enviar "abonado"',    fondo: '#cfe2f3' },
  { estado: PP_LISTO,     etiqueta: 'Aprobado',            fondo: '#d9ead3' },
  { estado: PP_ABONADO,   etiqueta: '✓ Abonado',           fondo: '#d9ead3' }
];

/** Donde empieza la cita del mail anterior. Todo lo que viene después no lo escribió el remitente. */
const PP_CORTES_CITA = [
  /\bEl (lun|mar|mi[eé]|jue|vie|s[aá]b|dom)\.?,[\s\S]{0,300}?escribi[oó]\s*:/i,
  /\bOn (Mon|Tue|Wed|Thu|Fri|Sat|Sun),[\s\S]{0,300}?wrote\s*:/i,
  /-{5,}\s*(Forwarded message|Mensaje reenviado)/i,
  /(^|\n)\s*>/
];

/* ══════════════════ MENÚ ══════════════════ */

function revisarAprobacionesPaypal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nombre = AGRUP_PREFIJO + formatearNombreHoja_(new Date());
  const hoja = ss.getSheetByName(nombre);

  if (!hoja) {
    ui.alert('No existe la hoja "' + nombre + '". Generá primero la agrupada.');
    return;
  }

  const filas = refrescarColumnaPaypal_(hoja);
  if (filas.length === 0) {
    ui.alert('No hay filas Paypal en "' + nombre + '".');
    return;
  }

  const resumen = resumenPaypal_(filas);
  const envios = enviosAbonadoPendientes_(filas);

  if (envios.length === 0) {
    ui.alert(resumen + '\n\nNo hay ningún "abonado" para enviar.');
    return;
  }

  // El mail sale con la cuenta de quien corre el script: si no sos vos, no se manda nada.
  const yo = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (yo !== PP_MI_MAIL) {
    ui.alert(resumen + '\n\nHay ' + envios.length + ' "abonado" para enviar, pero el script está ' +
             'corriendo con "' + (yo || 'usuario desconocido') + '" y no con ' + PP_MI_MAIL +
             '. No se manda nada.');
    return;
  }

  const lista = envios.map(e =>
    '• ' + e.ids.join(', ') + '\n   asunto: ' + e.asunto + '\n   a: ' + e.destinatarios
  ).join('\n');

  const resp = ui.alert(
    resumen + '\n\n' +
    'Se va a responder A TODOS en ' + envios.length + ' hilo(s) con:\n' +
    '"' + PP_TEXTO_ABONADO + '" + firma\n\n' + lista + '\n\n¿Confirmás?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const enviados = [];
  const errores = [];
  envios.forEach(e => {
    try {
      e.mensaje.replyAll(PP_TEXTO_ABONADO + '\n\n' + PP_FIRMA);
      enviados.push(e.ids.join(', '));
    } catch (err) {
      errores.push(e.ids.join(', ') + ': ' + err.message);
    }
  });

  // Se relee el hilo: las tareas recién respondidas pasan a YA ABONADO.
  refrescarColumnaPaypal_(hoja);

  ui.alert(
    '"Abonado" enviado en ' + enviados.length + ' hilo(s)' +
    (enviados.length ? ':\n' + enviados.map(x => '• ' + x).join('\n') : '.') +
    (errores.length ? '\n\n❌ No se pudo enviar:\n' + errores.join('\n') : '')
  );
}

/* ══════════════════ MARCA EN COMENTARIO ══════════════════ */

/**
 * Lee las filas Paypal de la agrupada, consulta Gmail por cada tarea y escribe la marca en
 * COMENTARIO. Solo lee Gmail. Devuelve las filas analizadas para que el menú arme los envíos.
 * Conserva el nombre de cuando escribía una columna aparte: lo llama Hojaagrupada.gs.
 */
function refrescarColumnaPaypal_(hoja) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];

  const C = colsAgrup_(hoja);
  if (C.TIPO < 0 || C.IDS < 0 || C.COMENT < 0) return [];
  const valores = hoja.getRange(2, 1, ultima - 1, C.ANCHO).getValues();
  // El comprobante es un enlace rich text ("Ver comprobante"): getValues() no trae la URL.
  const enlaces = C.COMPROBANTE > 0
    ? hoja.getRange(2, C.COMPROBANTE, ultima - 1, 1).getRichTextValues() : null;

  const cache = {};
  const filas = [];

  valores.forEach((v, i) => {
    if (String(valAgrup_(v, C, 'TIPO')).trim() !== PP_TIPO) return;
    const ids = String(valAgrup_(v, C, 'IDS') || '').split(',').map(s => s.trim()).filter(String);
    if (ids.length === 0) return;   // banda de tipo

    const rt = enlaces ? enlaces[i][0] : null;
    const comprobante = !!((rt && rt.getLinkUrl()) ||
                           String(valAgrup_(v, C, 'COMPROBANTE') || '').trim());

    const porId = {};
    ids.forEach(id => {
      if (!cache[id]) cache[id] = analizarIdPaypal_(id);
      porId[id] = estadoConComprobante_(cache[id], comprobante);
    });

    filas.push({
      fila: i + 2,
      ids: ids,
      comprobante: comprobante,
      porId: porId,
      comentario: valAgrup_(v, C, 'COMENT')
    });
  });

  escribirMarcaPaypal_(hoja, filas, C.COMENT);
  return filas;
}

function estadoConComprobante_(info, comprobante) {
  if (!comprobante) return { estado: info.estado, texto: info.texto, info: info };

  if (info.estado === PP_LISTO) return { estado: PP_A_ENVIAR, texto: info.texto, info: info };

  if ([PP_PENDIENTE, PP_REVISAR, PP_SIN_HILO].indexOf(info.estado) !== -1) {
    return {
      estado: PP_SIN_APROB,
      texto: 'La fila tiene comprobante pero no hay aprobación de Dai o Cami.\n' + info.texto,
      info: info
    };
  }
  return { estado: info.estado, texto: info.texto, info: info };
}

/** Escribe "PAYPAL: <estado>" al final de COMENTARIO, la nota con el detalle y el color. */
function escribirMarcaPaypal_(hoja, filas, colComent) {
  const etiqueta = e => (PP_GRAVEDAD.find(g => g.estado === e) || { etiqueta: e }).etiqueta;

  filas.forEach(f => {
    const estados = f.ids.map(id => f.porId[id].estado);
    const peor = PP_GRAVEDAD.find(g => estados.indexOf(g.estado) !== -1);
    const marca = PP_MARCA + (f.ids.length === 1
      ? etiqueta(estados[0])
      : f.ids.map(id => id + ' ' + etiqueta(f.porId[id].estado)).join('; '));
    const nota = f.ids.map(id =>
      (f.ids.length > 1 ? id + '\n' : '') + f.porId[id].estado + '\n' + f.porId[id].texto
    ).join('\n\n');

    const celda = hoja.getRange(f.fila, colComent)
      .setValue(ponerMarcaPaypal_(f.comentario, marca))
      .setNote(nota);
    if (peor) celda.setBackground(peor.fondo);
  });
}

/** Comentario sin la marca de PayPal. La usa también comentariosAgrupAnteriores_ (Pagoslb.gs). */
function quitarMarcaPaypal_(comentario) {
  return String(comentario == null ? '' : comentario).replace(RE_MARCA_PAYPAL, '').trim();
}

/** Pone la marca al final, reemplazando la anterior si la había. Lo tuyo queda antes de " | ". */
function ponerMarcaPaypal_(comentario, marca) {
  const resto = quitarMarcaPaypal_(comentario);
  return resto ? resto + ' | ' + marca : marca;
}

/* ══════════════════ LECTURA DE GMAIL ══════════════════ */

function analizarIdPaypal_(id) {
  try {
    const reId = new RegExp('(^|[^A-Z0-9])' + escaparRegex_(ppNorm_(id)) + '(?![0-9])');
    const hilos = GmailApp.search('subject:"' + id + '" newer_than:' + PP_DIAS_BUSQUEDA + 'd', 0, 20)
      .filter(h => {
        const asunto = ppNorm_(h.getFirstMessageSubject());
        return reId.test(asunto) && asunto.indexOf('APROBACION') !== -1;
      });

    if (hilos.length === 0) {
      return {
        estado: PP_SIN_HILO,
        texto: 'No hay ningún mail con "' + id + '" y "Aprobación" en el asunto (últimos ' +
               PP_DIAS_BUSQUEDA + ' días).'
      };
    }

    const msgs = [];
    hilos.forEach(h => h.getMessages().forEach(m => msgs.push(m)));
    msgs.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

    let ultima = null;
    msgs.forEach(m => { if (PP_APROBADORAS.indexOf(remitente_(m)) !== -1) ultima = m; });

    if (!ultima) {
      return { estado: PP_PENDIENTE, texto: 'Hay hilo, pero Dai y Cami todavía no respondieron.' };
    }

    const propio = textoPropio_(ultima.getPlainBody());
    const fecha = Utilities.formatDate(ultima.getDate(), Session.getScriptTimeZone(), 'dd-MM HH:mm');
    const cita = nombreAprobadora_(remitente_(ultima)) + ' (' + fecha + '): ' + recortar_(propio, 300);

    if (!esAprobacion_(propio)) return { estado: PP_REVISAR, texto: cita };

    const t0 = ultima.getDate().getTime();
    const abonado = msgs.some(m =>
      remitente_(m) === PP_MI_MAIL &&
      m.getDate().getTime() > t0 &&
      /ABONADO/.test(ppNorm_(textoPropio_(m.getPlainBody())))
    );

    return {
      estado: abonado ? PP_ABONADO : PP_LISTO,
      texto: cita,
      mensaje: ultima,
      asunto: ultima.getSubject(),
      hiloId: ultima.getThread().getId()
    };
  } catch (err) {
    return { estado: PP_ERROR, texto: 'Error leyendo Gmail: ' + err.message };
  }
}

/** Un envío por hilo: una fila con varios IDs del mismo hilo se responde una sola vez. */
function enviosAbonadoPendientes_(filas) {
  const porHilo = {};
  filas.forEach(f => f.ids.forEach(id => {
    const r = f.porId[id];
    if (r.estado !== PP_A_ENVIAR) return;
    const k = r.info.hiloId;
    if (!porHilo[k]) {
      porHilo[k] = {
        mensaje: r.info.mensaje,
        asunto: r.info.asunto,
        destinatarios: destinatariosReplyAll_(r.info.mensaje),
        ids: []
      };
    }
    if (porHilo[k].ids.indexOf(id) === -1) porHilo[k].ids.push(id);
  }));
  return Object.keys(porHilo).map(k => porHilo[k]);
}

function resumenPaypal_(filas) {
  const cuenta = {};
  filas.forEach(f => f.ids.forEach(id => {
    const e = f.porId[id].estado;
    cuenta[e] = (cuenta[e] || 0) + 1;
  }));
  return 'TAREAS PAYPAL DE LA AGRUPADA\n' +
    PP_GRAVEDAD.filter(g => cuenta[g.estado])
      .map(g => '  ' + g.estado + ': ' + cuenta[g.estado]).join('\n');
}

/* ══════════════════ TEXTO ══════════════════ */

/** Texto que escribió el remitente, sin la cita del mail anterior. */
function textoPropio_(cuerpo) {
  const t = String(cuerpo == null ? '' : cuerpo).replace(/\r/g, '');
  let corte = t.length;
  PP_CORTES_CITA.forEach(re => {
    const m = t.match(re);
    if (m && m.index < corte) corte = m.index;
  });
  return t.substring(0, corte).trim();
}

function esAprobacion_(texto) {
  const t = ppNorm_(texto);
  if (/OK\s+PARA\s+ABONAR\b(?!\s*\?)/.test(t)) return true;
  return /(^|[.!?¡¿:\n]\s*)(PAGOS?\s+)?APROBADOS?\b(?!\s*\?)/.test(t);
}

function ppNorm_(s) {
  return String(s == null ? '' : s).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function escaparRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extraerMails_(s) {
  return (String(s == null ? '' : s).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [])
    .map(x => x.toLowerCase());
}

function remitente_(m) {
  return extraerMails_(m.getFrom())[0] || '';
}

function nombreAprobadora_(mail) {
  if (mail === PP_APROBADORAS[0]) return 'Dai';
  if (mail === PP_APROBADORAS[1]) return 'Cami';
  return mail;
}

/** A quiénes les llega el replyAll: remitente + To + Cc, sin vos. */
function destinatariosReplyAll_(m) {
  const todos = extraerMails_([m.getFrom(), m.getTo(), m.getCc()].join(','));
  return todos.filter((x, i) => x !== PP_MI_MAIL && todos.indexOf(x) === i).join(', ');
}

function recortar_(s, n) {
  const t = String(s || '').replace(/\s+/g, ' ').trim();
  return t.length > n ? t.substring(0, n) + '…' : t;
}
