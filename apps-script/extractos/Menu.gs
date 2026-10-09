/**
 * Menú "Saldos" en la barra de EXTRACTOS BANCARIOS, para no entrar a Apps Script.
 * Aparece solo al abrir la planilla: después de pegar cambios, recargá la pestaña del navegador.
 *
 * Las fechas se piden y se muestran siempre en formato yyyy-MM-dd (2026-09-30).
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Saldos')
    .addItem('Correr todo (ayer)', 'menuCorrerTodo')
    .addSeparator()
    .addItem('Importar archivos de Drive', 'menuImportarArchivos')
    .addItem('Actualizar saldos (ayer)', 'menuSaldosAyer')
    .addItem('Actualizar saldos (otra fecha)…', 'menuSaldosFecha')
    .addSeparator()
    .addItem('Pasar a Base CP (ayer)', 'menuBaseCPAyer')
    .addItem('Pasar a Base CP (otra fecha)…', 'menuBaseCPFecha')
    .addItem('Recalcular rango…', 'menuRecalcular')
    .addItem('Cargar saldos desde CSV…', 'abrirCargaCSV')
    .addItem('Aplicar correcciones manuales ahora', 'menuAplicarCorrecciones')
    .addItem('Repasar últimos días pendientes', 'menuRepasar')
    .addSeparator()
    .addItem('Cierre de mes…', 'abrirCierre')
    .addItem('Verificar último cierre', 'menuVerificarCierre')
    .addSeparator()
    .addSubMenu(ui.createMenu('QuickBooks')
      .addItem('Conectar LLC…', 'menuConectarQBOLLC')
      .addItem('Conectar SRL…', 'menuConectarQBOSRL')
      .addItem('Estado de la conexión', 'menuEstadoQBO')
      .addItem('Listar cuentas de QBO', 'menuListarQBO')
      .addItem('Probar QBO (ayer)', 'menuProbarQBO')
      .addSeparator()
      .addItem('Desconectar LLC', 'menuDesconectarQBOLLC')
      .addItem('Desconectar SRL', 'menuDesconectarQBOSRL'))
    .addSubMenu(ui.createMenu('Colppy')
      .addItem('Listar cuentas y bancos de Colppy', 'menuListarColppy')
      .addItem('Probar Colppy (ayer)', 'menuProbarColppy')
      .addItem('Probar Colppy (otra fecha)…', 'menuProbarColppyFecha'))
    .addSeparator()
    .addItem('Ver estado del último día', 'menuEstado')
    .addItem('Probar valuación de FCIs', 'menuProbarFCI')
    .addItem('Listar cuentas de Mercury', 'menuListarMercury')
    .addSeparator()
    .addItem('Programar corridas automáticas', 'menuCrearTriggers')
    .addItem('Ver corridas programadas', 'menuVerTriggers')
    .addToUi();
}

// ---------- ACCIONES ----------
function menuCorrerTodo() {
  correrConAviso('Corriendo todo', () => {
    pipeline();
    actualizarSaldos();
    actualizarBaseCP();
    aplicarCorrecciones();
  }, 'Listo: importación, saldos, Base CP de ayer y correcciones pendientes.');
}

function menuImportarArchivos() {
  correrConAviso('Importando archivos', importarArchivos,
    'Listo. Revisá las hojas *_raw y la carpeta de Drive: lo que falló queda sin mover.');
}

function menuSaldosAyer() {
  correrConAviso('Actualizando saldos', () => actualizarSaldos(),
    'Listo. Mirá la hoja Saldos: fuente y estado dicen de dónde salió cada número.');
}

function menuSaldosFecha() {
  const f = pedirFecha('¿De qué fecha querés los saldos?');
  if (!f) return;
  correrConAviso('Actualizando saldos del ' + f, () => actualizarSaldos(f),
    'Listo. Ojo: Mercury y Kraken solo se leen en la corrida de ayer, para otra fecha se arrastran.');
}

function menuBaseCPAyer() {
  correrConAviso('Escribiendo Base CP', () => actualizarBaseCP(),
    'Listo. Verde = leído al día o corregido a mano; beige = el resto.');
}

function menuBaseCPFecha() {
  const f = pedirFecha('¿Qué fecha querés pasar a Base CP?');
  if (!f) return;
  correrConAviso('Escribiendo Base CP del ' + f, () => actualizarBaseCP(f), 'Listo.');
}

function menuRecalcular() {
  const desde = pedirFecha('Recalcular desde qué fecha');
  if (!desde) return;
  const hasta = pedirFecha('Hasta qué fecha');
  if (!hasta) return;
  if (hasta < desde) {
    SpreadsheetApp.getUi().alert('La fecha de fin es anterior a la de inicio.');
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Recalcular',
    `Se van a rehacer saldos y bloques del ${desde} al ${hasta}. ` +
    'Las correcciones manuales (columna H) no se tocan. ¿Sigo?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  correrConAviso('Recalculando', () => recalcular(desde, hasta), 'Listo.');
}

function menuAplicarCorrecciones() {
  correrConAviso('Aplicando correcciones manuales', aplicarCorrecciones, 'Listo.');
}

function menuRepasar() {
  correrConAviso('Repasando últimos días', repasarPendientes, 'Listo.');
}

function menuEstado() {
  const sh = hojaSaldos();
  const tz = Session.getScriptTimeZone();
  const filas = leerSaldos(sh).filter(r => r[0] instanceof Date);
  if (!filas.length) {
    SpreadsheetApp.getUi().alert('La hoja Saldos todavía está vacía.');
    return;
  }
  const ultima = filas
    .map(r => Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'))
    .sort()
    .pop();
  const delDia = filas.filter(r => Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') === ultima);
  const corregidas = delDia.filter(r => !vacio(r[COL_MANUAL - 1]));
  const resto = delDia.filter(r => vacio(r[COL_MANUAL - 1]));
  const leidas = resto.filter(r => String(r[5]).indexOf('al día') === 0);
  const arrastradas = resto.filter(r => String(r[4]).trim().toLowerCase() === FUENTE_ARRASTRE);
  const sinConfirmar = resto.filter(r => leidas.indexOf(r) < 0 && arrastradas.indexOf(r) < 0);

  const linea = r => `• ${r[1]} ${r[2]}: ${r[3]} — ${r[5]}`;
  const lineaManual = r => `• ${r[1]} ${r[2]}: ${r[COL_MANUAL - 1]} (fuente: ${r[3]}) — ${r[COL_MOTIVO - 1] || 'sin motivo'}`;
  const texto = [
    `Último día cargado: ${ultima} (${delDia.length} cuentas)`,
    '',
    `Leídas al día (${leidas.length}):`,
    leidas.map(linea).join('\n') || '—',
    '',
    `Leídas sin confirmar, en beige (${sinConfirmar.length}):`,
    sinConfirmar.map(linea).join('\n') || '—',
    '',
    `Corregidas a mano (${corregidas.length}):`,
    corregidas.map(lineaManual).join('\n') || '—',
    '',
    `Arrastradas (${arrastradas.length}):`,
    arrastradas.map(linea).join('\n') || '—'
  ].join('\n');
  SpreadsheetApp.getUi().alert('Estado de los saldos', texto, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuVerificarCierre() {
  try {
    const f = ultimoCierrePasado();
    SpreadsheetApp.getUi().alert('Cierre del ' + fmtCorto(f), verificarCierre(f, true),
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function menuProbarFCI() {
  try {
    SpreadsheetApp.getUi().alert('Valuación de FCIs', probarCAFCI(), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function menuListarMercury() {
  try {
    const texto = cuentasMercury()
      .map(c => `${c.status === 'active' ? '•' : '(archivada)'} ${c.name} — ${c.currentBalance}\n${c.id}`)
      .join('\n\n');
    SpreadsheetApp.getUi().alert('Cuentas de Mercury',
      texto + '\n\nLos ids de las que suman en "Mercury" van en la propiedad MERCURY_CUENTAS.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('No pude leer Mercury: ' + e.message);
  }
}

// ---------- QUICKBOOKS ----------
function menuConectarQBOLLC() { menuConectarQBO('LLC'); }
function menuConectarQBOSRL() { menuConectarQBO('SRL'); }

function menuConectarQBO(empresa) {
  try {
    mostrarConexionQBO(empresa);
  } catch (e) {
    SpreadsheetApp.getUi().alert('No se puede conectar: ' + e.message);
  }
}

function menuEstadoQBO() {
  try {
    SpreadsheetApp.getUi().alert('QuickBooks', estadoQBO(), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function menuListarQBO() {
  correrConAviso('Leyendo cuentas de QuickBooks', listarCuentasQBO, 'Listo: hoja Cuentas QBO.');
}

function menuProbarQBO() {
  try {
    SpreadsheetApp.getUi().alert('Prueba de QuickBooks', probarQBO(), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function menuDesconectarQBOLLC() { menuDesconectarQBO('LLC'); }
function menuDesconectarQBOSRL() { menuDesconectarQBO('SRL'); }

function menuDesconectarQBO(empresa) {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Desconectar ' + empresa,
    `Se revoca el acceso a QuickBooks de la ${empresa}. Sus cuentas vuelven a ser manuales ` +
    '(se arrastran y aparecen en el cierre de mes) hasta que se conecte de nuevo. ¿Sigo?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  correrConAviso('Desconectando ' + empresa, () => desconectarQBO(empresa), empresa + ' desconectada.');
}

// ---------- COLPPY ----------
function menuListarColppy() {
  correrConAviso('Leyendo Colppy', listarColppy, 'Listo: hojas Cuentas Colppy y Bancos Colppy.');
}

function menuProbarColppy() {
  try {
    SpreadsheetApp.getUi().alert('Prueba de Colppy', probarColppy(), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function menuProbarColppyFecha() {
  const f = pedirFecha('¿A qué fecha querés los saldos de Colppy?');
  if (!f) return;
  try {
    SpreadsheetApp.getUi().alert('Prueba de Colppy', probarColppy(f), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

// ---------- CORRIDAS ----------
function menuCrearTriggers() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Programar corridas',
    'Se reemplazan las corridas de este proceso: quedan dos por día (6 y 18) y el registro de ' +
    'correcciones manuales de la hoja Saldos. Otros triggers del proyecto no se tocan. ¿Sigo?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  correrConAviso('Programando', crearTriggersCompleto, 'Listo: corridas a las 6 y a las 18.');
}

function menuVerTriggers() {
  const t = ScriptApp.getProjectTriggers()
    .map(x => `• ${x.getHandlerFunction()} (${x.getEventType()})`)
    .join('\n');
  SpreadsheetApp.getUi().alert('Corridas programadas', t || 'No hay ninguna programada.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---------- AUXILIARES ----------
function pedirFecha(titulo) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(titulo, 'Formato: yyyy-MM-dd (ejemplo 2026-09-30)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const f = r.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f)) {
    ui.alert('Fecha inválida: ' + f + '. Tiene que ser yyyy-MM-dd.');
    return null;
  }
  return f;
}

// Si fn devuelve un texto, se muestra ese texto en vez del mensaje de éxito fijo.
function correrConAviso(mensaje, fn, exito) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(mensaje + '…', 'Saldos', -1);
  try {
    const r = fn();
    ss.toast(typeof r === 'string' && r ? r : exito, 'Saldos', 10);
  } catch (e) {
    ss.toast('Falló', 'Saldos', 3);
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
