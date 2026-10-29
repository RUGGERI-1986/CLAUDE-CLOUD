/**
 * DECISIÓN — vista filtrable de la hoja Condiciones.
 * No modifica Condiciones. Crea 2 hojas:
 *   - "Decisión_Datos" (oculta): fórmulas vivas que separan la columna L y reúnen A, B, I, J, M.
 *   - "Decisión": filtros desplegables (con varios proyectos a la vez) + resultados + envío de la fila elegida a FILTRO.
 *
 * Instalación: ejecutar 1 vez setupDecision() y 1 vez setupTriggers().
 * Volver a ejecutar setupDecision() reconstruye la hoja "Decisión" (los filtros vuelven a "Todos").
 */

const CONFIG_DECISION = {
  HOJA_CONDICIONES: 'Condiciones',
  HOJA_VISTA: 'Decisión',
  HOJA_DATOS: 'Decisión_Datos',
  HOJA_FILTRO: 'FILTRO',
  HOJA_AUXILIAR: 'Auxiliar',
  HOJA_PARAMETROS: 'Parámetros',
  HOJA_BASE: 'BASE',
  FILA_ENCABEZADO: 15,
  FILA_RESULTADOS: 16,
  FILAS_RESULTADOS: 500,
  MAX_LINKS: 50,
  TODOS: 'Todos',
  TODOS_PROYECTOS: 'Todos los proyectos'
};

/* ------------------------------------------------------------------ */
/* SETUP                                                               */
/* ------------------------------------------------------------------ */

function setupDecision() {
  const ss = SpreadsheetApp.getActive();
  const cond = ss.getSheetByName(CONFIG_DECISION.HOJA_CONDICIONES);
  if (!cond) throw new Error('No existe la hoja "' + CONFIG_DECISION.HOJA_CONDICIONES + '".');
  const columnas = columnasCondiciones_(cond);

  const datos = obtenerOCrearHojaDecision_(ss, CONFIG_DECISION.HOJA_DATOS);
  const vista = obtenerOCrearHojaDecision_(ss, CONFIG_DECISION.HOJA_VISTA);
  armarDatos_(datos, columnas);
  armarVista_(vista);
  datos.hideSheet();
  ss.setActiveSheet(vista);
}

/**
 * Ubica las columnas de Condiciones por su encabezado (no por letra), así no se rompe si insertan o
 * mueven columnas. "Tráfico Mínimo" es opcional: si no existe, se toma 0.
 */
function columnasCondiciones_(cond) {
  const enc = cond.getRange(1, 1, 1, cond.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const buscar = function (nombre, opcional) {
    const i = enc.indexOf(nombre);
    if (i < 0 && !opcional) {
      throw new Error('Condiciones no tiene la columna "' + nombre + '". Encabezados encontrados: ' + enc.filter(String).join(' | '));
    }
    return i;
  };
  return {
    cliente: buscar('Cliente'),
    proyecto: buscar('Proyecto'),
    idioma: buscar('Idioma Comunicación Cliente'),
    paises: buscar('Países Objetivo'),
    tipo: buscar('Tipo de Servicio - Cantidad'),
    trafico: buscar('Tráfico Mínimo', true)
  };
}

function letraColumna_(i) {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActive();
  const propios = ['alEditarDecision', 'crearMenuDecision'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (propios.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('alEditarDecision').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('crearMenuDecision').forSpreadsheet(ss).onOpen().create();
}

/* ------------------------------------------------------------------ */
/* HOJA OCULTA: DATOS                                                  */
/* ------------------------------------------------------------------ */

function armarDatos_(sh, columnas) {
  sh.clear();
  const C = "'" + CONFIG_DECISION.HOJA_CONDICIONES + "'!";
  const rango = function (i) { const l = letraColumna_(i); return C + l + '2:' + l; };
  const sinProyecto = rango(columnas.proyecto) + '=""';
  const L = rango(columnas.tipo);

  sh.getRange('A1:G1').setValues([[
    'Cliente', 'Proyecto', 'Intervalo', 'Cantidad', 'Tráfico mínimo', 'Países', 'Idioma'
  ]]);
  sh.getRange('A2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + rango(columnas.cliente) + '))');
  sh.getRange('B2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + rango(columnas.proyecto) + '))');
  sh.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(' + sinProyecto + ',"",IFERROR(REGEXEXTRACT(' + L + '&"","(\\d+-\\d+)"),"")))');
  sh.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(' + sinProyecto + ',"",IFERROR(VALUE(REGEXEXTRACT(' + L + '&"","->\\s*(\\d+)")),0)))');
  sh.getRange('E2').setFormula(columnas.trafico < 0
    ? '=ARRAYFORMULA(IF(' + sinProyecto + ',"",0))'
    : '=ARRAYFORMULA(IF(' + sinProyecto + ',"",IFERROR(VALUE(' + rango(columnas.trafico) + '),0)))');
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + rango(columnas.paises) + '))');
  sh.getRange('G2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + rango(columnas.idioma) + '))');

  // Listas para los desplegables (los proyectos los arma el script, según el cliente)
  const todos = '"' + CONFIG_DECISION.TODOS + '"';
  sh.getRange('I1:K1').setValues([['Lista clientes', 'Lista intervalos', 'Lista países']]);
  sh.getRange('I2').setFormula('={' + todos + ';SORT(UNIQUE(FILTER(A2:A,A2:A<>"")))}');
  sh.getRange('J2').setFormula('={' + todos + ';SORT(UNIQUE(FILTER(C2:C,C2:C<>"")))}');
  sh.getRange('K2').setFormula(
    '=ARRAYFORMULA({' + todos + ';SORT(UNIQUE(FILTER(TRIM(FLATTEN(SPLIT(TEXTJOIN(",",TRUE,F2:F),","))),' +
    'TRIM(FLATTEN(SPLIT(TEXTJOIN(",",TRUE,F2:F),",")))<>"")))})');
  sh.setFrozenRows(1);
}

/* ------------------------------------------------------------------ */
/* HOJA VISIBLE: DECISIÓN                                              */
/* ------------------------------------------------------------------ */

function armarVista_(sh) {
  const T = CONFIG_DECISION.TODOS;
  const TP = CONFIG_DECISION.TODOS_PROYECTOS;
  const ini = CONFIG_DECISION.FILA_RESULTADOS;
  const hdr = CONFIG_DECISION.FILA_ENCABEZADO;
  const nRes = CONFIG_DECISION.FILAS_RESULTADOS;
  const D = "'" + CONFIG_DECISION.HOJA_DATOS + "'!";
  const datos = sh.getParent().getSheetByName(CONFIG_DECISION.HOJA_DATOS);

  // Colores: los mismos de FILTRO
  const NAVY = '#1F2A3C', AZUL = '#DBE7FB', AZUL_SUAVE = '#EEF3FC', PANEL = '#F8F6F4';
  const VERDE = '#2E7D4F', VERDE_SUAVE = '#E9F5EE', ROJO_SUAVE = '#FDE8E8', SUAVE = '#6B7280', BORDE = '#B6C8EA';

  // Limpieza total (no se toca ninguna otra hoja)
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange('A1:Z1000').clearDataValidations();
  sh.setFrozenRows(0);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:J' + (ini + nRes)).setFontSize(10).setVerticalAlignment('middle');

  // Anchos y altos
  [44, 190, 200, 120, 95, 105, 330, 20, 200, 95].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  [56, 30, 14, 34].forEach(function (h, i) { sh.setRowHeight(i + 1, h); });
  sh.setRowHeights(5, 7, 28);
  sh.setRowHeight(12, 14);
  sh.setRowHeight(13, 42);
  sh.setRowHeight(14, 32);
  sh.setRowHeight(hdr, 30);
  sh.setRowHeights(ini, nRes, 24);

  // Título
  sh.getRange('A1:J1').merge().setBackground(NAVY);
  sh.getRange('A1').setValue('Decisión de asignación de medios')
    .setFontColor('#FFFFFF').setFontSize(20).setFontWeight('bold');
  sh.getRange('A2:J2').merge().setBackground(NAVY);
  sh.getRange('A2').setValue('Elegí en las celdas azules. Podés sumar varios proyectos. Los resultados se ordenan por cantidad pendiente.')
    .setFontColor('#D6DCE8').setFontSize(11);

  // Encabezados de sección
  sh.getRange('B4:G4').merge();
  sh.getRange('B4').setValue('1 · Qué querés ver').setFontSize(12).setFontWeight('bold');
  sh.getRange('I4:J4').merge();
  sh.getRange('I4').setValue('Resumen de lo filtrado').setFontSize(12).setFontWeight('bold');

  // Filtros
  sh.getRange('B5:B11').setValues([
    ['Cliente'], ['Proyectos elegidos'], ['Agregar o quitar proyecto'], ['Intervalo de métrica'],
    ['País objetivo'], ['Cantidad mínima'], ['Tráfico mínimo exigido (desde)']
  ]).setFontWeight('bold');
  sh.getRange('C5').setValue(T);
  sh.getRange('C6:G6').merge();
  sh.getRange('C6').setValue(TP);
  sh.getRange('C7').setValue('');
  sh.getRange('C8').setValue(T);
  sh.getRange('C9').setValue(T);
  sh.getRange('C10').setValue(1);
  sh.getRange('C11').setValue(0);
  [['C5'], ['C7'], ['C8'], ['C9'], ['C10'], ['C11']].forEach(function (a) {
    sh.getRange(a[0]).setBackground(AZUL).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('left')
      .setBorder(true, true, true, true, false, false, BORDE, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange('C6:G6').setBackground(AZUL_SUAVE).setFontSize(11).setFontWeight('bold').setWrap(true)
    .setBorder(true, true, true, true, false, false, BORDE, SpreadsheetApp.BorderStyle.SOLID);

  const pistas = [['D5', 'Los proyectos para sumar son sólo los de este cliente'],
    ['D7', 'Elegí uno para sumarlo; volvé a elegirlo para quitarlo'],
    ['D8', 'DR del medio: 0-20, 21-40, ...'],
    ['D9', 'Coincide si el país está en la lista del proyecto'],
    ['D10', '1 oculta los intervalos sin cantidad pendiente'],
    ['D11', 'Proyectos que exigen al menos este tráfico']];
  pistas.forEach(function (p) { sh.getRange(p[0]).setValue(p[1]).setFontSize(9).setFontColor(SUAVE); });

  // Resumen (panel gris)
  sh.getRange('I5:I8').setValues([['Combinaciones'], ['Proyectos distintos'], ['Clientes'], ['Links por conseguir']]);
  sh.getRange('J5:J8').setFormulas([
    ['=COUNTIF(B' + ini + ':B,"?*")'],
    ['=IFERROR(COUNTUNIQUE(C' + ini + ':C),0)'],
    ['=IFERROR(COUNTUNIQUE(B' + ini + ':B),0)'],
    ['=SUM(E' + ini + ':E)']
  ]);
  sh.getRange('I5:J8').setBackground(PANEL);
  sh.getRange('J5:J8').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('right');

  // Mensaje de estado
  sh.getRange('B13:J13').merge().setBackground(VERDE_SUAVE);
  sh.getRange('B13').setFormula(
    '=IF(B' + ini + '="","Sin resultados con estos filtros. Aflojá alguno (Cantidad mínima 0, o Todos).",' +
    'J5&" combinaciones de "&J6&" proyectos · "&J8&" links por conseguir")')
    .setFontSize(11).setFontWeight('bold').setWrap(true);

  // Resultados
  sh.getRange('B14:G14').merge();
  sh.getRange('B14').setValue('2 · Hacé clic en una fila y usá el menú Decisión → Enviar fila a FILTRO')
    .setFontSize(11).setFontWeight('bold');
  sh.getRange('A' + hdr).setValue('Nº');
  sh.getRange(hdr, 2, 1, 6).setValues([['Cliente', 'Proyecto', 'Intervalo de métrica', 'Cantidad', 'Tráfico mínimo', 'Países objetivo']]);
  sh.getRange(hdr, 1, 1, 7).setBackground(VERDE).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange('A' + hdr + ':A' + (ini + nRes)).setHorizontalAlignment('center');
  sh.getRange('D' + hdr + ':F' + (ini + nRes)).setHorizontalAlignment('center');
  sh.getRange('E' + ini + ':F' + (ini + nRes)).setNumberFormat('#,##0');
  sh.getRange('A' + ini + ':A' + (ini + nRes)).setFontSize(9).setFontColor(SUAVE);
  sh.getRange('A' + ini).setFormula('=ARRAYFORMULA(IF(B' + ini + ':B="","",ROW(B' + ini + ':B)-' + (ini - 1) + '))');

  const cond = function (celda, col) {
    return '((($C$' + celda + '="' + T + '")+(' + D + col + '2:' + col + '=$C$' + celda + '))>0)';
  };
  const proyectos = '((($C$6="' + TP + '")+REGEXMATCH(' + D + 'B2:B&"","^("&SUBSTITUTE(SUBSTITUTE($C$6,".","\\."),", ","|")&")$"))>0)';
  const pais = '((($C$9="' + T + '")+REGEXMATCH(' + D + 'F2:F&"","(^|,)\\s*"&$C$9&"\\s*(,|$)"))>0)';
  sh.getRange(ini, 2).setFormula(
    '=IFERROR(SORT(FILTER(' + D + 'A2:F,' + D + 'B2:B<>"",' +
    cond(5, 'A') + ',' + proyectos + ',' + cond(8, 'C') + ',' + pais + ',' +
    D + 'D2:D>=N($C$10),' + D + 'E2:E>=N($C$11)),4,FALSE,1,TRUE,2,TRUE),"")');

  // Formato condicional: mensaje en rojo si no hay resultados y filas alternadas
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B$' + ini + '=""')
      .setBackground(ROJO_SUAVE).setRanges([sh.getRange('B13:J13')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B' + ini + '<>"",ISEVEN(ROW()))')
      .setBackground('#F4F7FB').setRanges([sh.getRange(ini, 1, nRes, 7)]).build()
  ]);

  // Desplegables
  const regla = function (rango) {
    return SpreadsheetApp.newDataValidation().requireValueInRange(rango, true).setAllowInvalid(false).build();
  };
  sh.getRange('C5').setDataValidation(regla(datos.getRange('I2:I')));
  sh.getRange('C8').setDataValidation(regla(datos.getRange('J2:J')));
  sh.getRange('C9').setDataValidation(regla(datos.getRange('K2:K')));
  sh.getRange('C10:C11').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
  actualizarListaProyectos_(sh);

  sh.setTabColor('#2E6FDE');
}

/* ------------------------------------------------------------------ */
/* TRIGGERS (instalables, creados por setupTriggers)                   */
/* ------------------------------------------------------------------ */

function crearMenuDecision() {
  SpreadsheetApp.getUi().createMenu('Decisión')
    .addItem('Enviar fila a FILTRO', 'enviarAFiltro')
    .addItem('Limpiar filtros', 'limpiarFiltros')
    .addToUi();
}

/**
 * Hoja Decisión:
 *  - C5 (Cliente): limpia los proyectos elegidos y deja para sumar sólo los de ese cliente.
 *  - C7 (Agregar o quitar proyecto): suma el elegido a C6, o lo quita si ya estaba.
 * Las ediciones en FILTRO las atiende onEdit() de Autorrelleno.gs, que llama a cascadaFiltro_.
 */
function alEditarDecision(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG_DECISION.HOJA_VISTA) return;
  if (e.range.getColumn() !== 3 || e.range.getNumRows() !== 1) return;
  const fila = e.range.getRow();
  if (fila === 5) {
    sh.getRange('C6').setValue(CONFIG_DECISION.TODOS_PROYECTOS);
    sh.getRange('C7').setValue('');
    actualizarListaProyectos_(sh);
  } else if (fila === 7) {
    alternarProyecto_(sh);
  }
}

function alternarProyecto_(sh) {
  const TP = CONFIG_DECISION.TODOS_PROYECTOS;
  const picker = sh.getRange('C7');
  const elegido = String(picker.getValue()).trim();
  if (!elegido) return;
  const celda = sh.getRange('C6');
  let lista = String(celda.getValue()).split(',').map(function (p) { return p.trim(); })
    .filter(function (p) { return p && p !== TP; });
  let aviso;
  if (elegido === TP) {
    lista = [];
    aviso = 'Selección limpiada: se muestran todos los proyectos.';
  } else if (lista.indexOf(elegido) !== -1) {
    lista = lista.filter(function (p) { return p !== elegido; });
    aviso = 'Quitado ' + elegido + ' (' + lista.length + ' elegidos).';
  } else {
    lista.push(elegido);
    aviso = 'Agregado ' + elegido + ' (' + lista.length + ' elegidos).';
  }
  lista.sort();
  celda.setValue(lista.length ? lista.join(', ') : TP);
  picker.setValue('');
  sh.getParent().toast(aviso, 'Decisión', 4);
}

/** Deja en C7 sólo los proyectos del cliente elegido en C5 (todos si es "Todos"). */
function actualizarListaProyectos_(vista) {
  const datos = vista.getParent().getSheetByName(CONFIG_DECISION.HOJA_DATOS);
  const cliente = vista.getRange('C5').getValue();
  const filas = datos.getRange(2, 1, datos.getLastRow() - 1, 2).getValues();
  const proyectos = [];
  filas.forEach(function (f) {
    if (f[1] && (cliente === CONFIG_DECISION.TODOS || f[0] === cliente) && proyectos.indexOf(f[1]) === -1) {
      proyectos.push(f[1]);
    }
  });
  proyectos.sort();
  vista.getRange('C7').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([CONFIG_DECISION.TODOS_PROYECTOS].concat(proyectos), true).setAllowInvalid(false).build());
}

function limpiarFiltros() {
  const vista = SpreadsheetApp.getActive().getSheetByName(CONFIG_DECISION.HOJA_VISTA);
  vista.getRange('C5').setValue(CONFIG_DECISION.TODOS);
  vista.getRange('C6').setValue(CONFIG_DECISION.TODOS_PROYECTOS);
  vista.getRange('C7').setValue('');
  vista.getRange('C8').setValue(CONFIG_DECISION.TODOS);
  vista.getRange('C9').setValue(CONFIG_DECISION.TODOS);
  vista.getRange('C10').setValue(1);
  vista.getRange('C11').setValue(0);
  actualizarListaProyectos_(vista);
}

/* ------------------------------------------------------------------ */
/* FILTRO EN CASCADA: cada desplegable depende de los de arriba        */
/* ------------------------------------------------------------------ */

/** Ejecutar una vez a mano para aplicar la cascada con lo que hoy hay en FILTRO. */
function activarCascadaFiltro() {
  cascadaFiltro_(SpreadsheetApp.getActive().getSheetByName(CONFIG_DECISION.HOJA_FILTRO), 7, false);
}

/**
 * Recalcula los desplegables de FILTRO desde la fila `desde` (7, 8 o 9) hacia abajo.
 * C7 Tipo de servicio: intervalos que el proyecto tiene en Condiciones (y que Auxiliar admite).
 * C8 País: países objetivo del proyecto en ESE intervalo.
 * C9 Idioma: idioma del cliente + idiomas con negociación en BASE para ESE país.
 * Si el valor actual ya no está entre las opciones, pasa a la primera opción.
 * C10 y C11: `reiniciar` = true los vuelve a su valor por defecto; false conserva lo que haya
 * mientras sea válido (es lo que corresponde después del autorrelleno, que ya los completó).
 */
function cascadaFiltro_(filtro, desde, reiniciar) {
  const ss = filtro.getParent();
  const aux = ss.getSheetByName(CONFIG_DECISION.HOJA_AUXILIAR);
  const filas = leerCondicionesFiltro_(ss);
  const proyecto = String(filtro.getRange('C5').getValue()).trim();
  const delProyecto = filas.filter(function (r) { return r.proyecto === proyecto; });
  if (!delProyecto.length) {
    if (proyecto) ss.toast('"' + proyecto + '" no está en Condiciones: los desplegables quedan sin filtrar.', 'FILTRO', 6);
    aplicarOpciones_(filtro.getRange('C7'), [], aux.getRange('B2:B5'));
    aplicarOpciones_(filtro.getRange('C8'), [], aux.getRange('C2:C300'));
    aplicarOpciones_(filtro.getRange('C9'), [], aux.getRange('D2:D300'));
    filtro.getRange('C10').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, CONFIG_DECISION.MAX_LINKS).setAllowInvalid(false).setHelpText('Un número entre 1 y 50').build());
    filtro.getRange('C11').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Opcional: tráfico orgánico mínimo').build());
    return;
  }

  if (desde <= 7) {
    const validos = aux.getRange('B2:B5').getValues().flat().filter(String);
    const porTipo = {};
    delProyecto.forEach(function (r) {
      if (r.tipo && validos.indexOf(r.tipo) !== -1) porTipo[r.tipo] = (porTipo[r.tipo] || 0) + r.cantidad;
    });
    const tipos = Object.keys(porTipo).sort(function (a, b) { return porTipo[b] - porTipo[a] || (a < b ? -1 : 1); });
    aplicarOpciones_(filtro.getRange('C7'), tipos, aux.getRange('B2:B5'));
  }
  if (desde <= 8) {
    const tipo = filtro.getRange('C7').getValue();
    const paises = [];
    delProyecto.filter(function (r) { return r.tipo === tipo; }).forEach(function (r) {
      r.paises.forEach(function (p) { if (paises.indexOf(p) === -1) paises.push(p); });
    });
    aplicarOpciones_(filtro.getRange('C8'), paises, aux.getRange('C2:C300'));
  }
  if (desde <= 9) {
    const pais = filtro.getRange('C8').getValue();
    const cliente = delProyecto.map(function (r) { return r.idioma; }).filter(String)[0];
    const conMedios = [];
    const base = ss.getSheetByName(CONFIG_DECISION.HOJA_BASE);
    const n = base.getLastRow() - 1;
    if (n > 0) {
      base.getRange(2, 13, n, 6).getValues().forEach(function (b) { // M país, N idioma, ..., Q deal_type
        if (b[0] === pais && b[1] && b[4] && conMedios.indexOf(b[1]) === -1) conMedios.push(b[1]);
      });
    }
    // Primero los idiomas con negociación (el del cliente antes, si está); el del cliente sin medios queda al final.
    const idiomas = conMedios.slice().sort(function (a, b) { return (b === cliente) - (a === cliente); });
    if (cliente && idiomas.indexOf(cliente) === -1) idiomas.push(cliente);
    aplicarOpciones_(filtro.getRange('C9'), idiomas, aux.getRange('D2:D300'));
  }
  if (desde <= 10) {
    // Filas de Condiciones que corresponden a proyecto + tipo + país (si ninguna nombra el país, proyecto + tipo).
    const tipo = filtro.getRange('C7').getValue();
    const pais = filtro.getRange('C8').getValue();
    const deTipo = delProyecto.filter(function (r) { return r.tipo === tipo; });
    const dePais = deTipo.filter(function (r) { return r.paises.indexOf(pais) !== -1; });
    const usar = dePais.length ? dePais : deTipo;
    const pendiente = usar.reduce(function (s, r) { return s + r.cantidad; }, 0);
    const trafMin = usar.reduce(function (m, r) { return Math.max(m, r.trafico); }, 0);

    // C10: de 1 hasta lo pendiente (máx. 50). Sin pendiente, queda 1 y se avisa.
    const tope = pendiente > 0 ? Math.min(pendiente, CONFIG_DECISION.MAX_LINKS) : CONFIG_DECISION.MAX_LINKS;
    const c10 = filtro.getRange('C10');
    const v10 = Number(c10.getValue());
    c10.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, tope).setAllowInvalid(false)
      .setHelpText('Entre 1 y ' + tope + (pendiente > 0 ? ' (pendiente en Condiciones: ' + pendiente + ')' : '')).build());
    if (reiniciar || !(v10 >= 1 && v10 <= tope)) c10.setValue(pendiente > 0 ? tope : 1);
    if (pendiente <= 0) ss.toast('Condiciones no tiene cantidad pendiente para ' + tipo + ' / ' + pais + '.', 'FILTRO', 6);

    // C11: no menos que el tráfico mínimo del proyecto.
    const c11 = filtro.getRange('C11');
    const v11 = Number(c11.getValue());
    c11.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(trafMin).setAllowInvalid(false)
      .setHelpText('Mínimo del proyecto en Condiciones: ' + trafMin).build());
    if (reiniciar || (trafMin > 0 && !(v11 >= trafMin))) c11.setValue(trafMin > 0 ? trafMin : '');
  }
}

function aplicarOpciones_(celda, opciones, rangoOriginal) {
  const b = SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  if (!opciones.length) {
    celda.setDataValidation(b.requireValueInRange(rangoOriginal, true).build());
    return;
  }
  celda.setDataValidation(b.requireValueInList(opciones, true).build());
  if (opciones.indexOf(celda.getValue()) === -1) celda.setValue(opciones[0]);
}

function leerCondicionesFiltro_(ss) {
  const sh = ss.getSheetByName(CONFIG_DECISION.HOJA_CONDICIONES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const c = columnasCondiciones_(sh);
  return sh.getRange(2, 1, n, sh.getLastColumn()).getValues().map(function (f) {
    const m = String(f[c.tipo]).match(/(\d+-\d+)\s*->\s*(\d+)/);
    const t = c.trafico < 0 ? 0 : f[c.trafico];
    return {
      proyecto: String(f[c.proyecto]).trim(),
      idioma: String(f[c.idioma]).trim(),
      paises: String(f[c.paises]).split(',').map(function (p) { return p.trim(); }).filter(String),
      tipo: m ? 'Intervalo métrica ' + m[1] : '',
      cantidad: m ? Number(m[2]) : 0,
      trafico: typeof t === 'number' ? t : (Number(String(t).replace(/\./g, '').replace(',', '.')) || 0)
    };
  }).filter(function (r) { return r.proyecto; });
}

/* ------------------------------------------------------------------ */
/* ENVIAR FILA A FILTRO                                                */
/* ------------------------------------------------------------------ */

function enviarAFiltro() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const vista = ss.getActiveSheet();
  if (vista.getName() !== CONFIG_DECISION.HOJA_VISTA) {
    ui.alert('Ubicate en la hoja "' + CONFIG_DECISION.HOJA_VISTA + '" y hacé clic en una fila de resultados.');
    return;
  }
  const fila = vista.getActiveRange().getRow();
  if (fila < CONFIG_DECISION.FILA_RESULTADOS) {
    ui.alert('Hacé clic en una fila de resultados (desde la fila ' + CONFIG_DECISION.FILA_RESULTADOS + ').');
    return;
  }
  const v = vista.getRange(fila, 2, 1, 6).getValues()[0];
  const proyecto = v[1], intervalo = v[2], cantidad = Number(v[3]) || 0, trafico = Number(v[4]) || 0;
  if (!proyecto) {
    ui.alert('La fila seleccionada está vacía.');
    return;
  }

  const filtro = ss.getSheetByName(CONFIG_DECISION.HOJA_FILTRO);
  validarEncabezadosDecision_(filtro, {
    B5: 'Proyecto', B7: 'Tipo de servicio', B8: 'País', B9: 'Idioma', B10: 'Cantidad de links'
  });

  const tipo = 'Intervalo métrica ' + intervalo;
  const tipos = ss.getSheetByName(CONFIG_DECISION.HOJA_AUXILIAR).getRange('B2:B5').getValues().flat();
  if (tipos.indexOf(tipo) === -1) {
    ui.alert('FILTRO no trabaja con "' + tipo + '" (Auxiliar!B2:B5 sólo tiene: ' +
      tipos.filter(String).join(', ') + '). No se envió nada.');
    return;
  }
  const proyectos = ss.getSheetByName(CONFIG_DECISION.HOJA_PARAMETROS).getRange('A2:A60').getValues().flat();
  if (proyectos.indexOf(proyecto) === -1) {
    ui.alert('El proyecto ' + proyecto + ' no está en Parámetros!A2:A60, FILTRO no lo acepta. No se envió nada.');
    return;
  }

  // País: el filtro elegido si coincide; si hay uno solo, ese; si no, se pregunta.
  const lista = String(v[5]).split(',').map(function (p) { return p.trim(); }).filter(String);
  const paisFiltro = vista.getRange('C9').getValue();
  let pais;
  if (paisFiltro !== CONFIG_DECISION.TODOS && lista.indexOf(paisFiltro) !== -1) pais = paisFiltro;
  else if (lista.length === 1) pais = lista[0];
  else {
    const r = ui.prompt('¿Qué país? (' + lista.join(', ') + ')', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    pais = r.getResponseText().trim();
    if (lista.indexOf(pais) === -1) {
      ui.alert('"' + pais + '" no está entre los países del proyecto. No se envió nada.');
      return;
    }
  }

  const datos = ss.getSheetByName(CONFIG_DECISION.HOJA_DATOS);
  const fd = datos.getRange(2, 2, datos.getLastRow() - 1, 6).getValues();
  const idioma = (fd.filter(function (f) { return f[0] === proyecto; })[0] || [])[5] || '';

  const links = Math.min(cantidad, CONFIG_DECISION.MAX_LINKS);
  filtro.getRange('C5').setValue(proyecto);
  filtro.getRange('C7').setValue(tipo);
  filtro.getRange('C8').setValue(pais);
  if (idioma) filtro.getRange('C9').setValue(idioma);
  filtro.getRange('C10').setValue(links);
  filtro.getRange('C11').setValue(trafico > 0 ? trafico : '');
  filtro.getRange('D8').setValue('Cargado desde la hoja Decisión');
  filtro.getRange('D9').setValue('');
  cascadaFiltro_(filtro, 7, false);
  ss.setActiveSheet(filtro);
  ss.toast(proyecto + ' · ' + tipo + ' · ' + pais + ' · ' + links + ' links' +
    (cantidad > links ? ' (pedían ' + cantidad + ', FILTRO admite ' + CONFIG_DECISION.MAX_LINKS + ')' : ''),
    'Enviado a FILTRO', 8);
}

/* ------------------------------------------------------------------ */
/* UTILIDADES                                                          */
/* ------------------------------------------------------------------ */

function obtenerOCrearHojaDecision_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

function validarEncabezadosDecision_(sh, esperados) {
  Object.keys(esperados).forEach(function (a1) {
    const real = String(sh.getRange(a1).getValue()).trim();
    if (real.indexOf(esperados[a1]) !== 0) {
      throw new Error('Encabezado inesperado en "' + sh.getName() + '"!' + a1 +
        ': esperaba "' + esperados[a1] + '" y hay "' + real + '". No se escribió nada.');
    }
  });
}
