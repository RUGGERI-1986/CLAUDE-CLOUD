/**
 * DECISIÓN — vista filtrable de la hoja Condiciones.
 * No modifica Condiciones. Crea 2 hojas:
 *   - "Decisión_Datos" (oculta): fórmulas vivas que separan la columna L y reúnen A, B, I, J, M.
 *   - "Decisión": filtros desplegables + resultados + envío de la fila elegida a FILTRO.
 *
 * Instalación: ver "Cómo probarlo" en el chat. Ejecutar 1 vez setupDecision() y 1 vez setupTriggers().
 */

const CONFIG_DECISION = {
  HOJA_CONDICIONES: 'Condiciones',
  HOJA_VISTA: 'Decisión',
  HOJA_DATOS: 'Decisión_Datos',
  HOJA_FILTRO: 'FILTRO',
  HOJA_AUXILIAR: 'Auxiliar',
  HOJA_PARAMETROS: 'Parámetros',
  HOJA_BASE: 'BASE',
  FILA_ENCABEZADO: 12,
  FILA_RESULTADOS: 13,
  MAX_LINKS: 50,
  TODOS: 'Todos'
};

/* ------------------------------------------------------------------ */
/* SETUP                                                               */
/* ------------------------------------------------------------------ */

function setupDecision() {
  const ss = SpreadsheetApp.getActive();
  const cond = ss.getSheetByName(CONFIG_DECISION.HOJA_CONDICIONES);
  if (!cond) throw new Error('No existe la hoja "' + CONFIG_DECISION.HOJA_CONDICIONES + '".');
  validarEncabezadosDecision_(cond, {
    A1: 'Cliente',
    B1: 'Proyecto',
    I1: 'Idioma Comunicación Cliente',
    J1: 'Países Objetivo',
    L1: 'Tipo de Servicio - Cantidad',
    M1: 'Tráfico Mínimo'
  });

  const datos = obtenerOCrearHojaDecision_(ss, CONFIG_DECISION.HOJA_DATOS);
  const vista = obtenerOCrearHojaDecision_(ss, CONFIG_DECISION.HOJA_VISTA);
  armarDatos_(datos);
  armarVista_(vista);
  datos.hideSheet();
  ss.setActiveSheet(vista);
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

function armarDatos_(sh) {
  sh.clear();
  const C = "'" + CONFIG_DECISION.HOJA_CONDICIONES + "'!";
  const sinProyecto = C + 'B2:B=""';

  sh.getRange('A1:G1').setValues([[
    'Cliente', 'Proyecto', 'Intervalo', 'Cantidad', 'Tráfico mínimo', 'Países', 'Idioma'
  ]]);
  sh.getRange('A2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + C + 'A2:A))');
  sh.getRange('B2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + C + 'B2:B))');
  sh.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(' + sinProyecto + ',"",IFERROR(REGEXEXTRACT(' + C + 'L2:L&"","(\\d+-\\d+)"),"")))');
  sh.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(' + sinProyecto + ',"",IFERROR(VALUE(REGEXEXTRACT(' + C + 'L2:L&"","->\\s*(\\d+)")),0)))');
  sh.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(' + sinProyecto + ',"",IFERROR(VALUE(' + C + 'M2:M),0)))');
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + C + 'J2:J))');
  sh.getRange('G2').setFormula('=ARRAYFORMULA(IF(' + sinProyecto + ',"",' + C + 'I2:I))');

  // Listas para los desplegables
  const todos = '"' + CONFIG_DECISION.TODOS + '"';
  sh.getRange('I1:L1').setValues([['Lista clientes', 'Lista intervalos', 'Lista países', 'Lista proyectos']]);
  sh.getRange('I2').setFormula('={' + todos + ';SORT(UNIQUE(FILTER(A2:A,A2:A<>"")))}');
  sh.getRange('J2').setFormula('={' + todos + ';SORT(UNIQUE(FILTER(C2:C,C2:C<>"")))}');
  sh.getRange('K2').setFormula(
    '=ARRAYFORMULA({' + todos + ';SORT(UNIQUE(FILTER(TRIM(FLATTEN(SPLIT(TEXTJOIN(",",TRUE,F2:F),","))),' +
    'TRIM(FLATTEN(SPLIT(TEXTJOIN(",",TRUE,F2:F),",")))<>"")))})');
  sh.getRange('L2').setFormula('={' + todos + ';SORT(UNIQUE(FILTER(B2:B,B2:B<>"")))}');
  sh.setFrozenRows(1);
}

/* ------------------------------------------------------------------ */
/* HOJA VISIBLE: DECISIÓN                                              */
/* ------------------------------------------------------------------ */

function armarVista_(sh) {
  sh.clear();
  sh.getRange('A1:Z1000').clearDataValidations();
  sh.setHiddenGridlines(true);

  const D = "'" + CONFIG_DECISION.HOJA_DATOS + "'!";
  const datos = sh.getParent().getSheetByName(CONFIG_DECISION.HOJA_DATOS);

  sh.getRange('B1').setValue('Decisión: qué proyectos necesitan medios').setFontSize(16).setFontWeight('bold');
  sh.getRange('B2').setValue('Elegí en las celdas azules. Los resultados se actualizan solos, ordenados por cantidad pendiente.')
    .setFontColor('#666666');

  // Filtros
  sh.getRange('B3').setValue('1 · Filtros').setFontWeight('bold');
  sh.getRange('B4:B9').setValues([
    ['Cliente'], ['Proyecto'], ['Intervalo de métrica'], ['País objetivo'],
    ['Cantidad mínima'], ['Tráfico mínimo exigido (desde)']
  ]);
  sh.getRange('C4:C9').setValues([
    [CONFIG_DECISION.TODOS], [CONFIG_DECISION.TODOS], [CONFIG_DECISION.TODOS], [CONFIG_DECISION.TODOS], [1], [0]
  ]);
  sh.getRange('D4:D9').setValues([
    ['Al elegir cliente, Proyecto muestra sólo los suyos'],
    ['Código del proyecto'],
    ['DR del medio: 0-20, 21-40, ...'],
    ['Coincide si el país está en la lista del proyecto'],
    ['1 oculta los intervalos con cantidad 0'],
    ['Proyectos que exigen al menos este tráfico']
  ]).setFontColor('#666666');
  sh.getRange('C4:C9').setBackground('#D9E8FB').setHorizontalAlignment('left');

  const regla = function (rango) {
    return SpreadsheetApp.newDataValidation().requireValueInRange(rango, true).setAllowInvalid(false).build();
  };
  sh.getRange('C4').setDataValidation(regla(datos.getRange('I2:I')));
  sh.getRange('C5').setDataValidation(regla(datos.getRange('L2:L')));
  sh.getRange('C6').setDataValidation(regla(datos.getRange('J2:J')));
  sh.getRange('C7').setDataValidation(regla(datos.getRange('K2:K')));
  sh.getRange('C8:C9').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());

  // Resumen
  sh.getRange('B10').setFormula(
    '=IF(B' + CONFIG_DECISION.FILA_RESULTADOS + '="","Sin resultados: aflojá algún filtro (Cantidad mínima 0 o Todos).",' +
    'COUNTA(B' + CONFIG_DECISION.FILA_RESULTADOS + ':B)&" combinaciones · "&COUNTUNIQUE(C' + CONFIG_DECISION.FILA_RESULTADOS +
    ':C)&" proyectos · "&SUM(E' + CONFIG_DECISION.FILA_RESULTADOS + ':E)&" links por conseguir")')
    .setFontWeight('bold');
  sh.getRange('B11').setValue('2 · Hacé clic en una fila de resultados y usá el menú Decisión → Enviar fila a FILTRO')
    .setFontColor('#666666');

  // Resultados
  const hdr = sh.getRange(CONFIG_DECISION.FILA_ENCABEZADO, 2, 1, 6);
  hdr.setValues([['Cliente', 'Proyecto', 'Intervalo de métrica', 'Cantidad', 'Tráfico mínimo', 'Países objetivo']])
    .setBackground('#1F3864').setFontColor('#FFFFFF').setFontWeight('bold');
  const cond = function (celda, col) {
    return '(($' + 'C$' + celda + '="' + CONFIG_DECISION.TODOS + '")+(' + D + col + '2:' + col + '=$C$' + celda + ')>0)';
  };
  const pais = '((' + '$C$7="' + CONFIG_DECISION.TODOS + '")+REGEXMATCH(' + D + 'F2:F&"","(^|,)\\s*"&$C$7&"\\s*(,|$)")>0)';
  sh.getRange(CONFIG_DECISION.FILA_RESULTADOS, 2).setFormula(
    '=IFERROR(SORT(FILTER(' + D + 'A2:F,' + D + 'B2:B<>"",' +
    cond(4, 'A') + ',' + cond(5, 'B') + ',' + cond(6, 'C') + ',' + pais + ',' +
    D + 'D2:D>=N($C$8),' + D + 'E2:E>=N($C$9)),4,FALSE,1,TRUE,2,TRUE),"")');
  sh.getRange(CONFIG_DECISION.FILA_RESULTADOS, 5, 500, 2).setNumberFormat('#,##0').setHorizontalAlignment('right');

  sh.setFrozenRows(CONFIG_DECISION.FILA_ENCABEZADO);
  sh.setColumnWidth(1, 20);
  sh.setColumnWidth(2, 190);
  sh.setColumnWidth(3, 190);
  sh.setColumnWidth(4, 190);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 360);
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

/** Al cambiar Cliente, Proyecto sólo ofrece los proyectos de ese cliente. */
function alEditarDecision(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() === CONFIG_DECISION.HOJA_FILTRO) {
    const f = e.range.getRow();
    // C5 Proyecto, C7 Tipo, C8 País alteran los desplegables de abajo. C6 (mes) no depende de nada.
    if (e.range.getColumn() === 3 && (f === 5 || f === 7 || f === 8)) cascadaFiltro_(sh, f === 5 ? 7 : f + 1);
    return;
  }
  if (sh.getName() !== CONFIG_DECISION.HOJA_VISTA) return;
  if (e.range.getRow() !== 4 || e.range.getColumn() !== 3) return;
  actualizarListaProyectos_(sh);
}

/* ------------------------------------------------------------------ */
/* FILTRO EN CASCADA: cada desplegable depende de los de arriba        */
/* ------------------------------------------------------------------ */

/** Ejecutar una vez a mano para aplicar la cascada con lo que hoy hay en FILTRO. */
function activarCascadaFiltro() {
  cascadaFiltro_(SpreadsheetApp.getActive().getSheetByName(CONFIG_DECISION.HOJA_FILTRO), 7);
}

/**
 * Recalcula los desplegables de FILTRO desde la fila `desde` (7, 8 o 9) hacia abajo.
 * C7 Tipo de servicio: intervalos que el proyecto tiene en Condiciones (y que Auxiliar admite).
 * C8 País: países objetivo del proyecto en ESE intervalo.
 * C9 Idioma: idioma del cliente + idiomas con negociación en BASE para ESE país.
 * Si el valor actual ya no está entre las opciones, pasa a la primera opción.
 */
function cascadaFiltro_(filtro, desde) {
  const ss = filtro.getParent();
  const aux = ss.getSheetByName(CONFIG_DECISION.HOJA_AUXILIAR);
  const filas = leerCondicionesFiltro_(ss);
  const proyecto = String(filtro.getRange('C5').getValue()).trim();
  const delProyecto = filas.filter(function (r) { return r.proyecto === proyecto; });
  if (!delProyecto.length) {
    ss.toast('"' + proyecto + '" no está en Condiciones: los desplegables quedan sin filtrar.', 'FILTRO', 6);
    aplicarOpciones_(filtro.getRange('C7'), [], aux.getRange('B2:B5'));
    aplicarOpciones_(filtro.getRange('C8'), [], aux.getRange('C2:C300'));
    aplicarOpciones_(filtro.getRange('C9'), [], aux.getRange('D2:D300'));
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
    const idiomas = [];
    const cliente = delProyecto.map(function (r) { return r.idioma; }).filter(String)[0];
    if (cliente) idiomas.push(cliente);
    const base = ss.getSheetByName(CONFIG_DECISION.HOJA_BASE);
    const n = base.getLastRow() - 1;
    if (n > 0) {
      base.getRange(2, 13, n, 6).getValues().forEach(function (b) { // M país, N idioma, ..., Q deal_type
        if (b[0] === pais && b[1] && b[4] && idiomas.indexOf(b[1]) === -1) idiomas.push(b[1]);
      });
    }
    aplicarOpciones_(filtro.getRange('C9'), idiomas, aux.getRange('D2:D300'));
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
  return sh.getRange(2, 1, n, 13).getValues().map(function (f) {
    const m = String(f[11]).match(/(\d+-\d+)\s*->\s*(\d+)/);
    return {
      proyecto: String(f[1]).trim(),
      idioma: String(f[8]).trim(),
      paises: String(f[9]).split(',').map(function (p) { return p.trim(); }).filter(String),
      tipo: m ? 'Intervalo métrica ' + m[1] : '',
      cantidad: m ? Number(m[2]) : 0
    };
  }).filter(function (r) { return r.proyecto; });
}

function actualizarListaProyectos_(vista) {
  const ss = vista.getParent();
  const datos = ss.getSheetByName(CONFIG_DECISION.HOJA_DATOS);
  const cliente = vista.getRange('C4').getValue();
  const celda = vista.getRange('C5');
  let regla;
  if (cliente === CONFIG_DECISION.TODOS) {
    regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(datos.getRange('L2:L'), true).setAllowInvalid(false).build();
  } else {
    const filas = datos.getRange(2, 1, datos.getLastRow() - 1, 2).getValues();
    const proyectos = [];
    filas.forEach(function (f) {
      if (f[0] === cliente && f[1] && proyectos.indexOf(f[1]) === -1) proyectos.push(f[1]);
    });
    proyectos.sort();
    regla = SpreadsheetApp.newDataValidation()
      .requireValueInList([CONFIG_DECISION.TODOS].concat(proyectos), true).setAllowInvalid(false).build();
  }
  celda.setDataValidation(regla).setValue(CONFIG_DECISION.TODOS);
}

function limpiarFiltros() {
  const vista = SpreadsheetApp.getActive().getSheetByName(CONFIG_DECISION.HOJA_VISTA);
  vista.getRange('C4:C9').setValues([
    [CONFIG_DECISION.TODOS], [CONFIG_DECISION.TODOS], [CONFIG_DECISION.TODOS], [CONFIG_DECISION.TODOS], [1], [0]
  ]);
  actualizarListaProyectos_(vista);
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
  const paisFiltro = vista.getRange('C7').getValue();
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
  cascadaFiltro_(filtro, 7);
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
