/**
 * Autorrelleno.gs — archivo NUEVO de la planilla "Asignación de medios" (no reemplaza CargarBulk.gs)
 * Versión 2026-10-29 d (llama a cascadaFiltro_ de Decision.gs para restringir los desplegables; reacciona también a C8)
 * Qué hace: al elegir el proyecto en FILTRO!C5 completa solo Tipo de servicio, País, Idioma,
 * Cantidad y Tráfico mínimo con los datos de la pestaña Condiciones. Todo queda editable a mano.
 * Si después cambiás el Tipo de servicio (C7), recalcula País, Idioma, Cantidad y Tráfico para ese intervalo.
 * Reglas:
 *  - Intervalo: el primero de Condiciones con cantidad > 0 (si no hay, el primero del proyecto).
 *  - País: entre los países objetivo del intervalo, el que tiene más medios (negociación vigente en el mes,
 *    con cupo, métrica dentro del intervalo y tráfico >= mínimo). La aclaración se escribe en FILTRO!D8.
 *  - Idioma: "Idioma Comunicación Cliente" de Condiciones. Si ese idioma no está entre los idiomas con
 *    negociación, no se toca C9 y se avisa en FILTRO!D9.
 *  - Cantidad: la de Condiciones para el intervalo y país (máximo 50). Si es 0, no se toca C10.
 * Es un activador simple (onEdit): no hay que instalar nada ni dar permisos nuevos.
 * Importante: no puede existir otra función onEdit en el proyecto (hoy no hay ninguna).
 * Depende de Decision.gs (cascadaFiltro_): después de autorrellenar, deja los desplegables de C7:C11
 * con las opciones que corresponden a lo elegido arriba.
 */

var AUTO = {
  FILTRO: 'FILTRO',
  PROYECTO: 5, INTERVALO: 7, PAIS: 8, IDIOMA: 9, CANTIDAD: 10, TRAFICO: 11,   // filas de la columna C
  COND_SHEET: 'Condiciones',
  COND_HEADERS: { B: 'Proyecto', I: 'Idioma Comunicación Cliente', J: 'Países Objetivo',
                  F: 'Cantidad', L: 'Tipo de Servicio - Cantidad', M: 'Tráfico Mínimo' },
  PARAM_SHEET: 'Parámetros',     // A = proyecto, C = métrica (DR o DA)
  AUX_SHEET: 'Auxiliar',         // D2:D = idiomas con negociación
  BASE_SHEET: 'BASE'             // G domain_rating, I authority_score, M país, F tráfico, AI negociación, AJ cupo
};

function onEdit(e) {
  try {
    var r = e && e.range;
    if (!r || r.getSheet().getName() !== AUTO.FILTRO || r.getColumn() !== 3 || r.getNumRows() !== 1) return;
    var fila = r.getRow();
    if (fila === AUTO.PAIS) {
      // País cambiado a mano: Idioma, Cantidad y Tráfico se ajustan a ese país (Decision.gs)
      cascadaFiltro_(r.getSheet(), 9, true);
      return;
    }
    if (fila !== AUTO.PROYECTO && fila !== AUTO.INTERVALO) return;
    autorrellenarFiltro_(e.source, fila === AUTO.PROYECTO);
    // Desplegables según lo que quedó elegido (Decision.gs)
    cascadaFiltro_(r.getSheet(), fila === AUTO.PROYECTO ? 7 : 8, false);
  } catch (err) {
    console.error(new Date(), 'onEdit', err);
    if (e && e.source) e.source.toast('Autorrelleno falló: ' + err.message, 'Asignación', 8);
  }
}

// Función para probar a mano desde el editor (usa el proyecto que esté elegido en FILTRO!C5)
function probarAutorrelleno() {
  autorrellenarFiltro_(SpreadsheetApp.getActive(), true);
}

// Las columnas de flags de BASE tienen formato fecha: getValues() devuelve Date (1 = 31/12/1899).
// Esto lo vuelve al número que la celda realmente tiene.
function num_(v) {
  if (v instanceof Date) return Math.round((v.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
  if (v === '' || v === null) return '';
  var n = Number(v);
  return isNaN(n) ? '' : n;
}

function autorrellenarFiltro_(ss, cambioProyecto) {
  var filtro = ss.getSheetByName(AUTO.FILTRO);
  var proyecto = String(filtro.getRange(AUTO.PROYECTO, 3).getValue()).trim();
  if (!proyecto) return;

  // 1) Condiciones del proyecto
  var cond = ss.getSheetByName(AUTO.COND_SHEET);
  var datos = cond.getRange(1, 2, cond.getLastRow(), 13).getValues();   // B:N
  var enc = datos[0];
  var esperados = { 0: AUTO.COND_HEADERS.B, 7: AUTO.COND_HEADERS.I, 8: AUTO.COND_HEADERS.J,
                    4: AUTO.COND_HEADERS.F, 10: AUTO.COND_HEADERS.L, 11: AUTO.COND_HEADERS.M };
  for (var k in esperados) {
    if (String(enc[k]).trim() !== esperados[k]) {
      throw new Error('Cambiaron los encabezados de Condiciones (esperaba "' + esperados[k] + '"). No completé nada.');
    }
  }
  var filas = [];
  datos.slice(1).forEach(function (d) {
    if (String(d[0]).trim() !== proyecto) return;
    var m = String(d[10]).match(/^(.*?)\s*->\s*(\d+)\s*$/);
    if (!m) return;
    filas.push({
      intervalo: m[1].trim(), cantidad: Number(m[2]) || 0,
      paises: String(d[8]).split(',').map(function (p) { return p.trim(); }).filter(String),
      idioma: String(d[7]).trim(), trafico: d[11]
    });
  });
  if (!filas.length) {
    filtro.getRange(AUTO.PAIS, 4).setValue('Autorrelleno: ' + proyecto + ' no está en Condiciones. Completá a mano.');
    return;
  }

  // 2) Intervalo
  var intervalo;
  if (cambioProyecto) {
    var conCant = filas.filter(function (f) { return f.cantidad > 0; });
    intervalo = (conCant.length ? conCant[0] : filas[0]).intervalo;
  } else {
    intervalo = String(filtro.getRange(AUTO.INTERVALO, 3).getValue()).trim();
  }
  var delIntervalo = filas.filter(function (f) { return f.intervalo === intervalo; });
  if (!delIntervalo.length) {
    filtro.getRange(AUTO.PAIS, 4).setValue('Autorrelleno: ' + proyecto + ' no tiene "' + intervalo + '" en Condiciones.');
    return;
  }
  var paises = [];
  delIntervalo.forEach(function (f) { f.paises.forEach(function (p) { if (paises.indexOf(p) < 0) paises.push(p); }); });

  // 3) Medios por país (negociación vigente + cupo + métrica en el intervalo + tráfico mínimo)
  var traficoMin = delIntervalo[0].trafico === '' ? 0 : Number(delIntervalo[0].trafico) || 0;
  var rango = intervalo.match(/(\d+)-(\d+)/);
  var min = rango ? Number(rango[1]) : 0, max = rango ? Number(rango[2]) : 100;
  var prm = ss.getSheetByName(AUTO.PARAM_SHEET);
  var metrica = 'DR';
  prm.getRange(2, 1, Math.max(prm.getLastRow() - 1, 1), 3).getValues().forEach(function (p) {
    if (String(p[0]).trim() === proyecto && String(p[2]).trim() === 'DA') metrica = 'DA';
  });
  var base = ss.getSheetByName(AUTO.BASE_SHEET);
  var n = base.getLastRow() - 1;
  var traf = base.getRange(2, 6, n, 1).getValues();
  var dr = base.getRange(2, 7, n, 1).getValues();
  var da = base.getRange(2, 9, n, 1).getValues();
  var pais = base.getRange(2, 13, n, 1).getValues();
  var flags = base.getRange(2, 35, n, 2).getValues();   // AI f_negociacion, AJ f_cupo
  var cuenta = {};
  paises.forEach(function (p) { cuenta[p] = 0; });
  for (var i = 0; i < n; i++) {
    var p = String(pais[i][0]).trim();
    if (!(p in cuenta) || num_(flags[i][0]) !== 1 || num_(flags[i][1]) !== 1) continue;
    var v = num_(metrica === 'DA' ? da[i][0] : dr[i][0]);
    if (v === '' || v < min || v > max) continue;
    if (traficoMin && !(num_(traf[i][0]) >= traficoMin)) continue;
    cuenta[p]++;
  }
  var elegido = paises[0], mejor = -1;
  paises.forEach(function (p) { if (cuenta[p] > mejor) { mejor = cuenta[p]; elegido = p; } });
  var otros = paises.filter(function (p) { return p !== elegido; })
    .map(function (p) { return p + ' (' + cuenta[p] + ')'; }).join(', ');
  var nota = mejor > 0
    ? 'Autorrelleno: ' + elegido + ' es el país con más medios (' + mejor + ')' + (otros ? '. Otros: ' + otros : '')
    : 'Autorrelleno: ningún país objetivo tiene medios; dejé ' + elegido + ' (primero de Condiciones)';

  // 4) Cantidad, idioma, tráfico de la fila de Condiciones del país elegido
  var fila = delIntervalo.filter(function (f) { return f.paises.indexOf(elegido) >= 0; })[0] || delIntervalo[0];
  var aux = ss.getSheetByName(AUTO.AUX_SHEET);
  var idiomas = aux.getRange(2, 4, Math.max(aux.getLastRow() - 1, 1), 1).getValues()
    .map(function (r) { return String(r[0]).trim(); });
  var notaIdioma = '';

  // 5) Escribir (solo celdas azules + notas de la columna D)
  if (cambioProyecto) filtro.getRange(AUTO.INTERVALO, 3).setValue(intervalo);
  filtro.getRange(AUTO.PAIS, 3).setValue(elegido);
  if (idiomas.indexOf(fila.idioma) >= 0) {
    filtro.getRange(AUTO.IDIOMA, 3).setValue(fila.idioma);
    notaIdioma = 'Autorrelleno: idioma del cliente en Condiciones';
  } else {
    notaIdioma = 'Autorrelleno: "' + fila.idioma + '" no tiene negociación; elegí el idioma a mano';
  }
  if (fila.cantidad > 0) filtro.getRange(AUTO.CANTIDAD, 3).setValue(Math.min(fila.cantidad, 50));
  filtro.getRange(AUTO.TRAFICO, 3).setValue(fila.trafico === '' ? '' : fila.trafico);
  filtro.getRange(AUTO.PAIS, 4).setValue(nota);
  filtro.getRange(AUTO.IDIOMA, 4).setValue(notaIdioma);
  SpreadsheetApp.flush();   // fuerza el recálculo de C12 y del diagnóstico con los valores recién escritos
}
