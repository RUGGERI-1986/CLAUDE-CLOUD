/**
 * Va en el mismo proyecto de EXTRACTOS BANCARIOS, como archivo aparte.
 *
 * 1) actualizarSaldos(fecha) → hoja "Saldos" de EXTRACTOS.
 *    Deja una fila por cuenta y por día para TODAS las cuentas de Base CP, no solo las automáticas.
 *    La lista de cuentas sale del último bloque de Base CP, así que se mantienen sincronizadas solas.
 *      - Cuentas con API, archivo, CAFCI o QBO: se leen y quedan con fuente propia y estado "al día".
 *      - El resto: se arrastra el último valor, con fuente "arrastre" y estado
 *        "arrastrado desde dd/MM/yyyy", que es la fecha de la última actualización real.
 *    El proceso automático escribe solo las columnas A:G. Las fuentes automáticas se releen si la fila
 *    no quedó "al día" o si se leyó antes de que terminara el día; las filas de arrastre se recalculan.
 *    No acepta la fecha de hoy ni fechas futuras: el día todavía no cerró.
 *
 *    CORRECCIÓN MANUAL: columna H "saldo manual" (y el porqué en I "motivo").
 *      - Si H tiene valor, ese es el saldo de la fila, venga de donde venga: Base CP lo usa, lo pinta
 *        verde, y el arrastre de los días siguientes lo toma.
 *      - El script nunca escribe ni borra H/I (salvo el formulario de cierre, que escribe ahí).
 *      - Al editar H, un trigger anota la fecha y la próxima corrida (6 o 18) rehace saldos y Base CP
 *        desde esa fecha. Para aplicarlo ya: menú Saldos → Aplicar correcciones manuales ahora.
 *      - Borrar H vuelve al valor de la fuente.
 *    Las filas con fuente "manual" de la versión anterior se migran solas una vez (D → H).
 *
 * 2) actualizarBaseCP(fecha) → hoja "Base CP" del Cash Position.
 *    Si no existe el bloque de esa fecha, copia el último debajo (fórmulas y formato) y le pone la
 *    fecha. Si ya existe, lo actualiza en su lugar, aunque sea de una fecha anterior.
 *    Color de la columna I: verde = leído al día o corregido a mano; beige = el resto.
 *
 * 3) recalcular(desde, hasta) → rehace saldos y bloques de un rango, en formato 'yyyy-MM-dd'.
 *    Si "hasta" es hoy o posterior, se corta en ayer.
 * 4) repasarPendientes() → vuelve a leer los últimos días con cuentas automáticas sin "al día"
 *    (QBO del viernes que se carga el lunes, archivo subido tarde).
 * 5) FCIs (FCI.gs), QuickBooks (QBO.gs) y cierre de mes (Cierre.gs) se suman a las fuentes.
 * 6) Colppy (Colppy.gs): respaldo de todas las cuentas en pesos de "Mapeo Colppy". Si la fuente propia
 *    (CSV, tenencias) no tiene el dato al día, el saldo sale de Colppy, con control contra el último
 *    dato propio. Las cuentas sin fuente propia (Consultatio SRL ARS, Plazo fijo) salen de Colppy.
 *
 * Fecha de corte por defecto: ayer. Con la de hoy, Base Cotiz. todavía no tiene cotización.
 * Mercury y Kraken dan el saldo del momento: solo se leen en la corrida de ayer y una vez por día.
 * Corridas automáticas: 6 y 18 (QBO tiene el día anterior recién después de las 15).
 *
 * Propiedades del script (además de WISE_TOKEN, PAYPAL_CLIENT_ID y PAYPAL_SECRET):
 *   MERCURY_TOKEN, MERCURY_CUENTAS (ids separados por coma; ver listarCuentasMercury)
 *   KRAKEN_KEY, KRAKEN_SECRET   (permiso "Query Funds" solamente, sin restricción de IP)
 *   PAYPAL_JIC_CLIENT_ID, PAYPAL_JIC_SECRET   (opcional; si Paypal JIC está mapeada en QBO, manda QBO)
 *   QBO_*   ver QBO.gs
 *   COBERTURA_<hoja>, RECALCULAR_DESDE, RECALCULAR_MARCA, SALDOS_MIGRADO_H   las escribe el script
 */

const CP_ID = '1eusEAEt9PzHXTH8rR66pnkLz7bC8rN87L_a_7xXA_sw';
const HOJA_CP = 'Base CP';
const HOJA_COTIZ = 'Base Cotiz.';
const HOJA_SALDOS = 'Saldos';
const HEADERS_SALDOS = ['fecha', 'cuenta', 'moneda', 'saldo', 'fuente', 'estado', 'actualizado', 'saldo manual', 'motivo'];
const N_AUTO = 7;      // columnas A:G: las únicas que escribe el proceso automático
const COL_MANUAL = 8;  // H
const COL_MOTIVO = 9;  // I
const COL_CP = { fecha: 2, cuenta: 3, moneda: 6, saldo: 9 };
const VERDE = '#00ff00';
const BEIGE = '#fff2cc';
const FUENTE_ARRASTRE = 'arrastre';
const FUENTE_MANUAL = 'manual'; // versión anterior; hoy solo lo usa la migración
const HORAS_CORRIDA = [6, 18];
const HANDLERS_PROPIOS = ['pipelineCompleto', 'pipeline', 'importarTodo', 'registrarEdicionSaldos'];
const DIAS_REPASO = 5;
const MAX_MS_CORRIDA = 4.5 * 60 * 1000; // margen contra el límite de 6 minutos de Apps Script
const _INICIO_EJECUCION = Date.now();

// Cuentas automáticas. fn devuelve: número | {valor, fresco, estado?} | null (no configurada → se arrastra).
// unaVez: el saldo es del momento, así que se lee una sola vez y solo en la corrida de ayer.
const FUENTES_SALDO = {
  'Wise|USD':          { fuente: 'Wise API',     unaVez: false, fn: ctx => saldoWise(ctx, 'USD') },
  'Paypal SRL|USD':    { fuente: 'PayPal API',   unaVez: false, fn: ctx => saldoPayPal(ctx, 'PAYPAL_CLIENT_ID', 'PAYPAL_SECRET') },
  'Paypal JIC|USD':    { fuente: 'PayPal API',   unaVez: false, fn: ctx => saldoPayPal(ctx, 'PAYPAL_JIC_CLIENT_ID', 'PAYPAL_JIC_SECRET') },
  'Mercury|USD':       { fuente: 'Mercury API',  unaVez: true,  fn: ctx => saldoMercury(ctx) },
  'Kraken|USD':        { fuente: 'Kraken API',   unaVez: true,  fn: ctx => saldoKraken(ctx) },
  'Banco Galicia|ARS': { fuente: 'Galicia_raw',  unaVez: false, fn: ctx => saldoDesdeRaw(ctx, 'Galicia_raw', { moneda: 'ARS' }) },
  'Payoneer|USD':      { fuente: 'Payoneer_raw', unaVez: false, fn: ctx => saldoDesdeRaw(ctx, 'Payoneer_raw', { moneda: 'USD', estados: ['Completed'] }) },
  'Banco ICBC|ARS':    { fuente: 'ICBC_ARS_raw', unaVez: false, fn: ctx => saldoDesdeRaw(ctx, 'ICBC_ARS_raw', { moneda: 'ARS' }) },
  'Banco ICBC|USD':    { fuente: 'ICBC_USD_raw', unaVez: false, fn: ctx => saldoDesdeRaw(ctx, 'ICBC_USD_raw', { moneda: 'USD' }) }
};

// Fuentes fijas + una por cada FCI de Tenencias FCI + una por cada cuenta mapeada en QBO.
// Si una cuenta está en más de un lugar, manda el último (QBO). Colppy queda como respaldo de las
// cuentas en pesos de "Mapeo Colppy" (ver conColppy en Colppy.gs).
function todasLasFuentes() {
  return conColppy(Object.assign({}, FUENTES_SALDO, fuentesFCI(), fuentesQBO()));
}

// ---------- AUXILIARES DE LA HOJA SALDOS ----------
function vacio(v) {
  return v === '' || v == null;
}

// Valor vigente de una fila de Saldos: H si tiene algo, si no D.
function valorFila(r) {
  if (!vacio(r[COL_MANUAL - 1])) {
    const v = typeof r[COL_MANUAL - 1] === 'number' ? r[COL_MANUAL - 1] : NaN;
    return { valor: v, manual: true, valido: isFinite(v) };
  }
  if (vacio(r[3])) return { valor: null, manual: false, valido: false };
  const v = Number(r[3]);
  return { valor: v, manual: false, valido: isFinite(v) };
}

// Fecha (dd/MM/yyyy) de la última actualización real detrás de una fila.
function desdeDeFila(r, fechaStr, manual) {
  if (!manual && String(r[4]).trim().toLowerCase() === FUENTE_ARRASTRE) {
    const m = String(r[5]).match(/arrastrado desde (\d{2}\/\d{2}\/\d{4})/);
    if (m) return m[1];
  }
  return fechaStr.split('-').reverse().join('/');
}

function leerSaldos(sh) {
  return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_SALDOS.length).getValues() : [];
}

function tiempoAgotado() {
  return Date.now() - _INICIO_EJECUCION > MAX_MS_CORRIDA;
}

function hojaSaldos(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_SALDOS);
  if (!sh) {
    sh = ss.insertSheet(HOJA_SALDOS);
    sh.getRange(1, 1, 1, HEADERS_SALDOS.length).setValues([HEADERS_SALDOS]);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy');
    sh.getRange('G:G').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.setFrozenRows(1);
    PropertiesService.getScriptProperties().setProperty('SALDOS_MIGRADO_H', new Date().toISOString());
    return sh;
  }
  if (sh.getMaxColumns() < HEADERS_SALDOS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS_SALDOS.length - sh.getMaxColumns());
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS_SALDOS.length).setValues([HEADERS_SALDOS]);
  } else {
    const extra = sh.getRange(1, COL_MANUAL, 1, 2).getValues()[0].map(x => String(x).trim());
    if (extra[0] !== HEADERS_SALDOS[COL_MANUAL - 1] || extra[1] !== HEADERS_SALDOS[COL_MOTIVO - 1]) {
      if (extra[0] || extra[1]) {
        throw new Error(`la hoja Saldos tiene otra cosa en H1:I1 (${extra.join(' | ')}); ` +
                        'movela para agregar "saldo manual" y "motivo"');
      }
      sh.getRange(1, COL_MANUAL, 1, 2).setValues([[HEADERS_SALDOS[COL_MANUAL - 1], HEADERS_SALDOS[COL_MOTIVO - 1]]]);
    }
  }
  migrarManuales(sh);
  return sh;
}

// Una sola vez: las filas con fuente "manual" pasan su saldo a H, que es el mecanismo nuevo.
function migrarManuales(sh) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SALDOS_MIGRADO_H')) return;
  let n = 0;
  leerSaldos(sh).forEach((r, i) => {
    if (String(r[4]).trim().toLowerCase() !== FUENTE_MANUAL || !vacio(r[COL_MANUAL - 1])) return;
    const estado = String(r[5]).trim();
    const motivo = !vacio(r[COL_MOTIVO - 1]) ? r[COL_MOTIVO - 1]
      : (estado.indexOf('cierre:') === 0 ? estado.replace('cierre:', 'cierre de mes:') : 'migrado de fuente manual');
    sh.getRange(i + 2, 5, 1, 2).setValues([[FUENTE_ARRASTRE, 'ver saldo manual']]);
    sh.getRange(i + 2, COL_MANUAL, 1, 2).setValues([[r[3], motivo]]);
    n++;
  });
  props.setProperty('SALDOS_MIGRADO_H', new Date().toISOString());
  if (n) console.log(`Saldos: ${n} fila(s) con fuente manual migradas a la columna "saldo manual"`);
}

// ---------- PASO 1: HOJA "SALDOS" ----------
function actualizarSaldos(fechaCorte) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const ayerStr = fmt(new Date(Date.now() - 864e5));
  const fechaStr = fechaCorte || ayerStr;
  // El día de hoy no cerró: un saldo leído ahora quedaría "al día" con un valor de mitad de día.
  if (fechaStr > ayerStr) throw new Error(`el ${fechaStr} es hoy o futuro: solo se cargan días cerrados`);
  const ini = Utilities.parseDate(fechaStr, tz, 'yyyy-MM-dd');
  const fin = new Date(ini.getTime() + 864e5 - 1000);
  const ctx = { fechaStr, ayerStr, ini, fin, tz };

  // La lista de cuentas la manda Base CP
  const bloque = ultimoBloque(SpreadsheetApp.openById(CP_ID));
  const fuentes = todasLasFuentes();

  const sh = hojaSaldos(ss);
  const ultimo = {};   // último valor vigente ANTES de la fecha de corte
  const delDia = {};   // fila ya existente de la fecha de corte
  leerSaldos(sh).forEach((r, i) => {
    if (!(r[0] instanceof Date)) return;
    const f = fmt(r[0]);
    if (f > fechaStr) return;
    const clave = String(r[1]).trim() + '|' + String(r[2]).trim();
    const v = valorFila(r);
    if (f === fechaStr) {
      delDia[clave] = { fila: i + 2, fuente: String(r[4]).trim(), estado: String(r[5]).trim(), manual: v.manual,
                        actualizado: r[6] instanceof Date ? r[6].getTime() : 0 };
    } else if (v.valido && (!ultimo[clave] || f >= ultimo[clave].fecha)) {
      ultimo[clave] = { fecha: f, saldo: v.valor, desde: desdeDeFila(r, f, v.manual) };
    }
  });

  const errores = [];
  const leidos = [];
  const agregar = [];
  bloque.cuentas.forEach(c => {
    const clave = c.cuenta + '|' + c.moneda;
    const existente = delDia[clave];
    if (existente && existente.manual) return; // corregida a mano (H): no se toca

    const f = fuentes[clave];
    if (f) {
      // Ya leída = misma fuente, al día y leída después del cierre del día. Si quedó "último
      // disponible" (archivo subido tarde, VCP sin publicar, QBO sin actualizar) o se leyó antes de
      // que terminara el día, se vuelve a leer. Lo tomado de Colppy se relee siempre: el contador
      // puede cargar asientos de esa fecha después.
      const yaLeida = existente && existente.fuente === f.fuente && existente.fuente !== FUENTE_COLPPY &&
                      existente.estado.indexOf('al día') === 0 && existente.actualizado > fin.getTime();
      const fueraDeVentana = f.unaVez && fechaStr !== ayerStr; // saldo del momento: no sirve para otra fecha
      if (!yaLeida && !fueraDeVentana) {
        try {
          const r = f.fn(ctx);
          if (r != null) {
            const s = typeof r === 'number' ? { valor: r, fresco: true } : r;
            if (!isFinite(s.valor)) throw new Error('saldo no numérico: ' + s.valor);
            // La fuente puede traer su propio estado (ej: FCI con suscripciones sin tenencia nueva)
            const estado = s.estado || (s.fresco !== false ? 'al día' : 'último disponible');
            const fila = [ini, c.cuenta, c.moneda, Math.round(s.valor * 100) / 100, s.fuente || f.fuente, estado, new Date()];
            if (existente) sh.getRange(existente.fila, 1, 1, N_AUTO).setValues([fila]);
            else agregar.push(fila);
            leidos.push(`${clave}: ${fila[3]}${estado.indexOf('al día') === 0 ? '' : ' (' + estado + ')'}`);
            return;
          }
        } catch (e) {
          errores.push(`${clave}: ${e.message}`);
        }
      } else if (existente && existente.fuente.toLowerCase() !== FUENTE_ARRASTRE) {
        return; // ya resuelta para esta fecha con una lectura de la fuente
      }
      // Mercury/Kraken fuera de su ventana: si la fila es de arrastre, se recalcula abajo con el
      // valor vigente (ej: después de corregir a mano un día anterior).
    }

    // Arrastre: último valor vigente. Solo pisa filas de arrastre, nunca lecturas de una fuente.
    if (existente && existente.fuente.toLowerCase() !== FUENTE_ARRASTRE) return;
    const prev = ultimo[clave];
    // Primer día del historial: no hay valor anterior en Saldos. Si la fila ya existe se deja como
    // está; si no, el último bloque de Base CP la tomaría con el valor de HOY, no el de esa fecha.
    if (!prev && existente) return;
    const saldo = prev ? prev.saldo : c.saldo;
    const desde = prev ? prev.desde : bloque.fechaStr.split('-').reverse().join('/');
    if (vacio(saldo)) return;
    const fila = [ini, c.cuenta, c.moneda, saldo, FUENTE_ARRASTRE, 'arrastrado desde ' + desde, new Date()];
    if (existente) sh.getRange(existente.fila, 1, 1, N_AUTO).setValues([fila]);
    else agregar.push(fila);
  });

  if (agregar.length) {
    sh.getRange(sh.getLastRow() + 1, 1, agregar.length, N_AUTO).setValues(agregar);
  }

  console.log(`Saldos ${fechaStr}: ${leidos.length} leídas, ${agregar.length} filas nuevas. ` +
              (leidos.join(' | ') || 'sin lecturas'));
  if (errores.length) {
    console.error(errores.join('\n'));
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      `Saldos ${fechaStr}: ${errores.length} error(es)`, errores.join('\n'));
  }
}

// Último bloque de Base CP: fecha, fila inicial, cantidad de filas y cuentas con su saldo.
function ultimoBloque(cp) {
  const tz = cp.getSpreadsheetTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const sh = cp.getSheetByName(HOJA_CP);
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, COL_CP.saldo).getValues();
  let iFin = datos.length - 1;
  while (iFin >= 0 && !(datos[iFin][COL_CP.fecha - 1] instanceof Date)) iFin--;
  if (iFin < 0) throw new Error('Base CP no tiene fechas');
  const ultima = fmt(datos[iFin][COL_CP.fecha - 1]);
  let iIni = iFin;
  while (iIni > 0 && datos[iIni - 1][COL_CP.fecha - 1] instanceof Date &&
         fmt(datos[iIni - 1][COL_CP.fecha - 1]) === ultima) iIni--;

  const cuentas = [];
  for (let i = iIni; i <= iFin; i++) {
    cuentas.push({
      cuenta: String(datos[i][COL_CP.cuenta - 1]).trim(),
      moneda: String(datos[i][COL_CP.moneda - 1]).trim(),
      saldo: datos[i][COL_CP.saldo - 1]
    });
  }
  return { fechaStr: ultima, filaIni: iIni + 2, n: iFin - iIni + 1, cuentas };
}

// ---------- PASO 2: BLOQUE EN BASE CP ----------
function actualizarBaseCP(fechaCorte) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60 * 1000)) throw new Error('hay otra corrida en curso');
  try {
    const tzScript = Session.getScriptTimeZone();
    const ayerStr = Utilities.formatDate(new Date(Date.now() - 864e5), tzScript, 'yyyy-MM-dd');
    const fechaStr = fechaCorte || ayerStr;
    if (fechaStr > ayerStr) throw new Error(`el ${fechaStr} es hoy o futuro: solo se pasan días cerrados`);

    // Saldos de esa fecha (H si hay corrección manual, si no D)
    const saldos = {};
    const invalidos = [];
    leerSaldos(hojaSaldos(SpreadsheetApp.getActiveSpreadsheet())).forEach(r => {
      if (!(r[0] instanceof Date)) return;
      if (Utilities.formatDate(r[0], tzScript, 'yyyy-MM-dd') !== fechaStr) return;
      const clave = String(r[1]).trim() + '|' + String(r[2]).trim();
      const v = valorFila(r);
      if (v.manual && !v.valido) {
        invalidos.push(`${clave}: el saldo manual "${r[COL_MANUAL - 1]}" no es un número`);
        return;
      }
      if (!v.valido) return;
      saldos[clave] = { valor: v.valor, fresco: v.manual || String(r[5]).trim().indexOf('al día') === 0 };
    });

    const cp = SpreadsheetApp.openById(CP_ID);
    const tz = cp.getSpreadsheetTimeZone();
    const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const ini = Utilities.parseDate(fechaStr, tz, 'yyyy-MM-dd');

    // Sin cotización de la fecha, Base CP da #N/A en J y rompe los SUMIF de KPI.
    // Además del día se controla el código de la columna C: Base CP busca "<serial><moneda>"
    // (ej: 46265USD). Una fecha con hora (46266,5USD) pasa el control de día pero no matchea.
    const cot = cp.getSheetByName(HOJA_COTIZ);
    const serial = Math.round((ini.getTime() - Date.UTC(1899, 11, 30)) / 864e5);
    const monedas = new Set(cot.getRange(2, 1, cot.getLastRow() - 1, 3).getValues()
      .filter(r => r[0] instanceof Date && fmt(r[0]) === fechaStr &&
                   String(r[2]).trim() === serial + String(r[1]).trim())
      .map(r => String(r[1]).trim()));
    const sinCotiz = ['USD', 'ARS', 'BTC'].filter(m => !monedas.has(m));
    if (sinCotiz.length) {
      throw new Error(`Base Cotiz. no tiene ${sinCotiz.join(', ')} del ${fechaStr} ` +
                      `con código ${serial}<moneda> (revisar fecha con hora en el archivo de cotizaciones)`);
    }

    const sh = cp.getSheetByName(HOJA_CP);
    const fechas = sh.getRange(2, COL_CP.fecha, sh.getLastRow() - 1, 1).getValues().map(r => r[0]);

    // ¿Ya existe el bloque de esa fecha?
    let iIni = -1, iFin = -1;
    fechas.forEach((f, i) => {
      if (f instanceof Date && fmt(f) === fechaStr) { if (iIni < 0) iIni = i; iFin = i; }
    });

    let fila, n;
    const nuevo = iIni < 0;
    if (!nuevo) {
      fila = iIni + 2;
      n = iFin - iIni + 1;
    } else {
      const bloque = ultimoBloque(cp);
      if (bloque.fechaStr > fechaStr) {
        console.log(`no hay bloque del ${fechaStr} y Base CP ya tiene ${bloque.fechaStr}: no se escribe.`);
        return;
      }
      n = bloque.n;
      fila = bloque.filaIni + n;
      const nCols = sh.getLastColumn();
      const faltan = fila + n - 1 - sh.getMaxRows();
      if (faltan > 0) sh.insertRowsAfter(sh.getMaxRows(), faltan);
      const ocupadas = sh.getRange(fila, COL_CP.fecha, n, 2).getValues()
        .some(r => r[0] !== '' || r[1] !== '');
      if (ocupadas) throw new Error(`hay datos debajo del último bloque (fila ${fila}); no se escribe`);
      sh.getRange(bloque.filaIni, 1, n, nCols).copyTo(sh.getRange(fila, 1, n, nCols));
      sh.getRange(fila, COL_CP.fecha, n, 1).setValues(Array.from({ length: n }, () => [ini]));
    }

    // Columna I y colores (celda por celda, para no pisar fórmulas del resto)
    const bloqueDatos = sh.getRange(fila, 1, n, COL_CP.saldo).getValues();
    const colI = sh.getRange(fila, COL_CP.saldo, n, 1);
    const colores = colI.getBackgrounds();
    let escritas = 0;
    bloqueDatos.forEach((r, i) => {
      const clave = String(r[COL_CP.cuenta - 1]).trim() + '|' + String(r[COL_CP.moneda - 1]).trim();
      const s = saldos[clave];
      if (!s) return;
      sh.getRange(fila + i, COL_CP.saldo).setValue(s.valor);
      colores[i][0] = s.fresco ? VERDE : BEIGE;
      escritas++;
      delete saldos[clave];
    });
    colI.setBackgrounds(colores);

    const sobrantes = Object.keys(saldos);
    console.log(`Base CP ${fechaStr} (${nuevo ? 'bloque nuevo' : 'actualización'}): ${escritas} filas escritas`);
    const avisos = [];
    if (sobrantes.length) avisos.push('Cuentas de la hoja Saldos que no están en el bloque de Base CP: ' + sobrantes.join(', '));
    if (invalidos.length) avisos.push('No se escribieron (corregir la columna H de Saldos):\n' + invalidos.join('\n'));
    if (avisos.length) {
      console.error(avisos.join('\n'));
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), `Base CP ${fechaStr}: revisar`, avisos.join('\n\n'));
    }
  } finally {
    lock.releaseLock();
  }
}

// Rehace saldos y bloques de un rango: recalcular('2026-09-30', '2026-10-02'). Corta en ayer.
function recalcular(desdeStr, hastaStr) {
  const tz = Session.getScriptTimeZone();
  const ayer = Utilities.formatDate(new Date(Date.now() - 864e5), tz, 'yyyy-MM-dd');
  if (hastaStr > ayer) hastaStr = ayer;
  const fin = Utilities.parseDate(hastaStr, tz, 'yyyy-MM-dd');
  let d = Utilities.parseDate(desdeStr, tz, 'yyyy-MM-dd');
  while (d.getTime() <= fin.getTime()) {
    const f = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    try {
      actualizarSaldos(f);
      actualizarBaseCP(f);
    } catch (e) {
      console.error(`${f}: ${e.message}`);
    }
    d = Utilities.parseDate(Utilities.formatDate(new Date(d.getTime() + 864e5 + 36e5), tz, 'yyyy-MM-dd'),
                            tz, 'yyyy-MM-dd');
  }
}

// ---------- CORRECCIONES MANUALES (columna H) ----------
// Trigger instalable onEdit: anota desde qué fecha hay que rehacer cuando se edita H.
function registrarEdicionSaldos(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== HOJA_SALDOS) return;
  const c1 = e.range.getColumn();
  const c2 = c1 + e.range.getNumColumns() - 1;
  if (c1 > COL_MANUAL || c2 < COL_MANUAL) return;
  const r1 = Math.max(2, e.range.getRow());
  const n = e.range.getRow() + e.range.getNumRows() - r1;
  if (n <= 0) return;
  const tz = Session.getScriptTimeZone();
  const fechas = sh.getRange(r1, 1, n, 1).getValues()
    .map(r => r[0])
    .filter(d => d instanceof Date)
    .map(d => Utilities.formatDate(d, tz, 'yyyy-MM-dd'))
    .sort();
  if (!fechas.length) return;
  const props = PropertiesService.getScriptProperties();
  const actual = props.getProperty('RECALCULAR_DESDE');
  if (!actual || fechas[0] < actual) props.setProperty('RECALCULAR_DESDE', fechas[0]);
  props.setProperty('RECALCULAR_MARCA', String(Date.now()));
  e.source.toast('Corrección anotada: se aplica a Base CP en la próxima corrida ' +
                 '(o menú Saldos → Aplicar correcciones manuales ahora).', 'Saldos', 6);
}

// Rehace saldos y Base CP desde la fecha más vieja corregida hasta ayer. Si se queda sin tiempo,
// deja anotado desde dónde seguir.
function aplicarCorrecciones() {
  const props = PropertiesService.getScriptProperties();
  const desde = props.getProperty('RECALCULAR_DESDE');
  if (!desde) return 'No hay correcciones pendientes.';
  const marca = props.getProperty('RECALCULAR_MARCA');
  const ayer = sumarDias(hoyStr(), -1);
  const errores = [];
  let f = desde;
  while (f <= ayer) {
    if (tiempoAgotado()) {
      if (props.getProperty('RECALCULAR_MARCA') === marca) props.setProperty('RECALCULAR_DESDE', f);
      return `Correcciones aplicadas hasta el ${fmtCorto(sumarDias(f, -1))}; el resto sigue en la próxima corrida.`;
    }
    try {
      actualizarSaldos(f);
      actualizarBaseCP(f);
    } catch (e) {
      errores.push(`${f}: ${e.message}`);
    }
    f = sumarDias(f, 1);
  }
  // Si mientras corría se corrigió otra fila, la marca cambió: queda anotado para la próxima.
  if (props.getProperty('RECALCULAR_MARCA') === marca) {
    props.deleteProperty('RECALCULAR_DESDE');
    props.deleteProperty('RECALCULAR_MARCA');
  }
  if (errores.length) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      `Correcciones manuales: ${errores.length} fecha(s) con error`, errores.join('\n'));
  }
  return desde > ayer ? 'La corrección es de hoy: entra en la corrida de mañana.'
    : `Correcciones aplicadas del ${fmtCorto(desde)} al ${fmtCorto(ayer)}.` +
      (errores.length ? ` ${errores.length} fecha(s) con error (ver mail).` : '');
}

// ---------- REPASO DE DÍAS PENDIENTES ----------
// Cada corrida mira solo ayer. Esto vuelve sobre los últimos días con cuentas automáticas sin
// "al día" (ej: el viernes de QBO, que se carga el lunes a la tarde) y sobre los días tomados de
// Colppy que hay que releer (ver colppyFechasARepasar en Colppy.gs).
function repasarPendientes() {
  const tz = Session.getScriptTimeZone();
  const hoy = hoyStr();
  const ayer = sumarDias(hoy, -1);
  const desde = sumarDias(hoy, -1 - DIAS_REPASO);
  const desdeColppy = sumarDias(hoy, -1 - COLPPY.DIAS_REPASO);
  const fuentes = todasLasFuentes();
  const fechas = new Set();
  const filasColppy = [];
  leerSaldos(hojaSaldos()).forEach(r => {
    if (!(r[0] instanceof Date) || !vacio(r[COL_MANUAL - 1])) return;
    const f = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    if (f >= ayer) return;
    const clave = String(r[1]).trim() + '|' + String(r[2]).trim();
    if (f >= desdeColppy && String(r[4]).trim() === FUENTE_COLPPY) {
      filasColppy.push({ fecha: f, clave, valor: Number(r[3]) });
    }
    if (f < desde) return;
    const fu = fuentes[clave];
    if (fu && !fu.unaVez && String(r[5]).trim().indexOf('al día') !== 0) fechas.add(f);
  });
  if (filasColppy.length) {
    try {
      colppyFechasARepasar(filasColppy, fuentes, desde).forEach(f => fechas.add(f));
    } catch (e) {
      console.error('repaso Colppy: ' + e.message);
    }
  }
  const lista = Array.from(fechas).sort();
  const hechas = [];
  for (const f of lista) {
    if (tiempoAgotado()) {
      console.warn('repaso cortado por tiempo antes del ' + f);
      break;
    }
    try {
      actualizarSaldos(f);
      actualizarBaseCP(f);
      hechas.push(fmtCorto(f));
    } catch (e) {
      console.error(`repaso ${f}: ${e.message}`);
    }
  }
  return hechas.length ? 'Repasados: ' + hechas.join(', ') : 'No había días pendientes.';
}

// ---------- WISE ----------
function saldoWise(ctx, moneda) {
  const token = PropertiesService.getScriptProperties().getProperty('WISE_TOKEN');
  const get = path => JSON.parse(UrlFetchApp.fetch('https://api.wise.com' + path,
    { headers: { Authorization: 'Bearer ' + token } }).getContentText());

  const perfil = get('/v2/profiles').find(x => x.type === 'BUSINESS');
  if (!perfil) throw new Error('no hay perfil BUSINESS');
  const bal = get(`/v4/profiles/${perfil.id}/balances?types=STANDARD`).find(b => b.currency === moneda);
  if (!bal) throw new Error('no hay balance ' + moneda);

  // Saldo al cierre de la fecha de corte, no el del momento de la corrida
  const st = get(`/v1/profiles/${perfil.id}/balance-statements/${bal.id}/statement.json?currency=${moneda}` +
                 `&intervalStart=${ctx.ini.toISOString()}&intervalEnd=${ctx.fin.toISOString()}&type=COMPACT`);
  if (!st.endOfStatementBalance) throw new Error('el extracto no trae endOfStatementBalance');
  return Number(st.endOfStatementBalance.value);
}

// ---------- PAYPAL ----------
function saldoPayPal(ctx, propId, propSecret) {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty(propId);
  const secret = p.getProperty(propSecret);
  if (!id || !secret) return null; // sin credenciales: se arrastra

  const base = 'https://api-m.paypal.com';
  const token = JSON.parse(UrlFetchApp.fetch(base + '/v1/oauth2/token', {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(id + ':' + secret) },
    payload: { grant_type: 'client_credentials' }
  }).getContentText()).access_token;

  // Saldo a las 23:59:59 de la fecha de corte (PayPal tarda hasta 3 h en reflejarlo)
  const asOf = Utilities.formatDate(ctx.fin, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const r = JSON.parse(UrlFetchApp.fetch(
    `${base}/v1/reporting/balances?currency_code=USD&as_of_time=${encodeURIComponent(asOf)}`,
    { headers: { Authorization: 'Bearer ' + token } }).getContentText());
  const b = (r.balances || []).find(x => x.currency === 'USD');
  if (!b) throw new Error('no hay balance USD');
  return Number(b.total_balance.value);
}

// ---------- MERCURY ----------
function cuentasMercury() {
  const token = PropertiesService.getScriptProperties().getProperty('MERCURY_TOKEN');
  if (!token) throw new Error('falta MERCURY_TOKEN');
  const r = UrlFetchApp.fetch('https://api.mercury.com/api/v1/accounts',
    { headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } });
  return JSON.parse(r.getContentText()).accounts || [];
}

// Correr una vez a mano: lista las cuentas para armar MERCURY_CUENTAS.
function listarCuentasMercury() {
  cuentasMercury().forEach(c =>
    console.log([c.id, c.name, c.nickname || '', c.kind, c.status, c.currentBalance].join(' | ')));
}

function saldoMercury(ctx) {
  const ids = (PropertiesService.getScriptProperties().getProperty('MERCURY_CUENTAS') || '')
    .split(',').map(s => s.trim()).filter(Boolean);
  if (!ids.length) return null;
  const cuentas = cuentasMercury();
  return ids.reduce((total, id) => {
    const c = cuentas.find(x => x.id === id);
    if (!c) throw new Error('no encuentro la cuenta ' + id);
    return total + Number(c.currentBalance);
  }, 0);
}

// ---------- KRAKEN ----------
function saldoKraken(ctx) {
  const p = PropertiesService.getScriptProperties();
  const key = p.getProperty('KRAKEN_KEY');
  const secret = p.getProperty('KRAKEN_SECRET');
  if (!key || !secret) return null;

  const path = '/0/private/Balance';
  const nonce = String(Date.now() * 1000);
  const body = 'nonce=' + nonce;
  const sha = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, nonce + body, Utilities.Charset.UTF_8);
  const firma = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_512,
    Utilities.newBlob(path).getBytes().concat(sha), Utilities.base64Decode(secret));

  const r = JSON.parse(UrlFetchApp.fetch('https://api.kraken.com' + path, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: body,
    headers: { 'API-Key': key, 'API-Sign': Utilities.base64Encode(firma) }
  }).getContentText());
  if (r.error && r.error.length) throw new Error(r.error.join(', '));

  // USDT spot y sus variantes de Earn/staking (USDT.F, USDT.B, …), valuadas 1:1 como la fila USDT
  return Object.keys(r.result || {})
    .filter(k => k.split('.')[0] === 'USDT')
    .reduce((total, k) => total + Number(r.result[k]), 0);
}

// ---------- DESDE HOJAS *_raw (Galicia, Payoneer, ICBC) ----------
// opts.moneda: solo filas de esa moneda. opts.estados: solo filas con esos estados.
function saldoDesdeRaw(ctx, hoja, opts) {
  opts = opts || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(hoja);
  if (!sh || sh.getLastRow() < 2) return null;
  const tzRaw = ss.getSpreadsheetTimeZone();
  const fmt = d => Utilities.formatDate(d, tzRaw, 'yyyy-MM-dd');

  // columnas raw: id, fecha, descripcion, monto, fee, moneda, tipo, estado, saldo
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
    .filter(r => r[1] instanceof Date && r[8] !== '' && fmt(r[1]) <= ctx.fechaStr)
    .filter(r => !opts.moneda || String(r[5]).trim() === opts.moneda)
    .filter(r => !opts.estados || opts.estados.indexOf(String(r[7]).trim()) >= 0)
    .sort((a, b) => a[1] - b[1]);
  if (!filas.length) return null;

  const dia = fmt(filas[filas.length - 1][1]);
  const delDia = filas.filter(r => fmt(r[1]) === dia);

  // Varios movimientos el mismo día (Galicia no trae hora): el cierre es el saldo que no es
  // "saldo previo" (saldo − monto) de ningún otro movimiento de ese día.
  const c = x => Math.round(Number(x) * 100);
  const previos = new Set(delDia.map(r => c(r[8]) - c(r[3])));
  const cierres = delDia.filter(r => !previos.has(c(r[8])));
  const cierre = cierres.length === 1 ? cierres[0] : delDia[delDia.length - 1];

  // "Al día" si se procesó un archivo subido después del cierre de la fecha de corte,
  // así un fin de semana sin movimientos igual queda en verde.
  const cobertura = Number(PropertiesService.getScriptProperties().getProperty('COBERTURA_' + hoja) || 0);
  return { valor: Number(cierre[8]), fresco: cobertura > ctx.fin.getTime() && cierres.length === 1 };
}

// ---------- ORQUESTACIÓN ----------
function pipelineCompleto() {
  [pipeline, actualizarSaldos, actualizarBaseCP, aplicarCorrecciones, repasarPendientes,
   recordatorioCierre, verificarCierre].forEach(fn => {
    try {
      fn();
    } catch (e) {
      console.error(`${fn.name}: ${e.message}`);
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), `${fn.name} falló`, e.message);
    }
  });
}

// Corridas a las 6 y a las 18 + el trigger que anota las correcciones de la columna H.
// Solo borra los triggers de este proceso; los de otros scripts del proyecto quedan.
function crearTriggersCompleto() {
  ScriptApp.getProjectTriggers()
    .filter(t => HANDLERS_PROPIOS.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  HORAS_CORRIDA.forEach(h => ScriptApp.newTrigger('pipelineCompleto').timeBased().everyDays(1).atHour(h).create());
  ScriptApp.newTrigger('registrarEdicionSaldos').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
}
