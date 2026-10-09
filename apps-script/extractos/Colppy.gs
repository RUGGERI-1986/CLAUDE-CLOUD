/**
 * Colppy (contabilidad de la SRL) → EXTRACTOS BANCARIOS. Archivo nuevo en el proyecto de EXTRACTOS.
 *
 * Colppy es el RESPALDO DIARIO de todas las cuentas en pesos de "Mapeo Colppy" (conColppy, que usa
 * todasLasFuentes de saldosbasecp.gs):
 *   - Si la fuente propia de la cuenta (CSV del banco, Tenencias FCI) tiene el dato al día, manda ella.
 *   - Si no (no se subió el CSV, falta la tenencia o el VCP), el saldo sale de Colppy, con fuente
 *     "Colppy" en la hoja Saldos.
 *   - Las cuentas sin fuente propia (Consultatio SRL ARS, Plazo fijo) salen siempre de Colppy.
 *   - Solo cuentas en ARS: Colppy tiene las cuentas en dólares pasadas a pesos. Las filas USD del
 *     mapeo se usan solo en "Probar Colppy".
 * Control (sin mail): el valor de Colppy se compara con el último dato de la fuente propia.
 *   Si difiere más del umbral (5% FCI, 1% el resto) y más de 50.000 ARS, la fila queda en beige con
 *   el detalle en la columna estado. Sin dato propio para comparar: "al día (Colppy, sin control)".
 *   Cada comparación queda en la hoja "Control Colppy", una fila por cuenta y día.
 * Los días tomados de Colppy se releen solos (repasarPendientes) si Colppy cambió después o si, en
 * los últimos días, apareció el dato propio (CSV subido tarde).
 *
 * Solo lectura: el script únicamente puede llamar a las operaciones de COLPPY_OPERACIONES_PERMITIDAS.
 * Cualquier otra (alta_asiento, borrar_asiento, etc.) se rechaza antes de llegar a Colppy.
 *
 * Saldo a una fecha: suma de todos los movimientos del libro diario de la cuenta desde el inicio
 * hasta esa fecha (Debe − Haber). Verificado contra producción (09/10/2026):
 *   - leer_saldoCuenta responde "éxito" pero devuelve 0 para cualquier cuenta: NO se usa.
 *   - listar_movimientosdiario acepta fromDate / toDate en formato aaaa-mm-dd y devuelve la lista en
 *     "movimientos" con idPlanCuenta, fechaContable, DebitoCredito (D/C) e Importe.
 *   - Si el formato de fecha es otro, IGNORA el filtro sin dar error y devuelve todo: por eso se
 *     controla que cada movimiento esté dentro del rango pedido.
 *   - Sin paginar devuelve solo 50 movimientos: se pide por páginas (start / limit).
 *   - Los movimientos con isNIIF = 1 (libro NIIF) se excluyen para no contarlos dos veces.
 * Control: el saldo calculado de cada banco se compara con el saldo de Tesorería/listado_banco.
 *
 * Hojas:
 *   "Mapeo Colppy"   cuenta | moneda | código Colppy | descripción Colppy | notas
 *                    Varias filas con la misma cuenta y moneda se suman (FCI Galicia = 4 cuentas FIMA).
 *                    "descripción Colppy" la completa "Listar cuentas y bancos".
 *   "Cuentas Colppy" volcado del plan de cuentas imputables (se reescribe en cada listado).
 *   "Bancos Colppy"  volcado de Tesorería/listado_banco (se reescribe en cada listado).
 *   "Control Colppy" la escribe el script: Colppy contra la fuente propia, por cuenta y día.
 *   "Balance Colppy" balance de sumas y saldos al último día de cada mes descargado (todas las cuentas
 *                    con movimientos). Una tanda de filas por mes: al volver a descargar un mes se
 *                    reemplazan solo las filas de ese mes. No se edita a mano: se trabaja en otra hoja
 *                    que lea de esta (SUMIFS por mes y código).
 *
 * Propiedades del script (Configuración del proyecto → Propiedades del script):
 *   COLPPY_INTEGRADOR_USUARIO    mail del usuario integrador (api.colppy.com/registro/)
 *   COLPPY_INTEGRADOR_PASSWORD   contraseña del usuario integrador (texto plano: el script la pasa a MD5)
 *   COLPPY_USUARIO               mail del usuario de la plataforma Colppy (el que crea la SRL, solo lectura)
 *   COLPPY_PASSWORD              contraseña de ese usuario (texto plano: el script la pasa a MD5)
 *   COLPPY_ID_EMPRESA            opcional: id de empresa, si el login devuelve otra
 *   COLPPY_ENTORNO               opcional: "staging" para el entorno de prueba
 *   COLPPY_SESION                la escribe el script (clave de sesión, dura 12 h)
 */

const COLPPY = {
  URL_PROD: 'https://login.colppy.com/lib/frontera2/service.php',
  URL_STAGING: 'https://staging.colppy.com/lib/frontera2/service.php',
  HOJA_MAPEO: 'Mapeo Colppy',
  HOJA_CUENTAS: 'Cuentas Colppy',
  HOJA_BANCOS: 'Bancos Colppy',
  DIARIO_DESDE: '2001-01-01',    // inicio del historial: el saldo es la suma desde acá
  DIAS_FUTURO: 366,              // se piden también movimientos con fecha futura, para informarlos
  LOTE: 500,                     // movimientos por página del libro diario
  MAX_PAGINAS: 200,              // tope de páginas por corrida (100.000 movimientos)
  SESION_MS: 11 * 3600 * 1000,   // la sesión dura 12 h desde el último uso; se renueva antes
  HOJA_CONTROL: 'Control Colppy',
  MONEDA: 'ARS',                 // única moneda que Colppy tiene en su moneda original
  UMBRAL_FCI: 0.05,              // cuentas que empiezan con "FCI"
  UMBRAL_OTRAS: 0.01,            // bancos y el resto
  PISO_ARS: 50000,               // una diferencia menor no se marca aunque supere el porcentaje
  DIAS_REPASO: 40,               // días hacia atrás en que se relee lo tomado de Colppy si cambió
  HOJA_BALANCE: 'Balance Colppy'
};
const FUENTE_COLPPY = 'Colppy';
const HEADERS_CONTROL_COLPPY = ['fecha', 'cuenta', 'moneda', 'Colppy', 'fuente propia', 'valor fuente propia',
                                'diferencia', 'diferencia %', 'resultado', 'va a Saldos', 'actualizado'];
const HEADERS_BALANCE_COLPPY = ['mes', 'código', 'cuenta', 'tipo de cuenta', 'saldo inicial',
                                'debe del mes', 'haber del mes', 'saldo al cierre', 'actualizado'];
const HEADERS_MAPEO_COLPPY = ['cuenta', 'moneda', 'código Colppy', 'descripción Colppy', 'notas'];
const MAPEO_COLPPY_INICIAL = [
  ['Banco Galicia', 'ARS', '111100', 'respaldo si no hay CSV'],
  ['Banco ICBC', 'ARS', '111101', 'respaldo si no hay CSV'],
  ['FCI Galicia', 'ARS', '113001', 'Fima Premium A'],
  ['FCI Galicia', 'ARS', '113002', 'Fima Ahorro Pesos'],
  ['FCI Galicia', 'ARS', '113003', 'Fima Ahorro Plus A'],
  ['FCI Galicia', 'ARS', '113005', 'Fima Ahorro Pesos (otra clase)'],
  ['FCI ICBC', 'ARS', '113004', 'Alpha Pesos'],
  ['Consultatio SRL', 'ARS', '113006', 'One 678 ARS'],
  ['Plazo fijo Galicia', 'ARS', '113000', ''],
  ['Banco ICBC', 'USD', '111105', 'solo diagnóstico: Colppy la tiene en pesos'],
  ['Consultatio SRL', 'USD', '113007', 'solo diagnóstico: Colppy la tiene en pesos']
];
const HEADERS_CUENTAS_COLPPY = ['Id', 'código', 'descripción', 'tipo de cuenta'];
const HEADERS_BANCOS_COLPPY = ['idBanco', 'nombre', 'banco', 'nro. cuenta', 'moneda', 'cuenta contable',
                               'código', 'saldo', 'última conciliación', 'multimoneda'];
// Solo lectura. Usuario/iniciar_sesion se usa únicamente desde colppySesion().
const COLPPY_OPERACIONES_PERMITIDAS = {
  Contabilidad: ['listar_cuentasdiario', 'listar_movimientosdiario'],
  Tesoreria: ['listado_banco']
};
const _colppyCache = {};

// ---------- CONEXIÓN ----------
function colppyMd5(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s), Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function colppyCredenciales() {
  const p = PropertiesService.getScriptProperties();
  const c = {
    intUsuario: p.getProperty('COLPPY_INTEGRADOR_USUARIO'),
    intPassword: p.getProperty('COLPPY_INTEGRADOR_PASSWORD'),
    usuario: p.getProperty('COLPPY_USUARIO'),
    password: p.getProperty('COLPPY_PASSWORD')
  };
  const nombres = { intUsuario: 'COLPPY_INTEGRADOR_USUARIO', intPassword: 'COLPPY_INTEGRADOR_PASSWORD',
                    usuario: 'COLPPY_USUARIO', password: 'COLPPY_PASSWORD' };
  const faltan = Object.keys(nombres).filter(k => !c[k]).map(k => nombres[k]);
  if (faltan.length) throw new Error('faltan en Propiedades del script: ' + faltan.join(', '));
  return c;
}

function colppyUrl() {
  return String(PropertiesService.getScriptProperties().getProperty('COLPPY_ENTORNO') || '').toLowerCase() === 'staging'
    ? COLPPY.URL_STAGING : COLPPY.URL_PROD;
}

function colppyPost(cuerpo) {
  const r = UrlFetchApp.fetch(colppyUrl(), {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(cuerpo),
    muteHttpExceptions: true
  });
  const txt = r.getContentText();
  if (r.getResponseCode() !== 200) throw new Error(`Colppy respondió HTTP ${r.getResponseCode()}: ${txt.slice(0, 300)}`);
  try {
    return JSON.parse(txt);
  } catch (e) {
    throw new Error('Colppy no devolvió JSON: ' + txt.slice(0, 300));
  }
}

// La documentación muestra dos formatos: {result, response: {...}} y {success, message, data}.
function colppyRespuesta(j, operacion) {
  const res = j.result || {};
  const r = j.response !== undefined ? j.response : j;
  if (res.estado !== undefined && Number(res.estado) !== 0) {
    throw new Error(`${operacion}: ${res.mensaje || 'estado ' + res.estado}`);
  }
  if (r && r.success === false) throw new Error(`${operacion}: ${r.message || 'success = false'}`);
  return r;
}

// Clave de sesión guardada en COLPPY_SESION; se pide una nueva si venció o si se fuerza.
function colppySesion(forzar) {
  const props = PropertiesService.getScriptProperties();
  if (!forzar) {
    const raw = props.getProperty('COLPPY_SESION');
    if (raw) {
      const s = JSON.parse(raw);
      if (Date.now() - s.usado < COLPPY.SESION_MS) return s;
    }
  }
  const c = colppyCredenciales();
  let r;
  try {
    r = colppyRespuesta(colppyPost({
      auth: { usuario: c.intUsuario, password: colppyMd5(c.intPassword) },
      service: { provision: 'Usuario', operacion: 'iniciar_sesion' },
      parameters: { usuario: c.usuario, password: colppyMd5(c.password) }
    }), 'iniciar_sesion');
  } catch (e) {
    throw new Error(e.message + '. Si dice "contraseña incorrecta" (602): revisá que el usuario integrador ' +
                    'esté en COLPPY_INTEGRADOR_* y el de la plataforma en COLPPY_USUARIO / COLPPY_PASSWORD, sin invertirlos.');
  }
  const d = r.data || {};
  if (!d.claveSesion) throw new Error('iniciar_sesion no devolvió claveSesion');
  const s = {
    clave: d.claveSesion,
    usuario: c.usuario,
    idEmpresa: String(props.getProperty('COLPPY_ID_EMPRESA') || d.idEmpresa || ''),
    usado: Date.now()
  };
  if (!s.idEmpresa) throw new Error('el login no devolvió idEmpresa: cargalo en la propiedad COLPPY_ID_EMPRESA');
  props.setProperty('COLPPY_SESION', JSON.stringify(s));
  return s;
}

// Llama a una operación de lectura. Si falla con la sesión guardada, vuelve a iniciar sesión una vez.
function colppyLlamar(provision, operacion, params) {
  if ((COLPPY_OPERACIONES_PERMITIDAS[provision] || []).indexOf(operacion) < 0) {
    throw new Error(`operación no permitida: ${provision}/${operacion} (este script solo lee)`);
  }
  const c = colppyCredenciales();
  const pedir = s => colppyRespuesta(colppyPost({
    auth: { usuario: c.intUsuario, password: colppyMd5(c.intPassword) },
    service: { provision, operacion },
    parameters: Object.assign({ sesion: { usuario: s.usuario, claveSesion: s.clave }, idEmpresa: s.idEmpresa }, params || {})
  }), operacion);
  let s = colppySesion(false);
  let r;
  try {
    r = pedir(s);
  } catch (e) {
    console.warn(`${operacion} falló con la sesión guardada (${e.message}); inicio sesión de nuevo`);
    s = colppySesion(true);
    r = pedir(s);
  }
  s.usado = Date.now();
  PropertiesService.getScriptProperties().setProperty('COLPPY_SESION', JSON.stringify(s));
  return r;
}

// ---------- DATOS ----------
function colppyNum(v) {
  if (typeof v === 'number') return v;
  let s = String(v == null ? '' : v).trim();
  if (s === '') return 0;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,56
  else if (s.indexOf(',') >= 0) s = s.replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');                      // 1.234.567
  const n = Number(s);
  if (!isFinite(n)) throw new Error('importe no numérico en la respuesta de Colppy: ' + v);
  return n;
}

// "20-01-2026", "20/01/2026" o "2026-01-20" → "2026-01-20"
function colppyFecha(v) {
  const s = String(v == null ? '' : v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return null;
}

function colppyRedondear(n) {
  return Math.round(n * 100) / 100;
}

function colppyCuentas() {
  if (_colppyCache.cuentas) return _colppyCache.cuentas;
  const r = colppyLlamar('Contabilidad', 'listar_cuentasdiario', { query: '' });
  _colppyCache.cuentas = r.cuentas || r.data || []; // el listado viene en "cuentas", no en "data"
  return _colppyCache.cuentas;
}

function colppyCuentaPorCodigo(codigo) {
  const c = colppyCuentas().find(x => String(x.idPlanCuenta).trim() === String(codigo).trim());
  if (!c) throw new Error(`la cuenta ${codigo} no está en el plan de cuentas imputables de Colppy`);
  return c;
}

// Libro diario entre dos fechas (yyyy-MM-dd), de todas las cuentas. Se cachea por corrida.
// Producción acepta fromDate / toDate en aaaa-mm-dd y devuelve la lista en "movimientos".
// Sin paginar devuelve como máximo 50 movimientos: se pide por páginas con start / limit y se
// controla que el paginado funcione (si Colppy lo ignora, repite movimientos y se frena).
function colppyMovimientos(desdeStr, hastaStr) {
  const k = 'movs|' + desdeStr + '|' + hastaStr;
  if (_colppyCache[k]) return _colppyCache[k];
  const lista = [];
  const vistos = new Set();
  let niif = 0, total = 0, paginas = 0, start = 0;
  while (true) {
    if (++paginas > COLPPY.MAX_PAGINAS) {
      throw new Error(`el libro diario tiene más de ${COLPPY.MAX_PAGINAS * COLPPY.LOTE} movimientos: subí MAX_PAGINAS`);
    }
    const r = colppyLlamar('Contabilidad', 'listar_movimientosdiario',
      { fromDate: desdeStr, toDate: hastaStr, start, limit: COLPPY.LOTE });
    const crudos = Array.isArray(r.movimientos) ? r.movimientos : (Array.isArray(r.data) ? r.data : null);
    if (!crudos) throw new Error('listar_movimientosdiario no devolvió la lista "movimientos" (claves: ' + Object.keys(r).join(', ') + ')');
    let nuevos = 0;
    crudos.forEach(m => {
      const id = String(m.idDiario || '');
      if (!id) throw new Error('movimiento sin idDiario: ' + JSON.stringify(m).slice(0, 200));
      if (vistos.has(id)) return;
      vistos.add(id);
      nuevos++;
      total++;
      if (String(m.isNIIF) === '1') { niif++; return; }
      const fecha = colppyFecha(m.fechaContable);
      const dc = String(m.DebitoCredito || '').trim().toUpperCase();
      if (!fecha) throw new Error('movimiento sin fechaContable legible: ' + JSON.stringify(m).slice(0, 200));
      if (dc !== 'D' && dc !== 'C') throw new Error('movimiento con DebitoCredito inesperado: ' + JSON.stringify(m).slice(0, 200));
      // Si Colppy ignoró el filtro de fechas, el saldo saldría mal: se frena.
      if (fecha < desdeStr || fecha > hastaStr) {
        throw new Error(`Colppy devolvió un movimiento del ${fecha} fuera del rango pedido (${desdeStr} a ${hastaStr}): ignoró el filtro de fechas`);
      }
      const importe = colppyNum(m.Importe);
      lista.push({ cuenta: String(m.idPlanCuenta || '').trim(), fecha, importe: dc === 'D' ? importe : -importe });
    });
    if (crudos.length && !nuevos) {
      throw new Error(`Colppy repitió los mismos movimientos en la página ${paginas}: no acepta el paginado start / limit`);
    }
    if (crudos.length < COLPPY.LOTE) {
      // Página incompleta = última. Si vino justo el tope viejo de 50 con lote mayor, el limit se ignoró.
      if (paginas === 1 && crudos.length === 50 && COLPPY.LOTE !== 50) {
        throw new Error('Colppy devolvió exactamente 50 movimientos: ignoró el parámetro limit, el libro diario está incompleto');
      }
      break;
    }
    start += crudos.length;
  }
  lista.niif = niif;
  lista.total = total;
  lista.paginas = paginas;
  _colppyCache[k] = lista;
  return lista;
}

// Todo el historial, desde el inicio hasta un año adelante (incluye asientos con fecha futura).
function colppyDiarioCompleto() {
  return colppyMovimientos(COLPPY.DIARIO_DESDE, sumarDias(hoyStr(), COLPPY.DIAS_FUTURO));
}

// Movimientos del diario agrupados por código de cuenta. Se arma una vez por corrida.
function colppyPorCuenta() {
  if (_colppyCache.porCuenta) return _colppyCache.porCuenta;
  const por = {};
  colppyDiarioCompleto().forEach(m => (por[m.cuenta] = por[m.cuenta] || []).push(m));
  _colppyCache.porCuenta = por;
  return por;
}

// Saldo de una cuenta al cierre de fechaStr = suma (Debe − Haber) de sus movimientos hasta esa fecha.
function colppySaldoAFecha(codigo, fechaStr) {
  const c = colppyCuentaPorCodigo(codigo);
  const cod = String(c.idPlanCuenta).trim();
  const hoy = hoyStr();
  let valor = 0, n = 0, total = 0, futuro = 0, nFuturo = 0;
  (colppyPorCuenta()[cod] || []).forEach(m => {
    total += m.importe;
    if (m.fecha <= fechaStr) { valor += m.importe; n++; }
    if (m.fecha > hoy) { futuro += m.importe; nFuturo++; }
  });
  return {
    descripcion: c.Descripcion,
    valor: colppyRedondear(valor),
    n,
    total: colppyRedondear(total),
    nFuturo,
    futuro: colppyRedondear(futuro)
  };
}

function colppyBancos() {
  const r = colppyLlamar('Tesoreria', 'listado_banco', {});
  return Array.isArray(r.data) ? r.data : [];
}

// ---------- MAPEO ----------
function hojaMapeoColppy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(COLPPY.HOJA_MAPEO);
  if (!sh) {
    sh = ss.insertSheet(COLPPY.HOJA_MAPEO);
    sh.getRange('C:C').setNumberFormat('@'); // códigos como texto
    sh.getRange(1, 1, 1, HEADERS_MAPEO_COLPPY.length).setValues([HEADERS_MAPEO_COLPPY]).setFontWeight('bold');
    const filas = MAPEO_COLPPY_INICIAL.map(r => [r[0], r[1], r[2], '', r[3]]);
    sh.getRange(2, 1, filas.length, HEADERS_MAPEO_COLPPY.length).setValues(filas);
    sh.setFrozenRows(1);
    return sh;
  }
  const hdr = sh.getRange(1, 1, 1, HEADERS_MAPEO_COLPPY.length).getValues()[0].map(x => String(x).trim());
  if (hdr.join('|') !== HEADERS_MAPEO_COLPPY.join('|')) {
    throw new Error(`la hoja ${COLPPY.HOJA_MAPEO} no tiene los encabezados esperados (${HEADERS_MAPEO_COLPPY.join(' | ')})`);
  }
  return sh;
}

function leerMapeoColppy() {
  const sh = hojaMapeoColppy();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_MAPEO_COLPPY.length).getValues()
    .map((r, i) => ({
      cuenta: String(r[0]).trim(),
      moneda: String(r[1]).trim().toUpperCase(),
      codigo: String(r[2]).trim(),
      notas: String(r[4]).trim(),
      fila: i + 2
    }))
    .filter(m => m.cuenta && m.codigo);
}

// ---------- RESPALDO DIARIO (lo usa todasLasFuentes de saldosbasecp.gs) ----------
// Cuentas en pesos del mapeo: {"Banco Galicia|ARS": ['111100'], "FCI Galicia|ARS": ['113001', …]}.
function colppyGrupos() {
  const grupos = {};
  leerMapeoColppy().filter(m => m.moneda === COLPPY.MONEDA).forEach(m => {
    const k = m.cuenta + '|' + m.moneda;
    (grupos[k] = grupos[k] || []).push(m.codigo);
  });
  return grupos;
}

function colppySaldoGrupo(codigos, fechaStr) {
  return colppyRedondear(codigos.reduce((t, cod) => t + colppySaldoAFecha(cod, fechaStr).valor, 0));
}

// Envuelve las fuentes: cada cuenta del mapeo pasa a leerse con saldoConColppy. Las que no tenían
// fuente propia quedan con fuente "Colppy". Si el mapeo no se puede leer, las fuentes quedan como estaban.
function conColppy(fuentes) {
  let grupos;
  try {
    grupos = colppyGrupos();
  } catch (e) {
    console.error('Colppy: no se pudo leer el mapeo, sin respaldo de Colppy: ' + e.message);
    return fuentes;
  }
  Object.keys(grupos).forEach(clave => {
    const propia = fuentes[clave];
    const codigos = grupos[clave];
    fuentes[clave] = {
      fuente: propia ? propia.fuente : FUENTE_COLPPY,
      unaVez: propia ? propia.unaVez : false,
      fn: ctx => saldoConColppy(ctx, clave, propia, codigos)
    };
  });
  return fuentes;
}

// Fuente propia al día → manda. Si no, Colppy, con control contra el último dato propio.
// Devuelve el mismo formato que las demás fuentes, más "fuente" cuando el saldo sale de Colppy.
function saldoConColppy(ctx, clave, propia, codigos) {
  let p = null, errPropia = '';
  if (propia) {
    try {
      const r = propia.fn(ctx);
      if (r != null) {
        p = typeof r === 'number' ? { valor: r, fresco: true } : r;
        if (!isFinite(p.valor)) throw new Error('saldo no numérico: ' + p.valor);
      }
    } catch (e) {
      p = null;
      errPropia = e.message;
    }
  }
  const propiaAlDia = !!p && p.fresco !== false && (!p.estado || p.estado.indexOf('al día') === 0);

  let c = null, errColppy = '';
  try {
    c = colppySaldoGrupo(codigos, ctx.fechaStr);
  } catch (e) {
    errColppy = e.message;
  }

  const control = c == null ? null : colppyControl(clave, c, p);
  const usada = propiaAlDia ? propia.fuente : (c != null ? FUENTE_COLPPY : (p ? propia.fuente : ''));
  try {
    registrarControlColppy(ctx, clave, c, propia, p, control, usada, errColppy);
  } catch (e) {
    console.error(`Control Colppy ${clave}: ${e.message}`);
  }

  if (propiaAlDia) return p;
  if (c == null) {
    if (p) return Object.assign({}, p, { estado: (p.estado || 'último disponible') + ' · Colppy falló: ' + errColppy });
    throw new Error(`sin dato propio${errPropia ? ' (' + errPropia + ')' : ''} y Colppy falló: ${errColppy}`);
  }
  const estado = control.estado + (errPropia ? ` · ${propia.fuente} falló: ${errPropia}` : '');
  return { valor: c, fresco: control.ok, fuente: FUENTE_COLPPY, estado };
}

// Compara Colppy con el último dato de la fuente propia (p puede ser null: sin control).
function colppyControl(clave, c, p) {
  if (!p) return { ok: true, resultado: 'sin control', estado: 'al día (Colppy, sin control)' };
  const dif = colppyRedondear(c - p.valor);
  const pct = p.valor ? Math.abs(dif) / Math.abs(p.valor) : (dif ? Infinity : 0);
  const umbral = clave.indexOf('FCI') === 0 ? COLPPY.UMBRAL_FCI : COLPPY.UMBRAL_OTRAS;
  const pctTxt = isFinite(pct) ? (pct * 100).toFixed(1) + '%' : 'sin base';
  const ok = Math.abs(dif) <= COLPPY.PISO_ARS || pct <= umbral;
  return {
    ok, dif, pct,
    resultado: ok ? 'OK' : 'DIFERENCIA',
    estado: ok ? `al día (Colppy; difiere ${pctTxt} del último dato propio)`
               : `Colppy: difiere ${pctTxt} (${dif}) del último dato propio (${p.valor}); umbral ${umbral * 100}%`
  };
}

// Hoja "Control Colppy": una fila por fecha y cuenta; si ya existe, se reescribe.
function registrarControlColppy(ctx, clave, c, propia, p, control, usada, errColppy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(COLPPY.HOJA_CONTROL);
  if (!sh) {
    sh = ss.insertSheet(COLPPY.HOJA_CONTROL);
    sh.getRange(1, 1, 1, HEADERS_CONTROL_COLPPY.length).setValues([HEADERS_CONTROL_COLPPY]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy');
    sh.getRange('D:G').setNumberFormat('#,##0.00');
    sh.getRange('H:H').setNumberFormat('0.0%');
    sh.getRange('K:K').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.setFrozenRows(1);
  }
  if (!_colppyCache.control) {
    const idx = {};
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach((r, i) => {
        if (r[0] instanceof Date) {
          idx[Utilities.formatDate(r[0], ctx.tz, 'yyyy-MM-dd') + '|' + String(r[1]).trim() + '|' + String(r[2]).trim()] = i + 2;
        }
      });
    }
    _colppyCache.control = idx;
  }
  const [cuenta, moneda] = clave.split('|');
  const resultado = c == null ? 'error Colppy: ' + errColppy : control.resultado;
  const fila = [ctx.ini, cuenta, moneda, c == null ? '' : c, propia ? propia.fuente : '', p ? p.valor : '',
                control && p ? control.dif : '', control && p && isFinite(control.pct) ? control.pct : '',
                resultado, usada, new Date()];
  const k = ctx.fechaStr + '|' + clave;
  const n = _colppyCache.control[k] || sh.getLastRow() + 1;
  _colppyCache.control[k] = n;
  sh.getRange(n, 1, 1, fila.length).setValues([fila]);
  sh.getRange(n, 9).setBackground(c != null && control.ok ? null : BEIGE);
}

// Fechas a releer entre las filas de Saldos con fuente "Colppy" (filas: [{fecha, clave, valor}]):
//   - el saldo de Colppy de esa fecha cambió (el contador cargó asientos después), o
//   - la cuenta tiene fuente propia y la fecha es posterior a desdePropia (el CSV pudo subirse tarde).
function colppyFechasARepasar(filas, fuentes, desdePropia) {
  const grupos = colppyGrupos();
  const fechas = new Set();
  filas.forEach(x => {
    if (fechas.has(x.fecha)) return;
    const fu = fuentes[x.clave];
    if (fu && fu.fuente !== FUENTE_COLPPY && x.fecha >= desdePropia) {
      fechas.add(x.fecha);
      return;
    }
    const codigos = grupos[x.clave];
    if (!codigos) return;
    if (Math.abs(colppySaldoGrupo(codigos, x.fecha) - x.valor) >= 0.01) fechas.add(x.fecha);
  });
  return fechas;
}

// ---------- LISTADO ----------
// Vuelca el plan de cuentas y los bancos, y completa "descripción Colppy" en el mapeo.
function listarColppy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cuentas = colppyCuentas();
  if (!cuentas.length) throw new Error('listar_cuentasdiario no devolvió cuentas');
  const bancos = colppyBancos();

  const shC = ss.getSheetByName(COLPPY.HOJA_CUENTAS) || ss.insertSheet(COLPPY.HOJA_CUENTAS);
  shC.clearContents();
  shC.getRange('B:B').setNumberFormat('@');
  shC.getRange(1, 1, 1, HEADERS_CUENTAS_COLPPY.length).setValues([HEADERS_CUENTAS_COLPPY]).setFontWeight('bold');
  shC.getRange(2, 1, cuentas.length, HEADERS_CUENTAS_COLPPY.length)
    .setValues(cuentas.map(c => [String(c.Id || ''), String(c.idPlanCuenta || ''), String(c.Descripcion || ''), String(c.idTipoCuenta || '')]));
  shC.setFrozenRows(1);

  const shB = ss.getSheetByName(COLPPY.HOJA_BANCOS) || ss.insertSheet(COLPPY.HOJA_BANCOS);
  shB.clearContents();
  shB.getRange(1, 1, 1, HEADERS_BANCOS_COLPPY.length).setValues([HEADERS_BANCOS_COLPPY]).setFontWeight('bold');
  if (bancos.length) {
    shB.getRange(2, 1, bancos.length, HEADERS_BANCOS_COLPPY.length).setValues(bancos.map(b => [
      String(b.idBanco || ''), String(b.Nombre || ''), String(b.Banco || ''), String(b.nroCuenta || ''),
      String(b.Moneda || ''), String(b.idPlanCuenta || ''), String(b.planCuentaId || ''),
      b.saldo === undefined || b.saldo === null || b.saldo === '' ? '' : colppyNum(b.saldo),
      String(b.ultimaConciliacion || ''), b.esMultimoneda === undefined ? '' : String(b.esMultimoneda)
    ]));
    shB.getRange(2, 8, bancos.length, 1).setNumberFormat('#,##0.00');
  }
  shB.setFrozenRows(1);

  const map = hojaMapeoColppy();
  if (map.getLastRow() > 1) {
    const codigos = map.getRange(2, 3, map.getLastRow() - 1, 1).getValues();
    const desc = codigos.map(r => {
      const cod = String(r[0]).trim();
      if (!cod) return [''];
      const c = cuentas.find(x => String(x.idPlanCuenta).trim() === cod);
      return [c ? c.Descripcion : 'NO EXISTE en Colppy'];
    });
    map.getRange(2, 4, desc.length, 1).setValues(desc);
  }
  return `${cuentas.length} cuentas en "${COLPPY.HOJA_CUENTAS}" y ${bancos.length} banco(s) en "${COLPPY.HOJA_BANCOS}". ` +
         `Revisá la columna "descripción Colppy" de "${COLPPY.HOJA_MAPEO}".`;
}

// ---------- BALANCE ----------
// Último mes cerrado, en yyyy-MM.
function colppyUltimoMesCerrado() {
  return sumarDias(hoyStr().slice(0, 8) + '01', -1).slice(0, 7);
}

// Balance de sumas y saldos al último día del mes (yyyy-MM), calculado con el libro diario:
// saldo inicial = Debe − Haber hasta el último día del mes anterior; debe y haber = movimientos del mes;
// saldo al cierre = inicial + debe − haber. Se reemplazan las filas de ese mes en "Balance Colppy".
function descargarBalanceColppy(mesStr) {
  const mes = String(mesStr || colppyUltimoMesCerrado()).trim();
  const m = mes.match(/^(\d{4})-(\d{2})$/);
  if (!m || +m[2] < 1 || +m[2] > 12) throw new Error(`mes inválido: ${mes}. Tiene que ser yyyy-MM (ejemplo 2026-09)`);
  const ini = mes + '-01';
  const siguiente = +m[2] === 12 ? `${+m[1] + 1}-01-01` : `${m[1]}-${String(+m[2] + 1).padStart(2, '0')}-01`;
  const fin = sumarDias(siguiente, -1);
  if (fin >= hoyStr()) throw new Error(`el ${fmtCorto(fin)} todavía no cerró: solo meses terminados`);

  const plan = {};
  colppyCuentas().forEach(c => (plan[String(c.idPlanCuenta).trim()] = c));
  const por = {};
  colppyDiarioCompleto().forEach(x => {
    if (x.fecha > fin) return;
    const a = por[x.cuenta] = por[x.cuenta] || { inicial: 0, debe: 0, haber: 0 };
    if (x.fecha < ini) a.inicial += x.importe;
    else if (x.importe >= 0) a.debe += x.importe;
    else a.haber -= x.importe;
  });

  const tz = Session.getScriptTimeZone();
  const fechaMes = Utilities.parseDate(fin, tz, 'yyyy-MM-dd');
  const ahora = new Date();
  let sinPlan = 0, totSaldo = 0, totDebe = 0, totHaber = 0;
  const nuevas = Object.keys(por)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
    .map(cod => {
      const a = por[cod];
      const inicial = colppyRedondear(a.inicial), debe = colppyRedondear(a.debe), haber = colppyRedondear(a.haber);
      if (!inicial && !debe && !haber) return null;
      const c = plan[cod];
      if (!c) sinPlan++;
      const saldo = colppyRedondear(inicial + debe - haber);
      totSaldo += saldo; totDebe += debe; totHaber += haber;
      return [fechaMes, cod, c ? String(c.Descripcion || '') : '(no está en el plan de cuentas imputables)',
              c ? String(c.idTipoCuenta || '') : '', inicial, debe, haber, saldo, ahora];
    })
    .filter(Boolean);
  if (!nuevas.length) throw new Error(`Colppy no tiene movimientos hasta el ${fmtCorto(fin)}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(COLPPY.HOJA_BALANCE) || ss.insertSheet(COLPPY.HOJA_BALANCE);
  const n = HEADERS_BALANCE_COLPPY.length;
  const otras = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues()
        .filter(r => r[0] instanceof Date && Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') !== fin)
    : [];
  const filas = otras.concat(nuevas).sort((a, b) =>
    (a[0].getTime() - b[0].getTime()) || String(a[1]).localeCompare(String(b[1]), undefined, { numeric: true }));
  sh.clearContents();
  sh.getRange(1, 1, 1, n).setValues([HEADERS_BALANCE_COLPPY]).setFontWeight('bold');
  sh.getRange(2, 2, filas.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, filas.length, n).setValues(filas);
  sh.getRange(2, 1, filas.length, 1).setNumberFormat('yyyy-MM-dd');
  sh.getRange(2, 5, filas.length, 4).setNumberFormat('#,##0.00');
  sh.getRange(2, 9, filas.length, 1).setNumberFormat('yyyy-MM-dd HH:mm');
  sh.setFrozenRows(1);

  totSaldo = colppyRedondear(totSaldo);
  const cuadra = Math.abs(totSaldo) < 0.01 && Math.abs(colppyRedondear(totDebe - totHaber)) < 0.01;
  return `Balance al ${fmtCorto(fin)}: ${nuevas.length} cuentas en "${COLPPY.HOJA_BALANCE}". ` +
         (cuadra ? 'Cuadra: saldos suman 0 y debe = haber del mes.'
                 : `NO CUADRA: saldos suman ${totSaldo}, debe del mes ${colppyRedondear(totDebe)}, haber ${colppyRedondear(totHaber)}.`) +
         (sinPlan ? ` ${sinPlan} código(s) con movimientos no están en el plan de cuentas imputables.` : '');
}

// ---------- DIAGNÓSTICO ----------
// Saldo de cada cuenta mapeada al cierre de la fecha, comparado con lo que hoy tiene la hoja Saldos,
// y control del cálculo: saldo calculado de cada banco contra el saldo que informa Tesorería.
function probarColppy(fechaStr) {
  const tz = Session.getScriptTimeZone();
  const ayer = sumarDias(hoyStr(), -1);
  const f = fechaStr || ayer;
  if (f > ayer) throw new Error(`el ${f} es hoy o futuro: solo días cerrados`);

  const s = colppySesion(false);
  const lineas = [
    `Entorno: ${colppyUrl() === COLPPY.URL_STAGING ? 'staging' : 'producción'} · usuario ${s.usuario} · empresa ${s.idEmpresa}`
  ];

  const diario = colppyDiarioCompleto();
  const fechas = diario.map(m => m.fecha).sort();
  lineas.push(`Libro diario: ${diario.total} movimiento(s) leídos en ${diario.paginas} página(s)` +
              (diario.niif ? `, ${diario.niif} NIIF excluidos` : '') +
              (fechas.length ? ` · del ${fmtCorto(fechas[0])} al ${fmtCorto(fechas[fechas.length - 1])}` : ''));

  // Control del cálculo: bancos de Tesorería (saldo del momento) contra la suma de todo el diario
  lineas.push('');
  lineas.push('Control: saldo calculado (todo el diario) contra Tesorería:');
  try {
    const bancos = colppyBancos();
    if (!bancos.length) lineas.push('  (Tesorería no devolvió bancos)');
    bancos.forEach(b => {
      const cod = String(b.planCuentaId || '').trim();
      const tes = colppyNum(b.saldo);
      const calc = colppyRedondear(diario.filter(m => m.cuenta === cod).reduce((t, m) => t + m.importe, 0));
      const dif = colppyRedondear(calc - tes);
      lineas.push(`  ${cod} ${b.Nombre} (${b.Moneda}): calculado ${calc} · Tesorería ${tes} · ` +
                  (Math.abs(dif) < 0.01 ? 'OK' : `DIFERENCIA ${dif}`));
    });
  } catch (e) {
    lineas.push('  ERROR ' + e.message);
  }

  // Valor de la hoja Saldos de esa fecha, para comparar
  const enSaldos = {};
  leerSaldos(hojaSaldos()).forEach(r => {
    if (!(r[0] instanceof Date) || Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') !== f) return;
    const v = valorFila(r);
    if (v.valido) enSaldos[String(r[1]).trim() + '|' + String(r[2]).trim()] = { valor: v.valor, fuente: String(r[4]).trim() };
  });

  const grupos = {};
  leerMapeoColppy().forEach(m => {
    const k = m.cuenta + '|' + m.moneda;
    (grupos[k] = grupos[k] || []).push(m);
  });
  lineas.push('');
  lineas.push(`Saldos al ${fmtCorto(f)} (suma Debe − Haber hasta esa fecha):`);
  if (!Object.keys(grupos).length) lineas.push(`La hoja "${COLPPY.HOJA_MAPEO}" no tiene códigos cargados.`);

  Object.keys(grupos).forEach(k => {
    lineas.push('');
    lineas.push(k);
    let total = 0, ok = true;
    grupos[k].forEach(m => {
      try {
        const x = colppySaldoAFecha(m.codigo, f);
        total += x.valor;
        lineas.push(`  ${x.descripcion}: ${x.valor} (${x.n} mov.) · hoy ${x.total}` +
                    (x.nFuturo ? ` · ${x.nFuturo} con fecha futura por ${x.futuro}` : ''));
      } catch (e) {
        ok = false;
        lineas.push(`  ${m.codigo}: ERROR ${e.message}`);
      }
    });
    if (!ok) return;
    total = colppyRedondear(total);
    const h = enSaldos[k];
    if (!h) {
      lineas.push(`  = ${total} · en Saldos: sin fila del ${fmtCorto(f)}`);
    } else {
      const dif = colppyRedondear(total - h.valor);
      const pct = h.valor ? ` (${(dif / Math.abs(h.valor) * 100).toFixed(1)}%)` : '';
      lineas.push(`  = ${total} · en Saldos ${h.valor} (${h.fuente}) · diferencia ${dif}${pct}`);
    }
  });

  const texto = lineas.join('\n');
  console.log(texto);
  return texto;
}
