/**
 * Colppy (contabilidad de la SRL) → EXTRACTOS BANCARIOS. Archivo nuevo en el proyecto de EXTRACTOS.
 *
 * FASE 1: SOLO DIAGNÓSTICO. Este archivo se conecta a Colppy, lista cuentas y bancos y calcula el
 * saldo de las cuentas mapeadas a una fecha, para comparar contra la hoja Saldos. NO escribe en
 * Saldos ni en Base CP y no cambia ninguna fuente. La integración (Galicia con respaldo de Colppy,
 * FCI e ICBC desde Colppy con control) es la fase 2, una vez validados los números.
 *
 * Solo lectura: el script únicamente puede llamar a las operaciones de COLPPY_OPERACIONES_PERMITIDAS.
 * Cualquier otra (alta_asiento, borrar_asiento, etc.) se rechaza antes de llegar a Colppy.
 *
 * Saldo a una fecha: la API no tiene "saldo a fecha". Se toma el saldo actual de la cuenta
 * (Contabilidad/leer_saldoCuenta) y se le restan los movimientos con fecha contable posterior
 * (Contabilidad/listar_movimientosdiario), con signo Debe − Haber. VERIFICAR con "Probar Colppy":
 *   - que el signo dé bien (Banco Galicia tiene que dar el mismo signo que en Colppy);
 *   - si leer_saldoCuenta incluye movimientos con fecha futura (la prueba los informa aparte).
 *
 * Hojas:
 *   "Mapeo Colppy"   cuenta | moneda | código Colppy | descripción Colppy | notas
 *                    Varias filas con la misma cuenta y moneda se suman (FCI Galicia = 4 cuentas FIMA).
 *                    "descripción Colppy" la completa "Listar cuentas y bancos".
 *   "Cuentas Colppy" volcado del plan de cuentas imputables (se reescribe en cada listado).
 *   "Bancos Colppy"  volcado de Tesorería/listado_banco (se reescribe en cada listado).
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
  DIAS_FUTURO: 366,              // movimientos con fecha futura que se descuentan del saldo actual
  SESION_MS: 11 * 3600 * 1000    // la sesión dura 12 h desde el último uso; se renueva antes
};
const HEADERS_MAPEO_COLPPY = ['cuenta', 'moneda', 'código Colppy', 'descripción Colppy', 'notas'];
const MAPEO_COLPPY_INICIAL = [
  ['Banco Galicia', 'ARS', '111100', 'fase 2: respaldo si no hay CSV'],
  ['Banco ICBC', 'ARS', '111101', 'fase 2: fuente, control 1% contra CSV'],
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
  Contabilidad: ['listar_cuentasdiario', 'leer_saldoCuenta', 'listar_movimientosdiario'],
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

function colppyDMY(fechaStr) {
  return fechaStr.split('-').reverse().join('-'); // yyyy-MM-dd → dd-MM-yyyy
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

function colppySaldoActual(descripcion) {
  const r = colppyLlamar('Contabilidad', 'leer_saldoCuenta', { planCuenta: descripcion });
  const v = r.saldo !== undefined ? r.saldo : (r.data && r.data.saldo);
  if (v === undefined || v === null || v === '') throw new Error(`leer_saldoCuenta no devolvió saldo para ${descripcion}`);
  return colppyNum(v);
}

// Movimientos del diario entre dos fechas (yyyy-MM-dd), de todas las cuentas. Se cachea por corrida.
function colppyMovimientos(desdeStr, hastaStr) {
  const k = 'movs|' + desdeStr + '|' + hastaStr;
  if (_colppyCache[k]) return _colppyCache[k];
  const r = colppyLlamar('Contabilidad', 'listar_movimientosdiario',
    { fechaDesde: colppyDMY(desdeStr), fechaHasta: colppyDMY(hastaStr) });
  const lista = (Array.isArray(r.data) ? r.data : []).map(m => ({
    desc: String(m.Descripcion || '').trim(),
    debe: colppyNum(m.Debito),
    haber: colppyNum(m.Credito),
    fecha: colppyFecha(m.fechaContable)
  }));
  _colppyCache[k] = lista;
  return lista;
}

function colppyEsDeCuenta(mov, cuenta) {
  return mov.desc === String(cuenta.Descripcion).trim() ||
         mov.desc.split(' - ')[0].trim() === String(cuenta.idPlanCuenta).trim();
}

// Saldo de una cuenta al cierre de fechaStr = saldo actual − (Debe − Haber) de los movimientos posteriores.
function colppySaldoAFecha(codigo, fechaStr) {
  const c = colppyCuentaPorCodigo(codigo);
  const actual = colppySaldoActual(c.Descripcion);
  const hoy = hoyStr();
  const movs = colppyMovimientos(sumarDias(fechaStr, 1), sumarDias(hoy, COLPPY.DIAS_FUTURO))
    .filter(m => colppyEsDeCuenta(m, c));
  let posterior = 0, futuro = 0, nFuturo = 0, sinFecha = 0;
  movs.forEach(m => {
    const d = m.debe - m.haber;
    posterior += d;
    if (!m.fecha) sinFecha++;
    else if (m.fecha > hoy) { futuro += d; nFuturo++; }
  });
  return {
    descripcion: c.Descripcion,
    actual,
    n: movs.length,
    posterior: colppyRedondear(posterior),
    nFuturo,
    futuro: colppyRedondear(futuro),
    sinFecha,
    valor: colppyRedondear(actual - posterior)
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

// ---------- DIAGNÓSTICO ----------
// Saldo de cada cuenta mapeada al cierre de la fecha, comparado con lo que hoy tiene la hoja Saldos.
function probarColppy(fechaStr) {
  const tz = Session.getScriptTimeZone();
  const ayer = sumarDias(hoyStr(), -1);
  const f = fechaStr || ayer;
  if (f > ayer) throw new Error(`el ${f} es hoy o futuro: solo días cerrados`);

  const s = colppySesion(false);
  const lineas = [
    `Entorno: ${colppyUrl() === COLPPY.URL_STAGING ? 'staging' : 'producción'} · usuario ${s.usuario} · empresa ${s.idEmpresa}`,
    `Saldos al ${fmtCorto(f)} (saldo actual − movimientos posteriores, signo Debe − Haber):`
  ];

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
  if (!Object.keys(grupos).length) lineas.push(`La hoja "${COLPPY.HOJA_MAPEO}" no tiene códigos cargados.`);

  Object.keys(grupos).forEach(k => {
    lineas.push('');
    lineas.push(k);
    let total = 0, ok = true;
    grupos[k].forEach(m => {
      try {
        const x = colppySaldoAFecha(m.codigo, f);
        total += x.valor;
        lineas.push(`  ${x.descripcion}: ${x.valor} = actual ${x.actual} − ${x.n} mov. posterior(es) por ${x.posterior}` +
                    (x.nFuturo ? ` (de ellos ${x.nFuturo} con fecha futura por ${x.futuro})` : '') +
                    (x.sinFecha ? ` ⚠ ${x.sinFecha} sin fecha legible` : ''));
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

  lineas.push('');
  lineas.push('Bancos (Tesorería/listado_banco, saldo del momento):');
  try {
    const bancos = colppyBancos();
    if (!bancos.length) lineas.push('  (sin bancos)');
    bancos.forEach(b => lineas.push(`  ${b.Nombre} · ${b.Moneda} · saldo ${b.saldo} · multimoneda ${b.esMultimoneda} · ` +
                                    `conciliado al ${b.ultimaConciliacion || '—'} · cuenta ${b.planCuentaId || b.idPlanCuenta}`));
  } catch (e) {
    lineas.push('  ERROR ' + e.message);
  }

  const texto = lineas.join('\n');
  console.log(texto);
  return texto;
}
