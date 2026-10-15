/**
 * MONTO A PAGAR DESDE LAS ÓRDENES DE PAGO DE COLPPY
 *
 * abrirOrdenesColppy() → ítem de menú. Subís los PDF de orden de pago que genera Colppy (la orden,
 * con o sin el certificado de retención en el mismo archivo) y el script completa MONTO A PAGAR en
 * las filas A de la agrupada de hoy.
 *
 * QUÉ SE LEE DE CADA PDF (en el navegador, con pdf.js; no usa el OCR de Drive)
 *   proveedor      la línea que sigue a "Datos del Proveedor";
 *   este pago      suma de "Este Pago" de los comprobantes de la ORDEN (no del certificado);
 *   retenido       suma de los importes entre paréntesis de "Retenciones Realizadas";
 *   total a pagar  "Total a Pagar: 186.620,00".
 * Control: este pago − retenido tiene que dar el total a pagar. Si no cierra, no se usa.
 *
 * CON QUÉ FILA SE CRUZA
 * Fila tipo A de la agrupada de hoy, con ESTADO vacío o CARGADO EN BANCO, cuyo BENEFICIARIO tiene
 * las mismas palabras que el proveedor (mismoConjuntoDePalabras_) y cuyo MONTO (el de la factura)
 * es igual a "este pago". Hacen falta las dos cosas y una sola fila: si no, no se escribe nada y se
 * dice por qué.
 *
 * QUÉ ESCRIBE
 * MONTO A PAGAR = total a pagar, solo si la celda está vacía. Si ya tiene otro valor no la pisa:
 * lo muestra como conflicto y lo resolvés a mano.
 *
 * DEPENDE de Pagoslb.gs (formatearNombreHoja_, mismoConjuntoDePalabras_, aNumero_, redondear2_,
 * formatearImporte_, EST_CARGADO), Hojaagrupada.gs (AGRUP_PREFIJO, colsAgrup_, valAgrup_) y
 * CargaGalicia.gs (CG_TIPO_CON_RETENCION, CG_HEADER_PAGAR).
 */

function abrirOrdenesColppy() {
  const nombre = AGRUP_PREFIJO + formatearNombreHoja_(new Date());
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No existe la hoja "' + nombre + '". Generá primero la agrupada.');
    return;
  }
  if (colsAgrup_(hoja).PAGAR < 0) {
    SpreadsheetApp.getUi().alert('La agrupada de hoy no tiene la columna ' + CG_HEADER_PAGAR +
                                 '. Corré "Actualizar hoja agrupada del día" primero.');
    return;
  }
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(ocHtml_()).setWidth(760).setHeight(520),
    'MONTO A PAGAR desde Colppy'
  );
}

/** Cruza las órdenes leídas con las filas A. aplicar = true escribe; false solo informa. */
function ocCruzar(ordenes, aplicar) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(AGRUP_PREFIJO + formatearNombreHoja_(new Date()));
  if (!hoja) return { ok: false, motivo: 'No existe la agrupada de hoy.' };
  const C = colsAgrup_(hoja);
  if (C.PAGAR < 0 || C.IDS < 0 || C.TIPO < 0) return { ok: false, motivo: 'La agrupada no tiene las columnas esperadas.' };

  const valores = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 1), C.ANCHO).getValues();
  const filasA = [];
  valores.forEach((v, i) => {
    if (String(valAgrup_(v, C, 'TIPO')).trim() !== CG_TIPO_CON_RETENCION) return;
    if (String(valAgrup_(v, C, 'IDS') || '').trim() === '') return;
    const est = String(valAgrup_(v, C, 'ESTADO') || '').trim().toUpperCase();
    if (est !== '' && est !== EST_CARGADO) return;
    filasA.push({
      fila: i + 2,
      benef: String(valAgrup_(v, C, 'BENEF') || ''),
      ids: String(valAgrup_(v, C, 'IDS')),
      monto: aNumero_(valAgrup_(v, C, 'MONTO')),
      pagar: aNumero_(valAgrup_(v, C, 'PAGAR'))
    });
  });

  const res = [];
  const usadas = {};
  (ordenes || []).forEach(o => {
    const r = { archivo: o.archivo, proveedor: o.nombre || '?', totalAPagar: o.totalAPagar, estado: '', detalle: '', fila: null };
    res.push(r);
    if (o.error) { r.estado = 'error'; r.detalle = o.error; return; }
    if (o.totalAPagar === null || o.estePago === null) {
      r.estado = 'error'; r.detalle = 'no encontré "Total a Pagar" o los comprobantes pagados'; return;
    }
    if (Math.abs(redondear2_(o.estePago - o.retenido) - o.totalAPagar) > 0.01) {
      r.estado = 'error';
      r.detalle = 'no cierra: este pago ' + formatearImporte_(o.estePago) + ' − retenido ' +
                  formatearImporte_(o.retenido) + ' ≠ total a pagar ' + formatearImporte_(o.totalAPagar);
      return;
    }
    const porNombre = filasA.filter(f => mismoConjuntoDePalabras_(f.benef, o.nombre));
    const porAmbos = porNombre.filter(f => f.monto !== null && Math.abs(f.monto - o.estePago) <= 0.01);
    if (porAmbos.length === 1) {
      const f = porAmbos[0];
      r.fila = f.fila;
      r.detalle = f.benef + ' (' + f.ids + ') · factura ' + formatearImporte_(f.monto);
      if (usadas[f.fila]) { r.estado = 'conflicto'; r.detalle += ' · otra orden ya apunta a esta fila'; usadas[f.fila].estado = 'conflicto'; return; }
      usadas[f.fila] = r;
      if (f.pagar === null) r.estado = 'escribir';
      else if (Math.abs(f.pagar - o.totalAPagar) <= 0.01) r.estado = 'igual';
      else { r.estado = 'conflicto'; r.detalle += ' · ya tiene ' + formatearImporte_(f.pagar) + ', no se pisa'; }
      return;
    }
    r.estado = 'sin fila';
    if (porAmbos.length > 1) {
      r.detalle = porAmbos.length + ' filas coinciden en nombre y monto: cargalo a mano';
    } else if (porNombre.length) {
      r.detalle = 'el proveedor está, pero el monto no coincide: orden ' + formatearImporte_(o.estePago) +
                  ' vs ' + porNombre.map(f => (f.monto === null ? 'sin monto' : formatearImporte_(f.monto)) + ' (' + f.ids + ')').join(', ');
    } else {
      const porMonto = filasA.filter(f => f.monto !== null && Math.abs(f.monto - o.estePago) <= 0.01);
      r.detalle = 'ninguna fila A pendiente con ese proveedor' +
                  (porMonto.length ? '. Con ese monto está: ' + porMonto.map(f => f.benef).join(', ') + ' (revisá el nombre)' : '');
    }
  });

  let escritas = 0;
  if (aplicar) {
    res.forEach(r => {
      if (r.estado !== 'escribir') return;
      hoja.getRange(r.fila, C.PAGAR).setValue(r.totalAPagar)
        .setNote('Total a Pagar de la orden de Colppy (' + r.archivo + ').');
      r.estado = 'escrito';
      escritas++;
    });
  }
  return { ok: true, filas: res, escritas: escritas };
}

/**
 * La lectura del PDF es la misma que usa el panel de lote de Comprobantes.gs para la documentación
 * contable (texto con los saltos de línea de pdf.js), que ya funciona con estas órdenes.
 */
function ocHtml_() {
  return `
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"></script>
<style>
  body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; margin: 0; padding: 12px; color: #222; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 10px; }
  th, td { border: 1px solid #ddd; padding: 5px 7px; text-align: left; vertical-align: top; }
  th { background: #eee; }
  .escribir, .escrito { color: #137333; font-weight: bold; }
  .igual { color: #666; }
  .conflicto, .error, .sinfila { color: #cc0000; font-weight: bold; }
  button { padding: 8px 14px; font-size: 13px; border-radius: 4px; border: 1px solid #0a0a8c; background: #0a0a8c; color: #fff; font-weight: bold; cursor: pointer; margin-top: 10px; }
  button:disabled { opacity: .45; cursor: default; }
  #msg { margin-top: 8px; }
</style>
<div>Elegí los PDF de orden de pago de Colppy (pueden traer el certificado de retención en el mismo archivo).</div>
<input type="file" id="archivos" multiple accept="application/pdf" style="margin-top:8px">
<div id="msg"></div>
<div id="tabla"></div>
<button id="aplicar" style="display:none">Escribir MONTO A PAGAR</button>
<script>
  var ORDENES = [];
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function fmt(n) { return n === null || n === undefined ? '—' : Number(n).toFixed(2).replace('.', ','); }
  function importeAR(s) { return Number(String(s).replace(/\\./g, '').replace(',', '.')); }

  function seccion(t, desde, hasta) {
    var i = t.search(desde);
    if (i < 0) return '';
    var r = t.substring(i);
    var j = r.search(hasta);
    return j > 0 ? r.substring(0, j) : r;
  }

  function parsear(texto) {
    var lineas = texto.split('\\n').map(function (l) { return l.trim(); });
    var nombre = null;
    for (var k = 0; k < lineas.length; k++) {
      if (/^Datos del Proveedor/i.test(lineas[k])) { nombre = lineas[k + 1] || null; break; }
    }
    var tp = texto.match(/Total\\s+a\\s+Pagar\\s*:?\\s*\\$?\\s*([\\d.]+,\\d{2})/i);
    // Solo la orden: el certificado repite "Comprobantes Pagados" con la base imponible.
    var comp = seccion(texto, /Comprobantes\\s+Pagados/i, /Retenciones\\s+Realizadas|Intereses\\s*\\/|Total\\s+a\\s+Pagar/i);
    var estePago = null;
    comp.split('\\n').forEach(function (l) {
      if (!/^\\s*\\d{3,5}-\\d{6,8}\\b/.test(l)) return;
      var m = l.match(/([\\d.]+,\\d{2})\\s*$/);
      if (m) estePago = Math.round(((estePago || 0) + importeAR(m[1])) * 100) / 100;
    });
    var ret = seccion(texto, /Retenciones\\s+Realizadas/i, /Intereses\\s*\\/|Valores\\s+Entregados|Total\\s+a\\s+Pagar/i);
    var retenido = 0;
    (ret.match(/\\(([\\d.]+,\\d{2})\\)/g) || []).forEach(function (x) { retenido += importeAR(x.replace(/[()]/g, '')); });
    return {
      nombre: nombre,
      totalAPagar: tp ? importeAR(tp[1]) : null,
      estePago: estePago,
      retenido: Math.round(retenido * 100) / 100
    };
  }

  function leer(f) {
    return new Promise(function (ok) {
      var fr = new FileReader();
      fr.onload = function () {
        pdfjsLib.getDocument({ data: new Uint8Array(fr.result), isEvalSupported: false }).promise.then(function (doc) {
          var ps = [];
          for (var p = 1; p <= doc.numPages; p++) ps.push(doc.getPage(p).then(function (pg) { return pg.getTextContent(); }));
          return Promise.all(ps);
        }).then(function (cs) {
          var texto = cs.map(function (c) {
            return c.items.map(function (it) { return it.str + (it.hasEOL ? '\\n' : ' '); }).join('');
          }).join('\\n');
          var o = parsear(texto); o.archivo = f.name; ok(o);
        }, function (e) { ok({ archivo: f.name, error: 'no pude leer el PDF: ' + e.message }); });
      };
      fr.readAsArrayBuffer(f);
    });
  }

  var ETIQUETA = { escribir: 'se escribe', escrito: '✓ escrito', igual: 'ya estaba', conflicto: 'conflicto', error: 'error', 'sin fila': 'sin fila' };

  function pintar(r) {
    if (!r.ok) { $('msg').innerHTML = '<span class="error">' + esc(r.motivo) + '</span>'; return; }
    $('tabla').innerHTML = '<table><tr><th>Archivo</th><th>Proveedor</th><th>Total a pagar</th><th>Estado</th><th>Detalle</th></tr>' +
      r.filas.map(function (f) {
        return '<tr><td>' + esc(f.archivo) + '</td><td>' + esc(f.proveedor) + '</td><td>' + fmt(f.totalAPagar) +
          '</td><td class="' + f.estado.replace(' ', '') + '">' + esc(ETIQUETA[f.estado] || f.estado) + '</td><td>' + esc(f.detalle) + '</td></tr>';
      }).join('') + '</table>';
    var n = r.filas.filter(function (f) { return f.estado === 'escribir'; }).length;
    $('aplicar').style.display = n ? '' : 'none';
    $('aplicar').textContent = 'Escribir MONTO A PAGAR en ' + n + ' fila(s)';
    $('aplicar').disabled = false;
  }

  $('archivos').onchange = function () {
    var files = Array.prototype.slice.call(this.files);
    if (!files.length) return;
    if (!window.pdfjsLib) { $('msg').innerHTML = '<span class="error">No cargó pdf.js. Revisá la conexión y reabrí la ventana.</span>'; return; }
    $('msg').textContent = 'Leyendo ' + files.length + ' PDF…';
    Promise.all(files.map(leer)).then(function (os) {
      ORDENES = os;
      $('msg').textContent = 'Cruzando con la agrupada…';
      google.script.run.withSuccessHandler(function (r) { $('msg').textContent = ''; pintar(r); })
        .withFailureHandler(function (e) { $('msg').innerHTML = '<span class="error">' + esc(e.message) + '</span>'; })
        .ocCruzar(ORDENES, false);
    });
  };

  $('aplicar').onclick = function () {
    this.disabled = true;
    google.script.run.withSuccessHandler(function (r) {
      pintar(r);
      $('aplicar').style.display = 'none';
      $('msg').innerHTML = '<b>' + r.escritas + '</b> fila(s) con MONTO A PAGAR escrito.';
    }).withFailureHandler(function (e) { $('msg').innerHTML = '<span class="error">' + esc(e.message) + '</span>'; })
      .ocCruzar(ORDENES, true);
  };
</script>
`;
}
