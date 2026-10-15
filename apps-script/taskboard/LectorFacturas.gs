/**
 * LECTOR DE FACTURAS EN EL NAVEGADOR
 *
 * Reemplaza al OCR de Drive en "Validar beneficiario y monto (filas seleccionadas)".
 *
 * POR QUÉ
 * Las facturas de ARCA vienen como PDF con texto real adentro. Pasarlas por el OCR de Drive las
 * convierte en un Google Doc que reordena el texto (en la de Carbone queda "$ 120.000,00 Total",
 * con el importe ANTES del rótulo) y además gasta una cuota por minuto. Leído directo del PDF y
 * reordenado por posición en la página, el texto sale como se ve impreso:
 *     Total  $ 120.000,00
 *     SUBTOTAL  IVA  5,00  %  TOTAL $
 *     104.761,90  5.238,10  110.000,00
 * y los mismos rótulos que ya usaba la validación lo leen bien.
 *
 * QR DE ARCA
 * Si la factura trae el QR obligatorio de ARCA, se decodifica y su importe manda sobre el texto:
 * es el total que el emisor le informó a ARCA. Se le pasa a la validación como una línea
 * "[QR ARCA] importe: ..." (ver parsearImporteFactura_ en Pagoslb.gs). No todas lo traen: la A de
 * Jornada Web tiene el código de barras viejo, que no lleva importe. Esas se leen por el texto.
 * Las monedas distintas de pesos se ignoran en el QR: la validación compara en pesos.
 *
 * CÓMO CORRE
 *   1. abrirLectorFacturas() abre la ventana con las filas seleccionadas de la hoja del día.
 *   2. La ventana pide cada factura al servidor (lfArchivo), la lee con pdf.js y busca el QR con
 *      ZXing. Todo en tu navegador: no usa la cuota de OCR de Drive.
 *   3. Manda los textos a lfValidar, que corre validarFilas_ (la misma validación de siempre:
 *      beneficiario, monto, pagos agrupados, MONTO FACTURA) con esos textos.
 * Si una factura es imagen (JPG/PNG), o el PDF es escaneado y no tiene texto, se cae al OCR de
 * Drive como antes, sumándole el QR si se pudo leer.
 *
 * DEPENDE de Pagoslb.gs (validarFilas_, contextoAgrupacion_, anclasDeGrupo_, claveAgrupacion_,
 * mapaStatusPorTask_, extraerIdDeDrive_, obtenerTextoFactura_, HEADERS, COL, TIPOS_VALIDABLES,
 * STATUS_AGRUPACION).
 */

function abrirLectorFacturas() {
  const hoja = SpreadsheetApp.getActiveSheet();
  if (!/^\d{2}-\d{2}-\d{4}$/.test(hoja.getName())) {
    SpreadsheetApp.getUi().alert('Parate sobre la hoja del día (DD-MM-YYYY) y seleccioná las filas a validar.\n' +
                                 'Estás en "' + hoja.getName() + '".');
    return;
  }
  const r = hoja.getActiveRange();
  const datos = lfFacturasDeSeleccion_(hoja, r.getRow(), r.getNumRows());
  if (datos.ids.length === 0) {
    SpreadsheetApp.getUi().alert('En las filas seleccionadas no hay ninguna A o C con factura.');
    return;
  }
  const html = HtmlService.createTemplate(lfHtml_());
  html.datos = JSON.stringify({ hoja: hoja.getName(), fila: r.getRow(), cant: r.getNumRows(), ids: datos.ids });
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(520).setHeight(300), 'Validar facturas');
}

/**
 * IDs de Drive de las facturas que la validación va a necesitar: las de las filas seleccionadas y,
 * para las filas agrupadas sin factura propia, la de la fila que la tiene (aunque esté fuera de la
 * selección). Sin eso, la validación de esas filas iría al OCR de Drive.
 */
function lfFacturasDeSeleccion_(hoja, filaInicio, cantFilas) {
  const statusPorTask = mapaStatusPorTask_();
  const ctx = contextoAgrupacion_(hoja, statusPorTask);
  const valores = hoja.getRange(filaInicio, 1, cantFilas, HEADERS.length).getValues();
  const ids = [];
  const agregar = link => {
    const id = extraerIdDeDrive_(link);
    if (id && ids.indexOf(id) === -1) ids.push(id);
  };
  valores.forEach((v, i) => {
    if (filaInicio + i === 1) return;
    const task = v[COL.TASK - 1];
    if (task === '' || TIPOS_VALIDABLES.indexOf(v[COL.TIPO - 1]) === -1) return;
    const link = String(v[COL.INVOICE - 1] || '').trim();
    if (link !== '') { agregar(link); return; }
    if (statusPorTask[String(task)] !== STATUS_AGRUPACION) return;
    anclasDeGrupo_(task, claveAgrupacion_(v[COL.DATO - 1], v[COL.BENEF - 1]), ctx)
      .forEach(a => agregar(a.link));
  });
  return { ids: ids };
}

/** Una factura en base64, para que la lea el navegador. */
function lfArchivo(id) {
  try {
    const archivo = DriveApp.getFileById(id);
    const mime = archivo.getMimeType();
    // Un Google Doc no es un PDF: se lo deja al OCR/lectura de Drive del lado del servidor.
    if (mime === MimeType.GOOGLE_DOCS) return { ok: true, id: id, mime: mime, data: '' };
    return { ok: true, id: id, mime: mime, data: Utilities.base64Encode(archivo.getBlob().getBytes()) };
  } catch (e) {
    return { ok: false, id: id, motivo: e.message };
  }
}

/**
 * Corre la validación con lo que leyó el navegador.
 * lecturas: id → { texto, qr } (texto puede venir vacío; qr es el importe en pesos o null).
 */
function lfValidar(nombreHoja, filaInicio, cantFilas, lecturas) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) return 'No encuentro la hoja "' + nombreHoja + '".';

  const cache = {};
  const avisos = [];
  Object.keys(lecturas || {}).forEach(id => {
    const l = lecturas[id];
    let texto = String(l.texto || '');
    if (texto.replace(/\s/g, '').length < 20) {
      // Imagen o PDF escaneado: sin texto propio, se lee con el OCR de Drive como antes.
      try { texto = obtenerTextoFactura_(id); }
      catch (e) { avisos.push(id + ': sin texto y el OCR de Drive falló (' + e.message + ')'); texto = ''; }
    }
    if (typeof l.qr === 'number') texto = '[QR ARCA] importe: ' + l.qr.toFixed(2) + '\n' + texto;
    if (texto !== '') cache[id] = { texto: texto };
  });

  return validarFilas_(hoja, filaInicio, cantFilas, cache) +
         (avisos.length ? '\n\n⚠️ ' + avisos.join('\n⚠️ ') : '');
}

/**
 * pdf.js y su worker van como <script> desde cdnjs (igual que el panel de lote de Comprobantes.gs:
 * con el worker cargado en la página no se crea un Web Worker de otro dominio). ZXing desde
 * jsdelivr, versión fija.
 * El texto se arma por renglones: los trozos con la misma altura en la página van en la misma
 * línea, de izquierda a derecha. Así queda como se ve impreso, que es lo que esperan los rótulos.
 * El QR se busca renderizando la primera página a 3x y, si no aparece, a 4,5x: en la factura de
 * Carbone el QR viene como una imagen de 69 px y a 2x no se llega a leer.
 */
function lfHtml_() {
  return `
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@zxing/library@0.20.0/umd/index.min.js"></script>
<style>
  body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; margin: 0; padding: 12px; color: #222; }
  #estado { margin-bottom: 8px; }
  #barra { height: 8px; background: #eee; border-radius: 4px; overflow: hidden; margin-bottom: 10px; }
  #barra div { height: 100%; width: 0; background: #0a0a8c; transition: width .2s; }
  #resumen { white-space: pre-wrap; font-size: 12px; background: #fafafa; border: 1px solid #ddd; padding: 8px; max-height: 190px; overflow: auto; display: none; }
  .err { color: #cc0000; }
</style>
<div id="estado">Preparando…</div>
<div id="barra"><div></div></div>
<div id="resumen"></div>
<script>
  var D = <?!= datos ?>;
  var lecturas = {}, hechas = 0, qrs = 0;
  function $(id) { return document.getElementById(id); }
  function estado(t) { $('estado').textContent = t; }

  function bytesDe(b64) {
    var bin = atob(b64), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }

  function textoPorLineas(items) {
    var filas = [];
    items.forEach(function (it) {
      if (!it.str || !it.str.trim()) return;
      var y = it.transform[5], x = it.transform[4], h = Math.abs(it.transform[3]) || 8;
      var f = null;
      for (var i = 0; i < filas.length; i++) if (Math.abs(filas[i].y - y) <= Math.max(2, h * 0.4)) { f = filas[i]; break; }
      if (!f) { f = { y: y, items: [] }; filas.push(f); }
      f.items.push({ x: x, s: it.str.trim() });
    });
    filas.sort(function (a, b) { return b.y - a.y; });
    return filas.map(function (f) {
      return f.items.sort(function (a, b) { return a.x - b.x; }).map(function (i) { return i.s; }).join('  ');
    }).join('\\n');
  }

  /** Importe en pesos del QR de ARCA, o null. */
  function importeQr(texto) {
    var m = String(texto || '').match(/[?&]p=([^&#]+)/);
    if (!m) return null;
    try {
      var j = JSON.parse(atob(decodeURIComponent(m[1]).replace(/ /g, '+').replace(/-/g, '+').replace(/_/g, '/')));
      if (j.moneda && j.moneda !== 'PES') return null;
      return typeof j.importe === 'number' ? j.importe : Number(j.importe);
    } catch (e) { return null; }
  }

  function leerQrDeCanvas(canvas) {
    if (!window.ZXing) return null;
    var ctx = canvas.getContext('2d');
    var d = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    var g = new Uint8ClampedArray(canvas.width * canvas.height);
    for (var i = 0; i < g.length; i++) g[i] = (d[i * 4] * 299 + d[i * 4 + 1] * 587 + d[i * 4 + 2] * 114) / 1000;
    try {
      var lum = new ZXing.RGBLuminanceSource(g, canvas.width, canvas.height);
      var hints = new Map(); hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
      return new ZXing.QRCodeReader().decode(new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(lum)), hints).getText();
    } catch (e) { return null; }
  }

  function qrDePagina(page) {
    var escalas = [3, 4.5];
    function probar(k) {
      if (k >= escalas.length) return Promise.resolve(null);
      var vp = page.getViewport({ scale: escalas[k] });
      var c = document.createElement('canvas');
      c.width = vp.width; c.height = vp.height;
      var ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      return page.render({ canvasContext: ctx, viewport: vp }).promise.then(function () {
        var t = leerQrDeCanvas(c);
        return t ? t : probar(k + 1);
      });
    }
    return probar(0);
  }

  function qrDeImagen(b64, mime) {
    return new Promise(function (ok) {
      var img = new Image();
      img.onload = function () {
        var c = document.createElement('canvas');
        c.width = img.width * 2; c.height = img.height * 2;
        var ctx = c.getContext('2d'); ctx.imageSmoothingEnabled = false;
        ctx.drawImage(img, 0, 0, c.width, c.height);
        ok(leerQrDeCanvas(c));
      };
      img.onerror = function () { ok(null); };
      img.src = 'data:' + mime + ';base64,' + b64;
    });
  }

  function leerPdf(b64) {
    return pdfjsLib.getDocument({ data: bytesDe(b64), isEvalSupported: false }).promise.then(function (doc) {
      var paginas = [];
      for (var p = 1; p <= doc.numPages; p++) paginas.push(doc.getPage(p));
      return Promise.all(paginas).then(function (pgs) {
        return Promise.all(pgs.map(function (pg) { return pg.getTextContent(); })).then(function (cs) {
          var texto = cs.map(function (c) { return textoPorLineas(c.items); }).join('\\n');
          return qrDePagina(pgs[0]).then(function (qr) { return { texto: texto, qr: importeQr(qr) }; });
        });
      });
    });
  }

  function leerUna(n) {
    $('barra').firstChild.style.width = Math.round(n / D.ids.length * 100) + '%';
    if (n >= D.ids.length) return validar();
    estado('Leyendo factura ' + (n + 1) + ' de ' + D.ids.length + '…');
    google.script.run.withSuccessHandler(function (r) {
      var sigue = function (l) {
        if (l) { lecturas[r.id] = l; if (typeof l.qr === 'number') qrs++; }
        leerUna(n + 1);
      };
      if (!r.ok) { sigue(null); return; }
      if (!r.data) { sigue({ texto: '', qr: null }); return; }
      if (r.mime === 'application/pdf' && window.pdfjsLib) {
        leerPdf(r.data).then(sigue, function () { sigue({ texto: '', qr: null }); });
      } else if (/^image\\//.test(r.mime)) {
        qrDeImagen(r.data, r.mime).then(function (q) { sigue({ texto: '', qr: importeQr(q) }); });
      } else {
        sigue({ texto: '', qr: null });
      }
    }).withFailureHandler(function () { leerUna(n + 1); }).lfArchivo(D.ids[n]);
  }

  function validar() {
    estado('Validando ' + Object.keys(lecturas).length + ' factura(s) (' + qrs + ' con QR de ARCA)…');
    google.script.run.withSuccessHandler(function (txt) {
      estado('Listo.');
      $('resumen').style.display = 'block';
      $('resumen').textContent = txt;
      google.script.host.setHeight(Math.min(560, 140 + txt.split('\\n').length * 16));
    }).withFailureHandler(function (e) {
      estado('Error del servidor.');
      $('resumen').style.display = 'block';
      $('resumen').innerHTML = '<span class="err">' + e.message + '</span>';
    }).lfValidar(D.hoja, D.fila, D.cant, lecturas);
  }

  if (!window.pdfjsLib) estado('No cargó pdf.js: las facturas se van a leer con el OCR de Drive.');
  leerUna(0);
</script>
`;
}
