/**
 * RECORDATORIO DIARIO DE COMPROBANTES
 *
 * Reemplaza al popup fijo "¿CARGASTE LOS COMPROBANTES DE AYER?". En vez de preguntar, mira los
 * datos y lista lo que falta. Sale una vez por día, desde el trigger instalable "Al abrir" que ya
 * existe (apunta a checkComprobantesPopup: el nombre no cambió a propósito).
 *
 * QUÉ CUENTA COMO "FALTA COMPROBANTE"
 *   1) En la última AGRUP anterior a hoy (un lunes, la del viernes): grupos con ESTADO = TRANSFERIDO
 *      y la columna COMPROBANTE vacía.
 *   2) En la hoja del día de esa misma fecha: filas con la marca "⚠️ PAGADO EL dd-MM - FALTA
 *      COMPROBANTE". Eso agarra transferencias de días anteriores que ya salieron de la agrupada
 *      y siguen sin comprobante (ej. BID-LIN-1-17, pagada el 29-09).
 *   En los dos casos, una tarea que ya NO está en base se descarta: si salió de base, el
 *   comprobante ya se procesó en AppSheet. Si base no se puede leer, no se descarta nada (se
 *   muestra de más, nunca de menos).
 *
 * La agrupada anterior se lee por nombre de encabezado (colsAgrup_), no por posición: puede ser
 * de un día en que la agrupada tenía otras columnas.
 *
 * Si no falta nada: un toast "todo OK", sin ventana.
 * Si no hay ninguna AGRUP anterior: el popup genérico de siempre.
 *
 * ⚠️ BORRAR la función checkComprobantesPopup vieja del archivo donde esté. Si quedan dos funciones
 * con el mismo nombre, Apps Script no avisa: usa la última que lee y no sabés cuál es.
 *
 * DEPENDE de Pagoslb.gs (formatearNombreHoja_, parsearNombreHoja_, parsearNombreHojaAgrup_,
 * hojasAnterioresA_, indicesPorHeader_, mapaStatusPorTask_, formatearImporte_, aNumero_,
 * MARCA_PAGADO_PREFIJO, MARCA_PAGADO_SUFIJO, EST_TRANSFERIDO) y de Hojaagrupada.gs
 * (AGRUP_PREFIJO, colsAgrup_, valAgrup_).
 */

function checkComprobantesPopup() {
  const props = PropertiesService.getUserProperties();
  const hoyTxt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (props.getProperty('comprobantesPopupFecha') === hoyTxt) return; // ya se mostró hoy
  props.setProperty('comprobantesPopupFecha', hoyTxt);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoy = parsearNombreHoja_(formatearNombreHoja_(new Date()));
  const agrups = hojasAnterioresA_(ss, hoy, h => parsearNombreHojaAgrup_(h.getName()));
  if (agrups.length === 0) {
    popupComprobantesGenerico_();
    return;
  }

  const ag = agrups[0].hoja;
  const fechaAg = ag.getName().replace(AGRUP_PREFIJO, '');
  const enBase = mapaStatusPorTask_();
  const baseLeida = Object.keys(enBase).length > 0;
  const sigueEnBase = t => !baseLeida || Object.prototype.hasOwnProperty.call(enBase, t);

  const faltan = [];
  const yaListadas = {};

  // 1) Agrupada anterior: TRANSFERIDO sin comprobante.
  const n = ag.getLastRow() - 1;
  const C = colsAgrup_(ag);
  if (n > 0 && C.ESTADO > 0 && C.IDS > 0) {
    const valores = ag.getRange(2, 1, n, C.ANCHO).getValues();
    const enlaces = C.COMPROBANTE > 0 ? ag.getRange(2, C.COMPROBANTE, n, 1).getRichTextValues() : null;
    valores.forEach((v, i) => {
      if (String(valAgrup_(v, C, 'ESTADO')).trim().toUpperCase() !== EST_TRANSFERIDO) return;
      const rt = enlaces ? enlaces[i][0] : null;
      const tieneComprobante = !!((rt && rt.getLinkUrl()) ||
                                  String(valAgrup_(v, C, 'COMPROBANTE') || '').trim());
      if (tieneComprobante) return;
      const ids = String(valAgrup_(v, C, 'IDS') || '').split(',').map(s => s.trim())
        .filter(String).filter(sigueEnBase);
      if (ids.length === 0) return;
      ids.forEach(t => { yaListadas[t] = true; });
      const monto = aNumero_(valAgrup_(v, C, 'MONTO'));
      faltan.push({
        cuando: 'transferido el ' + fechaAg.substring(0, 5),
        benef: String(valAgrup_(v, C, 'BENEF') || '(sin beneficiario)'),
        ids: ids,
        monto: (monto === null ? '' : formatearImporte_(monto)) + ' ' + String(valAgrup_(v, C, 'MONEDA') || '')
      });
    });
  }

  // 2) Hoja del día de esa fecha: marcas "PAGADO EL ... FALTA COMPROBANTE" que vienen de antes.
  const hojaDia = ss.getSheetByName(fechaAg);
  if (hojaDia && hojaDia.getLastRow() > 1) {
    const idx = indicesPorHeader_(hojaDia);
    if (idx['TASK'] >= 0 && idx['ESTADO'] >= 0) {
      hojaDia.getRange(2, 1, hojaDia.getLastRow() - 1, hojaDia.getLastColumn()).getValues().forEach(v => {
        const estado = String(v[idx['ESTADO']] || '');
        if (estado.indexOf(MARCA_PAGADO_PREFIJO) !== 0 || estado.indexOf(MARCA_PAGADO_SUFIJO) === -1) return;
        const t = String(v[idx['TASK']] || '').trim();
        if (!t || yaListadas[t] || !sigueEnBase(t)) return;
        yaListadas[t] = true;
        const monto = idx['MONTO'] >= 0 ? aNumero_(v[idx['MONTO']]) : null;
        faltan.push({
          cuando: estado.replace(MARCA_PAGADO_PREFIJO, 'pagado el ').replace(MARCA_PAGADO_SUFIJO, ''),
          benef: String((idx['BENEFICIARIO'] >= 0 && v[idx['BENEFICIARIO']]) || '(sin beneficiario)'),
          ids: [t],
          monto: (monto === null ? '' : formatearImporte_(monto)) + ' ' +
                 String(idx['MONEDA'] >= 0 ? v[idx['MONEDA']] || '' : '')
        });
      });
    }
  }

  if (faltan.length === 0) {
    ss.toast('Comprobantes al día (revisado contra AGRUP ' + fechaAg + ').', 'Recordatorio', 6);
    return;
  }

  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const filas = faltan.map(f =>
    '<tr><td>' + esc(f.ids.join(', ')) + '</td><td>' + esc(f.benef) + '</td>' +
    '<td style="text-align:right;white-space:nowrap">' + esc(f.monto) + '</td>' +
    '<td style="white-space:nowrap">' + esc(f.cuando) + '</td></tr>'
  ).join('');

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:8px 4px;">' +
    '<div style="font-size:22px;font-weight:bold;color:#c0392b;margin-bottom:6px;">⚠️ ' +
    faltan.length + ' pago(s) sin comprobante</div>' +
    '<div style="font-size:13px;color:#555;margin-bottom:12px;">Ya salieron, pero el comprobante ' +
    'no está cargado. Revisado contra AGRUP ' + esc(fechaAg) + ' y la hoja del día de esa fecha.</div>' +
    '<table style="border-collapse:collapse;font-size:13px;width:100%;">' +
    '<tr style="background:#434343;color:#fff;"><th style="text-align:left;padding:4px">Tarea(s)</th>' +
    '<th style="text-align:left;padding:4px">Beneficiario</th><th style="padding:4px">Monto</th>' +
    '<th style="text-align:left;padding:4px">Cuándo</th></tr>' + filas + '</table>' +
    '<style>td{padding:4px;border-bottom:1px solid #ddd;vertical-align:top}</style></div>'
  ).setWidth(640).setHeight(Math.min(160 + faltan.length * 34, 560));

  SpreadsheetApp.getUi().showModalDialog(html, 'Recordatorio diario');
}

/** El popup de siempre, para cuando no hay una agrupada anterior contra la que revisar. */
function popupComprobantesGenerico_() {
  const html = HtmlService.createHtmlOutput(
    '<div style="text-align:center; font-family:Arial, sans-serif; padding:30px;">' +
    '<div style="font-size:60px;">⚠️</div>' +
    '<div style="font-size:30px; font-weight:bold; color:#c0392b; margin:20px 0;">' +
    '¿CARGASTE LOS COMPROBANTES<br>DE AYER?' +
    '</div>' +
    '<div style="font-size:16px; color:#555;">No encontré una agrupada anterior para revisarlo solo.</div>' +
    '</div>'
  ).setWidth(450).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Recordatorio diario');
}

/**
 * Para probar sin esperar a mañana: correla desde el editor y después recargá la planilla (F5).
 * Solo borra la marca de "ya se mostró hoy"; el popup lo dispara el trigger al abrir.
 * (Desde el editor no se puede mostrar una ventana: getUi() solo anda con la planilla abierta.)
 */
function resetRecordatorioComprobantes() {
  PropertiesService.getUserProperties().deleteProperty('comprobantesPopupFecha');
}
