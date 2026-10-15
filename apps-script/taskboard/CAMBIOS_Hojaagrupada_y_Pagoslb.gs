/* ═══════════════════════════════════════════════════════════════════════════════════════════
 * HOJAAGRUPADA.GS — reemplazar AGRUP_HEADERS y AGRUP_COL (están al principio del archivo)
 * MONTO A PAGAR pasa a ser columna propia de la agrupada, entre COMENTARIO ANTERIOR y LINKS.
 * Nada más cambia en este archivo: todo lo demás usa AGRUP_COL o busca por nombre de header.
 * ═══════════════════════════════════════════════════════════════════════════════════════════ */

const AGRUP_HEADERS = [
  'TIPO', 'BENEFICIARIO', 'SUBGRUPO', 'N° TX', 'TRANSACCIONES (IDs)', 'MONTO', 'MONEDA',
  'CHECK', 'ESTADO', 'ORIGEN', 'COMENTARIO', 'COMENTARIO ANTERIOR', 'MONTO A PAGAR', 'LINKS', 'PDF',
  'COMPROBANTE'
];

/** Número de columna (1-based) de cada campo de la hoja agrupada. */
const AGRUP_COL = {
  TIPO: 1, BENEF: 2, SUBGRUPO: 3, NTX: 4, IDS: 5, MONTO: 6, MONEDA: 7,
  CHECK: 8, ESTADO: 9, ORIGEN: 10, COMENT: 11, COMENT_ANT: 12, PAGAR: 13, LINKS: 14, PDF: 15,
  COMPROBANTE: 16
};


/* ═══════════════════════════════════════════════════════════════════════════════════════════
 * PAGOSLB.GS — reemplazar la función comentariosAgrupAnteriores_ entera
 * Único cambio: saca la marca "PAYPAL: ..." del comentario antes de arrastrarlo
 * (quitarMarcaPaypal_, en PaypalAprobaciones.gs).
 * ═══════════════════════════════════════════════════════════════════════════════════════════ */

/**
 * Recorre las hojas "AGRUP DD-MM-YYYY" anteriores a hoy, de la más nueva a la más vieja, y arma
 * dos índices con el comentario más reciente que encuentre:
 *   porTask[TASK]              → para la hoja del día. La fila agrupada cubre varias tareas, así
 *                                que su comentario se reparte a todas las que figuran en IDs.
 *   porGrupo[tipo|nombre|mon]  → para la hoja agrupada de hoy, que se compara grupo contra grupo.
 *
 * La clave del día es TASK sola, no TASK+MONTO: la hoja agrupada guarda el monto SUMADO del grupo,
 * no el de cada tarea, así que no hay con qué armar la clave compuesta.
 *
 * La marca "PAYPAL: ..." que escribe PaypalAprobaciones.gs al final del comentario NO se arrastra:
 * cambia con cada estado de PayPal y COMENTARIO ANTERIOR es parte de la clave de grupo.
 *
 * Devuelve { porTask:{}, porGrupo:{} }. Si no hay ninguna hoja agrupada anterior, ambos vacíos.
 */
function comentariosAgrupAnteriores_(ss, nombreHojaHoy) {
  const hoy = parsearNombreHoja_(nombreHojaHoy);
  const res = { porTask: {}, porGrupo: {} };
  if (!hoy) return res;

  const anteriores = hojasAnterioresA_(ss, hoy, h => parsearNombreHojaAgrup_(h.getName()));

  anteriores.forEach(o => {
    const ultimaFila = o.hoja.getLastRow();
    if (ultimaFila < 2) return;

    const ancho = Math.max(o.hoja.getLastColumn(), 1);
    const hdr = o.hoja.getRange(1, 1, 1, ancho).getValues()[0]
      .map(h => String(h == null ? '' : h).trim().toUpperCase());

    const iTipo = hdr.indexOf('TIPO');
    const iBenef = hdr.indexOf('BENEFICIARIO');
    const iIds = hdr.indexOf('TRANSACCIONES (IDS)');
    const iMoneda = hdr.indexOf('MONEDA');
    const iComent = hdr.indexOf('COMENTARIO');
    if (iComent < 0 || iIds < 0) return; // hoja agrupada vieja, sin columna COMENTARIO

    const etiqueta = Utilities.formatDate(o.fecha, ss.getSpreadsheetTimeZone(), 'dd-MM');
    const valores = o.hoja.getRange(2, 1, ultimaFila - 1, ancho).getValues();

    valores.forEach(v => {
      const texto = quitarMarcaPaypal_(v[iComent]);
      if (texto === '') return;
      const registro = { fecha: etiqueta, texto: texto };

      String(v[iIds] || '').split(',').forEach(id => {
        const task = id.trim();
        if (task === '') return;
        if (!res.porTask[task]) res.porTask[task] = registro;
      });

      if (iTipo >= 0) {
        const clave = claveGrupoAgrup_(v[iTipo], iBenef >= 0 ? v[iBenef] : '', iMoneda >= 0 ? v[iMoneda] : '');
        if (!res.porGrupo[clave]) res.porGrupo[clave] = registro;
      }
    });
  });

  return res;
}
