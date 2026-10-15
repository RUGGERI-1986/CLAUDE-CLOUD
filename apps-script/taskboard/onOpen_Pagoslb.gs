function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Pagos LB')
    .addItem('Generar hoja del día', 'menuGenerarHojaDia')
    .addItem('Actualizar hoja del día', 'menuActualizarHojaDia')
    .addItem('Validar beneficiario y monto (filas seleccionadas)', 'validarSeleccion')
    .addItem('[DIAGNÓSTICO] Ver texto OCR de fila', 'diagnosticoOcrFila')
    .addToUi();

  ui.createMenu('Agrupado')
    .addItem('Generar hoja agrupada del día', 'menuGenerarAgrupada')
    .addItem('Actualizar hoja agrupada del día', 'menuActualizarAgrupada')
    .addSeparator()
    .addItem('Galicia: generar archivo de carga al banco', 'generarArchivoGalicia')
    .addItem('Generar hoja de APROBACIÓN', 'generarHojaAprobacion')
    .addItem('PayPal: revisar aprobaciones y confirmar abonados', 'revisarAprobacionesPaypal')
    .addSeparator()
    .addItem('Armar comprobante de pago (panel)', 'abrirPanelComprobante')
    .addItem('Armar comprobantes EN LOTE', 'abrirPanelComprobantesLote')
    .addToUi();
}
