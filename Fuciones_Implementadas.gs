function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Inventario')
    .addItem('Registrar Salida (Préstamo)', 'formularioSalida')
    .addItem('Registrar Entrada (Devolución)', 'formularioEntrada')
    .addItem('Consultar Material', 'formularioConsultaMaterial')
    .addToUi();
}
