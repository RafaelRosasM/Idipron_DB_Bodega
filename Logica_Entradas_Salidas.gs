// Botones para el front
function botonSalida() { formularioSalida(); }
function botonEntrada() { formularioEntrada(); }

// Formularios Salida/Entrada
function formularioSalida() {
  const html = HtmlService.createHtmlOutputFromFile('salida')
    .setWidth(600)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Salida');
}

function formularioEntrada() {
  const html = HtmlService.createHtmlOutputFromFile('entrada')
    .setWidth(600)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Entrada');
}

// ============================
// Procesar Salida
// ============================
function registrarSalida(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const materiales = ss.getSheetByName('MATERIALES');
  const historial = ss.getSheetByName('HISTORIAL_INVENTARIO');

  if (!Number.isInteger(datos.cantidad) || datos.cantidad <= 0) {
    return '❌ La cantidad debe ser un número entero positivo mayor que 0';
  }

  const data = materiales.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === datos.material.toString().trim()) {
      let stockActual = parseInt(data[i][2], 10) || 0;
      if (stockActual >= datos.cantidad) {
        materiales.getRange(i+1, 3).setValue(stockActual - datos.cantidad);
        historial.appendRow([new Date(), datos.responsable, datos.material, datos.cantidad, 'SALIDA']);
        return '✅ Salida registrada con éxito';
      } else {
        return '❌ Stock insuficiente';
      }
    }
  }
  return '❌ Material no encontrado';
}

// ============================
// Procesar Entrada
// ============================
function registrarEntrada(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const materiales = ss.getSheetByName('MATERIALES');
  const historial = ss.getSheetByName('HISTORIAL_INVENTARIO');

  if (!Number.isInteger(datos.cantidad) || datos.cantidad <= 0) {
    return '❌ La cantidad debe ser un número entero positivo mayor que 0';
  }

  const data = materiales.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === datos.material.toString().trim()) {
      let stockActual = parseInt(data[i][2], 10) || 0;
      materiales.getRange(i+1, 3).setValue(stockActual + datos.cantidad);
      historial.appendRow([new Date(), datos.responsable, datos.material, datos.cantidad, 'ENTRADA']);
      return '✅ Entrada registrada con éxito';
    }
  }
  return '❌ Material no encontrado';
}

// ============================
// Listas desplegables
// ============================
function obtenerResponsables() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RESPONSABLES');
  const data = hoja.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0]?.toString().trim() || "",
    nombre: row[1]?.toString().trim() || "",
    contacto: row[2]?.toString().trim() || "",
    upi: row[3]?.toString().trim() || "",
    contratoFin: row[4]?.toString().trim() || ""
  }));
}

function obtenerMateriales() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MATERIALES');
  const grupos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GRUPO_MATERIAL').getDataRange().getValues();
  const dictGrupos = {};
  for (let i = 1; i < grupos.length; i++) dictGrupos[grupos[i][0]] = grupos[i][1];

  const data = hoja.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0],
    nombre: row[1],
    stock: row[2],
    elemento: row[3],
    serial: row[4],
    valorHistorico: row[5],
    grupo: dictGrupos[row[6]] || ""
  }));
}
