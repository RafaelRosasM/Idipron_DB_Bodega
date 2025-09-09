function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Inventario')
    .addItem('Registrar Salida (Préstamo)', 'formularioSalida')
    .addItem('Registrar Entrada (Devolución)', 'formularioEntrada')
    .addItem('Consultar Material', 'formularioConsultaMaterial')
    .addToUi();
}

// Botones para el front
function botonSalida() { formularioSalida(); }
function botonEntrada() { formularioEntrada(); }
function botonConsultaMaterial() { formularioConsultaMaterial(); }

// Formularios
function formularioSalida() {
  const html = HtmlService.createHtmlOutputFromFile('salida')
    .setWidth(400)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Salida');
}

function formularioEntrada() {
  const html = HtmlService.createHtmlOutputFromFile('entrada')
    .setWidth(400)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Entrada');
}

function formularioConsultaMaterial() {
  var html = HtmlService.createHtmlOutputFromFile('consultaMaterial')
    .setWidth(700)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Consulta de Materiales');
}

// ============================
// Procesar Salida
// ============================
function registrarSalida(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const materiales = ss.getSheetByName('MATERIALES');
  const historial = ss.getSheetByName('HISTORIAL_INVENTARIO');

  // ✅ Validación de cantidad
  if (!Number.isInteger(datos.cantidad) || datos.cantidad <= 0) {
    return '❌ La cantidad debe ser un número entero positivo mayor que 0';
  }

  const data = materiales.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
  Logger.log("Comparando: Hoja=" + data[i][0] + " | Modal=" + datos.material);

  if (data[i][0].toString().trim() === datos.material.toString().trim()) {
    let stockActual = parseInt(data[i][2], 10) || 0;
    Logger.log("✅ Material encontrado, stock actual: " + stockActual);

    if (stockActual >= datos.cantidad) {
      materiales.getRange(i+1, 3).setValue(stockActual - datos.cantidad);
      historial.appendRow([
        new Date(),
        datos.responsable,
        datos.material,
        datos.cantidad,
        'SALIDA'
      ]);
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

  // ✅ Validación de cantidad
  if (!Number.isInteger(datos.cantidad) || datos.cantidad <= 0) {
    return '❌ La cantidad debe ser un número entero positivo mayor que 0';
  }

  const data = materiales.getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
  Logger.log("Comparando: Hoja=" + data[i][0] + " | Modal=" + datos.material);

  if (data[i][0].toString().trim() === datos.material.toString().trim()) { // ID
    let stockActual = parseInt(data[i][2], 10) || 0; // STOCK (si está vacío, arranca en 0)
    Logger.log("✅ Material encontrado, stock actual: " + stockActual);

    materiales.getRange(i+1, 3).setValue(stockActual + datos.cantidad);

    historial.appendRow([
      new Date(),
      datos.responsable,
      datos.material,
      datos.cantidad,
      'ENTRADA'
    ]);

    return '✅ Entrada registrada con éxito';
  }
}
  return '❌ Material no encontrado';
}

// ============================
// Listas desplegables para formularios
// ============================
function obtenerResponsables() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RESPONSABLES');
  const data = hoja.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0] != null ? row[0].toString().trim() : "",
    nombre: row[1] != null ? row[1].toString().trim() : "",
    contacto: row[2] != null ? row[2].toString().trim() : "",
    upi: row[3] != null ? row[3].toString().trim() : "",
    contratoFin: row[4] != null ? row[4].toString().trim() : ""
  }));
}

function obtenerMateriales() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MATERIALES');
  const grupos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GRUPO_MATERIAL').getDataRange().getValues();

  // crear diccionario grupoId → nombreGrupo
  const dictGrupos = {};
  for (let i = 1; i < grupos.length; i++) {
    dictGrupos[grupos[i][0]] = grupos[i][1];
  }

  const data = hoja.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0],                // ID
    nombre: row[1],            // DESCRIPCION
    stock: row[2],             // STOCK
    elemento: row[3],
    serial: row[4],
    valorHistorico: row[5],
    grupo: dictGrupos[row[6]] || ""  // Buscar el grupo por ID
  }));
}

// ============================
// Devuelve sugerencias para autocompletar
// ============================
function obtenerSugerencias(tipo, valor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("MATERIALES"); 
  const datos = hoja.getDataRange().getValues();
  let col = tipo === "id" ? 0 : 1; // Col 0 = ID, Col 1 = Descripción

  return datos
    .map(r => r[col])
    .filter(v => v && v.toString().toLowerCase().includes(valor.toLowerCase()))
    .slice(0, 25); // Máximo 25 sugerencias
}

// ============================
// Procesar Consulta
// ============================
function buscarMaterial(id, descripcion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const materialesSheet = ss.getSheetByName("MATERIALES");
  const historialSheet = ss.getSheetByName("HISTORIAL_INVENTARIO");
  const responsablesSheet = ss.getSheetByName("RESPONSABLES");
  const gruposSheet = ss.getSheetByName("GRUPO_MATERIAL");

  const materiales = materialesSheet.getDataRange().getValues();
  const historial = historialSheet.getDataRange().getValues();
  const responsables = responsablesSheet.getDataRange().getValues();
  const grupos = gruposSheet.getDataRange().getValues();

  // Diccionario grupos
  const dictGrupos = {};
  for (let i = 1; i < grupos.length; i++) {
    dictGrupos[grupos[i][0]] = grupos[i][1];
  }

  let resultados = [];

  for (let i = 1; i < materiales.length; i++) {
    let matID = materiales[i][0];       // ID
    let matDesc = materiales[i][1];     // Descripción
    let stock = materiales[i][2];       // Stock
    let elemento = materiales[i][3];
    let serial = materiales[i][4];
    let valorHistorico = materiales[i][5];
    let grupoID = materiales[i][6];
    let grupoNombre = dictGrupos[grupoID] || "";

    // Filtro por ID o descripción
    if ((id && matID.toString().toLowerCase().includes(id.toLowerCase())) ||
        (descripcion && matDesc.toLowerCase().includes(descripcion.toLowerCase()))) {

      // 📌 Agrupar historial por responsable
      let prestamosPorResponsable = {};

      for (let j = 1; j < historial.length; j++) {
        let responsable = historial[j][1];
        let materialID = historial[j][2];
        let cantidad = historial[j][3];
        let tipo = historial[j][4]; // "Salida" o "Entrada"

        if (materialID == matID) {
          if (!prestamosPorResponsable[responsable]) {
            prestamosPorResponsable[responsable] = 0;
          }
          if (tipo.toLowerCase() === "salida") {
            prestamosPorResponsable[responsable] += cantidad;
          } else if (tipo.toLowerCase() === "entrada") {
            prestamosPorResponsable[responsable] -= cantidad;
          }
        }
      }

      // 📌 Construir resultados
      let tienePrestamos = false;
      for (let resp in prestamosPorResponsable) {
        let prestado = prestamosPorResponsable[resp];
        if (prestado > 0) {
          tienePrestamos = true;

          // Buscar info extra en RESPONSABLES
          let upi = "", contacto = "", contratoFin = "";
          for (let r = 1; r < responsables.length; r++) {
            if (responsables[r][0] == resp) { 
              upi = responsables[r][3];
              contacto = responsables[r][2];
              contratoFin = responsables[r][4];
              break;
            }
          }

          resultados.push({
            id: matID,
            descripcion: matDesc,
            stock: stock,
            elemento: elemento,
            serial: serial,
            valorHistorico: valorHistorico,
            grupo: grupoNombre,
            responsable: resp,
            upi: upi,
            contacto: contacto,
            contratoFin: contratoFin,
            cantidadPrestada: prestado
          });
        }
      }

      // 📌 Si no hay préstamos → solo mostrar info general
      if (!tienePrestamos) {
        resultados.push({
          id: matID,
          descripcion: matDesc,
          stock: stock,
          elemento: elemento,
          serial: serial,
          valorHistorico: valorHistorico,
          grupo: grupoNombre,
          responsable: "",
          upi: "",
          contacto: "",
          contratoFin: "",
          cantidadPrestada: 0
        });
      }
    }
  }

  return resultados;
}
