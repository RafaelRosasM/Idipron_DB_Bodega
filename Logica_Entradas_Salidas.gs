// ============================
// Botones para el front
// ============================
function botonSalida() { formularioSalida(); }
function botonEntrada() { formularioEntrada(); }

// ============================
// Formularios Salida/Entrada
// ============================

function formularioSalida() {
  validarPermiso("salida");
  const html = HtmlService.createHtmlOutputFromFile('salida')
    .setWidth(600)
    .setHeight(600); // Ajusté un poco para el nuevo campo
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Salida');
}

function formularioEntrada() {
  validarPermiso("entrada");
  const html = HtmlService.createHtmlOutputFromFile('entrada')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Entrada');
}

// ============================
// Procesar Salida
// ============================
function registrarSalida(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const materiales = ss.getSheetByName('EQUIPOS');
  const historial = ss.getSheetByName('HISTORIAL_INVENTARIO');

  if (!Number.isInteger(datos.cantidad) || datos.cantidad <= 0) {
    return '❌ La cantidad debe ser un número entero positivo mayor que 0';
  }

  if (!datos.fecha) {
    return '❌ La fecha es obligatoria';
  }

  // Parsear la fecha manualmente (YYYY-MM-DD → DD/MM/YYYY sin desfase horario)
  let partes = datos.fecha.split("-"); // [YYYY, MM, DD]
  let fechaMovimiento = `${partes[2]}/${partes[1]}/${partes[0]}`;

  // Fecha de registro con hora real de Colombia
  let fechaRegistro = Utilities.formatDate(
    new Date(),
    "GMT-5",
    "dd/MM/yyyy HH:mm:ss"
  );


  const data = materiales.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === datos.material.toString().trim()) {
      let stockActual = parseInt(data[i][2], 10) || 0;
      if (stockActual >= datos.cantidad) {
        materiales.getRange(i+1, 3).setValue(stockActual - datos.cantidad);

        // Agregamos datos.upi como último campo (si no viene, guardamos "")
        historial.appendRow([
          new Date(),                 // FECHA_REGISTRO
          datos.responsable,          // ID_RESPONSABLE
          datos.material,             // ID_MATERIAL
          datos.cantidad,             // CANTIDAD
          'SALIDA',                   // TIPO
          fechaMovimiento,            // FECHA_MOVIMIENTO
          datos.upi || ""             // UPI (nuevo campo al final)
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
  const materiales = ss.getSheetByName('EQUIPOS');
  const historial = ss.getSheetByName('HISTORIAL_INVENTARIO');

  if (!Number.isInteger(datos.cantidad) || datos.cantidad <= 0) {
    return '❌ La cantidad debe ser un número entero positivo mayor que 0';
  }

  if (!datos.fecha) {
    return '❌ La fecha es obligatoria';
  }

  // Parsear la fecha manualmente (YYYY-MM-DD → DD/MM/YYYY sin desfase horario)
  let partes = datos.fecha.split("-"); // [YYYY, MM, DD]
  let fechaMovimiento = `${partes[2]}/${partes[1]}/${partes[0]}`;

  // Fecha de registro con hora real de Colombia
  let fechaRegistro = Utilities.formatDate(
    new Date(),
    "GMT-5",
    "dd/MM/yyyy HH:mm:ss"
  );


  const data = materiales.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === datos.material.toString().trim()) {
      let stockActual = parseInt(data[i][2], 10) || 0;
      materiales.getRange(i+1, 3).setValue(stockActual + datos.cantidad);

      // Agregamos datos.upi como último campo (si no viene, guardamos "")
      historial.appendRow([
        new Date(),                 // FECHA_REGISTRO
        datos.responsable,          // ID_RESPONSABLE
        datos.material,             // ID_MATERIAL
        datos.cantidad,             // CANTIDAD
        'ENTRADA',                  // TIPO
        fechaMovimiento,            // FECHA_MOVIMIENTO
        datos.upi || ""             // UPI (nuevo campo al final)
      ]);

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

// ============================
// Buscar Equipos (Búsqueda dinámica por ID o Placa)
// ============================
function buscarEquiposPorID(termino) {
  if (!termino || termino.trim().length < 3) return []; // Solo buscar si hay 3+ caracteres

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EQUIPOS');
  const data = hoja.getDataRange().getValues();
  const resultados = [];

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0]?.toString().toLowerCase() || "";
    const nombre = data[i][1]?.toString() || "";
    if (id.includes(termino.toLowerCase())) {
      resultados.push({
        id: data[i][0],
        nombre: nombre
      });
    }
  }
  return resultados;
}

// ============================
// Obtener detalle del equipo seleccionado
// ============================
function obtenerEquipoPorID(idBuscado) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EQUIPOS');
  const data = hoja.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString().trim() === idBuscado.trim()) {
      return {
        id: data[i][0],
        nombre: data[i][1],
        stock: data[i][2],
        elemento: data[i][3],
        serial: data[i][4],
        valorHistorico: data[i][5]
      };
    }
  }
  return null;
}

// ============================
// Obtener lista de UPIs (para poblar los selects en los modales)
// Hoja esperada: "UPIS" con columnas: ID | UPI_NOMBRE
// Devuelve array de nombres (UPI_NOMBRE)
// ============================
function obtenerUPI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('UPIS');
  if (!hoja) return [];
  const data = hoja.getDataRange().getValues();
  // Asumimos que la 1ª fila es encabezado y la segunda columna (index 1) es UPI_NOMBRE
  return data.slice(1).map(row => (row[1] ? row[1].toString().trim() : ""));
}
