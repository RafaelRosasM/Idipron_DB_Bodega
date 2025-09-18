// Botón de formulario
function botonConsultaMaterial() { formularioConsultaMaterial(); }

function formularioConsultaMaterial() {
  const html = HtmlService.createHtmlOutputFromFile('consultaMaterial')
    .setWidth(850)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Consulta de Materiales');
}

// Sugerencias de autocompletado
function obtenerSugerencias(tipo, valor) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MATERIALES");
  const datos = hoja.getDataRange().getValues();
  return datos.slice(1)
    .filter(r => r[0] && r[1] &&
      (tipo === "id"
        ? r[0].toString().toLowerCase().includes(valor.toLowerCase())
        : r[1].toString().toLowerCase().includes(valor.toLowerCase())))
    .map(r => ({ id: r[0], desc: r[1] }))
    .slice(0, 25);
}

//
// Helper: convierte varios formatos a Date de forma segura.
// Si no puede parsear, devuelve new Date(0) (epoch) para que no interrumpa la búsqueda.
function safeDate(val) {
  // Si ya es Date válido
  if (val instanceof Date && !isNaN(val)) return val;

  if (val == null || val === "") return new Date(0);

  // Si viene como número (raro con getValues, pero por si acaso)
  if (typeof val === "number") {
    var dnum = new Date(val);
    if (!isNaN(dnum)) return dnum;
  }

  // Intento directo
  var d = new Date(val);
  if (!isNaN(d)) return d;

  // Intento dd/mm/yyyy[ hh:mm:ss]
  var s = String(val).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}:\d{2}(?::\d{2})?))?/);
  if (m) {
    var day = parseInt(m[1], 10);
    var month = parseInt(m[2], 10);
    var year = parseInt(m[3], 10);
    if (year < 100) year += 2000;
    var time = m[4] || "00:00:00";
    // construir ISO
    var iso = year + "-" + (month < 10 ? "0"+month : month) + "-" + (day < 10 ? "0"+day : day) + "T" + time;
    var d2 = new Date(iso);
    if (!isNaN(d2)) return d2;
  }

  // Fallback: epoch (0) -> será considerado la fecha más antigua
  return new Date(0);
}

//

// Consulta de materiales
function buscarMaterial(id, descripcion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const matSheet = ss.getSheetByName("MATERIALES");
    const histSheet = ss.getSheetByName("HISTORIAL_INVENTARIO");
    const respSheet = ss.getSheetByName("RESPONSABLES");

    const materiales = matSheet.getDataRange().getValues();
    const historial = histSheet.getDataRange().getValues();
    const responsables = respSheet.getDataRange().getValues();

    Logger.log('buscarMaterial START | id=%s | desc=%s | materiales=%s | historial=%s',
               id, descripcion, materiales.length, historial.length);

    const resultados = [];
    const tz = Session.getScriptTimeZone();

    for (let i = 1; i < materiales.length; i++) {
      const matID   = materiales[i][0] != null ? String(materiales[i][0]).trim() : "";
      const matDesc = materiales[i][1] != null ? String(materiales[i][1]).trim() : "";
      const stock   = materiales[i][2] != null ? materiales[i][2] : "";

      // Filtrar por ID o descripción (ignora si campo de búsqueda vacío)
      if (
        (id && matID.toLowerCase().indexOf(String(id).toLowerCase()) !== -1) ||
        (descripcion && matDesc.toLowerCase().indexOf(String(descripcion).toLowerCase()) !== -1)
      ) {
        Logger.log('MATCH material -> %s : %s', matID, matDesc);

        // Buscar el último movimiento del historial para este material (sin usar sort)
        let latest = null; // { date: Date, tipo: string, responsableID: string, raw: row }
        for (let j = 1; j < historial.length; j++) {
          const row = historial[j];
          const histMatID = row[2] != null ? String(row[2]).trim() : "";
          if (histMatID === matID) {
            const d = safeDate(row[0]);
            if (!latest || d.getTime() > latest.date.getTime()) {
              latest = {
                date: d,
                tipo: row[4] != null ? String(row[4]).trim() : "",
                responsableID: row[1] != null ? String(row[1]).trim() : "",
                raw: row
              };
            }
            Logger.log('  historial coincide: %s | fecha(raw)=%s | tipo=%s | resp=%s',
                       JSON.stringify(row), row[0], row[4], row[1]);
          }
        }

        Logger.log('  ultimo movimiento encontrado: %s', latest ? JSON.stringify(latest.raw) : 'NINGUNO');

        // Caso: SIN historial -> solo ID, descripcion, stock
        if (!latest) {
          resultados.push({
            id: matID,
            descripcion: matDesc,
            stock: stock,
            responsable: "",
            upi: "",
            contacto: "",
            contratoFin: "",
            fechaPrestamo: "",
            fechaDevolucion: ""
          });
          continue;
        }

        // Formatear fecha para mostrar (si existe)
        const fechaFormateada = Utilities.formatDate(latest.date, tz, "dd/MM/yyyy HH:mm:ss");

        // Si el ultimo movimiento es SALIDA -> está prestado: mostrar datos del responsable y fechaPrestamo
        if (String(latest.tipo).toLowerCase() === "salida") {
          // Buscar datos del responsable
          let respNombre = "", upi = "", contacto = "", contratoFin = "";
          for (let r = 1; r < responsables.length; r++) {
            if (String(responsables[r][0]).trim() === latest.responsableID) {
              respNombre  = responsables[r][1] != null ? String(responsables[r][1]).trim() : "";
              contacto    = responsables[r][2] != null ? String(responsables[r][2]).trim() : "";
              upi         = responsables[r][3] != null ? String(responsables[r][3]).trim() : "";
              contratoFin = responsables[r][4] != null ? String(responsables[r][4]).trim() : "";
              break;
            }
          }

          resultados.push({
            id: matID,
            descripcion: matDesc,
            stock: stock,
            responsable: respNombre,
            upi: upi,
            contacto: contacto,
            contratoFin: contratoFin,
            fechaPrestamo: fechaFormateada,
            fechaDevolucion: ""
          });

        } else {
          // Ultimo movimiento = ENTRADA -> está devuelto: mostrar fechaDevolucion solamente
          resultados.push({
            id: matID,
            descripcion: matDesc,
            stock: stock,
            responsable: "",
            upi: "",
            contacto: "",
            contratoFin: "",
            fechaPrestamo: "",
            fechaDevolucion: fechaFormateada
          });
        }
      } // end if filter
    } // end for materiales

    Logger.log('buscarMaterial END - resultados encontrados: ' + resultados.length);
    return resultados;

  } catch (err) {
    Logger.log('ERROR buscarMaterial: ' + err.stack || err);
    // re-lanzo para que el cliente reciba el error (withFailureHandler)
    throw err;
  }
}
