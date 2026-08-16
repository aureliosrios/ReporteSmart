// ==============================================================================
// GOOGLE APPS SCRIPT - CONECTOR DE BASE DE DATOS VIVA EN GOOGLE SHEETS
// Proyecto: Redes Sanitarias de Agua Potable y Alcantarillado
// Vinculado a: https://aureliosrios.github.io/ReporteSmart/
// ==============================================================================

var TAB_NAME_LOGS = "04_LOG_FIELD_ENTRIES";

var SPREADSHEET_ID_OR_URL = null;

function getSpreadsheet() {
  try {
    if (SPREADSHEET_ID_OR_URL) {
      if (SPREADSHEET_ID_OR_URL.indexOf("https://") === 0) {
        return SpreadsheetApp.openByUrl(SPREADSHEET_ID_OR_URL);
      } else {
        return SpreadsheetApp.openById(SPREADSHEET_ID_OR_URL);
      }
    }
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (err) {
    console.error("Error al obtener hoja:", err);
  }
  return null;
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "ERROR", message: "Payload JSON no recibido" });
    }

    var rawContent = JSON.parse(e.postData.contents);
    var records = Array.isArray(rawContent) ? rawContent : [rawContent];

    var ss = getSpreadsheet();
    if (!ss) {
      return responseJSON({ 
        status: "ERROR", 
        message: "No se pudo acceder al Google Sheet activo. Asegúrate de abrir Apps Script desde el menú Extensiones > Apps Script de tu hoja de cálculo actual." 
      });
    }

    var sheet = ss.getSheetByName(TAB_NAME_LOGS);

    if (!sheet) {
      sheet = ss.insertSheet(TAB_NAME_LOGS);
      sheet.appendRow([
        "ID Registro", "Fecha", "Rol Responsable", "Código WBS", 
        "Código Recurso/Partida", "Descripción / Detalle", "Cantidad Campo", 
        "Unidad", "P.U. (Busca en Maestro)", "Subtotal Monto (S/)", 
        "Categoría EVM", "Origen HTML"
      ]);
    }

    // Determinar el siguiente número secuencial para el ID LOG-YYYYMMDD-XXX
    var nextNum = 1;
    var todayStr = new Date().toISOString().split("T")[0].replace(/-/g, ""); // YYYYMMDD
    var startRow = sheet.getLastRow();
    
    if (startRow >= 2) {
      var ids = sheet.getRange(2, 1, startRow - 1, 1).getValues();
      for (var r = ids.length - 1; r >= 0; r--) {
        var currId = ids[r][0].toString();
        if (currId.indexOf("LOG-") === 0) {
          var parts = currId.split("-");
          if (parts.length === 3) {
            var lastNum = parseInt(parts[2], 10);
            if (!isNaN(lastNum)) {
              nextNum = lastNum + 1;
              break;
            }
          }
        }
      }
    }

    var insertedCount = 0;

    for (var i = 0; i < records.length; i++) {
      var data = records[i];
      
      // Calcular la fila exacta de forma correlativa para evitar que getLastRow() devuelva
      // un valor desactualizado dentro de bucles rápidos antes de que se limpie la caché de Sheets.
      var currentRow = startRow + i + 1;

      // Generación secuencial de ID Registro para evitar códigos UUID complejos
      var numStr = nextNum.toString();
      while (numStr.length < 3) {
        numStr = "0" + numStr;
      }
      var idVal = "LOG-" + todayStr + "-" + numStr;
      nextNum++;

      var fechaVal = data.fecha || new Date().toISOString().split("T")[0];
      var rolVal = data.rol || data.emisor_rol || "Sin Rol";
      var wbsVal = data.wbs || data.wbs_codigo || "WBS-100";
      
      // Forzar formato de texto para códigos numéricos como "01.01" para evitar que Sheets los convierta en números decimales (ej. 1.01) y rompa el VLOOKUP
      var codVal = data.codigoRecurso || data.codigo_recurso_partida || "MO_PEON";
      codVal = String(codVal);
      if (codVal.match(/^\d+(\.\d+)+$/) || !isNaN(codVal)) {
        codVal = "'" + codVal;
      }

      var detVal = data.detalle || data.descripcion || "";
      var cantVal = Number(data.cantidad) || 0.0;
      var undVal = data.unidad || "und";
      var catVal = data.tipo || data.categoria_evm || "AC_MO";
      var origVal = data.origen_html || "almacenero.html";

      sheet.appendRow([
        idVal, fechaVal, rolVal, wbsVal, codVal, detVal, cantVal, undVal,
        "", "", catVal, origVal
      ]);

      // Centrar el contenido de las columnas A, D, E y H en la fila insertada
      sheet.getRange(currentRow, 1).setHorizontalAlignment("center"); // Col A (ID Registro)
      sheet.getRange(currentRow, 4).setHorizontalAlignment("center"); // Col D (Código WBS)
      sheet.getRange(currentRow, 5).setHorizontalAlignment("center"); // Col E (Código Recurso/Partida)
      sheet.getRange(currentRow, 8).setHorizontalAlignment("center"); // Col H (Unidad)

      // Fórmulas en notación A1 limpia y alineada con la estructura del Excel
      // Busca en '05_MAESTRO_RECURSOS' (Col D: P.U. Meta Oficial) o '06_MAESTRO_PARTIDAS_EV' (Col F: P.U. Directo Meta)
      var rangePU = sheet.getRange(currentRow, 9);
      rangePU.setFormula("=IFERROR(VLOOKUP(E" + currentRow + ", '05_MAESTRO_RECURSOS'!A:D, 4, FALSE), IFERROR(VLOOKUP(E" + currentRow + ", '06_MAESTRO_PARTIDAS_EV'!B:F, 5, FALSE), 0))");
      rangePU.setNumberFormat("S/ #,##0.00");

      var rangeSubtotal = sheet.getRange(currentRow, 10);
      rangeSubtotal.setFormula("=ROUND(G" + currentRow + " * I" + currentRow + ", 2)");
      rangeSubtotal.setNumberFormat("S/ #,##0.00");

      insertedCount++;
    }

    return responseJSON({
      status: "SUCCESS",
      message: "Se insertaron " + insertedCount + " registro(s) correctamente en la pestaña " + TAB_NAME_LOGS + " de Google Sheets.",
      count: insertedCount
    });

  } catch (error) {
    return responseJSON({
      status: "ERROR",
      message: "Error al procesar registro en Google Sheets: " + error.toString()
    });
  }
}

function doGet(e) {
  try {
    return responseJSON({
      status: "ONLINE",
      version: "1.0.3",
      proyecto: "Redes Sanitarias de Agua Potable y Alcantarillado",
      pestana_destino: TAB_NAME_LOGS,
      github_pages: "https://aureliosrios.github.io/ReporteSmart/",
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    return responseJSON({
      status: "ERROR",
      message: "Error en doGet: " + error.toString()
    });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Función de prueba directa para autorizar permisos de Google Sheets
function probarConexion() {
  var ss = getSpreadsheet();
  if (ss) {
    Logger.log("✅ Conexión exitosa al Google Sheet: " + ss.getName());
    var sheet = ss.getSheetByName(TAB_NAME_LOGS);
    if (sheet) {
      Logger.log("✅ Pestaña encontrada: " + TAB_NAME_LOGS + " con " + sheet.getLastRow() + " filas.");
      try {
        // Intentar escribir un valor de prueba en una celda temporal y luego borrarlo
        var tempRange = sheet.getRange("Z1");
        tempRange.setValue("TEST_WRITE");
        tempRange.clearContent();
        Logger.log("✅ Permiso de ESCRITURA confirmado con éxito.");
      } catch (writeErr) {
        Logger.log("❌ ERROR DE ESCRITURA: Tu cuenta no tiene permisos de Editor en este documento. Detalle: " + writeErr.toString());
      }
    } else {
      Logger.log("⚠️ La pestaña " + TAB_NAME_LOGS + " no existe aún (se creará automáticamente).");
    }
  } else {
    Logger.log("❌ No se pudo conectar a la hoja de cálculo.");
  }
}

// Función sin try-catch para forzar la ventana emergente de permisos de Google
function forzarAutorizacion() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1nlN-U7iJFBlGNS-q0xBxSzJWSDI2uMr_dCZ2eU6W3hE/edit");
  Logger.log("✅ Documento abierto con éxito: " + ss.getName());
}
