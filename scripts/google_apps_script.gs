// ==============================================================================
// GOOGLE APPS SCRIPT - CONECTOR DE BASE DE DATOS VIVA EN GOOGLE SHEETS
// Proyecto: Redes Sanitarias de Agua Potable y Alcantarillado
// Vinculado a: https://aureliosrios.github.io/ReporteSmart/
// ==============================================================================

var TAB_NAME_LOGS = "04_LOG_FIELD_ENTRIES";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "ERROR", message: "Payload JSON no recibido" });
    }

    var rawContent = JSON.parse(e.postData.contents);
    var records = Array.isArray(rawContent) ? rawContent : [rawContent];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var lastRow = sheet.getLastRow();
    
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
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

      var nextRow = sheet.getLastRow();

      // Fórmulas en notación A1 limpia y alineada con la estructura del Excel
      // Busca en '05_MAESTRO_RECURSOS' (Col D: P.U. Meta Oficial) o '06_MAESTRO_PARTIDAS_EV' (Col F: P.U. Directo Meta)
      var rangePU = sheet.getRange(nextRow, 9);
      rangePU.setFormula("=IFERROR(VLOOKUP(E" + nextRow + ", '05_MAESTRO_RECURSOS'!A:D, 4, FALSE), IFERROR(VLOOKUP(E" + nextRow + ", '06_MAESTRO_PARTIDAS_EV'!B:F, 5, FALSE), 0))");
      rangePU.setNumberFormat("S/ #,##0.00");

      var rangeSubtotal = sheet.getRange(nextRow, 10);
      rangeSubtotal.setFormula("=ROUND(G" + nextRow + " * I" + nextRow + ", 2)");
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
