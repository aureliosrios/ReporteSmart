import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(root, "outputs", "reportabilidad_ro");
await fs.mkdir(outputDir, { recursive: true });

const readData = async (file) => JSON.parse(await fs.readFile(path.join(root, "docs", "data", file), "utf8"));
const presupuesto = await readData("presupuesto_con_apu.json");
const base = await readData("base_datos_reportabilidad.json");
const schema = await readData("reportabilidad_schema.json");

const wbsMap = Object.fromEntries((base.catalogos?.wbs || []).map((row) => [row.codigo, row]));
const sourceByType = {
  AC_MO: "tareador.html",
  AC_MAT: "almacenero.html",
  AC_EQP: "administradora.html",
  AC_SUB: "administradora.html",
  EV_PRODUCCION: "ing_campo.html",
};

function normalizeLog(raw, index) {
  const tipo = raw.tipo || raw.categoria_evm || "";
  const log = {
    id: String(raw.id || raw.id_registro || `IMPORT-${index + 1}`),
    fecha: raw.fecha || "",
    rol: raw.rol || raw.emisor_rol || "",
    wbs: raw.wbs || raw.wbs_codigo || "",
    codigoRecurso: raw.codigoRecurso || raw.codigo_recurso_partida || "",
    detalle: raw.detalle || raw.descripcion || "",
    cantidad: Number(raw.cantidad) || 0,
    unidad: raw.unidad || "",
    pu: Number(raw.pu ?? raw.costo_unitario) || 0,
    costo: Number(raw.costo ?? raw.costo_total_pen) || 0,
    tipo,
    origen_html: sourceByType[tipo] || "origen_no_identificado",
    estado_validacion: raw.estado_validacion || "VALIDO",
  };
  return log;
}

const registros = (base.registros_diarios || []).map(normalizeLog);
const registrosValidos = registros.filter((log) => log.estado_validacion === "VALIDO");

function wbsForItem(item) {
  if (item === "01.01") return "WBS-100";
  if (item.startsWith("01.02")) return "WBS-200";
  if (item.startsWith("01.03")) return "WBS-300";
  return "WBS-400";
}

const apu = (presupuesto.analisis_precios_unitarios || []).map((row) => ({
  wbs: wbsForItem(row.item),
  item: row.item,
  descripcion: row.descripcion,
  unidad: row.unidad,
  metrado: Number(row.metrado) || 0,
  pu_meta: Number(row.precio_unitario_directo) || 0,
  bac: Number(row.costo_total_partida_directo) || 0,
}));

const bacByWbs = Object.fromEntries(Object.keys(wbsMap).map((code) => [code, apu.filter((row) => row.wbs === code).reduce((sum, row) => sum + row.bac, 0)]));
const acByWbs = Object.fromEntries(Object.keys(wbsMap).map((code) => [code, registrosValidos.filter((row) => row.wbs === code && row.tipo.startsWith("AC_")).reduce((sum, row) => sum + row.costo, 0)]));
const evByWbs = Object.fromEntries(Object.keys(wbsMap).map((code) => [code, registrosValidos.filter((row) => row.wbs === code && row.tipo === "EV_PRODUCCION").reduce((sum, row) => sum + row.costo, 0)]));
const pvByWbs = Object.fromEntries(Object.keys(wbsMap).map((code) => [code, Number(wbsMap[code].pv_acumulado_pen) || 0]));

const indicadores = Object.fromEntries(Object.keys(wbsMap).map((code) => {
  const bac = bacByWbs[code];
  const pv = pvByWbs[code];
  const ev = evByWbs[code];
  const ac = acByWbs[code];
  const cpi = ac > 0 ? ev / ac : 1;
  const spi = pv > 0 ? ev / pv : 1;
  const eac = cpi > 0 ? ac + (bac - ev) / cpi : bac;
  return [code, { bac, pv, ev, ac, cv: ev - ac, sv: ev - pv, cpi, spi, eac, variacion_eac: bac - eac }];
}));

const consolidated = {
  metadata: {
    proyecto: base.metadata?.proyecto || presupuesto.proyecto?.nombre || "",
    codigo_proyecto: base.metadata?.codigo_proyecto || "HU-CEDROS-2026",
    fecha_generacion: new Date().toISOString(),
    version_esquema: schema.version_esquema,
  },
  contrato: schema,
  wbs: Object.values(wbsMap),
  apu_presupuesto: apu,
  registros_diarios: registros,
  indicadores_evm: indicadores,
  calidad_datos: {
    registros_recibidos: registros.length,
    registros_validos: registrosValidos.length,
    duplicados_detectados: registros.length - registrosValidos.length,
    regla_deduplicacion: "fecha + rol + WBS + código + cantidad + PU + costo + tipo",
  },
};

await fs.writeFile(path.join(outputDir, "reportabilidad_consolidada.json"), JSON.stringify(consolidated, null, 2), "utf8");

const workbook = Workbook.create();
const control = workbook.worksheets.add("CONTROL");
const wbsSheet = workbook.worksheets.add("WBS");
const apuSheet = workbook.worksheets.add("APU_PRESUPUESTO");
const logsSheet = workbook.worksheets.add("REGISTROS_DIARIOS");
const evmSheet = workbook.worksheets.add("EVM_WBS");
const catalogSheet = workbook.worksheets.add("CATALOGOS");

const navy = "#1E293B";
const blue = "#0284C7";
const green = "#16A34A";
const purple = "#7E22CE";
const light = "#F8FAFC";
const white = "#FFFFFF";
const currency = '"S/" #,##0.00';

function title(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range.split(":")[0]).values = [[text]];
  sheet.getRange(range).format = { fill: navy, font: { bold: true, color: white, size: 14 }, horizontalAlignment: "center", verticalAlignment: "center" };
}

function header(sheet, range) {
  sheet.getRange(range).format = { fill: blue, font: { bold: true, color: white }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "all", style: "thin", color: "#CBD5E1" } };
}

title(control, "A1:F1", "SISTEMA DE REPORTABILIDAD RO - CONTROL LOCAL");
control.getRange("A3:B9").values = [
  ["Proyecto", consolidated.metadata.proyecto],
  ["Código", consolidated.metadata.codigo_proyecto],
  ["Fecha generación", consolidated.metadata.fecha_generacion],
  ["Registros recibidos", registros.length],
  ["Registros válidos", registrosValidos.length],
  ["Duplicados detectados", registros.length - registrosValidos.length],
  ["Fuente", "HTML → JSON → Excel (fase local)"],
];
control.getRange("A3:A9").format = { fill: light, font: { bold: true } };
control.getRange("A3:B9").format.borders = { preset: "all", style: "thin", color: "#CBD5E1" };
control.getRange("D3:F3").values = [["Indicador", "Valor", "Lectura"]];
header(control, "D3:F3");
control.getRange("D4:F8").values = [
  ["BAC total", null, "Presupuesto APU"],
  ["PV acumulado", null, "Valor planificado"],
  ["EV acumulado", null, "Valor ganado"],
  ["AC acumulado", null, "Costo real válido"],
  ["CPI global", null, "EV / AC"],
];
control.getRange("E4").formulas = [["=SUM('EVM_WBS'!C4:C7)"]];
control.getRange("E5").formulas = [["=SUM('EVM_WBS'!D4:D7)"]];
control.getRange("E6").formulas = [["=SUM('EVM_WBS'!E4:E7)"]];
control.getRange("E7").formulas = [["=SUM('EVM_WBS'!F4:F7)"]];
control.getRange("E8").formulas = [["=IF(E7>0,E6/E7,1)"]];
control.getRange("E4:E7").format.numberFormat = currency;
control.getRange("E8").format.numberFormat = "0.00";

title(wbsSheet, "A1:D1", "CATÁLOGO WBS Y LÍNEA BASE");
wbsSheet.getRange("A3:D3").values = [["Código WBS", "Descripción", "BAC desde APU", "PV acumulado"]];
header(wbsSheet, "A3:D3");
wbsSheet.getRange(`A4:D${3 + Object.keys(wbsMap).length}`).values = Object.values(wbsMap).map((row) => [row.codigo, row.nombre, bacByWbs[row.codigo], Number(row.pv_acumulado_pen) || 0]);
wbsSheet.getRange(`C4:D${3 + Object.keys(wbsMap).length}`).format.numberFormat = currency;

title(apuSheet, "A1:G1", "APU Y PRESUPUESTO META");
apuSheet.getRange("A3:G3").values = [["WBS", "Item", "Descripción", "Unidad", "Metrado", "P.U. Meta", "BAC"]];
header(apuSheet, "A3:G3");
apuSheet.getRange(`A4:G${3 + apu.length}`).values = apu.map((row) => [row.wbs, row.item, row.descripcion, row.unidad, row.metrado, row.pu_meta, row.bac]);
apuSheet.getRange(`E4:E${3 + apu.length}`).format.numberFormat = "#,##0.00";
apuSheet.getRange(`F4:G${3 + apu.length}`).format.numberFormat = currency;

title(logsSheet, "A1:N1", "REGISTROS DIARIOS NORMALIZADOS DESDE HTML");
logsSheet.getRange("A3:N3").values = [["ID Registro", "Fecha", "Rol", "WBS", "Código", "Detalle", "Cantidad", "Unidad", "P.U.", "Costo", "Tipo EVM", "Estado", "Origen HTML", "ID Parte"]];
header(logsSheet, "A3:N3");
logsSheet.getRange(`A4:N${3 + registros.length}`).values = registros.map((row) => [row.id, row.fecha, row.rol, row.wbs, row.codigoRecurso, row.detalle, row.cantidad, row.unidad, row.pu, row.costo, row.tipo, row.estado_validacion, row.origen_html, row.id_parte]);
logsSheet.getRange(`G4:G${3 + registros.length}`).format.numberFormat = "#,##0.00";
logsSheet.getRange(`I4:J${3 + registros.length}`).format.numberFormat = currency;

title(evmSheet, "A1:K1", "EVM Y RESULTADO OPERATIVO POR WBS");
evmSheet.getRange("A3:K3").values = [["WBS", "Descripción", "BAC", "PV", "EV", "AC", "CV", "CPI", "SPI", "EAC", "Estado"]];
header(evmSheet, "A3:K3");
evmSheet.getRange("A4:B7").values = Object.values(wbsMap).map((row) => [row.codigo, row.nombre]);
const lastLog = Math.max(4, 3 + registros.length);
const lastApu = 3 + apu.length;
for (let row = 4; row <= 7; row += 1) {
  evmSheet.getRange(`C${row}:K${row}`).formulas = [[
    `=SUMIF('APU_PRESUPUESTO'!$A$4:$A$${lastApu},A${row},'APU_PRESUPUESTO'!$G$4:$G$${lastApu})`,
    `=SUMIF('WBS'!$A$4:$A$7,A${row},'WBS'!$D$4:$D$7)`,
    `=SUMIFS('REGISTROS_DIARIOS'!$J$4:$J$${lastLog},'REGISTROS_DIARIOS'!$D$4:$D$${lastLog},A${row},'REGISTROS_DIARIOS'!$K$4:$K$${lastLog},"EV_PRODUCCION",'REGISTROS_DIARIOS'!$L$4:$L$${lastLog},"VALIDO")`,
    `=SUM(SUMIFS('REGISTROS_DIARIOS'!$J$4:$J$${lastLog},'REGISTROS_DIARIOS'!$D$4:$D$${lastLog},A${row},'REGISTROS_DIARIOS'!$K$4:$K$${lastLog},"AC_MO",'REGISTROS_DIARIOS'!$L$4:$L$${lastLog},"VALIDO"),SUMIFS('REGISTROS_DIARIOS'!$J$4:$J$${lastLog},'REGISTROS_DIARIOS'!$D$4:$D$${lastLog},A${row},'REGISTROS_DIARIOS'!$K$4:$K$${lastLog},"AC_MAT",'REGISTROS_DIARIOS'!$L$4:$L$${lastLog},"VALIDO"),SUMIFS('REGISTROS_DIARIOS'!$J$4:$J$${lastLog},'REGISTROS_DIARIOS'!$D$4:$D$${lastLog},A${row},'REGISTROS_DIARIOS'!$K$4:$K$${lastLog},"AC_EQP",'REGISTROS_DIARIOS'!$L$4:$L$${lastLog},"VALIDO"),SUMIFS('REGISTROS_DIARIOS'!$J$4:$J$${lastLog},'REGISTROS_DIARIOS'!$D$4:$D$${lastLog},A${row},'REGISTROS_DIARIOS'!$K$4:$K$${lastLog},"AC_SUB",'REGISTROS_DIARIOS'!$L$4:$L$${lastLog},"VALIDO"))`,
    `=E${row}-F${row}`,
    `=IF(F${row}>0,E${row}/F${row},1)`,
    `=IF(D${row}>0,E${row}/D${row},1)`,
    `=IF(H${row}>0,F${row}+(C${row}-E${row})/H${row},C${row})`,
    `=IF(F${row}=0,"SIN INICIAR",IF(H${row}<1,"ALERTA","OK"))`,
  ]];
}
evmSheet.getRange("C4:G7").format.numberFormat = currency;
evmSheet.getRange("H4:I7").format.numberFormat = "0.00";
evmSheet.getRange("J4:J7").format.numberFormat = currency;
evmSheet.getRange("A8:K8").values = [["TOTAL", "Proyecto", null, null, null, null, null, null, null, null, ""]];
evmSheet.getRange("C8:J8").formulas = [["=SUM(C4:C7)", "=SUM(D4:D7)", "=SUM(E4:E7)", "=SUM(F4:F7)", "=E8-F8", "=IF(F8>0,E8/F8,1)", "=IF(D8>0,E8/D8,1)", "=IF(H8>0,F8+(C8-E8)/H8,C8)"]];
evmSheet.getRange("C8:G8").format.numberFormat = currency;
evmSheet.getRange("H8:I8").format.numberFormat = "0.00";
evmSheet.getRange("J8").format.numberFormat = currency;
evmSheet.getRange("A8:K8").format = { fill: "#DBEAFE", font: { bold: true }, borders: { preset: "all", style: "thin", color: "#CBD5E1" } };

title(catalogSheet, "A1:E1", "CATÁLOGOS MAESTROS PARA LOS FORMULARIOS HTML");
catalogSheet.getRange("A3:E3").values = [["Tipo", "Código", "Descripción", "Unidad", "Precio"]];
header(catalogSheet, "A3:E3");
const catalogRows = [];
for (const [group, resources] of Object.entries(presupuesto.precios_recursos_maestros || {})) {
  for (const [code, item] of Object.entries(resources)) catalogRows.push([group, code, item.desc, item.unit, Number(item.precio ?? item.precio_unitario ?? 0)]);
}
catalogSheet.getRange(`A4:E${3 + catalogRows.length}`).values = catalogRows;
catalogSheet.getRange(`E4:E${3 + catalogRows.length}`).format.numberFormat = currency;

for (const sheet of [control, wbsSheet, apuSheet, logsSheet, evmSheet, catalogSheet]) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(3);
  const used = sheet.getUsedRange();
  used.format.borders = { preset: "inside", style: "thin", color: "#E2E8F0" };
  used.format.wrapText = false;
  used.format.autofitColumns();
}
logsSheet.getRange(`F4:F${3 + registros.length}`).format.columnWidth = 36;
logsSheet.getRange(`N4:N${3 + registros.length}`).format.columnWidth = 32;
apuSheet.getRange(`C4:C${3 + apu.length}`).format.columnWidth = 48;
wbsSheet.getRange("B4:B7").format.columnWidth = 42;
evmSheet.getRange("B4:B8").format.columnWidth = 42;

await fs.writeFile(path.join(outputDir, "README.txt"), "Fase local: los HTML capturan registros, app.js los normaliza a JSON y este libro Excel los consolida. La siguiente fase conectará el mismo esquema a GitHub y Google Sheets.\n", "utf8");
const preview = await workbook.render({ sheetName: "EVM_WBS", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(path.join(outputDir, "preview_evm_wbs.png"), new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(path.join(outputDir, "Reportabilidad_RO_Local.xlsx"));

console.log(JSON.stringify({ outputDir, json: "reportabilidad_consolidada.json", excel: "Reportabilidad_RO_Local.xlsx", registros: registros.length, validos: registrosValidos.length, duplicados: registros.length - registrosValidos.length }, null, 2));
