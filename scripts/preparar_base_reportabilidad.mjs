import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const legacy = path.join(root, "legacy", "originales");
const dataDir = path.join(root, "docs", "data");
await fs.mkdir(dataDir, { recursive: true });

const read = async (file) => JSON.parse(await fs.readFile(path.join(legacy, file), "utf8"));
const oldBase = await read("base_datos_ro_diaria.json");
const presupuesto = await read("presupuesto_con_apu.json");
const schema = await read("reportabilidad_schema.json");

const sourceByType = {
  AC_MO: "tareador.html",
  AC_MAT: "almacenero.html",
  AC_EQP: "administradora.html",
  AC_SUB: "administradora.html",
  EV_PRODUCCION: "ing_campo.html",
};

const seenIds = new Set();
const registros = (oldBase.registros_diarios_logs || []).map((raw, index) => {
  const tipo = raw.tipo || raw.categoria_evm || "";
  const originalId = String(raw.id_registro || raw.id || `LEGACY-${index + 1}`).trim();
  const id = seenIds.has(originalId) ? `${originalId}-DUP-${index + 1}` : originalId;
  const estado = seenIds.has(originalId) ? "OBSERVADO_DUPLICADO" : "VALIDO";
  seenIds.add(originalId);

  return {
    id_registro: id,
    id_parte: `PARTE-${raw.fecha || "SIN_FECHA"}-${raw.rol || "SIN_ROL"}`,
    fecha: raw.fecha || "",
    rol: raw.rol || raw.emisor_rol || "",
    wbs_codigo: raw.wbs || raw.wbs_codigo || "",
    codigo_recurso_partida: raw.codigoRecurso || raw.codigo_recurso_partida || "",
    descripcion: raw.detalle || raw.descripcion || "",
    cantidad: Number(raw.cantidad) || 0,
    unidad: raw.unidad || "",
    costo_unitario: Number(raw.pu ?? raw.costo_unitario) || 0,
    costo_total_pen: Number(raw.costo ?? raw.costo_total_pen) || 0,
    categoria_evm: tipo,
    origen_html: sourceByType[tipo] || "origen_no_identificado",
    estado_validacion: estado,
    fecha_recepcion: raw.fecha_recepcion || null,
  };
});

const baseCanonica = {
  metadata: {
    proyecto: presupuesto.proyecto?.nombre || oldBase.metadata_proyecto?.nombre_proyecto || "",
    codigo_proyecto: oldBase.metadata_proyecto?.codigo_proyecto || "HU-CEDROS-2026",
    version_esquema: "3.0.0",
    moneda: presupuesto.proyecto?.moneda || "PEN",
    fuente: "HTML → JSON canónico → Excel / Google Sheets",
  },
  catalogos: {
    wbs: oldBase.nodos_wbs_estructura || [],
    recursos: presupuesto.precios_recursos_maestros || {},
  },
  presupuesto_apu: (presupuesto.analisis_precios_unitarios || []).map((apu) => ({
    codigo_partida: apu.item,
    descripcion: apu.descripcion,
    unidad: apu.unidad,
    metrado_meta: Number(apu.metrado) || 0,
    precio_unitario_meta: Number(apu.precio_unitario_directo) || 0,
    costo_total_meta: Number(apu.costo_total_partida_directo) || 0,
  })),
  registros_diarios: registros,
  control_calidad: {
    registros_recibidos: registros.length,
    registros_validos: registros.filter((r) => r.estado_validacion === "VALIDO").length,
    registros_observados: registros.filter((r) => r.estado_validacion !== "VALIDO").length,
    regla_duplicado: "solo se considera duplicado el mismo id_registro original",
  },
  contrato: schema,
};

await fs.writeFile(path.join(dataDir, "base_datos_reportabilidad.json"), JSON.stringify(baseCanonica, null, 2), "utf8");
await fs.copyFile(path.join(legacy, "presupuesto_con_apu.json"), path.join(dataDir, "presupuesto_con_apu.json"));
await fs.copyFile(path.join(legacy, "reportabilidad_schema.json"), path.join(dataDir, "reportabilidad_schema.json"));
console.log(JSON.stringify(baseCanonica.control_calidad, null, 2));
