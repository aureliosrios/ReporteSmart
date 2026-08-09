import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dataDir = path.join(root, "docs", "data");
const catalogDir = path.join(dataDir, "catalogos");
await fs.mkdir(catalogDir, { recursive: true });

const presupuesto = JSON.parse(await fs.readFile(path.join(dataDir, "presupuesto_con_apu.json"), "utf8"));
const maestros = presupuesto.precios_recursos_maestros || {};

const wbs = [
  { codigo: "WBS-100", nombre: "Obras Preliminares, Trazo y Movilización" },
  { codigo: "WBS-200", nombre: "Red de Alcantarillado, Zanjas y Buzones" },
  { codigo: "WBS-300", nombre: "Red de Agua Potable y Conexiones Domiciliarias" },
  { codigo: "WBS-400", nombre: "Pruebas Hidráulicas, Desinfección y Entrega" },
];

const partidas = (presupuesto.analisis_precios_unitarios || []).map((item) => ({
  codigo_partida: item.item,
  descripcion: item.descripcion,
  unidad: item.unidad,
  metrado_meta: item.metrado,
  precio_unitario_meta: item.precio_unitario_directo,
  costo_total_meta: item.costo_total_partida_directo,
}));

const write = async (name, data) => fs.writeFile(path.join(catalogDir, name), JSON.stringify(data, null, 2), "utf8");
await write("wbs.json", wbs);
await write("recursos_mano_obra.json", maestros.mano_obra || {});
await write("recursos_almacen.json", maestros.materiales || {});
await write("recursos_equipos_servicios.json", { equipos: maestros.equipos || {}, subcontratos: maestros.subcontratos || {} });
await write("partidas_ev.json", partidas);
await write("catalogos_manifest.json", {
  version: "1.0.0",
  uso: "Catálogos consumidos por los portales HTML",
  archivos: {
    wbs: "wbs.json",
    mano_obra: "recursos_mano_obra.json",
    almacenero: "recursos_almacen.json",
    equipos_servicios: "recursos_equipos_servicios.json",
    partidas_ev: "partidas_ev.json",
  },
});

console.log(`Catálogos HTML generados en ${catalogDir}`);
