# Portal de Reportabilidad RO

Este directorio es la aplicación web preparada para publicarse en GitHub Pages.

- `index.html`: hub de acceso por rol.
- `tareador.html`: captura de mano de obra.
- `almacenero.html`: captura de materiales.
- `administradora.html`: captura de equipos y servicios.
- `ing_campo.html`: captura de avance físico y EV.
- `dashboard_ro.html`: lectura local de indicadores EVM.
- `assets/`: JavaScript y estilos.
- `data/base_datos_reportabilidad.json`: única base JSON operativa.
- `data/catalogos/`: JSON maestros que alimentan los formularios.
  - `recursos_almacen.json`: materiales del portal de almacén.
  - `recursos_mano_obra.json`: categorías del tareo.
  - `recursos_equipos_servicios.json`: equipos y subcontratos.
  - `partidas_ev.json`: partidas para avance físico.
  - `wbs.json`: frentes de trabajo.

La aplicación consulta los catálogos de `data/catalogos/` y guarda los registros en `data/base_datos_reportabilidad.json`. En local, los envíos se dirigen al servidor Python; en la siguiente etapa se configurará el endpoint de Google Apps Script.
