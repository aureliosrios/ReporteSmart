# Portal de Reportabilidad RO

Este directorio es la aplicación web preparada para publicarse en GitHub Pages.

- `index.html`: hub de acceso por rol.
- `tareador.html`: captura de mano de obra.
- `almacenero.html`: captura de materiales.
- `administradora.html`: captura de equipos y servicios.
- `ing_campo.html`: captura de avance físico y EV.
- `dashboard_ro.html`: lectura local de indicadores EVM.
- `assets/`: JavaScript y estilos.
- `data/`: única base JSON operativa y catálogos maestros.

La aplicación consulta `data/base_datos_reportabilidad.json`. En local, los envíos se dirigen al servidor Python; en la siguiente etapa se configurará el endpoint de Google Apps Script.
