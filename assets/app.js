// Motor de Inteligencia y Lógica para Portales Independientes con Sincronización en Tiempo Real a Excel

document.addEventListener('DOMContentLoaded', () => {
    const DEFAULT_PRESUPUESTO_DATA = {
        "proyecto": {
            "nombre": "Redes Sanitarias de Agua Potable y Alcantarillado - Habilitación Urbana Los Cedros",
            "tipo_presupuesto": "Comercial / Licitación",
            "duracion_semanas": 8,
            "duracion_dias_calendario": 60,
            "moneda": "PEN",
            "jornal_estandar_hh": 8.0
        },
        "precios_recursos_maestros": {
            "mano_obra": {
                "MO_CAPATAZ": { "desc": "CAPATAZ", "unit": "hh", "precio": 32.50 },
                "MO_OPERARIO": { "desc": "OPERARIO", "unit": "hh", "precio": 26.80 },
                "MO_OFICIAL": { "desc": "OFICIAL", "unit": "hh", "precio": 22.40 },
                "MO_PEON": { "desc": "PEON", "unit": "hh", "precio": 20.10 }
            },
            "materiales": {
                "MAT_ESTACA": { "desc": "ESTACAS Y MADERA PARA TRAZO", "unit": "pza", "precio": 4.50 },
                "MAT_PINTURA": { "desc": "PINTURA SPRAY MARCADOR DE ZANJA", "unit": "gla", "precio": 18.00 },
                "MAT_CORDEL": { "desc": "CORDEL / NYLON", "unit": "m", "precio": 0.80 },
                "MAT_ARENA_CAMAS": { "desc": "ARENA FINA SELECCIONADA PARA CAMA E=0.10M", "unit": "m3", "precio": 45.00 },
                "MAT_TUB_PVC_200": { "desc": "TUBERIA PVC UF DN 200MM SERIE S-20 ALCANTARILLADO", "unit": "m", "precio": 42.00 },
                "MAT_LUBRICANTE": { "desc": "LUBRICANTE PARA TUBERIA CON ESPIGA Y CAMPANA", "unit": "kg", "precio": 24.00 },
                "MAT_MATERIAL_PR": { "desc": "MATERIAL DE PRÉSTAMO AFIRMADO PARA RELLENO COMPACTADO", "unit": "m3", "precio": 38.00 },
                "MAT_AGUA": { "desc": "AGUA PARA COMPACTACIÓN Y PRUEBAS", "unit": "m3", "precio": 12.00 },
                "MAT_BUZON_PREF": { "desc": "BUZON PREFABRICADO CONCRETO H=1.5-2.5M CON MARCO Y TAPA", "unit": "und", "precio": 1050.00 },
                "MAT_CAJA_REGISTRO": { "desc": "CAJA DE REGISTRO PREFABRICADA DE CONCRETO 12X24 WITH TAPA", "unit": "und", "precio": 145.00 },
                "MAT_TUB_PVC_160_CONEX": { "desc": "TUBERIA PVC SAL DN 160MM PARA ACOMETIDA ALCANTARILLADO", "unit": "m", "precio": 22.00 },
                "MAT_TUB_PVC_110_AGUA": { "desc": "TUBERIA PVC C-10 DN 110MM AGUA POTABLE", "unit": "m", "precio": 26.50 },
                "MAT_ACCESORIOS_AGUA": { "desc": "ACCESORIOS Y CODOS PVC AGUA 110MM (PROMEDIO/M)", "unit": "glb", "precio": 5.00 },
                "MAT_CAJA_AGUA": { "desc": "CAJA TERMOFORMADA PARA MEDIDOR + ABRAZADERA + LLAVE DE PASO", "unit": "und", "precio": 62.00 },
                "MAT_TUB_HDPE_1_2": { "desc": "TUBERIA HDPE 1/2 PARA ACOMETIDA AGUA", "unit": "m", "precio": 3.80 },
                "MAT_HIPOCLORITO": { "desc": "HIPOCLORITO DE CALCIO 70% PARA DESINFECCION", "unit": "kg", "precio": 35.00 },
                "MAT_CARTEL": { "desc": "CARTEL DE OBRA 3.60X2.40M", "unit": "und", "precio": 1200.00 },
                "MAT_CASETA": { "desc": "CASETA DE ALMACEN Y OFICINA PROVISIONAL", "unit": "glb", "precio": 2300.00 }
            },
            "equipos": {
                "EQ_HERRAMIENTAS": { "desc": "HERRAMIENTAS MANUALES", "unit": "%MO", "precio": 0.00 },
                "EQ_EXCAVADORA": { "desc": "EXCAVADORA SOBRE ORUGAS 140 HP", "unit": "hm", "precio": 185.00 },
                "EQ_RETROEXCAVADORA": { "desc": "RETROEXCAVADORA SOBRE LLANTAS 62 HP", "unit": "hm", "precio": 130.00 },
                "EQ_PLANCHA": { "desc": "PLANCHA COMPACTADORA 7 HP", "unit": "hm", "precio": 25.00 },
                "EQ_RODILLO_CAMINANTE": { "desc": "RODILLO VIBRATORIO CAMINANTE 1.5 TN", "unit": "hm", "precio": 45.00 },
                "EQ_CAMION_CISTERNA": { "desc": "CAMION CISTERNA 4X2 (AGUA) 2,000 GLN", "unit": "hm", "precio": 140.00 },
                "EQ_CAMION_GRUAN": { "desc": "CAMION GRUA 5 TN PARA MONTAJE DE BUZONES", "unit": "hm", "precio": 160.00 },
                "EQ_BOMBA_PRUEBA": { "desc": "BOMBA HIDROSTÁTICA DE PRUEBA DE PRESIÓN", "unit": "hm", "precio": 30.00 },
                "EQ_ESTACION_TOTAL": { "desc": "EQUIPO DE TOPOGRAFIA ESTACION TOTAL", "unit": "hm", "precio": 18.00 }
            },
            "subcontratos": {
                "SUB_MOVILIZACION": { "desc": "SUBCONTRATO DE MOVILIZACION Y DESMOVILIZACION DE MAQUINARIA", "unit": "glb", "precio": 4000.00 },
                "SUB_ENSAYOS_COMPAC": { "desc": "SUBCONTRATO DE ENSAYOS DE DENSIDAD DE CAMPO (PROCTOR/DENSIDAD)", "unit": "und", "precio": 80.00 },
                "SUB_MONTAJE_BUZON": { "desc": "SUBCONTRATO SERVICIO ASENTADO Y ANCLAJE BUZONES PREFABRICADOS", "unit": "und", "precio": 250.00 },
                "SUB_PRUEBA_LAB": { "desc": "SUBCONTRATO PRUEBAS DE LABORATORIO DE AGUA (BACTERIOLOGICO/FISICOQUIMICO)", "unit": "glb", "precio": 3800.00 }
            }
        },
        "analisis_precios_unitarios": [
            { "item": "01.01", "descripcion": "Obras Preliminares y Trabajos Provisionales", "unidad": "GLB", "metrado": 1.0, "precio_unitario_directo": 14500.00, "costo_total_partida_directo": 14500.00 },
            { "item": "01.02.01", "descripcion": "Trazo, Nivelación y Replanteo de Zanjas", "unidad": "M", "metrado": 2400.0, "precio_unitario_directo": 3.50, "costo_total_partida_directo": 8400.00 },
            { "item": "01.02.02", "descripcion": "Excavación de Zanja H=1.50m - 2.20m a Máquina (Terreno Normal)", "unidad": "M3", "metrado": 3840.0, "precio_unitario_directo": 14.50, "costo_total_partida_directo": 55680.00 },
            { "item": "01.02.03", "descripcion": "Preparación y Colocación de Cama de Arena e=0.10m", "unidad": "M", "metrado": 2400.0, "precio_unitario_directo": 12.80, "costo_total_partida_directo": 30720.00 },
            { "item": "01.02.04", "descripcion": "Suministro e Instalación de Tubería PVC UF DN 200mm Serie S-20 para Alcantarillado", "unidad": "M", "metrado": 1400.0, "precio_unitario_directo": 58.00, "costo_total_partida_directo": 81200.00 },
            { "item": "01.02.05", "descripcion": "Relleno Compactado de Zanja en Capas de 0.20m con Maquinaria/Plancha", "unidad": "M3", "metrado": 3300.0, "precio_unitario_directo": 16.00, "costo_total_partida_directo": 52800.00 },
            { "item": "01.02.06", "descripcion": "Construcción de Buzones Prefabricados de Concreto h=1.50m - 2.50m (Inc. Marco y Tapa)", "unidad": "UND", "metrado": 32.0, "precio_unitario_directo": 1850.00, "costo_total_partida_directo": 59200.00 },
            { "item": "01.02.07", "descripcion": "Conexiones Domiciliarias de Alcantarillado (Caja de Registro + Acometida)", "unidad": "UND", "metrado": 120.0, "precio_unitario_directo": 320.00, "costo_total_partida_directo": 38400.00 },
            { "item": "01.03.01", "descripcion": "Suministro e Instalación de Tubería PVC C-10 DN 110mm para Agua Potable", "unidad": "M", "metrado": 1000.0, "precio_unitario_directo": 39.00, "costo_total_partida_directo": 39000.00 },
            { "item": "01.03.02", "descripcion": "Conexiones Domiciliarias de Agua Potable (Caja de Agua + Abrazadera + Acometida)", "unidad": "UND", "metrado": 120.0, "precio_unitario_directo": 135.00, "costo_total_partida_directo": 16250.00 },
            { "item": "01.04.01", "descripcion": "Pruebas Hidráulicas de Redes de Agua y Alcantarillado + Desinfección", "unidad": "GLB", "metrado": 1.0, "precio_unitario_directo": 12050.00, "costo_total_partida_directo": 12050.00 }
        ]
    };

    let presupuestoData = DEFAULT_PRESUPUESTO_DATA;
    let baseDatosData = null;
    let dbLogs = readLocalLogs();
    let dbConsolidado = [];

    let moMap = {};
    let matMap = {};
    let eqpMap = {};

    let wbsOptions = [
        { code: "WBS-100", name: "WBS-100: Obras Preliminares, Trazo y Movilización" },
        { code: "WBS-200", name: "WBS-200: Red de Alcantarillado, Zanjas y Buzones" },
        { code: "WBS-300", name: "WBS-300: Red de Agua Potable y Conexiones Domiciliarias" },
        { code: "WBS-400", name: "WBS-400: Pruebas Hidráulicas, Desinfección y Entrega" }
    ];

    // Limpieza de URLs antiguas o rotas del localStorage para evitar conflictos en otras PCs
    const oldBrokenUrls = [
        'https://script.google.com/macros/s/AKfycbwecahQY_jC4kqtZiYkGSZKj5LRvgG4HHC1GOHUIvDF0obE6_kek_x8ebhZs_zd3Mp9/exec',
        'https://script.google.com/macros/s/AKfycbzpFQScHLEAe6_YdJuozbDtnIa_Wbr1JFmkzeexy1sVNv_mRr0gFOvdJ--Eb9YVxCCB/exec',
        'https://script.google.com/macros/s/AKfycbxK63b9x3zv000PeM_jw6JkjrJYe0FiKxAtEWcarBYwJpR5dO2ckcpjK3aA3R9HKF0O/exec',
        'https://script.google.com/macros/s/AKfycbyNDuJGaES1RMLoR82uC0qUtxZDKfixPsh5v4sD0IhE1EC5JOJwZTye5mjLdG3i3cbI/exec',
        'https://script.google.com/macros/s/AKfycbwr7q4wcOR_bmCsDUMudFPkU42vi1NKrYfCROwzMacGoKzN2obnBsb-a1-p82FEjX1V/exec',
        'https://script.google.com/macros/s/AKfycbzTWtvYvXmWMLoQdADUULNIRXO5ki-OmBHwbJIdeqn5ZIPE5qupFsTHftXid0Jfvze6/exec',
        'https://script.google.com/macros/s/AKfycbwqSvJoI4RLwCPKgIO_yFMXa8p9f2Zd2rtkDSjZfZnc6iaUrXzEijH8sWa4_zRgXpQ/exec',
        'https://script.google.com/macros/s/AKfycbxvQXdRmbwR20YkHhL6dbwnFG-ZqRm8-NQE65i7xXPVAW-GXE8_BT7LrDZTlevi_QQc/exec'
    ];
    const storedEndpoint = localStorage.getItem('ro_api_endpoint');
    if (storedEndpoint && oldBrokenUrls.indexOf(storedEndpoint) !== -1) {
        localStorage.removeItem('ro_api_endpoint');
    }

    const DATA_PATH = 'data/';
    const CATALOG_PATH = `${DATA_PATH}catalogos/`;
    const API_ENDPOINT = window.RO_API_ENDPOINT || localStorage.getItem('ro_api_endpoint') || 'https://script.google.com/macros/s/AKfycbycOxMRY83hp6SU185I942sojJ-UieDEsRzaFP3VfhjJ5vIZ0WFjUj5Vt__5YCeR4tD/exec';

    initApp();

    function readLocalLogs() {
        // En el dashboard no leemos borradores del localStorage, solo base oficial
        if (location.pathname.indexOf('dashboard') !== -1 || document.getElementById('tbody-wbs-evm')) {
            return [];
        }
        try {
            const raw = JSON.parse(localStorage.getItem('ro_unified_logs')) || [];
            return raw.map(normalizeLog).filter(Boolean);
        } catch (err) {
            console.warn('No se pudo leer el almacenamiento local de logs:', err);
            return [];
        }
    }

    function normalizeLog(log) {
        if (!log || typeof log !== 'object') return null;
        return {
            id: log.id || log.id_registro || `LOCAL-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            id_parte: log.id_parte || '',
            fecha: log.fecha || '',
            rol: log.rol || log.emisor_rol || '',
            wbs: log.wbs || log.wbs_codigo || '',
            codigoRecurso: log.codigoRecurso || log.codigo_recurso_partida || '',
            detalle: log.detalle || log.descripcion || '',
            cantidad: Number(log.cantidad) || 0,
            unidad: log.unidad || '',
            pu: Number(log.pu ?? log.costo_unitario) || 0,
            costo: Number(log.costo ?? log.costo_total_pen) || 0,
            tipo: log.tipo || log.categoria_evm || '',
            origen_html: log.origen_html || '',
            estado_validacion: log.estado_validacion || 'VALIDO'
        };
    }

    function mergeLogs(...sources) {
        const unique = new Map();
        sources.flat().map(normalizeLog).filter(log => log && log.estado_validacion !== 'OBSERVADO_DUPLICADO').forEach(log => {
            const key = String(log.id);
            unique.set(key, log);
        });
        return [...unique.values()];
    }

    async function initApp() {
        setupGlobalButtons();

        try {
            const resp = await fetch(`${DATA_PATH}presupuesto_con_apu.json`);
            if (resp.ok) {
                presupuestoData = await resp.json();
            }
        } catch (err) {
            console.warn("Uso de base de datos offline:", err);
        }

        try {
            const catalogResponses = await Promise.all([
                fetch(`${CATALOG_PATH}wbs.json`),
                fetch(`${CATALOG_PATH}recursos_mano_obra.json`),
                fetch(`${CATALOG_PATH}recursos_almacen.json`),
                fetch(`${CATALOG_PATH}recursos_equipos_servicios.json`),
                fetch(`${CATALOG_PATH}partidas_ev.json`)
            ]);
            if (catalogResponses.every(response => response.ok)) {
                const [wbs, manoObra, almacen, equiposServicios, partidas] = await Promise.all(catalogResponses.map(response => response.json()));
                wbsOptions = wbs.map(row => ({ code: row.codigo, name: `${row.codigo}: ${row.nombre}` }));
                presupuestoData = {
                    ...presupuestoData,
                    precios_recursos_maestros: {
                        mano_obra: manoObra,
                        materiales: almacen,
                        equipos: equiposServicios.equipos || {},
                        subcontratos: equiposServicios.subcontratos || {}
                    },
                    analisis_precios_unitarios: partidas.map(row => ({
                        item: row.codigo_partida,
                        descripcion: row.descripcion,
                        unidad: row.unidad,
                        metrado: row.metrado_meta,
                        precio_unitario_directo: row.precio_unitario_meta,
                        costo_total_partida_directo: row.costo_total_meta
                    }))
                };
            }
        } catch (err) {
            console.warn("Catálogos separados no disponibles; se usa el presupuesto maestro:", err);
        }

        try {
            const dbResp = await fetch(`${DATA_PATH}base_datos_reportabilidad.json`);
            if (dbResp.ok) {
                baseDatosData = await dbResp.json();
                dbLogs = mergeLogs(dbLogs, baseDatosData.registros_diarios || []);
                localStorage.setItem('ro_unified_logs', JSON.stringify(dbLogs));
            }
        } catch (err) {
            console.warn("Base operativa no disponible; se usan los logs locales:", err);
        }

        // Sincronización en tiempo real desde Google Sheets (si está disponible)
        try {
            const sheetResp = await fetch(API_ENDPOINT);
            if (sheetResp.ok) {
                const sheetData = await sheetResp.json();
                if (sheetData && sheetData.status === 'SUCCESS' && Array.isArray(sheetData.logs)) {
                    // Convertir los registros de Google Sheets al formato esperado del cliente
                    const remoteLogs = sheetData.logs.map(row => ({
                        id: row.id,
                        fecha: row.fecha,
                        rol: row.rol,
                        wbs: row.wbs,
                        codigoRecurso: row.codigoRecurso,
                        detalle: row.detalle,
                        cantidad: Number(row.cantidad) || 0,
                        unidad: row.unidad,
                        pu: Number(row.pu) || 0,
                        costo: Number(row.costo) || 0,
                        tipo: row.tipo,
                        origen_html: row.origen_html
                    }));
                    dbLogs = mergeLogs(dbLogs, remoteLogs);
                    if (Array.isArray(sheetData.consolidado)) {
                        dbConsolidado = sheetData.consolidado;
                        console.log(`✅ Sincronizados ${dbConsolidado.length} registros consolidados (03_CONSOLIDADO_DIARIO_EVM_WBS) desde Google Sheets.`);
                    }
                    localStorage.setItem('ro_unified_logs', JSON.stringify(dbLogs));
                    console.log(`✅ Sincronizados ${remoteLogs.length} registros en tiempo real desde Google Sheets.`);
                }
            }
        } catch (err) {
            console.warn("No se pudo obtener datos en tiempo real desde Google Sheets:", err);
        }

        buildResourceMaps();

        if (document.getElementById('rows-tareador-only')) setupTareadorPortal();
        if (document.getElementById('rows-almacenero-only')) setupAlmaceneroPortal();
        if (document.getElementById('rows-admin-only')) setupAdminPortal();
        if (document.getElementById('rows-campo-only')) setupCampoPortal();
        if (document.getElementById('tbody-wbs-evm')) {
            setupDashboardCorteControl();
            updateDashboardViews();
        }
    }

    function setupDashboardCorteControl() {
        const inputCorte = document.getElementById('dash-fecha-corte');
        const btnHoy = document.getElementById('btn-hoy-corte');
        if (!inputCorte) return;

        // Obtener fechas reales de los registros
        const evLogs = dbLogs.filter(l => l.tipo === 'EV_PRODUCCION' && l.fecha);
        const acLogs = dbLogs.filter(l => ['AC_MO','AC_MAT','AC_EQP','AC_SUB'].includes(l.tipo) && l.fecha);
        const fechasReales = [...new Set([...evLogs, ...acLogs].map(l => l.fecha))].sort();

        const defaultDate = fechasReales.length > 0 ? fechasReales[fechasReales.length - 1] : new Date().toISOString().split('T')[0];
        inputCorte.value = defaultDate;

        inputCorte.addEventListener('change', () => {
            updateDashboardViews();
        });

        if (btnHoy) {
            btnHoy.addEventListener('click', () => {
                inputCorte.value = fechasReales.length > 0 ? fechasReales[fechasReales.length - 1] : new Date().toISOString().split('T')[0];
                updateDashboardViews();
            });
        }
    }

    function buildResourceMaps() {
        moMap = presupuestoData.precios_recursos_maestros.mano_obra;
        matMap = presupuestoData.precios_recursos_maestros.materiales;
        eqpMap = { ...presupuestoData.precios_recursos_maestros.equipos, ...presupuestoData.precios_recursos_maestros.subcontratos };
    }

    function getWBSSelectHTML(defaultWBS = "WBS-200") {
        return wbsOptions.map(w => `<option value="${w.code}" ${w.code === defaultWBS ? 'selected' : ''}>${w.name}</option>`).join('');
    }

    function createRecordId(prefix = 'REG') {
        const uuid = window.crypto?.randomUUID ? window.crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
        return `RO-${prefix}-${uuid}`;
    }

    function getRecordId(row) {
        if (!row.dataset.recordId) row.dataset.recordId = createRecordId();
        return row.dataset.recordId;
    }

    // FUNCIÓN CENTRALIZADA DE ENVIÓ SIMULTÁNEO AL BACKEND EXCEL (SOPORTA FILE:// Y HTTP://)
    async function syncLogsToExcelServer(newRecords) {
        const payload = newRecords.map(r => ({
            id: r.id,
            fecha: r.fecha,
            rol: r.rol,
            wbs: r.wbs,
            codigoRecurso: r.codigoRecurso,
            detalle: r.detalle,
            cantidad: r.cantidad,
            unidad: r.unidad,
            tipo: r.tipo,
            origen_html: r.origen_html || `${location.pathname.split('/').pop()}`
        }));

        const apiUrl = window.RO_API_ENDPOINT || localStorage.getItem('ro_api_endpoint') || 'https://script.google.com/macros/s/AKfycbycOxMRY83hp6SU185I942sojJ-UieDEsRzaFP3VfhjJ5vIZ0WFjUj5Vt__5YCeR4tD/exec';

        try {
            // Usamos Content-Type text/plain y mode: no-cors para evitar el preflight OPTIONS y el bloqueo por redirección de Google Apps Script
            await fetch(apiUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(payload)
            });

            dbLogs = mergeLogs(dbLogs, newRecords);
            localStorage.setItem('ro_unified_logs', JSON.stringify(dbLogs));
            alert(`✅ ${newRecords.length} registro(s) enviado(s) exitosamente a Google Sheets.`);
            if (document.getElementById('tbody-wbs-evm')) updateDashboardViews();
        } catch (err) {
            console.warn("Servidor backend no disponible; guardando en almacenamiento web local:", err);
            dbLogs = mergeLogs(dbLogs, newRecords);
            localStorage.setItem('ro_unified_logs', JSON.stringify(dbLogs));
            alert(`⚠️ No se pudo conectar a Google Sheets. Guardado localmente.`);
        }
    }

    // --- 1. PORTAL TAREADOR ---
    function setupTareadorPortal() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('form-date-tareador').value = today;

        const tbody = document.getElementById('rows-tareador-only');
        tbody.innerHTML = '';

        const addRow = (wbsVal = "WBS-200") => {
            const tr = document.createElement('tr');
            tr.dataset.recordId = createRecordId('MO');
            const moOptions = Object.keys(moMap).map(k => `<option value="${k}" data-precio="${moMap[k].precio}">${moMap[k].desc} (S/ ${moMap[k].precio.toFixed(2)}/hh)</option>`).join('');

            tr.innerHTML = `
                <td><select class="input-wbs">${getWBSSelectHTML(wbsVal)}</select></td>
                <td><select class="input-cat">${moOptions}</select></td>
                <td><input type="number" class="input-num" value="1" min="1"></td>
                <td><input type="number" class="input-hh" value="8.0" step="0.5"></td>
                <td><button type="button" class="btn-icon-danger btn-del"><i class="fa-solid fa-xmark"></i></button></td>
            `;

            tbody.appendChild(tr);
            bindCalcEvents(tr, calcTareadorSummary);
            calcTareadorSummary();
        };

        document.getElementById('btn-add-row-tareador').onclick = () => addRow();
        addRow("WBS-200");

        document.getElementById('btn-save-tareador').onclick = () => {
            const fecha = document.getElementById('form-date-tareador').value;
            const recordsToSave = [];

            tbody.querySelectorAll('tr').forEach(tr => {
                const wbs = tr.querySelector('.input-wbs').value;
                const sel = tr.querySelector('.input-cat');
                const opt = sel.options[sel.selectedIndex];
                const codigoRecurso = sel.value;
                const catName = opt.textContent.split('(')[0];
                const precio = parseFloat(opt.dataset.precio);
                const num = parseFloat(tr.querySelector('.input-num').value) || 0;
                const hh = parseFloat(tr.querySelector('.input-hh').value) || 0;
                const totalHH = num * hh;
                const costo = totalHH * precio;

                if (totalHH > 0) {
                    recordsToSave.push({
                        id: getRecordId(tr), id_parte: `PARTE-${fecha}-TAREO`, fecha, rol: "Tareador (Bildin)", wbs,
                        codigoRecurso, tipo: "AC_MO", detalle: `${catName} (${num} pers. x ${hh}h)`,
                        cantidad: totalHH, unidad: "hh", pu: precio, costo
                    });
                }
            });

            if (recordsToSave.length > 0) syncLogsToExcelServer(recordsToSave);
        };
    }

    function calcTareadorSummary() {
        let totalHH = 0, totalCosto = 0;
        document.querySelectorAll('#rows-tareador-only tr').forEach(tr => {
            const num = parseFloat(tr.querySelector('.input-num').value) || 0;
            const hh = parseFloat(tr.querySelector('.input-hh').value) || 0;
            const sel = tr.querySelector('.input-cat');
            const precio = parseFloat(sel.options[sel.selectedIndex].dataset.precio) || 0;
            const rowHH = num * hh;
            totalHH += rowHH;
            totalCosto += (rowHH * precio);
        });

        document.getElementById('summary-tareador-only').innerHTML = 
            `Total: <strong>${totalHH.toFixed(2)} HH</strong> | Costo: <strong>S/ ${totalCosto.toLocaleString('es-PE', {minimumFractionDigits: 2})}</strong>`;
    }

    // --- 2. PORTAL ALMACENERO ---
    function setupAlmaceneroPortal() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('form-date-almacenero').value = today;

        const tbody = document.getElementById('rows-almacenero-only');
        tbody.innerHTML = '';

        const addRow = (wbsVal = "WBS-200") => {
            const tr = document.createElement('tr');
            tr.dataset.recordId = createRecordId('MAT');
            const matOptions = Object.keys(matMap).map(k => `<option value="${k}" data-precio="${matMap[k].precio}" data-unit="${matMap[k].unit}">${matMap[k].desc} [${matMap[k].unit}] - S/ ${matMap[k].precio.toFixed(2)}</option>`).join('');

            tr.innerHTML = `
                <td><select class="input-wbs">${getWBSSelectHTML(wbsVal)}</select></td>
                <td><select class="input-mat">${matOptions}</select></td>
                <td><input type="number" class="input-cant" value="10" step="0.01"></td>
                <td><input type="text" class="input-pu" readonly style="background:rgba(255,255,255,0.05);"></td>
                <td><button type="button" class="btn-icon-danger btn-del"><i class="fa-solid fa-xmark"></i></button></td>
            `;

            tbody.appendChild(tr);
            bindCalcEvents(tr, calcAlmaceneroSummary);
            calcAlmaceneroSummary();
        };

        document.getElementById('btn-add-row-almacenero').onclick = () => addRow();
        addRow("WBS-200");

        document.getElementById('btn-save-almacenero').onclick = () => {
            const fecha = document.getElementById('form-date-almacenero').value;
            const recordsToSave = [];

            tbody.querySelectorAll('tr').forEach(tr => {
                const wbs = tr.querySelector('.input-wbs').value;
                const sel = tr.querySelector('.input-mat');
                const opt = sel.options[sel.selectedIndex];
                const codigoRecurso = sel.value;
                const desc = opt.textContent.split('[')[0];
                const unidad = opt.dataset.unit;
                const precio = parseFloat(opt.dataset.precio);
                const cant = parseFloat(tr.querySelector('.input-cant').value) || 0;
                const costo = cant * precio;

                if (cant > 0) {
                    recordsToSave.push({
                        id: getRecordId(tr), id_parte: `PARTE-${fecha}-ALMACEN`, fecha, rol: "Almacenero", wbs,
                        codigoRecurso, tipo: "AC_MAT", detalle: desc, cantidad: cant, unidad, pu: precio, costo
                    });
                }
            });

            if (recordsToSave.length > 0) syncLogsToExcelServer(recordsToSave);
        };
    }

    function calcAlmaceneroSummary() {
        let totalCosto = 0;
        document.querySelectorAll('#rows-almacenero-only tr').forEach(tr => {
            const sel = tr.querySelector('.input-mat');
            const opt = sel.options[sel.selectedIndex];
            const precio = parseFloat(opt.dataset.precio) || 0;
            const cant = parseFloat(tr.querySelector('.input-cant').value) || 0;

            tr.querySelector('.input-pu').value = `S/ ${precio.toFixed(2)}`;
            totalCosto += (cant * precio);
        });

        document.getElementById('summary-almacenero-only').innerHTML = 
            `Total Insumos: <strong>S/ ${totalCosto.toLocaleString('es-PE', {minimumFractionDigits: 2})}</strong>`;
    }

    // --- 3. PORTAL ADMINISTRADORA ---
    function setupAdminPortal() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('form-date-admin').value = today;

        const tbody = document.getElementById('rows-admin-only');
        tbody.innerHTML = '';

        const addRow = (wbsVal = "WBS-200") => {
            const tr = document.createElement('tr');
            tr.dataset.recordId = createRecordId('EQP');
            const eqpOptions = Object.keys(eqpMap).filter(k => k !== "EQ_HERRAMIENTAS").map(k => `<option value="${k}" data-precio="${eqpMap[k].precio}" data-unit="${eqpMap[k].unit}">${eqpMap[k].desc} [${eqpMap[k].unit}] - S/ ${eqpMap[k].precio.toFixed(2)}</option>`).join('');

            tr.innerHTML = `
                <td><select class="input-wbs">${getWBSSelectHTML(wbsVal)}</select></td>
                <td><select class="input-eqp">${eqpOptions}</select></td>
                <td><input type="number" class="input-cant" value="8" step="0.5"></td>
                <td><input type="text" class="input-pu" readonly style="background:rgba(255,255,255,0.05);"></td>
                <td><button type="button" class="btn-icon-danger btn-del"><i class="fa-solid fa-xmark"></i></button></td>
            `;

            tbody.appendChild(tr);
            bindCalcEvents(tr, calcAdminSummary);
            calcAdminSummary();
        };

        document.getElementById('btn-add-row-admin').onclick = () => addRow();
        addRow("WBS-200");

        document.getElementById('btn-save-admin').onclick = () => {
            const fecha = document.getElementById('form-date-admin').value;
            const recordsToSave = [];

            tbody.querySelectorAll('tr').forEach(tr => {
                const wbs = tr.querySelector('.input-wbs').value;
                const sel = tr.querySelector('.input-eqp');
                const opt = sel.options[sel.selectedIndex];
                const codigoRecurso = sel.value;
                const desc = opt.textContent.split('[')[0];
                const unidad = opt.dataset.unit;
                const precio = parseFloat(opt.dataset.precio);
                const cant = parseFloat(tr.querySelector('.input-cant').value) || 0;
                const costo = cant * precio;

                if (cant > 0) {
                    recordsToSave.push({
                        id: getRecordId(tr), id_parte: `PARTE-${fecha}-ADMIN`, fecha, rol: "Administradora", wbs,
                        codigoRecurso, tipo: codigoRecurso.startsWith('SUB_') ? "AC_SUB" : "AC_EQP", detalle: desc, cantidad: cant, unidad, pu: precio, costo
                    });
                }
            });

            if (recordsToSave.length > 0) syncLogsToExcelServer(recordsToSave);
        };
    }

    function calcAdminSummary() {
        let totalCosto = 0;
        document.querySelectorAll('#rows-admin-only tr').forEach(tr => {
            const sel = tr.querySelector('.input-eqp');
            const opt = sel.options[sel.selectedIndex];
            const precio = parseFloat(opt.dataset.precio) || 0;
            const cant = parseFloat(tr.querySelector('.input-cant').value) || 0;

            tr.querySelector('.input-pu').value = `S/ ${precio.toFixed(2)}`;
            totalCosto += (cant * precio);
        });

        document.getElementById('summary-admin-only').innerHTML = 
            `Total Equipos: <strong>S/ ${totalCosto.toLocaleString('es-PE', {minimumFractionDigits: 2})}</strong>`;
    }

    // --- 4. PORTAL ING. CAMPO ---
    function setupCampoPortal() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('form-date-campo').value = today;

        const tbody = document.getElementById('rows-campo-only');
        tbody.innerHTML = '';

        const addRow = (wbsVal = "WBS-200") => {
            const tr = document.createElement('tr');
            tr.dataset.recordId = createRecordId('EV');
            const partidaOptions = presupuestoData.analisis_precios_unitarios.map(apu => 
                `<option value="${apu.item}" data-pu="${apu.precio_unitario_directo}" data-unit="${apu.unidad}" data-desc="${apu.descripcion}">${apu.item} - ${apu.descripcion} (${apu.unidad})</option>`
            ).join('');

            tr.innerHTML = `
                <td><select class="input-wbs">${getWBSSelectHTML(wbsVal)}</select></td>
                <td><select class="input-partida">${partidaOptions}</select></td>
                <td><input type="text" class="input-tramo" value="Tramo Calle 1"></td>
                <td><input type="number" class="input-metrado" value="50" step="0.01"></td>
                <td><input type="text" class="input-unidad" readonly style="background:rgba(255,255,255,0.05); text-align:center; font-weight:bold;"></td>
                <td><button type="button" class="btn-icon-danger btn-del"><i class="fa-solid fa-xmark"></i></button></td>
            `;

            tbody.appendChild(tr);

            const updateUnitCell = () => {
                const sel = tr.querySelector('.input-partida');
                const opt = sel.options[sel.selectedIndex];
                tr.querySelector('.input-unidad').value = opt ? opt.dataset.unit : '-';
            };

            tr.querySelector('.input-partida').addEventListener('change', updateUnitCell);
            updateUnitCell();

            bindCalcEvents(tr, calcCampoSummary);
            calcCampoSummary();
        };

        document.getElementById('btn-add-row-campo').onclick = () => addRow();
        addRow("WBS-200");

        document.getElementById('btn-save-campo').onclick = () => {
            const fecha = document.getElementById('form-date-campo').value;
            const recordsToSave = [];

            tbody.querySelectorAll('tr').forEach(tr => {
                const wbs = tr.querySelector('.input-wbs').value;
                const sel = tr.querySelector('.input-partida');
                const opt = sel.options[sel.selectedIndex];
                const itemCode = opt.value;
                const pu = parseFloat(opt.dataset.pu);
                const unidad = opt.dataset.unit;
                const tramo = tr.querySelector('.input-tramo').value;
                const metrado = parseFloat(tr.querySelector('.input-metrado').value) || 0;
                const evCosto = metrado * pu;

                if (metrado > 0) {
                    const descPartida = opt.dataset.desc || '';
                    const detalleCompuesto = (descPartida + ' - ' + tramo).toUpperCase();
                    recordsToSave.push({
                        id: getRecordId(tr), id_parte: `PARTE-${fecha}-CAMPO`, fecha, rol: "Ing. de Campo", wbs,
                        codigoRecurso: itemCode, tipo: "EV_PRODUCCION", detalle: detalleCompuesto,
                        cantidad: metrado, unidad, pu, costo: evCosto
                    });
                }
            });

            if (recordsToSave.length > 0) syncLogsToExcelServer(recordsToSave);
        };
    }

    function calcCampoSummary() {
        let totalEV = 0;
        document.querySelectorAll('#rows-campo-only tr').forEach(tr => {
            const sel = tr.querySelector('.input-partida');
            const opt = sel.options[sel.selectedIndex];
            const pu = parseFloat(opt.dataset.pu) || 0;
            const metrado = parseFloat(tr.querySelector('.input-metrado').value) || 0;
            totalEV += (metrado * pu);
        });

        document.getElementById('summary-campo-only').innerHTML = 
            `Total Valor Ganado ($EV$): <strong>S/ ${totalEV.toLocaleString('es-PE', {minimumFractionDigits: 2})}</strong>`;
    }

    function bindCalcEvents(tr, calcFn) {
        tr.querySelectorAll('input, select').forEach(el => {
            el.addEventListener('input', calcFn);
            el.addEventListener('change', calcFn);
        });
        tr.querySelector('.btn-del').onclick = () => {
            const parent = tr.parentElement;
            if (parent.children.length > 1) {
                tr.remove();
                calcFn();
            } else {
                alert("Debe mantener al menos 1 fila.");
            }
        };
    }

    // --- 5. DASHBOARD RO Y CUANTIFICACIÓN POR WBS ---
    function updateDashboardViews() {
        const inputCorte = document.getElementById('dash-fecha-corte');
        const fechaCorte = inputCorte ? inputCorte.value : null;
        const evm = calculateEVMByWBS(fechaCorte);

        document.getElementById('dash-pv').textContent = `S/ ${evm.totalPV.toLocaleString('es-PE', {minimumFractionDigits:2})}`;
        document.getElementById('dash-ev').textContent = `S/ ${evm.totalEV.toLocaleString('es-PE', {minimumFractionDigits:2})}`;
        document.getElementById('dash-ac').textContent = `S/ ${evm.totalAC.toLocaleString('es-PE', {minimumFractionDigits:2})}`;
        document.getElementById('dash-eac').textContent = `S/ ${evm.eacGlobal.toLocaleString('es-PE', {minimumFractionDigits:2})}`;

        document.getElementById('dash-spi-sub').textContent = `SPI: ${evm.spiGlobal.toFixed(2)}`;
        document.getElementById('dash-cpi-sub').textContent = `CPI: ${evm.cpiGlobal.toFixed(2)}`;
        document.getElementById('dash-desvio-sub').textContent = `Desvío Meta: S/ ${evm.desvioGlobal.toLocaleString('es-PE', {minimumFractionDigits:2})}`;

        const tbodyCuant = document.getElementById('tbody-cuantificacion-wbs');
        if (tbodyCuant) {
            tbodyCuant.innerHTML = '';
            Object.values(evm.wbsResMap).forEach(w => {
                const tr = document.createElement('tr');
                const difClass = w.dif >= 0 ? 'text-green' : 'text-red';
                tr.innerHTML = `
                    <td><strong>${w.code}</strong></td>
                    <td>${w.name}</td>
                    <td>S/ ${w.acMO.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td>S/ ${w.acMAT.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td>S/ ${w.acEQP.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td><strong>S/ ${w.ac.toLocaleString('es-PE', {minimumFractionDigits:2})}</strong></td>
                    <td style="color:#4ade80; font-weight:bold;">S/ ${w.ev.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td style="font-weight:bold;" class="${difClass}">S/ ${w.dif.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td style="font-weight:bold; color:${w.cpi < 1.0 && w.ac > 0 ? '#f87171' : '#4ade80'}">${w.cpi.toFixed(2)}</td>
                `;
                tbodyCuant.appendChild(tr);
            });
        }

        const tbodyWBS = document.getElementById('tbody-wbs-evm');
        if (tbodyWBS) {
            tbodyWBS.innerHTML = '';
            Object.values(evm.wbsResMap).forEach(w => {
                const tr = document.createElement('tr');
                const badgeClass = w.ac === 0 ? 'badge-info' : w.cpi >= 1.0 ? 'badge-success' : 'badge-danger';
                const badgeText = w.ac === 0 ? 'Sin Iniciar' : w.cpi >= 1.0 ? 'OK / Saludable' : 'Alerta Sobrecosto';

                tr.innerHTML = `
                    <td><strong>${w.code}</strong></td>
                    <td>${w.name}</td>
                    <td>S/ ${w.bac.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td>S/ ${w.pv.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td>S/ ${w.ev.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td>S/ ${w.ac.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td style="color:${w.cpi < 1.0 && w.ac > 0 ? '#f87171' : '#4ade80'}; font-weight:bold;">${w.cpi.toFixed(2)}</td>
                    <td style="font-weight:bold;">${w.spi.toFixed(2)}</td>
                    <td>S/ ${w.eac.toLocaleString('es-PE', {minimumFractionDigits:2})}</td>
                    <td><span class="badge ${badgeClass}">${badgeText}</span></td>
                `;
                tbodyWBS.appendChild(tr);
            });
        }
        renderSCurve(evm, fechaCorte);
    }

    function calculateEVMByWBS(fechaCorte) {
        const wbsDefinitions = buildWBSBaseline();

        // Encontrar la fecha de inicio del proyecto (primer registro disponible)
        const evLogsAll = dbLogs.filter(l => l.tipo === 'EV_PRODUCCION' && l.fecha);
        const acLogsAll = dbLogs.filter(l => ['AC_MO','AC_MAT','AC_EQP','AC_SUB'].includes(l.tipo) && l.fecha);
        const fechasReales = [...new Set([...evLogsAll, ...acLogsAll].map(l => l.fecha))].sort();
        const fechaInicioStr = fechasReales.length > 0 ? fechasReales[0] : new Date().toISOString().split('T')[0];

        // Días transcurridos hasta la fecha de corte
        const corteDateStr = fechaCorte || (fechasReales.length > 0 ? fechasReales[fechasReales.length - 1] : new Date().toISOString().split('T')[0]);
        const inicioDate = new Date(fechaInicioStr + 'T00:00:00');
        const corteDate = new Date(corteDateStr + 'T00:00:00');
        const diffTime = corteDate - inicioDate;
        const t = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1);
        const duracionDias = presupuestoData.proyecto?.duracion_dias_calendario || 60;

        const wbsResMap = {};
        let totalBAC = 0, totalPV = 0, totalEV = 0, totalAC = 0;

        // Intentar usar los datos reales consolidados de la pestaña 03 del Google Sheet si están disponibles
        const consFiltrado = dbConsolidado.filter(c => c.fecha === corteDateStr);
        const usarSheetConsolidado = consFiltrado.length > 0;

        if (usarSheetConsolidado) {
            console.log(`📊 Usando datos consolidados reales de Google Sheets para la fecha de corte: ${corteDateStr}`);
            consFiltrado.forEach(c => {
                const code = c.wbs;
                const def = wbsDefinitions[code] || { name: c.descripcion || code };
                
                const bac = c.bac || 0;
                const pv = c.pv_acumulado || 0;
                const ev = c.ev_acumulado || 0;
                const ac = c.ac_acumulado || 0;
                const dif = c.variacion || (ev - ac);

                const cpi = ac > 0 ? ev / ac : 1.0;
                const spi = pv > 0 ? ev / pv : 1.0;
                const eac = cpi > 0 ? ac + (bac - ev) / cpi : bac;

                totalBAC += bac;
                totalPV += pv;
                totalEV += ev;
                totalAC += ac;

                // Calcular desglose de costos AC por componente de forma retroactiva para la Tabla 1
                const logsMO = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_MO" && l.fecha && l.fecha <= corteDateStr);
                const logsMAT = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_MAT" && l.fecha && l.fecha <= corteDateStr);
                const logsEQP = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_EQP" && l.fecha && l.fecha <= corteDateStr);
                const logsSUB = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_SUB" && l.fecha && l.fecha <= corteDateStr);

                const acMO = logsMO.reduce((sum, l) => sum + (l.costo || 0), 0);
                const acMAT = logsMAT.reduce((sum, l) => sum + (l.costo || 0), 0);
                const acEQP = logsEQP.reduce((sum, l) => sum + (l.costo || 0), 0);
                const acSUB = logsSUB.reduce((sum, l) => sum + (l.costo || 0), 0);

                wbsResMap[code] = { code, name: def.name, bac, pv, ev, ac, acMO, acMAT, acEQP, acSUB, dif, cpi, spi, eac };
            });
        } else {
            // FALLBACK LOCAL: Si está offline o no hay registros para esta fecha en la pestaña 03
            Object.keys(wbsDefinitions).forEach(code => {
                const def = wbsDefinitions[code];
                totalBAC += def.bac;

                // Calcular PV acumulado lineal para este WBS hasta la fecha de corte
                const pvWBS = (def.bac / duracionDias) * t;
                def.pv = Math.max(0, Math.min(def.bac, pvWBS));
                totalPV += def.pv;

                // Filtrar logs de costos y producción para este WBS que estén en el rango (inicio hasta fecha de corte)
                const logsMO = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_MO" && l.fecha && l.fecha <= corteDateStr);
                const logsMAT = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_MAT" && l.fecha && l.fecha <= corteDateStr);
                const logsEQP = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_EQP" && l.fecha && l.fecha <= corteDateStr);
                const logsSUB = dbLogs.filter(l => l.wbs === code && l.tipo === "AC_SUB" && l.fecha && l.fecha <= corteDateStr);
                const logsEV = dbLogs.filter(l => l.wbs === code && l.tipo === "EV_PRODUCCION" && l.fecha && l.fecha <= corteDateStr);

                const acMO = logsMO.reduce((sum, l) => sum + l.costo, 0);
                const acMAT = logsMAT.reduce((sum, l) => sum + l.costo, 0);
                const acEQP = logsEQP.reduce((sum, l) => sum + l.costo, 0);
                const acSUB = logsSUB.reduce((sum, l) => sum + l.costo, 0);

                const acTotal = acMO + acMAT + acEQP + acSUB;
                const evTotal = logsEV.reduce((sum, l) => sum + l.costo, 0);

                totalAC += acTotal;
                totalEV += evTotal;

                const cpi = acTotal > 0 ? evTotal / acTotal : 1.0;
                const spi = def.pv > 0 ? evTotal / def.pv : 1.0;
                const eac = cpi > 0 ? acTotal + (def.bac - evTotal) / cpi : def.bac;
                const dif = evTotal - acTotal;

                wbsResMap[code] = { code, name: def.name, bac: def.bac, pv: def.pv, ev: evTotal, ac: acTotal, acMO, acMAT, acEQP, acSUB, dif, cpi, spi, eac };
            });
        }

        const cpiGlobal = totalAC > 0 ? totalEV / totalAC : 1.0;
        const spiGlobal = totalPV > 0 ? totalEV / totalPV : 1.0;
        const eacGlobal = cpiGlobal > 0 ? totalAC + (totalBAC - totalEV) / cpiGlobal : totalBAC;
        const desvioGlobal = totalBAC - eacGlobal;

        return { totalBAC, totalPV, totalEV, totalAC, cpiGlobal, spiGlobal, eacGlobal, desvioGlobal, wbsResMap };
    }

    function buildWBSBaseline() {
        const definitions = {};
        wbsOptions.forEach(w => {
            definitions[w.code] = {
                name: w.name.replace(/^WBS-\d+:\s*/, ''),
                bac: 0,
                pv: 0
            };
        });

        (presupuestoData.analisis_precios_unitarios || []).forEach(apu => {
            const wbs = apu.item === '01.01' ? 'WBS-100' :
                apu.item.startsWith('01.02') ? 'WBS-200' :
                apu.item.startsWith('01.03') ? 'WBS-300' : 'WBS-400';
            if (definitions[wbs]) {
                definitions[wbs].bac += Number(apu.costo_total_partida_directo) || 0;
            }
        });

        const wbsSource = baseDatosData?.catalogos?.wbs || baseDatosData?.nodos_wbs_estructura || [];
        wbsSource.forEach(row => {
            const code = row.codigo || row.wbs_codigo || row.wbs;
            if (definitions[code]) {
                definitions[code].pv = Number(row.pv_acumulado_pen || row.pv_pen || row.pv || 0);
                if (row.nombre) definitions[code].name = row.nombre;
            }
        });

        return definitions;
    }

    function setupGlobalButtons() {
        const btnExport = document.getElementById('btn-export-db') || document.getElementById('btn-export-db-dash');
        if (btnExport) {
            btnExport.onclick = () => {
                const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(dbLogs, null, 2));
                const downloadAnchor = document.createElement('a');
                downloadAnchor.setAttribute("href", dataStr);
                downloadAnchor.setAttribute("download", `RO_Logs_Unificados_${new Date().toISOString().split('T')[0]}.json`);
                document.body.appendChild(downloadAnchor);
                downloadAnchor.click();
                downloadAnchor.remove();
            };
        }

        const btnAnalizar = document.getElementById('btn-analizar-ia');
        if (btnAnalizar) {
            btnAnalizar.onclick = runAIGapAnalysis;
        }
    }

    function runAIGapAnalysis() {
        const evm = calculateEVMByWBS();
        const diagDiv = document.getElementById('ai-diag-text');
        const recDiv = document.getElementById('ai-rec-text');
        const alertWBS = Object.values(evm.wbsResMap).filter(w => w.cpi < 1.0 && w.ac > 0);

        if (alertWBS.length === 0) {
            diagDiv.innerHTML = `<p>✅ <strong>Estado Saludable por WBS:</strong> El Resultado Operativo en todos los nodos WBS se encuentra dentro del presupuesto meta.</p>`;
            recDiv.innerHTML = `<p>Mantener la disciplina de captura diaria por frentes WBS.</p>`;
        } else {
            let html = '<ul>';
            alertWBS.forEach(w => {
                html += `<li><strong>${w.code} (${w.name}):</strong> CPI = ${w.cpi.toFixed(2)}. EV Ganado: S/ ${w.ev.toLocaleString()} vs. Costo Incurrido: S/ ${w.ac.toLocaleString()}. Sobrecosto acumulado: S/ ${(w.ac - w.ev).toLocaleString()}.</li>`;
            });
            html += '</ul>';

            diagDiv.innerHTML = `<p>🚨 <strong>Brechas Detectadas a Nivel WBS Multi-Frente:</strong></p>${html}`;
            recDiv.innerHTML = `
                <ul>
                    <li><strong>Ajuste de Rendimientos HH en WBS-200 (Tareador):</strong> Revisar la distribución de cuadrillas de peones y oficiales por tramo.</li>
                    <li><strong>Control de Equipos en Stand-by (Administradora):</strong> Verificar horas máquina de excavadora reportadas en WBS-200.</li>
                </ul>
            `;
        }
    }

    // --- 6. CURVA S: PV acumulado, EV acumulado, AC acumulado por fecha ---
    let sCurveChartInstance = null;

    function renderSCurve(evm, fechaCorte) {
        const canvas = document.getElementById('chart-scurve');
        const noDataDiv = document.getElementById('scurve-no-data');
        if (!canvas) return;

        // Recopilar todas las fechas con registros
        const evLogs  = dbLogs.filter(l => l.tipo === 'EV_PRODUCCION' && l.fecha);
        const acLogs  = dbLogs.filter(l => ['AC_MO','AC_MAT','AC_EQP','AC_SUB'].includes(l.tipo) && l.fecha);

        // Si no hay ningún registro, mostrar placeholder
        if (evLogs.length === 0 && acLogs.length === 0) {
            canvas.style.display = 'none';
            if (noDataDiv) noDataDiv.style.display = 'block';
            if (sCurveChartInstance) { sCurveChartInstance.destroy(); sCurveChartInstance = null; }
            return;
        }
        canvas.style.display = 'block';
        if (noDataDiv) noDataDiv.style.display = 'none';

        // Recopilar todas las fechas únicas reales y ordenarlas
        const fechasRealesSet = new Set();
        [...evLogs, ...acLogs].forEach(l => fechasRealesSet.add(l.fecha));
        const fechasReales = [...fechasRealesSet].sort();

        const fechaInicioStr = fechasReales[0];
        const fechaLimite = fechaCorte || fechasReales[fechasReales.length - 1];

        const labelsDates = [];
        const dataPV = [], dataEV = [], dataAC = [];

        const usarSheetConsolidado = dbConsolidado.length > 0;

        if (usarSheetConsolidado) {
            // Obtener todas las fechas del consolidado del Google Sheet ordenadas
            const fechasProyecto = [...new Set(dbConsolidado.map(c => c.fecha))].sort();

            fechasProyecto.forEach((f) => {
                // Sumar el PV acumulado de todas las WBS para la fecha f
                const pvCons = dbConsolidado.filter(c => c.fecha === f).reduce((sum, c) => sum + (c.pv_acumulado || 0), 0);
                dataPV.push(parseFloat(pvCons.toFixed(2)));

                // Solo graficar EV y AC hasta la fecha límite seleccionada (fecha de corte)
                if (f <= fechaLimite) {
                    const evCons = dbConsolidado.filter(c => c.fecha === f).reduce((sum, c) => sum + (c.ev_acumulado || 0), 0);
                    const acCons = dbConsolidado.filter(c => c.fecha === f).reduce((sum, c) => sum + (c.ac_acumulado || 0), 0);
                    dataEV.push(parseFloat(evCons.toFixed(2)));
                    dataAC.push(parseFloat(acCons.toFixed(2)));
                } else {
                    dataEV.push(null);
                    dataAC.push(null);
                }

                // Formatear la etiqueta de fecha (YYYY-MM-DD -> DD/MM)
                const parts = f.split('-');
                labelsDates.push(parts.length === 3 ? `${parts[2]}/${parts[1]}` : f);
            });
        } else {
            // FALLBACK LOCAL: Si está offline
            const duracionDias = presupuestoData.proyecto?.duracion_dias_calendario || 60;
            const totalPV = evm.totalBAC;

            // Generar el rango completo de fechas del proyecto (ej. 60 días) para proyectar todo el PV
            const fechasProyecto = [];
            const inicioDate = new Date(fechaInicioStr + 'T00:00:00');
            for (let i = 0; i < duracionDias; i++) {
                const d = new Date(inicioDate);
                d.setDate(inicioDate.getDate() + i);
                const yyyy = d.getFullYear();
                const mm = String(d.getMonth() + 1).padStart(2, '0');
                const dd = String(d.getDate()).padStart(2, '0');
                fechasProyecto.push(`${yyyy}-${mm}-${dd}`);
            }

            fechasProyecto.forEach((f, i) => {
                // PV acumulado lineal para el día i + 1
                const pvAcumulado = Math.min(totalPV, (totalPV / duracionDias) * (i + 1));
                dataPV.push(parseFloat(pvAcumulado.toFixed(2)));

                // Solo graficar EV y AC hasta la fecha límite seleccionada (fecha de corte)
                if (f <= fechaLimite) {
                    // Acumular todos los registros reales de forma retroactiva hasta la fecha f
                    const evHastaFecha = evLogs.filter(l => l.fecha <= f).reduce((sum, l) => sum + (l.costo || 0), 0);
                    const acHastaFecha = acLogs.filter(l => l.fecha <= f).reduce((sum, l) => sum + (l.costo || 0), 0);
                    dataEV.push(parseFloat(evHastaFecha.toFixed(2)));
                    dataAC.push(parseFloat(acHastaFecha.toFixed(2)));
                } else {
                    // null en Chart.js corta la línea para indicar el periodo restante/futuro
                    dataEV.push(null);
                    dataAC.push(null);
                }

                // Formatear la etiqueta de fecha (YYYY-MM-DD -> DD/MM)
                const parts = f.split('-');
                labelsDates.push(parts.length === 3 ? `${parts[2]}/${parts[1]}` : f);
            });
        }

        // Destruir instancia anterior si existe
        if (sCurveChartInstance) {
            sCurveChartInstance.destroy();
            sCurveChartInstance = null;
        }

        const ctx = canvas.getContext('2d');

        // Gradientes con colores industriales de construcción
        const gradPV = ctx.createLinearGradient(0, 0, 0, 340);
        gradPV.addColorStop(0, 'rgba(29, 78, 216, 0.25)'); // Azul Marino Obra
        gradPV.addColorStop(1, 'rgba(29, 78, 216, 0.0)');

        const gradEV = ctx.createLinearGradient(0, 0, 0, 340);
        gradEV.addColorStop(0, 'rgba(234, 179, 8, 0.25)'); // Amarillo Caterpillar
        gradEV.addColorStop(1, 'rgba(234, 179, 8, 0.0)');

        const gradAC = ctx.createLinearGradient(0, 0, 0, 340);
        gradAC.addColorStop(0, 'rgba(234, 88, 12, 0.25)'); // Naranja de Seguridad
        gradAC.addColorStop(1, 'rgba(234, 88, 12, 0.0)');

        sCurveChartInstance = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labelsDates,
                datasets: [
                    {
                        label: 'PV — Valor Planificado (S/)',
                        data: dataPV,
                        borderColor: '#1d4ed8', // Azul Marino Obra
                        backgroundColor: gradPV,
                        borderWidth: 2.5,
                        pointBackgroundColor: '#1d4ed8',
                        pointRadius: 0, // No pintar puntos para una curva de 60 días limpia, solo en el hover
                        pointHoverRadius: 6,
                        fill: true,
                        tension: 0.2
                    },
                    {
                        label: 'EV — Valor Ganado (S/)',
                        data: dataEV,
                        borderColor: '#eab308', // Amarillo Caterpillar
                        backgroundColor: gradEV,
                        borderWidth: 2.5,
                        pointBackgroundColor: '#eab308',
                        pointRadius: 0,
                        pointHoverRadius: 6,
                        fill: true,
                        tension: 0.2
                    },
                    {
                        label: 'AC — Costo Real (S/)',
                        data: dataAC,
                        borderColor: '#ea580c', // Naranja de Seguridad
                        backgroundColor: gradAC,
                        borderWidth: 2.5,
                        pointBackgroundColor: '#ea580c',
                        pointRadius: 0,
                        pointHoverRadius: 6,
                        fill: true,
                        tension: 0.2
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        align: 'end',
                        labels: {
                            color: '#cbd5e1',
                            font: { family: 'Outfit', size: 12, weight: '500' },
                            boxWidth: 14,
                            boxHeight: 14,
                            padding: 20,
                            usePointStyle: true,
                            pointStyle: 'circle'
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        titleColor: '#e2e8f0',
                        bodyColor: '#94a3b8',
                        borderColor: 'rgba(148,163,184,0.15)',
                        borderWidth: 1,
                        padding: 12,
                        titleFont: { family: 'Outfit', size: 13, weight: '600' },
                        bodyFont: { family: 'Outfit', size: 12 },
                        callbacks: {
                            label: (context) => {
                                const val = context.parsed.y;
                                if (val === null || val === undefined) return null;
                                return ` ${context.dataset.label.split(' — ')[0]}: S/ ${val.toLocaleString('es-PE', { minimumFractionDigits: 2 })}`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { color: 'rgba(148,163,184,0.05)', drawBorder: false },
                        ticks: {
                            color: '#64748b',
                            font: { family: 'Outfit', size: 10 },
                            maxRotation: 45,
                            autoSkip: true,
                            maxTicksLimit: 12
                        }
                    },
                    y: {
                        grid: { color: 'rgba(148,163,184,0.05)', drawBorder: false },
                        ticks: {
                            color: '#64748b',
                            font: { family: 'Outfit', size: 11 },
                            callback: (val) => 'S/ ' + val.toLocaleString('es-PE', { minimumFractionDigits: 0 })
                        },
                        beginAtZero: true
                    }
                }
            }
        });
    }
});
