import http.server
import socketserver
import json
import os
import threading
import traceback
import uuid
from urllib.parse import urlparse

PORT = 8080
LOCK = threading.Lock()
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DB_PATH = os.path.join(ROOT_DIR, 'docs', 'data', 'base_datos_reportabilidad.json')

class BackendSyncHandler(http.server.SimpleHTTPRequestHandler):
    def send_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors_headers()
        self.end_headers()

    def do_POST(self):
        parsed_path = urlparse(self.path).path
        
        if parsed_path.rstrip('/').endswith('/api/save-log'):
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length)
                raw_logs = json.loads(post_data.decode('utf-8'))
                with LOCK:
                    if os.path.exists(DB_PATH):
                        with open(DB_PATH, 'r', encoding='utf-8') as f:
                            db_data = json.load(f)
                    else:
                        db_data = {"metadata": {}, "catalogos": {}, "presupuesto_apu": [], "registros_diarios": []}

                    if not isinstance(raw_logs, list):
                        raw_logs = [raw_logs]

                    existing_ids = {
                        str(log.get('id_registro') or log.get('id') or '')
                        for log in db_data.get('registros_diarios', [])
                    }
                    inserted_count = 0

                    for r in raw_logs:
                        normalized_log = {
                            "id_registro": str(r.get('id_registro') or r.get('id') or f"RO-{uuid.uuid4()}"),
                            "id_parte": r.get('id_parte') or '',
                            "fecha": r.get('fecha') or '',
                            "rol": r.get('rol') or r.get('emisor_rol') or 'Tareador (Bildin)',
                            "wbs_codigo": r.get('wbs_codigo') or r.get('wbs') or 'WBS-200',
                            "codigo_recurso_partida": r.get('codigo_recurso_partida') or r.get('codigoRecurso') or 'MO_OPERARIO',
                            "descripcion": r.get('descripcion') or r.get('detalle') or '',
                            "cantidad": float(r.get('cantidad', 0)),
                            "unidad": r.get('unidad') or '',
                            "costo_unitario": float(r.get('costo_unitario') or r.get('pu') or 0),
                            "costo_total_pen": float(r.get('costo_total_pen') or r.get('costo') or 0),
                            "categoria_evm": r.get('categoria_evm') or r.get('tipo') or 'AC_MO',
                            "origen_html": r.get('origen_html') or '',
                            "estado_validacion": 'VALIDO',
                            "fecha_recepcion": r.get('fecha_recepcion') or ''
                        }
                        if normalized_log["id_registro"] in existing_ids:
                            continue
                        db_data.setdefault('registros_diarios', []).append(normalized_log)
                        existing_ids.add(normalized_log["id_registro"])
                        inserted_count += 1

                    with open(DB_PATH, 'w', encoding='utf-8') as f:
                        json.dump(db_data, f, indent=2, ensure_ascii=False)

                self.send_response(200)
                self.send_cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                
                msg = "Registro(s) guardado(s) en la base JSON canónica. El Excel se genera desde scripts/generar_reportabilidad_excel.mjs."

                response_payload = {
                    "status": "success",
                    "message": msg,
                    "excel_updated": False,
                    "total_registros_db": len(db_data.get('registros_diarios', []))
                }
                self.wfile.write(json.dumps(response_payload).encode('utf-8'))

            except Exception as e:
                err_text = traceback.format_exc()
                with open("server_error.log", "w", encoding="utf-8") as ef:
                    ef.write(err_text)
                self.send_response(500)
                self.send_cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                err_payload = {"status": "error", "message": str(e), "traceback": err_text}
                self.wfile.write(json.dumps(err_payload).encode('utf-8'))
        else:
            self.send_response(404)
            self.send_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            err_payload = {"status": "error", "message": "Endpoint POST no encontrado"}
            self.wfile.write(json.dumps(err_payload).encode('utf-8'))

def run_server():
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), BackendSyncHandler) as httpd:
        print(f"Servidor Backend POST/GET activo en http://localhost:{PORT}")
        httpd.serve_forever()

if __name__ == "__main__":
    run_server()
