import urllib.request, json, re

payload = [{
    'id': 'TEST-FINAL-SYNC-OK',
    'fecha': '2026-08-15',
    'rol': 'Ing. de Campo',
    'wbs': 'WBS-200',
    'codigoRecurso': '01.02.03',
    'detalle': 'PRUEBA FINAL COMPLETA DESPUES DE INCÓGNITO',
    'cantidad': 9,
    'unidad': 'M',
    'tipo': 'EV_PRODUCCION',
    'origen_html': 'ing_campo.html'
}]
req_post = urllib.request.Request('https://script.google.com/macros/s/AKfycbxvQXdRmbwR20YkHhL6dbwnFG-ZqRm8-NQE65i7xXPVAW-GXE8_BT7LrDZTlevi_QQc/exec',
                                 data=json.dumps(payload).encode('utf-8'),
                                 headers={'Content-Type': 'text/plain'})
try:
    res = urllib.request.urlopen(req_post)
    print('SUCCESS RESPONSE:', res.read().decode('utf-8'))
except urllib.error.HTTPError as e:
    print('ERROR CODE:', e.code)
    body = e.read().decode('utf-8')
    match = re.search(r'<div style="text-align:center[^"]*">(.*?)</div>', body)
    if match:
        print('ERROR MSG:', match.group(1))
    else:
        print('FULL BODY:', body[:500])
