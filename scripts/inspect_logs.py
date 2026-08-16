import urllib.request, json

res = urllib.request.urlopen('https://script.google.com/macros/s/AKfycbxvQXdRmbwR20YkHhL6dbwnFG-ZqRm8-NQE65i7xXPVAW-GXE8_BT7LrDZTlevi_QQc/exec')
data = json.loads(res.read().decode('utf-8'))
wbs200_logs = [l for l in data.get('logs', []) if l.get('wbs') == 'WBS-200']
print('WBS-200 LOGS COUNT:', len(wbs200_logs))
for l in wbs200_logs:
    print(f"Code={l.get('codigoRecurso')}, Qty={l.get('cantidad')}, PU={l.get('pu')}, Costo={l.get('costo')}, Tipo={l.get('tipo')}, Detalle={l.get('detalle')}")
