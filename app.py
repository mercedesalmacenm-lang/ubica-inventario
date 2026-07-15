import unicodedata
import json
import os
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from flask import Flask, request, jsonify, render_template_string
import time
import threading

SHEET_ID = os.environ.get("SHEET_ID", "1FLznJQ0PBxqnMNRPI_JEgv_QD7o7RoiodZLZmDzGE6Y")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

COL_CODIGO = "articulo"
COL_DESCRIPCION = "descripcion"
COL_UBICACION = "ubicacion"
COL_CANTIDAD = "existencia"

MAX_RESULTADOS = 20
PUERTO = int(os.environ.get("PORT", 8080))

# Claves por almacén: JSON mapping {"clave": "NOMBRE PESTAÑA"}
# Si está vacío, cualquiera puede ver cualquier almacén (modo admin)
ALMACEN_KEYS = {}
_keys_raw = os.environ.get("ALMACEN_KEYS", "")
if _keys_raw:
    try:
        ALMACEN_KEYS = json.loads(_keys_raw)
    except Exception:
        pass

app = Flask(__name__)

def get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_info = json.loads(creds_json)
        return Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    creds_file = os.environ.get("GOOGLE_CREDENTIALS_FILE", "./credentials.json")
    return Credentials.from_service_account_file(creds_file, scopes=SCOPES)

def get_sheet():
    creds = get_credentials()
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)

def normalizar(texto):
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))

_cache = {"almacenes": None, "ts": 0}
_cache_ws = {}
TTL = 300

def get_almacenes():
    now = time.time()
    if _cache["almacenes"] is not None and now - _cache["ts"] < TTL:
        return _cache["almacenes"]
    try:
        sheet = get_sheet()
        worksheets = sheet.worksheets()
        names = [ws.title for ws in worksheets if ws.title.lower() != "hoja de cálculo 1"]
        if not names:
            names = [ws.title for ws in worksheets]
        _cache["almacenes"] = names
        _cache["ts"] = now
        return names
    except Exception as e:
        print(f"[ERROR] get_almacenes: {e}")
        return []

def get_dataframe(almacen):
    now = time.time()
    if almacen in _cache_ws and now - _cache_ws[almacen]["ts"] < TTL:
        return _cache_ws[almacen]["df"]
    try:
        sheet = get_sheet()
        ws = sheet.worksheet(almacen)
        data = ws.get_all_values()
        if not data:
            return pd.DataFrame()
        df = pd.DataFrame(data[1:], columns=[normalizar(c) for c in data[0]])
        df = df.fillna("")
        _cache_ws[almacen] = {"df": df, "ts": now}
        return df
    except Exception as e:
        print(f"[ERROR] get_dataframe({almacen}): {e}")
        return pd.DataFrame()

PAGINA = """
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ubica | Consulta de inventario</title>
<style>
  :root{
    --carbon:#181a1d;
    --panel:#212327;
    --steel:#3a3d43;
    --amber:#f5a623;
    --amber-dim:#8a5f16;
    --paper:#ececea;
    --muted:#8b8d92;
    --ok:#5aa96a;
    --miss:#c1443c;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background:var(--carbon);
    color:var(--paper);
    font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
    min-height:100vh;
  }
  #splash{
    position:fixed;inset:0;z-index:9999;
    background:var(--carbon);
    display:flex;flex-direction:column;align-items:center;justify-content:center;
    transition:opacity .4s ease;
  }
  #splash.hide{opacity:0;pointer-events:none;}
  #splash .splash-stripes{position:absolute;top:0;left:0;right:0;height:6px;background:repeating-linear-gradient(45deg,var(--amber) 0 14px,var(--carbon) 14px 28px);}
  #splash .splash-title{font-size:28px;font-weight:800;margin:0 0 8px;}
  #splash .splash-sub{color:var(--muted);font-size:14px;margin:0 0 28px;}
  .splash-spinner{width:36px;height:36px;border:3px solid var(--steel);border-top-color:var(--amber);border-radius:50%;animation:spin .8s linear infinite;margin-bottom:18px;}
  @keyframes spin{to{transform:rotate(360deg);}}
  #splash .splash-msg{color:var(--muted);font-size:13px;text-align:center;max-width:300px;line-height:1.5;}
  #splash .splash-msg .err{color:var(--miss);margin-top:8px;display:none;}
  .stripes{height:6px;background:repeating-linear-gradient(45deg,var(--amber) 0 14px,var(--carbon) 14px 28px);}
  header{padding:28px 20px 18px;border-bottom:1px solid var(--steel);}
  .eyebrow{font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:var(--amber);font-weight:700;margin:0 0 6px;}
  h1{margin:0;font-size:26px;font-weight:800;}
  .sub{color:var(--muted);font-size:14px;margin-top:6px;}
  .almacen-wrap{max-width:640px;margin:0 auto;padding:16px 18px 0;}
  .almacen-label{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);margin-bottom:6px;}
  #almacen{width:100%;font-size:16px;padding:12px 14px;border-radius:8px;border:2px solid var(--steel);background:var(--panel);color:var(--paper);outline:none;appearance:none;-webkit-appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%238b8d92' fill='none' stroke-width='2'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 14px center;cursor:pointer;}
  #almacen:focus{border-color:var(--amber)}
  main{max-width:640px;margin:0 auto;padding:22px 18px 60px;}
  .search-wrap{position:relative;margin-bottom:22px;}
  #q{width:100%;font-size:19px;padding:16px 16px 16px 46px;border-radius:10px;border:2px solid var(--steel);background:var(--panel);color:var(--paper);outline:none;transition:border-color .15s ease;}
  #q:focus{border-color:var(--amber);}
  .search-wrap::before{content:"";position:absolute;left:16px;top:50%;width:16px;height:16px;transform:translateY(-50%);border:2px solid var(--muted);border-radius:50%;box-shadow:6px 6px 0 -3px var(--muted);}
  .status{font-size:13px;color:var(--muted);min-height:18px;margin-bottom:14px;}
  .pagination{display:flex;justify-content:center;align-items:center;gap:16px;margin-top:20px;padding:14px 0;}
  .pag-btn{background:var(--panel);color:var(--paper);border:1px solid var(--steel);border-radius:6px;padding:8px 18px;font-size:14px;cursor:pointer;transition:all .15s ease;font-family:inherit;}
  .pag-btn:hover:not([disabled]){border-color:var(--amber);color:var(--amber);}
  .pag-btn[disabled]{opacity:.35;cursor:not-allowed;}
  .pag-info{font-size:13px;color:var(--muted);min-width:60px;text-align:center;}
  .card{background:var(--panel);border:1px solid var(--steel);border-left:5px solid var(--amber);border-radius:8px;padding:14px 16px;margin-bottom:12px;}
  .card .cod{font-family:"SF Mono",Consolas,Menlo,monospace;font-size:13px;color:var(--amber);letter-spacing:.03em;}
  .card .desc{font-size:16px;font-weight:600;margin:2px 0 10px;}
  .meta{display:flex;gap:18px;flex-wrap:wrap;font-size:14px;}
  .meta div span{display:block;font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;}
  .meta .ubic{color:var(--ok);font-weight:700;font-size:16px;}
  .empty,.hint{color:var(--muted);font-size:14px;padding:20px 4px;text-align:center;}
  .err{color:var(--miss);}
  footer{text-align:center;color:var(--muted);font-size:12px;padding-bottom:30px;}
  .help-btn{position:fixed;bottom:22px;right:22px;width:50px;height:50px;border-radius:50%;background:var(--amber);color:var(--carbon);border:none;font-size:22px;font-weight:800;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.4);z-index:100;transition:transform .15s ease;}
  .help-btn:hover{transform:scale(1.08);}
  .modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:200;justify-content:center;align-items:center;padding:18px;}
  .modal-overlay.active{display:flex;}
  .modal{background:var(--panel);border:1px solid var(--steel);border-radius:12px;max-width:480px;width:100%;max-height:85vh;overflow-y:auto;padding:28px 24px;}
  .modal h2{margin:0 0 16px;font-size:20px;color:var(--amber);}
  .modal h3{margin:18px 0 8px;font-size:15px;color:var(--paper);}
  .modal p,.modal li{font-size:14px;color:var(--muted);line-height:1.6;margin:0 0 10px;}
  .modal ul{padding-left:20px;margin:0 0 10px;}
  .modal li{margin-bottom:6px;}
  .modal .close-btn{display:block;width:100%;padding:12px;margin-top:18px;background:var(--steel);color:var(--paper);border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;}
  .modal .close-btn:hover{background:var(--amber);color:var(--carbon);}
</style>
</head>
<body>
  <div id="splash">
    <div class="splash-stripes"></div>
    <div class="splash-spinner"></div>
    <p class="splash-title">Ubica</p>
    <p class="splash-sub">Consulta de inventario</p>
    <p class="splash-msg">Preparando el sistema...<span class="err" id="splashErr"></span></p>
  </div>
  <div class="stripes"></div>
  <header>
    <p class="eyebrow">Consulta rápida de inventario</p>
    <h1>¿Dónde está?</h1>
    <p class="sub">Selecciona tu almacén y busca el artículo</p>
  </header>
  <div class="almacen-wrap">
    <div class="almacen-label">Almacén</div>
    <div class="almacen-fijo" id="almacenFijo" style="display:none"></div>
    <select id="almacen"><option value="">Cargando almacenes...</option></select>
  </div>
  <main>
    <div class="search-wrap">
      <input id="q" type="text" placeholder="Ej. tornillo 1/4 o MF-1234" autofocus autocomplete="off">
    </div>
    <div class="status" id="status"></div>
    <div id="resultados"><p class="hint">Los resultados aparecerán aquí mientras escribes.</p></div>
  </main>
  <footer>Datos según el último export de BOSS · <span id="fechaExport">cargando...</span></footer>
  <button class="help-btn" id="helpBtn" title="Ayuda">?</button>
  <div class="modal-overlay" id="modal">
    <div class="modal">
      <h2>Cómo usar Ubica</h2>
      <h3>Seleccionar tu almacén</h3>
      <p>Elige tu almacén en el menú desplegable. Solo verás los artículos de tu almacén.</p>
      <h3>Buscar un artículo</h3>
      <p>Escribe en el buscador el código o nombre. Los resultados se muestran mientras escribes.</p>
      <ul>
        <li>Puedes buscar por código, descripción o parte del nombre.</li>
        <li>No importa mayúsculas o minúsculas.</li>
        <li>La <strong>ubicación</strong> en verde es donde buscar.</li>
      </ul>
      <h3>Mantener actualizado</h3>
      <p>Cada almacén ejecuta el monitor en su PC. Exportás desde BOSS, el monitor lo sube solo.</p>
      <button class="close-btn" id="closeModal">Entendido</button>
    </div>
  </div>
<script>
const ALMACEN_FIJO = {{ almacen_fijo_js }};
const splash = document.getElementById('splash');

function ocultarSplash(){
  splash.classList.add('hide');
  setTimeout(() => splash.remove(), 500);
  if(ALMACEN_FIJO){
    document.getElementById('almacenFijo').textContent = ALMACEN_FIJO;
    document.getElementById('almacenFijo').style.display = 'block';
    document.getElementById('almacen').style.display = 'none';
    document.querySelector('.almacen-label').textContent = 'Tu almacén';
  } else {
    cargarAlmacenes();
  }
}

async function esperarServidor(){
  for(let i = 0; i < 40; i++){
    try{ const r = await fetch('/api/health', {signal: AbortSignal.timeout(5000)}); if(r.ok) return true; }catch(e){}
    await new Promise(ok => setTimeout(ok, 2000));
  }
  return false;
}

(async function(){
  const ok = await esperarServidor();
  ocultarSplash();
})();

const input = document.getElementById('q');
const resultados = document.getElementById('resultados');
const status = document.getElementById('status');
const almacenSelect = document.getElementById('almacen');
let timer = null;
let currentPage = 1;
let currentQ = '';

function getAlmacen(){
  return ALMACEN_FIJO || almacenSelect.value;
}

async function cargarAlmacenes(){
  try{
    const r = await fetch('/api/almacenes');
    const d = await r.json();
    if(d.almacenes && d.almacenes.length){
      almacenSelect.innerHTML = d.almacenes.map(a => '<option value="'+a+'">'+a+'</option>').join('');
      const saved = localStorage.getItem('ubica_almacen');
      if(saved && d.almacenes.includes(saved)) almacenSelect.value = saved;
    } else {
      almacenSelect.innerHTML = '<option value="">Sin almacenes disponibles</option>';
    }
  }catch(e){
    almacenSelect.innerHTML = '<option value="">Error cargando</option>';
  }
}

almacenSelect.addEventListener('change', () => {
  localStorage.setItem('ubica_almacen', almacenSelect.value);
  const q = input.value.trim();
  if(q) buscar(q, 1);
});

input.addEventListener('input', () => {
  clearTimeout(timer);
  currentPage = 1;
  const q = input.value.trim();
  if(!q){ resultados.innerHTML = '<p class="hint">Los resultados aparecerán aquí mientras escribes.</p>'; status.textContent = ''; return; }
  status.textContent = 'Buscando...';
  timer = setTimeout(() => buscar(q, 1), 300);
});

async function buscar(q, page){
  page = page || 1;
  try{
    const almacen = getAlmacen();
    const res = await fetch('/api/buscar?q=' + encodeURIComponent(q) + '&page=' + page + '&almacen=' + encodeURIComponent(almacen));
    const data = await res.json();
    if(data.error){ status.textContent=''; resultados.innerHTML='<p class="empty err">'+data.error+'</p>'; return; }
    const items = data.resultados;
    if(!items.length){ status.textContent=''; resultados.innerHTML='<p class="empty">No encontré nada para "'+q+'"</p>'; return; }
    currentPage = data.page;
    currentQ = q;
    status.textContent = data.paginas > 1
      ? items.length+' de '+data.total+' resultados (página '+data.page+'/'+data.paginas+')'
      : items.length + ' resultado(s)';
    let html = items.map(it => '<div class="card"><div class="cod">'+it.codigo+'</div><div class="desc">'+it.descripcion+'</div><div class="meta"><div><span>Ubicación</span><div class="ubic">'+(it.ubicacion||'N/D')+'</div></div>'+(it.cantidad?'<div><span>Cantidad</span><div>'+it.cantidad+'</div></div>':'')+'</div></div>').join('');
    if(data.paginas > 1){
      html += '<div class="pagination"><button class="pag-btn" onclick="irPagina('+(currentPage-1)+')"'+(currentPage<=1?' disabled':'')+'>&#9664; Anterior</button><span class="pag-info">'+currentPage+' / '+data.paginas+'</span><button class="pag-btn" onclick="irPagina('+(currentPage+1)+')"'+(currentPage>=data.paginas?' disabled':'')+'>Siguiente &#9654;</button></div>';
    }
    resultados.innerHTML = html;
  }catch(e){ status.textContent=''; resultados.innerHTML='<p class="empty err">Error de conexión.</p>'; }
}

function irPagina(page){ buscar(currentQ, page); resultados.scrollIntoView({behavior:'smooth', block:'start'}); }

document.getElementById('helpBtn').addEventListener('click', () => document.getElementById('modal').classList.add('active'));
document.getElementById('closeModal').addEventListener('click', () => document.getElementById('modal').classList.remove('active'));
document.getElementById('modal').addEventListener('click', e => { if(e.target.id==='modal') e.target.classList.remove('active'); });
</script>
</body>
</html>
"""


@app.route("/")
def index():
    clave = request.args.get("k", "").strip()
    almacen_fijo = ""
    if clave and clave in ALMACEN_KEYS:
        almacen_fijo = ALMACEN_KEYS[clave]
    almacen_fijo_js = json.dumps(almacen_fijo) if almacen_fijo else '""'
    sub_text = "Busca el artículo de tu almacén" if almacen_fijo else "Selecciona tu almacén y busca el artículo"
    return render_template_string(PAGINA, almacen_fijo_js=almacen_fijo_js, sub_text=sub_text)


@app.route("/api/health")
def api_health():
    return jsonify({"status": "ok"})


@app.route("/api/validate-key")
def api_validate_key():
    clave = request.args.get("k", "").strip()
    if clave and clave in ALMACEN_KEYS:
        return jsonify({"valid": True, "almacen": ALMACEN_KEYS[clave]})
    return jsonify({"valid": False})


@app.route("/api/almacenes")
def api_almacenes():
    return jsonify({"almacenes": get_almacenes()})


@app.route("/api/buscar")
def api_buscar():
    consulta = request.args.get("q", "").strip()
    almacen = request.args.get("almacen", "").strip()
    if not consulta:
        return jsonify({"resultados": [], "total": 0})
    if not almacen:
        almacenes = get_almacenes()
        if almacenes:
            almacen = almacenes[0]
        else:
            return jsonify({"error": "No hay almacenes disponibles"})

    df = get_dataframe(almacen)
    if df.empty:
        return jsonify({"error": "No hay datos para este almacén"})

    if COL_CODIGO not in df.columns or COL_DESCRIPCION not in df.columns:
        return jsonify({"error": f"Columnas no encontradas. Disponibles: {list(df.columns)}"})

    consulta_norm = normalizar(consulta)
    mask = df[COL_CODIGO].apply(normalizar).str.contains(consulta_norm, na=False) | \
           df[COL_DESCRIPCION].apply(normalizar).str.contains(consulta_norm, na=False)
    filtrado = df[mask]

    total = len(filtrado)
    page = int(request.args.get("page", 1))
    start = (page - 1) * MAX_RESULTADOS
    paginado = filtrado.iloc[start:start + MAX_RESULTADOS]

    items = []
    for _, fila in paginado.iterrows():
        items.append({
            "codigo": fila.get(COL_CODIGO, ""),
            "descripcion": fila.get(COL_DESCRIPCION, ""),
            "ubicacion": fila.get(COL_UBICACION, ""),
            "cantidad": fila.get(COL_CANTIDAD, None) if COL_CANTIDAD in df.columns else None,
        })

    return jsonify({"resultados": items, "total": total, "page": page, "paginas": max(1, -(-total // MAX_RESULTADOS))})


threading.Thread(target=get_almacenes, daemon=True).start()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PUERTO, debug=False)
