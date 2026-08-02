import unicodedata
import json
import os
import base64
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from flask import Flask, request, jsonify, render_template_string, Response
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

ALMACEN_PINS = {"mercedes": "2703", "bolivar": "0611"}

app = Flask(__name__)

FAVICON_B64 = "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"

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
TTL = 60

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
        header_idx = None
        for idx, row in enumerate(data[:10]):
            vals = [normalizar(str(c)) for c in row]
            if any("almacen" in v for v in vals):
                header_idx = idx
                break
        if header_idx is None:
            header_idx = 0
        headers = [normalizar(str(c)) for c in data[header_idx]]
        df = pd.DataFrame(data[header_idx+1:], columns=headers)
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
<link rel="icon" type="image/png" href="/favicon.png">
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

  .logo{display:block;max-width:180px;height:auto;margin:0 auto 14px;}
  .stripes{height:6px;background:repeating-linear-gradient(45deg,var(--amber) 0 14px,var(--carbon) 14px 28px);}
  header{padding:28px 20px 18px;border-bottom:1px solid var(--steel);text-align:center;}
  .eyebrow{font-size:12px;letter-spacing:.18em;text-transform:uppercase;color:var(--amber);font-weight:700;margin:0 0 6px;}
  h1{margin:0;font-size:26px;font-weight:800;}
  .sub{color:var(--muted);font-size:14px;margin-top:6px;}
  .almacen-wrap{max-width:640px;margin:0 auto;padding:16px 18px 0;}
  .almacen-label{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);margin-bottom:6px;}
  .almacen-row{display:flex;gap:12px;align-items:flex-start;}
  .almacen-row .col{flex:1;min-width:0;}
  #almacen,#ubic{width:100%;font-size:16px;padding:12px 14px;border-radius:8px;border:2px solid var(--steel);background:var(--panel);color:var(--paper);outline:none;appearance:none;-webkit-appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%238b8d92' fill='none' stroke-width='2'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 14px center;cursor:pointer;}
  #almacen:focus,#ubic:focus{border-color:var(--amber)}
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
  .pin-overlay{position:fixed;inset:0;z-index:300;background:var(--carbon);display:flex;align-items:center;justify-content:center;padding:18px;}
  .pin-overlay.hidden{display:none;}
  .pin-box{background:var(--panel);border:1px solid var(--steel);border-radius:12px;max-width:360px;width:100%;padding:32px 24px;text-align:center;}
  .pin-box h2{margin:0 0 6px;font-size:20px;color:var(--amber);}
  .pin-box p{color:var(--muted);font-size:14px;margin:0 0 20px;}
  .pin-box .pin-almacen{font-size:16px;font-weight:700;color:var(--paper);margin-bottom:16px;}
  .pin-inputs{display:flex;gap:10px;justify-content:center;margin-bottom:20px;}
  .pin-inputs input{width:48px;height:56px;text-align:center;font-size:24px;font-weight:700;border:2px solid var(--steel);border-radius:8px;background:var(--panel);color:var(--paper);outline:none;transition:border-color .15s ease;}
  .pin-inputs input:focus{border-color:var(--amber);}
  .pin-btn{width:100%;padding:14px;font-size:16px;font-weight:700;background:var(--amber);color:var(--carbon);border:none;border-radius:8px;cursor:pointer;transition:opacity .15s ease;}
  .pin-btn:hover{opacity:.85;}
  .pin-btn:disabled{opacity:.4;cursor:not-allowed;}
  .pin-error{color:var(--miss);font-size:13px;margin-top:10px;display:none;}
</style>
</head>
<body>

  <div class="pin-overlay" id="pinOverlay" style="display:none">
    <div class="pin-box">
      <h2>Ubica</h2>
      <p>Ingresá el código de acceso</p>
      <div class="pin-almacen" id="pinAlmacen"></div>
      <div class="pin-inputs">
        <input type="tel" maxlength="1" class="pin-digit" id="pd1" inputmode="numeric" autocomplete="off">
        <input type="tel" maxlength="1" class="pin-digit" id="pd2" inputmode="numeric" autocomplete="off">
        <input type="tel" maxlength="1" class="pin-digit" id="pd3" inputmode="numeric" autocomplete="off">
        <input type="tel" maxlength="1" class="pin-digit" id="pd4" inputmode="numeric" autocomplete="off">
      </div>
      <button class="pin-btn" id="pinBtn" disabled>Entrar</button>
      <div class="pin-error" id="pinError">Código incorrecto</div>
    </div>
  </div>

  <div class="stripes"></div>
  <header>
    <img src="data:image/png;base64,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" alt="Logo" class="logo">
    <p class="eyebrow">Consulta rápida de inventario</p>
    <h1>¿Dónde está?</h1>
    <p class="sub">{{ sub_text }}</p>
  </header>
  <div class="almacen-wrap">
    <div class="almacen-row">
      <div class="col">
        <div class="almacen-label">Almacén</div>
        <div class="almacen-fijo" id="almacenFijo" style="display:none"></div>
        <select id="almacen"><option value="">Cargando almacenes...</option></select>
      </div>
      <div class="col" id="ubicWrap" style="display:none">
        <div class="almacen-label">Ubicación</div>
        <select id="ubic"><option value="">Todas</option></select>
      </div>
    </div>
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
const ALMACEN_FIJO = {{ almacen_fijo_js | safe }};

(function(){
  if(ALMACEN_FIJO){
    const pinKey = 'ubica_pin_' + ALMACEN_FIJO;
    const savedPin = localStorage.getItem(pinKey);
    if(savedPin){
      iniciarApp();
    } else {
      mostrarPantallaPIN();
    }
  } else {
    iniciarApp();
  }
})();

function iniciarApp(){
  document.getElementById('pinOverlay').style.display = 'none';
  if(ALMACEN_FIJO){
    document.getElementById('almacenFijo').textContent = ALMACEN_FIJO;
    document.getElementById('almacenFijo').style.display = 'block';
    document.getElementById('almacen').style.display = 'none';
    document.querySelector('.almacen-label').textContent = 'Tu almacén';
    cargarFecha(ALMACEN_FIJO);
    cargarUbicaciones(ALMACEN_FIJO);
  } else {
    cargarAlmacenes();
  }
}

async function cargarUbicaciones(almacen){
  const wrap = document.getElementById('ubicWrap');
  const select = document.getElementById('ubic');
  select.innerHTML = '<option value="">Todas</option>';
  wrap.style.display = 'none';
  if(!almacen) return;
  try{
    const r = await fetch('/api/ubicaciones?almacen=' + encodeURIComponent(almacen));
    const d = await r.json();
    if(d.ubicaciones && d.ubicaciones.length){
      select.innerHTML = '<option value="">Todas</option>' + d.ubicaciones.map(u => '<option value="'+u+'">'+u+'</option>').join('');
      wrap.style.display = 'block';
    }
  }catch(e){}
}

async function cargarFecha(almacen){
  try{
    const r = await fetch('/api/fecha?almacen=' + encodeURIComponent(almacen));
    const d = await r.json();
    if(d.fecha) document.getElementById('fechaExport').textContent = d.fecha;
  }catch(e){}
}

function mostrarPantallaPIN(){
  const overlay = document.getElementById('pinOverlay');
  overlay.style.display = 'flex';
  document.getElementById('pinAlmacen').textContent = ALMACEN_FIJO;
  const digits = document.querySelectorAll('.pin-digit');
  digits[0].focus();
  digits.forEach((d, i) => {
    d.addEventListener('input', () => {
      if(d.value && i < 3) digits[i+1].focus();
      checkPin();
    });
    d.addEventListener('keydown', (e) => {
      if(e.key === 'Backspace' && !d.value && i > 0) digits[i-1].focus();
      if(e.key === 'Enter') submitPin();
    });
  });
}

function checkPin(){
  const digits = document.querySelectorAll('.pin-digit');
  const allFilled = Array.from(digits).every(d => d.value.length === 1);
  document.getElementById('pinBtn').disabled = !allFilled;
}

async function submitPin(){
  const digits = document.querySelectorAll('.pin-digit');
  const pin = Array.from(digits).map(d => d.value).join('');
  const btn = document.getElementById('pinBtn');
  btn.disabled = true;
  btn.textContent = 'Verificando...';
  try{
    const r = await fetch('/api/validate-pin', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({k: ALMACEN_FIJO, pin: pin})
    });
    const data = await r.json();
    if(data.valid){
      localStorage.setItem('ubica_pin_' + ALMACEN_FIJO, pin);
      iniciarApp();
    } else {
      document.getElementById('pinError').style.display = 'block';
      digits.forEach(d => { d.value = ''; d.style.borderColor = 'var(--miss)'; });
      setTimeout(() => {
        document.getElementById('pinError').style.display = 'none';
        digits.forEach(d => d.style.borderColor = '');
        digits[0].focus();
      }, 1500);
      btn.disabled = false;
      btn.textContent = 'Entrar';
    }
  }catch(e){
    btn.disabled = false;
    btn.textContent = 'Entrar';
    document.getElementById('pinError').textContent = 'Error de conexión';
    document.getElementById('pinError').style.display = 'block';
    setTimeout(() => document.getElementById('pinError').style.display = 'none', 2000);
  }
}

document.getElementById('pinBtn').addEventListener('click', submitPin);

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
  for(let intento = 0; intento < 5; intento++){
    try{
      const r = await fetch('/api/almacenes');
      const d = await r.json();
      if(d.almacenes && d.almacenes.length){
        almacenSelect.innerHTML = d.almacenes.map(a => '<option value="'+a+'">'+a+'</option>').join('');
        const saved = localStorage.getItem('ubica_almacen');
        if(saved && d.almacenes.includes(saved)) almacenSelect.value = saved;
        return;
      } else {
        almacenSelect.innerHTML = '<option value="">Sin almacenes disponibles</option>';
        return;
      }
    }catch(e){
      await new Promise(ok => setTimeout(ok, 3000));
    }
  }
  almacenSelect.innerHTML = '<option value="">Error cargando almacenes</option>';
}

almacenSelect.addEventListener('change', () => {
  localStorage.setItem('ubica_almacen', almacenSelect.value);
  if(almacenSelect.value) cargarFecha(almacenSelect.value);
  cargarUbicaciones(almacenSelect.value);
  const q = input.value.trim();
  if(q) buscar(q, 1);
});

document.getElementById('ubic').addEventListener('change', () => {
  const q = input.value.trim();
  buscar(q, 1);
  if(!q){
    status.textContent = 'Mostrando artículos de la ubicación seleccionada...';
  }
});

input.addEventListener('input', () => {
  clearTimeout(timer);
  currentPage = 1;
  const q = input.value.trim();
  if(!q){
    const ubic = document.getElementById('ubic').value;
    if(ubic){
      status.textContent = 'Mostrando artículos de la ubicación seleccionada...';
      timer = setTimeout(() => buscar(q, 1), 300);
    } else {
      resultados.innerHTML = '<p class="hint">Los resultados aparecerán aquí mientras escribes.</p>';
      status.textContent = '';
    }
    return;
  }
  status.textContent = 'Buscando...';
  timer = setTimeout(() => buscar(q, 1), 300);
});

async function buscar(q, page){
  page = page || 1;
  try{
    const almacen = getAlmacen();
    const ubic = document.getElementById('ubic').value;
    status.textContent = 'Buscando...';
    const res = await fetch('/api/buscar?q=' + encodeURIComponent(q) + '&page=' + page + '&almacen=' + encodeURIComponent(almacen) + '&ubic=' + encodeURIComponent(ubic));
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
  }catch(e){
    status.textContent = 'Conectando con el servidor...';
    setTimeout(() => buscar(q, page), 5000);
  }
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
    sub_text = "Busca un artículo en el almacén" if almacen_fijo else "Selecciona tu almacén y busca el artículo"
    return render_template_string(PAGINA, almacen_fijo_js=almacen_fijo_js, sub_text=sub_text)


@app.route("/api/health")
def api_health():
    return jsonify({"status": "ok"})


@app.route("/favicon.png")
def favicon():
    return Response(base64.b64decode(FAVICON_B64), mimetype="image/png")


@app.route("/api/fecha")
def api_fecha():
    almacen = request.args.get("almacen", "").strip()
    if not almacen:
        almacenes = get_almacenes()
        almacen = almacenes[0] if almacenes else ""
    if not almacen:
        return jsonify({"fecha": ""})
    try:
        sheet = get_sheet()
        ws = sheet.worksheet(almacen)
        data = ws.get_all_values()
        for i in range(min(5, len(data))):
            for cell in data[i]:
                cell = str(cell).strip()
                if cell and "-" in cell and ":" in cell:
                    parts = cell.split("-")
                    if len(parts) >= 2:
                        return jsonify({"fecha": parts[-1].strip()})
        return jsonify({"fecha": ""})
    except Exception as e:
        print(f"[ERROR] api_fecha: {e}")
        return jsonify({"fecha": ""})


@app.route("/api/validate-key")
def api_validate_key():
    clave = request.args.get("k", "").strip()
    if clave and clave in ALMACEN_KEYS:
        return jsonify({"valid": True, "almacen": ALMACEN_KEYS[clave]})
    return jsonify({"valid": False})


@app.route("/api/validate-pin", methods=["POST"])
def api_validate_pin():
    data = request.get_json()
    clave = data.get("k", "").strip()
    pin = data.get("pin", "").strip()
    if not clave or not pin:
        return jsonify({"valid": False})
    expected_pin = ALMACEN_PINS.get(clave.lower(), "")
    if expected_pin and pin == expected_pin and clave.lower() in ALMACEN_KEYS:
        return jsonify({"valid": True, "almacen": ALMACEN_KEYS[clave.lower()]})
    return jsonify({"valid": False})


@app.route("/api/almacenes")
def api_almacenes():
    return jsonify({"almacenes": get_almacenes()})


@app.route("/api/ubicaciones")
def api_ubicaciones():
    almacen = request.args.get("almacen", "").strip()
    if not almacen:
        return jsonify({"ubicaciones": []})
    df = get_dataframe(almacen)
    if df.empty or COL_UBICACION not in df.columns:
        return jsonify({"ubicaciones": []})
    ubics = sorted({str(v).strip() for v in df[COL_UBICACION] if str(v).strip()})
    return jsonify({"ubicaciones": ubics})


@app.route("/api/buscar")
def api_buscar():
    consulta = request.args.get("q", "").strip()
    almacen = request.args.get("almacen", "").strip()
    ubic = request.args.get("ubic", "").strip()
    if not consulta and not ubic:
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

    if ubic:
        mask = df[COL_UBICACION].apply(normalizar) == normalizar(ubic)
    else:
        mask = pd.Series([True] * len(df))
    if consulta:
        consulta_norm = normalizar(consulta)
        mask = mask & (df[COL_CODIGO].apply(normalizar).str.contains(consulta_norm, na=False) |
                       df[COL_DESCRIPCION].apply(normalizar).str.contains(consulta_norm, na=False))
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
