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

ALMACEN_PINS = {"mercedes": "2703", "bolivar": "0611"}

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
    <p class="eyebrow">Consulta rápida de inventario</p>
    <h1>¿Dónde está?</h1>
    <p class="sub">{{ sub_text }}</p>
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
  } else {
    cargarAlmacenes();
  }
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
    status.textContent = 'Buscando...';
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
    sub_text = "Busca el artículo de tu almacén" if almacen_fijo else "Selecciona tu almacén y busca el artículo"
    return render_template_string(PAGINA, almacen_fijo_js=almacen_fijo_js, sub_text=sub_text)


@app.route("/api/health")
def api_health():
    return jsonify({"status": "ok"})


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
