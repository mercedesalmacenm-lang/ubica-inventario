import os
import re
import unicodedata
import pandas as pd
from flask import Flask, request, jsonify, render_template_string
import urllib.request
import gspread
from google.oauth2.service_account import Credentials
import openpyxl
from datetime import datetime

# ======================= CONFIGURACION =======================

SHEET_ID = os.environ.get("SHEET_ID", "1FLznJQ0PBxqnMNRPI_JEgv_QD7o7RoiodZLZmDzGE6Y")
GOOGLE_SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

COL_CODIGO = "articulo"
COL_DESCRIPCION = "descripcion"
COL_UBICACION = "ubicacion"
COL_CANTIDAD = "existencia"

HEADER_ROW = 5
MAX_RESULTADOS = 20

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ===============================================================

app = Flask(__name__)

def get_gspread():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_info = __import__("json").loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("./credentials.json", scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)

def is_boss_file(filename):
    name = filename.replace(".xlsx", "")
    return bool(re.match(r"^\d{14,16}$", name))

def normalizar(texto):
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))

def fecha_archivo():
    try:
        resp = urllib.request.urlopen(GOOGLE_SHEETS_URL)
        lineas = [l.decode("utf-8", errors="ignore").strip() for l in resp.readlines()[:10]]
        for linea in lineas:
            m = re.search(r"(\d{1,2}/\w{3}/\d{2}\s+\d{1,2}:\d{2})", linea)
            if m:
                return m.group(1)
    except:
        pass
    return "desconocida"

def cargar_inventario():
    skip = list(range(0, HEADER_ROW - 1)) if HEADER_ROW > 1 else None
    df = pd.read_csv(GOOGLE_SHEETS_URL, dtype=str, skiprows=skip)
    df.columns = [normalizar(c) for c in df.columns]
    return df.fillna("")

def subir_a_sheets(data):
    sheet = get_gspread()
    worksheets = sheet.worksheets()
    if worksheets:
        ws = worksheets[0]
        ws.clear()
        ws.update(range_name="A1", values=data)
    else:
        ws = sheet.add_worksheet(title="BOSS", rows=len(data)+10, cols=len(data[0]) if data else 10)
        ws.update(range_name="A1", values=data)
    return len(data)


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
  .stripes{
    height:6px;
    background:repeating-linear-gradient(45deg,var(--amber) 0 14px,var(--carbon) 14px 28px);
  }
  header{
    padding:28px 20px 18px;
    border-bottom:1px solid var(--steel);
  }
  .eyebrow{
    font-size:12px;
    letter-spacing:.18em;
    text-transform:uppercase;
    color:var(--amber);
    font-weight:700;
    margin:0 0 6px;
  }
  h1{
    margin:0;
    font-size:26px;
    font-weight:800;
  }
  .sub{
    color:var(--muted);
    font-size:14px;
    margin-top:6px;
  }
  main{
    max-width:640px;
    margin:0 auto;
    padding:22px 18px 60px;
  }
  .search-wrap{
    position:relative;
    margin-bottom:22px;
  }
  #q{
    width:100%;
    font-size:19px;
    padding:16px 16px 16px 46px;
    border-radius:10px;
    border:2px solid var(--steel);
    background:var(--panel);
    color:var(--paper);
    outline:none;
    transition:border-color .15s ease;
  }
  #q:focus{border-color:var(--amber);}
  .search-wrap::before{
    content:"";
    position:absolute;
    left:16px;
    top:50%;
    width:16px;
    height:16px;
    transform:translateY(-50%);
    border:2px solid var(--muted);
    border-radius:50%;
    box-shadow:6px 6px 0 -3px var(--muted);
  }
  .status{
    font-size:13px;
    color:var(--muted);
    min-height:18px;
    margin-bottom:14px;
  }
  .card{
    background:var(--panel);
    border:1px solid var(--steel);
    border-left:5px solid var(--amber);
    border-radius:8px;
    padding:14px 16px;
    margin-bottom:12px;
  }
  .card .cod{
    font-family:"SF Mono",Consolas,Menlo,monospace;
    font-size:13px;
    color:var(--amber);
    letter-spacing:.03em;
  }
  .card .desc{
    font-size:16px;
    font-weight:600;
    margin:2px 0 10px;
  }
  .meta{
    display:flex;
    gap:18px;
    flex-wrap:wrap;
    font-size:14px;
  }
  .meta div span{
    display:block;
    font-size:11px;
    color:var(--muted);
    text-transform:uppercase;
    letter-spacing:.06em;
  }
  .meta .ubic{
    color:var(--ok);
    font-weight:700;
    font-size:16px;
  }
  .empty, .hint{
    color:var(--muted);
    font-size:14px;
    padding:20px 4px;
    text-align:center;
  }
  .err{color:var(--miss);}
  footer{
    text-align:center;
    color:var(--muted);
    font-size:12px;
    padding-bottom:30px;
  }

  .upload-section{
    margin-top:30px;
    padding:20px;
    border:2px dashed var(--steel);
    border-radius:12px;
    text-align:center;
    transition:border-color .2s ease;
    cursor:pointer;
  }
  .upload-section:hover,.upload-section.dragover{
    border-color:var(--amber);
  }
  .upload-section h3{
    margin:0 0 8px;
    font-size:16px;
    color:var(--amber);
  }
  .upload-section p{
    margin:0;
    font-size:13px;
    color:var(--muted);
  }
  #fileInput{display:none;}
  .upload-status{
    margin-top:14px;
    font-size:14px;
    display:none;
  }
  .upload-status.ok{color:var(--ok);}
  .upload-status.error{color:var(--miss);}
  .upload-status.loading{color:var(--amber);}

  .nav-btns{
    display:flex;
    gap:10px;
    margin-top:24px;
    justify-content:center;
  }
  .nav-btns a{
    padding:10px 18px;
    border-radius:8px;
    text-decoration:none;
    font-size:14px;
    font-weight:600;
    border:2px solid var(--steel);
    color:var(--paper);
    background:var(--panel);
    transition:all .15s ease;
  }
  .nav-btns a:hover{
    border-color:var(--amber);
    color:var(--amber);
  }
  .nav-btns a.active{
    background:var(--amber);
    color:var(--carbon);
    border-color:var(--amber);
  }

  .help-btn{
    position:fixed;
    bottom:22px;
    right:22px;
    width:50px;
    height:50px;
    border-radius:50%;
    background:var(--amber);
    color:var(--carbon);
    border:none;
    font-size:22px;
    font-weight:800;
    cursor:pointer;
    box-shadow:0 4px 16px rgba(0,0,0,.4);
    z-index:100;
  }
  .help-btn:hover{transform:scale(1.08);}
  .modal-overlay{
    display:none;
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.65);
    z-index:200;
    justify-content:center;
    align-items:center;
    padding:18px;
  }
  .modal-overlay.active{display:flex;}
  .modal{
    background:var(--panel);
    border:1px solid var(--steel);
    border-radius:12px;
    max-width:480px;
    width:100%;
    max-height:85vh;
    overflow-y:auto;
    padding:28px 24px;
  }
  .modal h2{margin:0 0 16px;font-size:20px;color:var(--amber);}
  .modal h3{margin:18px 0 8px;font-size:15px;color:var(--paper);}
  .modal p,.modal li{font-size:14px;color:var(--muted);line-height:1.6;margin:0 0 10px;}
  .modal ul{padding-left:20px;margin:0 0 10px;}
  .modal li{margin-bottom:6px;}
  .modal .close-btn{
    display:block;width:100%;padding:12px;margin-top:18px;
    background:var(--steel);color:var(--paper);border:none;
    border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;
  }
  .modal .close-btn:hover{background:var(--amber);color:var(--carbon);}
</style>
</head>
<body>
  <div class="stripes"></div>
  <header>
    <p class="eyebrow">Almacen Mercedes</p>
    <h1>Ubica</h1>
    <p class="sub">Consulta de inventario</p>
  </header>
  <main>
    <div class="nav-btns">
      <a href="#" class="active" id="navBuscar">Buscar</a>
      <a href="#" id="navSubir">Subir BOSS</a>
    </div>

    <div id="panelBuscar">
      <div style="margin-top:22px">
        <div class="search-wrap">
          <input id="q" type="text" placeholder="Codigo o nombre del articulo" autofocus autocomplete="off">
        </div>
        <div class="status" id="status"></div>
        <div id="resultados">
          <p class="hint">Escribe para buscar.</p>
        </div>
      </div>
    </div>

    <div id="panelSubir" style="display:none; margin-top:22px;">
      <div class="upload-section" id="dropZone">
        <h3>Exportar de BOSS</h3>
        <p>Arrastra el archivo .xlsx aqui o haz clic para seleccionarlo</p>
        <p style="margin-top:8px; font-size:12px;">El nombre debe ser numerico (ej: 2026071221275882)</p>
        <input type="file" id="fileInput" accept=".xlsx">
      </div>
      <div class="upload-status" id="uploadStatus"></div>
    </div>

    <div class="nav-btns" style="margin-top:30px">
      <a href="#" id="helpBtn">Como usar</a>
    </div>
  </main>
  <footer>Datos segun el ultimo export de BOSS</footer>

  <div class="modal-overlay" id="modal">
    <div class="modal">
      <h2>Como usar Ubica</h2>
      <h3>Buscar</h3>
      <p>Escribe el codigo o nombre del articulo. Los resultados aparecen mientras escribes.</p>
      <h3>Actualizar inventario</h3>
      <p>Exporta el reporte desde BOSS como archivo .xlsx. Luego ve a la pestana <strong>Subir BOSS</strong> y selecciona el archivo.</p>
      <p>El archivo debe llamarse solo numeros (ej: 2026071221275882.xlsx).</p>
      <button class="close-btn" id="closeModal">Entendido</button>
    </div>
  </div>

<script>
const input = document.getElementById('q');
const resultados = document.getElementById('resultados');
const status = document.getElementById('status');
let timer = null;

input.addEventListener('input', () => {
  clearTimeout(timer);
  const q = input.value.trim();
  if(!q){
    resultados.innerHTML = '<p class="hint">Escribe para buscar.</p>';
    status.textContent = '';
    return;
  }
  status.textContent = 'Buscando...';
  timer = setTimeout(() => buscar(q), 300);
});

async function buscar(q){
  try{
    const res = await fetch('/api/buscar?q=' + encodeURIComponent(q));
    const data = await res.json();
    if(data.error){
      status.textContent = '';
      resultados.innerHTML = '<p class="empty err">' + data.error + '</p>';
      return;
    }
    const items = data.resultados;
    if(items.length === 0){
      status.textContent = '';
      resultados.innerHTML = '<p class="empty">No encontre nada para "' + q + '"</p>';
      return;
    }
    status.textContent = data.total > items.length
      ? 'Mostrando ' + items.length + ' de ' + data.total
      : items.length + ' resultado(s)';
    resultados.innerHTML = items.map(it => `
      <div class="card">
        <div class="cod">${it.codigo}</div>
        <div class="desc">${it.descripcion}</div>
        <div class="meta">
          <div><span>Ubicacion</span><div class="ubic">${it.ubicacion || 'N/D'}</div></div>
          ${it.cantidad !== null ? '<div><span>Cantidad</span><div>' + it.cantidad + '</div></div>' : ''}
        </div>
      </div>
    `).join('');
  }catch(e){
    status.textContent = '';
    resultados.innerHTML = '<p class="empty err">Error de conexion.</p>';
  }
}

document.getElementById('navBuscar').addEventListener('click', e => {
  e.preventDefault();
  document.getElementById('panelBuscar').style.display = '';
  document.getElementById('panelSubir').style.display = 'none';
  document.getElementById('navBuscar').classList.add('active');
  document.getElementById('navSubir').classList.remove('active');
  input.focus();
});
document.getElementById('navSubir').addEventListener('click', e => {
  e.preventDefault();
  document.getElementById('panelBuscar').style.display = 'none';
  document.getElementById('panelSubir').style.display = '';
  document.getElementById('navSubir').classList.add('active');
  document.getElementById('navBuscar').classList.remove('active');
});

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const uploadStatus = document.getElementById('uploadStatus');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  if(e.dataTransfer.files.length) subirArchivo(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => {
  if(fileInput.files.length) subirArchivo(fileInput.files[0]);
});

async function subirArchivo(file){
  uploadStatus.style.display = 'block';
  uploadStatus.className = 'upload-status loading';
  uploadStatus.textContent = 'Subiendo ' + file.name + '...';

  const formData = new FormData();
  formData.append('archivo', file);

  try{
    const res = await fetch('/api/subir', { method: 'POST', body: formData });
    const data = await res.json();
    if(data.error){
      uploadStatus.className = 'upload-status error';
      uploadStatus.textContent = data.error;
    }else{
      uploadStatus.className = 'upload-status ok';
      uploadStatus.textContent = 'OK: ' + data.filas + ' filas subidas desde ' + data.archivo;
    }
  }catch(e){
    uploadStatus.className = 'upload-status error';
    uploadStatus.textContent = 'Error de conexion.';
  }
}

document.getElementById('helpBtn').addEventListener('click', e => {
  e.preventDefault();
  document.getElementById('modal').classList.add('active');
});
document.getElementById('closeModal').addEventListener('click', () => {
  document.getElementById('modal').classList.remove('active');
});
document.getElementById('modal').addEventListener('click', e => {
  if(e.target.id === 'modal') document.getElementById('modal').classList.remove('active');
});
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(PAGINA)

@app.route("/api/buscar")
def api_buscar():
    consulta = request.args.get("q", "").strip()
    if not consulta:
        return jsonify({"resultados": [], "total": 0})
    try:
        df = cargar_inventario()
    except FileNotFoundError:
        return jsonify({"error": "No encuentro el archivo de inventario."})
    except Exception as e:
        return jsonify({"error": f"Error leyendo inventario: {e}"})
    if COL_CODIGO not in df.columns or COL_DESCRIPCION not in df.columns:
        return jsonify({"error": f"Columnas no coinciden. Encontradas: {list(df.columns)}"})
    consulta_norm = normalizar(consulta)
    mask = df[COL_CODIGO].apply(normalizar).str.contains(consulta_norm, na=False) | \
           df[COL_DESCRIPCION].apply(normalizar).str.contains(consulta_norm, na=False)
    filtrado = df[mask]
    items = []
    for _, fila in filtrado.head(MAX_RESULTADOS).iterrows():
        items.append({
            "codigo": fila.get(COL_CODIGO, ""),
            "descripcion": fila.get(COL_DESCRIPCION, ""),
            "ubicacion": fila.get(COL_UBICACION, ""),
            "cantidad": fila.get(COL_CANTIDAD, None) if COL_CANTIDAD in df.columns else None,
        })
    return jsonify({"resultados": items, "total": len(filtrado)})

@app.route("/api/subir", methods=["POST"])
def api_subir():
    if "archivo" not in request.files:
        return jsonify({"error": "No se envio ningun archivo."})
    archivo = request.files["archivo"]
    if not archivo.filename:
        return jsonify({"error": "Nombre de archivo vacio."})
    if not archivo.filename.endswith(".xlsx"):
        return jsonify({"error": "Solo se permiten archivos .xlsx"})
    if not is_boss_file(archivo.filename):
        return jsonify({"error": f"'{archivo.filename}' no tiene formato BOSS (debe ser solo numeros, ej: 2026071221275882.xlsx)"})

    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws_excel = wb.active
        data = []
        for row in ws_excel.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                data.append([str(cell) if cell is not None else "" for cell in row])
        wb.close()

        if not data:
            return jsonify({"error": "El archivo esta vacio."})

        filas = subir_a_sheets(data)
        return jsonify({"archivo": archivo.filename, "filas": filas})
    except Exception as e:
        return jsonify({"error": f"Error procesando: {e}"})

@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)), debug=False)
