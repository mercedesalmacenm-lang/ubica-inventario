"""
Web de consulta de ubicaciones de inventario (BOSS)
-----------------------------------------------------
No requiere instalar nada en los celulares/PCs de tus compañeros: solo
abren un link en el navegador (Chrome, el que sea) dentro de la red de
la oficina/almacén.

Requisitos (instalar una sola vez en la PC que va a servir la web):
    pip install flask pandas openpyxl

Configura las variables en la sección "CONFIGURACIÓN" antes de correrlo.
"""

import unicodedata
import pandas as pd
from flask import Flask, request, jsonify, render_template_string
import urllib.request

# ======================= CONFIGURACIÓN =======================

# ID de tu hoja de Google Sheets (parte de la URL entre /d/ y /edit)
SHEET_ID = "1FLznJQ0PBxqnMNRPI_JEgv_QD7o7RoiodZLZmDzGE6Y"

# URL de exportación como CSV (no cambiar)
GOOGLE_SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# Nombres de columnas en tu Excel (ajusta si en tu export se llaman distinto)
COL_CODIGO = "articulo"
COL_DESCRIPCION = "descripcion"
COL_UBICACION = "ubicacion"
COL_CANTIDAD = "existencia"  # opcional, pon None si no existe

# Fila en la que están los encabezados (1 = primera fila, 2 = segunda, etc.)
# Si tu hoja tiene filas vacías o títulos antes de los encabezados, ajusta esto.
HEADER_ROW = 5

MAX_RESULTADOS = 20

# Puerto en el que corre la web (no lo cambies salvo que ya esté ocupado)
PUERTO = 8080

# ===============================================================

app = Flask(__name__)


def fecha_archivo():
    try:
        resp = urllib.request.urlopen(GOOGLE_SHEETS_URL)
        lineas = [l.decode("utf-8", errors="ignore").strip() for l in resp.readlines()[:10]]
        for linea in lineas:
            # Buscar una línea que contenga una fecha tipo "dd/Mes/yy hh:mm"
            import re
            m = re.search(r'(\d{1,2}/\w{3}/\d{2}\s+\d{1,2}:\d{2})', linea)
            if m:
                return m.group(1)
    except:
        pass
    return "desconocida"


def normalizar(texto: str) -> str:
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def cargar_inventario() -> pd.DataFrame:
    skip = list(range(1, HEADER_ROW)) if HEADER_ROW > 1 else None
    df = pd.read_csv(GOOGLE_SHEETS_URL, dtype=str, skiprows=skip)
    df.columns = [normalizar(c) for c in df.columns]
    return df.fillna("")


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
    letter-spacing:-0.01em;
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
  #q:focus{
    border-color:var(--amber);
  }
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
  #q:focus ~ .scanline{
    opacity:1;
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
  .err{
    color:var(--miss);
  }
  footer{
    text-align:center;
    color:var(--muted);
    font-size:12px;
    padding-bottom:30px;
  }
</style>
</head>
<body>
  <div class="stripes"></div>
  <header>
    <p class="eyebrow">Almacén Mercedes · Consulta rápida</p>
    <h1>¿Dónde está?</h1>
    <p class="sub">Escribe el código o el nombre del artículo</p>
  </header>
  <main>
    <div class="search-wrap">
      <input id="q" type="text" placeholder="Ej. tornillo 1/4 o MF-1234" autofocus autocomplete="off">
    </div>
    <div class="status" id="status"></div>
    <div id="resultados">
      <p class="hint">Los resultados aparecerán aquí mientras escribes.</p>
    </div>
  </main>
  <footer>Datos según el último export de BOSS · {{ fecha_export }}</footer>

<script>
const input = document.getElementById('q');
const resultados = document.getElementById('resultados');
const status = document.getElementById('status');
let timer = null;

input.addEventListener('input', () => {
  clearTimeout(timer);
  const q = input.value.trim();
  if(!q){
    resultados.innerHTML = '<p class="hint">Los resultados aparecerán aquí mientras escribes.</p>';
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
      resultados.innerHTML = '<p class="empty">No encontré nada para "' + q + '"</p>';
      return;
    }
    status.textContent = data.total > items.length
      ? 'Mostrando ' + items.length + ' de ' + data.total + ' resultados'
      : items.length + ' resultado(s)';
    resultados.innerHTML = items.map(it => `
      <div class="card">
        <div class="cod">${it.codigo}</div>
        <div class="desc">${it.descripcion}</div>
        <div class="meta">
          <div><span>Ubicación</span><div class="ubic">${it.ubicacion || 'N/D'}</div></div>
          ${it.cantidad !== null ? `<div><span>Cantidad</span><div>${it.cantidad}</div></div>` : ''}
        </div>
      </div>
    `).join('');
  }catch(e){
    status.textContent = '';
    resultados.innerHTML = '<p class="empty err">Error de conexión con el servidor.</p>';
  }
}
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(PAGINA, fecha_export=fecha_archivo())


@app.route("/api/buscar")
def api_buscar():
    consulta = request.args.get("q", "").strip()
    if not consulta:
        return jsonify({"resultados": [], "total": 0})

    try:
        df = cargar_inventario()
    except FileNotFoundError:
        return jsonify({"error": "No encuentro el archivo de inventario. Avisa al admin."})
    except Exception as e:
        return jsonify({"error": f"Error leyendo el inventario: {e}"})

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


if __name__ == "__main__":
    # host="0.0.0.0" permite que otras personas en la misma red (wifi/cable
    # de la oficina) accedan desde su celular usando la IP de esta PC.
    app.run(host="0.0.0.0", port=PUERTO, debug=False)
