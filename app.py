"""
Gestor de Finanzas — Backend Flask
Clasifica transacciones y las escribe en el Excel automáticamente.
Sin API key: clasificación por reglas (gratis).
Con CLAUDE_API_KEY: clasificación inteligente con IA.
"""

from flask import Flask, request, jsonify, render_template_string
from datetime import datetime
import openpyxl, os, re

app = Flask(__name__)

EXCEL_PATH    = os.environ.get("EXCEL_PATH", "Finanzas_Miguel_2026.xlsx")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")

MESES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

# Transporte ANTES que Hogar para evitar que "gas" matchee "gasolina"
REGLAS = {
    "Comida":      ["mercadona","lidl","aldi","carrefour","dia","eroski","supermercado",
                    "fruteria","panaderia","pan","comida","cena","almuerzo","desayuno",
                    "cafe","bar","restaurante","pizza","burger","mcdonalds","kfc","kebab",
                    "delivery","glovo","uber eats","just eat","sushi","bocadillo",
                    "fruta","verdura","carne","pescado"],
    "Transporte":  ["gasolina","combustible","repsol","bp","cepsa","parking","aparcar",
                    "autobus","metro","cercanias","renfe","tren","taxi","uber","cabify",
                    "blablacar","peaje","itv","seguro coche","coche","moto","bici","patinete"],
    "Hogar":       ["alquiler","hipoteca","piso","casa","habitacion","luz","electricidad",
                    "agua","gas natural","factura gas","wifi","internet","fibra","movistar",
                    "vodafone","orange","comunidad","seguro hogar","ikea","leroy",
                    "bricolaje","mueble","reparacion"],
    "Salud":       ["farmacia","medicamento","medicina","doctor","medico","dentista",
                    "hospital","clinica","consulta","gym","gimnasio","proteina","suplemento",
                    "seguro medico","fisio","fisioterapia"],
    "Ocio":        ["netflix","spotify","amazon prime","disney","hbo","youtube premium",
                    "suscripcion","cine","teatro","concierto","viaje","vuelo","hotel",
                    "airbnb","booking","vacaciones","videojuego","steam","libro","amazon",
                    "fnac","copa","cerveza","discoteca","ocio"],
    "Inversiones": ["bolsa","accion","etf","fondo","inversion","crypto","bitcoin",
                    "ethereum","trading","degiro","ahorro","deposito","plan de pensiones",
                    "indexa","myinvestor","finizens"],
    "Trabajo":     ["material","herramienta","ordenador","portatil","software","licencia",
                    "curso","formacion","libro tecnico","office","adobe","asesoria",
                    "gestor","notario","autonomo","cuota autonomos"],
    "Ingresos":    ["salario","nomina","sueldo","paga","ingreso","transferencia recibida",
                    "freelance","cliente","factura cobrada","devolucion","reembolso",
                    "premio","bonus","extra"],
}

def clasificar_por_reglas(descripcion):
    desc = descripcion.lower()
    for categoria, palabras in REGLAS.items():
        for palabra in palabras:
            if re.search(r'\b' + re.escape(palabra) + r'\b', desc):
                return categoria, ("Ingreso" if categoria == "Ingresos" else "Gasto")
    for categoria, palabras in REGLAS.items():
        for palabra in palabras:
            if palabra in desc:
                return categoria, ("Ingreso" if categoria == "Ingresos" else "Gasto")
    return "Otros", "Gasto"

def clasificar_con_claude(descripcion):
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
        cats = ", ".join(list(REGLAS.keys()) + ["Otros"])
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001", max_tokens=50,
            messages=[{"role":"user","content":
                f'Clasifica esta transacción: "{descripcion}"\n'
                f'Categorías: {cats}\n'
                f'Responde SOLO: CATEGORIA|Ingreso o CATEGORIA|Gasto'}])
        resp = msg.content[0].text.strip()
        if "|" in resp:
            cat, tipo = resp.split("|", 1)
            cat, tipo = cat.strip(), tipo.strip()
            if cat in list(REGLAS.keys()) + ["Otros"] and tipo in ["Ingreso","Gasto"]:
                return cat, tipo
    except Exception:
        pass
    return clasificar_por_reglas(descripcion)

def clasificar(descripcion):
    if CLAUDE_API_KEY:
        return clasificar_con_claude(descripcion)
    return clasificar_por_reglas(descripcion)

def parsear_monto(texto):
    texto = texto.lower().replace(",", ".")
    nums = re.findall(r'\d+\.?\d*', texto)
    return float(nums[-1]) if nums else 0.0

def añadir_al_excel(descripcion, monto, categoria, tipo, notas=""):
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
    except FileNotFoundError:
        return -1
    ws = wb["Transacciones"]
    next_row = 5
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=2, max_col=2):
        for cell in row:
            if cell.value is not None:
                next_row = cell.row + 1
    hoy = datetime.now()
    ws[f"B{next_row}"] = descripcion
    ws[f"C{next_row}"] = monto
    ws[f"D{next_row}"] = hoy.strftime("%d/%m/%Y")
    ws[f"E{next_row}"] = categoria
    ws[f"F{next_row}"] = tipo
    ws[f"G{next_row}"] = MESES_ES[hoy.month]
    ws[f"H{next_row}"] = notas
    wb.save(EXCEL_PATH)
    return next_row


# ── HTML (mobile-first, una sola página) ──────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>💰 Mis Finanzas</title>
<style>
:root{--accent:#6366F1;--green:#22C55E;--red:#EF4444;--amber:#F59E0B;--bg:#F8FAFC;--card:#fff;--text:#1E293B;--muted:#64748B;--border:#E2E8F0}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:48px}
.header{background:linear-gradient(135deg,#1E293B,#334155);padding:20px 16px 24px;text-align:center;color:#fff}
.header h1{font-size:22px;font-weight:800}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.stats{display:grid;grid-template-columns:1fr 1fr 1fr;background:#0F172A;padding:0 16px 14px;gap:6px}
.stat{text-align:center;padding:10px 4px;border-radius:10px}
.stat .v{font-size:17px;font-weight:800;color:#fff}
.stat .l{font-size:9px;color:#94A3B8;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
.stat.g .v{color:#4ADE80}.stat.r .v{color:#F87171}.stat.b .v{color:#818CF8}
.card{margin:14px;background:var(--card);border-radius:18px;box-shadow:0 4px 20px rgba(0,0,0,.07);overflow:hidden}
.ch{padding:13px 16px;color:#fff;font-size:13px;font-weight:700;display:flex;align-items:center;gap:8px}
.ch.a{background:var(--accent)}.ch.d{background:#1E293B}
.ia{padding:14px}
label{font-size:11px;font-weight:700;color:var(--muted);display:block;margin-bottom:5px;text-transform:uppercase;letter-spacing:.3px}
.tipo{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:12px}
.tb{padding:11px;border:2px solid var(--border);border-radius:11px;background:#fff;font-size:13px;font-weight:700;cursor:pointer;text-align:center;color:var(--muted);transition:all .15s}
.tb.gasto.sel{border-color:var(--red);background:#FFF5F5;color:var(--red)}
.tb.ingreso.sel{border-color:var(--green);background:#F0FFF4;color:var(--green)}
.row2{display:grid;grid-template-columns:1fr 100px;gap:8px;margin-bottom:10px}
input{width:100%;padding:13px 14px;border:2px solid var(--border);border-radius:11px;font-size:16px;color:var(--text);background:var(--bg);outline:none;-webkit-appearance:none}
input:focus{border-color:var(--accent);background:#fff}
input::placeholder{color:#94A3B8}
.exs{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:12px}
.ex{padding:6px 11px;background:#EEF2FF;color:var(--accent);border:none;border-radius:20px;font-size:11px;font-weight:600;cursor:pointer}
.ex:active{background:#C7D2FE}
.cats{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:12px}
.cb{padding:8px 3px;border:2px solid var(--border);border-radius:9px;background:#fff;font-size:10px;font-weight:600;text-align:center;cursor:pointer;color:var(--muted);transition:all .15s}
.cb.sel{border-color:var(--accent);background:#EEF2FF;color:var(--accent)}
.cb span{display:block;font-size:15px;margin-bottom:2px}
.btn{width:100%;padding:15px;background:var(--accent);color:#fff;border:none;border-radius:13px;font-size:16px;font-weight:700;cursor:pointer;transition:background .2s,transform .1s}
.btn:active{background:#4F46E5;transform:scale(.98)}
.btn:disabled{background:#94A3B8;cursor:not-allowed}
.toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(100px);background:#1E293B;color:#fff;padding:12px 22px;border-radius:50px;font-size:13px;font-weight:600;transition:transform .3s;z-index:100;white-space:nowrap;box-shadow:0 8px 24px rgba(0,0,0,.3)}
.toast.show{transform:translateX(-50%) translateY(0)}
.toast.ok{background:#166534}.toast.err{background:#991B1B}
.txn{display:flex;align-items:center;padding:11px 14px;border-bottom:1px solid var(--border);gap:11px}
.txn:last-child{border-bottom:none}
.ti{width:38px;height:38px;border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0}
.td{flex:1;min-width:0}
.tn{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tm{font-size:11px;color:var(--muted);margin-top:1px}
.ta{font-size:15px;font-weight:800;flex-shrink:0}
.ta.g{color:var(--green)}.ta.r{color:var(--red)}
.empty{padding:28px 16px;text-align:center;color:var(--muted);font-size:13px}
.empty big{display:block;font-size:36px;margin-bottom:8px}
</style>
</head>
<body>
<div class="header"><h1>💰 Mis Finanzas</h1><p>Añade un gasto o ingreso · Se guarda en tu Excel</p></div>
<div class="stats">
  <div class="stat g"><div class="v" id="si">—</div><div class="l">Ingresos</div></div>
  <div class="stat r"><div class="v" id="sg">—</div><div class="l">Gastos</div></div>
  <div class="stat b"><div class="v" id="ss">—</div><div class="l">Saldo</div></div>
</div>
<div class="card">
  <div class="ch a">✏️ Nueva transacción</div>
  <div class="ia">
    <label>Tipo</label>
    <div class="tipo">
      <button class="tb gasto sel" onclick="setTipo('Gasto')">💸 Gasto</button>
      <button class="tb ingreso" onclick="setTipo('Ingreso')">💼 Ingreso</button>
    </div>
    <label>Descripción</label>
    <div class="row2">
      <input type="text" id="desc" placeholder="pan, gasolina, salario…" autocomplete="off" autocorrect="off">
      <input type="number" id="monto" placeholder="€" step="0.01" min="0">
    </div>
    <div class="exs">
      <button class="ex" onclick="q('pan',1)">🍞 Pan 1€</button>
      <button class="ex" onclick="q('mercadona',65)">🛒 Mercadona</button>
      <button class="ex" onclick="q('gasolina',48)">⛽ Gasolina</button>
      <button class="ex" onclick="q('gym',35)">🏋️ Gym</button>
      <button class="ex" onclick="q('netflix',12.99)">📺 Netflix</button>
      <button class="ex" onclick="q('salario',1500,true)">💼 Salario</button>
    </div>
    <label>Categoría <span style="color:#94A3B8;font-weight:400;text-transform:none">(opcional)</span></label>
    <div class="cats">
      <button class="cb" data-cat="Comida"      onclick="setCat(this)"><span>🍽️</span>Comida</button>
      <button class="cb" data-cat="Hogar"       onclick="setCat(this)"><span>🏠</span>Hogar</button>
      <button class="cb" data-cat="Transporte"  onclick="setCat(this)"><span>🚗</span>Transp.</button>
      <button class="cb" data-cat="Salud"       onclick="setCat(this)"><span>💊</span>Salud</button>
      <button class="cb" data-cat="Ocio"        onclick="setCat(this)"><span>🎉</span>Ocio</button>
      <button class="cb" data-cat="Inversiones" onclick="setCat(this)"><span>📈</span>Invers.</button>
      <button class="cb" data-cat="Trabajo"     onclick="setCat(this)"><span>💼</span>Trabajo</button>
      <button class="cb" data-cat="Otros"       onclick="setCat(this)"><span>📦</span>Otros</button>
    </div>
    <button class="btn" id="btn" onclick="enviar()">✅ Añadir al Excel</button>
  </div>
</div>
<br>
<div class="card">
  <div class="ch d">📋 Últimas transacciones <span id="cnt" style="font-size:10px;opacity:.6;margin-left:6px"></span></div>
  <div id="list"><div class="empty"><big>📊</big>Cargando…</div></div>
</div>
<div class="toast" id="toast"></div>
<script>
const ICONS={Comida:'🍽️',Hogar:'🏠',Transporte:'🚗',Salud:'💊',Ocio:'🎉',Inversiones:'📈',Trabajo:'💼',Ingresos:'💰',Otros:'📦'};
const BGS={Comida:'#DCFCE7',Hogar:'#DBEAFE',Transporte:'#FEF3C7',Salud:'#FCE7F3',Ocio:'#F3E8FF',Inversiones:'#CCFBF1',Trabajo:'#EEF2FF',Ingresos:'#D1FAE5',Otros:'#F1F5F9'};
let tipo='Gasto',cat=null;
function setTipo(t){tipo=t;document.querySelectorAll('.tb').forEach(b=>b.classList.remove('sel'));document.querySelector('.tb.'+t.toLowerCase()).classList.add('sel')}
function setCat(b){document.querySelectorAll('.cb').forEach(x=>x.classList.remove('sel'));if(cat===b.dataset.cat){cat=null}else{cat=b.dataset.cat;b.classList.add('sel')}}
function q(d,m,ing=false){document.getElementById('desc').value=d;document.getElementById('monto').value=m;if(ing)setTipo('Ingreso');else setTipo('Gasto')}
function fmt(n){return n>=1000?(n/1000).toFixed(1)+'k€':n.toFixed(0)+'€'}
function toast(msg,type){const t=document.getElementById('toast');t.textContent=msg;t.className='toast '+type;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),3000)}
async function stats(){try{const r=await fetch('/api/stats');const d=await r.json();document.getElementById('si').textContent=fmt(d.ingresos);document.getElementById('sg').textContent=fmt(d.gastos);const s=d.ingresos-d.gastos;document.getElementById('ss').textContent=(s>=0?'':'-')+fmt(Math.abs(s))}catch(e){}}
async function hist(){try{const r=await fetch('/api/transacciones?limit=15');const d=await r.json();const txns=d.transacciones||[];document.getElementById('cnt').textContent=txns.length+' recientes';if(!txns.length){document.getElementById('list').innerHTML='<div class="empty"><big>📊</big>¡Añade tu primera transacción!</div>';return}document.getElementById('list').innerHTML=txns.map(t=>`<div class="txn"><div class="ti" style="background:${BGS[t.categoria]||'#F1F5F9'}">${ICONS[t.categoria]||'📦'}</div><div class="td"><div class="tn">${t.descripcion}</div><div class="tm">${t.categoria} · ${t.fecha}</div></div><div class="ta ${t.tipo==='Ingreso'?'g':'r'}">${t.tipo==='Ingreso'?'+':'-'}${parseFloat(t.monto).toFixed(2)}€</div></div>`).join('')}catch(e){document.getElementById('list').innerHTML='<div class="empty"><big>⚠️</big>Sin conexión</div>'}}
async function enviar(){const desc=document.getElementById('desc').value.trim();const monto=parseFloat(document.getElementById('monto').value);if(!desc){toast('⚠️ Escribe una descripción','err');return}if(!monto||monto<=0){toast('⚠️ Escribe un monto válido','err');return}const btn=document.getElementById('btn');btn.disabled=true;btn.textContent='⏳ Guardando…';try{const body={descripcion:desc,monto,tipo};if(cat)body.categoria=cat;const r=await fetch('/api/transaccion',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});const d=await r.json();if(d.ok){toast(`✅ ${d.categoria} · ${monto.toFixed(2)}€`,'ok');document.getElementById('desc').value='';document.getElementById('monto').value='';cat=null;document.querySelectorAll('.cb').forEach(b=>b.classList.remove('sel'));await stats();await hist()}else toast('❌ '+(d.error||'Error'),'err')}catch(e){toast('❌ Sin conexión','err')}finally{btn.disabled=false;btn.textContent='✅ Añadir al Excel'}}
['desc','monto'].forEach(id=>document.getElementById(id).addEventListener('keypress',e=>{if(e.key==='Enter')enviar()}));
stats();hist();
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/api/transaccion", methods=["POST"])
def api_nueva():
    data = request.get_json()
    if not data:
        return jsonify({"ok": False, "error": "Sin datos"}), 400
    descripcion  = data.get("descripcion", "").strip()
    monto        = float(data.get("monto", 0))
    tipo_forzado = data.get("tipo")
    cat_forzada  = data.get("categoria")
    if not descripcion or monto <= 0:
        return jsonify({"ok": False, "error": "Datos inválidos"}), 400
    all_cats = list(REGLAS.keys()) + ["Otros"]
    if cat_forzada and cat_forzada in all_cats:
        categoria = cat_forzada
        tipo = tipo_forzado or ("Ingreso" if categoria == "Ingresos" else "Gasto")
    else:
        categoria, tipo = clasificar(descripcion)
        if tipo_forzado in ["Ingreso", "Gasto"]:
            tipo = tipo_forzado
            if tipo == "Ingreso":
                categoria = "Ingresos"
    fila = añadir_al_excel(descripcion, monto, categoria, tipo)
    return jsonify({"ok": True, "descripcion": descripcion, "monto": monto,
                    "categoria": categoria, "tipo": tipo, "fila": fila})

@app.route("/api/stats")
def api_stats():
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["Transacciones"]
        ingresos = gastos = 0.0
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
            if row[1] and isinstance(row[1], (int, float)):
                if row[4] == "Ingreso":
                    ingresos += row[1]
                else:
                    gastos += row[1]
        return jsonify({"ingresos": round(ingresos,2), "gastos": round(gastos,2)})
    except Exception as e:
        return jsonify({"ingresos": 0, "gastos": 0, "error": str(e)})

@app.route("/api/transacciones")
def api_transacciones():
    limit = int(request.args.get("limit", 20))
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["Transacciones"]
        txns = []
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
            desc, monto, fecha, cat, tipo = row[0], row[1], row[2], row[3], row[4]
            if desc and monto:
                txns.append({"descripcion": str(desc), "monto": float(monto),
                             "fecha": str(fecha) if fecha else "",
                             "categoria": str(cat) if cat else "Otros",
                             "tipo": str(tipo) if tipo else "Gasto"})
        txns.reverse()
        return jsonify({"transacciones": txns[:limit]})
    except Exception as e:
        return jsonify({"transacciones": [], "error": str(e)})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
