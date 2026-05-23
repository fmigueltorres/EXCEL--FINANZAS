"""
Gestor de Finanzas — Backend Flask
Diseño: tonos marrones/beige, glassmorphism, hover effects, logo MN
"""

from flask import Flask, request, jsonify, render_template_string
from datetime import datetime
import openpyxl, os, re

app = Flask(__name__)

EXCEL_PATH     = os.environ.get("EXCEL_PATH", "Finanzas_Miguel_2026.xlsx")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")

MESES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

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


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>MN · Finanzas</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Inter:wght@300;400;500;600;700&display=swap');

:root {
  --bg1: #C8B8A2;
  --bg2: #B8A590;
  --bg3: #A89278;
  --glass: rgba(210,195,175,0.45);
  --glass2: rgba(235,225,210,0.55);
  --glass-border: rgba(255,248,238,0.5);
  --dark: #2C1F14;
  --dark2: #3D2E1E;
  --mid: #6B5240;
  --muted: #9A8470;
  --cream: #F5EEE4;
  --green: #4A7C5A;
  --green-bg: rgba(74,124,90,0.15);
  --red: #8B3A2E;
  --red-bg: rgba(139,58,46,0.15);
  --gold: #C4914A;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Inter', sans-serif;
  background: linear-gradient(145deg, #D4C4AE 0%, #C2AF97 40%, #B09878 100%);
  min-height: 100vh;
  padding-bottom: 60px;
  color: var(--dark);
}

/* ── Navbar ── */
.nav {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 16px 20px;
  background: rgba(44,31,20,0.12);
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  border-bottom: 1px solid var(--glass-border);
  position: sticky;
  top: 0;
  z-index: 50;
}

.logo {
  display: flex;
  align-items: center;
  gap: 10px;
}

.logo-svg { width: 42px; height: 42px; }

.logo-text {
  font-family: 'Playfair Display', serif;
  font-size: 18px;
  font-weight: 700;
  color: var(--dark);
  letter-spacing: 1px;
}

.nav-title {
  font-family: 'Inter', sans-serif;
  font-size: 11px;
  font-weight: 500;
  color: var(--mid);
  letter-spacing: 2px;
  text-transform: uppercase;
}

.nav-date {
  font-size: 11px;
  color: var(--muted);
}

/* ── Stats bar ── */
.stats {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 10px;
  padding: 16px;
}

.stat {
  background: var(--glass2);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  border: 1px solid var(--glass-border);
  border-radius: 16px;
  padding: 14px 10px;
  text-align: center;
  transition: transform 0.2s ease, box-shadow 0.2s ease;
  cursor: default;
}

.stat:hover {
  transform: translateY(-3px);
  box-shadow: 0 8px 24px rgba(44,31,20,0.15);
}

.stat .v {
  font-family: 'Playfair Display', serif;
  font-size: 18px;
  font-weight: 700;
  color: var(--dark);
}

.stat .l {
  font-size: 9px;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 1px;
  margin-top: 3px;
}

.stat.g .v { color: var(--green); }
.stat.r .v { color: var(--red); }
.stat.b .v { color: var(--gold); }

/* ── Cards ── */
.card {
  margin: 0 16px 16px;
  background: var(--glass);
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  border: 1px solid var(--glass-border);
  border-radius: 20px;
  overflow: hidden;
  box-shadow: 0 4px 24px rgba(44,31,20,0.1);
}

.ch {
  padding: 14px 18px;
  font-family: 'Playfair Display', serif;
  font-size: 14px;
  font-weight: 700;
  color: var(--cream);
  background: rgba(44,31,20,0.65);
  backdrop-filter: blur(10px);
  display: flex;
  align-items: center;
  gap: 8px;
  letter-spacing: 0.3px;
}

.ia { padding: 16px; }

/* ── Labels ── */
label {
  font-size: 10px;
  font-weight: 600;
  color: var(--muted);
  display: block;
  margin-bottom: 6px;
  text-transform: uppercase;
  letter-spacing: 1px;
}

/* ── Tipo toggle ── */
.tipo { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 14px; }

.tb {
  padding: 12px;
  border: 1.5px solid rgba(107,82,64,0.3);
  border-radius: 12px;
  background: rgba(235,225,210,0.3);
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  text-align: center;
  color: var(--mid);
  transition: all 0.2s ease;
  font-family: 'Inter', sans-serif;
}

.tb:hover {
  background: rgba(235,225,210,0.6);
  transform: translateY(-1px);
  box-shadow: 0 4px 12px rgba(44,31,20,0.12);
}

.tb.gasto.sel  { border-color: var(--red);   background: var(--red-bg);   color: var(--red); }
.tb.ingreso.sel { border-color: var(--green); background: var(--green-bg); color: var(--green); }

/* ── Inputs ── */
.row2 { display: grid; grid-template-columns: 1fr 95px; gap: 8px; margin-bottom: 12px; }

input {
  width: 100%;
  padding: 13px 15px;
  border: 1.5px solid rgba(107,82,64,0.25);
  border-radius: 12px;
  font-size: 15px;
  color: var(--dark);
  background: rgba(245,238,228,0.5);
  outline: none;
  -webkit-appearance: none;
  font-family: 'Inter', sans-serif;
  transition: all 0.2s ease;
}

input:focus {
  border-color: var(--gold);
  background: rgba(245,238,228,0.85);
  box-shadow: 0 0 0 3px rgba(196,145,74,0.15);
}

input::placeholder { color: rgba(107,82,64,0.5); }

/* ── Quick examples ── */
.exs { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 14px; }

.ex {
  padding: 6px 12px;
  background: rgba(196,145,74,0.15);
  color: var(--dark2);
  border: 1px solid rgba(196,145,74,0.3);
  border-radius: 20px;
  font-size: 11px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s ease;
  font-family: 'Inter', sans-serif;
}

.ex:hover {
  background: rgba(196,145,74,0.3);
  transform: translateY(-1px);
  box-shadow: 0 3px 8px rgba(44,31,20,0.12);
}

.ex:active { transform: scale(0.96); }

/* ── Category buttons ── */
.cats { display: grid; grid-template-columns: repeat(4,1fr); gap: 6px; margin-bottom: 14px; }

.cb {
  padding: 9px 3px;
  border: 1.5px solid rgba(107,82,64,0.2);
  border-radius: 11px;
  background: rgba(245,238,228,0.35);
  font-size: 10px;
  font-weight: 500;
  text-align: center;
  cursor: pointer;
  color: var(--mid);
  transition: all 0.2s ease;
  font-family: 'Inter', sans-serif;
}

.cb:hover {
  background: rgba(245,238,228,0.7);
  border-color: var(--gold);
  transform: translateY(-2px);
  box-shadow: 0 4px 10px rgba(44,31,20,0.12);
}

.cb.sel {
  border-color: var(--gold);
  background: rgba(196,145,74,0.2);
  color: var(--dark);
  font-weight: 700;
}

.cb span { display: block; font-size: 16px; margin-bottom: 3px; }

/* ── Submit button ── */
.btn {
  width: 100%;
  padding: 15px;
  background: rgba(44,31,20,0.75);
  color: var(--cream);
  border: 1px solid rgba(255,248,238,0.2);
  border-radius: 13px;
  font-size: 15px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.25s ease;
  font-family: 'Playfair Display', serif;
  letter-spacing: 0.5px;
  backdrop-filter: blur(10px);
}

.btn:hover {
  background: rgba(44,31,20,0.9);
  transform: translateY(-2px);
  box-shadow: 0 8px 20px rgba(44,31,20,0.25);
}

.btn:active { transform: scale(0.98); }
.btn:disabled { background: rgba(107,82,64,0.3); color: var(--muted); cursor: not-allowed; transform: none; }

/* ── Toast ── */
.toast {
  position: fixed;
  bottom: 24px;
  left: 50%;
  transform: translateX(-50%) translateY(120px);
  background: rgba(44,31,20,0.92);
  color: var(--cream);
  padding: 13px 24px;
  border-radius: 50px;
  font-size: 13px;
  font-weight: 500;
  transition: transform 0.3s cubic-bezier(0.34,1.56,0.64,1);
  z-index: 100;
  white-space: nowrap;
  box-shadow: 0 8px 32px rgba(44,31,20,0.3);
  border: 1px solid rgba(255,248,238,0.15);
  font-family: 'Inter', sans-serif;
}

.toast.show { transform: translateX(-50%) translateY(0); }
.toast.ok  { background: rgba(36,76,46,0.92); }
.toast.err { background: rgba(100,30,20,0.92); }

/* ── Transaction list ── */
.txn {
  display: flex;
  align-items: center;
  padding: 12px 16px;
  border-bottom: 1px solid rgba(107,82,64,0.12);
  gap: 12px;
  transition: background 0.2s ease;
  cursor: default;
}

.txn:hover { background: rgba(245,238,228,0.35); }
.txn:last-child { border-bottom: none; }

.ti {
  width: 40px; height: 40px;
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 18px;
  flex-shrink: 0;
  background: rgba(245,238,228,0.6);
  border: 1px solid rgba(196,145,74,0.2);
  transition: transform 0.2s ease;
}

.txn:hover .ti { transform: scale(1.08); }

.td { flex: 1; min-width: 0; }
.tn { font-size: 13px; font-weight: 600; color: var(--dark); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.tm { font-size: 11px; color: var(--muted); margin-top: 2px; }
.ta { font-size: 15px; font-weight: 700; flex-shrink: 0; font-family: 'Playfair Display', serif; }
.ta.g { color: var(--green); }
.ta.r { color: var(--red); }

.empty { padding: 32px 16px; text-align: center; color: var(--muted); font-size: 13px; }
.empty big { display: block; font-size: 38px; margin-bottom: 10px; }

/* ── Divider ── */
.divider { height: 1px; background: linear-gradient(to right, transparent, rgba(196,145,74,0.3), transparent); margin: 4px 16px; }
</style>
</head>
<body>

<!-- Navbar -->
<div class="nav">
  <div class="logo">
    <!-- Globe MN SVG -->
    <svg class="logo-svg" viewBox="0 0 80 80" fill="none" xmlns="http://www.w3.org/2000/svg">
      <circle cx="40" cy="40" r="34" stroke="#2C1F14" stroke-width="1.5" fill="none"/>
      <ellipse cx="40" cy="40" rx="34" ry="14" stroke="#2C1F14" stroke-width="1" fill="none" opacity="0.6"/>
      <ellipse cx="40" cy="40" rx="34" ry="24" stroke="#2C1F14" stroke-width="1" fill="none" opacity="0.4"/>
      <ellipse cx="40" cy="40" rx="20" ry="34" stroke="#2C1F14" stroke-width="1" fill="none" opacity="0.5"/>
      <ellipse cx="40" cy="40" rx="8" ry="34" stroke="#2C1F14" stroke-width="1" fill="none" opacity="0.4"/>
      <line x1="6" y1="40" x2="74" y2="40" stroke="#2C1F14" stroke-width="1" opacity="0.3"/>
      <line x1="40" y1="6" x2="40" y2="74" stroke="#2C1F14" stroke-width="1" opacity="0.3"/>
      <text x="40" y="47" font-family="Georgia, serif" font-size="20" font-weight="700" fill="#2C1F14" text-anchor="middle" letter-spacing="1">MN</text>
    </svg>
    <span class="logo-text">MN</span>
  </div>
  <span class="nav-title">Finanzas</span>
  <span class="nav-date" id="fecha"></span>
</div>

<!-- Stats -->
<div class="stats">
  <div class="stat g"><div class="v" id="si">—</div><div class="l">Ingresos</div></div>
  <div class="stat r"><div class="v" id="sg">—</div><div class="l">Gastos</div></div>
  <div class="stat b"><div class="v" id="ss">—</div><div class="l">Saldo</div></div>
</div>

<!-- Nueva transacción -->
<div class="card">
  <div class="ch">✦ &nbsp;Nueva transacción</div>
  <div class="ia">
    <label>Tipo de movimiento</label>
    <div class="tipo">
      <button class="tb gasto sel" onclick="setTipo('Gasto')">↓ Gasto</button>
      <button class="tb ingreso" onclick="setTipo('Ingreso')">↑ Ingreso</button>
    </div>

    <label>Descripción &amp; Monto</label>
    <div class="row2">
      <input type="text" id="desc" placeholder="pan, gasolina, salario…" autocomplete="off" autocorrect="off" spellcheck="false">
      <input type="number" id="monto" placeholder="€" step="0.01" min="0">
    </div>

    <div class="exs">
      <button class="ex" onclick="q('pan',1)">🍞 Pan</button>
      <button class="ex" onclick="q('mercadona',65)">🛒 Mercadona</button>
      <button class="ex" onclick="q('gasolina',48)">⛽ Gasolina</button>
      <button class="ex" onclick="q('gym',35)">🏋️ Gym</button>
      <button class="ex" onclick="q('netflix',12.99)">📺 Netflix</button>
      <button class="ex" onclick="q('salario',1500,true)">💼 Salario</button>
    </div>

    <label>Categoría <span style="text-transform:none;letter-spacing:0;font-weight:400;opacity:.7">(opcional · la detecta sola)</span></label>
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

    <button class="btn" id="btn" onclick="enviar()">Registrar movimiento</button>
  </div>
</div>

<div class="divider"></div>

<!-- Historial -->
<div class="card">
  <div class="ch">◈ &nbsp;Últimos movimientos <span id="cnt" style="font-size:10px;opacity:.5;margin-left:6px;font-family:Inter,sans-serif;font-weight:400"></span></div>
  <div id="list"><div class="empty"><big>◎</big>Cargando movimientos…</div></div>
</div>

<div class="toast" id="toast"></div>

<script>
const ICONS={Comida:'🍽️',Hogar:'🏠',Transporte:'🚗',Salud:'💊',Ocio:'🎉',Inversiones:'📈',Trabajo:'💼',Ingresos:'💰',Otros:'📦'};
let tipo='Gasto', cat=null;

// Fecha actual
const now=new Date();
document.getElementById('fecha').textContent=now.toLocaleDateString('es-ES',{day:'numeric',month:'short'});

function setTipo(t){
  tipo=t;
  document.querySelectorAll('.tb').forEach(b=>b.classList.remove('sel'));
  document.querySelector('.tb.'+t.toLowerCase()).classList.add('sel');
}

function setCat(b){
  document.querySelectorAll('.cb').forEach(x=>x.classList.remove('sel'));
  if(cat===b.dataset.cat){cat=null}else{cat=b.dataset.cat;b.classList.add('sel');}
}

function q(d,m,ing=false){
  document.getElementById('desc').value=d;
  document.getElementById('monto').value=m;
  if(ing)setTipo('Ingreso');else setTipo('Gasto');
}

function fmt(n){
  if(n===0)return '0€';
  return n>=1000?(n/1000).toFixed(1)+'k€':n.toFixed(0)+'€';
}

function toast(msg,type){
  const t=document.getElementById('toast');
  t.textContent=msg;t.className='toast '+type;t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),3200);
}

async function stats(){
  try{
    const r=await fetch('/api/stats');
    const d=await r.json();
    document.getElementById('si').textContent=fmt(d.ingresos);
    document.getElementById('sg').textContent=fmt(d.gastos);
    const s=d.ingresos-d.gastos;
    document.getElementById('ss').textContent=(s<0?'-':'')+fmt(Math.abs(s));
  }catch(e){}
}

async function hist(){
  try{
    const r=await fetch('/api/transacciones?limit=15');
    const d=await r.json();
    const txns=d.transacciones||[];
    document.getElementById('cnt').textContent=txns.length+' recientes';
    if(!txns.length){
      document.getElementById('list').innerHTML='<div class="empty"><big>◎</big>¡Añade tu primer movimiento!</div>';
      return;
    }
    document.getElementById('list').innerHTML=txns.map(t=>`
      <div class="txn">
        <div class="ti">${ICONS[t.categoria]||'📦'}</div>
        <div class="td">
          <div class="tn">${t.descripcion}</div>
          <div class="tm">${t.categoria} · ${t.fecha}</div>
        </div>
        <div class="ta ${t.tipo==='Ingreso'?'g':'r'}">${t.tipo==='Ingreso'?'+':'-'}${parseFloat(t.monto).toFixed(2)}€</div>
      </div>`).join('');
  }catch(e){
    document.getElementById('list').innerHTML='<div class="empty"><big>⚠</big>Sin conexión</div>';
  }
}

async function enviar(){
  const desc=document.getElementById('desc').value.trim();
  const monto=parseFloat(document.getElementById('monto').value);
  if(!desc){toast('Escribe una descripción','err');return;}
  if(!monto||monto<=0){toast('Escribe un monto válido','err');return;}
  const btn=document.getElementById('btn');
  btn.disabled=true;btn.textContent='Guardando…';
  try{
    const body={descripcion:desc,monto,tipo};
    if(cat)body.categoria=cat;
    const r=await fetch('/api/transaccion',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    const d=await r.json();
    if(d.ok){
      toast(`${d.categoria}  ·  ${monto.toFixed(2)}€  guardado`,'ok');
      document.getElementById('desc').value='';
      document.getElementById('monto').value='';
      cat=null;
      document.querySelectorAll('.cb').forEach(b=>b.classList.remove('sel'));
      await stats();await hist();
    }else toast('Error: '+(d.error||'desconocido'),'err');
  }catch(e){toast('Sin conexión','err');}
  finally{btn.disabled=false;btn.textContent='Registrar movimiento';}
}

['desc','monto'].forEach(id=>
  document.getElementById(id).addEventListener('keypress',e=>{if(e.key==='Enter')enviar();})
);

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
