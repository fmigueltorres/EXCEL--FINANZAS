# 🚀 Guía de Despliegue — Gestor de Finanzas

## Lo que tienes en esta carpeta
- `app.py` — la web app completa (backend + frontend en un solo archivo)
- `requirements.txt` — dependencias Python
- `render.yaml` — configuración para Render.com
- `Finanzas_Miguel_2026.xlsx` — tu Excel con datos de ejemplo

---

## PASO 1 — Instalar Git (si no lo tienes)
1. Ve a https://git-scm.com/download/win
2. Descarga e instala con opciones por defecto
3. Reinicia el ordenador

---

## PASO 2 — Subir el proyecto a GitHub
1. Ve a https://github.com y crea una cuenta (gratis)
2. Clic en **"New repository"** → nombre: `finanzas-personal` → **Create**
3. Abre la carpeta `finanzas_app` en el Explorador de archivos
4. Haz clic derecho en la carpeta → **"Git Bash Here"** (o abre PowerShell aquí)
5. Ejecuta estos comandos uno a uno:

```bash
git init
git add .
git commit -m "Gestor de finanzas inicial"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/finanzas-personal.git
git push -u origin main
```
*(Cambia TU_USUARIO por tu usuario de GitHub)*

---

## PASO 3 — Crear cuenta en Render.com y desplegar
1. Ve a https://render.com → **"Get Started for Free"**
2. Regístrate con tu cuenta de GitHub
3. Clic en **"New +"** → **"Web Service"**
4. Conecta tu repositorio `finanzas-personal`
5. Configuración:
   - **Name:** finanzas-miguel
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
6. Clic en **"Create Web Service"**
7. Espera 2-3 minutos mientras despliega
8. Te dará una URL tipo: `https://finanzas-miguel.onrender.com` ✅

---

## PASO 4 — Crear tu API Key de Claude (opcional pero recomendado)
> Sin API key la app funciona perfectamente con reglas. Con API key la clasificación es más inteligente.

1. Ve a https://console.anthropic.com
2. Crea una cuenta con tu email
3. Ve a **"API Keys"** → **"Create Key"**
4. Copia la key (empieza por `sk-ant-...`)
5. En Render.com → tu servicio → **"Environment"** → **"Add Environment Variable"**:
   - Key: `CLAUDE_API_KEY`
   - Value: `sk-ant-TU_KEY_AQUI`
6. Render desplegará automáticamente con la nueva variable

---

## PASO 5 — Añadir el Excel al servidor
> El Excel necesita estar en el servidor para que la app lo actualice.

**Opción A (más simple) — Google Drive sync:**
Por ahora la app funciona con el Excel incluido en el repositorio.
Para actualizarlo, simplemente haz `git push` con el Excel nuevo.

**Opción B — Persistencia real:**
Cuando quieras dar el siguiente paso, dile a Claude:
*"Quiero que mi Excel se guarde en Google Drive y la app lo actualice ahí"*

---

## PASO 6 — Usar desde el móvil
1. Abre tu URL de Render en Safari/Chrome del móvil
2. En iPhone: **Compartir → Añadir a pantalla de inicio**
3. En Android: **Menú → Añadir a pantalla de inicio**
4. ¡Ya tienes la app instalada como si fuera nativa! 📱

---

## Cómo actualizar la app en el futuro
Cuando quieras cambiar algo, modifica los archivos y ejecuta:
```bash
git add .
git commit -m "descripción del cambio"
git push
```
Render despliega automáticamente en 1-2 minutos.

---

## Solución de problemas comunes

| Problema | Solución |
|----------|----------|
| La app tarda en cargar | Normal en tier gratuito, el servidor "duerme". Espera 30 seg. |
| "Excel not found" | Asegúrate de que el .xlsx está en el repositorio |
| Clasificación errónea | Selecciona la categoría manualmente con los botones |
| Quiero añadir una categoría | Dile a Claude: "Añade la categoría X a mis reglas" |

---

## Estructura del proyecto
```
finanzas_app/
├── app.py                    ← toda la lógica (no tocar sin Claude)
├── requirements.txt          ← dependencias
├── render.yaml               ← config despliegue
├── Finanzas_Miguel_2026.xlsx ← tu Excel
└── GUIA_DESPLIEGUE.md        ← este archivo
```

---

*Creado con Claude · Mayo 2026*
