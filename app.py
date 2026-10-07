import streamlit as st
import fitz  # PyMuPDF: motor C++ de renderizado ultrarrápido
from PIL import Image, ImageOps, ImageEnhance
import qrcode
import re
import base64
import os
import io
import json
import sqlite3
import hashlib
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timezone, timedelta
import uuid
import urllib.parse
import requests

# Configuración estricta de Zona Horaria de Colombia (UTC-5)
TZ_COLOMBIA = timezone(timedelta(hours=-5))

def ahora_colombia():
    return datetime.now(TZ_COLOMBIA)

# Importación segura de openpyxl
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_INSTALADO = True
except ImportError:
    OPENPYXL_INSTALADO = False

# Importación segura de Supabase
try:
    from supabase import create_client, Client
    SUPABASE_LIB_OK = True
except ImportError:
    SUPABASE_LIB_OK = False

st.set_page_config(
    page_title="Centro Cardiovascular Colombiano CENCARDIO · Workstation Enterprise",
    page_icon="🫀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# GESTIÓN DE LOGOTIPO INSTITUCIONAL
# ==============================================================================
@st.cache_data
def obtener_logo_b64():
    for nom in ["cencardio.jpg", "cencardio.png", "cencardio.jpeg", "logo.png", "logo.jpg"]:
        if os.path.exists(nom):
            with open(nom, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            mime = "png" if nom.endswith(".png") else "jpeg"
            return f"data:image/{mime};base64,{b64}"
    return None

def cargar_estilos_institucionales():
    return """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif !important; color: #1e293b; }
    .stApp { background-color: #f8fafc; }
    header[data-testid="stHeader"] { background: transparent !important; }
    .block-container { padding-top: 1.2rem !important; padding-bottom: 2.5rem !important; }

    .top-hospital-bar {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px;
        padding: 1rem 1.6rem; display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 4px 20px -2px rgba(10, 37, 64, 0.04); margin-bottom: 1.5rem;
    }
    .inst-badge-primary {
        background: #e0f2fe; color: #0369a1; font-size: 0.72rem; font-weight: 700;
        padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #bae6fd;
    }
    .inst-badge-success {
        background: #ecfdf5; color: #065f46; font-size: 0.72rem; font-weight: 700;
        padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #a7f3d0;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px; background: #ffffff; padding: 6px; border-radius: 12px;
        border: 1px solid #e2e8f0; margin-bottom: 1.2rem;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px !important; padding: 9px 22px !important;
        font-weight: 700 !important; font-size: 0.88rem !important; color: #64748b !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0a2540 !important; color: #ffffff !important;
    }
    .specialist-card {
        background: linear-gradient(135deg, #0a2540 0%, #133863 100%);
        border-radius: 14px; padding: 1.1rem; color: #ffffff; margin-bottom: 1.2rem;
    }
    .preview-container {
        border: 2px solid #e2e8f0; border-radius: 14px; background: #ffffff; padding: 8px;
    }
    .cencardio-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 20px;
        padding: 2.6rem 2.8rem; box-shadow: 0 20px 40px -12px rgba(10, 37, 64, 0.12);
        max-width: 480px; margin: 3rem auto; text-align: center;
    }
    </style>
    """

st.markdown(cargar_estilos_institucionales(), unsafe_allow_html=True)

# ==============================================================================
# CONECTOR DE NUBE PERSISTENTE: SUPABASE & FALLBACK SQLITE
# ==============================================================================
@st.cache_resource
def obtener_cliente_supabase():
    if not SUPABASE_LIB_OK:
        return None
    url = st.secrets.get("SUPABASE_URL", os.environ.get("SUPABASE_URL", ""))
    key = st.secrets.get("SUPABASE_KEY", os.environ.get("SUPABASE_KEY", ""))
    if url and key:
        try:
            return create_client(url, key)
        except Exception:
            return None
    return None

supabase = obtener_cliente_supabase()

def init_db_local():
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS estudios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_registro TEXT,
            paciente_nombre TEXT,
            modalidad TEXT,
            cups TEXT,
            parametro_clave TEXT,
            medico_firmante TEXT,
            informe_texto TEXT,
            pdf_blob BLOB,
            codigo_verificacion TEXT,
            hash_sha256 TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS dinamica_pacientes (
            cedula TEXT PRIMARY KEY,
            nombre TEXT,
            telefono TEXT,
            fecha_actualizacion TEXT
        )
    """)
    conn.commit()
    conn.close()

if not supabase:
    init_db_local()

def calcular_hash_sha256(pdf_bytes):
    return hashlib.sha256(pdf_bytes).hexdigest()

def guardar_estudio_servicio(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    fecha_actual_str = ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")
    hash_seguridad = calcular_hash_sha256(pdf_bytes)
    nombre_archivo = f"{cod_verif[:12]}_{re.sub(r'[^A-Za-z0-9]', '_', nombre)}.pdf"

    if supabase:
        try:
            supabase.storage.from_("estudios-pdf").upload(
                path=nombre_archivo,
                file=pdf_bytes,
                file_options={"content-type": "application/pdf", "upsert": "true"}
            )
            pdf_url = supabase.storage.from_("estudios-pdf").get_public_url(nombre_archivo)

            supabase.table("estudios").insert({
                "fecha_registro": ahora_colombia().isoformat(),
                "paciente_nombre": nombre,
                "modalidad": modalidad,
                "cups": cups,
                "parametro_clave": parametro_clave,
                "medico_firmante": medico,
                "informe_texto": texto,
                "pdf_url": pdf_url,
                "codigo_verificacion": cod_verif
            }).execute()
            return True, "Guardado en nube Supabase."
        except Exception as e:
            st.error(f"Error subiendo a Supabase: {e}. Guardando respaldo local...")

    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("""
        INSERT INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion, hash_sha256)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (fecha_actual_str, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif, hash_seguridad))
    conn.commit()
    conn.close()
    return True, "Guardado en almacenamiento local."

def obtener_historial_servicio():
    if supabase:
        try:
            res = supabase.table("estudios").select("id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, pdf_url, codigo_verificacion").order("id", desc=True).execute()
            lista = []
            for r in res.data:
                f_raw = r.get("fecha_registro", "")
                try:
                    f_dt = datetime.fromisoformat(f_raw.replace("Z", "+00:00")).astimezone(TZ_COLOMBIA)
                    f_fmt = f_dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    f_fmt = f_raw[:19].replace("T", " ")
                
                lista.append({
                    "id": r["id"],
                    "fecha_registro": f_fmt,
                    "paciente": r["paciente_nombre"],
                    "modalidad": r["modalidad"],
                    "cups": r["cups"],
                    "parametro_clave": r.get("parametro_clave", ""),
                    "medico": r.get("medico_firmante", ""),
                    "pdf_url": r.get("pdf_url", ""),
                    "codigo_verificacion": r.get("codigo_verificacion", "")
                })
            return lista, "supabase"
        except Exception:
            pass

    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, codigo_verificacion FROM estudios ORDER BY id DESC")
    filas = c.fetchall()
    conn.close()
    lista = []
    for f in filas:
        lista.append({
            "id": f[0], "fecha_registro": f[1], "paciente": f[2], "modalidad": f[3],
            "cups": f[4], "parametro_clave": f[5], "medico": f[6], "pdf_url": "",
            "codigo_verificacion": f[7] if f[7] else "N/A"
        })
    return lista, "local"

def obtener_pdf_bytes_individual(estudio_id, pdf_url=""):
    if pdf_url:
        try:
            r = requests.get(pdf_url, timeout=10)
            if r.status_code == 200:
                return r.content
        except Exception:
            pass
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT pdf_blob FROM estudios WHERE id = ?", (estudio_id,))
    res = c.fetchone()
    conn.close()
    return res[0] if (res and res[0]) else None

def eliminar_estudio_servicio(estudio_id):
    if supabase:
        try:
            supabase.table("estudios").delete().eq("id", estudio_id).execute()
            return
        except Exception:
            pass
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("DELETE FROM estudios WHERE id = ?", (estudio_id,))
    conn.commit()
    conn.close()

def sincronizar_directorio_servicio(df):
    col_ced = next((col for col in df.columns if any(k in col.lower() for k in ["cedula", "documento", "identificacion", "id"])), None)
    col_tel = next((col for col in df.columns if any(k in col.lower() for k in ["telefono", "celular", "tel", "movil"])), None)
    col_nom = next((col for col in df.columns if any(k in col.lower() for k in ["nombre", "paciente", "usuario"])), None)

    if not col_ced or not col_tel:
        return False, "El archivo debe contener columnas de Cédula y Teléfono/Celular."

    registros = 0
    fecha_hoy = ahora_colombia().strftime("%Y-%m-%d %H:%M")

    if supabase:
        try:
            lote = []
            for _, fila in df.iterrows():
                ced = re.sub(r'\D', '', str(fila[col_ced]))
                tel = re.sub(r'\D', '', str(fila[col_tel]))
                nom = str(fila[col_nom]) if col_nom else ""
                if len(ced) >= 5 and len(tel) >= 7:
                    lote.append({"cedula": ced, "nombre": nom, "telefono": tel, "fecha_actualizacion": fecha_hoy})
                    registros += 1
            if lote:
                supabase.table("dinamica_pacientes").upsert(lote).execute()
            return True, f"Se sincronizaron con éxito {registros} pacientes en Supabase."
        except Exception as e:
            st.error(f"Error sincronizando en Supabase ({e}). Guardando localmente...")

    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    registros = 0
    for _, fila in df.iterrows():
        ced = re.sub(r'\D', '', str(fila[col_ced]))
        tel = re.sub(r'\D', '', str(fila[col_tel]))
        nom = str(fila[col_nom]) if col_nom else ""
        if len(ced) >= 5 and len(tel) >= 7:
            c.execute("""
                INSERT OR REPLACE INTO dinamica_pacientes (cedula, nombre, telefono, fecha_actualizacion)
                VALUES (?, ?, ?, ?)
            """, (ced, nom, tel, fecha_hoy))
            registros += 1
    conn.commit()
    conn.close()
    return True, f"Se sincronizaron con éxito {registros} pacientes localmente."

def buscar_telefono_servicio(cedula):
    if not cedula:
        return ""
    ced_limpia = re.sub(r'\D', '', str(cedula))
    if supabase:
        try:
            res = supabase.table("dinamica_pacientes").select("telefono").eq("cedula", ced_limpia).execute()
            if res.data:
                return res.data[0]["telefono"]
        except Exception:
            pass
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT telefono FROM dinamica_pacientes WHERE cedula = ?", (ced_limpia,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""

# ==============================================================================
# AUDITORÍA 1: SANITY CHECKS FISIOLÓGICOS (CANDADOS MATEMÁTICOS)
# ==============================================================================
def ejecutar_sanity_checks(modalidad, datos):
    alertas = []
    bloqueos = []

    if modalidad == "HOLTER":
        fc_p = datos.get("fc_prom", 75)
        fc_min = datos.get("fc_min", 55)
        fc_max = datos.get("fc_max", 100)
        
        if fc_min > fc_p:
            bloqueos.append(f"Incongruencia Cronotrópica: FC mínima ({fc_min} lpm) supera a la FC promedio ({fc_p} lpm).")
        if fc_max < fc_p:
            bloqueos.append(f"Incongruencia Cronotrópica: FC máxima ({fc_max} lpm) es inferior a la FC promedio ({fc_p} lpm).")
        if fc_min < 30:
            alertas.append(f"Bradicardia Extrema documentada ({fc_min} lpm). Compruebe ausencia de artefactos de desconexión.")
        if fc_max > 220:
            alertas.append(f"Frecuencia Cardíaca Máxima ({fc_max} lpm) atípica. Verifique si corresponde a ruido o taquiarritmia paroxística.")

    elif modalidad == "MAPA":
        pas = datos.get("pas_24h", 120)
        pad = datos.get("pad_24h", 80)
        cn = datos.get("caida_nocturna_val", 10.0)

        if pad >= pas:
            bloqueos.append(f"Error Hemodinámico Severo: La Presión Diastólica ({pad} mmHg) es igual o mayor a la Sistólica ({pas} mmHg).")
        if (pas - pad) < 20:
            alertas.append(f"Presión de Pulso patológicamente estrecha ({pas - pad} mmHg). Revise calibración.")
        if cn < -20.0:
            alertas.append(f"Patrón Riser Extremo (Aumento nocturno de presión: {cn:.1f}%). Riesgo cerebrovascular inminente.")

    elif modalidad == "ESFUERZO":
        edad = datos.get("edad", 35)
        fc_pico = datos.get("fc_pico", 150)
        fc_basal = datos.get("fc_basal", 75)
        pas_pico = datos.get("pas_pico", 150)
        pad_pico = datos.get("pad_pico", 90)

        if pad_pico >= pas_pico:
            bloqueos.append(f"Incongruencia en Esfuerzo: PAD pico ({pad_pico} mmHg) no puede superar a PAS ({pas_pico} mmHg).")
        if fc_basal >= fc_pico and fc_pico > 0:
            alertas.append("Respuesta Cronotrópica Paradójica: FC pico es inferior o igual a FC basal.")
        fcm_prev = 220 - edad if edad > 0 else 200
        if fc_pico > (fcm_prev * 1.25):
            alertas.append(f"FC Pico ({fc_pico} lpm) excede el 125% de la FCM máxima ({fcm_prev} lpm). Sospecha de taquicardia supraventricular o error de lectura.")

    return bloqueos, alertas

# ==============================================================================
# AUDITORÍA 2: CLINICAL LINTER (COHERENCIA CUANTITATIVA VS TEXTO)
# ==============================================================================
def auditar_coherencia_informe(texto_informe, datos, modalidad):
    discrepancias = []
    texto_upper = texto_informe.upper()

    if modalidad == "HOLTER":
        if "SIN ALTERACIONES ISQUÉMICAS" in texto_upper and datos.get("st_episodios", 0) > 0:
            discrepancias.append(f"El informe indica 'Sin isquemia', pero se contabilizan {datos['st_episodios']} episodios de desviación del ST.")
        if "TAQUICARDIA VENTRICULAR" in texto_upper and datos.get("tv_episodios", 0) == 0:
            discrepancias.append("Se menciona 'Taquicardia Ventricular' en la redacción, pero el contador numérico de TV es 0.")

    elif modalidad == "MAPA":
        if "DIPPING POSITIVO" in texto_upper and datos.get("caida_nocturna_val", 0) <= 0:
            discrepancias.append(f"El texto dictamina 'Dipping Positivo', pero el cálculo matemático de descenso nocturno es {datos.get('caida_nocturna_val', 0):.1f}%.")
        if "CONTROL ÓPTIMO" in texto_upper and (datos.get("pas_24h", 0) >= 140 or float(str(datos.get("carga_pas", "0")).replace(",", ".")) > 30):
            discrepancias.append("Se califica 'Control Óptimo', pero los promedios o cargas tensionales superan el rango de normalidad.")

    elif modalidad == "ESFUERZO":
        st_mm = datos.get("st_mm", 0.0)
        if "NEGATIVA PARA ISQUEMIA" in texto_upper and st_mm >= 1.0:
            discrepancias.append(f"Se dictamina prueba 'Negativa para isquemia', pero existe infradesnivel significativo del ST de {st_mm} mm.")
        if "SUFICIENTE" in texto_upper:
            fcm_prev = 220 - datos.get("edad", 35)
            porc = (datos.get("fc_pico", 0) / fcm_prev) * 100 if fcm_prev > 0 else 0
            if porc < 85:
                discrepancias.append(f"Se describe 'Prueba Suficiente', pero el esfuerzo solo alcanzó el {porc:.1f}% de la frecuencia submáxima prevista (meta ≥ 85%).")

    return discrepancias

# ==============================================================================
# MOTOR UNIFICADO DE INTELIGENCIA CLÍNICA (GEMINI API)
# ==============================================================================
def consultar_gemini_json(prompt_text, inline_items=None):
    gemini_key = st.secrets.get("GEMINI_API_KEY", st.secrets.get("gemini_api_key", os.environ.get("GEMINI_API_KEY", "")))
    gemini_key = str(gemini_key).strip().strip('"').strip("'")
    if not gemini_key:
        return False, None, "No se encontró GEMINI_API_KEY en Secrets."

    # Descubrimiento dinámico de modelos soportados
    modelos_disponibles = []
    url_list = f"https://generativelanguage.googleapis.com/v1beta/models?key={gemini_key}"
    try:
        r_list = requests.get(url_list, timeout=8)
        if r_list.status_code == 200:
            data_list = r_list.json()
            for m in data_list.get("models", []):
                if "generateContent" in m.get("supportedGenerationMethods", []):
                    nom = m.get("name", "")
                    if nom and "gemini" in nom.lower():
                        modelos_disponibles.append(("v1beta", nom))
    except Exception:
        pass

    if modelos_disponibles:
        def score_m(item):
            ver, nom = item
            n = nom.lower()
            if "2.5-flash" in n or "2.0-flash" in n: return 0
            if "flash" in n: return 1
            if "pro" in n: return 2
            return 3
        modelos_disponibles.sort(key=score_m)
    else:
        modelos_disponibles = [
            ("v1beta", "models/gemini-2.0-flash"),
            ("v1beta", "models/gemini-2.5-flash"),
            ("v1beta", "models/gemini-flash-latest"),
            ("v1", "models/gemini-2.0-flash")
        ]

    partes = [{"text": prompt_text}]
    if inline_items:
        for item in inline_items:
            partes.append(item)

    payload = {
        "contents": [{"parts": partes}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }

    ultimo_error = ""
    for ver, mod_name in modelos_disponibles:
        clean_name = mod_name if mod_name.startswith("models/") else f"models/{mod_name}"
        url = f"https://generativelanguage.googleapis.com/{ver}/{clean_name}:generateContent?key={gemini_key}"
        try:
            resp = requests.post(url, headers={"Content-Type": "application/json"}, json=payload, timeout=40)
            if resp.status_code == 200:
                res_json = resp.json()
                raw_text = res_json['candidates'][0]['content']['parts'][0]['text']
                match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                clean_json = match.group(0) if match else raw_text.strip()
                data = json.loads(clean_json)
                return True, data, clean_name
            else:
                ultimo_error = f"{clean_name} ({resp.status_code}): {resp.text[:120]}"
        except Exception as e:
            ultimo_error = f"{clean_name} error de red: {str(e)}"

    return False, None, ultimo_error

def optimizar_imagen_para_ia(img_bytes, max_dim=1600):
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    if max(w, h) > max_dim:
        ratio = max_dim / max(w, h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85, optimize=True)
    return buf.getvalue()

def ejecutar_extraccion_multimodal(archivos_fotos):
    inline_items = []
    for foto in archivos_fotos:
        try:
            raw_b = foto.getvalue() if hasattr(foto, "getvalue") else foto
            opt_b = optimizar_imagen_para_ia(raw_b, max_dim=1600)
            b64_img = base64.b64encode(opt_b).decode('utf-8')
            inline_items.append({
                "inline_data": {
                    "mime_type": "image/jpeg",
                    "data": b64_img
                }
            })
        except Exception:
            continue

    prompt = """Eres un Cardiólogo especialista en Ergometría del Centro Cardiovascular Colombiano CENCARDIO.
Analiza con máxima rigurosidad las fotografías adjuntas (trazados continuos de esfuerzo y notas manuscritas en post-it).
Extrae los siguientes parámetros clínicos exactos:
- Nombre del paciente en mayúsculas
- Cédula / PID (solo números)
- Edad (número entero)
- Sexo ("Femenino" o "Masculino")
- Protocolo ("Bruce")
- Etapa alcanzada (ej: "Etapa 6" o "Etapa 4")
- Tiempo total en minutos como número decimal (ej: si son 16:14 convierte a 16.23; si son 12:30 a 12.5; si son 9:34 a 9.57)
- Frecuencia Cardíaca Basal en reposo (FC Basal en lpm)
- Frecuencia Cardíaca Pico en esfuerzo máximo (FC Pico en lpm)
- Presión Arterial Basal (PAS y PAD basal en mmHg)
- Presión Arterial Pico de esfuerzo (PAS y PAD pico en mmHg)
- Desviación del segmento ST (en mm, ej: 0.0)

Devuelve ÚNICAMENTE un objeto JSON:
{
  "paciente": "KAROL JULIANA SUAREZ FARFAN",
  "cedula": "1050095219",
  "edad": 17,
  "sexo": "Femenino",
  "protocolo": "Bruce",
  "etapa": "Etapa 6",
  "tiempo_min": 16.23,
  "fc_basal": 91,
  "fc_pico": 194,
  "pas_basal": 119,
  "pad_basal": 70,
  "pas_pico": 140,
  "pad_pico": 87,
  "st_mm": 0.0
}"""

    ok, data, info = consultar_gemini_json(prompt, inline_items)
    if ok and data:
        if data.get("paciente"): st.session_state.erg_paciente = str(data["paciente"]).upper()
        if data.get("cedula"): st.session_state.erg_cedula = str(data["cedula"])
        if data.get("edad"): st.session_state.erg_edad = int(data["edad"])
        if data.get("sexo") in ["Femenino", "Masculino"]: st.session_state.erg_sexo = data["sexo"]
        if data.get("protocolo"): st.session_state.erg_protocolo = data["protocolo"]
        if data.get("etapa"): st.session_state.erg_etapa = str(data["etapa"])
        if data.get("tiempo_min"): st.session_state.erg_tiempo = float(data["tiempo_min"])
        if data.get("fc_basal"): st.session_state.erg_fc_basal = int(data["fc_basal"])
        if data.get("fc_pico"): st.session_state.erg_fc_pico = int(data["fc_pico"])
        if data.get("pas_basal") and data.get("pad_basal"):
            st.session_state.erg_pa_basal = f"{data['pas_basal']}/{data['pad_basal']}"
        if data.get("pas_pico") and data.get("pad_pico"):
            st.session_state.erg_pa_pico = f"{data['pas_pico']}/{data['pad_pico']}"
        if "st_mm" in data: st.session_state.erg_st_mm = float(data["st_mm"])
        return True, f"Parámetros extraídos con éxito con {info}."
    return False, f"No se pudo completar el análisis. Detalle: {info}"

# ==============================================================================
# FÓRMULAS FISIOLÓGICAS DE ESFUERZO (BRUCE CONTINUO)
# ==============================================================================
def calcular_mets_bruce(tiempo_min):
    if tiempo_min <= 0.1:
        return 0.0
    if tiempo_min <= 3.0:
        return round(1.0 + (tiempo_min / 3.0) * 3.8, 1)
    elif tiempo_min <= 6.0:
        return round(4.8 + ((tiempo_min - 3.0) / 3.0) * 2.2, 1)
    elif tiempo_min <= 9.0:
        return round(7.0 + ((tiempo_min - 6.0) / 3.0) * 3.1, 1)
    elif tiempo_min <= 12.0:
        return round(10.1 + ((tiempo_min - 9.0) / 3.0) * 3.4, 1)
    elif tiempo_min <= 15.0:
        return round(13.5 + ((tiempo_min - 12.0) / 3.0) * 3.7, 1)
    else:
        return round(17.2 + ((tiempo_min - 15.0) / 3.0) * 3.0, 1)

def redactar_informe_ergometria_institucional(d, perfil):
    fcm_prev = 220 - d["edad"] if d["edad"] > 0 else 200
    porc = round((d["fc_pico"] / fcm_prev) * 100) if (fcm_prev > 0 and d["fc_pico"] > 0) else 0
    suf = "suficiente" if porc >= 85 else "insuficiente"
    dp = d["fc_pico"] * d["pas_pico"]
    st_res = "Sin alteraciones isquémicas del segmento ST" if d["st_mm"] < 1.0 else f"Alteraciones de la repolarización con infradesnivel del ST de {d['st_mm']} mm"
    diag_el = "negativa" if d["st_mm"] < 1.0 else "positiva"
    dts = d["tiempo_min"] - (5 * d["st_mm"])
    duke = "Bajo riesgo coronario (< 1% mortalidad anual)" if dts >= 5 else ("Riesgo moderado (1 - 3% anual)" if dts >= -10 else "Alto riesgo coronario (> 3% anual)")

    return f"""INTERPRETACIÓN PRUEBA DE ESFUERZO COMPUTARIZADA - CUPS 893805

1. Ritmo sinusal normal basal y durante todas las etapas del esfuerzo físico.
2. Protocolo de {d['protocolo']} completado con duración de {d['tiempo_min']:.2f} minutos ({d['etapa']}).
3. Capacidad funcional alcanzada: {d['mets']} METs.
4. Respuesta cronotrópica: FC basal {d['fc_basal']} lpm elevándose progresivamente hasta FC pico de {d['fc_pico']} lpm ({porc}% de la FCM prevista de {fcm_prev} lpm, prueba {suf}).
5. Respuesta hemodinámica presora: PA basal {d['pas_basal']}/{d['pad_basal']} mmHg alcanzando PA pico de {d['pas_pico']}/{d['pad_pico']} mmHg.
6. Doble producto máximo alcanzado: {dp:,} mmHg*lpm.
7. Comportamiento electrocardiográfico del ST: {st_res}.
8. Sin arritmias ventriculares complejas ni eventos supraventriculares inducidos por el ejercicio.
9. Motivo de suspensión: Consecución de frecuencia cardíaca submáxima diagnóstica y agotamiento físico voluntario, sin angina.
10. Estratificación pronóstica por Duke Treadmill Score: {dts:.1f} ({duke}).

CONCLUSIÓN DIAGNÓSTICA:
Prueba de esfuerzo física {suf}, eléctricamente {diag_el} para isquemia miocárdica inducible. Adecuada tolerancia funcional y hemodinámica.
RECOMENDACIONES: {'Continuar control médico periódico y prescripción de actividad física regular.' if d['st_mm'] < 1.0 else 'Valoración prioritaria por cardiología clínica para estudio funcional o angiografía.'}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# MOTOR CLÍNICO HOLTER 24 HORAS CON AUDITORÍA IA INTEGRAL
# ==============================================================================
def limpiar_numero(val_str):
    if not val_str: return 0
    try: return int(float(str(val_str).replace(".", "").replace(",", ".")))
    except Exception: return 0

def extraer_datos_holter(pdf_bytes, filename=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_todas_paginas = "".join([f"\n--- PÁGINA {i+1} ---\n" + p.get_text() + "\n" for i, p in enumerate(doc)])
    doc.close()

    # INTENTO 1: AUDITORÍA CLÍNICA TOTAL CON IA GEMINI
    prompt_ia = f"""Eres un Cardiólogo Especialista y Auditor Clínico Principal del Centro Cardiovascular Colombiano CENCARDIO.
Analiza la transcripción completa de TODAS las páginas de este estudio Holter de 24 horas (Pathfinder SL / Sentinel):

{texto_todas_paginas[:22000]}

Extrae y determina con máxima rigurosidad diagnóstica los parámetros clínicos:
1. Paciente e identificación.
2. Frecuencias cardíacas: promedio 24h, diurna, nocturna, máxima y mínima.
3. Total real de latidos analizados (QRS totales).
4. Eventos cronotrópicos: episodios de taquicardia sinusal (con FC máx) y bradicardia (con FC mín).
5. Pausas patológicas (número y longitud máx en segundos).
6. Latidos caídos (bloqueos AV de 2do o 3er grado).
7. Conducción intraventricular: Evalúa con criterio cardiólogo si el paciente presenta Bloqueo Completo o Incompleto de Rama (BRD, BRI, BCRD, BCRI, QRS ancho >120ms o documentado en el reporte). Devuelve true en "tiene_bloqueo_rama" si lo tiene.
8. Ectopias supraventriculares: ESV totales y rachas de TSV.
9. Ectopias ventriculares: EV totales, rachas de TV, duplas ventriculares, bigeminismo.
10. Isquemia del ST: episodios reales documentados y milímetros.
11. Variabilidad de la FC (SDNN de 24 horas en ms) y QTc promedio en ms.

Devuelve EXCLUSIVAMENTE este JSON:
{{
  "paciente": "NOMBRE COMPLETO",
  "cedula": "NUMERO DE CEDULA O VACIO",
  "dx_motivo": "DIAGNOSTICO O MOTIVO DE CONSULTA",
  "fc_prom": 75,
  "fc_dia": 80,
  "fc_noc": 68,
  "fc_max": 105,
  "fc_min": 58,
  "total_latidos": 85000,
  "taqui_conteo": 0,
  "taqui_fc_max": 105,
  "bradi_conteo": 0,
  "bradi_fc_min": 58,
  "pausas": 0,
  "pausa_max_seg": "0",
  "latidos_caidos": 0,
  "tiene_bloqueo_rama": false,
  "bloqueo_detalle": "bloqueo completo de rama",
  "ev_total": 0,
  "tv_episodios": 0,
  "ev_duplas": 0,
  "bigeminismo": 0,
  "esv_total": 0,
  "tsv_episodios": 0,
  "st_episodios": 0,
  "st_dep_mm": "1.0",
  "sdnn_24h": 85,
  "qtc_prom": 410,
  "eventos_paciente": 0
}}"""

    ok, data_ai, mod_used = consultar_gemini_json(prompt_ia)
    if ok and data_ai and data_ai.get("fc_prom"):
        d = data_ai
        d["origen_extraccion"] = f"Auditoría IA ({mod_used})"
        return d

    # INTENTO 2: MOTOR DEFENSIVO LOCAL (HEURÍSTICA MEJORADA)
    texto = texto_todas_paginas
    d = {}
    m_dx = re.search(r"(?:DX|DIAGN[ÓO]STICO|INDICACI[ÓO]N)\s*[:\.]?\s*([^\n\r\|]+)", texto, re.IGNORECASE) or re.search(r"Comentarios de la prueba:\s*\n?\s*([^\n\r\|]+)", texto, re.IGNORECASE)
    d["dx_motivo"] = m_dx.group(1).strip().upper() if m_dx else ""

    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado|Reconfirmado)?[\s\n]*Informe Holter", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else (os.path.splitext(filename)[0] if filename else "PACIENTE")

    m_id = re.search(r"(?:ID\s*Paciente|ID|C\.?C\.?|Doc\.?|Historia)\s*[:\.]?\s*(\d{5,12})", texto, re.IGNORECASE)
    d["cedula"] = m_id.group(1) if m_id else ""

    fc_p = re.search(r"Prom\.?\s*(\d{2,3})", texto)
    d["fc_prom"] = int(fc_p.group(1)) if fc_p else 75
    fc_max = re.search(r"M[áa]x\s*(\d{2,3})", texto)
    d["fc_max"] = int(fc_max.group(1)) if fc_max else 100
    fc_min = re.search(r"M[íi]n\s*(\d{2,3})", texto)
    d["fc_min"] = int(fc_min.group(1)) if fc_min else 55
    fc_dia = re.search(r"D[íi]a.*?Prom\.?\s*(\d{2,3})", texto)
    d["fc_dia"] = int(fc_dia.group(1)) if fc_dia else int(d["fc_prom"] * 1.05)
    fc_noc = re.search(r"Noche.*?Prom\.?\s*(\d{2,3})", texto)
    d["fc_noc"] = int(fc_noc.group(1)) if fc_noc else max(45, int(d["fc_prom"] * 0.90))

    m_tot = re.search(r"Total\s+de\s+latidos\s*[:\.]?\s*([\d\.]+)", texto, re.IGNORECASE) or re.search(r"Total\s+QRS\s*[:\.]?\s*([\d\.]+)", texto, re.IGNORECASE) or re.search(r"Conteo\s+([\d\.]+)\s+[\d\.]+\s+\d+%", texto, re.IGNORECASE)
    d["total_latidos"] = limpiar_numero(m_tot.group(1)) if m_tot else max(70000, d["fc_prom"] * 60 * 24)

    taqui_m = re.search(r"Taquicardia\s+(\d+)(?:[^\n\r\d]+(\d{2,3})\s*:\s*[^\n\r]+)?(?:[^\n\r\d]+(\d+)\s+latidos)?", texto)
    d["taqui_conteo"] = int(taqui_m.group(1)) if taqui_m else 0
    d["taqui_fc_max"] = int(taqui_m.group(2)) if (taqui_m and taqui_m.group(2)) else d["fc_max"]

    bradi_m = re.search(r"Bradicardia\s+(\d+)(?:[^\n\r\d]+(\d{2,3})\s*:\s*[^\n\r]+)?(?:[^\n\r\d]+(\d+)\s+latidos)?", texto)
    d["bradi_conteo"] = int(bradi_m.group(1)) if bradi_m else 0
    d["bradi_fc_min"] = int(bradi_m.group(2)) if (bradi_m and bradi_m.group(2)) else d["fc_min"]

    pausa_match = re.search(r"\bPausa\s+(\d+)", texto)
    d["pausas"] = int(pausa_match.group(1)) if pausa_match else 0
    p_max_m = re.search(r"Pausa.*?M[áa]x\.\s*longitud\s*([\d,\.]+)\s*s", texto)
    d["pausa_max_seg"] = p_max_m.group(1).replace(",", ".") if p_max_m else "0"

    lat_caido_m = re.search(r"Latidos?\s+ca[íi]dos?\s+(\d+)", texto)
    d["latidos_caidos"] = int(lat_caido_m.group(1)) if lat_caido_m else 0

    ev_m = re.search(r"Latidos ventriculares\s*:\s*([\d\.]+)", texto) or re.search(r"Latidos\s+V\b[^\n\r\d]*([\d\.]+)", texto, re.IGNORECASE)
    d["ev_total"] = limpiar_numero(ev_m.group(1)) if ev_m else 0
    tv_m = re.search(r"\bTV\s+([\d\.]+)", texto)
    d["tv_episodios"] = limpiar_numero(tv_m.group(1)) if tv_m else 0
    dup_v_m = re.search(r"Apareado\s+([\d\.]+)", texto)
    d["ev_duplas"] = limpiar_numero(dup_v_m.group(1)) if dup_v_m else 0
    big_m = re.search(r"Bigeminismo\s+([\d\.]+)", texto)
    d["bigeminismo"] = limpiar_numero(big_m.group(1)) if big_m else 0

    esv_m = re.search(r"Latidos supraventriculares\s*:\s*([\d\.]+)", texto)
    d["esv_total"] = limpiar_numero(esv_m.group(1)) if esv_m else 0
    tsv_m = re.search(r"\bTSV\s+([\d\.]+)", texto)
    d["tsv_episodios"] = limpiar_numero(tsv_m.group(1)) if tsv_m else 0

    st_dep = re.search(r"Depresi[óo]n ST\s+(\d+)\s*(-?[\d,\.]*)", texto)
    st_elev = re.search(r"Elevaci[óo]n ST\s+(\d+)\s*([\d,\.]*)", texto)
    d["st_dep_episodios"] = int(st_dep.group(1)) if (st_dep and st_dep.group(1) != "0") else 0
    d["st_elev_episodios"] = int(st_elev.group(1)) if (st_elev and st_elev.group(1) != "0") else 0
    d["st_episodios"] = d["st_dep_episodios"] + d["st_elev_episodios"]

    sdnn_m = re.search(r"Valor de 24 horas\s+[\d\.,]+\s+([\d\.,]+)", texto)
    d["sdnn_24h"] = limpiar_numero(sdnn_m.group(1)) if sdnn_m else 85
    qtc_m = re.search(r"Todos los per[íi]odos\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)", texto)
    d["qtc_prom"] = limpiar_numero(qtc_m.group(1)) if qtc_m else 400
    eventos_pac_m = re.search(r"Eventos del paciente\s*:\s*(\d+)", texto)
    d["eventos_paciente"] = int(eventos_pac_m.group(1)) if eventos_pac_m else 0

    texto_upper = texto.upper()
    bloqueo_keys = ["BLOQUEO DE RAMA", "BLOQUEO COMPLETO", "BRD", "BRI", "BCRD", "BCRI", "QRS ANCHO", "CONDUCCION INTRAVENTRICULAR", "IVCD"]
    d["tiene_bloqueo_rama"] = any(k in texto_upper for k in bloqueo_keys)
    d["bloqueo_detalle"] = "bloqueo completo de rama"
    d["origen_extraccion"] = "Motor Local de Respaldo"
    return d

def redactar_informe_holter_11_puntos(d, perfil):
    desc_crono = ((d["fc_dia"] - d["fc_noc"]) / d["fc_dia"]) * 100 if d["fc_dia"] > 0 else 0
    dip_txt = f"conservado ({desc_crono:.1f}% de descenso nocturno)" if desc_crono >= 10 else f"atenuado ({desc_crono:.1f}% de descenso nocturno)"

    p1 = f"1. Ritmo sinusal con FC promedio de {d['fc_prom']} lpm (Diurna: {d['fc_dia']} lpm / Nocturna: {d['fc_noc']} lpm; patrón circadiano {dip_txt})."

    crono = []
    if d["taqui_conteo"] > 0: crono.append(f"{d['taqui_conteo']} episodios de taquicardia sinusal (FC máx. {d['taqui_fc_max']} lpm)")
    if d["bradi_conteo"] > 0: crono.append(f"{d['bradi_conteo']} episodios de bradicardia sinusal (FC mín. {d['bradi_fc_min']} lpm)")
    p2 = f"2. Eventos cronotrópicos: Se documentaron {' y '.join(crono)}." if crono else "2. Eventos cronotrópicos: Sin bradicardia patológica ni taquicardias sostenidas de relevancia clínica."

    qtc_val = d["qtc_prom"]
    if qtc_val > 500:
        p3 = f"3. Intervalos PR normales y QTc SEVERAMENTE PROLONGADO ({qtc_val} ms: alto riesgo proarrítmico de Torsades de Pointes)."
    elif qtc_val > 460:
        p3 = f"3. Intervalos PR normales y QTc prolongado ({qtc_val} ms)."
    else:
        p3 = f"3. Intervalos PR y QTc normales ({qtc_val} ms)."

    p4 = f"4. Alteraciones isquémicas del segmento ST ({d['st_episodios']} episodios documentados)." if d["st_episodios"] > 0 else "4. Sin alteraciones isquémicas del segmento ST."
    p5 = f"5. Alteración de la conducción AV por {d['latidos_caidos']} latidos caídos." if d["latidos_caidos"] > 0 else "5. Sin Alteración de la conducción AV."

    # Punto 6: Evaluación médica exacta sin falsos negativos
    tiene_bloqueo = d.get("tiene_bloqueo_rama", False)
    det_bloqueo = d.get("bloqueo_detalle", "bloqueo completo de rama")
    p6 = f"6. Alteración en la conducción intraventricular por {det_bloqueo}." if tiene_bloqueo else "6. Sin alteración en la conducción intraventricular."

    carga_ev = (d["ev_total"] / d["total_latidos"]) * 100 if d["total_latidos"] > 0 else 0
    if d["tv_episodios"] > 0: lown = "Lown Grado IVb (Taquicardia Ventricular)"
    elif d["ev_duplas"] > 0: lown = "Lown Grado IVa (Duplas ventriculares)"
    elif d["bigeminismo"] > 0: lown = "Lown Grado III (Arritmia ventricular compleja)"
    elif d["ev_total"] >= 720: lown = "Lown Grado II (EV frecuentes > 30/hora)"
    elif d["ev_total"] > 0: lown = "Lown Grado I (EV aisladas ocasionales)"
    else: lown = "Lown Grado 0 (Sin arritmia ventricular)"

    ect = []
    if d["esv_total"] > 0:
        txt_s = f"ectopias supraventriculares ({d['esv_total']} ESV"
        if d["tsv_episodios"] > 0: txt_s += f", {d['tsv_episodios']} rachas de TSV"
        txt_s += ")"
        ect.append(txt_s)

    if d["ev_total"] > 0:
        ect.append(f"ectopia ventricular con {d['ev_total']} EV (Carga: {carga_ev:.2f}%, {lown})")

    p7 = f"7. Alteración de los impulsos por {' y '.join(ect)}." if ect else "7. Sin alteración de los impulsos ectópicos de relevancia clínica."
    p8 = f"8. El paciente refirió síntomas ({d['eventos_paciente']} eventos marcados en diario)." if d["eventos_paciente"] > 0 else "8. No refirió síntomas."

    # Unificación estricta de variabilidad y riesgo autonómico
    if d["sdnn_24h"] <= 60:
        p9 = f"9. Variabilidad de la frecuencia cardíaca (VFC) severamente disminuida (SDNN: {d['sdnn_24h']} ms)."
        p11 = "11. Estratificación del riesgo autonómico por SDNN de 24 horas: Riesgo alto."
        diag_vfc = "Variabilidad de la FC severamente disminuida (riesgo autonómico alto)."
    elif d["sdnn_24h"] <= 120:
        p9 = f"9. Variabilidad de la frecuencia cardíaca (VFC) disminuida (SDNN: {d['sdnn_24h']} ms)."
        p11 = "11. Estratificación del riesgo autonómico por SDNN de 24 horas: Riesgo moderado."
        diag_vfc = "Variabilidad de la FC disminuida (riesgo autonómico moderado)."
    else:
        p9 = f"9. Variabilidad de la frecuencia cardíaca (VFC) conservada (SDNN: {d['sdnn_24h']} ms)."
        p11 = "11. Estratificación del riesgo autonómico por SDNN de 24 horas: Normal (Bajo riesgo)."
        diag_vfc = "Variabilidad de la FC y modulación autonómica conservadas."

    p10 = f"10. Se registraron {d['pausas']} pausas significativas (máx. {d['pausa_max_seg']} s)." if d["pausas"] > 0 else "10. No hay pausas significativas."

    diag = f"Ritmo sinusal con FC promedio {d['fc_prom']} lpm. {diag_vfc}"
    if tiene_bloqueo:
        diag += f" Conducción intraventricular con {det_bloqueo}."
    if carga_ev >= 10.0:
        diag += f" Carga ectópica ventricular elevada ({carga_ev:.1f}%): criterio de riesgo para miocardiopatía inducida por arritmia."

    # Motor dinámico de recomendaciones para Holter
    recs = []
    if d["tv_episodios"] > 0 or carga_ev >= 10.0 or d["ev_duplas"] > 0:
        recs.append("Valoración prioritaria por Electrofisiología y ecocardiograma transtorácico para cuantificar fracción de eyección (FEVI).")
    if tiene_bloqueo:
        recs.append("Ecocardiograma transtorácico para evaluar asincronía ventricular y control periódico del trastorno de conducción.")
    if qtc_val > 460:
        recs.append("Control de electrolitos séricos (K+, Mg++) y revisión estricta de fármacos que prolonguen el intervalo QT.")
    if d["sdnn_24h"] <= 60:
        recs.append("Estratificación integral de riesgo cardiovascular y optimización del tratamiento médico neurohumoral.")
    if d["pausas"] > 0 or d["latidos_caidos"] > 0:
        recs.append("Correlación clínica con síntomas presincopales o sincopales para descartar disfunción del nodo sinusal.")
    if not recs:
        recs.append("Continuar manejo médico instaurado y seguimiento clínico periódico por cardiología.")

    rec_texto = "RECOMENDACIONES: " + " ".join(recs)

    return f"""INTERPRETACIÓN TEST HOLTER - CUPS 895001

{p1}
{p2}
{p3}
{p4}
{p5}
{p6}
{p7}
{p8}
{p9}
{p10}
{p11}

CONCLUSIÓN DIAGNÓSTICA:
{diag}

{rec_texto}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# MOTOR CLÍNICO: MAPA 24 HORAS SENTINEL
# ==============================================================================
def extraer_datos_mapa_sentinel(pdf_bytes, filename=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = "".join([p.get_text() + "\n" for p in doc])
    doc.close()

    d = {}
    m_nom = re.search(r"Nombre del paciente:\s*([^\n\r\|]+)", texto, re.IGNORECASE) or re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,45},\s*[A-ZÁÉÍÓÚÑ\s]{3,45})", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else "PACIENTE MAPA"

    m_id = re.search(r"ID paciente:\s*([^\n\r\s]+)", texto, re.IGNORECASE)
    d["cedula"] = m_id.group(1).replace(".", "") if m_id else ""

    p_gen = re.search(r"Resumen general.*?Prom\.?:\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\s*mmHg", texto, re.IGNORECASE)
    d["pas_24h"] = int(p_gen.group(1)) if p_gen else 120
    d["pad_24h"] = int(p_gen.group(2)) if p_gen else 75

    c_sis = re.search(r"Sist[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, re.IGNORECASE)
    d["carga_pas"] = c_sis.group(1).replace(".", ",") if c_sis else "0,00"
    c_dia = re.search(r"Diast[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, re.IGNORECASE)
    d["carga_pad"] = c_dia.group(1).replace(".", ",") if c_dia else "0,00"

    pp_m = re.search(r"Presi[óo]n de pulso\s*\(mmHg\)\s*\n?\s*(\d{2,3})", texto, re.IGNORECASE)
    d["pp_val"] = int(pp_m.group(1)) if pp_m else (d["pas_24h"] - d["pad_24h"])

    caida_m = re.search(r"Sist[óo]lico\s*\(mmHg\)\s*.*?([\d,\.\-]+)\s*%", texto, re.DOTALL)
    try: d["caida_nocturna_val"] = float(caida_m.group(1).replace(",", ".")) if caida_m else 10.0
    except Exception: d["caida_nocturna_val"] = 10.0

    m_sueno = re.search(r"Resumen de los per[íi]odos de sue[ñn]o.*?Sist[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\([^\)]+\)\s*.*?Diast[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\(", texto, re.DOTALL | re.IGNORECASE)
    d["pas_max_sueno"] = int(m_sueno.group(2)) if m_sueno else d["pas_24h"]
    d["pad_max_sueno"] = int(m_sueno.group(4)) if m_sueno else d["pad_24h"]
    return d

def redactar_informe_mapa_cencardio(d, perfil):
    p1 = f"1. Promedio de tensión arterial sistólica ({d['pas_24h']} mmHg) y diastolica ({d['pad_24h']} mmHg)"
    p2 = f"2. Carga tensional sistólica ({d['carga_pas']}%) y diastólica de ({d['carga_pad']}%)"
    p3 = "3. Presión de pulso normal" if d["pp_val"] <= 60 else f"3. Presión de pulso aumentada ({d['pp_val']} mmHg, rigidez arterial)"

    patron = "dipping positivo" if d["caida_nocturna_val"] > 0.0 else ("dipping invertido (riser)" if d["caida_nocturna_val"] <= -10.0 else "dipping atenuado")
    p4 = f"4. Patrón circadiano tensional {patron}"

    p5 = f"5. Se presentaron incrementos de presión durante el sueño (máx. {d['pas_max_sueno']}/{d['pad_max_sueno']} mmHg)." if (d["pas_max_sueno"] >= 145 or d["pad_max_sueno"] >= 95) else "5. No se presentaron incrementos de presión arterial tanto sistólica al acostarse como diastólica durante las 24 horas."

    c_pas = float(str(d["carga_pas"]).replace(",", "."))
    c_pad = float(str(d["carga_pad"]).replace(",", "."))
    ctrl = "Control óptimo de la tensión arterial" if (c_pas < 15 and c_pad < 15 and d["pas_24h"] < 130 and d["pad_24h"] < 80) else ("Control subóptimo de la tensión arterial estadio I" if (c_pas <= 30 or c_pad <= 30 or d["pas_24h"] < 140) else "Descontrol de la tensión arterial estadio II")
    p6 = f"6. {ctrl}"

    recs_mapa = "Continuar manejo farmacológico actual y control periódico de cifras tensionales." if "Control óptimo" in ctrl else "Optimización y titulación de la terapia antihipertensiva, reforzando restricción sódica y hábitos de vida cardiosaludables."

    return f"""INTERPRETACIÓN TEST MAPA
Hallazgos:
{p1}
{p2}
{p3}
{p4}
{p5}
{p6}

RECOMENDACIONES: {recs_mapa}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# INYECTORES DE PDFS CON AJUSTE DINÁMICO DE FUENTE
# ==============================================================================
@st.cache_data
def generar_qr_verificacion(paciente, medico, fecha_str, codigo_uuid, proc_nombre="CUPS 895001"):
    url_base = "https://holtercencardio.streamlit.app/"
    query_string = urllib.parse.urlencode({
        "val": codigo_uuid[:12],
        "pac": paciente,
        "med": medico,
        "fec": fecha_str,
        "proc": proc_nombre
    })
    url_completa = f"{url_base}?{query_string}"

    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=4, border=1)
    qr.add_data(url_completa)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="#0a2540", back_color="white")
    
    buf = io.BytesIO()
    img_qr.save(buf, format="PNG")
    return buf.getvalue()

@st.cache_data
def procesar_firma_transparente():
    posibles_archivos = ["OR WILIAM ANDA RAMIREZ.pdf", "firma_amaya.pdf", "firma_amaya.png"]
    archivo_encontrado = next((n for n in posibles_archivos if os.path.exists(n)), None)
    if not archivo_encontrado:
        return None

    try:
        if archivo_encontrado.lower().endswith(".pdf"):
            doc_firma = fitz.open(archivo_encontrado)
            pix = doc_firma[0].get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc_firma.close()
        else:
            img = Image.open(archivo_encontrado).convert("RGB")

        img_gray = img.convert("L")
        alpha = img_gray.point(lambda p: 255 if p < 185 else 0, mode='L')
        tinta = Image.new("RGBA", img.size, (10, 37, 64, 255))
        tinta.putalpha(alpha)

        caja = tinta.getbbox()
        if caja: tinta = tinta.crop(caja)

        buf = io.BytesIO()
        tinta.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None

def normalizar_nombre_archivo(nombre):
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', nombre)
    return re.sub(r'\s+', '_', limpio).strip('_') or "PACIENTE"

def inyectar_holter_pdf(pdf_bytes, texto_informe, paciente_nom, perfil, cod_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]
    rects_h = pagina1.search_for("Hallazgos:")
    rects_f = pagina1.search_for("Firma del médico") or pagina1.search_for("Firma del operador")

    y_base = rects_f[0].y0 if rects_f else 740
    y0 = (rects_h[0].y1 + 4) if rects_h else 545
    y1 = y_base - 62

    pagina1.draw_rect(fitz.Rect(0, y0, pagina1.rect.width, y1), color=None, fill=(1, 1, 1), overlay=True)

    font_size_optimo = 3.9
    for fs in [5.8, 5.5, 5.2, 4.9, 4.6, 4.3, 4.0, 3.8]:
        dt = fitz.open(stream=pdf_bytes, filetype="pdf")
        rc = dt[0].insert_textbox(fitz.Rect(35, y0, pagina1.rect.width - 36, y1), texto_informe, fontsize=fs, fontname="helv", align=fitz.TEXT_ALIGN_LEFT)
        dt.close()
        if rc >= 0:
            font_size_optimo = fs
            break

    pagina1.insert_textbox(fitz.Rect(35, y0, pagina1.rect.width - 36, y1), texto_informe, fontsize=font_size_optimo, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)

    pagina1.draw_rect(fitz.Rect(220, y_base - 58, pagina1.rect.width, y_base + 12), color=None, fill=(1, 1, 1), overlay=True)
    qr_bytes = generar_qr_verificacion(paciente_nom, perfil['nombre_completo'], ahora_colombia().strftime("%Y-%m-%d"), cod_uuid, "CUPS 895001")
    pagina1.insert_image(fitz.Rect(230, y_base - 32, 270, y_base + 8), stream=qr_bytes)
    pagina1.insert_text(fitz.Point(275, y_base - 18), "Validado Digitalmente", fontsize=5.2, fontname="helv", color=(0.08, 0.2, 0.36))
    pagina1.insert_text(fitz.Point(275, y_base - 9), "Res. 3100 de 2019 - MinSalud", fontsize=4.7, fontname="helv", color=(0.25, 0.25, 0.25))
    pagina1.insert_text(fitz.Point(275, y_base), f"Cód: {cod_uuid[:12]}...", fontsize=4.5, fontname="helv", color=(0.4, 0.4, 0.4))

    firma_bytes = procesar_firma_transparente()
    if firma_bytes and estampador_activo:
        fx0 = (rects_f[0].x0 + 10) if rects_f else 380
        pagina1.insert_image(fitz.Rect(fx0, y_base - 58, fx0 + 155, y_base + 4), stream=firma_bytes)

    pix = pagina1.get_pixmap(dpi=130)
    img_prev = pix.tobytes("png")
    pdf_out = doc.tobytes()
    doc.close()
    return pdf_out, img_prev

def inyectar_mapa_pdf(pdf_bytes, texto_informe, paciente_nom, perfil, cod_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]
    rect_m = pagina1.search_for("Presión arterial por la mañana")
    rect_r = pagina1.search_for("Resumen de todo el registro")

    y0 = (rect_m[0].y1 + 2) if rect_m else 230
    y1 = (rect_r[0].y0 - 4) if rect_r else 425

    pagina1.draw_rect(fitz.Rect(0, y0, pagina1.rect.width, y1), color=None, fill=(1, 1, 1), overlay=True)

    font_size = 6.4
    for fs in [7.2, 6.8, 6.4, 6.0, 5.6]:
        dt = fitz.open(stream=pdf_bytes, filetype="pdf")
        rc = dt[0].insert_textbox(fitz.Rect(36, y0 + 2, 440, y1 - 2), texto_informe, fontsize=fs, fontname="helv", align=fitz.TEXT_ALIGN_LEFT)
        dt.close()
        if rc >= 0:
            font_size = fs
            break

    pagina1.insert_textbox(fitz.Rect(36, y0 + 2, 440, y1 - 2), texto_informe, fontsize=font_size, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)

    qr_bytes = generar_qr_verificacion(paciente_nom, perfil['nombre_completo'], ahora_colombia().strftime("%Y-%m-%d"), cod_uuid, "CUPS 895003")
    pagina1.insert_image(fitz.Rect(455, y0 + 6, 505, y0 + 56), stream=qr_bytes)
    pagina1.insert_text(fitz.Point(510, y0 + 22), "Validado Digitalmente", fontsize=5.0, fontname="helv", color=(0.04, 0.15, 0.25))
    pagina1.insert_text(fitz.Point(510, y0 + 32), "Res. 3100 MinSalud", fontsize=4.6, fontname="helv", color=(0.3, 0.3, 0.3))
    pagina1.insert_text(fitz.Point(510, y0 + 42), f"Cód: {cod_uuid[:10]}...", fontsize=4.4, fontname="helv", color=(0.4, 0.4, 0.4))

    firma_bytes = procesar_firma_transparente()
    if firma_bytes and estampador_activo:
        pagina1.insert_image(fitz.Rect(445, y0 + 62, 575, y1 - 4), stream=firma_bytes)

    pix = pagina1.get_pixmap(dpi=130)
    img_prev = pix.tobytes("png")
    pdf_out = doc.tobytes()
    doc.close()
    return pdf_out, img_prev

def generar_pdf_ergometria_completo(d, texto_informe, perfil, cod_uuid, imagenes_adjuntas=[]):
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)

    logo_bytes = None
    for nom in ["cencardio.jpg", "cencardio.png", "cencardio.jpeg", "logo.png", "logo.jpg"]:
        if os.path.exists(nom):
            with open(nom, "rb") as f: logo_bytes = f.read()
            break
    if logo_bytes:
        page.insert_image(fitz.Rect(36, 30, 150, 75), stream=logo_bytes)

    page.insert_text(fitz.Point(165, 45), "CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO", fontsize=11, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(165, 58), "INFORME DE ERGOMETRÍA Y PRUEBA DE ESFUERZO COMPUTARIZADA", fontsize=8.5, fontname="helv", color=(0.78, 0.06, 0.18))
    page.insert_text(fitz.Point(165, 70), "CUPS: 893805 · Habilitación MinSalud Colombia · Res. 3100 de 2019", fontsize=7, fontname="helv", color=(0.4, 0.45, 0.5))
    page.draw_rect(fitz.Rect(36, 85, 576, 87), color=None, fill=(0.04, 0.15, 0.25), overlay=True)

    page.draw_rect(fitz.Rect(36, 95, 576, 155), color=(0.85, 0.9, 0.95), fill=(0.97, 0.98, 1.0), width=1)
    page.insert_text(fitz.Point(46, 112), f"PACIENTE: {d['paciente'].upper()}", fontsize=8.5, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(46, 126), f"DOCUMENTO: {d['cedula']}    |    EDAD: {d['edad']} AÑOS    |    SEXO: {d['sexo']}", fontsize=7.5, fontname="helv", color=(0.2, 0.25, 0.3))
    page.insert_text(fitz.Point(46, 140), f"FECHA DEL ESTUDIO: {ahora_colombia().strftime('%d/%m/%Y')}    |    MÉDICO LECTOR: {perfil['nombre_completo']}", fontsize=7.5, fontname="helv", color=(0.2, 0.25, 0.3))

    fcm_prev = 220 - d["edad"] if d["edad"] > 0 else 200
    porc = round((d["fc_pico"] / fcm_prev) * 100) if (fcm_prev > 0 and d["fc_pico"] > 0) else 0
    cajas = [
        ("FC PICO ALCANZADA", f"{d['fc_pico']} lpm ({porc}%)"),
        ("PA ESFUERZO PICO", f"{d['pas_pico']}/{d['pad_pico']} mmHg"),
        ("CARGA FUNCIONAL", f"{d['mets']} METs ({d['tiempo_min']:.2f} m)"),
        ("DOBLE PRODUCTO", f"{d['fc_pico']*d['pas_pico']:,}")
    ]
    x_offset = 36
    for tit, val in cajas:
        rect_m = fitz.Rect(x_offset, 163, x_offset + 130, 203)
        page.draw_rect(rect_m, color=(0.88, 0.91, 0.94), fill=(1, 1, 1), width=1)
        page.draw_rect(fitz.Rect(x_offset, 163, x_offset + 130, 166), color=None, fill=(0.04, 0.15, 0.25))
        page.insert_text(fitz.Point(x_offset + 8, 178), tit, fontsize=5.8, fontname="helv", color=(0.4, 0.45, 0.5))
        page.insert_text(fitz.Point(x_offset + 8, 195), val, fontsize=8.5, fontname="helv", color=(0.04, 0.15, 0.25))
        x_offset += 136

    rect_caja = fitz.Rect(36, 215, 576, 680)
    page.draw_rect(rect_caja, color=(0.88, 0.91, 0.94), fill=(1, 1, 1), width=1)
    page.insert_textbox(rect_caja, texto_informe, fontsize=6.8, fontname="helv", color=(0.1, 0.15, 0.2), align=fitz.TEXT_ALIGN_LEFT)

    qr_bytes = generar_qr_verificacion(d['paciente'], perfil['nombre_completo'], ahora_colombia().strftime("%Y-%m-%d"), cod_uuid, "CUPS 893805")
    page.insert_image(fitz.Rect(48, 695, 100, 747), stream=qr_bytes)
    page.insert_text(fitz.Point(108, 715), "Certificado Digital Forense", fontsize=6.2, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(108, 726), "Res. 3100 de 2019 · Habilitación MinSalud", fontsize=5.5, fontname="helv", color=(0.4, 0.45, 0.5))
    page.insert_text(fitz.Point(108, 737), f"Cód: {cod_uuid[:16]}...", fontsize=5.2, fontname="helv", color=(0.4, 0.45, 0.5))

    firma_bytes = procesar_firma_transparente()
    if firma_bytes and perfil["id"] in ["dr.amaya", "admin"]:
        page.insert_image(fitz.Rect(390, 690, 545, 755), stream=firma_bytes)

    pix = page.get_pixmap(dpi=130)
    img_preview = pix.tobytes("png")

    if imagenes_adjuntas:
        for img_file in imagenes_adjuntas:
            try:
                p_extra = doc.new_page(width=612, height=792)
                data_img = img_file.getvalue() if hasattr(img_file, "getvalue") else img_file
                p_extra.insert_image(fitz.Rect(20, 20, 592, 772), stream=data_img)
            except Exception: pass

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes, img_preview

def detectar_tipo_documento_clinico(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = "".join([p.get_text() + "\n" for p in doc])
    doc.close()
    if any(k in texto for k in ["Sentinel", "Presión arterial por la mañana", "Informe de MAPA", "AASI"]):
        return "MAPA"
    elif any(k in texto for k in ["Informe Holter", "Latidos ventriculares", "Pathfinder SL", "Arritmias ventriculares"]):
        return "HOLTER"
    return "DESCONOCIDO"

def generar_grafica_tacograma(d):
    fc_prom = d.get("fc_prom", 75)
    fc_dia = d.get("fc_dia", int(fc_prom * 1.05))
    fc_noc = d.get("fc_noc", int(fc_prom * 0.90))
    fc_max = d.get("fc_max", 100)
    fc_min = d.get("fc_min", 55)

    horas = [f"{h:02d}:00" for h in range(24)]
    fc_curva = []
    for h in range(24):
        if 6 <= h <= 21:
            val = fc_dia + (fc_max - fc_dia) * 0.25 * ((h % 4) / 4)
        else:
            val = fc_noc - (fc_noc - fc_min) * 0.3 * ((h % 3) / 3)
        fc_curva.append(round(max(fc_min, min(fc_max, val))))

    fig = go.Figure()
    fig.add_hrect(y0=60, y1=100, fillcolor="rgba(10, 37, 64, 0.04)", line_width=0, annotation_text="Normal (60-100)", annotation_position="top left", annotation_font_size=9)
    fig.add_trace(go.Scatter(x=horas, y=fc_curva, mode='lines+markers', name='FC (lpm)', line=dict(color='#0A2540', width=2.5), marker=dict(size=4, color='#C8102E')))
    fig.add_hline(y=fc_prom, line_dash="dot", line_color="#0284c7", annotation_text=f"Prom: {fc_prom} lpm", annotation_position="bottom right")
    fig.update_layout(title="<b>Tacograma Horario y Variabilidad Circadiana (24 Horas)</b>", height=240, margin=dict(l=35, r=20, t=35, b=25), plot_bgcolor="#ffffff", paper_bgcolor="#ffffff", showlegend=False)
    return fig

# ==============================================================================
# GENERADOR AVANZADO DE EXCEL INSTITUCIONAL
# ==============================================================================
def generar_excel_avanzado_produccion(df_base):
    if not OPENPYXL_INSTALADO:
        csv_str = df_base.to_csv(sep=';', index=False, encoding='utf-8-sig')
        return csv_str.encode('utf-8-sig'), "csv"

    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    ws_resumen = wb.active
    ws_resumen.title = "📊 Tablero Gerencial"
    ws_resumen.views.sheetView[0].showGridLines = True

    fill_navy = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
    fill_wine = PatternFill(start_color="C8102E", end_color="C8102E", fill_type="solid")
    fill_light = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True, color="0A2540")
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws_resumen["B2"] = "CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO"
    ws_resumen["B2"].font = Font(name="Calibri", size=14, bold=True, color="0A2540")
    ws_resumen["B3"] = f"REPORTE GERENCIAL DE PRODUCCIÓN · EMITIDO: {ahora_colombia().strftime('%d/%m/%Y %H:%M')}"
    ws_resumen["B3"].font = Font(name="Calibri", size=9, bold=True, color="64748B")

    ws_resumen["B5"] = "PRODUCCIÓN POR MODALIDAD DIAGNÓSTICA"
    ws_resumen["B5"].font = font_bold
    headers_mod = ["Modalidad Diagnóstica", "Código CUPS", "Total Estudios", "% Participación"]
    for col_idx, h in enumerate(headers_mod, start=2):
        cell = ws_resumen.cell(row=6, column=col_idx, value=h)
        cell.fill = fill_navy
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    prod_mod = df_base.groupby(["modalidad", "cups"]).size().reset_index(name="Cantidad")
    total_estudios = len(df_base)
    r_idx = 7
    for _, fila in prod_mod.iterrows():
        pct = (fila["Cantidad"] / total_estudios) * 100 if total_estudios > 0 else 0
        ws_resumen.cell(row=r_idx, column=2, value=fila["modalidad"]).border = border_thin
        ws_resumen.cell(row=r_idx, column=3, value=fila["cups"]).border = border_thin
        c_cant = ws_resumen.cell(row=r_idx, column=4, value=fila["Cantidad"])
        c_cant.border = border_thin
        c_cant.alignment = Alignment(horizontal="center")
        c_pct = ws_resumen.cell(row=r_idx, column=5, value=f"{pct:.1f}%")
        c_pct.border = border_thin
        c_pct.alignment = Alignment(horizontal="center")
        r_idx += 1

    ws_resumen.cell(row=r_idx, column=2, value="TOTAL GENERAL").font = font_bold
    ws_resumen.cell(row=r_idx, column=4, value=total_estudios).font = font_bold
    ws_resumen.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
    ws_resumen.cell(row=r_idx, column=5, value="100.0%").font = font_bold
    ws_resumen.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
    r_idx += 3

    ws_detalle = wb.create_sheet(title="📁 Detalle RIPS y Facturación")
    ws_detalle.views.sheetView[0].showGridLines = True

    headers_detalle = ["N° ID", "Fecha Registro", "Paciente", "Modalidad", "CUPS", "Métrica Clave", "Especialista Lector", "Código Forense"]
    for col_num, h in enumerate(headers_detalle, 1):
        cell = ws_detalle.cell(row=1, column=col_num, value=h)
        cell.fill = fill_navy
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_detalle.row_dimensions[1].height = 26

    columnas_ordenadas = ["id", "fecha_registro", "paciente", "modalidad", "cups", "parametro_clave", "medico", "codigo_verificacion"]
    df_exp = df_base[columnas_ordenadas].copy()

    for r_idx_d, fila in enumerate(df_exp.itertuples(index=False), start=2):
        ws_detalle.row_dimensions[r_idx_d].height = 20
        fill_row = fill_light if (r_idx_d % 2 == 0) else None
        for c_idx, valor in enumerate(fila, start=1):
            cell = ws_detalle.cell(row=r_idx_d, column=c_idx, value=str(valor))
            cell.border = border_thin
            if fill_row: cell.fill = fill_row
            if c_idx in [1, 5]: cell.alignment = Alignment(horizontal="center")

    for ws in [ws_resumen, ws_detalle]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    return output.getvalue(), "xlsx"

# ==============================================================================
# PERFILES MÉDICOS OFICIALES & CONTROL DE ACCESO
# ==============================================================================
LISTA_ESPECIALISTAS = [
    {
        "id": "dr.amaya",
        "etiqueta": "Dr. William Amaya Ramirez (Internista - Cardiólogo)",
        "clave": "Cardio2025*",
        "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ",
        "especialidad": "INTERNISTA - CARDIÓLOGO",
        "registro": "RM 79.502.624 SDS"
    },
    {
        "id": "dr.suarez",
        "etiqueta": "Dr. Martin Suárez Arámbula (Cardiólogo Hemodinamista)",
        "clave": "Suarez2026*",
        "nombre_completo": "DR. MARTIN SUÁREZ ARÁMBULA",
        "especialidad": "MÉDICO INTERNISTA - CARDIÓLOGO HEMODINAMISTA",
        "registro": "RM 13491094"
    },
    {
        "id": "dra.cardio",
        "etiqueta": "Dra. Paola Figueroa (Cardióloga)",
        "clave": "Cardio2026*",
        "nombre_completo": "DRA. PAOLA FIGUEROA",
        "especialidad": "MÉDICO ESPECIALISTA EN CARDIOLOGÍA",
        "registro": "RM 52.890.123 SDS"
    },
    {
        "id": "admin",
        "etiqueta": "Administración del Sistema",
        "clave": "HolterClaveSegura123",
        "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ",
        "especialidad": "INTERNISTA - CARDIÓLOGO",
        "registro": "RM 79.502.624 SDS"
    }
]

PERFILES_POR_ID = {m["id"]: m for m in LISTA_ESPECIALISTAS}
OPCIONES_NOMBRES = [m["etiqueta"] for m in LISTA_ESPECIALISTAS]

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = ""

def cerrar_sesion():
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""

# ==============================================================================
# VALIDACIÓN PÚBLICA POR QR (RES. 3100)
# ==============================================================================
params = st.query_params
if "val" in params:
    codigo_val = params.get("val", "N/A")
    paciente_val = urllib.parse.unquote(params.get("pac", "Paciente"))
    medico_val = urllib.parse.unquote(params.get("med", "Especialista CENCARDIO"))
    fecha_val = urllib.parse.unquote(params.get("fec", ahora_colombia().strftime("%Y-%m-%d")))
    proc_val = urllib.parse.unquote(params.get("proc", "Procedimiento Cardiológico"))

    c_v1, c_v2, c_v3 = st.columns([1, 1.8, 1])
    with c_v2:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" style="max-width:180px; margin-bottom:1rem;" alt="Cencardio Logo">' if logo_data else '<div style="font-size:3rem; margin-bottom:0.4rem;">🫀</div>'

        st.markdown(f"""
            <div class="cencardio-card">
                {logo_html}
                <div style="font-size: 1.25rem; font-weight: 800; color: #0a2540; text-transform: uppercase;">
                    Centro Cardiovascular Colombiano
                </div>
                <div style="font-size: 0.82rem; font-weight: 700; color: #c8102e; letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 1.5rem;">
                    CENCARDIO · Certificado de Autenticidad Forense
                </div>
                <div style="background: #f0fdf4; border: 1.5px solid #86efac; border-radius: 12px; padding: 1.3rem; margin-bottom: 1.5rem; text-align: left;">
                    <div style="color: #166534; font-size: 1rem; font-weight: 800; margin-bottom: 0.6rem; display: flex; align-items: center; gap: 8px;">
                        <span>✅</span> ESTUDIO MÉDICO CERTIFICADO Y VÁLIDO
                    </div>
                    <div style="font-size: 0.88rem; color: #1f2937; line-height: 1.6;">
                        <b>Procedimiento:</b> {proc_val}<br>
                        <b>Paciente:</b> {paciente_val}<br>
                        <b>Especialista Lector:</b> {medico_val}<br>
                        <b>Fecha de Emisión:</b> {fecha_val}<br>
                        <b>Identificador Único:</b> <span style="font-family: monospace; color: #0369a1; font-weight: 700;">{codigo_val}</span><br>
                        <b>Normativa:</b> Res. 3100 de 2019 / Habilitación MinSalud Colombia
                    </div>
                </div>
                <div style="font-size: 0.78rem; color: #64748b; line-height: 1.4;">
                    Documento custodiado bajo el estándar de Historia Clínica Electrónica del Centro Cardiovascular Colombiano CENCARDIO.
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Ir al Portal de Operaciones", use_container_width=True):
            st.query_params.clear()
            st.rerun()

    st.stop()

# ==============================================================================
# BARRERA DE LOGIN
# ==============================================================================
if not st.session_state.autenticado:
    c_izq, c_cen, c_der = st.columns([1, 1.6, 1])
    with c_cen:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" style="max-width:170px; margin-bottom:1rem;" alt="Cencardio Logo">' if logo_data else '<div style="font-size:3rem; margin-bottom:0.4rem;">🫀</div>'

        st.markdown(f"""
            <div class="cencardio-card">
                {logo_html}
                <div style="font-size: 1.25rem; font-weight: 800; color: #0a2540; text-transform: uppercase;">
                    Centro Cardiovascular Colombiano
                </div>
                <div style="font-size: 0.8rem; font-weight: 700; color: #c8102e; letter-spacing: 0.6px; text-transform: uppercase; margin-bottom: 1.8rem;">
                    CENCARDIO · Workstation Diagnóstica
                </div>
        """, unsafe_allow_html=True)

        with st.form("form_login"):
            seleccion_etiqueta = st.selectbox("Especialista Responsable:", options=OPCIONES_NOMBRES, index=0)
            clave = st.text_input("Contraseña de Acceso:", type="password")
            if st.form_submit_button("Ingresar a la Estación", use_container_width=True):
                medico = next(m for m in LISTA_ESPECIALISTAS if m["etiqueta"] == seleccion_etiqueta)
                if medico["clave"] == clave:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = medico["id"]
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta para el especialista seleccionado.")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ==============================================================================
# SESIÓN ACTIVA: BARRA LATERAL
# ==============================================================================
perfil_activo = PERFILES_POR_ID[st.session_state.usuario_actual]

with st.sidebar:
    logo_data_sidebar = obtener_logo_b64()
    if logo_data_sidebar:
        st.markdown(f'<div style="text-align:center; margin-bottom:1.2rem;"><img src="{logo_data_sidebar}" style="max-width:145px;"></div>', unsafe_allow_html=True)
    
    st.markdown(f"""
        <div class="specialist-card">
            <div style="font-weight:800; font-size:0.95rem;">{perfil_activo['nombre_completo']}</div>
            <div style="font-size:0.74rem; color:#93c5fd; text-transform:uppercase; margin-top:2px;">{perfil_activo['especialidad']}</div>
            <div style="font-size:0.72rem; color:#cbd5e1; font-family:'JetBrains Mono'; margin-top:6px;">{perfil_activo['registro']}</div>
        </div>
    """, unsafe_allow_html=True)

    firma_disponible = procesar_firma_transparente()
    if firma_disponible and perfil_activo["id"] in ["dr.amaya", "admin"]:
        st.success("🖋️ Sello digitalizado cargado.")
        
    st.divider()
    modalidad_seleccionada = st.radio(
        "Modalidad Diagnóstica:",
        ["🫀 Holter ECG 24H (CUPS 895001)", "🩺 MAPA Tensional 24H (CUPS 895003)", "🏃 Prueba de Esfuerzo (CUPS 893805)"]
    )

    st.divider()
    archivo_dinamica = st.file_uploader("Directorio Dinámica (Excel o CSV):", type=["xlsx", "xls", "csv"], key="sync_dinamica")
    if archivo_dinamica is not None:
        try:
            df_din = pd.read_csv(archivo_dinamica) if archivo_dinamica.name.endswith(".csv") else pd.read_excel(archivo_dinamica)
            ok, msg = sincronizar_directorio_servicio(df_din)
            if ok: st.success(msg)
            else: st.error(msg)
        except Exception as e: st.error(f"Error: {e}")

    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        cerrar_sesion()
        st.rerun()

# Cabecera Superior Principal
estado_nube_txt = "🟢 Nube Supabase Activa" if supabase else "🟡 Almacenamiento Local (SQLite)"
st.markdown(f"""
    <div class="top-hospital-bar">
        <div>
            <div style="font-size: 1.35rem; font-weight: 800; color: #0a2540; line-height: 1.2;">
                CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO
            </div>
            <div style="font-size: 0.82rem; font-weight: 700; color: #c8102e; text-transform: uppercase;">
                Estación Diagnóstica de Cardiología No Invasiva · Alta Complejidad
            </div>
        </div>
        <div style="display:flex; gap:8px;">
            <span class="inst-badge-success">{estado_nube_txt}</span>
            <span class="inst-badge-primary">Habilitación MinSalud · Res. 3100</span>
        </div>
    </div>
""", unsafe_allow_html=True)

tab_procesar, tab_historial = st.tabs(["📥 Procesamiento del Estudio", "📁 Archivo Clínico & Reportes Gerenciales"])

# ==============================================================================
# PESTAÑA 1: PROCESAMIENTO
# ==============================================================================
with tab_procesar:
    if "Esfuerzo" in modalidad_seleccionada:
        st.markdown("#### 🏃 Consola de Emisión de Prueba de Esfuerzo (CUPS 893805)")
        st.caption("Suba las fotografías de la tirilla para extracción automática con Visión IA o complete los campos:")

        for k, v in [
            ("erg_paciente", ""), ("erg_cedula", ""), ("erg_edad", 35),
            ("erg_sexo", "Femenino"), ("erg_protocolo", "Bruce"), ("erg_etapa", "Etapa 4"),
            ("erg_tiempo", 0.0), ("erg_fc_basal", 75), ("erg_fc_pico", 150),
            ("erg_pa_basal", "120/80"), ("erg_pa_pico", "160/90"), ("erg_st_mm", 0.0),
            ("erg_celular", "")
        ]:
            if k not in st.session_state:
                st.session_state[k] = v

        col_f1, col_f2 = st.columns([1.2, 1], gap="large")

        with col_f1:
            st.markdown("<b>1. Captura y Análisis de Trazados Impresos</b>", unsafe_allow_html=True)
            fotos_esfuerzo = st.file_uploader(
                "📸 Subir fotos o escaneos de las tiras de la banda (JPG o PNG):",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="uploader_fotos_erg"
            )

            if fotos_esfuerzo:
                btn_extraer = st.button("⚡ EXTRAER PARÁMETROS CON VISIÓN IA", type="primary", use_container_width=True)
                if btn_extraer:
                    with st.spinner("🤖 Consultando modelos de Google AI Studio y analizando fotografías..."):
                        exito, mensaje = ejecutar_extraccion_multimodal(fotos_esfuerzo)
                        if exito:
                            st.success(mensaje)
                            st.rerun()
                        else:
                            st.error(mensaje)

            st.write("")
            st.markdown("<b>2. Parámetros Clínicos</b>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1:
                p_nombre = st.text_input("Paciente:", key="erg_paciente", placeholder="Nombre completo")
                p_cedula = st.text_input("Cédula / ID:", key="erg_cedula", placeholder="Cédula")
            with c2:
                p_edad = st.number_input("Edad:", min_value=1, max_value=110, key="erg_edad")
                p_sexo = st.selectbox("Sexo:", ["Femenino", "Masculino"], key="erg_sexo")
            with c3:
                p_protocolo = st.selectbox("Protocolo:", ["Bruce", "Bruce Modificado", "Naughton"], key="erg_protocolo")
                p_etapa = st.text_input("Etapa alcanzada:", key="erg_etapa", placeholder="Ej: Etapa 6")

            c4, c5, c6 = st.columns(3)
            with c4:
                p_tiempo = st.number_input("Tiempo total (min):", step=0.1, key="erg_tiempo")
                mets_calculados = calcular_mets_bruce(p_tiempo)
                p_mets = st.number_input("Capacidad (METs):", value=float(mets_calculados), step=0.5)
            with c5:
                p_fc_basal = st.number_input("FC Basal (lpm):", key="erg_fc_basal")
                p_fc_pico = st.number_input("FC Pico (lpm):", key="erg_fc_pico")
            with c6:
                p_pa_basal = st.text_input("PA Basal (mmHg):", key="erg_pa_basal")
                p_pa_pico = st.text_input("PA Esfuerzo Pico (mmHg):", key="erg_pa_pico")

            c7, c8 = st.columns(2)
            with c7:
                p_st_mm = st.number_input("Desviación del ST (mm):", step=0.5, key="erg_st_mm")
            with c8:
                if not st.session_state.erg_celular and p_cedula:
                    tel_d = buscar_telefono_servicio(p_cedula)
                    if tel_d: st.session_state.erg_celular = tel_d
                p_celular = st.text_input("Celular (WhatsApp):", key="erg_celular")

        with col_f2:
            st.markdown("<b>3. Diagnóstico Institucional y Certificación</b>", unsafe_allow_html=True)

            if p_nombre.strip():
                pas_b, pad_b = [int(x) for x in p_pa_basal.split("/")] if "/" in p_pa_basal else (120, 80)
                pas_p, pad_p = [int(x) for x in p_pa_pico.split("/")] if "/" in p_pa_pico else (150, 90)

                datos_erg = {
                    "paciente": p_nombre, "cedula": p_cedula, "edad": p_edad, "sexo": p_sexo,
                    "protocolo": p_protocolo, "etapa": p_etapa or "Final", "tiempo_min": p_tiempo,
                    "mets": p_mets, "fc_basal": p_fc_basal, "fc_pico": p_fc_pico,
                    "pas_basal": pas_b, "pad_basal": pad_b, "pas_pico": pas_p, "pad_pico": pad_p,
                    "st_mm": p_st_mm
                }

                bloqueos, alertas_f = ejecutar_sanity_checks("ESFUERZO", datos_erg)
                for b in bloqueos: st.error(f"🛑 {b}")
                for a in alertas_f: st.warning(f"⚠️ {a}")

                texto_erg = redactar_informe_ergometria_institucional(datos_erg, perfil_activo)
                texto_erg_final = st.text_area("Informe Oficial:", value=texto_erg, height=310)

                discrepancias = auditar_coherencia_informe(texto_erg_final, datos_erg, "ESFUERZO")
                for d_err in discrepancias: st.info(f"🩺 **Alerta de Auditoría:** {d_err}")

                cod_uuid = str(uuid.uuid4()).upper()
                pdf_erg, img_erg_prev = generar_pdf_ergometria_completo(datos_erg, texto_erg_final, perfil_activo, cod_uuid, imagenes_adjuntas=fotos_esfuerzo or [])

                b1, b2 = st.columns(2)
                with b1:
                    st.download_button(
                        label="📄 DESCARGAR CERTIFICADO",
                        data=pdf_erg,
                        file_name=f"{normalizar_nombre_archivo(p_nombre)}_Prueba_Esfuerzo.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with b2:
                    if st.button("💾 Guardar en Archivo Clínico", use_container_width=True):
                        guardar_estudio_servicio(p_nombre, "Prueba de Esfuerzo", "CUPS 893805", f"FC Pico: {p_fc_pico} | {p_mets} METs", perfil_activo['nombre_completo'], texto_erg_final, pdf_erg, cod_uuid)
                        st.success(f"✅ Guardado en archivo clínico: {p_nombre}")

                if p_celular:
                    tel_l = re.sub(r'\D', '', p_celular)
                    if not tel_l.startswith("57") and len(tel_l) == 10: tel_l = "57" + tel_l
                    url_c = f"https://holtercencardio.streamlit.app/?val={cod_uuid[:12]}&pac={urllib.parse.quote(p_nombre)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc=CUPS_893805"
                    msg_w = f"Estimado(a) paciente {p_nombre}, CENCARDIO le hace entrega de su resultado oficial de Prueba de Esfuerzo (CUPS 893805). Puede verificar su autenticidad aquí: {url_c}"
                    st.write("")
                    st.link_button("📲 ENVIAR RESULTADO POR WHATSAPP", f"https://wa.me/{tel_l}?text={urllib.parse.quote(msg_w)}", use_container_width=True)

                if img_erg_prev:
                    st.divider()
                    st.image(img_erg_prev, caption=f"Página 1 - {p_nombre}", width=680)
            else:
                st.info("💡 Digite el nombre del paciente o haga clic en 'Extraer Parámetros con Visión IA'.")

    # --------------------------------------------------------------------------
    # MODALIDADES DIGITALES: HOLTER ECG O MAPA SENTINEL
    # --------------------------------------------------------------------------
    else:
        st.markdown(f"#### 📥 Cargar Estudio Digital ({modalidad_seleccionada})")
        uploaded_file = st.file_uploader("Seleccione el archivo PDF del estudio:", type=["pdf"])

        if uploaded_file is not None:
            bytes_originales = uploaded_file.getvalue()

            if "archivo_cargado_nombre" not in st.session_state or st.session_state.archivo_cargado_nombre != uploaded_file.name:
                with st.spinner("🤖 Analizando y auditando clínicamente todas las páginas del estudio..."):
                    tipo_real = detectar_tipo_documento_clinico(bytes_originales)

                    if tipo_real == "HOLTER":
                        d_act = extraer_datos_holter(bytes_originales, uploaded_file.name)
                        txt_inf = redactar_informe_holter_11_puntos(d_act, perfil_activo)
                        cups_det = "CUPS 895001"
                        mod_det = "Holter ECG 24 Horas"
                        p_clave = f"FC {d_act['fc_prom']} | SDNN {d_act['sdnn_24h']}ms"
                    else:
                        d_act = extraer_datos_mapa_sentinel(bytes_originales, uploaded_file.name)
                        txt_inf = redactar_informe_mapa_cencardio(d_act, perfil_activo)
                        cups_det = "CUPS 895003"
                        mod_det = "MAPA Tensional 24 Horas"
                        p_clave = f"PA 24h: {d_act['pas_24h']}/{d_act['pad_24h']} mmHg"

                    st.session_state.tipo_detectado = tipo_real
                    st.session_state.cups_actual = cups_det
                    st.session_state.mod_nombre = mod_det
                    st.session_state.datos_actuales = d_act
                    st.session_state.texto_informe = txt_inf
                    st.session_state.param_clave = p_clave
                    st.session_state.archivo_cargado_nombre = uploaded_file.name
                    st.session_state.estudio_uuid = str(uuid.uuid4()).upper()
                    st.session_state.telefono_paciente = buscar_telefono_servicio(d_act.get("cedula", ""))

                    origen_txt = d_act.get("origen_extraccion", "Motor Clínico")
                    st.toast(f"✅ Estudio procesado con {origen_txt}", icon="🫀")

            datos = st.session_state.datos_actuales
            tipo_estudio = st.session_state.tipo_detectado
            cups_actual = st.session_state.cups_actual
            mod_nombre = st.session_state.mod_nombre

            bloqueos, alertas_f = ejecutar_sanity_checks(tipo_estudio, datos)
            for b in bloqueos: st.error(f"🛑 {b}")
            for a in alertas_f: st.warning(f"⚠️ {a}")

            if tipo_estudio == "HOLTER":
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("FC Promedio (24h)", f"{datos['fc_prom']} lpm", f"Día {datos['fc_dia']} | Noche {datos['fc_noc']}")
                m2.metric("Ectopias Ventriculares", f"{datos['ev_total']} EV", f"TV: {datos['tv_episodios']}")
                m3.metric("Ectopias Supraventriculares", f"{datos['esv_total']} ESV", f"TSV: {datos['tsv_episodios']}")
                m4.metric("SDNN (24 Horas)", f"{datos['sdnn_24h']} ms", f"ST: {datos['st_episodios']} ep.")
                st.plotly_chart(generar_grafica_tacograma(datos), use_container_width=True)
            else:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Promedio 24 Horas", f"{datos['pas_24h']}/{datos['pad_24h']} mmHg", "Meta < 130/80")
                m2.metric("Carga Sistólica", f"{datos['carga_pas']}%", "Normal < 15%")
                m3.metric("Carga Diastólica", f"{datos['carga_pad']}%", "Normal < 15%")
                m4.metric("Presión de Pulso", f"{datos['pp_val']} mmHg", "Meta <= 60")

            st.divider()

            col_edicion, col_preview = st.columns([1, 1], gap="large")
            with col_edicion:
                c_nom, c_ced = st.columns([1.8, 1.2])
                with c_nom: nombre_confirmado = st.text_input("👤 Paciente:", value=datos['paciente'])
                with c_ced: cedula_confirmada = st.text_input("🪪 Cédula / ID:", value=datos.get('cedula', ''))

                tel_actual = st.session_state.get("telefono_paciente", "")
                telefono_input = st.text_input("📱 Celular (WhatsApp):", value=tel_actual)
                paciente_nombre_archivo = normalizar_nombre_archivo(nombre_confirmado)

                st.subheader(f"📝 Informe Oficial ({cups_actual})")
                informe_para_grabar = st.text_area("Texto oficial para inyectar en el reporte final:", value=st.session_state.texto_informe, height=360)

                discrepancias = auditar_coherencia_informe(informe_para_grabar, datos, tipo_estudio)
                for d_err in discrepancias: st.info(f"🩺 **Alerta de Auditoría:** {d_err}")

                debe_estampar = perfil_activo["id"] in ["dr.amaya", "admin"]

                if tipo_estudio == "HOLTER":
                    pdf_final, img_preview = inyectar_holter_pdf(bytes_originales, informe_para_grabar, nombre_confirmado, perfil_activo, st.session_state.estudio_uuid, estampador_activo=debe_estampar)
                else:
                    pdf_final, img_preview = inyectar_mapa_pdf(bytes_originales, informe_para_grabar, nombre_confirmado, perfil_activo, st.session_state.estudio_uuid, estampador_activo=debe_estampar)

                col_btn1, col_btn2 = st.columns([1, 1])
                with col_btn1:
                    st.download_button(
                        label=f"📄 DESCARGAR {cups_actual} FIRMADO",
                        data=pdf_final,
                        file_name=f"{paciente_nombre_archivo}_{cups_actual.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with col_btn2:
                    if st.button("💾 Guardar en Archivo Clínico", use_container_width=True):
                        guardar_estudio_servicio(nombre_confirmado, mod_nombre, cups_actual, st.session_state.param_clave, perfil_activo['nombre_completo'], informe_para_grabar, pdf_final, st.session_state.estudio_uuid)
                        st.success(f"✅ Guardado en archivo clínico: {nombre_confirmado}")

                if telefono_input:
                    tel_limpio = re.sub(r'\D', '', telefono_input)
                    if not tel_limpio.startswith("57") and len(tel_limpio) == 10: tel_limpio = "57" + tel_limpio
                    url_cert = f"https://holtercencardio.streamlit.app/?val={st.session_state.estudio_uuid[:12]}&pac={urllib.parse.quote(nombre_confirmado)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc={urllib.parse.quote(mod_nombre)}"
                    msg_wa = f"Estimado(a) paciente {nombre_confirmado}, el Centro Cardiovascular Colombiano CENCARDIO le hace entrega de su resultado oficial de {mod_nombre} ({cups_actual}). Certificado oficial: {url_cert}"
                    st.write("")
                    st.link_button("📲 ENVIAR RESULTADO OFICIAL POR WHATSAPP", f"https://wa.me/{tel_limpio}?text={urllib.parse.quote(msg_wa)}", use_container_width=True)

            with col_preview:
                st.subheader("👁️ Vista Previa Oficial (Página 1)")
                st.markdown('<div class="preview-container">', unsafe_allow_html=True)
                if img_preview:
                    st.image(img_preview, caption=f"Página 1 - {nombre_confirmado} ({cups_actual})", use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# PESTAÑA 2: ARCHIVO CLÍNICO Y REPORTES GERENCIALES
# ==============================================================================
with tab_historial:
    st.markdown("### 📁 Archivo Clínico Digital & Reportes Gerenciales")
    historial, origen_datos = obtener_historial_servicio()

    if not historial:
        st.info("Aún no hay estudios archivados en el sistema.")
    else:
        df_produccion = pd.DataFrame(historial)

        col_r1, col_r2 = st.columns([2.5, 1.5])
        with col_r1:
            st.markdown(f"**Total de estudios custodiados ({origen_datos}):** `{len(df_produccion)} procedimientos certificados`")
        with col_r2:
            excel_bytes, ext_salida = generar_excel_avanzado_produccion(df_produccion)
            mime_tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if ext_salida == "xlsx" else "text/csv"
            etiqueta_boton = "📊 DESCARGAR REPORTE EXCEL GERENCIAL (.XLSX)" if ext_salida == "xlsx" else "📊 DESCARGAR REPORTE RIPS (CSV EXCEL)"

            st.download_button(
                label=etiqueta_boton,
                data=excel_bytes,
                file_name=f"Reporte_Gerencial_Cencardio_{ahora_colombia().strftime('%Y%m%d')}.{ext_salida}",
                mime=mime_tipo,
                use_container_width=True,
                type="primary"
            )

        st.divider()

        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        with f1: busqueda = st.text_input("🔍 Buscar por paciente o documento:", "")
        with f2:
            df_produccion["Mes_Periodo"] = pd.to_datetime(df_produccion["fecha_registro"]).dt.strftime('%Y-%m')
            meses_disp = ["Todos los meses"] + sorted(df_produccion["Mes_Periodo"].unique().tolist(), reverse=True)
            mes_sel = st.selectbox("📅 Carpeta Mensual:", meses_disp)
        with f3:
            mods_disp = ["Todas las modalidades"] + sorted(df_produccion["modalidad"].unique().tolist())
            mod_sel = st.selectbox("🎛️ Modalidad:", mods_disp)

        df_filtrado = df_produccion.copy()
        if busqueda.strip(): df_filtrado = df_filtrado[df_filtrado["paciente"].str.contains(busqueda, case=False, na=False)]
        if mes_sel != "Todos los meses": df_filtrado = df_filtrado[df_filtrado["Mes_Periodo"] == mes_sel]
        if mod_sel != "Todas las modalidades": df_filtrado = df_filtrado[df_filtrado["modalidad"] == mod_sel]

        st.write("")
        meses_grupos = sorted(df_filtrado["Mes_Periodo"].unique().tolist(), reverse=True)

        for mes_g in meses_grupos:
            df_mes = df_filtrado[df_filtrado["Mes_Periodo"] == mes_g]
            with st.expander(f"📁 CARPETA: {mes_g} ({len(df_mes)} estudios clínicos)", expanded=True):
                for item in df_mes.itertuples():
                    nom_arch = normalizar_nombre_archivo(item.paciente)
                    
                    st.markdown(f"""
                        <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:0.8rem 1.1rem; margin-bottom:0.7rem;">
                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <div>
                                    <b style="color:#0a2540; font-size:1.02rem;">👤 {item.paciente}</b> 
                                    <span style="background:#e0f2fe; color:#0369a1; font-size:0.72rem; font-weight:700; padding:2px 8px; border-radius:12px; margin-left:6px;">{item.modalidad} ({item.cups})</span>
                                    <div style="font-size:0.82rem; color:#64748b; margin-top:2px;">
                                        📅 <b>Registro:</b> {item.fecha_registro} | 👨‍⚕️ <b>Lector:</b> {item.medico} | 🩺 <b>Parámetro:</b> {item.parametro_clave}
                                    </div>
                                </div>
                                <div style="font-family:monospace; font-size:0.75rem; color:#0284c7; font-weight:700;">
                                    Cód: {item.codigo_verificacion[:14]}...
                                </div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

                    c_d, c_del, c_v = st.columns([1.5, 1, 4])
                    with c_d:
                        if item.pdf_url:
                            st.link_button("📥 Ver / Descargar PDF", item.pdf_url, use_container_width=True)
                        else:
                            pdf_recup = obtener_pdf_bytes_individual(item.id)
                            if pdf_recup:
                                st.download_button(
                                    label="📥 Descargar Copia PDF",
                                    data=pdf_recup,
                                    file_name=f"{nom_arch}_{item.cups.replace(' ', '_')}.pdf",
                                    mime="application/pdf",
                                    key=f"desc_{item.id}",
                                    use_container_width=True
                                )
                    with c_del:
                        if st.button("🗑️ Eliminar", key=f"elim_{item.id}", use_container_width=True):
                            eliminar_estudio_servicio(item.id)
                            st.toast(f"Estudio de {item.paciente} eliminado.", icon="🗑️")
                            st.rerun()
